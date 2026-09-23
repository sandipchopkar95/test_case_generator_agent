#!/usr/bin/env python3
"""
Test Case Generation Agent
----------------------------------
Reads a PRD / user story / Jira ticket text and generates structured
QA/functional test cases using an LLM via OpenRouter (default model:
NVIDIA Nemotron 3 Ultra).

Usage:
    python generate_tests.py --input requirements.txt --jira-id PROJ-123 --output test_cases.xlsx
    python generate_tests.py --story "As a user, ..." --jira-id PROJ-123 --output test_cases.xlsx
    python generate_tests.py --jira-ticket PROJ-123 --output test_cases.xlsx
"""

from __future__ import annotations

import base64
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import copy
from pathlib import Path
from typing import Callable
from urllib.parse import urlparse

from openai import OpenAI
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension, RowDimension


try:
    from dotenv import load_dotenv
    load_dotenv(dotenv_path=Path(__file__).with_name(".env"), override=True)
except ImportError:
    pass  # dotenv is optional; env vars can be set directly

OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"

# Free tier by default. Swap to "nvidia/nemotron-3-ultra-550b-a55b" (paid, no
# rate limit) via the MODEL env var if you hit free-tier limits.
DEFAULT_MODEL = "nvidia/nemotron-3.5-lightning:free"
MODEL = os.environ.get("OPENROUTER_MODEL", DEFAULT_MODEL)
MODEL_OPTIONS = {
    "Claude": "anthropic/claude-sonnet-4.6",
    "ChatGPT": "openai/gpt-4o",
    "NVIDIA Nemotron 3": MODEL,
}
MAX_REQUIREMENT_CHUNK_CHARS = 4_000
LONG_REQUIREMENT_THRESHOLD = 8_000
LONG_REQUIREMENT_BATCH_TOKENS = 12_288
FREE_REQUIREMENT_BATCH_TOKENS = 4_096
FREE_FALLBACK_MAX_TOKENS = 2_048
FREE_REQUIREMENT_CHUNK_CHARS = 1_000
FREE_MAX_CASES_PER_BATCH = 6
OPENROUTER_RETRY_LIMIT = 3
OPENROUTER_TIMEOUT_RETRY_LIMIT = 0

FREE_JSON_SYSTEM_PROMPT = """Return only valid JSON. Do not include analysis, reasoning, markdown, or
text outside the JSON object. Use exactly this shape:
{"test_cases":[{"scenario":"...","description":"...","preconditions":"...","steps":[{"name":"Action Label","instruction":"...","expected_result":"..."}],"priority":"High|Medium|Low","test_type":"UI|Functional","is_negative_case":"Yes|No","automation_candidate":"Yes|No"}],"open_questions":[]}

Every test case must have at least one step. Every step must contain non-empty
name, instruction, and expected_result strings. Return no more than 6 concise
test cases and do not invent requirements."""

COVERAGE_PLAN_PROMPT = """Analyze the requirement below and return ONLY valid JSON.

Create exactly two balanced, non-overlapping QA coverage groups. Together they must cover every requirement, clarification, validation, visibility rule, boundary, and dependency. Do not generate test cases. Each scope must name the feature areas and details it owns so a QA engineer can generate cases only for that scope.

Return exactly:
{{"groups": [{{"name": "...", "scope": "..."}}, {{"name": "...", "scope": "..."}}]}}

Requirement:
{requirement}
"""

COVERAGE_AUDIT_PROMPT = """Review the generated QA test scenarios against the complete requirement.

Return ONLY valid JSON in this exact shape:
{{"missing_coverage": ["specific requirement, acceptance criterion, validation, state, or dependency that has no adequate scenario"]}}

List only genuinely uncovered items. Do not request duplicate cases, out-of-scope work, or behavior that is not specified. An empty list means every explicit requirement is covered by at least one scenario.

Requirement:
{requirement}

Generated scenarios:
{scenarios}
"""

# ---------------------------------------------------------------------------
# Excel columns must remain in this order for downstream test-management import.
WORKBOOK_COLUMNS = [
    "ID",
    "Scenario",
    "Description",
    "Preconditions",
    "Instructions (test step)",
    "Expected results (test step)",
    "Name (test step)",
    "Type (test step)",
    "Status",
    "Portal",
    "Fix Version",
    "Module Name",
    "Requirement ID",
    "User Role",
    "Assigned to",
    "Priority",
    "Test Type",
    "Created By",
    "Is Negative Case",
    "Automation Candidate",
    "Is Automated",
]

DEFAULT_METADATA = {
    "Status": "Not Run",
    "Portal": "Client",
    "Fix Version": "Release 6",
    "Module Name": "Residences",
    "User Role": "CLIENT REP",
    "Assigned to": "Joshi, Nimisha (nimisha.joshi)",
    "Created By": "Chopkar, Sandip (sandip.chopkar)",
}

# JSON schema the model must return test cases in, expressed as an
# OpenAI-style function/tool definition (OpenRouter is OpenAI-compatible).
# ---------------------------------------------------------------------------
TEST_CASE_TOOL = {
    "type": "function",
    "function": {
        "name": "submit_test_cases",
        "description": "Submit the generated structured test cases.",
        "parameters": {
            "type": "object",
            "properties": {
                "test_cases": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "id": {"type": "string", "description": "e.g. TC-001"},
                            "scenario": {"type": "string"},
                            "description": {"type": "string"},
                            "requirement_id": {
                                "type": "string",
                                "description": "Jira ID or requirement identifier",
                            },
                            "preconditions": {"type": "string"},
                            "steps": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "instruction": {"type": "string"},
                                        "expected_result": {"type": "string"},
                                        "name": {
                                            "type": "string",
                                            "description": "Short action-oriented step name, usually 2-6 words, such as 'Open Resident Profile', 'Apply Date Filter', or 'Verify Activity Order'. Do not use Step 1, Test Step, Verify, Check, or a full sentence.",
                                        },
                                    },
                                    "required": ["instruction", "expected_result", "name"],
                                },
                            },
                            "priority": {
                                "type": "string",
                                "enum": ["High", "Medium", "Low"],
                            },
                            "test_type": {
                                "type": "string",
                                "enum": ["UI", "Functional"],
                            },
                            "is_negative_case": {"type": "string", "enum": ["Yes", "No"]},
                            "automation_candidate": {"type": "string", "enum": ["Yes", "No"]},
                        },
                        "required": [
                            "scenario",
                            "description",
                            "preconditions",
                            "steps",
                            "priority",
                            "test_type",
                            "is_negative_case",
                            "automation_candidate",
                        ],
                    },
                },
                "open_questions": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Ambiguities or missing info in the requirements that a human should clarify.",
                },
            },
            "required": ["test_cases"],
        },
    },
}

SYSTEM_PROMPT = """You are a senior QA engineer generating detailed manual test cases from a Jira story or pasted business requirement.

You are generating an execution-ready QA workbook. Analyze the feature before writing cases: identify actors and roles, entry points, pages, fields, controls, business rules, validations, dependencies, permissions, status transitions, related modules, data relationships, boundary values, error states, navigation paths, persistence, integrations, UI requirements, automation opportunities, and regression impact.

For a focused Jira story, stay strictly within the feature and acceptance criteria described in that story. Do not generate generic regression, performance, API, security, cross-module, or unrelated navigation cases unless the story explicitly requires them. Prefer a smaller set of distinct, execution-ready cases over exhaustive repetition. Use the approved team workbook examples as a style and granularity reference, not as additional requirements.

Uploaded Figma screenshots or UI reference images are authoritative visual source material when provided. Use only text and visual details that are readable in the images. Do not invent UI labels, messages, controls, limits, or behavior that cannot be established from the story, requirements, comments, or images.

Jira comments are included under `Comments / Clarifications`. Treat comments as authoritative clarifications or supplementary requirements when they resolve an ambiguity or gap in the story. Do not treat unrelated discussion, status updates, or opinions as requirements. If comments conflict with the story or acceptance criteria, flag the conflict in `open_questions` instead of silently choosing one.

Coverage must include every applicable scenario without mechanically inventing irrelevant cases:
- Positive, negative, validation, UI, functional, regression, integration, boundary, and edge scenarios.
- Create, edit, update, save, submit, confirm, cancel, delete, activate, transfer, archive, restore, search, filter, sort, pagination, navigation, related-record updates, and persistence where applicable.
- Required, empty, whitespace-only, minimum/exact/below/maximum/above limits, valid/invalid format, special characters, leading/trailing spaces, duplicates, case sensitivity, copy/paste, invalid combinations, and dependency validations where specified.
- Default, empty, loading, disabled, error, unauthorized, session-expired, backend-failure, and network-failure states where applicable.
- Save/cancel/confirm/discard, browser back, breadcrumb, reload, re-login, duplicate-click/double-submit protection, and persistence where applicable.
- Browser forward, tab/sub-tab navigation, direct URL access, navigation after save/cancel/create/delete, and unsaved-change navigation where applicable.
- For important actions, cover default, enabled, disabled, loading, success, failure, single-click, double-click, and rapid-repeat behavior. Verify repeated actions do not create duplicates or execute more than once.
- Verify data integrity: only intended fields change, existing related records remain unchanged unless explicitly required, unrelated records are not modified, and delete/archive affects only the intended record.
- Verify exact UI text, page titles, subtitles, icons, labels, visual representation, clickable cards/tiles, and Figma-defined presentation when those sources are provided. Cover user-facing validation, confirmation, warning, informational, and fallback messages without inventing wording.
- Cover maximum-character limits, truncation, ellipsis, overflow, tooltip/hover behavior, chips, tags, badges, and formatting only when the feature contains those controls or defines those limits.
- Cover conditional UI and verbiage based on status, dates, ownership, lease, record state, current/upcoming/previous/historical records, and field dependencies when applicable.
- Cover sorting and ordering in both ascending/descending directions and most-recently-added-first behavior when the feature supports sorting or ordered records.
- For date-driven features, cover From Date/To Date validation, past/current/future dates, default date values, related-record date relationships, and scheduled jobs or exact trigger timing when specified.
- Cover configuration-driven UI: PSP-configured enabled/disabled entities or sections and conditional section formation/removal when configuration controls the feature.
- Verify state, status, and date changes are reflected consistently across related tiles, screens, dashboards, widgets, resident/owner/staff views, notifications, reports, apartment/unit overviews, and activity/history when impacted.
- Verify cross-screen consistency and related-entity data reflection after updates, reloads, navigation, and other applicable persistence events.
- RBAC only when roles or permissions are defined or reasonably implied. Cover authorized, viewer/restricted, direct URL, and backend enforcement where applicable.
- Cross-tab, dashboard, widget, people/resident/owner/staff/friends-and-family, related-record, search/filter, notification, report, and activity/history regression only when impacted.
- Date/time and upload coverage only when dates/times or uploads exist in the feature.
- Chip/tag/badge/selection truncation and boundary coverage only when those controls exist and limits are specified.

Guidelines:
- Use only the requirements, user story, acceptance criteria, feature details, screenshots, Figma references, provided UI text, and explicitly stated existing behavior as source of truth. Do not invent labels, messages, buttons, fields, statuses, navigation destinations, limits, or business rules. If exact UI wording is absent, describe observable behavior without invented quotation marks.
- `scenario` must be unique, descriptive, action-oriented, and use the form `Verify that <role> can/cannot <action> when <condition>` where appropriate.
- `description` must independently explain what is tested, the condition/input, and the expected business behavior. Never use vague descriptions such as `Verify functionality`, `Verify page`, or `Check save functionality`.
- Every step must be sequential, concise, action-oriented, and independently executable by a new tester. Every step must have a specific expected result directly tied to that step. Never use `works correctly`, `successful`, `as expected`, or other generic expected results.
- Every step `name` must be a short, meaningful action label of approximately 2-6 words that summarizes the step. Use title case where natural, such as `Open Resident Profile`, `Enter Resident Name`, `Apply Date Filter`, `Submit Search`, or `Verify Empty State`. Do not use `Step 1`, `Step 2`, `Test Step`, `Action`, `Verify`, `Check`, a full sentence, or the entire instruction as the name. Keep names distinct within each scenario.
- Mention exact UI text only when it exists in the source material.
- Use `test_type` only as `UI` or `Functional`; never use Regression, Smoke, Sanity, Integration, API, or Performance as a test type.
- Use priority High for core flow, authorization, mandatory validation, data integrity, critical navigation, and duplicate prevention; Medium for standard validation, secondary behavior, search/filter, and dependencies; Low for minor UI and low-impact edge behavior.
- Set automation_candidate to Yes for stable, repeatable, high-value core flows, validation, RBAC, duplicate prevention, boundary, and regression-critical tests. Use No for unstable visual or exploratory behavior. `Is Automated` is always No.
- Mark `is_negative_case` Yes only when the expected outcome is rejection, validation failure, unauthorized access, error handling, duplicate prevention, or another negative path.
- Preserve related records unless the requirements explicitly require them to be created, removed, archived, or modified.
- Ensure every acceptance criterion has at least one scenario. Avoid duplicate scenarios unless the role, rule, validation, boundary, permission, state, data condition, navigation path, or outcome is meaningfully different.
- For each applicable input, consider required, empty, whitespace-only, format, boundary, duplicate, case, copy/paste, special-character, and dependency behavior; do not invent limits that are not defined.
- For dates, uploads, chips, tags, badges, or compact text, include their applicable boundary and persistence coverage only when present in the source feature.
- Before returning cases, remove duplicates, confirm acceptance-criterion traceability, confirm applicable negative/UI/state coverage, and ensure every expected result is observable and independently verifiable.
- Return complete test cases, not analysis or a summary. Put ambiguities in `open_questions`.
- Call the submit_test_cases function to return your output. Do not respond in plain text.
"""

SELF_REVIEW_PROMPT = """You previously generated the following test cases for this requirement.
Critique your own output: are there missing negative cases, boundary values, user roles, or flows?
If you find gaps, output an IMPROVED and COMPLETE set of test cases (not just the additions) by
calling submit_test_cases. If coverage is already solid, resubmit the same set unchanged.

Original requirement:
{requirement}

Previously generated test cases:
{test_cases_json}
"""

GENERIC_EXPECTED_RESULTS = {
    "works correctly",
    "successful",
    "successfully",
    "as expected",
    "system behaves correctly",
    "data is displayed properly",
}

GENERIC_STEP_NAMES = {
    "action",
    "check",
    "step",
    "step 1",
    "step 2",
    "step 3",
    "test step",
    "verify",
}


def _normalise_step_name(name: str, instruction: str) -> str:
    """Keep model-generated step labels short while retaining their action."""
    candidate = re.sub(r"^\s*\d+[.)]\s*", "", name).strip()
    candidate_key = candidate.casefold()
    if candidate_key in GENERIC_STEP_NAMES or re.fullmatch(r"(?:test )?step\s*\d+", candidate_key):
        candidate = instruction.strip()
    candidate = re.sub(r"^\s*\d+[.)]\s*", "", candidate)
    words = candidate.split()
    if len(words) > 6:
        candidate = " ".join(words[:6]).rstrip(".,:;-")
    candidate = re.sub(r"\s+(?:and|or|then|when|while)$", "", candidate, flags=re.IGNORECASE)
    return candidate[:80].strip()


def _load_workbook_examples(path: Path, limit: int = 8) -> list[dict]:
    from openpyxl import load_workbook

    workbook = load_workbook(path, read_only=True, data_only=True)
    worksheet = workbook.active
    headers = [cell.value for cell in next(worksheet.iter_rows())]
    column_index = {str(header): index for index, header in enumerate(headers) if header}
    required = {
        "ID",
        "Scenario",
        "Description",
        "Preconditions",
        "Instructions (test step)",
        "Expected results (test step)",
        "Name (test step)",
        "Priority",
        "Test Type",
    }
    if not required.issubset(column_index):
        raise ValueError("Workbook examples must contain the generated test-case columns.")

    examples = []
    current = None
    for values in worksheet.iter_rows(min_row=2, values_only=True):
        scenario = values[column_index["Scenario"]]
        if scenario:
            if current:
                examples.append(current)
                if len(examples) >= limit:
                    break
            current = {
                "scenario": scenario,
                "description": values[column_index["Description"]] or "",
                "preconditions": values[column_index["Preconditions"]] or "",
                "steps": [],
                "priority": values[column_index["Priority"]] or "Medium",
                "test_type": values[column_index["Test Type"]] or "Functional",
            }
        if current:
            current["steps"].append(
                {
                    "name": values[column_index["Name (test step)"]] or "",
                    "instruction": values[column_index["Instructions (test step)"]] or "",
                    "expected_result": values[column_index["Expected results (test step)"]] or "",
                }
            )
    if current and len(examples) < limit:
        examples.append(current)
    return examples


def load_few_shot_examples(path: str | None) -> str:
    """Load team examples from JSON or an approved Excel workbook."""
    if not path:
        return ""
    p = Path(path)
    if not p.exists():
        print(f"Warning: few-shot examples file '{path}' not found, skipping.", file=sys.stderr)
        return ""
    if p.suffix.casefold() == ".xlsx":
        examples = _load_workbook_examples(p)
    else:
        examples = json.loads(p.read_text())
    return (
        "\n\nHere are examples of test cases written in our team's preferred style. "
        "Match this format, tone, and level of detail:\n\n"
        + json.dumps(examples, indent=2)
    )


def extract_jira_id(value: str) -> str:
    """Extract a Jira key from either a key such as RES-123 or a Jira URL."""
    match = re.search(r"\b([A-Z][A-Z0-9]+-\d+)\b", value.strip(), re.IGNORECASE)
    if not match:
        raise ValueError("Enter a Jira issue key or a Jira link containing one, such as RES-123.")
    return match.group(1).upper()


def fetch_jira_ticket(
    ticket_id: str,
    jira_base_url: str | None = None,
    jira_email: str | None = None,
    jira_api_token: str | None = None,
) -> str:
    """Fetch a Jira ticket's description + acceptance criteria via the Jira REST API.

    Requires environment variables:
        JIRA_BASE_URL   e.g. https://yourcompany.atlassian.net
        JIRA_EMAIL
        JIRA_API_TOKEN
    """
    import requests
    from requests.auth import HTTPBasicAuth

    jira_reference = ticket_id.strip()
    parsed_reference = urlparse(jira_reference)
    base_url = jira_base_url or os.environ.get("JIRA_BASE_URL")
    if parsed_reference.scheme in {"http", "https"} and parsed_reference.netloc:
        base_url = f"{parsed_reference.scheme}://{parsed_reference.netloc}"
    email = jira_email or os.environ.get("JIRA_EMAIL")
    token = jira_api_token or os.environ.get("JIRA_API_TOKEN")

    if not all([base_url, email, token]):
        print(
            "Error: Jira URL/key and credentials are required. Set JIRA_EMAIL and "
            "JIRA_API_TOKEN in .env. For a Jira key without a full URL, also set "
            "JIRA_BASE_URL.",
            file=sys.stderr,
        )
        sys.exit(1)

    jira_id = extract_jira_id(ticket_id)
    auth = HTTPBasicAuth(email, token)
    headers = {"Accept": "application/json"}
    url = f"{base_url.rstrip('/')}/rest/api/3/issue/{jira_id}"
    resp = requests.get(url, auth=auth, headers=headers, timeout=30)
    if resp.status_code == 404:
        raise RuntimeError(
            f"Jira issue {jira_id} was not found at {base_url}. "
            "Check the Jira site, issue key, and that the API user can access this project."
        )
    if resp.status_code in {401, 403}:
        raise RuntimeError(
            f"Jira rejected the request ({resp.status_code}). Check JIRA_EMAIL, "
            "JIRA_API_TOKEN, and project permissions."
        )
    resp.raise_for_status()
    data = resp.json()

    fields = data.get("fields", {})
    summary = fields.get("summary", "")
    description = _adf_to_text(fields.get("description"))

    comments_url = f"{base_url.rstrip('/')}/rest/api/3/issue/{jira_id}/comment"
    comments = []
    start_at = 0
    while True:
        comments_response = requests.get(
            comments_url,
            auth=auth,
            headers=headers,
            params={"startAt": start_at, "maxResults": 100},
            timeout=30,
        )
        if comments_response.status_code in {401, 403}:
            raise RuntimeError(
                f"Jira returned {comments_response.status_code} while reading comments. "
                "Check the API user's comment-view permission for this project."
            )
        if not comments_response.ok:
            break
        page = comments_response.json()
        page_comments = page.get("comments", [])
        comments.extend(page_comments)
        start_at += len(page_comments)
        if not page_comments or start_at >= page.get("total", start_at):
            break

    comment_text = []
    for comment in comments:
        body = _adf_to_text(comment.get("body"))
        if not body:
            continue
        author = (comment.get("author") or {}).get("displayName", "Unknown author")
        created = comment.get("created", "")
        comment_text.append(f"[{created}] {author}:\n{body}")

    comments_section = "\n\nComments / Clarifications:\n" + "\n\n".join(comment_text) if comment_text else ""

    return f"Title: {summary}\n\nDescription / Acceptance Criteria:\n{description}{comments_section}"


def _adf_to_text(adf) -> str:
    """Very small Atlassian Document Format -> plain text flattener."""
    if not adf:
        return ""
    if isinstance(adf, str):
        return adf

    text_parts = []

    def walk(node):
        if isinstance(node, dict):
            if node.get("type") == "text":
                text_parts.append(node.get("text", ""))
            for child in node.get("content", []):
                walk(child)
            if node.get("type") in ("paragraph", "heading", "listItem"):
                text_parts.append("\n")
        elif isinstance(node, list):
            for item in node:
                walk(item)

    walk(adf)
    return "".join(text_parts).strip()


def _extract_tool_result(response) -> dict:
    """Pull the submit_test_cases arguments out of an OpenAI-style response,
    with a fallback to parsing raw JSON from the message content in case the
    model didn't use a tool call (some models are inconsistent about this)."""
    choices = getattr(response, "choices", None) if response is not None else None
    if not choices:
        response_text = str(response)[:2_000]
        raise RuntimeError(
            "OpenRouter returned no response choices. The provider may have stopped "
            f"the request before generation. Response: {response_text}"
        )

    message = getattr(choices[0], "message", None)
    if message is None:
        raise RuntimeError(
            "OpenRouter returned a response without a message. "
            f"Response: {str(response)[:2_000]}"
        )

    tool_calls = getattr(message, "tool_calls", None)
    if tool_calls:
        for call in tool_calls:
            function = getattr(call, "function", None)
            if function and getattr(function, "name", None) == "submit_test_cases":
                arguments = getattr(function, "arguments", "")
                if arguments:
                    return _parse_model_json(arguments)

    # Fallback: model replied in plain text instead of calling the tool.
    content = str(getattr(message, "content", None) or "").strip()
    if content.startswith("```"):
        content = content.strip("`")
        if content.startswith("json"):
            content = content[4:]
        content = content.strip()
    if content:
        try:
            return _parse_model_json(content)
        except ValueError:
            pass

    raise RuntimeError(
        "Model did not return structured test cases as expected. Raw response:\n" + str(message)
    )


def _parse_model_json(content: str) -> dict:
    """Parse model JSON, repairing small syntax defects common in long LLM output."""
    # Some reasoning models occasionally start a step object as `['name': ...]`
    # instead of `[{"name": ...}]`. Repair that precise structural typo before
    # attempting normal JSON parsing.
    content = re.sub(
        r'("steps"\s*:\s*)\[\s*"(?=(?:name|instruction|expected_result)"\s*:)',
        r'\1[{"',
        content,
    )
    try:
        result = json.loads(content)
    except json.JSONDecodeError as original_error:
        try:
            from json_repair import repair_json

            result = repair_json(content, return_objects=True)
        except Exception as repair_error:
            raise ValueError("Model response is not valid JSON.") from repair_error
        if not isinstance(result, dict):
            raise ValueError("Repaired model response is not a JSON object.") from original_error
    if not isinstance(result, dict):
        raise ValueError("Model response is not a JSON object.")
    return result


def _image_data_url(image: tuple[str, bytes]) -> str:
    filename, image_bytes = image
    extension = Path(filename).suffix.lower().lstrip(".") or "png"
    mime_type = "jpeg" if extension in {"jpg", "jpeg"} else extension
    return f"data:image/{mime_type};base64,{base64.b64encode(image_bytes).decode('ascii')}"


def _multimodal_user_content(
    text: str,
    images: list[tuple[str, bytes]] | None = None,
):
    if not images:
        return text
    content = [{"type": "text", "text": text}]
    for image in images:
        content.append(
            {
                "type": "image_url",
                "image_url": {"url": _image_data_url(image)},
            }
        )
    return content


def _is_image_support_error(error: Exception) -> bool:
    message = str(error).casefold()
    return "image input" in message or "image support" in message or "filter by image" in message


def _openrouter_completion(client: OpenAI, **kwargs):
    """Retry transient OpenRouter throttling without duplicating application work."""
    for attempt in range(OPENROUTER_RETRY_LIMIT + 1):
        try:
            return client.chat.completions.create(**kwargs)
        except Exception as error:
            status_code = getattr(error, "status_code", None)
            message = str(error).casefold()
            transient = status_code == 429 or status_code in {408, 409, 500, 502, 503, 504}
            transient = transient or any(
                marker in message for marker in ("rate limit", "too many requests", "timed out", "timeout")
            )
            retry_limit = (
                OPENROUTER_TIMEOUT_RETRY_LIMIT
                if status_code == 504 or "timeout" in message or "timed out" in message
                else OPENROUTER_RETRY_LIMIT
            )
            if not transient or attempt >= retry_limit:
                raise
            retry_after = getattr(error, "headers", {}).get("retry-after") if getattr(error, "headers", None) else None
            try:
                delay = min(30, max(2, float(retry_after))) if retry_after else min(30, 2 ** attempt * 2)
            except (TypeError, ValueError):
                delay = min(30, 2 ** attempt * 2)
            time.sleep(delay)


def _generate_json_fallback(
    client: OpenAI,
    requirement_text: str,
    few_shot: str,
    images: list[tuple[str, bytes]] | None = None,
    model: str = MODEL,
    max_tokens: int = 12288,
) -> dict:
    """Retry without tool calling for models that return reasoning but no tool call."""
    fallback_prompt = f"""Return ONLY valid JSON. Do not explain your reasoning and do not use markdown.

Generate the complete QA test-case object for the requirement below using exactly this top-level shape:
{{"test_cases": [{{"scenario": "...", "description": "...", "preconditions": "...", "steps": [{{"instruction": "...", "expected_result": "...", "name": "..."}}], "priority": "High|Medium|Low", "test_type": "UI|Functional", "is_negative_case": "Yes|No", "automation_candidate": "Yes|No"}}], "open_questions": []}}

Use the complete QA rules in the system prompt. Do not invent UI text or requirements. Every step needs a specific expected result.
Generate no more than {FREE_MAX_CASES_PER_BATCH if ':free' in model else 20} concise test cases for this request. Cover the most important distinct requirements in this scope; do not produce a long exhaustive list.

Requirement:
{requirement_text}
{few_shot}
"""
    system_prompt = FREE_JSON_SYSTEM_PROMPT if ":free" in model else SYSTEM_PROMPT + "\nReturn JSON in the user-requested shape if tool calling is unavailable."
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": _multimodal_user_content(fallback_prompt, images)},
    ]
    try:
        response = _openrouter_completion(
            client,
            model=model,
            max_tokens=max_tokens,
            messages=messages,
            response_format={"type": "json_object"},
        )
        parsed_result = _extract_tool_result(response)
        try:
            return validate_generated_result(parsed_result)
        except ValueError as validation_error:
            return _repair_generated_result(
                client,
                requirement_text,
                parsed_result,
                model,
                validation_error,
            )
    except Exception as structured_error:
        if images and _is_image_support_error(structured_error):
            return _generate_json_fallback(client, requirement_text, few_shot, images=None, model=model, max_tokens=max_tokens)
        if _is_provider_unavailable(structured_error):
            raise RuntimeError(
                f"OpenRouter could not use model '{model}'. "
                "The selected provider returned 404; choose another model in Settings."
            ) from structured_error
        if _is_provider_timeout(structured_error):
            raise RuntimeError(
                "OpenRouter timed out before returning test cases. "
                "Retry this run, use a paid model, or shorten the requirement."
            ) from structured_error
        try:
            response = _openrouter_completion(
                client,
                model=model,
                max_tokens=max_tokens,
                messages=messages,
            )
            return validate_generated_result(_extract_tool_result(response))
        except Exception as plain_error:
            if _is_provider_unavailable(plain_error):
                raise RuntimeError(
                    f"OpenRouter could not use model '{model}'. "
                    "The selected provider returned 404; choose another model in Settings."
                ) from plain_error
            if _is_provider_timeout(plain_error):
                raise RuntimeError(
                    "OpenRouter timed out before returning test cases. "
                    "Retry this run, use a paid model, or shorten the requirement."
                ) from plain_error
            raise RuntimeError(
                "The model returned reasoning but no structured test cases. "
                f"JSON retry failed: {structured_error}; plain-text retry failed: {plain_error}"
            ) from plain_error


def _is_provider_timeout(error: Exception) -> bool:
    """Identify provider-side timeouts before reporting a parsing failure."""
    status_code = getattr(error, "status_code", None)
    message = str(error).casefold()
    return status_code == 504 or "timeout" in message or "timed out" in message


def _repair_generated_result(
    client: OpenAI,
    requirement_text: str,
    result: dict,
    model: str,
    validation_error: ValueError,
) -> dict:
    """Ask the model to repair a nearly valid response without redoing analysis."""
    repair_prompt = f"""Repair this JSON test-case object and return only valid JSON.

Validation error: {validation_error}
Every test case must contain at least one step. Every step must contain non-empty
name, instruction, and expected_result fields. Preserve the requirements and
do not add behavior that is absent from the source.

Requirement:
{requirement_text}

JSON to repair:
{json.dumps(result, ensure_ascii=False)}
"""
    response = _openrouter_completion(
        client,
        model=model,
        max_tokens=FREE_FALLBACK_MAX_TOKENS if ":free" in model else 4096,
        messages=[
            {"role": "system", "content": FREE_JSON_SYSTEM_PROMPT if ":free" in model else SYSTEM_PROMPT},
            {"role": "user", "content": repair_prompt},
        ],
        response_format={"type": "json_object"},
    )
    return validate_generated_result(_extract_tool_result(response))


def _is_provider_unavailable(error: Exception) -> bool:
    """Identify a model/provider route that OpenRouter cannot serve."""
    status_code = getattr(error, "status_code", None)
    message = str(error).casefold()
    return status_code == 404 or "provider returned error" in message


def validate_generated_result(result: dict) -> dict:
    """Reject output that cannot satisfy the workbook's execution contract."""
    test_cases = result.get("test_cases")
    if not isinstance(test_cases, list) or not test_cases:
        raise ValueError("The model returned no test cases.")

    scenarios = set()
    for case_number, test_case in enumerate(test_cases, start=1):
        if not isinstance(test_case, dict):
            raise ValueError(f"Test case {case_number} is malformed.")
        # Free-tier models occasionally omit a classification field even when
        # the rest of the case is complete. Use conservative import-safe values.
        test_case.setdefault("priority", "Medium")
        test_case.setdefault("test_type", "Functional")
        test_case.setdefault("is_negative_case", "No")
        test_case.setdefault("automation_candidate", "No")
        for field in ("scenario", "description", "preconditions", "priority", "test_type"):
            if not str(test_case.get(field, "")).strip():
                raise ValueError(f"Test case {case_number} is missing '{field}'.")
        scenario_key = test_case["scenario"].strip().casefold()
        if scenario_key in scenarios:
            raise ValueError(f"Duplicate scenario returned: {test_case['scenario']}")
        scenarios.add(scenario_key)
        if test_case["priority"] not in {"High", "Medium", "Low"}:
            raise ValueError(f"Unsupported priority in test case {case_number}.")
        if test_case["test_type"] not in {"UI", "Functional"}:
            raise ValueError(f"Unsupported Test Type in test case {case_number}.")
        if test_case.get("is_negative_case") not in {"Yes", "No"}:
            raise ValueError(f"Invalid negative-case value in test case {case_number}.")
        if test_case.get("automation_candidate") not in {"Yes", "No"}:
            raise ValueError(f"Invalid automation-candidate value in test case {case_number}.")

        steps = test_case.get("steps")
        if not isinstance(steps, list) or not steps:
            raise ValueError(f"Test case {case_number} has no test steps.")
        for step_number, step in enumerate(steps, start=1):
            if not isinstance(step, dict):
                raise ValueError(f"Test case {case_number}, step {step_number} is malformed.")
            instruction = str(step.get("instruction", "")).strip()
            expected_result = str(step.get("expected_result", "")).strip()
            name = _normalise_step_name(str(step.get("name", "")), instruction)
            if not instruction or not expected_result or not name:
                raise ValueError(f"Test case {case_number}, step {step_number} is incomplete.")
            step["name"] = name
            name_key = name.casefold()
            if name_key in GENERIC_STEP_NAMES or re.fullmatch(r"(?:test )?step\s*\d+", name_key):
                raise ValueError(
                    f"Test case {case_number}, step {step_number} has a generic step name."
                )
            if len(name.split()) > 6:
                raise ValueError(
                    f"Test case {case_number}, step {step_number} step name is too long."
                )
            if expected_result.casefold() in GENERIC_EXPECTED_RESULTS:
                raise ValueError(
                    f"Test case {case_number}, step {step_number} has a generic expected result."
                )
    return result


def _generate_test_case_batch(
    client: OpenAI,
    requirement_text: str,
    few_shot: str,
    images: list[tuple[str, bytes]] | None = None,
    model: str = MODEL,
    max_tokens: int = 8192,
) -> dict:
    if ":free" in model:
        return _generate_json_fallback(
            client,
            requirement_text,
            few_shot,
            images,
            model,
            max_tokens=min(max_tokens, FREE_FALLBACK_MAX_TOKENS),
        )

    user_content = _multimodal_user_content(requirement_text + few_shot, images)

    try:
        response = _openrouter_completion(
            client,
            model=model,
            max_tokens=max_tokens,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_content},
            ],
            tools=[TEST_CASE_TOOL],
            tool_choice={"type": "function", "function": {"name": "submit_test_cases"}},
        )
    except Exception as error:
        if images and _is_image_support_error(error):
            response = _openrouter_completion(
                client,
                model=model,
                max_tokens=max_tokens,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": requirement_text + few_shot},
                ],
                tools=[TEST_CASE_TOOL],
                tool_choice={"type": "function", "function": {"name": "submit_test_cases"}},
            )
        else:
            raise

    try:
        return validate_generated_result(_extract_tool_result(response))
    except (RuntimeError, ValueError):
        return _generate_json_fallback(client, requirement_text, few_shot, images, model, max_tokens=max_tokens)


def _split_requirement_for_generation(
    requirement_text: str,
    max_chars: int = MAX_REQUIREMENT_CHUNK_CHARS,
) -> list[str]:
    """Split long PRDs on paragraph boundaries so one response is not truncated.

    A very detailed story can require more output than a provider permits in one
    response. Each returned batch is independently valid, then all batches are
    merged below. Short stories retain the existing single-request behavior.
    """
    if len(requirement_text) <= max_chars:
        return [requirement_text]

    paragraphs = [part.strip() for part in re.split(r"\n\s*\n", requirement_text) if part.strip()]
    title = next((part for part in paragraphs if part.casefold().startswith("title:")), "")
    chunks: list[str] = []
    current: list[str] = []
    current_length = 0
    for paragraph in paragraphs:
        paragraph_length = len(paragraph) + 2
        if paragraph_length > max_chars:
            if current:
                chunks.append("\n\n".join(current))
                current = []
                current_length = 0
            chunks.extend(
                paragraph[start : start + max_chars]
                for start in range(0, len(paragraph), max_chars)
            )
            continue
        if current and current_length + paragraph_length > max_chars:
            chunks.append("\n\n".join(current))
            current = []
            current_length = 0
        current.append(paragraph)
        current_length += paragraph_length
    if current:
        chunks.append("\n\n".join(current))

    if title:
        return [chunk if chunk.startswith(title) else f"{title}\n\n{chunk}" for chunk in chunks]
    return chunks


def _merge_generated_batches(results: list[dict]) -> dict:
    """Merge batches while retaining the first version of any duplicate scenario."""
    merged_cases = []
    open_questions = []
    seen_scenarios = set()
    seen_questions = set()
    for result in results:
        for test_case in result.get("test_cases", []):
            scenario_key = str(test_case.get("scenario", "")).strip().casefold()
            if scenario_key and scenario_key not in seen_scenarios:
                seen_scenarios.add(scenario_key)
                merged_cases.append(test_case)
        for question in result.get("open_questions") or []:
            question_key = str(question).strip().casefold()
            if question_key and question_key not in seen_questions:
                seen_questions.add(question_key)
                open_questions.append(question)
    return validate_generated_result({"test_cases": merged_cases, "open_questions": open_questions})


def _fallback_coverage_groups(requirement_text: str) -> list[dict]:
    """Create bounded scopes without an AI planning response if needed."""
    chunks = _split_requirement_for_generation(
        requirement_text,
        max_chars=FREE_REQUIREMENT_CHUNK_CHARS,
    )
    return [
        {
            "name": "Complete requirement" if len(chunks) == 1 else f"Requirement coverage chunk {index}",
            "scope": chunk,
        }
        for index, chunk in enumerate(chunks, start=1)
    ]


def _plan_coverage_groups(client: OpenAI, requirement_text: str, model: str) -> list[dict]:
    """Use a small response to assign the full requirement to two QA scopes."""
    try:
        response = _openrouter_completion(
            client,
            model=model,
            max_tokens=1200,
            messages=[{"role": "user", "content": COVERAGE_PLAN_PROMPT.format(requirement=requirement_text)}],
            response_format={"type": "json_object"},
        )
        planned = _parse_model_json((response.choices[0].message.content or "").strip())
        groups = planned.get("groups")
        if not isinstance(groups, list) or len(groups) != 2:
            raise ValueError("Coverage plan did not contain two groups.")
        normalised = [
            {"name": str(group.get("name", "")).strip(), "scope": str(group.get("scope", "")).strip()}
            for group in groups
            if isinstance(group, dict)
        ]
        if len(normalised) != 2 or any(not group["scope"] for group in normalised):
            raise ValueError("Coverage plan contains an incomplete group.")
        return normalised
    except Exception:
        return _fallback_coverage_groups(requirement_text)


def _generate_coverage_group(
    client: OpenAI,
    requirement_text: str,
    few_shot: str,
    group: dict,
    model: str,
) -> dict:
    if ":free" in model:
        # Free-tier endpoints are latency-sensitive. Each fallback group already
        # contains a bounded requirement chunk, so sending the full story and
        # large reference context again only increases timeout risk.
        compact_prompt = f"""Generate up to {FREE_MAX_CASES_PER_BATCH} concise QA test cases for this requirement chunk.

Return only the required JSON test-case object. Cover only behavior explicitly present in the chunk. Every case needs executable steps with observable expected results.

Requirement chunk:
{group['scope']}
"""
        compact_examples = ""
        return _generate_test_case_batch(
            client,
            compact_prompt,
            compact_examples,
            images=None,
            model=model,
            max_tokens=FREE_FALLBACK_MAX_TOKENS,
        )

    group_prompt = f"""Generate complete test cases only for the assigned coverage group below.

    Do not omit any explicit condition, validation, visibility rule, boundary, dependency, or state change in the assigned group. Do not generate cases that belong exclusively to the other coverage group.
    Generate no more than {FREE_MAX_CASES_PER_BATCH if ':free' in model else 20} concise test cases for this assigned scope. Return complete JSON before adding more cases.

Assigned group: {group['name']}
Assigned scope:
{group['scope']}

Full requirement source (use it only to understand the assigned scope and its dependencies):
{requirement_text}
"""
    return _generate_test_case_batch(
        client,
        group_prompt,
        few_shot,
        images=None,
        model=model,
        max_tokens=LONG_REQUIREMENT_BATCH_TOKENS,
    )


def _find_missing_coverage(client: OpenAI, requirement_text: str, result: dict, model: str) -> list[str]:
    """Use a small review response to find explicit requirements without a case."""
    scenario_summary = [
        {"scenario": case.get("scenario", ""), "description": case.get("description", "")}
        for case in result.get("test_cases", [])
    ]
    try:
        response = _openrouter_completion(
            client,
            model=model,
            max_tokens=1600,
            messages=[
                {
                    "role": "user",
                    "content": COVERAGE_AUDIT_PROMPT.format(
                        requirement=requirement_text,
                        scenarios=json.dumps(scenario_summary, ensure_ascii=False),
                    ),
                }
            ],
            response_format={"type": "json_object"},
        )
        audit = _parse_model_json((response.choices[0].message.content or "").strip())
        missing_coverage = audit.get("missing_coverage", [])
        if not isinstance(missing_coverage, list):
            return []
        return [str(item).strip() for item in missing_coverage if str(item).strip()][:20]
    except Exception:
        # Generation remains available if a provider does not support JSON mode
        # for the lightweight audit request.
        return []


def _generate_missing_coverage(
    client: OpenAI,
    requirement_text: str,
    few_shot: str,
    missing_coverage: list[str],
    model: str,
) -> dict:
    prompt = f"""Generate additional test cases only for the uncovered requirement points below.

Do not repeat existing scenarios. Each point must receive one or more independently executable test cases with observable expected results.

Uncovered points:
{json.dumps(missing_coverage, ensure_ascii=False, indent=2)}

Full requirement source:
{requirement_text}
"""
    return _generate_test_case_batch(
        client, prompt, few_shot, images=None, model=model, max_tokens=LONG_REQUIREMENT_BATCH_TOKENS
    )


def generate_test_cases(
    client: OpenAI,
    requirement_text: str,
    few_shot: str,
    images: list[tuple[str, bytes]] | None = None,
    model: str = MODEL,
    progress_callback: Callable[[str], None] | None = None,
    coverage_groups: list[dict] | None = None,
) -> dict:
    """Generate complete coverage, parallelizing two groups for detailed stories."""
    if len(requirement_text) <= LONG_REQUIREMENT_THRESHOLD and not (
        ":free" in model and len(requirement_text) > FREE_REQUIREMENT_CHUNK_CHARS
    ):
        return _generate_test_case_batch(client, requirement_text, few_shot, images, model)

    if progress_callback:
        progress_callback("Planning complete coverage for this detailed requirement...")
    groups = coverage_groups or (
        _fallback_coverage_groups(requirement_text)
        if ":free" in model
        else _plan_coverage_groups(client, requirement_text, model)
    )
    if progress_callback and ":free" not in model:
        progress_callback("Generating two focused coverage groups in parallel...")

    results: list[dict | None] = [None] * len(groups)
    if ":free" in model:
        if progress_callback:
            progress_callback("Generating focused coverage groups sequentially for NVIDIA free-tier reliability...")
        for index, group in enumerate(groups):
            if progress_callback:
                progress_callback(f"Generating coverage group {index + 1} of {len(groups)}...")
            results[index] = _generate_coverage_group(client, requirement_text, few_shot, group, model)
    else:
        failures: list[int] = []
        with ThreadPoolExecutor(max_workers=min(2, len(groups))) as executor:
            futures = {
                executor.submit(_generate_coverage_group, client, requirement_text, few_shot, group, model): index
                for index, group in enumerate(groups)
            }
            for future in as_completed(futures):
                index = futures[future]
                try:
                    results[index] = future.result()
                except Exception:
                    failures.append(index)

        for index in failures:
            if progress_callback:
                progress_callback(f"Retrying coverage group {index + 1} after provider throttling...")
            results[index] = _generate_coverage_group(client, requirement_text, few_shot, groups[index], model)

    if progress_callback:
        progress_callback("Combining coverage groups and validating the workbook data...")
    results = [result for result in results if result is not None]
    merged_result = _merge_generated_batches(results)
    if ":free" in model:
        merged_result["coverage_audit"] = {
            "missing_points_addressed": 0,
            "status": "skipped_for_free_tier_speed",
        }
        return merged_result

    if progress_callback:
        progress_callback("Auditing scenario traceability against the complete requirement...")
    missing_coverage = _find_missing_coverage(client, requirement_text, merged_result, model)
    if missing_coverage:
        if progress_callback:
            progress_callback(f"Generating targeted cases for {len(missing_coverage)} uncovered requirement point(s)...")
        merged_result = _merge_generated_batches(
            [merged_result, _generate_missing_coverage(client, requirement_text, few_shot, missing_coverage, model)]
        )
    merged_result["coverage_audit"] = {
        "missing_points_addressed": len(missing_coverage),
        "status": "reviewed",
    }
    return merged_result


def self_review(client: OpenAI, requirement_text: str, result: dict, model: str = MODEL) -> dict:
    prompt = SELF_REVIEW_PROMPT.format(
        requirement=requirement_text,
        test_cases_json=json.dumps(result.get("test_cases", []), indent=2),
    )

    response = _openrouter_completion(
        client,
        model=model,
        max_tokens=8192,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
        tools=[TEST_CASE_TOOL],
        tool_choice={"type": "function", "function": {"name": "submit_test_cases"}},
    )

    try:
        return validate_generated_result(_extract_tool_result(response))
    except RuntimeError:
        return result  # fallback: keep original if review pass fails to parse


def _normalise_jira_id(jira_id: str) -> str:
    value = "".join(char if char.isalnum() or char in "-_" else "-" for char in jira_id.strip())
    return value.strip("-_") or "REQ"


def _metadata_value(test_case: dict, key: str, default: str) -> str:
    value = test_case.get(key, default)
    if isinstance(value, bool):
        return "Yes" if value else "No"
    return str(value or default)


def _build_workbook_rows(result: dict, jira_id: str, metadata: dict | None = None) -> list[list[str]]:
    rows = []
    normalised_id = _normalise_jira_id(jira_id)
    provided_metadata = metadata or {}
    scenario_metadata = {**DEFAULT_METADATA, **provided_metadata}
    for case_number, test_case in enumerate(result.get("test_cases", []), start=1):
        case_id = f"{normalised_id}_{case_number:03d}"
        steps = test_case.get("steps") or []
        if not steps:
            continue

        test_type = test_case.get("test_type", "Functional")
        if test_type not in {"UI", "Functional"}:
            test_type = "Functional"
        first_step_metadata = {
            **scenario_metadata,
            "Requirement ID": normalised_id,
            "Is Automated": "No",
            "Priority": provided_metadata.get("Priority", test_case.get("priority", "Medium")),
            "Test Type": provided_metadata.get("Test Type", test_type),
            "Is Negative Case": provided_metadata.get(
                "Is Negative Case", _metadata_value(test_case, "is_negative_case", "No")
            ),
            "Automation Candidate": provided_metadata.get(
                "Automation Candidate", _metadata_value(test_case, "automation_candidate", "No")
            ),
        }
        first_step = steps[0]
        first_row = {column: "" for column in WORKBOOK_COLUMNS}
        first_row.update(first_step_metadata)
        first_row.update(
            {
                "ID": case_id,
                "Scenario": test_case.get("scenario", ""),
                "Description": test_case.get("description", test_case.get("scenario", "")),
                "Preconditions": test_case.get("preconditions", ""),
                "Instructions (test step)": first_step.get("instruction", ""),
                "Expected results (test step)": first_step.get("expected_result", ""),
                "Name (test step)": first_step.get("name", ""),
                "Type (test step)": "Step",
            }
        )
        rows.append([first_row[column] for column in WORKBOOK_COLUMNS])

        for step in steps[1:]:
            rows.append(
                [
                    case_id if column == "ID" else step.get("instruction", "")
                    if column == "Instructions (test step)"
                    else step.get("expected_result", "")
                    if column == "Expected results (test step)"
                    else step.get("name", "")
                    if column == "Name (test step)"
                    else "Step"
                    if column == "Type (test step)"
                    else ""
                    for column in WORKBOOK_COLUMNS
                ]
            )
    return rows


def _prepare_workbook(template_path: str | None):
    if template_path:
        path = Path(template_path)
        if not path.exists():
            raise FileNotFoundError(f"Workbook template not found: {template_path}")
        workbook = load_workbook(path)
        worksheet = workbook.active
        headers = [worksheet.cell(row=1, column=index).value for index in range(1, len(WORKBOOK_COLUMNS) + 1)]
        if headers != WORKBOOK_COLUMNS:
            raise ValueError(
                "Template headers must contain exactly the required 21 columns in the specified order."
            )
        return workbook, worksheet

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Test Cases"
    worksheet.append(WORKBOOK_COLUMNS)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in worksheet[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(WORKBOOK_COLUMNS))}1"
    for index, column in enumerate(WORKBOOK_COLUMNS, start=1):
        worksheet.column_dimensions[get_column_letter(index)].width = max(14, min(34, len(column) + 4))
    return workbook, worksheet


def write_workbook(
    result: dict,
    output_path: str,
    jira_id: str,
    template_path: str | None = None,
    metadata: dict | None = None,
):
    workbook, worksheet = _prepare_workbook(template_path)
    rows = _build_workbook_rows(result, jira_id, metadata)
    last_row = len(rows) + 1

    # Templates often contain formatted placeholder columns/rows beyond the
    # import contract. Remove them so the downloaded workbook is exact.
    if worksheet.max_column > len(WORKBOOK_COLUMNS):
        worksheet.delete_cols(
            len(WORKBOOK_COLUMNS) + 1,
            worksheet.max_column - len(WORKBOOK_COLUMNS),
        )
    if worksheet.max_row > last_row:
        worksheet.delete_rows(last_row + 1, worksheet.max_row - last_row)

    # A template may contain collapsed/hidden placeholder rows or saved filter
    # criteria. Clear those view settings so every generated step is visible.
    for row_number in range(1, max(worksheet.max_row, 1) + 1):
        existing_row = worksheet.row_dimensions[row_number]
        worksheet.row_dimensions[row_number] = RowDimension(
            worksheet,
            index=row_number,
            height=existing_row.height,
            hidden=False,
            collapsed=False,
        )
    for column_number in range(1, len(WORKBOOK_COLUMNS) + 1):
        column_letter = get_column_letter(column_number)
        existing_column = worksheet.column_dimensions[column_letter]
        worksheet.column_dimensions[column_letter] = ColumnDimension(
            worksheet,
            index=column_letter,
            width=existing_column.width,
        )
    worksheet.auto_filter.filterColumn = []
    for row in worksheet.iter_rows(min_row=2, max_col=len(WORKBOOK_COLUMNS)):
        for cell in row:
            cell.value = None

    for row_number, values in enumerate(rows, start=2):
        for column_number, value in enumerate(values, start=1):
            cell = worksheet.cell(row=row_number, column=column_number)
            cell.value = value
            cell.alignment = copy(cell.alignment)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    last_row = max(1, len(rows) + 1)
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(WORKBOOK_COLUMNS))}{last_row}"
    worksheet.freeze_panes = "A2"
    for row_number in range(1, last_row + 1):
        worksheet.row_dimensions[row_number].hidden = False
        worksheet.row_dimensions[row_number].collapsed = False
    for column_number in range(1, len(WORKBOOK_COLUMNS) + 1):
        column_dimension = worksheet.column_dimensions[get_column_letter(column_number)]
        column_dimension.hidden = False
        column_dimension.collapsed = False
    workbook.save(output_path)
    validate_workbook(output_path, jira_id, metadata)
    print(f"\u2714 Wrote {len(rows)} step rows for {len(result.get('test_cases', []))} scenarios to {output_path}")

    open_questions = result.get("open_questions") or []
    if open_questions:
        print("\n\u26a0 Open questions / ambiguities the model flagged:")
        for question in open_questions:
            print(f"  - {question}")


def validate_workbook(workbook_path: str, jira_id: str, metadata: dict | None = None):
    """Verify the exported workbook before it becomes available for download."""
    workbook = load_workbook(workbook_path)
    if len(workbook.worksheets) != 1:
        raise ValueError("Generated workbook must contain exactly one worksheet.")
    worksheet = workbook.active
    headers = [worksheet.cell(row=1, column=index).value for index in range(1, len(WORKBOOK_COLUMNS) + 1)]
    if headers != WORKBOOK_COLUMNS or worksheet.max_column != len(WORKBOOK_COLUMNS):
        raise ValueError("Generated workbook does not contain the required 21 columns in order.")

    expected_metadata = {**DEFAULT_METADATA, **(metadata or {})}
    normalised_id = _normalise_jira_id(jira_id)
    seen_ids = set()
    scenario_ids = set()
    for row_number in range(2, worksheet.max_row + 1):
        values = [worksheet.cell(row=row_number, column=index).value for index in range(1, len(WORKBOOK_COLUMNS) + 1)]
        case_id = values[0]
        if not case_id or not values[4] or not values[5] or not values[6] or values[7] != "Step":
            raise ValueError(f"Generated workbook has an incomplete step row at row {row_number}.")
        seen_ids.add(case_id)
        if case_id not in scenario_ids:
            scenario_ids.add(case_id)
            expected_case_id = f"{normalised_id}_{len(scenario_ids):03d}"
            if case_id != expected_case_id:
                raise ValueError(f"Scenario IDs are not sequential at row {row_number}.")
            if not all(values[index] for index in (1, 2, 3, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)):
                raise ValueError(f"Scenario metadata is incomplete at row {row_number}.")
            fixed_metadata = {
                8: expected_metadata["Status"],
                9: expected_metadata["Portal"],
                10: expected_metadata["Fix Version"],
                11: expected_metadata["Module Name"],
                12: normalised_id,
                13: expected_metadata["User Role"],
                14: expected_metadata["Assigned to"],
                17: expected_metadata["Created By"],
                20: "No",
            }
            if any(values[index] != expected for index, expected in fixed_metadata.items()):
                raise ValueError(f"Scenario metadata is incorrect at row {row_number}.")
        else:
            if any(value not in (None, "") for value in values[1:4] + values[8:21]):
                raise ValueError(f"Follow-up row {row_number} contains scenario metadata.")
    if not scenario_ids or not seen_ids:
        raise ValueError("Generated workbook contains no test cases.")


def main():
    from qa_generator.cli import run

    run()


if __name__ == "__main__":
    main()
