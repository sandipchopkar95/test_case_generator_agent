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

import argparse
import base64
import json
import os
import re
import sys
from copy import copy
from pathlib import Path
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
DEFAULT_MODEL = "nvidia/nemotron-3-ultra-550b-a55b:free"
MODEL = os.environ.get("OPENROUTER_MODEL", DEFAULT_MODEL)
MODEL_OPTIONS = {
    "Claude": "anthropic/claude-sonnet-4.6",
    "ChatGPT": "openai/gpt-4o",
    "NVIDIA Nemotron 3": DEFAULT_MODEL,
}

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


def load_few_shot_examples(path: str | None) -> str:
    """Load optional few-shot examples (your team's own test cases) to steer style/format."""
    if not path:
        return ""
    p = Path(path)
    if not p.exists():
        print(f"Warning: few-shot examples file '{path}' not found, skipping.", file=sys.stderr)
        return ""
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
    resp = requests.get(url, auth=auth, headers=headers)
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
    message = response.choices[0].message

    tool_calls = getattr(message, "tool_calls", None)
    if tool_calls:
        for call in tool_calls:
            if call.function.name == "submit_test_cases":
                return json.loads(call.function.arguments)

    # Fallback: model replied in plain text instead of calling the tool.
    content = (message.content or "").strip()
    if content.startswith("```"):
        content = content.strip("`")
        if content.startswith("json"):
            content = content[4:]
        content = content.strip()
    if content:
        try:
            return json.loads(content)
        except json.JSONDecodeError:
            pass

    raise RuntimeError(
        "Model did not return structured test cases as expected. Raw response:\n" + str(message)
    )


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


def _generate_json_fallback(
    client: OpenAI,
    requirement_text: str,
    few_shot: str,
    images: list[tuple[str, bytes]] | None = None,
    model: str = MODEL,
) -> dict:
    """Retry without tool calling for models that return reasoning but no tool call."""
    fallback_prompt = f"""Return ONLY valid JSON. Do not explain your reasoning and do not use markdown.

Generate the complete QA test-case object for the requirement below using exactly this top-level shape:
{{"test_cases": [{{"scenario": "...", "description": "...", "preconditions": "...", "steps": [{{"instruction": "...", "expected_result": "...", "name": "..."}}], "priority": "High|Medium|Low", "test_type": "UI|Functional", "is_negative_case": "Yes|No", "automation_candidate": "Yes|No"}}], "open_questions": []}}

Use the complete QA rules in the system prompt. Do not invent UI text or requirements. Every step needs a specific expected result.

Requirement:
{requirement_text}
{few_shot}
"""
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT + "\nReturn JSON in the user-requested shape if tool calling is unavailable."},
        {"role": "user", "content": _multimodal_user_content(fallback_prompt + few_shot, images)},
    ]
    try:
        response = client.chat.completions.create(
            model=model,
            max_tokens=12288,
            messages=messages,
            response_format={"type": "json_object"},
        )
        return validate_generated_result(_extract_tool_result(response))
    except Exception as structured_error:
        if images and _is_image_support_error(structured_error):
            return _generate_json_fallback(client, requirement_text, few_shot, images=None, model=model)
        try:
            response = client.chat.completions.create(
                model=model,
                max_tokens=12288,
                messages=messages,
            )
            return validate_generated_result(_extract_tool_result(response))
        except Exception as plain_error:
            raise RuntimeError(
                "The model returned reasoning but no structured test cases. "
                f"JSON retry failed: {structured_error}; plain-text retry failed: {plain_error}"
            ) from plain_error


def validate_generated_result(result: dict) -> dict:
    """Reject output that cannot satisfy the workbook's execution contract."""
    test_cases = result.get("test_cases")
    if not isinstance(test_cases, list) or not test_cases:
        raise ValueError("The model returned no test cases.")

    scenarios = set()
    for case_number, test_case in enumerate(test_cases, start=1):
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


def generate_test_cases(
    client: OpenAI,
    requirement_text: str,
    few_shot: str,
    images: list[tuple[str, bytes]] | None = None,
    model: str = MODEL,
) -> dict:
    user_content = _multimodal_user_content(requirement_text + few_shot, images)

    try:
        response = client.chat.completions.create(
            model=model,
            max_tokens=8192,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_content},
            ],
            tools=[TEST_CASE_TOOL],
            tool_choice={"type": "function", "function": {"name": "submit_test_cases"}},
        )
    except Exception as error:
        if images and _is_image_support_error(error):
            response = client.chat.completions.create(
                model=model,
                max_tokens=8192,
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
    except RuntimeError:
        return _generate_json_fallback(client, requirement_text, few_shot, images, model)


def self_review(client: OpenAI, requirement_text: str, result: dict, model: str = MODEL) -> dict:
    prompt = SELF_REVIEW_PROMPT.format(
        requirement=requirement_text,
        test_cases_json=json.dumps(result.get("test_cases", []), indent=2),
    )

    response = client.chat.completions.create(
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
    parser = argparse.ArgumentParser(description="Generate QA test cases from requirements via OpenRouter.")
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--input", help="Path to a local .txt/.md file with the PRD/requirements/ticket text.")
    source.add_argument("--story", help="Pasted Jira story or business requirement text.")
    source.add_argument("--jira-ticket", help="Jira ticket ID to fetch, e.g. PROJ-123.")
    parser.add_argument("--jira-id", help="Jira ID used in generated IDs, e.g. RES-123.")
    parser.add_argument("--template", help="Optional .xlsx template with the required 21 columns.")
    parser.add_argument("--output", default="test_cases.xlsx", help="Output workbook path (default: test_cases.xlsx)")
    parser.add_argument("--examples", help="Optional path to a JSON file of few-shot example test cases.")
    parser.add_argument(
        "--self-review",
        action="store_true",
        help="Run a second pass where the model critiques and improves its own coverage.",
    )
    args = parser.parse_args()

    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        print(
            "Error: OPENROUTER_API_KEY environment variable not set.\n"
            "Set it in a .env file (see .env.example) or with `export OPENROUTER_API_KEY=your-key`.",
            file=sys.stderr,
        )
        sys.exit(1)

    client = OpenAI(
        base_url=OPENROUTER_BASE_URL,
        api_key=api_key,
    )

    if args.input:
        requirement_text = Path(args.input).read_text()
        jira_id = args.jira_id or "REQ"
    elif args.story:
        requirement_text = args.story
        jira_id = args.jira_id or "REQ"
    else:
        requirement_text = fetch_jira_ticket(args.jira_ticket)
        jira_id = extract_jira_id(args.jira_ticket)

    few_shot = load_few_shot_examples(args.examples)

    print(f"Generating test cases with {MODEL}...")
    result = generate_test_cases(client, requirement_text, few_shot)

    if args.self_review:
        print("Running self-review pass...")
        result = self_review(client, requirement_text, result)

    write_workbook(result, args.output, jira_id, args.template)


if __name__ == "__main__":
    main()
