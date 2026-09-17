#!/usr/bin/env python3
"""
Test Case Generation Agent
----------------------------------
Reads a PRD / user story / Jira ticket text and generates structured
QA/functional test cases using an LLM via OpenRouter (default model:
NVIDIA Nemotron 3 Ultra).

Usage:
    python generate_tests.py --input requirements.txt --output test_cases.csv
    python generate_tests.py --input requirements.txt --output test_cases.csv --self-review
    python generate_tests.py --jira-ticket PROJ-123 --output test_cases.csv
"""

import argparse
import json
import os
import sys
from pathlib import Path

import pandas as pd
from openai import OpenAI

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # dotenv is optional; env vars can be set directly

OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"

# Free tier by default. Swap to "nvidia/nemotron-3-ultra-550b-a55b" (paid, no
# rate limit) via the MODEL env var if you hit free-tier limits.
DEFAULT_MODEL = "nvidia/nemotron-3-ultra-550b-a55b:free"
MODEL = os.environ.get("OPENROUTER_MODEL", DEFAULT_MODEL)

# ---------------------------------------------------------------------------
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
                            "title": {"type": "string"},
                            "requirement_ref": {
                                "type": "string",
                                "description": "Which requirement / AC this test case maps to",
                            },
                            "preconditions": {"type": "string"},
                            "steps": {
                                "type": "array",
                                "items": {"type": "string"},
                            },
                            "expected_result": {"type": "string"},
                            "priority": {
                                "type": "string",
                                "enum": ["High", "Medium", "Low"],
                            },
                            "type": {
                                "type": "string",
                                "enum": [
                                    "Functional",
                                    "Negative",
                                    "Edge Case",
                                    "Regression",
                                    "Alternate Flow",
                                ],
                            },
                        },
                        "required": [
                            "id",
                            "title",
                            "preconditions",
                            "steps",
                            "expected_result",
                            "priority",
                            "type",
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

SYSTEM_PROMPT = """You are a senior QA engineer generating test cases from product requirements, PRDs, or Jira tickets.

For each requirement or acceptance criterion you are given, generate test cases covering:
- Happy path / positive scenarios
- Negative scenarios (invalid input, unauthorized access, error handling)
- Edge cases (boundary values, empty states, max/min limits, concurrency where relevant)
- Alternate flows (different user roles, different entry points, different devices/platforms if mentioned)

Guidelines:
- Test steps must be specific and executable by a QA tester with no prior context of the ticket.
- Do not invent requirements that aren't stated or reasonably implied. If something is ambiguous or
  underspecified, generate your best-effort test case AND add a note to `open_questions`.
- Prioritize test cases based on how core the flow is to the feature (High = core happy path /
  security-critical, Medium = common alternate flows, Low = rare edge cases).
- Keep titles short and descriptive (e.g. "User cannot submit form with empty required field").
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


def fetch_jira_ticket(ticket_id: str) -> str:
    """Fetch a Jira ticket's description + acceptance criteria via the Jira REST API.

    Requires environment variables:
        JIRA_BASE_URL   e.g. https://yourcompany.atlassian.net
        JIRA_EMAIL
        JIRA_API_TOKEN
    """
    import requests
    from requests.auth import HTTPBasicAuth

    base_url = os.environ.get("JIRA_BASE_URL")
    email = os.environ.get("JIRA_EMAIL")
    token = os.environ.get("JIRA_API_TOKEN")

    if not all([base_url, email, token]):
        print(
            "Error: JIRA_BASE_URL, JIRA_EMAIL, and JIRA_API_TOKEN must be set in your .env "
            "file to use --jira-ticket.",
            file=sys.stderr,
        )
        sys.exit(1)

    url = f"{base_url}/rest/api/3/issue/{ticket_id}"
    resp = requests.get(url, auth=HTTPBasicAuth(email, token), headers={"Accept": "application/json"})
    resp.raise_for_status()
    data = resp.json()

    fields = data.get("fields", {})
    summary = fields.get("summary", "")
    description = _adf_to_text(fields.get("description"))

    return f"Title: {summary}\n\nDescription / Acceptance Criteria:\n{description}"


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


def generate_test_cases(client: OpenAI, requirement_text: str, few_shot: str) -> dict:
    user_content = requirement_text + few_shot

    response = client.chat.completions.create(
        model=MODEL,
        max_tokens=4096,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_content},
        ],
        tools=[TEST_CASE_TOOL],
        tool_choice={"type": "function", "function": {"name": "submit_test_cases"}},
    )

    return _extract_tool_result(response)


def self_review(client: OpenAI, requirement_text: str, result: dict) -> dict:
    prompt = SELF_REVIEW_PROMPT.format(
        requirement=requirement_text,
        test_cases_json=json.dumps(result.get("test_cases", []), indent=2),
    )

    response = client.chat.completions.create(
        model=MODEL,
        max_tokens=4096,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
        tools=[TEST_CASE_TOOL],
        tool_choice={"type": "function", "function": {"name": "submit_test_cases"}},
    )

    try:
        return _extract_tool_result(response)
    except RuntimeError:
        return result  # fallback: keep original if review pass fails to parse


def write_csv(result: dict, output_path: str):
    rows = []
    for tc in result.get("test_cases", []):
        rows.append(
            {
                "ID": tc.get("id", ""),
                "Title": tc.get("title", ""),
                "Requirement Ref": tc.get("requirement_ref", ""),
                "Preconditions": tc.get("preconditions", ""),
                "Steps": "\n".join(
                    f"{i+1}. {s}" for i, s in enumerate(tc.get("steps", []))
                ),
                "Expected Result": tc.get("expected_result", ""),
                "Priority": tc.get("priority", ""),
                "Type": tc.get("type", ""),
            }
        )
    df = pd.DataFrame(rows)
    df.to_csv(output_path, index=False)
    print(f"\u2714 Wrote {len(rows)} test cases to {output_path}")

    open_qs = result.get("open_questions") or []
    if open_qs:
        print("\n\u26a0 Open questions / ambiguities the model flagged:")
        for q in open_qs:
            print(f"  - {q}")


def main():
    parser = argparse.ArgumentParser(description="Generate QA test cases from requirements via OpenRouter.")
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--input", help="Path to a local .txt/.md file with the PRD/requirements/ticket text.")
    source.add_argument("--jira-ticket", help="Jira ticket ID to fetch, e.g. PROJ-123.")
    parser.add_argument("--output", default="test_cases.csv", help="Output CSV path (default: test_cases.csv)")
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
    else:
        requirement_text = fetch_jira_ticket(args.jira_ticket)

    few_shot = load_few_shot_examples(args.examples)

    print(f"Generating test cases with {MODEL}...")
    result = generate_test_cases(client, requirement_text, few_shot)

    if args.self_review:
        print("Running self-review pass...")
        result = self_review(client, requirement_text, result)

    write_csv(result, args.output)


if __name__ == "__main__":
    main()
