"""Command-line orchestration for test-case generation."""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from openai import OpenAI

from learning_memory import (
    DEFAULT_LEARNING_STORE,
    build_learning_context,
    create_learning_repository,
    save_generation,
)


def run() -> None:
    """Parse CLI arguments and execute one generation run."""
    from generate_tests import (
        MODEL,
        OPENROUTER_BASE_URL,
        extract_jira_id,
        fetch_jira_ticket,
        generate_test_cases,
        load_few_shot_examples,
        self_review,
        write_workbook,
    )

    parser = argparse.ArgumentParser(description="Generate QA test cases from requirements via OpenRouter.")
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--input", help="Path to a local .txt/.md file with the PRD/requirements/ticket text.")
    source.add_argument("--story", help="Pasted Jira story or business requirement text.")
    source.add_argument("--jira-ticket", help="Jira ticket ID to fetch, e.g. PROJ-123.")
    parser.add_argument("--jira-id", help="Jira ID used in generated IDs, e.g. RES-123.")
    parser.add_argument("--template", help="Optional .xlsx template with the required 21 columns.")
    parser.add_argument("--output", default="test_cases.xlsx", help="Output workbook path (default: test_cases.xlsx)")
    parser.add_argument("--examples", help="Optional path to a JSON or Excel file of few-shot examples.")
    parser.add_argument(
        "--learning-store",
        default=str(DEFAULT_LEARNING_STORE),
        help="Local JSON memory for offline use when MongoDB is not configured.",
    )
    parser.add_argument("--mongodb-uri", help="MongoDB Atlas URI for shared team learning memory.")
    parser.add_argument("--mongodb-database", help="MongoDB database name (default: test-case-learning).")
    parser.add_argument("--no-learning", action="store_true", help="Do not use or save past-work references for this run.")
    parser.add_argument("--self-review", action="store_true", help="Run a second pass to improve coverage.")
    args = parser.parse_args()

    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        print(
            "Error: OPENROUTER_API_KEY environment variable not set.\n"
            "Set it in a .env file (see .env.example) or with `export OPENROUTER_API_KEY=your-key`.",
            file=sys.stderr,
        )
        raise SystemExit(1)

    client = OpenAI(base_url=OPENROUTER_BASE_URL, api_key=api_key, timeout=180, max_retries=0)
    if args.input:
        requirement_text = Path(args.input).read_text(encoding="utf-8")
        jira_id = args.jira_id or "REQ"
    elif args.story:
        requirement_text = args.story
        jira_id = args.jira_id or "REQ"
    else:
        requirement_text = fetch_jira_ticket(args.jira_ticket)
        jira_id = extract_jira_id(args.jira_ticket)

    few_shot = load_few_shot_examples(args.examples)
    learning_context = ""
    reference_count = 0
    learning_repository = None
    if not args.no_learning:
        learning_repository = create_learning_repository(
            args.mongodb_uri, args.mongodb_database, args.learning_store
        )
        learning_context, reference_count = build_learning_context(requirement_text, learning_repository)
    if reference_count:
        print(f"Using {reference_count} relevant past-work reference(s).")

    print(f"Generating test cases with {MODEL}...")
    result = generate_test_cases(client, requirement_text, few_shot + learning_context)
    if args.self_review:
        print("Running self-review pass...")
        result = self_review(client, requirement_text, result)

    write_workbook(result, args.output, jira_id, args.template)
    if learning_repository is not None:
        save_generation(requirement_text, jira_id, result, learning_repository)
        location = "shared cloud" if learning_repository.is_cloud else "local"
        print(f"Saved this successful generation to {location} past-work memory.")
