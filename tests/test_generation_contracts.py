import pytest

from generate_tests import (
    _build_workbook_rows,
    _merge_generated_batches,
    _split_requirement_for_generation,
    extract_jira_id,
    validate_generated_result,
)


def test_extract_jira_id_accepts_key_and_url():
    assert extract_jira_id("sen-25136") == "SEN-25136"
    assert extract_jira_id("https://example.atlassian.net/browse/RES-123") == "RES-123"


def test_extract_jira_id_rejects_missing_key():
    with pytest.raises(ValueError):
        extract_jira_id("not-a-jira-reference")


def test_split_requirement_preserves_title():
    chunks = _split_requirement_for_generation("Title: Example\n\n" + "a" * 12, max_chars=8)
    assert len(chunks) > 1
    assert all(chunk.startswith("Title: Example") for chunk in chunks)


def test_validation_normalises_missing_classification_defaults():
    result = {
        "test_cases": [
            {
                "scenario": "Verify search",
                "description": "Search returns matching records.",
                "preconditions": "Records exist.",
                "steps": [
                    {
                        "name": "Enter Search Term",
                        "instruction": "Enter a resident name.",
                        "expected_result": "The matching name is accepted in the search field.",
                    }
                ],
            }
        ]
    }

    validated = validate_generated_result(result)
    assert validated["test_cases"][0]["priority"] == "Medium"
    assert validated["test_cases"][0]["test_type"] == "Functional"


def test_merge_generated_batches_deduplicates_scenarios_and_questions():
    case = {
        "scenario": "Verify search",
        "description": "Search works.",
        "preconditions": "Records exist.",
        "steps": [{"name": "Search Records", "instruction": "Search.", "expected_result": "Results appear."}],
        "priority": "Medium",
        "test_type": "Functional",
        "is_negative_case": "No",
        "automation_candidate": "Yes",
    }
    merged = _merge_generated_batches(
        [
            {"test_cases": [case], "open_questions": ["Question"]},
            {"test_cases": [case], "open_questions": ["question"]},
        ]
    )
    assert len(merged["test_cases"]) == 1
    assert merged["open_questions"] == ["Question"]


def test_workbook_rows_keep_metadata_only_on_first_step():
    result = {
        "test_cases": [
            {
                "scenario": "Verify save",
                "description": "Save persists changes.",
                "preconditions": "User can edit.",
                "steps": [
                    {"name": "Edit Record", "instruction": "Edit the record.", "expected_result": "The edit form opens."},
                    {"name": "Save Record", "instruction": "Save the record.", "expected_result": "The changes persist."},
                ],
                "priority": "High",
                "test_type": "Functional",
                "is_negative_case": "No",
                "automation_candidate": "Yes",
            }
        ]
    }
    rows = _build_workbook_rows(result, "SEN-1")
    assert len(rows) == 2
    assert rows[0][1] == "Verify save"
    assert rows[1][1] == ""
    assert rows[0][8] == "Not Run"
    assert rows[1][8] == ""
