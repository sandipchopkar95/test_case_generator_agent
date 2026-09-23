from pathlib import Path

from learning_memory import LocalLearningRepository, build_learning_context


def test_local_learning_replaces_duplicate_requirement(tmp_path: Path):
    repository = LocalLearningRepository(tmp_path / "learning.json")
    result = {"test_cases": [{"scenario": "Verify resident search"}]}

    repository.save_generation("Resident search", "SEN-1", result)
    repository.save_generation("  resident   search ", "SEN-2", result)

    records = repository.load_records()
    assert len(records) == 1
    assert records[0]["jira_id"] == "SEN-2"


def test_learning_context_requires_material_relevance(tmp_path: Path):
    repository = LocalLearningRepository(tmp_path / "learning.json")
    repository.save_generation(
        "Resident search and filtering",
        "SEN-1",
        {"test_cases": [{"scenario": "Verify resident filter"}]},
    )

    context, count = build_learning_context("Unrelated billing export", repository)
    assert context == ""
    assert count == 0
