from pathlib import Path

from qa_generator.run_state import RunStateStore


def test_run_state_round_trip_restores_workbook(tmp_path: Path):
    store = RunStateStore(tmp_path / ".test_case_run" / "state.json")
    output_path = store.output_path("SEN-1_test_cases.xlsx")
    output_path.write_bytes(b"workbook")
    original = {
        "jira_id": "SEN-1",
        "requirement_text": "A requirement",
        "coverage_groups": [{"name": "Group", "scope": "Scope"}],
        "memory_ready": True,
        "output_path": str(output_path),
        "download_name": output_path.name,
    }

    store.save(original)
    restored = {}
    assert store.restore(restored) is True
    assert restored["jira_id"] == "SEN-1"
    assert restored["memory_ready"] is True
    assert restored["download_bytes"] == b"workbook"


def test_run_state_rejects_output_outside_run_directory(tmp_path: Path):
    store = RunStateStore(tmp_path / ".test_case_run" / "state.json")
    outside = tmp_path / "outside.xlsx"
    outside.write_bytes(b"secret")
    store.save({"output_path": str(outside)})

    restored = {}
    store.restore(restored)
    assert "download_bytes" not in restored


def test_clear_removes_owned_artifacts(tmp_path: Path):
    store = RunStateStore(tmp_path / ".test_case_run" / "state.json")
    output_path = store.output_path("output.xlsx")
    output_path.write_bytes(b"data")
    state = {"output_path": str(output_path), "jira_id": "SEN-1"}
    store.save(state)

    store.clear(state)
    assert not output_path.exists()
    assert not store.state_path.exists()
    assert "jira_id" not in state
