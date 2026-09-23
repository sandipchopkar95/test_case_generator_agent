"""Durable, non-secret state for a resumable generation run."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, MutableMapping


STATE_KEYS = (
    "jira_id",
    "requirement_text",
    "coverage_groups",
    "few_shot",
    "learning_context",
    "reference_count",
    "memory_ready",
    "download_name",
    "output_path",
    "scenario_count",
    "preview_rows",
)


class RunStateStore:
    """Persist and restore workflow state without storing API credentials."""

    def __init__(self, state_path: str | Path):
        self.state_path = Path(state_path)
        self.output_directory = self.state_path.parent

    def save(self, session_state: MutableMapping[str, Any]) -> None:
        state = {key: session_state[key] for key in STATE_KEYS if key in session_state}
        self.output_directory.mkdir(parents=True, exist_ok=True)
        temporary_path = self.state_path.with_suffix(".tmp")
        temporary_path.write_text(json.dumps(state), encoding="utf-8")
        temporary_path.replace(self.state_path)

    def restore(self, session_state: MutableMapping[str, Any]) -> bool:
        if session_state.get("run_state_restored") or not self.state_path.exists():
            return False
        try:
            state = json.loads(self.state_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            state = {}
        if isinstance(state, dict):
            session_state.update({key: state[key] for key in STATE_KEYS if key in state})
            output_path = state.get("output_path")
            if output_path and self._is_owned_output(output_path):
                output_file = Path(output_path)
                if output_file.exists():
                    session_state["download_bytes"] = output_file.read_bytes()
        session_state["run_state_restored"] = True
        return bool(state)

    def clear(self, session_state: MutableMapping[str, Any]) -> None:
        output_path = session_state.get("output_path")
        if output_path and self._is_owned_output(output_path):
            Path(output_path).unlink(missing_ok=True)
        self.state_path.unlink(missing_ok=True)
        for key in (*STATE_KEYS, "download_bytes"):
            session_state.pop(key, None)

    def output_path(self, file_name: str) -> Path:
        safe_name = Path(file_name).name
        self.output_directory.mkdir(parents=True, exist_ok=True)
        return self.output_directory / safe_name

    def _is_owned_output(self, output_path: str | Path) -> bool:
        try:
            Path(output_path).resolve().relative_to(self.output_directory.resolve())
            return True
        except ValueError:
            return False
