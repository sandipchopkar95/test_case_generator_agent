"""Configuration and filesystem locations for the application."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class AppConfig:
    """Non-secret application configuration resolved from the environment."""

    root: Path
    openrouter_base_url: str = "https://openrouter.ai/api/v1"
    default_learning_database: str = "test-case-learning"
    run_directory_name: str = ".test_case_run"

    @classmethod
    def from_file(cls, file_path: str | Path) -> "AppConfig":
        return cls(root=Path(file_path).resolve().parent)

    @property
    def run_directory(self) -> Path:
        return self.root / self.run_directory_name

    @property
    def run_state_path(self) -> Path:
        return self.run_directory / "state.json"

    @property
    def local_learning_path(self) -> Path:
        return self.root / ".test_case_learning.json"

    @property
    def examples_path(self) -> Path:
        return self.root / "sample_test_cases.xlsx"

    def secret(self, name: str, streamlit_secrets: object | None = None) -> str:
        """Resolve a secret without making local development require Streamlit."""
        if streamlit_secrets is not None:
            try:
                value = streamlit_secrets.get(name)
                if value:
                    return str(value)
            except (AttributeError, FileNotFoundError):
                pass
        return os.environ.get(name, "")
