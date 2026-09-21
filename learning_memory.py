"""Shared retrieval memory for the test-case generator.

The memory is reference retrieval, not model training. MongoDB Atlas is used
when configured; a local JSON file remains available for offline CLI use.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Protocol

DEFAULT_LEARNING_STORE = Path(__file__).with_name(".test_case_learning.json")
MAX_RECORDS = 2_000
MAX_REFERENCES = 8
MAX_CASES_PER_REFERENCE = 4

_STOP_WORDS = {"a", "an", "and", "are", "as", "at", "be", "by", "for", "from", "in", "is", "it", "of", "on", "or", "that", "the", "this", "to", "with", "will"}
_CONCEPTS = {
    "authentication": {"login", "logon", "signin", "password", "session", "mfa", "otp", "reset"},
    "permissions": {"role", "permission", "access", "authorized", "unauthorized", "rbac", "viewer"},
    "notifications": {"notification", "email", "message", "alert", "reminder"},
    "records": {"resident", "tenant", "client", "user", "profile", "record", "owner", "staff"},
    "workflow": {"create", "edit", "update", "delete", "archive", "restore", "submit", "approve"},
    "search": {"search", "filter", "sort", "pagination", "list", "dashboard"},
    "integration": {"api", "integration", "webhook", "sync", "import", "export"},
}


class LearningRepository(Protocol):
    is_cloud: bool

    def load_records(self) -> list[dict]: ...
    def save_generation(self, requirement_text: str, jira_id: str, result: dict) -> None: ...
    def clear(self) -> None: ...


def _tokens(text: str) -> set[str]:
    return {token for token in re.findall(r"[a-z0-9]{3,}", text.casefold()) if token not in _STOP_WORDS}


def _record_key(requirement: str) -> str:
    return hashlib.sha256(re.sub(r"\s+", " ", requirement).strip().casefold().encode("utf-8")).hexdigest()


def _valid_generation(requirement_text: str, result: dict) -> tuple[str, list[dict]] | None:
    requirement = requirement_text.strip()
    cases = result.get("test_cases") if isinstance(result, dict) else None
    return (requirement, cases) if requirement and isinstance(cases, list) and cases else None


def _new_record(key: str, requirement: str, jira_id: str, cases: list[dict]) -> dict:
    return {"requirement_key": key, "saved_at": datetime.now(timezone.utc).isoformat(), "jira_id": jira_id, "requirement": requirement, "test_cases": cases}


class LocalLearningRepository:
    is_cloud = False

    def __init__(self, path: str | Path = DEFAULT_LEARNING_STORE):
        self.path = Path(path)

    def load_records(self) -> list[dict]:
        if not self.path.exists():
            return []
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return []
        records = data.get("records", []) if isinstance(data, dict) else []
        return [record for record in records if isinstance(record, dict)]

    def save_generation(self, requirement_text: str, jira_id: str, result: dict) -> None:
        valid = _valid_generation(requirement_text, result)
        if not valid:
            return
        requirement, cases = valid
        key = _record_key(requirement)
        records = [record for record in self.load_records() if record.get("requirement_key") != key]
        records.append(_new_record(key, requirement, jira_id, cases))
        self.path.write_text(json.dumps({"version": 2, "records": records[-MAX_RECORDS:]}, indent=2), encoding="utf-8")

    def clear(self) -> None:
        if self.path.exists():
            self.path.unlink()


class MongoLearningRepository:
    """Atlas-backed repository. The URI must be held in Streamlit secrets."""

    is_cloud = True

    def __init__(self, uri: str, database: str = "test-case-learning"):
        try:
            from pymongo import MongoClient
        except ImportError as error:
            raise RuntimeError("Install pymongo to use MongoDB shared learning memory.") from error
        self.collection = MongoClient(uri, serverSelectionTimeoutMS=5_000)[database]["learning_records"]

    def load_records(self) -> list[dict]:
        try:
            return list(self.collection.find({}, {"_id": 0}).sort("saved_at", -1).limit(MAX_RECORDS))
        except Exception as error:
            raise RuntimeError("Could not read shared learning memory from MongoDB Atlas.") from error

    def save_generation(self, requirement_text: str, jira_id: str, result: dict) -> None:
        valid = _valid_generation(requirement_text, result)
        if not valid:
            return
        requirement, cases = valid
        record = _new_record(_record_key(requirement), requirement, jira_id, cases)
        try:
            self.collection.update_one({"requirement_key": record["requirement_key"]}, {"$set": record}, upsert=True)
            stale = list(self.collection.find({}, {"_id": 1}).sort("saved_at", -1).skip(MAX_RECORDS))
            if stale:
                self.collection.delete_many({"_id": {"$in": [item["_id"] for item in stale]}})
        except Exception as error:
            raise RuntimeError("Could not save shared learning memory to MongoDB Atlas.") from error

    def clear(self) -> None:
        try:
            self.collection.delete_many({})
        except Exception as error:
            raise RuntimeError("Could not clear shared learning memory from MongoDB Atlas.") from error


def create_learning_repository(mongodb_uri: str | None = None, mongodb_database: str | None = None, local_path: str | Path = DEFAULT_LEARNING_STORE) -> LearningRepository:
    uri = mongodb_uri or os.environ.get("MONGODB_URI")
    if uri:
        return MongoLearningRepository(uri, mongodb_database or os.environ.get("MONGODB_DATABASE", "test-case-learning"))
    return LocalLearningRepository(local_path)


def learning_record_count(repository: LearningRepository) -> int:
    return len(repository.load_records())


def clear_learning_records(repository: LearningRepository) -> None:
    repository.clear()


def save_generation(requirement_text: str, jira_id: str, result: dict, repository: LearningRepository) -> None:
    repository.save_generation(requirement_text, jira_id, result)


def _concepts(tokens: set[str]) -> set[str]:
    return {name for name, terms in _CONCEPTS.items() if tokens & terms}


def _referenced_jira_ids(text: str) -> set[str]:
    return set(re.findall(r"\b[A-Z][A-Z0-9]+-\d+\b", text.upper()))


def _record_search_text(record: dict) -> str:
    scenarios = " ".join(str(case.get("scenario", "")) for case in record.get("test_cases", []) if isinstance(case, dict))
    return f"{record.get('requirement', '')} {scenarios}"


def build_learning_context(requirement_text: str, repository: LearningRepository, limit: int = MAX_REFERENCES) -> tuple[str, int]:
    """Retrieve functional/dependency references, using only materially relevant records."""
    query_tokens = _tokens(requirement_text)
    query_concepts = _concepts(query_tokens)
    query_ids = _referenced_jira_ids(requirement_text)
    scored = []
    for record in repository.load_records():
        record_tokens = _tokens(_record_search_text(record))
        shared_terms = query_tokens & record_tokens
        shared_concepts = query_concepts & _concepts(record_tokens)
        jira_dependency = str(record.get("jira_id", "")).upper() in query_ids
        # One incidental term is not enough: require two terms, a shared domain
        # concept plus a term, or an explicit Jira dependency.
        if not (jira_dependency or len(shared_terms) >= 2 or (shared_concepts and shared_terms)):
            continue
        score = len(shared_terms) * 3 + len(shared_concepts) * 4 + (30 if jira_dependency else 0)
        scored.append((score, str(record.get("saved_at", "")), record))
    selected = [record for _, _, record in sorted(scored, key=lambda item: (item[0], item[1]), reverse=True)[:limit]]
    references = [{"past_jira_id": record.get("jira_id", ""), "past_requirement": str(record.get("requirement", ""))[:1600], "past_test_cases": record.get("test_cases", [])[:MAX_CASES_PER_REFERENCE]} for record in selected if isinstance(record.get("test_cases"), list)]
    if not references:
        return "", 0
    return ("\n\nRelevant past-work references. They may describe dependencies or related functionality. Use their coverage patterns and details only when supported by the current story or an explicit dependency; never copy unrelated behavior:\n" + json.dumps(references, ensure_ascii=False, indent=2), len(references))
