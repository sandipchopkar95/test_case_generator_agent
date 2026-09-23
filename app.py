from __future__ import annotations

import os
import json
import re
import tempfile
import time
from hmac import compare_digest
from pathlib import Path

import extra_streamlit_components as stx
import streamlit as st

from generate_tests import (
    DEFAULT_METADATA,
    MODEL,
    MODEL_OPTIONS,
    OpenAI,
    WORKBOOK_COLUMNS,
    _build_workbook_rows,
    _fallback_coverage_groups,
    _plan_coverage_groups,
    extract_jira_id,
    fetch_jira_ticket,
    generate_test_cases,
    load_few_shot_examples,
    write_workbook,
)
from learning_memory import (
    LocalLearningRepository,
    build_learning_context,
    clear_learning_records,
    create_learning_repository,
    learning_record_count,
    save_generation,
)
from qa_generator.config import AppConfig
from qa_generator.run_state import RunStateStore

SETTINGS_COOKIE = "qa_test_case_generator_settings"
COOKIE_MAX_AGE = 10 * 365 * 24 * 60 * 60
APP_CONFIG = AppConfig.from_file(__file__)
RUN_STATE_STORE = RunStateStore(APP_CONFIG.run_state_path)

cookie_manager = stx.CookieManager(key="qa-test-case-generator-cookies")
saved_settings = cookie_manager.get(SETTINGS_COOKIE) or {}
if isinstance(saved_settings, str):
    try:
        saved_settings = json.loads(saved_settings)
    except json.JSONDecodeError:
        saved_settings = {}


RUN_STATE_STORE.restore(st.session_state)


def _streamlit_secret(name: str) -> str:
    """Read a deployment secret without requiring it during local development."""
    try:
        return APP_CONFIG.secret(name, st.secrets)
    except (FileNotFoundError, AttributeError):
        return APP_CONFIG.secret(name)


try:
    LEARNING_REPOSITORY = create_learning_repository(
        _streamlit_secret("MONGODB_URI"), _streamlit_secret("MONGODB_DATABASE")
    )
    LEARNING_REPOSITORY_ERROR = ""
except RuntimeError as error:
    LEARNING_REPOSITORY = LocalLearningRepository(APP_CONFIG.local_learning_path)
    LEARNING_REPOSITORY_ERROR = str(error)


def _require_team_access() -> None:
    """Optionally restrict the public deployment before it can use shared memory."""
    expected_password = _streamlit_secret("APP_ACCESS_PASSWORD")
    if not expected_password or st.session_state.get("team_access_granted"):
        return
    st.title("QA Test Case Generator")
    supplied_password = st.text_input("Team access password", type="password")
    if supplied_password:
        if compare_digest(supplied_password, expected_password):
            st.session_state["team_access_granted"] = True
            st.rerun()
        else:
            st.error("Incorrect team access password.")
    st.stop()

st.set_page_config(page_title="QA Test Case Generator", page_icon="✓", layout="wide")
_require_team_access()

st.markdown(
    """
    <style>
    .block-container { max-width: 1180px; padding-top: 2.5rem; }
    [data-testid="stMetricValue"] { color: #155e75; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("QA Test Case Generator")
st.caption("Turn a Jira issue or pasted story into an execution-ready Excel workbook.")

with st.sidebar:
    with st.expander("Settings", expanded=False):
        st.caption("Only non-secret settings are saved in this browser. Enter API keys and tokens for each new session.")
        if st.button("Clear saved settings", use_container_width=True):
            cookie_manager.delete(SETTINGS_COOKIE)
            st.rerun()
        configured_model_label = next(
            (label for label, model_id in MODEL_OPTIONS.items() if model_id == MODEL),
            "Ling 3.0 Flash VL",
        )
        saved_model_label = saved_settings.get("model_label", configured_model_label)
        if saved_model_label not in MODEL_OPTIONS:
            saved_model_label = configured_model_label
        model_label = st.selectbox(
            "AI model",
            list(MODEL_OPTIONS),
            index=list(MODEL_OPTIONS).index(saved_model_label),
            help="All models are accessed through your OpenRouter API key.",
        )
        selected_model = MODEL_OPTIONS[model_label]
        openrouter_api_key = st.text_input(
            f"OpenRouter API key for {model_label}",
            value="",
            type="password",
            placeholder=f"Enter your OpenRouter key for {model_label}",
            help=f"Required to use {model_label} through OpenRouter.",
        )
        jira_base_url = st.text_input(
            "Jira base URL",
            value=saved_settings.get("jira_base_url", os.environ.get("JIRA_BASE_URL", "")),
            placeholder="https://yourcompany.atlassian.net",
        )
        jira_email = st.text_input(
            "Jira email",
            value="",
            placeholder="name@company.com",
        )
        jira_api_token = st.text_input(
            "Jira API token",
            value="",
            type="password",
            placeholder="Enter your Jira API token",
            help="Required when using a Jira issue link or key.",
        )
        settings_json = json.dumps(
            {
                "jira_base_url": jira_base_url,
                "model_label": model_label,
            }
        )
        if st.session_state.get("saved_settings_json") != settings_json:
            cookie_manager.set(
                SETTINGS_COOKIE,
                settings_json,
                max_age=COOKIE_MAX_AGE,
                same_site="strict",
            )
            st.session_state["saved_settings_json"] = settings_json

    st.header("Scenario defaults")
    metadata = {"Status": "Not Run"}
    metadata["Portal"] = st.selectbox(
        "Portal",
        ["Client", "PSP"],
        index=["Client", "PSP"].index(DEFAULT_METADATA["Portal"]),
    )
    metadata["Fix Version"] = st.text_input("Fix Version", value=DEFAULT_METADATA["Fix Version"])
    metadata["Module Name"] = st.text_input("Module Name", value=DEFAULT_METADATA["Module Name"])

    role_options = [
        "CLIENT REP",
        "CLIENT REP (VIEWER)",
        "TENANT REP",
        "CUSTOM ROLE",
        "PSP ADMIN",
        "RESIDENT",
        "STAFF",
        "FRIENDS AND FAMILY",
    ]
    selected_role = st.selectbox("User Role", role_options, index=0)
    if selected_role == "CUSTOM ROLE":
        custom_role = st.text_input("Custom role name", placeholder="Enter the custom role")
        metadata["User Role"] = custom_role.strip() or "CUSTOM ROLE"
    else:
        metadata["User Role"] = selected_role

    metadata["Assigned to"] = st.text_input("Assigned to", value=DEFAULT_METADATA["Assigned to"])
    metadata["Created By"] = st.text_input("Created By", value=DEFAULT_METADATA["Created By"])

    st.divider()
    figma_uploads = st.file_uploader(
        "Figma screenshots or UI references (optional)",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        help="Uploaded images are sent to the model as visual source material for UI test coverage.",
    )
    template_upload = st.file_uploader("Excel template (optional)", type=["xlsx"])
    examples_path = Path(__file__).with_name("sample_test_cases.xlsx")
    use_examples = st.checkbox("Use team examples", value=examples_path.exists())

    st.divider()
    st.header("Past-work learning")
    use_learning = st.checkbox(
        "Use and remember past work",
        value=True,
        help="Successful generations are saved only in this project and relevant past examples guide future work. This does not train the AI provider's model.",
    )
    try:
        learned_count = learning_record_count(LEARNING_REPOSITORY)
        memory_location = "shared cloud memory" if LEARNING_REPOSITORY.is_cloud else "local memory"
        st.caption(f"{learned_count} saved generation{'s' if learned_count != 1 else ''} in {memory_location}.")
    except RuntimeError as error:
        learned_count = 0
        use_learning = False
        st.warning(f"Past-work learning is unavailable: {error}")
    if st.button("Clear past-work memory", use_container_width=True, disabled=not learned_count):
        clear_learning_records(LEARNING_REPOSITORY)
        st.rerun()
    if st.button(
        "Clear current run",
        use_container_width=True,
        disabled=not RUN_STATE_STORE.state_path.exists(),
    ):
        RUN_STATE_STORE.clear(st.session_state)
        st.rerun()

source = st.radio("Requirement source", ["Jira link", "Pasted story"], horizontal=True)

if source == "Jira link":
    jira_reference = st.text_input(
        "Jira issue link or key",
        placeholder="https://yourcompany.atlassian.net/browse/RES-123",
    )
    story_text = ""
    jira_id = ""
else:
    jira_id = st.text_input("Jira ID", placeholder="RES-123")
    story_text = st.text_area(
        "Pasted Jira story",
        height=260,
        placeholder="Paste the story, acceptance criteria, screenshots text, or Figma notes here.",
    )
    jira_reference = ""

st.divider()
st.subheader("Generation steps")
st.caption("Run each stage separately so Jira retrieval and preparation are not repeated when model generation is slow.")

retrieve_clicked = st.button("1. Retrieve Jira ticket and create plan", use_container_width=True)
memory_clicked = st.button(
    "2. Check MongoDB and prepare references",
    disabled="requirement_text" not in st.session_state,
    use_container_width=True,
)
generate_clicked = st.button(
    "3. Generate test cases and build workbook",
    type="primary",
    disabled="memory_ready" not in st.session_state,
    use_container_width=True,
)

if retrieve_clicked:
    try:
        if source == "Jira link":
            if not jira_reference.strip():
                st.error("Enter a Jira issue link or key.")
                st.stop()
            staged_jira_id = extract_jira_id(jira_reference)
            with st.spinner(f"Fetching Jira issue {staged_jira_id}..."):
                staged_requirement = fetch_jira_ticket(
                    jira_reference,
                    jira_base_url=jira_base_url,
                    jira_email=jira_email,
                    jira_api_token=jira_api_token,
                )
        else:
            if not jira_id.strip():
                st.error("Enter the Jira ID used for the generated test-case IDs.")
                st.stop()
            if not story_text.strip():
                st.error("Paste the Jira story or acceptance criteria.")
                st.stop()
            staged_jira_id = extract_jira_id(jira_id)
            staged_requirement = story_text.strip()

        configured_openrouter_key = openrouter_api_key or os.environ.get("OPENROUTER_API_KEY")
        if not configured_openrouter_key:
            st.error("Enter an OpenRouter API key in Settings.")
            st.stop()

        if ":free" in selected_model:
            coverage_groups = _fallback_coverage_groups(staged_requirement)
        else:
            client = OpenAI(
                base_url="https://openrouter.ai/api/v1",
                api_key=configured_openrouter_key,
                timeout=90,
                max_retries=0,
            )
            with st.spinner("Creating the coverage plan..."):
                coverage_groups = _plan_coverage_groups(client, staged_requirement, selected_model)

        st.session_state["jira_id"] = staged_jira_id
        st.session_state["requirement_text"] = staged_requirement
        st.session_state["coverage_groups"] = coverage_groups
        st.session_state.pop("learning_context", None)
        st.session_state.pop("memory_ready", None)
        st.session_state.pop("download_bytes", None)
        st.session_state.pop("output_path", None)
        RUN_STATE_STORE.save(st.session_state)
        st.success(f"Retrieved {staged_jira_id} and created {len(coverage_groups)} coverage groups.")
    except Exception as error:
        st.error(f"Retrieval/planning failed: {error}")

if st.session_state.get("coverage_groups"):
    with st.expander("Coverage plan", expanded=False):
        for group in st.session_state["coverage_groups"]:
            st.markdown(f"**{group['name']}**")
            st.write(group["scope"])

if memory_clicked:
    try:
        requirement_text = st.session_state["requirement_text"]
        few_shot = load_few_shot_examples(str(examples_path)) if use_examples else ""
        learning_context, reference_count = ("", 0)
        if use_learning:
            learning_context, reference_count = build_learning_context(
                requirement_text, LEARNING_REPOSITORY
            )
        st.session_state["few_shot"] = few_shot
        st.session_state["learning_context"] = learning_context
        st.session_state["reference_count"] = reference_count
        st.session_state["memory_ready"] = True
        RUN_STATE_STORE.save(st.session_state)
        memory_location = "shared MongoDB" if LEARNING_REPOSITORY.is_cloud else "local memory"
        if LEARNING_REPOSITORY_ERROR:
            st.warning(f"MongoDB is unavailable, so this run uses local memory: {LEARNING_REPOSITORY_ERROR}")
        st.success(f"References prepared from {memory_location}: {reference_count} relevant record(s).")
    except Exception as error:
        st.error(f"Memory preparation failed: {error}")

if generate_clicked:
    started_at = time.monotonic()
    requirement_text = st.session_state["requirement_text"]
    jira_id = st.session_state["jira_id"]
    configured_openrouter_key = openrouter_api_key or os.environ.get("OPENROUTER_API_KEY")
    if not configured_openrouter_key:
        st.error("Enter an OpenRouter API key in Settings.")
    else:
        try:
            figma_images = [(upload.name, upload.getvalue()) for upload in figma_uploads]
            client = OpenAI(
                base_url="https://openrouter.ai/api/v1",
                api_key=configured_openrouter_key,
                timeout=180,
                max_retries=0,
            )
            reference_count = st.session_state.get("reference_count", 0)
            reference_suffix = f" with {reference_count} relevant past-work reference(s)" if reference_count else ""
            image_suffix = f" and {len(figma_images)} visual reference(s)" if figma_images else ""
            generation_progress = st.progress(0, text="Starting test-case generation...")

            def report_generation_progress(message: str) -> None:
                generation_status.write(message)
                group_match = re.search(r"coverage group (\d+) of (\d+)", message, re.IGNORECASE)
                if group_match:
                    current_group = int(group_match.group(1))
                    total_groups = max(1, int(group_match.group(2)))
                    percent = min(95, round(current_group / total_groups * 90))
                    generation_progress.progress(
                        percent,
                        text=f"Generating coverage group {current_group} of {total_groups}...",
                    )
                elif "Combining coverage groups" in message:
                    generation_progress.progress(94, text="Combining generated coverage...")
                elif "Auditing" in message:
                    generation_progress.progress(96, text="Auditing requirement coverage...")

            with st.status("Generating test cases", expanded=True) as generation_status:
                generation_status.write(f"Generating with {model_label}{reference_suffix}{image_suffix}...")
                result = generate_test_cases(
                    client,
                    requirement_text,
                    st.session_state.get("few_shot", "") + st.session_state.get("learning_context", ""),
                    images=figma_images,
                    model=selected_model,
                    progress_callback=report_generation_progress,
                    coverage_groups=st.session_state.get("coverage_groups"),
                )
                generation_status.write(f"Generated {len(result.get('test_cases', []))} scenarios. Building workbook...")
                template_path = None
                temporary_template = None
                output_path = RUN_STATE_STORE.output_path(f"{jira_id}_test_cases.xlsx")
                if template_upload is not None:
                    temporary_template = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                    temporary_template.write(template_upload.getvalue())
                    temporary_template.close()
                    template_path = temporary_template.name
                try:
                    write_workbook(result, str(output_path), jira_id, template_path=template_path, metadata=metadata)
                    if use_learning:
                        try:
                            save_generation(requirement_text, jira_id, result, LEARNING_REPOSITORY)
                        except RuntimeError as error:
                            generation_status.write(f"Workbook ready, but memory was not updated: {error}")
                    st.session_state["download_bytes"] = output_path.read_bytes()
                finally:
                    if temporary_template:
                        Path(temporary_template.name).unlink(missing_ok=True)
                st.session_state["download_name"] = f"{jira_id}_test_cases.xlsx"
                st.session_state["output_path"] = str(output_path)
                st.session_state["scenario_count"] = len(result.get("test_cases", []))
                st.session_state["preview_rows"] = [
                    dict(zip(WORKBOOK_COLUMNS, row))
                    for row in _build_workbook_rows(result, jira_id, metadata)
                ]
                RUN_STATE_STORE.save(st.session_state)
                generation_progress.progress(100, text="Generation complete")
                generation_status.update(label=f"Test cases ready in {time.monotonic() - started_at:.1f} seconds", state="complete")
            st.success(f"Generated {st.session_state['scenario_count']} test scenarios.")
        except Exception as error:
            st.error(f"Generation failed: {error}")

if st.session_state.get("download_bytes"):
    st.divider()
    st.subheader("Excel preview")
    st.caption("This preview uses the same 21 columns and rows as the downloaded workbook.")
    preview_rows = st.session_state.get("preview_rows", [])
    st.dataframe(
        preview_rows,
        column_order=WORKBOOK_COLUMNS,
        hide_index=True,
        height=560,
        use_container_width=True,
    )
    st.subheader("Generated workbook")
    st.download_button(
        "Download test case workbook",
        data=st.session_state["download_bytes"],
        file_name=st.session_state["download_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
