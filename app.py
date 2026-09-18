from __future__ import annotations

import os
import json
import tempfile
import time
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
    extract_jira_id,
    fetch_jira_ticket,
    generate_test_cases,
    load_few_shot_examples,
    write_workbook,
)

OUTPUT_PATH = Path(__file__).with_name("generated_test_cases.xlsx")
SETTINGS_COOKIE = "qa_test_case_generator_settings"
COOKIE_MAX_AGE = 10 * 365 * 24 * 60 * 60

cookie_manager = stx.CookieManager(key="qa-test-case-generator-cookies")
saved_settings = cookie_manager.get(SETTINGS_COOKIE) or {}
if isinstance(saved_settings, str):
    try:
        saved_settings = json.loads(saved_settings)
    except json.JSONDecodeError:
        saved_settings = {}

st.set_page_config(page_title="QA Test Case Generator", page_icon="✓", layout="wide")

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
        st.caption("Settings are saved in this browser until you clear them or clear site cookies.")
        if st.button("Clear saved settings", use_container_width=True):
            cookie_manager.delete(SETTINGS_COOKIE)
            st.rerun()
        model_label = st.selectbox(
            "AI model",
            list(MODEL_OPTIONS),
            index=list(MODEL_OPTIONS).index(
                saved_settings.get("model_label", "NVIDIA Nemotron 3")
                if saved_settings.get("model_label", "NVIDIA Nemotron 3") in MODEL_OPTIONS
                else "NVIDIA Nemotron 3"
            ),
            help="All models are accessed through your OpenRouter API key.",
        )
        selected_model = MODEL_OPTIONS[model_label]
        openrouter_api_key = st.text_input(
            f"OpenRouter API key for {model_label}",
            value=saved_settings.get("openrouter_api_key", ""),
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
            value=saved_settings.get("jira_email", ""),
            placeholder="name@company.com",
        )
        jira_api_token = st.text_input(
            "Jira API token",
            value=saved_settings.get("jira_api_token", ""),
            type="password",
            placeholder="Enter your Jira API token",
            help="Required when using a Jira issue link or key.",
        )
        cookie_manager.set(
            SETTINGS_COOKIE,
            json.dumps(
                {
                    "openrouter_api_key": openrouter_api_key,
                    "jira_base_url": jira_base_url,
                    "jira_email": jira_email,
                    "jira_api_token": jira_api_token,
                    "model_label": model_label,
                }
            ),
            max_age=COOKIE_MAX_AGE,
            same_site="strict",
        )

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
    examples_path = Path(__file__).with_name("examples.json")
    use_examples = st.checkbox("Use team examples", value=examples_path.exists())

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

generate_clicked = st.button("Generate test cases", type="primary", use_container_width=True)

if generate_clicked:
    started_at = time.monotonic()
    if OUTPUT_PATH.exists():
        OUTPUT_PATH.unlink()
    st.session_state.pop("download_bytes", None)
    st.session_state.pop("download_name", None)
    st.session_state.pop("preview_rows", None)

    progress_bar = st.progress(0, text="Starting generation...")
    with st.status("Preparing test-case generation", expanded=True) as generation_status:
        try:
            generation_status.write("Validating the requirement source and metadata...")
            progress_bar.progress(5, text="Validating input...")
            if source == "Jira link":
                if not jira_reference.strip():
                    st.error("Enter a Jira issue link or key.")
                    st.stop()
                jira_id = extract_jira_id(jira_reference)
                generation_status.write(f"Fetching Jira issue {jira_id}...")
                progress_bar.progress(15, text=f"Fetching Jira issue {jira_id}...")
                requirement_text = fetch_jira_ticket(
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
                jira_id = extract_jira_id(jira_id)
                requirement_text = story_text.strip()

            generation_status.write("Preparing text and visual references...")
            progress_bar.progress(20, text="Preparing source material...")

            configured_openrouter_key = openrouter_api_key or os.environ.get("OPENROUTER_API_KEY")
            if not configured_openrouter_key:
                st.error("Enter an OpenRouter API key in Settings.")
                st.stop()

            template_path = None
            temporary_template = None
            if template_upload is not None:
                generation_status.write("Loading the Excel template...")
                temporary_template = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
                temporary_template.write(template_upload.getvalue())
                temporary_template.close()
                template_path = temporary_template.name

            try:
                client = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=configured_openrouter_key)
                few_shot = load_few_shot_examples(str(examples_path)) if use_examples else ""
                figma_images = [(upload.name, upload.getvalue()) for upload in figma_uploads]
                generation_status.write(
                    f"Generating detailed test cases with {model_label}"
                    + (f" using {len(figma_images)} visual reference(s)..." if figma_images else "...")
                )
                progress_bar.progress(30, text="Generating test cases... This may take a moment.")
                result = generate_test_cases(
                    client,
                    requirement_text,
                    few_shot,
                    images=figma_images,
                    model=selected_model,
                )
                generation_status.write(f"Generated {len(result.get('test_cases', []))} scenarios. Building workbook...")
                progress_bar.progress(75, text="Building the Excel workbook...")
                write_workbook(
                    result,
                    str(OUTPUT_PATH),
                    jira_id,
                    template_path=template_path,
                    metadata=metadata,
                )
                generation_status.write("Workbook validated. Preparing preview and download...")
                progress_bar.progress(95, text="Preparing download...")
            finally:
                if temporary_template:
                    Path(temporary_template.name).unlink(missing_ok=True)

            st.session_state["download_bytes"] = OUTPUT_PATH.read_bytes()
            st.session_state["download_name"] = f"{jira_id}_test_cases.xlsx"
            st.session_state["scenario_count"] = len(result.get("test_cases", []))
            st.session_state["preview_rows"] = [
                dict(zip(WORKBOOK_COLUMNS, row))
                for row in _build_workbook_rows(result, jira_id, metadata)
            ]
            elapsed = time.monotonic() - started_at
            progress_bar.progress(100, text=f"Complete in {elapsed:.1f} seconds")
            generation_status.update(label="Test cases ready", state="complete", expanded=False)
            st.success(f"Generated {st.session_state['scenario_count']} test scenarios.")
        except Exception as error:
            generation_status.update(label="Generation failed", state="error", expanded=True)
            progress_bar.empty()
            st.error(str(error))

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
