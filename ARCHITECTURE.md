# Architecture

The project has two entry points:

- `app.py` is the Streamlit presentation layer.
- `generate_tests.py` remains the backwards-compatible script entry point.

## Package boundaries

`qa_generator/config.py` owns filesystem locations and environment-backed configuration.

`qa_generator/run_state.py` owns atomic persistence and recovery of non-secret workflow state. It validates that restored workbook paths remain inside `.test_case_run/`.

`qa_generator/cli.py` owns command-line parsing and orchestration. The generation engine is imported only when the CLI runs, so `generate_tests.py` can delegate to it without a circular import.

`generate_tests.py` owns the generation domain for now: Jira retrieval, OpenRouter calls, result validation, coverage planning, and workbook export. Its pure functions are covered by `tests/test_generation_contracts.py` and remain import-compatible for the Streamlit UI.

`learning_memory.py` owns local and MongoDB-backed reference retrieval. Its repository behavior is covered by `tests/test_learning_memory.py`.

## Local commands

```bash
streamlit run app.py
python generate_tests.py --input sample_input.txt --jira-id SEN-25136 --output test_cases.xlsx
python -m pytest -q
```

The app stores resumable local artifacts under `.test_case_run/`, which is ignored by Git. Secrets are never written to run state. Cloud deployments should use durable object storage or a job queue if generation must survive an instance restart.
