# Workflow Automation Tests

This folder contains small scripts to run Coze workflows from your local machine.

## Prereqs

- Python 3 installed
- `requests` and `python-dotenv` installed:

```bash
pip install requests python-dotenv
```

## Config

The scripts load environment variables from `backend/.env` (if present).

Required:

- `COZE_API_TOKEN`

Optional:

- `COZE_BASE_URL` (defaults to `https://api.coze.cn`)

## Run subscribe_using tests

```bash
python tests/test_subscribe_using.py
```

## Run subscribe_using suite (dataset + report)

```bash
python tests/run_subscribe_using_suite.py
```

Artifacts are written to:

- `tests/reports/subscribe_using_report.json`
- `tests/reports/subscribe_using_report.md`

## Run instant news generation-node suite

This suite targets the *test workflow* that contains only the core LLM node for
instant news generation (dedup/cluster/sort/format).

```bash
python tests/run_instant_news_suite.py
```

Artifacts are written to:

- `tests/reports/instant_news_report.json`
- `tests/reports/instant_news_report.md`
- `tests/reports/instant_news_test_data.xlsx`

Expected output:

- `PASS ...` lines
- exit code 0 on success, 1 on failure

## Run dispatch model suite (chat-level)

This suite tests the dispatch model via Coze Chat API (intent routing + parameter
extraction), rather than calling workflows directly.

```bash
python tests/run_dispatch_suite.py
```

Required setup:

- `tests/dispatch_suite.json` (already included)
- A test bot with STUB workflows for:
  - `get_instant_news_using`
  - `subscribe_using`
  - `send_mail_using`

Environment variables:

- `COZE_API_TOKEN` (required)
- `COZE_DISPATCH_TEST_BOT_ID` (optional, overrides bot_id in suite)
- `DISPATCH_USE_STREAM` (optional, default `true`, used for first-response latency)

Artifacts are written to:

- `tests/reports/dispatch_model/dispatch_report.json`
- `tests/reports/dispatch_model/dispatch_report.md`
- `tests/reports/dispatch_model/dispatch_test_data.xlsx`

## Run agent E2E suite (chat + real workflows)

This suite validates the production bot end-to-end with real workflows.
It covers intent routing, parameter extraction, workflow execution signals,
response quality scoring, and interaction quality scoring.

```bash
python tests/run_agent_e2e_suite.py
```

Required setup:

- `tests/agent_e2e_suite.json`
- Production/test bot with real workflows attached

Environment variables:

- `COZE_API_TOKEN` (required)
- `COZE_AGENT_E2E_BOT_ID` (optional, overrides bot_id in suite)
- `AGENT_E2E_USE_STREAM` (optional, default `true`)
- `AGENT_E2E_START_INDEX` (optional, 1-based start index)
- `AGENT_E2E_MAX_CASES` (optional, run subset for targeted rerun)

Artifacts are written to:

- `tests/reports/agent_e2e/agent_e2e_report.json`
- `tests/reports/agent_e2e/agent_e2e_report.md`
- `tests/reports/agent_e2e/agent_e2e_test_data.xlsx`

Excel sheets:

- `测试汇总`
- `测试明细`
- `验收标准`
- `失败案例分析`
- `迭代追踪`

## Notes

- These tests call Coze's workflow API; your machine must have internet access.
- The `subscribe_using` workflow has no LLM node, but it still calls Supabase via HTTP nodes.
- The test user_id/email are randomized to avoid colliding with real users.
- The agent E2E suite uses real workflow side effects. Confirm run scope with owner before full runs.
