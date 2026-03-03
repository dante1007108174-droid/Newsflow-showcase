"""Run the subscribe_using workflow suite and write a report.

Usage:
  python tests/run_subscribe_using_suite.py

Outputs:
  - tests/reports/subscribe_using_report.json
  - tests/reports/subscribe_using_report.md
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from coze_workflow_client import CozeWorkflowClient


ROOT = Path(__file__).resolve().parents[1]
SUITE_PATH = Path(__file__).resolve().parent / "subscribe_using_suite.json"
REPORT_DIR = Path(__file__).resolve().parent / "reports"


def _safe_print(s: str) -> None:
    try:
        print(s)
    except UnicodeEncodeError:
        # Fallback for Windows console encoding issues.
        print(s.encode("utf-8", errors="replace").decode("utf-8", errors="replace"))


def _now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _rand_token(n: int = 8) -> str:
    import random
    import string

    alphabet = string.ascii_lowercase + string.digits
    return "".join(random.choice(alphabet) for _ in range(n))


def _make_context() -> Dict[str, str]:
    tag = datetime.now().strftime("%Y%m%d_%H%M%S")
    email = f"autotest+{tag}_{_rand_token()}@example.com"
    email2 = f"autotest+{tag}_{_rand_token()}@example.com"
    user_id = f"autotest_subscribe_{tag}_{_rand_token()}"
    return {
        "user_id": user_id,
        "email": email,
        "email2": email2,
        "keyword": "AI",
    }


def _render_template(value: str, ctx: Dict[str, str]) -> str:
    if not isinstance(value, str):
        return value
    out = value
    for k, v in ctx.items():
        out = out.replace("{{" + k + "}}", v)
    return out


def _extract_field(text: str, label: str) -> Optional[str]:
    # Matches: "**邮箱**: xxx" and captures until newline.
    m = re.search(rf"\*\*{re.escape(label)}\*\*:\s*([^\n]+)", text)
    return m.group(1).strip() if m else None


@dataclass
class StepResult:
    step_index: int
    action: str
    params: Dict[str, Any]
    output: str
    passed: bool
    errors: List[str]
    duration_ms: int


def _assert_expectations(output: str, expect: Dict[str, Any], ctx: Dict[str, str]) -> List[str]:
    errors: List[str] = []

    if expect.get("non_empty") is True:
        if not output.strip():
            errors.append("output is empty")

    any_contains = expect.get("any_contains")
    if isinstance(any_contains, list) and any_contains:
        if not any((s in output) for s in any_contains if isinstance(s, str)):
            errors.append(f"output does not contain any of: {any_contains}")

    any_contains_all = expect.get("any_contains_all")
    if isinstance(any_contains_all, list) and any_contains_all:
        missing = [s for s in any_contains_all if isinstance(s, str) and s not in output]
        if missing:
            errors.append(f"output missing required substrings: {missing}")

    none_contains = expect.get("none_contains")
    if isinstance(none_contains, list) and none_contains:
        present = [s for s in none_contains if isinstance(s, str) and s in output]
        if present:
            errors.append(f"output contains forbidden substrings: {present}")

    must_contain_vars = expect.get("must_contain_vars")
    if isinstance(must_contain_vars, list) and must_contain_vars:
        for var_name in must_contain_vars:
            if not isinstance(var_name, str):
                continue
            val = ctx.get(var_name)
            if val and val not in output:
                errors.append(f"output missing variable value for {var_name}: {val!r}")

    # Optional structured checks for query results.
    if expect.get("must_have_fields"):
        # Example: ["邮箱", "主题"]
        for label in expect["must_have_fields"]:
            extracted = _extract_field(output, label)
            if extracted is None:
                errors.append(f"missing formatted field: {label}")

    return errors


def _call_workflow(
    client: CozeWorkflowClient,
    workflow_id: str,
    ctx: Dict[str, str],
    step: Dict[str, Any],
) -> Tuple[str, Dict[str, Any], int]:
    action = step.get("action")
    if not isinstance(action, str) or not action:
        raise ValueError("step.action must be a non-empty string")

    params = {
        "action": action,
        "input_user_id": ctx["user_id"],
        "ex_email": "",
        "new_email": "",
        "new_keyword": "",
    }
    for k in ("ex_email", "new_email", "new_keyword"):
        if k in step:
            params[k] = _render_template(str(step.get(k) or ""), ctx)

    # Coze occasionally returns a 200 with an empty/unknown payload (no output, no execute_id).
    # For automation stability, we retry a couple times on that specific condition.
    attempts = 3
    last_err: Optional[Exception] = None
    for attempt in range(1, attempts + 1):
        try:
            t0 = time.time()
            res = client.run_workflow(workflow_id, params, is_async=False, timeout_s=90)
            dt_ms = int((time.time() - t0) * 1000)
            return res.output or "", params, dt_ms
        except RuntimeError as e:
            last_err = e
            msg = str(e)
            if "did not return output or execute_id" in msg and attempt < attempts:
                time.sleep(0.5 * attempt)
                continue
            raise

    # Unreachable, but keeps type checkers happy.
    raise last_err  # type: ignore[misc]


def main() -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    if not SUITE_PATH.exists():
        raise SystemExit(f"Suite file not found: {SUITE_PATH}")

    suite = json.loads(SUITE_PATH.read_text(encoding="utf-8"))
    suite_meta = suite.get("suite") or {}
    workflow_id = str(suite_meta.get("workflow_id") or "").strip()
    if not workflow_id:
        raise SystemExit("suite.suite.workflow_id is required")

    cases = suite.get("cases")
    if not isinstance(cases, list) or not cases:
        raise SystemExit("suite.cases must be a non-empty list")

    client = CozeWorkflowClient.from_env()

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Write both a timestamped artifact and a stable "latest" artifact.
    report_json_path = REPORT_DIR / f"subscribe_using_report_{run_id}.json"
    report_md_path = REPORT_DIR / f"subscribe_using_report_{run_id}.md"
    report_json_latest_path = REPORT_DIR / "subscribe_using_report.json"
    report_md_latest_path = REPORT_DIR / "subscribe_using_report.md"

    report: Dict[str, Any] = {
        "suite": suite_meta,
        "run": {
            "run_id": run_id,
            "timestamp": _now_ts(),
            "base_url": os.getenv("COZE_BASE_URL") or os.getenv("COZE_API_BASE") or "https://api.coze.cn",
        },
        "results": [],
        "summary": {},
    }

    passed_cases = 0
    failed_cases = 0
    known_failed_cases = 0

    _safe_print(f"Running suite subscribe_using ({len(cases)} cases) at {_now_ts()}")

    for case in cases:
        case_id = str(case.get("id") or "")
        case_name = str(case.get("name") or "")
        case_name_zh = str(case.get("name_zh") or "")
        purpose = str(case.get("purpose") or "")
        known_issue = bool(case.get("known_issue"))
        issue_status = str(case.get("issue_status") or "")
        issue_note = str(case.get("issue_note") or "")
        steps = case.get("steps")
        if not isinstance(steps, list) or not steps:
            continue

        ctx = _make_context()
        case_result: Dict[str, Any] = {
            "id": case_id,
            "name": case_name,
            "name_zh": case_name_zh,
            "purpose": purpose,
            "known_issue": known_issue,
            "issue_status": issue_status,
            "issue_note": issue_note,
            "context": {
                # Keep context in report for debugging; it contains only synthetic values.
                "user_id": ctx["user_id"],
                "email": ctx["email"],
                "email2": ctx["email2"],
            },
            "steps": [],
            "passed": True,
        }

        for idx, step in enumerate(steps):
            expect = step.get("expect") or {}
            try:
                out, params, dt_ms = _call_workflow(client, workflow_id, ctx, step)
                errors = _assert_expectations(out, expect, ctx)
                passed = len(errors) == 0
            except Exception as e:
                out = ""
                params = {
                    "action": step.get("action"),
                    "input_user_id": ctx.get("user_id"),
                }
                dt_ms = 0
                errors = [f"exception: {type(e).__name__}: {e}"]
                passed = False

            sr = StepResult(
                step_index=idx,
                action=str(step.get("action") or ""),
                params=params,
                output=out,
                passed=passed,
                errors=errors,
                duration_ms=dt_ms,
            )
            case_result["steps"].append(sr.__dict__)

            if not passed:
                case_result["passed"] = False
                case_result["passed"] = False

        report["results"].append(case_result)

        # Classify failure type for reporting.
        is_known_fail = (not case_result["passed"]) and known_issue
        case_result["known_fail"] = bool(is_known_fail)

        if case_result["passed"]:
            passed_cases += 1
            _safe_print(f"PASS {case_id} {case_name}")
        elif is_known_fail:
            known_failed_cases += 1
            _safe_print(f"KNOWN_FAIL {case_id} {case_name}")
        else:
            failed_cases += 1
            _safe_print(f"FAIL {case_id} {case_name}")

    report["summary"] = {
        "total_cases": passed_cases + failed_cases + known_failed_cases,
        "passed_cases": passed_cases,
        # Keep backward-compatible meaning: failed_cases == unexpected failures.
        "failed_cases": failed_cases,
        "known_failed_cases": known_failed_cases,
    }

    report_json_content = json.dumps(report, ensure_ascii=False, indent=2)
    report_json_path.write_text(report_json_content, encoding="utf-8")
    report_json_latest_path.write_text(report_json_content, encoding="utf-8")

    # Build a compact markdown report.
    lines: List[str] = []
    lines.append(f"# subscribe_using workflow test report")
    lines.append("")
    lines.append(f"- Run ID: `{run_id}`")
    lines.append(f"- Time: `{report['run']['timestamp']}`")
    lines.append(f"- Workflow ID: `{workflow_id}`")
    lines.append(f"- Total: {report['summary']['total_cases']}")
    lines.append(f"- Passed: {passed_cases}")
    lines.append(f"- Failed: {failed_cases}")
    lines.append(f"- Known Failed: {known_failed_cases}")
    lines.append("")

    if known_failed_cases:
        lines.append("## Known Issues")
        lines.append("")
        for r in report["results"]:
            if not r.get("known_fail"):
                continue
            lines.append(f"### {r.get('id')} {r.get('name')}")
            if r.get("issue_status") or r.get("issue_note"):
                lines.append("")
                if r.get("issue_status"):
                    lines.append(f"- issue_status: {r.get('issue_status')}")
                if r.get("issue_note"):
                    lines.append(f"- issue_note: {r.get('issue_note')}")
            lines.append("")
            for s in r.get("steps", []):
                if s.get("passed"):
                    continue
                lines.append(f"- Step {s.get('step_index')} action={s.get('action')}")
                for err in s.get("errors", []):
                    lines.append(f"  - {err}")
            lines.append("")

    if failed_cases:
        lines.append("## Failures")
        lines.append("")
        for r in report["results"]:
            if r.get("passed") or r.get("known_fail"):
                continue
            lines.append(f"### {r.get('id')} {r.get('name')}")
            lines.append("")
            for s in r.get("steps", []):
                if s.get("passed"):
                    continue
                lines.append(f"- Step {s.get('step_index')} action={s.get('action')}")
                for err in s.get("errors", []):
                    lines.append(f"  - {err}")
            lines.append("")

    lines.append("## Artifacts")
    lines.append("")
    lines.append(f"- JSON: `{report_json_path}`")
    lines.append(f"- MD: `{report_md_path}`")
    lines.append("")

    report_md_content = "\n".join(lines)
    report_md_path.write_text(report_md_content, encoding="utf-8")
    report_md_latest_path.write_text(report_md_content, encoding="utf-8")

    _safe_print(f"Wrote report: {report_json_path}")
    _safe_print(f"Wrote report: {report_md_path}")
    _safe_print(f"Wrote report: {report_json_latest_path}")
    _safe_print(f"Wrote report: {report_md_latest_path}")

    # Only unexpected failures should fail the run.
    return 0 if failed_cases == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
