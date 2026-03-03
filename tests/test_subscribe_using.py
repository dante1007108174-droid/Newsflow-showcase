"""Automated smoke tests for the Coze subscribe_using workflow.

This hits Coze's workflow API directly, so inputs/outputs are fully controlled.
It does NOT require a model node; the only costs are network + backend services.
"""

from __future__ import annotations

import random
import re
import string
from datetime import datetime
import sys

from coze_workflow_client import CozeWorkflowClient


WORKFLOW_ID_SUBSCRIBE_USING = "7599988261473927214"


def _rand_suffix(n: int = 8) -> str:
    alphabet = string.ascii_lowercase + string.digits
    return "".join(random.choice(alphabet) for _ in range(n))


def _now_tag() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _make_user_id() -> str:
    return f"autotest_subscribe_{_now_tag()}_{_rand_suffix()}"


def _make_email() -> str:
    return f"autotest+{_now_tag()}_{_rand_suffix()}@example.com"


def _run(client: CozeWorkflowClient, *, action: str, input_user_id: str, ex_email: str = "", new_email: str = "", new_keyword: str = "") -> str:
    params = {
        "action": action,
        "input_user_id": input_user_id,
        "ex_email": ex_email,
        "new_email": new_email,
        "new_keyword": new_keyword,
    }
    res = client.run_workflow(WORKFLOW_ID_SUBSCRIBE_USING, params, is_async=False, timeout_s=60)
    return res.output or ""


def _assert_contains(haystack: str, needle: str, context: str) -> None:
    if needle not in haystack:
        raise AssertionError(f"Expected output to contain {needle!r} ({context}).\nActual: {haystack}")


def _extract_field(text: str, label: str) -> str | None:
    # Matches lines like: "📧 **邮箱**: xxx"
    m = re.search(rf"\*\*{re.escape(label)}\*\*:\s*([^\n]+)", text)
    return m.group(1).strip() if m else None

def test_update_then_query_then_unsubscribe(client: CozeWorkflowClient) -> None:
    user_id = _make_user_id()
    email = _make_email()
    keyword = "AI"

    # Update (create or modify)
    out_update = _run(
        client,
        action="修改订阅",
        input_user_id=user_id,
        new_email=email,
        new_keyword=keyword,
    )
    # We don't hard-assert exact phrasing; just ensure it's not empty.
    if not out_update.strip():
        raise AssertionError("Expected non-empty output for update_subscription")

    # Query
    out_query = _run(client, action="查询订阅", input_user_id=user_id)
    if not out_query.strip():
        raise AssertionError("Expected non-empty output for query_subscription")

    # Validate returned fields (the workflow formats this in Markdown)
    extracted_email = _extract_field(out_query, "邮箱")
    extracted_keyword = _extract_field(out_query, "主题")
    if extracted_email is None or extracted_keyword is None:
        raise AssertionError(f"Expected formatted subscription card. Actual: {out_query}")

    if extracted_email != email:
        raise AssertionError(f"Email mismatch. expected={email!r} got={extracted_email!r}. Output: {out_query}")
    if extracted_keyword != keyword:
        raise AssertionError(f"Keyword mismatch. expected={keyword!r} got={extracted_keyword!r}. Output: {out_query}")

    # Unsubscribe
    out_unsub = _run(client, action="取消订阅", input_user_id=user_id)
    if not out_unsub.strip():
        raise AssertionError("Expected non-empty output for unsubscribe")

    # Query after unsubscribe: accept either "not found" or inactive status
    out_query2 = _run(client, action="查询订阅", input_user_id=user_id)
    if not out_query2.strip():
        raise AssertionError("Expected non-empty output for query after unsubscribe")

    if ("⏸️" not in out_query2) and ("未找到" not in out_query2) and ("取消" not in out_query2) and ("inactive" not in out_query2):
        raise AssertionError(
            "Expected query-after-unsubscribe to indicate inactive or not found. "
            f"Actual: {out_query2}"
        )


def main() -> None:
    # Avoid Windows console encoding crashes on emoji output.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    client = CozeWorkflowClient.from_env()

    tests = [
        ("update_then_query_then_unsubscribe", test_update_then_query_then_unsubscribe),
    ]

    failed = 0
    for name, fn in tests:
        try:
            fn(client)
            print(f"PASS {name}")
        except Exception as e:
            failed += 1
            print(f"FAIL {name}: {e}")

    if failed:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
