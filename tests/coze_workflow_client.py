"""Small Coze Workflow API client for local automation.

We keep this dependency-light (requests + python-dotenv) so it can be used in
CI or on a dev machine.
"""

from __future__ import annotations

import os
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional
import json

import requests
from dotenv import load_dotenv
from requests.exceptions import RequestException, SSLError


DEFAULT_COZE_BASE_URL = "https://api.coze.cn"


def _load_backend_env() -> None:
    """Load backend/.env if present (do not print secrets)."""
    env_path = Path(__file__).resolve().parents[1] / "backend" / ".env"
    if env_path.exists():
        load_dotenv(env_path)


def _get_required_env(name: str) -> str:
    value = os.getenv(name)
    if value is None or value.strip() == "":
        raise RuntimeError(
            f"Missing required env var {name}. "
            f"Set it in backend/.env or in your shell environment."
        )
    return value


def _get_base_url() -> str:
    # Prefer COZE_BASE_URL (matches this repo), fall back to COZE_API_BASE (coze-py examples)
    base = (os.getenv("COZE_BASE_URL") or os.getenv("COZE_API_BASE") or DEFAULT_COZE_BASE_URL).strip()
    return base.rstrip("/")


def _extract_sync_output(payload: Any) -> Optional[str]:
    """Best-effort extraction of sync run output from Coze response JSON."""
    if not isinstance(payload, dict):
        return None

    data = payload.get("data")
    # Some APIs wrap as { data: { data: "..." } }
    if isinstance(data, dict):
        for key in ("data", "output", "content"):
            val = data.get(key)
            if isinstance(val, str):
                return val
        # Some responses might embed output under nested keys
        if isinstance(data.get("output"), dict) and isinstance(data["output"].get("content"), str):
            return data["output"]["content"]
    # Or { data: "..." }
    if isinstance(data, str):
        return data

    return None


def _extract_debug_url(payload: Any) -> Optional[str]:
    """Best-effort extraction of debug_url from Coze response JSON."""

    if not isinstance(payload, dict):
        return None
    v = payload.get("debug_url")
    if isinstance(v, str) and v.strip():
        return v.strip()
    data = payload.get("data")
    if isinstance(data, dict):
        v2 = data.get("debug_url")
        if isinstance(v2, str) and v2.strip():
            return v2.strip()
    return None


def _extract_execute_id(payload: Any) -> Optional[str]:
    """Best-effort extraction of execute_id from Coze response JSON."""

    if not isinstance(payload, dict):
        return None
    v = payload.get("execute_id")
    if isinstance(v, str) and v.strip():
        return v.strip()
    data = payload.get("data")
    if isinstance(data, dict):
        v2 = data.get("execute_id")
        if isinstance(v2, str) and v2.strip():
            return v2.strip()
    return None


def _normalize_output_text(text: str) -> str:
    """Coze may return workflow output as a JSON string wrapper.

    Example (string): {"content_type":1,"data":"..."}
    In that case, we want the inner `data` field.
    """
    s = (text or "").strip()
    if not s:
        return ""

    if s.startswith("{") and s.endswith("}"):
        try:
            obj = json.loads(s)
            if isinstance(obj, dict):
                inner = obj.get("data")
                if isinstance(inner, str):
                    return inner
        except Exception:
            # Not valid JSON; return as-is.
            return s

    return s


@dataclass
class WorkflowRunResult:
    output: str
    execute_id: Optional[str] = None
    debug_url: Optional[str] = None
    raw: Optional[Dict[str, Any]] = None


class CallLimiter:
    """Thread-safe API call counter with hard quota."""

    def __init__(self, max_calls: Optional[int] = None):
        self._max_calls = max_calls
        self._count = 0
        self._lock = threading.Lock()
        self._history: List[Dict[str, Any]] = []
        self._method_counts: Dict[str, int] = {}

    def check(self, method: str, url: str) -> None:
        with self._lock:
            self._count += 1
            self._method_counts[method] = self._method_counts.get(method, 0) + 1
            if len(self._history) >= 100:
                self._history.pop(0)
            self._history.append({"n": self._count, "method": method, "url": url})
            if self._max_calls is not None and self._count > self._max_calls:
                raise RuntimeError(
                    f"API call quota exceeded: {self._count} > {self._max_calls}. "
                    f"Last call: {method} {url}. "
                    f"Set COZE_MAX_CALLS env var to increase limit."
                )

    def stats(self) -> Dict[str, Any]:
        with self._lock:
            return {
                "total_calls": self._count,
                "max_calls": self._max_calls,
                "remaining": self._max_calls - self._count if self._max_calls else None,
                "method_counts": dict(self._method_counts),
                "history": list(self._history),
            }


class CozeWorkflowClient:
    def __init__(
        self,
        token: str,
        base_url: str,
        *,
        max_calls: Optional[int] = None,
        call_limiter: Optional[CallLimiter] = None,
    ):
        self._token = token
        self._base_url = base_url.rstrip("/")
        self._session = requests.Session()
        self._session.trust_env = True
        self._call_limiter = call_limiter or CallLimiter(max_calls=max_calls)

    @classmethod
    def from_env(
        cls,
        *,
        max_calls: Optional[int] = None,
        call_limiter: Optional[CallLimiter] = None,
    ) -> "CozeWorkflowClient":
        """Create client from environment variables (backend/.env)."""
        _load_backend_env()
        token = _get_required_env("COZE_API_TOKEN")
        return cls(
            token=token,
            base_url=_get_base_url(),
            max_calls=max_calls,
            call_limiter=call_limiter,
        )

    def _headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self._token}",
            "Content-Type": "application/json",
        }

    def _check_quota(self, method: str, url: str) -> None:
        self._call_limiter.check(method, url)

    def get_call_stats(self) -> Dict[str, Any]:
        return self._call_limiter.stats()

    def _ssl_verify(self) -> Any:
        v = (os.getenv("COZE_SSL_VERIFY") or "").strip().lower()
        if v in ("0", "false", "no", "off"):
            return False
        ca = (os.getenv("COZE_CA_BUNDLE") or "").strip()
        if ca:
            return ca
        return True

    def _request(
        self, method: str, url: str, *, is_trigger: bool = False, **kwargs: Any
    ) -> requests.Response:
        """Make an HTTP request with quota tracking.

        POST workflow triggers (is_trigger=True) never retry.
        GET polling requests retry once on transient errors.
        """
        self._check_quota(method, url)

        if is_trigger:
            retries = 0
        else:
            retries = int(os.getenv("COZE_REQUEST_RETRIES") or "1")

        backoff_s = float(os.getenv("COZE_RETRY_BACKOFF_S") or "0.6")

        last_err: Optional[BaseException] = None
        for attempt in range(1, retries + 2):
            try:
                resp = self._session.request(
                    method, url, verify=self._ssl_verify(), **kwargs
                )
                resp.raise_for_status()
                return resp
            except RequestException as e:
                last_err = e
                if attempt >= retries + 1:
                    raise
                time.sleep(backoff_s * attempt)

        # Should never reach here, but just in case
        raise last_err  # type: ignore[misc]

    def run_workflow(
        self,
        workflow_id: str,
        parameters: Dict[str, Any],
        *,
        is_async: bool = False,
        timeout_s: int = 300,
        poll_interval_s: float = 2.0,
    ) -> WorkflowRunResult:
        """Run a workflow and return its final output.

        If is_async=True, triggers the workflow then polls run history until done.
        POST trigger never retries. GET polling retries on transient errors.
        """
        run_url = f"{self._base_url}/v1/workflow/run"

        # Trigger: exactly 1 POST, no retry
        resp = self._request(
            "POST",
            run_url,
            is_trigger=True,
            headers=self._headers(),
            json={
                "workflow_id": workflow_id,
                "is_async": bool(is_async),
                "parameters": parameters,
            },
            timeout=timeout_s,
        )
        payload = resp.json()

        # Try sync output first
        sync_output = _extract_sync_output(payload)
        if sync_output is not None and sync_output != "":
            debug_url = _extract_debug_url(payload)
            return WorkflowRunResult(
                output=_normalize_output_text(sync_output),
                debug_url=debug_url,
                raw=payload,
            )

        # Async flow: need execute_id to poll
        execute_id = _extract_execute_id(payload)
        if not execute_id:
            raise RuntimeError(
                "Coze workflow run did not return output or execute_id. "
                "You may need to switch to streaming_run, or inspect the raw response."
            )

        # Poll run history until completion
        history_url = f"{self._base_url}/v1/workflows/{workflow_id}/run_histories/{execute_id}"
        deadline = time.time() + timeout_s

        while True:
            if time.time() > deadline:
                raise TimeoutError(
                    f"Timed out waiting for workflow {workflow_id} "
                    f"execute_id={execute_id}"
                )

            h = self._request(
                "GET",
                history_url,
                is_trigger=False,
                headers=self._headers(),
                timeout=60,
            )
            last_payload = h.json()

            data = last_payload.get("data") if isinstance(last_payload, dict) else None
            item = None
            if isinstance(data, list) and data:
                item = data[0]

            if isinstance(item, dict):
                status = item.get("execute_status")
                status_lower = str(status).lower() if status is not None else ""
                if status_lower == "running" or status == 1:
                    time.sleep(poll_interval_s)
                    continue
                if status_lower in ("fail", "failed") or status == 2:
                    raise RuntimeError(
                        item.get("error_message") or "Workflow run failed"
                    )

                output = item.get("output")
                if isinstance(output, str):
                    return WorkflowRunResult(
                        output=_normalize_output_text(output),
                        execute_id=execute_id,
                        debug_url=item.get("debug_url")
                        if isinstance(item.get("debug_url"), str)
                        else None,
                        raw=last_payload,
                    )

            # Unknown shape; wait and retry
            time.sleep(poll_interval_s)
