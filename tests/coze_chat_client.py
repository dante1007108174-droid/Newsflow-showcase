"""Lightweight Coze v3 chat client for dispatch-model evaluation.

This client is separate from `coze_workflow_client.py` because dispatch tests need
chat-level behavior (intent routing, slot extraction, follow-up questions), not direct
workflow invocation.
"""

from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests
from dotenv import load_dotenv


DEFAULT_COZE_BASE_URL = "https://api.coze.cn"


def _load_backend_env() -> None:
    env_path = Path(__file__).resolve().parents[1] / "backend" / ".env"
    if env_path.exists():
        _ = load_dotenv(env_path)


def _get_required_env(name: str) -> str:
    value = (os.getenv(name) or "").strip()
    if not value:
        raise RuntimeError(
            f"Missing required env var {name}. Set it in backend/.env or shell env."
        )
    return value


def _safe_json_loads(text: str) -> Optional[Any]:
    if not isinstance(text, str):
        return None
    s = text.strip()
    if not s:
        return None
    try:
        return json.loads(s)
    except Exception:
        return None


def _collect_strings(obj: Any) -> List[str]:
    out: List[str] = []
    if isinstance(obj, str):
        out.append(obj)
        return out
    if isinstance(obj, list):
        for item in obj:
            out.extend(_collect_strings(item))
        return out
    if isinstance(obj, dict):
        for v in obj.values():
            out.extend(_collect_strings(v))
        return out
    return out


def _normalize_skill_name(name: str) -> str:
    s = (name or "").strip()
    lower = s.lower()
    # Canonical production workflow names.
    if "get_instant_news_using" in lower:
        return "get_instant_news_using"
    if "subscribe_using" in lower:
        return "subscribe_using"
    if "send_mail_cache_using" in lower:
        return "send_mail_cache_using"
    if "send_mail_using" in lower:
        return "send_mail_using"
    # Allow STUB/tool names to map back to canonical names.
    if "get_instant_news" in lower:
        return "get_instant_news_using"
    if "subscribe" in lower:
        return "subscribe_using"
    if "send_mail" in lower:
        return "send_mail_using"
    return s


def _parse_stub_text(text: str) -> List[Dict[str, Any]]:
    """Parse lines like:

    [STUB] get_instant_news_using | keyword=AI | user_id=dispatch_test_user_001
    """
    parsed: List[Dict[str, Any]] = []
    if "[STUB]" not in text:
        return parsed

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line.startswith("[STUB]"):
            continue

        body = line[len("[STUB]") :].strip()
        parts = [p.strip() for p in body.split("|") if p.strip()]
        if not parts:
            continue

        skill = _normalize_skill_name(parts[0])
        params: Dict[str, str] = {}
        for p in parts[1:]:
            if "=" not in p:
                continue
            k, v = p.split("=", 1)
            params[k.strip()] = v.strip().strip('"').strip("'")

        parsed.append({"skill": skill, "params": params, "raw": line})

    return parsed


@dataclass
class ChatRunResult:
    conversation_id: str
    chat_id: str
    user_id: str
    status: str
    final_answer: str
    messages: List[Dict[str, Any]]
    function_calls: List[Dict[str, Any]]
    stub_calls: List[Dict[str, Any]]
    first_response_ms: Optional[int]
    total_ms: int
    raw_create: Optional[Dict[str, Any]] = None
    raw_retrieve_last: Optional[Dict[str, Any]] = None


class CozeChatClient:
    def __init__(self, token: str, base_url: str):
        self._token = token
        self._base_url = base_url.rstrip("/")
        self._session = requests.Session()
        self._session.trust_env = True

    @classmethod
    def from_env(cls) -> "CozeChatClient":
        _load_backend_env()
        token = _get_required_env("COZE_API_TOKEN")
        base_url = (
            os.getenv("COZE_BASE_URL")
            or os.getenv("COZE_API_BASE")
            or DEFAULT_COZE_BASE_URL
        ).strip()
        return cls(token=token, base_url=base_url)

    def _headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self._token}",
            "Content-Type": "application/json",
        }

    def create_conversation(self, *, bot_id: str) -> str:
        """Pre-create a conversation via POST /v1/conversation/create.

        This is required for multi-turn chat: all turns in the same group
        must share a conversation_id obtained from this endpoint.
        Passing a conversation_id directly to /v3/chat without pre-creating
        it results in each turn starting a new conversation.
        """
        resp = self._session.post(
            f"{self._base_url}/v1/conversation/create",
            headers=self._headers(),
            json={"bot_id": bot_id},
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        conv_id = ""
        if isinstance(data, dict):
            inner = data.get("data") or data
            conv_id = str(inner.get("id") or inner.get("conversation_id") or "")
        if not conv_id:
            raise RuntimeError(
                f"create_conversation returned no id: {json.dumps(data, ensure_ascii=False)}"
            )
        return conv_id

    @staticmethod
    def _append_system_context(message: str, user_id: str) -> str:
        return f'{message.strip()}\n\n[System Context: user_id="{user_id}"]'

    def _create_chat(
        self,
        *,
        bot_id: str,
        message: str,
        user_id: str,
        conversation_id: Optional[str],
        stream: bool,
        timeout_s: int,
    ) -> Dict[str, Any]:
        payload: Dict[str, Any] = {
            "bot_id": bot_id,
            "user_id": user_id,
            "stream": stream,
            "auto_save_history": True,
            "additional_messages": [
                {
                    "role": "user",
                    "content": self._append_system_context(message, user_id),
                    "content_type": "text",
                }
            ],
        }
        params: Dict[str, str] = {}
        if conversation_id:
            # Coze v3 chat reads conversation_id from query params.
            params["conversation_id"] = conversation_id

        resp = self._session.post(
            f"{self._base_url}/v3/chat",
            headers=self._headers(),
            params=params,
            json=payload,
            timeout=timeout_s,
        )
        resp.raise_for_status()
        return resp.json()

    def _retrieve_chat(self, *, chat_id: str, conversation_id: str, timeout_s: int) -> Dict[str, Any]:
        url = (
            f"{self._base_url}/v3/chat/retrieve"
            f"?chat_id={chat_id}"
            f"&conversation_id={conversation_id}"
        )
        try:
            resp = self._session.get(url, headers=self._headers(), timeout=timeout_s)
            resp.raise_for_status()
            return resp.json()
        except Exception:
            resp = self._session.post(url, headers=self._headers(), json={}, timeout=timeout_s)
            resp.raise_for_status()
            return resp.json()

    def _list_messages(self, *, chat_id: str, conversation_id: str, timeout_s: int) -> List[Dict[str, Any]]:
        url = (
            f"{self._base_url}/v3/chat/message/list"
            f"?chat_id={chat_id}"
            f"&conversation_id={conversation_id}"
        )
        resp = self._session.post(url, headers=self._headers(), json={}, timeout=timeout_s)
        resp.raise_for_status()
        payload = resp.json()
        messages = payload.get("data") if isinstance(payload, dict) else None
        return messages if isinstance(messages, list) else []

    @staticmethod
    def _is_finish_signal(text: str) -> bool:
        s = (text or "").strip()
        if not s:
            return False
        if '"msg_type":"generate_answer_finish"' in s:
            return True
        obj = _safe_json_loads(s)
        if not isinstance(obj, dict):
            return False
        if str(obj.get("msg_type") or "") == "generate_answer_finish":
            return True
        data = obj.get("data")
        if isinstance(data, str):
            nested = _safe_json_loads(data)
            if isinstance(nested, dict) and "finish_reason" in nested:
                fin_data = nested.get("FinData")
                if fin_data in (None, ""):
                    return True
        return False

    @staticmethod
    def _extract_answer_text(content: Any) -> str:
        if not isinstance(content, str):
            return ""
        raw = content.strip()
        if not raw or CozeChatClient._is_finish_signal(raw):
            return ""

        obj = _safe_json_loads(raw)
        if not isinstance(obj, dict):
            return raw

        for key in ("answer", "content", "text", "message"):
            v = obj.get(key)
            if isinstance(v, str):
                candidate = v.strip()
                if candidate and not CozeChatClient._is_finish_signal(candidate):
                    return candidate

        data = obj.get("data")
        if isinstance(data, str):
            nested = _safe_json_loads(data)
            if isinstance(nested, dict):
                for key in ("answer", "content", "text", "message", "FinData"):
                    v = nested.get(key)
                    if isinstance(v, str):
                        candidate = v.strip()
                        if candidate and not CozeChatClient._is_finish_signal(candidate):
                            return candidate

        return ""

    @staticmethod
    def _pick_final_answer(messages: List[Dict[str, Any]]) -> str:
        # Priority 1: explicit assistant answer messages.
        for i in range(len(messages) - 1, -1, -1):
            m = messages[i] if isinstance(messages[i], dict) else {}
            role = str(m.get("role") or "")
            msg_type = str(m.get("type") or "")
            if role != "assistant" or msg_type != "answer":
                continue
            text = CozeChatClient._extract_answer_text(m.get("content"))
            if text:
                return text

        # Priority 2: fallback to assistant text/object_string excluding finish signals.
        for i in range(len(messages) - 1, -1, -1):
            m = messages[i] if isinstance(messages[i], dict) else {}
            role = str(m.get("role") or "")
            msg_type = str(m.get("type") or "")
            content_type = str(m.get("content_type") or "")
            if role != "assistant":
                continue
            if msg_type in ("function_call", "tool_output", "tool_response", "verbose", "follow_up"):
                continue
            if content_type not in ("text", "object_string", ""):
                continue
            text = CozeChatClient._extract_answer_text(m.get("content"))
            if text:
                return text

        return ""

    @staticmethod
    def _extract_function_calls(messages: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        calls: List[Dict[str, Any]] = []
        for m in messages:
            if not isinstance(m, dict):
                continue
            msg_type = str(m.get("type") or "")
            if msg_type != "function_call":
                continue
            raw = m.get("content")
            if not isinstance(raw, str):
                continue
            obj = _safe_json_loads(raw)
            if not isinstance(obj, dict):
                continue
            name = _normalize_skill_name(str(obj.get("name") or ""))
            arguments = obj.get("arguments")
            if isinstance(arguments, str):
                parsed_args = _safe_json_loads(arguments)
                if isinstance(parsed_args, dict):
                    arguments = parsed_args
            if not isinstance(arguments, dict):
                arguments = {}
            calls.append({
                "skill": name,
                "params": arguments,
                "raw": obj,
            })
        return calls

    @staticmethod
    def _extract_stub_calls(messages: List[Dict[str, Any]], final_answer: str) -> List[Dict[str, Any]]:
        out: List[Dict[str, Any]] = []
        text_candidates: List[str] = []

        for m in messages:
            if not isinstance(m, dict):
                continue
            content = m.get("content")
            if isinstance(content, str):
                text_candidates.append(content)
                obj = _safe_json_loads(content)
                if obj is not None:
                    text_candidates.extend(_collect_strings(obj))

        if isinstance(final_answer, str) and final_answer.strip():
            text_candidates.append(final_answer)

        for t in text_candidates:
            out.extend(_parse_stub_text(t))

        dedup: List[Dict[str, Any]] = []
        seen: set[str] = set()
        for item in out:
            sig = f"{item.get('skill')}|{json.dumps(item.get('params') or {}, ensure_ascii=False, sort_keys=True)}"
            if sig in seen:
                continue
            seen.add(sig)
            dedup.append(item)
        return dedup

    def _send_non_stream(
        self,
        *,
        bot_id: str,
        message: str,
        user_id: str,
        conversation_id: Optional[str],
        max_polls: int,
        poll_interval_s: float,
        timeout_s: int,
    ) -> ChatRunResult:
        t0 = time.time()
        create_payload = self._create_chat(
            bot_id=bot_id,
            message=message,
            user_id=user_id,
            conversation_id=conversation_id,
            stream=False,
            timeout_s=timeout_s,
        )

        data = create_payload.get("data") if isinstance(create_payload, dict) else {}
        chat_id = str((data or {}).get("id") or "")
        conv_id = str((data or {}).get("conversation_id") or "")
        status = str((data or {}).get("status") or "")
        if not chat_id or not conv_id:
            raise RuntimeError("Coze chat create response missing chat_id or conversation_id")

        last_retrieve: Optional[Dict[str, Any]] = None
        final_status = status
        for _ in range(max_polls):
            if final_status == "completed":
                break
            time.sleep(poll_interval_s)
            r = self._retrieve_chat(chat_id=chat_id, conversation_id=conv_id, timeout_s=15)
            last_retrieve = r
            r_data = r.get("data") if isinstance(r, dict) else {}
            final_status = str((r_data or {}).get("status") or "")

        messages = self._list_messages(chat_id=chat_id, conversation_id=conv_id, timeout_s=30)
        answer = self._pick_final_answer(messages)
        function_calls = self._extract_function_calls(messages)
        stub_calls = self._extract_stub_calls(messages, answer)

        total_ms = int((time.time() - t0) * 1000)
        return ChatRunResult(
            conversation_id=conv_id,
            chat_id=chat_id,
            user_id=user_id,
            status=final_status,
            final_answer=answer,
            messages=messages,
            function_calls=function_calls,
            stub_calls=stub_calls,
            first_response_ms=total_ms,
            total_ms=total_ms,
            raw_create=create_payload if isinstance(create_payload, dict) else None,
            raw_retrieve_last=last_retrieve,
        )

    def _send_stream(
        self,
        *,
        bot_id: str,
        message: str,
        user_id: str,
        conversation_id: Optional[str],
        timeout_s: int,
    ) -> ChatRunResult:
        t0 = time.time()

        payload: Dict[str, Any] = {
            "bot_id": bot_id,
            "user_id": user_id,
            "stream": True,
            "auto_save_history": True,
            "additional_messages": [
                {
                    "role": "user",
                    "content": self._append_system_context(message, user_id),
                    "content_type": "text",
                }
            ],
        }
        params: Dict[str, str] = {}
        if conversation_id:
            # Keep context by passing conversation_id in query params.
            params["conversation_id"] = conversation_id

        url = f"{self._base_url}/v3/chat"
        resp = self._session.post(
            url,
            headers=self._headers(),
            params=params,
            json=payload,
            stream=True,
            timeout=(15, timeout_s),
        )
        resp.raise_for_status()

        upstream_chat_id = ""
        upstream_conv_id = ""
        final_status = "running"
        first_response_ms: Optional[int] = None

        event_name: Optional[str] = None
        data_lines: List[str] = []

        def flush_event() -> None:
            nonlocal event_name, data_lines, upstream_chat_id, upstream_conv_id, first_response_ms, final_status

            if not event_name:
                data_lines = []
                return

            data_str = "\n".join(data_lines).strip()
            obj: Dict[str, Any] = {}
            if data_str:
                parsed = _safe_json_loads(data_str)
                if isinstance(parsed, dict):
                    obj = parsed

            if event_name in ("conversation.chat.created", "conversation.chat.in_progress"):
                upstream_chat_id = str(obj.get("id") or upstream_chat_id)
                upstream_conv_id = str(obj.get("conversation_id") or upstream_conv_id)
            elif event_name == "conversation.message.delta":
                upstream_chat_id = str(obj.get("chat_id") or upstream_chat_id)
                upstream_conv_id = str(obj.get("conversation_id") or upstream_conv_id)
                if first_response_ms is None:
                    role = str(obj.get("role") or "")
                    content_type = str(obj.get("content_type") or "")
                    content = obj.get("content")
                    if role == "assistant" and content_type == "text" and isinstance(content, str) and content:
                        first_response_ms = int((time.time() - t0) * 1000)
            elif event_name in ("conversation.chat.completed", "done"):
                final_status = "completed"
            elif event_name == "error":
                final_status = "failed"

            event_name = None
            data_lines = []

        for raw_line in resp.iter_lines(decode_unicode=True):
            if raw_line is None:
                continue
            line = raw_line.strip()
            if not line:
                flush_event()
                if final_status in ("completed", "failed"):
                    break
                continue
            if line.startswith(":"):
                continue
            if line.startswith("event:"):
                event_name = line[len("event:") :].strip()
                continue
            if line.startswith("data:"):
                data_lines.append(line[len("data:") :].strip())
                continue

        if event_name or data_lines:
            flush_event()

        if not upstream_chat_id or not upstream_conv_id:
            raise RuntimeError("Streaming response missing chat_id/conversation_id")

        messages = self._list_messages(
            chat_id=upstream_chat_id,
            conversation_id=upstream_conv_id,
            timeout_s=30,
        )
        answer = self._pick_final_answer(messages)
        function_calls = self._extract_function_calls(messages)
        stub_calls = self._extract_stub_calls(messages, answer)
        total_ms = int((time.time() - t0) * 1000)
        if first_response_ms is None:
            first_response_ms = total_ms

        return ChatRunResult(
            conversation_id=upstream_conv_id,
            chat_id=upstream_chat_id,
            user_id=user_id,
            status=final_status,
            final_answer=answer,
            messages=messages,
            function_calls=function_calls,
            stub_calls=stub_calls,
            first_response_ms=first_response_ms,
            total_ms=total_ms,
            raw_create=None,
            raw_retrieve_last=None,
        )

    def chat_once(
        self,
        *,
        bot_id: str,
        message: str,
        user_id: str,
        conversation_id: Optional[str] = None,
        use_stream: bool = True,
        max_polls: int = 20,
        poll_interval_s: float = 0.8,
        timeout_s: int = 120,
    ) -> ChatRunResult:
        if use_stream:
            return self._send_stream(
                bot_id=bot_id,
                message=message,
                user_id=user_id,
                conversation_id=conversation_id,
                timeout_s=timeout_s,
            )
        return self._send_non_stream(
            bot_id=bot_id,
            message=message,
            user_id=user_id,
            conversation_id=conversation_id,
            max_polls=max_polls,
            poll_interval_s=poll_interval_s,
            timeout_s=timeout_s,
        )


def pick_primary_call(
    *,
    function_calls: List[Dict[str, Any]],
    stub_calls: List[Dict[str, Any]],
) -> Tuple[str, Dict[str, Any]]:
    """Pick one normalized call for evaluator.

    Priority:
    1) STUB call (contains canonical workflow name + params)
    2) function_call message from Coze
    """
    if stub_calls:
        first_stub = stub_calls[0]
        stub_skill = str(first_stub.get("skill") or "")
        stub_params = dict(first_stub.get("params") or {})
        if function_calls:
            # Prefer STUB skill name (often canonical), but merge params from
            # function_call to avoid losing structured fields (e.g. raw_query).
            first_func = function_calls[0]
            func_skill = str(first_func.get("skill") or "")
            func_params = dict(first_func.get("params") or {})
            merged_params = dict(stub_params)
            merged_params.update(func_params)
            return stub_skill or func_skill, merged_params
        if stub_params:
            return stub_skill, stub_params
        return stub_skill, {}
    if function_calls:
        first = function_calls[0]
        return str(first.get("skill") or ""), dict(first.get("params") or {})
    return "", {}
