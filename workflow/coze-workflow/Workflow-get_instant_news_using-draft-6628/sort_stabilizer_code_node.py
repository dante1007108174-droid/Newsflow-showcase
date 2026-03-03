"""Coze Code Node: deterministic hotspot sort stabilizer.

Usage in Coze workflow:
1) Add a Code node after the LLM output node.
2) Map the LLM output string to input key `content` (or `output`).
3) Use returned `result` as the final response text.

The node keeps each hotspot block text unchanged and only reorders blocks by:
1. report_count (desc)
2. media_type_count (desc)
3. publish_time (desc)
4. url/title deterministic tiebreaker (asc)
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple


_HEADER_RE = re.compile(r"^\s*(?:🔥+\s*)?(?:\*\*)?【\s*(?:热点|Hotspot)\s*\d+\s*】", re.I)


def _parse_dt(s: str) -> Optional[datetime]:
    if not s:
        return None
    s = s.strip()
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


def _extract_title(header: str) -> str:
    m = re.search(r"【\s*(?:热点|Hotspot)\s*\d+\s*】\s*(.+)$", header or "", flags=re.I)
    if not m:
        return (header or "").strip("* ")
    return m.group(1).strip("* ")


def _extract_url(block_text: str) -> str:
    m = re.search(r"\[[^\]]+\]\((https?://[^)]+)\)", block_text)
    if not m:
        return ""
    return m.group(1).strip().lower()


def _extract_report_count(block_text: str) -> int:
    m_cn = re.search(r"热度\s*[:：]\s*(\d+)\s*家媒体报道", block_text)
    if m_cn:
        return max(1, int(m_cn.group(1)))
    m_en = re.search(r"Heat\s*[:：]\s*Covered\s+by\s+(\d+)\s+sources?", block_text, flags=re.I)
    if m_en:
        return max(1, int(m_en.group(1)))
    return 1


def _extract_media_type_count(block_text: str) -> int:
    m_cn = re.search(r"跨\s*(\d+)\s*类媒体", block_text)
    if m_cn:
        return max(1, int(m_cn.group(1)))
    m_en = re.search(r"Across\s+(\d+)\s+media\s+types?", block_text, flags=re.I)
    if m_en:
        return max(1, int(m_en.group(1)))
    return 1


def _extract_publish_dt(block_text: str) -> Optional[datetime]:
    m = re.search(
        r"(?:时间|Time)\s*[:：]\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}(?::\d{2})?)",
        block_text,
        flags=re.I,
    )
    if not m:
        return None
    return _parse_dt(m.group(1))


def _split_blocks(content: str) -> Tuple[str, List[List[str]]]:
    lines = (content or "").splitlines()
    first_idx = None
    for idx, line in enumerate(lines):
        if _HEADER_RE.search(line):
            first_idx = idx
            break

    if first_idx is None:
        return content or "", []

    header = "\n".join(lines[:first_idx]).strip("\n")
    body_lines = lines[first_idx:]

    blocks: List[List[str]] = []
    current: List[str] = []
    for line in body_lines:
        if _HEADER_RE.search(line):
            if current:
                blocks.append(current)
            current = [line]
        else:
            if current:
                current.append(line)
    if current:
        blocks.append(current)

    return header, blocks


def _block_sort_key(block: List[str], idx: int):
    text = "\n".join(block)
    report_count = _extract_report_count(text)
    media_type_count = _extract_media_type_count(text)
    dt = _extract_publish_dt(text) or datetime.min
    url = _extract_url(text)
    title = _extract_title(block[0] if block else "").lower().strip()
    fallback = url or title or f"idx-{idx:04d}"
    # Desc by stats/time, then deterministic asc fallback.
    return (-report_count, -media_type_count, -int(dt.timestamp()), fallback)


def _rebuild_output(header: str, blocks: List[List[str]]) -> str:
    parts: List[str] = []
    if header:
        parts.append(header)
    for block in blocks:
        parts.append("\n".join(block).strip("\n"))
    return "\n\n".join([p for p in parts if p]).strip()


def _read_content(args: Dict) -> str:
    for key in ("content", "output", "text", "markdown"):
        val = args.get(key)
        if isinstance(val, str) and val.strip():
            return val
    params = args.get("params")
    if isinstance(params, dict):
        for key in ("content", "output", "text", "markdown"):
            val = params.get(key)
            if isinstance(val, str) and val.strip():
                return val
    return ""


def main(args):
    content = _read_content(args or {})
    if not content:
        return {
            "result": "",
            "changed": False,
            "debug_info": "empty input",
        }

    header, blocks = _split_blocks(content)
    if len(blocks) <= 1:
        return {
            "result": content,
            "changed": False,
            "debug_info": "<=1 hotspot block, skip",
        }

    indexed = list(enumerate(blocks))
    sorted_blocks = [blk for i, blk in sorted(indexed, key=lambda t: _block_sort_key(t[1], t[0]))]
    result = _rebuild_output(header, sorted_blocks)
    changed = result != content

    return {
        "result": result,
        "changed": changed,
        "debug_info": f"hotspot_blocks={len(blocks)}, reordered={changed}",
    }
