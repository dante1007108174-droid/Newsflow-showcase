"""Run gen_news workflow suite and write reports.

Usage:
  python tests/run_gen_news_suite.py

Outputs (tests/reports/):
  - gen_news_report_{run_id}.json
  - gen_news_report_{run_id}.md
  Plus stable "latest" files:
  - gen_news_report.json
  - gen_news_report.md

Notes:
  - Tests news generation workflow with dedup, source diversity, and source cap checks.
  - Each case runs the workflow twice to check stability.
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

from coze_workflow_client import CozeWorkflowClient


ROOT = Path(__file__).resolve().parents[1]
SUITE_PATH = Path(__file__).resolve().parent / "gen_news_suite.json"
REPORT_DIR = Path(__file__).resolve().parent / "reports"


def _safe_print(s: str) -> None:
    try:
        print(s)
    except UnicodeEncodeError:
        print(s.encode("utf-8", errors="replace").decode("utf-8", errors="replace"))


def _now_ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _load_json_file(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _save_json(path: Path, data: Dict[str, Any]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _is_infra_exception(e: Exception) -> bool:
    msg = f"{type(e).__name__}: {e}".lower()
    return any(
        k in msg
        for k in (
            "sslerror",
            "certificate_verify_failed",
            "hostname mismatch",
            "connectionerror",
            "readtimeout",
            "connecttimeout",
            "proxyerror",
            "temporary failure",
            "name resolution",
        )
    )


@dataclass
class CheckResult:
    name: str
    passed: bool
    reason: str = ""
    root_cause: str = ""
    suggestion: str = ""


@dataclass
class NewsItem:
    title: str
    tag: str
    time: str
    summary: str
    source: str
    link: str

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "NewsItem":
        return cls(
            title=str(data.get("title") or "").strip(),
            tag=str(data.get("tag") or "").strip(),
            time=str(data.get("time") or "").strip(),
            summary=str(data.get("summary") or "").strip(),
            source=str(data.get("source") or "").strip(),
            link=str(data.get("link") or "").strip(),
        )


@dataclass
class NewsOutput:
    top_news: List[NewsItem]
    quick_news: List[NewsItem]

    def all_items(self) -> List[NewsItem]:
        return self.top_news + self.quick_news

    def source_counts(self) -> Counter:
        return Counter(item.source for item in self.all_items() if item.source)


def _parse_news_output(output: str) -> Tuple[Optional[NewsOutput], str]:
    """Parse workflow output into NewsOutput."""
    if not output or not output.strip():
        return None, "输出为空"

    try:
        data = json.loads(output.strip())
    except json.JSONDecodeError as e:
        return None, f"JSON解析失败: {e}"

    if not isinstance(data, dict):
        return None, "输出不是JSON对象"

    top_news_raw = data.get("top_news") or []
    quick_news_raw = data.get("quick_news") or []

    if not isinstance(top_news_raw, list):
        return None, "top_news不是数组"
    if not isinstance(quick_news_raw, list):
        return None, "quick_news不是数组"

    top_news = [NewsItem.from_dict(item) for item in top_news_raw if isinstance(item, dict)]
    quick_news = [NewsItem.from_dict(item) for item in quick_news_raw if isinstance(item, dict)]

    return NewsOutput(top_news=top_news, quick_news=quick_news), ""


def _extract_entities(text: str) -> Set[str]:
    """Extract entities from text: Chinese 2-4 char words, English brand names, version numbers."""
    entities: Set[str] = set()
    if not text:
        return entities

    # Chinese words (2-4 characters)
    chinese_words = re.findall(r'[\u4e00-\u9fff]{2,4}', text)
    entities.update(chinese_words)

    # English brand names (capitalized words, possibly with numbers)
    english_brands = re.findall(r'\b[A-Z][a-zA-Z0-9]*(?:\s+[A-Z][a-zA-Z0-9]*)*\b', text)
    entities.update(english_brands)

    # Version numbers like 3.5, v2, GPT-4, etc.
    version_patterns = re.findall(r'\b(?:v\d+(?:\.\d+)?|\d+\.\d+|GPT-\d+|Claude-\d+)\b', text, re.I)
    entities.update(version_patterns)

    # Technical terms and acronyms
    tech_terms = re.findall(r'\b[A-Z]{2,}\b', text)
    entities.update(tech_terms)

    return entities


def _jaccard_similarity(set1: Set[str], set2: Set[str]) -> float:
    """Compute Jaccard similarity between two sets."""
    if not set1 and not set2:
        return 0.0
    intersection = len(set1 & set2)
    union = len(set1 | set2)
    return intersection / union if union > 0 else 0.0


def _check_summary_overlap(summary1: str, summary2: str) -> float:
    """Check overlap between two summaries."""
    if not summary1 or not summary2:
        return 0.0
    words1 = set(re.findall(r'[\u4e00-\u9fff]{2,4}', summary1))
    words2 = set(re.findall(r'[\u4e00-\u9fff]{2,4}', summary2))
    return _jaccard_similarity(words1, words2)


def _check_dedup(news_output: NewsOutput) -> CheckResult:
    """Check for duplicate news items based on title entity overlap."""
    items = news_output.all_items()
    if len(items) < 2:
        return CheckResult(name="dedup", passed=True)

    # Extract entities for each item
    item_entities: List[Tuple[int, Set[str], str, str]] = []
    for i, item in enumerate(items):
        entities = _extract_entities(item.title)
        item_entities.append((i, entities, item.title, item.summary))

    duplicates: List[str] = []
    threshold = 0.55

    for i in range(len(item_entities)):
        for j in range(i + 1, len(item_entities)):
            idx1, entities1, title1, summary1 = item_entities[i]
            idx2, entities2, title2, summary2 = item_entities[j]

            # Base similarity from title entities
            similarity = _jaccard_similarity(entities1, entities2)

            # Boost if summaries overlap significantly
            summary_overlap = _check_summary_overlap(summary1, summary2)
            if summary_overlap > 0.3:
                similarity = min(1.0, similarity + 0.2)

            if similarity >= threshold:
                duplicates.append(f"[{idx1}]{title1[:30]}... vs [{idx2}]{title2[:30]}... (相似度:{similarity:.2f})")

    if duplicates:
        return CheckResult(
            name="dedup",
            passed=False,
            reason=f"发现{len(duplicates)}组重复或高度相似条目: {'; '.join(duplicates[:3])}",
            root_cause="去重机制未能有效识别相似新闻",
            suggestion="优化去重算法，提高实体提取准确性，或添加更多去重特征"
        )

    return CheckResult(name="dedup", passed=True)


def _check_source_diversity(news_output: NewsOutput) -> CheckResult:
    """Check that there are at least 3 distinct sources."""
    source_counts = news_output.source_counts()
    distinct_sources = len(source_counts)

    if distinct_sources < 3:
        sources_list = list(source_counts.keys())
        return CheckResult(
            name="source_diversity",
            passed=False,
            reason=f"信源数量不足: 仅{distinct_sources}个(要求>=3): {sources_list}",
            root_cause="信源覆盖面不够广",
            suggestion="增加信源渠道，确保覆盖至少3个不同来源"
        )

    return CheckResult(name="source_diversity", passed=True)


def _check_source_cap(news_output: NewsOutput) -> CheckResult:
    """Check that no single source exceeds 40% of total items. Track 华尔街见闻 specifically."""
    items = news_output.all_items()
    total = len(items)

    if total == 0:
        return CheckResult(name="source_cap", passed=True)

    source_counts = news_output.source_counts()
    max_allowed_ratio = 0.40
    violations: List[str] = []

    wallstreet_count = 0
    for source, count in source_counts.items():
        ratio = count / total
        if ratio > max_allowed_ratio:
            violations.append(f"{source}: {count}/{total} ({ratio*100:.1f}%)")
        if "华尔街见闻" in source or "wallstreetcn" in source.lower():
            wallstreet_count = count

    wallstreet_ratio = wallstreet_count / total if total > 0 else 0

    if violations:
        return CheckResult(
            name="source_cap",
            passed=False,
            reason=f"信源占比超限(>40%): {'; '.join(violations)}; 华尔街见闻: {wallstreet_count}/{total} ({wallstreet_ratio*100:.1f}%)",
            root_cause="单一信源占比过高，信息来源不够均衡",
            suggestion="限制单一信源的最大占比，增加其他信源的抓取权重"
        )

    # Also warn if 华尔街见闻 is close to limit
    if wallstreet_ratio >= 0.35:
        return CheckResult(
            name="source_cap",
            passed=True,
            reason=f"警告: 华尔街见闻占比{wallstreet_ratio*100:.1f}%接近上限(40%)",
            root_cause="",
            suggestion=""
        )

    return CheckResult(
        name="source_cap",
        passed=True,
        reason=f"华尔街见闻: {wallstreet_count}/{total} ({wallstreet_ratio*100:.1f}%)"
    )


def _check_non_empty(news_output: NewsOutput) -> CheckResult:
    """Check that output is not empty."""
    items = news_output.all_items()
    if not items:
        return CheckResult(
            name="non_empty",
            passed=False,
            reason="输出为空，没有新闻条目",
            root_cause="工作流未返回任何新闻数据",
            suggestion="检查工作流执行状态和输入参数"
        )
    return CheckResult(name="non_empty", passed=True, reason=f"共{len(items)}条新闻")


def _check_valid_json(news_output: NewsOutput, parse_error: str) -> CheckResult:
    """Check that output is valid JSON with correct structure."""
    if parse_error:
        return CheckResult(
            name="valid_json",
            passed=False,
            reason=parse_error,
            root_cause="工作流输出不是有效的JSON格式",
            suggestion="检查工作流输出节点，确保输出格式为JSON"
        )
    return CheckResult(name="valid_json", passed=True)


def _run_checks(news_output: Optional[NewsOutput], parse_error: str) -> List[CheckResult]:
    """Run all checks on the news output."""
    results: List[CheckResult] = []

    # valid_json check first
    results.append(_check_valid_json(news_output, parse_error))

    if news_output is None:
        # Can't run other checks if parsing failed
        return results

    results.append(_check_non_empty(news_output))
    results.append(_check_dedup(news_output))
    results.append(_check_source_diversity(news_output))
    results.append(_check_source_cap(news_output))

    return results


def _call_workflow(
    client: CozeWorkflowClient, workflow_id: str, params: Dict[str, str]
) -> Tuple[str, int, Optional[str]]:
    """Call workflow and return output."""
    t0 = time.time()
    res = client.run_workflow(workflow_id, params, is_async=False, timeout_s=300)
    dt_ms = int((time.time() - t0) * 1000)
    return res.output or "", dt_ms, res.debug_url


def _run_one_case(workflow_id: str, case: Dict[str, Any]) -> Dict[str, Any]:
    """Run a single test case."""
    case_id = str(case.get("id") or "").strip()
    intent = str(case.get("intent") or "").strip()
    scenario = str(case.get("scenario") or "").strip()
    inputs = case.get("inputs") or {}

    keyword = str(inputs.get("keyword") or "").strip()
    if not case_id or not keyword:
        return {}

    params = {"keyword": keyword}
    t0 = time.time()
    debug_url: Optional[str] = None

    try:
        client = CozeWorkflowClient.from_env()
        output, dt_ms, debug_url = _call_workflow(client, workflow_id, params)
        duration_s = round(dt_ms / 1000.0, 3)

        # Parse output
        news_output, parse_error = _parse_news_output(output)

        # Run checks
        checks = _run_checks(news_output, parse_error)
        failed_checks = [c for c in checks if not c.passed]

        # Build source stats
        source_stats = {}
        if news_output:
            source_counts = news_output.source_counts()
            total_items = len(news_output.all_items())
            source_stats = {
                "total_items": total_items,
                "top_news_count": len(news_output.top_news),
                "quick_news_count": len(news_output.quick_news),
                "distinct_sources": len(source_counts),
                "source_distribution": dict(source_counts.most_common()),
            }

        if failed_checks:
            status = "FAIL"
            status_cn = "失败"
            issue = "; ".join(c.reason for c in failed_checks)
            root_cause = "; ".join(c.root_cause for c in failed_checks if c.root_cause)
            suggestion = "; ".join(c.suggestion for c in failed_checks if c.suggestion)
        else:
            status = "PASS"
            status_cn = "通过"
            issue = ""
            root_cause = ""
            suggestion = ""

        # Build indicator
        indicator_parts = []
        for c in checks:
            symbol = "通过" if c.passed else "失败"
            indicator_parts.append(f"{c.name}:{symbol}")
        indicator = "; ".join(indicator_parts)

        out = {
            "id": case_id,
            "intent": intent,
            "scenario": scenario,
            "params": params,
            "duration_ms": dt_ms,
            "duration_s": duration_s,
            "status": status,
            "status_cn": status_cn,
            "indicator": indicator,
            "issue": issue,
            "root_cause": root_cause,
            "suggestion": suggestion,
            "checks": [c.__dict__ for c in checks],
            "source_stats": source_stats,
            "output_preview": output[:500] if output else "",
            "output_full": output,
        }
        if debug_url and status != "PASS":
            out["debug_url"] = debug_url
        return out

    except Exception as e:
        duration_s = round((time.time() - t0), 3)
        status = "ERROR" if _is_infra_exception(e) else "FAIL"
        status_cn = "异常" if status == "ERROR" else "失败"

        out = {
            "id": case_id,
            "intent": intent,
            "scenario": scenario,
            "params": params,
            "duration_ms": int(duration_s * 1000),
            "duration_s": duration_s,
            "status": status,
            "status_cn": status_cn,
            "indicator": "",
            "issue": f"{type(e).__name__}: {e}",
            "root_cause": "工作流执行异常",
            "suggestion": "检查网络、API配置与工作流状态",
            "checks": [],
            "source_stats": {},
            "output_preview": "",
            "output_full": "",
        }
        if debug_url:
            out["debug_url"] = debug_url
        return out


def _aggregate_source_distribution(results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Aggregate source distribution across all cases."""
    total_counter: Counter = Counter()
    case_sources: Dict[str, List[str]] = {}

    for r in results:
        case_id = r.get("id", "unknown")
        stats = r.get("source_stats") or {}
        distribution = stats.get("source_distribution") or {}

        case_sources[case_id] = list(distribution.keys())
        for source, count in distribution.items():
            total_counter[source] += count

    total_items = sum(total_counter.values())
    distribution_pct = {}
    if total_items > 0:
        for source, count in total_counter.most_common():
            distribution_pct[source] = {
                "count": count,
                "percentage": round(count / total_items * 100, 1)
            }

    # Find 华尔街见闻 stats
    wallstreet_count = 0
    for source, count in total_counter.items():
        if "华尔街见闻" in source or "wallstreetcn" in source.lower():
            wallstreet_count += count
            break

    wallstreet_pct = round(wallstreet_count / total_items * 100, 1) if total_items > 0 else 0

    return {
        "total_items": total_items,
        "distinct_sources": len(total_counter),
        "source_distribution": distribution_pct,
        "case_sources": case_sources,
        "wallstreet_count": wallstreet_count,
        "wallstreet_percentage": wallstreet_pct,
    }


def _render_md(report: Dict[str, Any]) -> str:
    """Render Markdown report."""
    summary = report.get("summary") or {}
    run = report.get("run") or {}
    suite = report.get("suite") or {}
    agg_sources = report.get("aggregate_sources") or {}

    lines: List[str] = []
    lines.append("# gen_news 测试报告")
    lines.append("")
    lines.append(f"- Run ID: `{run.get('run_id')}`")
    lines.append(f"- 时间: `{run.get('timestamp')}`")
    lines.append(f"- Workflow ID: `{suite.get('workflow_id')}`")
    lines.append(f"- 版本: `{suite.get('version')}`")
    lines.append("")

    # Summary table
    lines.append("## 测试汇总")
    lines.append("")
    lines.append("| 指标 | 结果 |")
    lines.append("|------|------|")
    lines.append(f"| 总用例数 | {summary.get('total')} |")
    lines.append(f"| 通过 | {summary.get('passed')} |")
    lines.append(f"| 失败 | {summary.get('failed')} |")
    lines.append(f"| 异常 | {summary.get('errors')} |")
    lines.append(f"| 通过率 | {summary.get('pass_rate')} |")
    lines.append("")

    # Aggregate source distribution
    lines.append("## 信源分布汇总")
    lines.append("")
    lines.append(f"- 总条目数: {agg_sources.get('total_items')}")
    lines.append(f"- 不同信源数: {agg_sources.get('distinct_sources')}")
    lines.append(f"- 华尔街见闻: {agg_sources.get('wallstreet_count')}条 ({agg_sources.get('wallstreet_percentage')}%)")
    lines.append("")

    if agg_sources.get("source_distribution"):
        lines.append("| 信源 | 数量 | 占比 |")
        lines.append("|------|------|------|")
        for source, stats in list(agg_sources["source_distribution"].items())[:20]:
            lines.append(f"| {source} | {stats['count']} | {stats['percentage']}% |")
        lines.append("")

    # Per-case details
    lines.append("## 用例详情")
    for r in report.get("results", []) or []:
        lines.append("")
        lines.append(f"### {r.get('id')} - {r.get('scenario')}")
        lines.append(f"- 结果: {r.get('status_cn')}")
        lines.append(f"- 耗时: {r.get('duration_s')}s")
        lines.append(f"- 评估指标: {r.get('indicator')}")

        stats = r.get("source_stats") or {}
        if stats:
            lines.append(f"- 条目数: 共{stats.get('total_items')}条 (重要{stats.get('top_news_count')}, 快讯{stats.get('quick_news_count')})")
            lines.append(f"- 信源数: {stats.get('distinct_sources')}个")
            dist = stats.get("source_distribution") or {}
            if dist:
                dist_str = ", ".join(f"{s}:{c}" for s, c in list(dist.items())[:5])
                lines.append(f"- 信源分布: {dist_str}")

        if r.get("issue"):
            lines.append(f"- 问题: {r.get('issue')}")
        if r.get("root_cause"):
            lines.append(f"- 根因: {r.get('root_cause')}")
        if r.get("suggestion"):
            lines.append(f"- 建议: {r.get('suggestion')}")
        if r.get("debug_url"):
            lines.append(f"- Debug URL: {r.get('debug_url')}")

    return "\n".join(lines)


def main() -> int:
    # Ensure UTF-8 output
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    if not SUITE_PATH.exists():
        raise SystemExit(f"Suite file not found: {SUITE_PATH}")

    suite = _load_json_file(SUITE_PATH)
    suite_meta = suite.get("suite") or {}
    workflow_id = str(suite_meta.get("workflow_id") or "").strip()
    if not workflow_id:
        raise SystemExit("suite.workflow_id is required")

    cases = suite.get("cases")
    if not isinstance(cases, list) or not cases:
        raise SystemExit("suite.cases must be a non-empty list")

    # Validate Coze credentials early
    CozeWorkflowClient.from_env()

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_json_path = REPORT_DIR / f"gen_news_report_{run_id}.json"
    report_md_path = REPORT_DIR / f"gen_news_report_{run_id}.md"
    report_json_latest_path = REPORT_DIR / "gen_news_report.json"
    report_md_latest_path = REPORT_DIR / "gen_news_report.md"

    report: Dict[str, Any] = {
        "suite": suite_meta,
        "run": {
            "run_id": run_id,
            "timestamp": _now_ts(),
            "base_url": os.getenv("COZE_BASE_URL") or os.getenv("COZE_API_BASE") or "https://api.coze.cn",
        },
        "results": [],
        "summary": {},
        "aggregate_sources": {},
    }

    # Run cases with ThreadPoolExecutor
    max_workers = 2
    results_by_idx: Dict[int, Dict[str, Any]] = {}

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = {}
        for idx, case in enumerate(cases):
            fut = ex.submit(_run_one_case, workflow_id, case)
            futures[fut] = idx

        for fut in as_completed(futures):
            idx = futures[fut]
            try:
                result = fut.result() or {}
                results_by_idx[idx] = result
                case_id = result.get("id", f"case_{idx}")
                status = result.get("status_cn", "未知")
                _safe_print(f"Completed: {case_id} - {status}")
            except Exception as e:
                results_by_idx[idx] = {
                    "id": f"case_{idx}",
                    "intent": "",
                    "scenario": "(runner exception)",
                    "params": {},
                    "duration_ms": 0,
                    "duration_s": 0,
                    "status": "ERROR",
                    "status_cn": "异常",
                    "indicator": "",
                    "issue": f"{type(e).__name__}: {e}",
                    "root_cause": "测试脚本执行异常",
                    "suggestion": "检查测试脚本并发执行",
                    "checks": [],
                    "source_stats": {},
                    "output_preview": "",
                    "output_full": "",
                }
                _safe_print(f"Exception in case_{idx}: {e}")

    report["results"] = [results_by_idx.get(i, {}) for i in range(len(cases)) if results_by_idx.get(i)]

    # Calculate summary
    total = len(report["results"])
    passed = len([r for r in report["results"] if r.get("status") == "PASS"])
    failed = len([r for r in report["results"] if r.get("status") == "FAIL"])
    errors = len([r for r in report["results"] if r.get("status") == "ERROR"])
    denom = max(passed + failed, 1)
    pass_rate = f"{round((passed / denom) * 100, 1)}%"

    report["summary"] = {
        "total": total,
        "passed": passed,
        "failed": failed,
        "errors": errors,
        "pass_rate": pass_rate,
    }

    # Aggregate source distribution
    report["aggregate_sources"] = _aggregate_source_distribution(report["results"])

    # Strip output_full from JSON report to save space
    for r in report["results"]:
        if "output_full" in r:
            del r["output_full"]

    # Save reports
    _save_json(report_json_path, report)
    _save_json(report_json_latest_path, report)

    md = _render_md(report)
    report_md_path.write_text(md, encoding="utf-8")
    report_md_latest_path.write_text(md, encoding="utf-8")

    _safe_print(f"")
    _safe_print(f"Report JSON: {report_json_path}")
    _safe_print(f"Report MD: {report_md_path}")
    _safe_print(f"Summary: 通过={passed}, 失败={failed}, 异常={errors}, 通过率={pass_rate}")

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
