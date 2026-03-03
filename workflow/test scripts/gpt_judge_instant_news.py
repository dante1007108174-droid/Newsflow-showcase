"""GPT-as-judge scorer for instant_news quality dimensions.

Uses LLM (OpenAI/DeepSeek) to evaluate quality dimensions:
- dedup: whether duplicate events are properly merged
- ranking: whether hot/important news is ranked first
- filter_precision: whether returned news is highly relevant to query

Score: 0-10, >=7 is considered passing.

Usage:
  python scripts/gpt_judge_instant_news.py

Outputs:
  - tests/reports/instant_news_gpt_scores.json
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
from dotenv import load_dotenv


ROOT = Path(__file__).resolve().parents[1]
REPORT_PATH = ROOT / "tests" / "reports" / "instant_news_report.json"
OUTPUT_PATH = ROOT / "tests" / "reports" / "instant_news_gpt_scores.json"


def _load_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _save_json(path: Path, data: Dict[str, Any]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _get_api_key() -> str:
    """Get OpenAI API key from env"""
    load_dotenv(ROOT / "backend" / ".env")
    key = os.getenv("OPENAI_API_KEY") or os.getenv("DEEPSEEK_API_KEY")
    if not key:
        raise RuntimeError("Missing OPENAI_API_KEY or DEEPSEEK_API_KEY in environment")
    return key


def _get_api_base() -> str:
    """Get API base URL"""
    return os.getenv("OPENAI_API_BASE", "https://api.openai.com/v1")


def _get_model() -> str:
    """Get model name"""
    return os.getenv("GPT_JUDGE_MODEL", "gpt-4o-mini")


@dataclass
class QualityScore:
    dedup: float  # 0-10
    ranking: float  # 0-10
    relevance: float  # 0-10
    reasoning: str


def _call_llm(prompt: str) -> str:
    """Call LLM API"""
    api_key = _get_api_key()
    api_base = _get_api_base()
    model = _get_model()
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3,
    }
    
    resp = requests.post(
        f"{api_base}/chat/completions",
        headers=headers,
        json=payload,
        timeout=60,
    )
    resp.raise_for_status()
    
    data = resp.json()
    return data["choices"][0]["message"]["content"]


def _parse_score(text: str) -> QualityScore:
    """Parse LLM response to extract scores"""
    # Try to find scores in format: dedup: 8, ranking: 7, relevance: 9
    pattern = r"(?:去重|dedup)[^:]*[:：]\s*(\d+(?:\.\d+)?).*?(?:排序|ranking)[^:]*[:：]\s*(\d+(?:\.\d+)?).*?(?:相关性|relevance)[^:]*[:：]\s*(\d+(?:\.\d+)?)"
    match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    
    if match:
        dedup = float(match.group(1))
        ranking = float(match.group(2))
        relevance = float(match.group(3))
    else:
        # Try alternative patterns
        scores = re.findall(r"(\d+(?:\.\d+)?)\s*分", text)
        if len(scores) >= 3:
            dedup, ranking, relevance = float(scores[0]), float(scores[1]), float(scores[2])
        else:
            # Default fallback
            dedup = ranking = relevance = 5.0
    
    # Clamp to 0-10
    dedup = max(0, min(10, dedup))
    ranking = max(0, min(10, ranking))
    relevance = max(0, min(10, relevance))
    
    return QualityScore(
        dedup=dedup,
        ranking=ranking,
        relevance=relevance,
        reasoning=text.strip(),
    )


def _build_judge_prompt(
    case_name: str,
    keyword: str,
    source: str,
    limit: str,
    raw_query: str,
    news_input: List[Dict],
    output: str,
    review_items: List[str],
) -> str:
    """Build prompt for GPT judge"""
    
    prompt_parts = [
        "你是一位专业的新闻质量评估专家。请对以下AI生成的新闻摘要进行质量评估。",
        "",
        "=== 评估维度 ===",
    ]
    
    if "dedup" in review_items:
        prompt_parts.append("1. 去重质量 (0-10分)：相同事件是否被正确合并，无重复报道")
    if "ranking" in review_items:
        prompt_parts.append("2. 排序质量 (0-10分)：热点/重要新闻是否排在前面")
    if "filter_precision" in review_items:
        prompt_parts.append("3. 相关性 (0-10分)：返回的新闻是否与查询主题高度相关")
    
    prompt_parts.extend([
        "",
        "=== 输入参数 ===",
        f"查询关键词: {keyword or '未指定'}",
        f"指定来源: {source or '未指定'}",
        f"条数限制: {limit or '未指定'}",
        f"原始查询: {raw_query or '未指定'}",
        "",
        "=== 原始新闻数据（输入）===",
        json.dumps(news_input[:8], ensure_ascii=False, indent=2),  # 限制输入长度
        "",
        "=== 模型输出（待评估）===",
        output[:2000] if len(output) > 2000 else output,  # 限制输出长度
        "",
        "=== 评分要求 ===",
        "请按以下格式输出评分结果：",
        "去重质量: X分 - 简要理由",
        "排序质量: X分 - 简要理由",
        "相关性: X分 - 简要理由",
        "",
        "评分标准：",
        "- 10分: 完美，无可挑剔",
        "- 7-9分: 良好，有 minor issues",
        "- 4-6分: 一般，有明显问题",
        "- 0-3分: 差，严重不符合要求",
        "",
        "请给出具体分数和简要理由（每维度50字以内）。",
    ])
    
    return "\n".join(prompt_parts)


def judge_case(
    case_result: Dict[str, Any],
    suite_case: Dict[str, Any],
    datasets: Dict[str, Any],
) -> Optional[QualityScore]:
    """Judge a single case"""
    
    expect = suite_case.get("expect", {})
    review_items = expect.get("review_items", [])
    
    if not review_items:
        return None  # No quality review needed
    
    # Get input data
    dataset_key = suite_case.get("dataset", "")
    news_input = []
    if dataset_key == "slim_12":
        news_input = datasets.get("slim_12", [])
    elif dataset_key == "mock_dedup":
        # Build from mock_dedup data
        news_input = _build_mock_dedup_data()
    elif dataset_key == "mock_ranking":
        news_input = _build_mock_ranking_data()
    
    params = case_result.get("params", {})
    output = case_result.get("output", "")
    
    if not output:
        return QualityScore(
            dedup=0, ranking=0, relevance=0,
            reasoning="输出为空，无法评估"
        )
    
    prompt = _build_judge_prompt(
        case_name=case_result.get("name", ""),
        keyword=params.get("keyword", ""),
        source=params.get("source", ""),
        limit=params.get("limit", ""),
        raw_query=params.get("raw_query", ""),
        news_input=news_input,
        output=output,
        review_items=review_items,
    )
    
    try:
        response = _call_llm(prompt)
        score = _parse_score(response)
        return score
    except Exception as e:
        return QualityScore(
            dedup=0, ranking=0, relevance=0,
            reasoning=f"评分失败: {str(e)}"
        )


def _build_mock_dedup_data() -> List[Dict]:
    """Build mock dedup dataset for context"""
    return [
        {"title": "DeepSeek发布新一代推理模型", "source": "IT之家", "time_code": "2026-02-06 10:00"},
        {"title": "DeepSeek新模型性能大幅提升", "source": "虎嗅", "time_code": "2026-02-06 10:30"},
        {"title": "DeepSeek推理模型引发热议", "source": "36氪", "time_code": "2026-02-06 11:00"},
    ]


def _build_mock_ranking_data() -> List[Dict]:
    """Build mock ranking dataset for context"""
    return [
        {"title": "谷歌发布Gemini新版本", "source": "IT之家", "time_code": "2026-02-06 10:00"},
        {"title": "亚马逊AI裁员", "source": "钛媒体", "time_code": "2026-02-06 11:00"},
        {"title": "独立开发者AI笔记应用", "source": "虎嗅", "time_code": "2026-02-06 06:00"},
    ]


def main() -> int:
    print("=" * 60)
    print("GPT-as-Judge 质量维度评分")
    print("=" * 60)
    
    if not REPORT_PATH.exists():
        raise SystemExit(f"报告不存在: {REPORT_PATH}")
    
    report = _load_json(REPORT_PATH)
    suite = _load_json(ROOT / "tests" / "instant_news_suite.json")
    datasets = report.get("datasets", {})
    
    suite_cases_by_id = {c.get("id"): c for c in suite.get("cases", [])}
    
    scores = {}
    needs_review_cases = [
        r for r in report.get("results", [])
        if r.get("status") == "NEEDS_REVIEW"
    ]
    
    print(f"\n找到 {len(needs_review_cases)} 个待评审用例")
    print("开始评分...\n")
    
    for i, case_result in enumerate(needs_review_cases, 1):
        case_id = case_result.get("id")
        suite_case = suite_cases_by_id.get(case_id, {})
        
        print(f"[{i}/{len(needs_review_cases)}] 评估用例: {case_id} - {case_result.get('name', '')}")
        
        score = judge_case(case_result, suite_case, datasets)
        
        if score:
            scores[case_id] = {
                "case_name": case_result.get("name", ""),
                "dedup": score.dedup,
                "ranking": score.ranking,
                "relevance": score.relevance,
                "average": round((score.dedup + score.ranking + score.relevance) / 3, 1),
                "passed": (score.dedup >= 7 and score.ranking >= 7 and score.relevance >= 7),
                "reasoning": score.reasoning,
            }
            print(f"  去重: {score.dedup}/10")
            print(f"  排序: {score.ranking}/10")
            print(f"  相关性: {score.relevance}/10")
            print(f"  综合: {scores[case_id]['average']}/10 - {'通过' if scores[case_id]['passed'] else '未通过'}")
        else:
            print(f"  跳过（无需评审）")
        print()
    
    # Save scores
    output_data = {
        "generated_at": datetime.now().isoformat(),
        "model": _get_model(),
        "scores": scores,
        "summary": {
            "total_reviewed": len(scores),
            "passed": sum(1 for s in scores.values() if s["passed"]),
            "failed": sum(1 for s in scores.values() if not s["passed"]),
        }
    }
    
    _save_json(OUTPUT_PATH, output_data)
    print(f"评分结果已保存: {OUTPUT_PATH}")
    print(f"\n汇总: {output_data['summary']['passed']} 通过, {output_data['summary']['failed']} 未通过")
    
    return 0


if __name__ == "__main__":
    from datetime import datetime
    raise SystemExit(main())
