"""Re-run failed test cases with updated prompt."""

import json
import sys
from pathlib import Path

from coze_workflow_client import CozeWorkflowClient

# Configuration
WORKFLOW_ID = "7603678313790652467"
FAILED_CASES = [
    "GEN-05",  # 来源不存在-纽约时报
    "GEN-07",  # 多来源组合
    "GEN-15",  # 全部过期-无有效新闻
    "GEN-24",  # 英文输出-基础
    "GEN-25",  # 英文输出-含来源约束
    "GEN-27",  # 边界-空数据集
    "GEN-30",  # 复合-来源加条数加语言
]

def load_suite():
    """Load test suite from JSON."""
    suite_path = Path(__file__).parent / "instant_news_suite.json"
    with open(suite_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def load_rss_data():
    """Load RSS input data."""
    rss_path = Path(__file__).parent.parent / "docs" / "真实rss信息输入.txt"
    with open(rss_path, 'r', encoding='utf-8') as f:
        content = f.read().strip()
        data = json.loads(content)
        return data.get("news_list", [])

def run_single_case(client, case, rss_data):
    """Run a single test case."""
    case_id = case["id"]
    params = case["parameters"].copy()
    
    # Prepare news_list based on dataset
    dataset = case.get("dataset", "slim_12")
    if dataset == "slim_12":
        # Build slim_12 dataset
        news_list = build_slim_12(rss_data)
    elif dataset == "mock_empty":
        news_list = []
    elif dataset == "mock_all_expired":
        news_list = build_mock_all_expired()
    else:
        # For other mock datasets, use slim_12 as fallback
        news_list = build_slim_12(rss_data)
    
    # Set current_time if not provided
    if not params.get("current_time"):
        from datetime import datetime
        params["current_time"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    # Convert news_list to JSON string
    params["news_list"] = json.dumps(news_list, ensure_ascii=False)
    
    print(f"\n{'='*60}")
    print(f"Running: {case_id} - {case['name']}")
    print(f"Parameters: keyword={params.get('keyword')}, source={params.get('source')}, limit={params.get('limit')}, language={params.get('language')}")
    print(f"News list count: {len(news_list)}")
    
    try:
        result = client.run_workflow(
            workflow_id=WORKFLOW_ID,
            parameters=params,
            timeout_s=120
        )
        
        print(f"✓ Success!")
        print(f"Output preview:\n{result.output[:500]}...")
        return {
            "case_id": case_id,
            "status": "SUCCESS",
            "output": result.output,
            "debug_url": result.debug_url
        }
        
    except Exception as e:
        print(f"✗ Failed: {e}")
        return {
            "case_id": case_id,
            "status": "FAILED",
            "error": str(e)
        }

def build_slim_12(news_list):
    """Build slim_12 dataset from RSS data."""
    preferred = ["虎嗅", "华尔街见闻", "IT之家", "36氪", "钛媒体"]
    
    seen_titles = set()
    out = []
    
    def add(item):
        title = str(item.get("title") or "").strip()
        if not title or title in seen_titles:
            return
        seen_titles.add(title)
        out.append({
            "title": title,
            "source": str(item.get("source") or "").strip(),
            "time_code": str(item.get("time_code") or "").strip(),
            "link": str(item.get("link") or "").strip(),
            "description": str(item.get("description") or "").strip()[:280],
        })
    
    # 2 per preferred source
    for src in preferred:
        picked = 0
        for it in news_list:
            if str(it.get("source") or "").strip() != src:
                continue
            add(it)
            picked += 1
            if picked >= 2:
                break
    
    # Fill up to 12
    for it in news_list:
        if len(out) >= 12:
            break
        add(it)
    
    return out[:12]

def build_mock_all_expired():
    """Build all expired dataset."""
    from datetime import datetime, timedelta
    ct = datetime(2026, 2, 6, 12, 0)
    
    def t(hours_ago):
        return (ct - timedelta(hours=hours_ago)).strftime("%Y-%m-%d %H:%M")
    
    return [
        {
            "title": "三天前的AI新闻：某公司发布新功能",
            "source": "虎嗅",
            "time_code": t(72),
            "link": "https://example.com/old-1",
            "description": "该条已过期，用于测试时效过滤。",
        },
        {
            "title": "两天前的新闻：AI芯片供应链变化",
            "source": "华尔街见闻",
            "time_code": t(48),
            "link": "https://example.com/old-2",
            "description": "该条已过期，用于测试时效过滤。",
        },
        {
            "title": "一周前旧新闻：AI会议回顾",
            "source": "36氪",
            "time_code": t(168),
            "link": "https://example.com/old-3",
            "description": "该条已过期，用于测试时效过滤。",
        },
    ]

def main():
    print("="*60)
    print("Re-running Failed Test Cases")
    print("="*60)
    
    # Initialize client
    client = CozeWorkflowClient.from_env()
    
    # Load suite and RSS data
    suite = load_suite()
    rss_data = load_rss_data()
    
    # Find failed cases
    all_cases = {case["id"]: case for case in suite["suite"]["cases"]}
    
    results = []
    for case_id in FAILED_CASES:
        if case_id not in all_cases:
            print(f"⚠ Case {case_id} not found in suite")
            continue
        
        case = all_cases[case_id]
        result = run_single_case(client, case, rss_data)
        results.append(result)
    
    # Print summary
    print(f"\n{'='*60}")
    print("Summary")
    print("="*60)
    
    success_count = sum(1 for r in results if r["status"] == "SUCCESS")
    failed_count = len(results) - success_count
    
    print(f"Total: {len(results)}")
    print(f"Success: {success_count}")
    print(f"Failed: {failed_count}")
    
    # Save results
    output_path = Path(__file__).parent / "reports" / "failed_cases_rerun.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    
    print(f"\nResults saved to: {output_path}")
    
    return 0 if failed_count == 0 else 1

if __name__ == "__main__":
    sys.exit(main())
