import sys, io, json, re
sys.path.insert(0, 'tests')
# 复用测试脚本里的提取逻辑
from run_mail_push_suite import _normalize_llm_payload, _summary_char_len

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

try:
    with open('D:/daily-ai-news/tests/reports/mail_push_report_20260209_012241.json', 'r', encoding='utf-8') as f:
        report = json.load(f)
except FileNotFoundError:
    print("Report file not found")
    sys.exit(1)

print(f'{ "Case":<12} | { "Topic":<8} | { "Avg":<5} | { "Min":<3} | { "Max":<3} | { "Count"}')
print('-'*60)

total_lens = []

for r in report['results']:
    output = r.get('output', '')
    parsed, notes = _normalize_llm_payload(output)
    
    if not isinstance(parsed, dict):
        print(f"{r['id']:<12} | {r['topic']:<8} | Parse Fail | - | - | 0")
        continue
        
    top_news = parsed.get('top_news', [])
    if not isinstance(top_news, list):
        print(f"{r['id']:<12} | {r['topic']:<8} | No Top | - | - | 0")
        continue
        
    lens = []
    for item in top_news:
        if not isinstance(item, dict):
            continue
        summ = item.get('summary', '')
        l = _summary_char_len(summ)
        lens.append(l)
        total_lens.append(l)
        
    if lens:
        avg = sum(lens) / len(lens)
        print(f"{r['id']:<12} | {r['topic']:<8} | {avg:<5.1f} | {min(lens):<3} | {max(lens):<3} | {len(lens)}")
    else:
        print(f"{r['id']:<12} | {r['topic']:<8} | Empty | - | - | 0")

if total_lens:
    print(f'\nOverall Avg: {sum(total_lens)/len(total_lens):.1f}')
