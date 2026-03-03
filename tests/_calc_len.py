import sys, io, json, re

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

def count_chars(s):
    # Count non-whitespace characters (Chinese chars count as 1)
    return len(re.sub(r'\s+', '', s or ''))

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
    output_str = r.get('output', '')
    
    # 提取 JSON 中的 summary 字段
    # 匹配模式： "summary": "..." (处理转义引号)
    # 这里的正则假设 summary 内容中没有未转义的引号，这在标准 JSON 中是成立的
    summaries = re.findall(r'"summary":\s*"(.*?)(?<!\\)"', output_str)
    
    case_lens = []
    for s in summaries:
        # 处理 unicode 转义 (如 \u4e2d) 和普通转义 (\", \\, \n)
        try:
            # 补全为 JSON 字符串并解析，以处理所有转义
            decoded = json.loads(f'"{s}"')
            l = count_chars(decoded)
            if l > 10:  # 忽略太短的误匹配
                case_lens.append(l)
                total_lens.append(l)
        except:
            # Fallback: basic unescape
            clean = s.replace(r'\"', '"').replace(r'\\', '\\')
            l = count_chars(clean)
            if l > 10:
                case_lens.append(l)
                total_lens.append(l)

    if case_lens:
        avg = sum(case_lens) / len(case_lens)
        print(f"{r['id']:<12} | {r['topic']:<8} | {avg:<5.1f} | {min(case_lens):<3} | {max(case_lens):<3} | {len(case_lens)}")
    else:
        print(f"{r['id']:<12} | {r['topic']:<8} | N/A   | -   | -   | 0")

if total_lens:
    print(f'\nOverall Avg: {sum(total_lens)/len(total_lens):.1f}')
