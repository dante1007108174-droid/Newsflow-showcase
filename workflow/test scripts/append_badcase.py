"""Append GEN-05 badcase to badcase合集.xlsx"""

from openpyxl import load_workbook
from pathlib import Path

input_path = Path(r'D:\daily-ai-news\docs\badcase合集.xlsx')

wb = load_workbook(input_path)
ws = wb['Sheet1']

# 添加 GEN-05 badcase
new_row = [
    "GEN-05",  # ID
    "来源过滤失效",  # 名称
    "用户指定 source=纽约时报 时，模型未提示'无匹配'，而是返回了其他来源（华尔街见闻、IT之家等）的新闻",  # 问题描述
    "LLM Prompt 没有空结果处理逻辑，当指定来源不存在时，模型会'凑数'返回其他来源内容",  # 根因
    "用户明确要求特定来源却收到无关内容，体验受损；来源过滤功能失去信任；可能涉及合规风险",  # 影响
    "在 Prompt 中增加空结果处理：当 source 有值但无匹配时，必须输出'暂无来自[source]的相关新闻'，禁止返回其他来源内容",  # 建议
    "🔴 待修复",  # 问题状态
]

ws.append(new_row)

wb.save(input_path)
print(f"已添加 GEN-05 badcase 到 {input_path}")

# 验证
import pandas as pd
df = pd.read_excel(input_path, sheet_name='Sheet1')
print(f"\n现在共有 {len(df)} 条 badcase")
print(df.to_string(index=False))
