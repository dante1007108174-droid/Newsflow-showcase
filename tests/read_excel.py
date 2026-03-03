import pandas as pd
import sys
import json

sys.stdout.reconfigure(encoding='utf-8')
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', 300)

# File 1: 评估方式
file1 = r'D:\daily-ai-news\docs\生成层模型选型—即时查询\生成层模型选型_即时查询_内容生成_V2.xlsx'
xls1 = pd.ExcelFile(file1)
print('=== 生成层模型选型_即时查询_内容生成_V2.xlsx ===')
print('Sheets:', xls1.sheet_names)

for sheet in xls1.sheet_names:
    df = pd.read_excel(xls1, sheet_name=sheet)
    print(f'\n--- {sheet} ---')
    print('Columns:', list(df.columns))
    print(df.to_string())
    print()
