from docx import Document
import sys
sys.stdout.reconfigure(encoding='utf-8')
doc = Document(r'D:\daily-ai-news\docs\真实rss信息输入.docx')
for para in doc.paragraphs:
    if para.text.strip():
        print(para.text)
