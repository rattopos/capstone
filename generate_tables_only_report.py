import os
from bs4 import BeautifulSoup

TEMPLATE_DIR = "templates"
OUTPUT_FILE = "exports/통합_데이터표_보고서.html"

# 템플릿 파일 목록 (확장자 *_template.html)
template_files = [
    f for f in os.listdir(TEMPLATE_DIR)
    if f.endswith("_template.html")
]

all_tables = []

for filename in sorted(template_files):
    path = os.path.join(TEMPLATE_DIR, filename)
    with open(path, encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")
        # 모든 table 태그 추출
        tables = soup.find_all("table")
        for table in tables:
            # 테이블에 소속된 제목(섹션명) 추출 (h1~h3, .section-title 등)
            title = None
            for tag in ["h1", "h2", "h3", "div", "span"]:
                prev = table.find_previous(tag)
                if prev and ("title" in prev.get("class", []) or tag.startswith("h")):
                    title = prev.get_text(strip=True)
                    break
            all_tables.append((filename, title, str(table)))

# 통합 HTML 생성
table_html = ""
for fname, title, table in all_tables:
    table_html += f"<h2>{fname}"
    if title:
        table_html += f" - {title}"
    table_html += "</h2>\n"
    table_html += table + "\n"

html = f"""
<!DOCTYPE html>
<html lang='ko'>
<head>
    <meta charset='UTF-8'>
    <title>통합 데이터표 보고서</title>
    <style>
        table {{ border-collapse: collapse; margin-bottom: 32px; }}
        th, td {{ border: 1px solid #888; padding: 6px 12px; }}
        h2 {{ margin-top: 32px; }}
    </style>
</head>
<body>
    <h1>통합 데이터표 보고서</h1>
    {table_html}
</body>
</html>
"""

os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write(html)

print(f"통합 데이터표 보고서가 생성되었습니다: {OUTPUT_FILE}")
