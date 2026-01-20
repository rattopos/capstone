#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
단일 HTML 출력 생성기
- REPORT_ORDER에 정의된 모든 템플릿을 순서대로 렌더링하여 하나의 HTML 파일로 합칩니다.
- 페이지 나눔 없이 연속 출력
- '지수없음', '미제공', 'N/A' 등의 결측 표시가 포함되면 해당 행을 제거합니다.
"""
import re
from pathlib import Path

from config.settings import BASE_DIR, TEMPLATES_DIR
from config.reports import SECTOR_REPORTS
from report_generator import ReportGenerator

EXCEL_PATH = str((BASE_DIR / "분석표_25년 3분기_캡스톤(업데이트).xlsx").resolve())
OUTPUT_PATH = (BASE_DIR / "exports" / "지역경제동향_2025년_3분기_통합.html").resolve()

FORBIDDEN_TOKENS = ["지수없음", "미제공", "N/A", "None", "nan", "NaN"]


def extract_head_styles(html: str) -> list:
    """style 블록만 추출"""
    styles = re.findall(r"<style[^>]*>(.*?)</style>", html, flags=re.S | re.I)
    return styles or []


def extract_body_content(html: str) -> str:
    """body 내부 컨텐츠만 추출"""
    m = re.search(r"<body[^>]*>(.*?)</body>", html, flags=re.S | re.I)
    return m.group(1).strip() if m else html.strip()


def remove_forbidden_lines(html: str) -> str:
    """금지된 토큰을 포함하는 줄을 제거"""
    lines = html.splitlines()
    safe_lines = []
    for line in lines:
        if any(tok in line for tok in FORBIDDEN_TOKENS):
            continue
        safe_lines.append(line)
    return "\n".join(safe_lines)


def strip_page_wrappers(html: str) -> str:
        """섹션 내부의 .page 래퍼를 제거"""
        html = re.sub(r"<div\s+class=\"page\"\s*>", "<div class=\"section\">", html)
        html = re.sub(r"<div\s+class=\"page\s+([^\"]*)\"\s*>", r"<div class=\"section \1\">", html)
        return html


def build_single_html(styles: list, sections: list) -> str:
        head_styles = "\n\n".join(f"<style>\n{css}\n</style>" for css in styles)
        body = "\n\n".join(sections)
        return f"""<!DOCTYPE html>
<html lang=\"ko\">
<head>
    <meta charset=\"UTF-8\" />
    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />
    <title>지역경제동향 2025년 3분기 (통합)</title>
    {head_styles}
</head>
<body>
<div class=\"page\">
{body}
</div>
</body>
</html>"""


def main():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    generator = ReportGenerator(EXCEL_PATH)
    styles_acc = []
    sections_acc = []

    for report in SECTOR_REPORTS:
        rid = report["id"]
        html = generator.generate_html(rid)
        styles_acc.extend(extract_head_styles(html))
        section = extract_body_content(html)
        section = remove_forbidden_lines(section)
        section = strip_page_wrappers(section)
        sections_acc.append(section)

    single_html = build_single_html(styles_acc, sections_acc)
    Path(OUTPUT_PATH).write_text(single_html, encoding="utf-8")
    print(f"✅ 통합 HTML 생성 완료: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
