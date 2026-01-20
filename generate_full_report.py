#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
전체 통합 보도자료 HTML 생성 스크립트
- 생성 순서: 부문별 → 시도별 → 요약 (데이터 재사용 캐시 목적)
- 출력 순서: 요약 → 부문별 → 시도별 (요청된 전체통합 순서)
- 단일 HTML로 합쳐 exports 폴더에 저장 (페이지 분리 없음)
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

from config.settings import EXPORT_FOLDER
from config.reports import SECTOR_REPORTS, REGIONAL_REPORTS, SUMMARY_REPORTS
from services.excel_cache import get_excel_file
from services.report_generator import generate_report_html, generate_regional_report_html
from utils.excel_utils import extract_year_quarter_from_excel


def _strip_chart_elements(html_content: str) -> str:
    if not html_content:
        return html_content

    chart_class_pattern = (
        r"chart-container|chart-wrapper|chart-area|graph-container|"
        r"chart-canvas-wrapper|chart-title|chart-image-converted|svg-image-converted"
    )

    html_content = re.sub(
        rf'<div[^>]*class=["\"][^"\"]*(?:{chart_class_pattern})[^"\"]*["\"][^>]*>.*?</div>',
        '',
        html_content,
        flags=re.DOTALL | re.IGNORECASE,
    )

    html_content = re.sub(r'<canvas[^>]*>.*?</canvas>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
    html_content = re.sub(r'<canvas[^>]*/?>', '', html_content, flags=re.IGNORECASE)
    html_content = re.sub(r'<svg[^>]*>.*?</svg>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
    html_content = re.sub(
        rf'<img[^>]*class=["\"][^"\"]*(?:{chart_class_pattern})[^"\"]*["\"][^>]*>',
        '',
        html_content,
        flags=re.DOTALL | re.IGNORECASE,
    )

    return html_content


def _strip_placeholders(html_content: str) -> str:
    if not html_content:
        return html_content

    html_content = re.sub(
        r'<span[^>]*class=["\"][^"\"]*\beditable-placeholder\b[^"\"]*["\"][^>]*>.*?</span>',
        '',
        html_content,
        flags=re.DOTALL | re.IGNORECASE,
    )
    html_content = re.sub(r'\[[^\]]*입력 필요\]', '', html_content)
    html_content = re.sub(r'\[\s*\]', '', html_content)

    return html_content


def _strip_page_wrapper(html_content: str) -> str:
    if not html_content:
        return html_content

    page_open_pattern = r'<div[^>]*class=["\"][^"\"]*\bpage\b[^"\"]*["\"][^>]*>'
    if not re.search(page_open_pattern, html_content, flags=re.IGNORECASE):
        return html_content

    html_content = re.sub(page_open_pattern, '', html_content, count=1, flags=re.IGNORECASE)
    html_content = re.sub(r'</div>\s*$', '', html_content.strip(), count=1)
    return html_content


def _add_table_inline_styles(html_content: str) -> str:
    html_content = re.sub(
        r'<table([^>]*)>',
        r'<table\1 style="border-collapse: collapse; width: 100%; margin: 10px 0; border-top: 2px solid #000000; border-bottom: 2px solid #000000; font-family: \'Malgun Gothic\', \'맑은 고딕\', \'Dotum\', \'돋움\', sans-serif; font-size: 11pt;">',
        html_content,
    )
    html_content = re.sub(
        r'<th([^>]*)>',
        r'<th\1 style="border-top: 1px solid #888; border-bottom: 1px solid #888; border-left: 1px solid #DDDDDD; border-right: 1px solid #DDDDDD; padding: 4px 6px; text-align: center; vertical-align: middle; background-color: #F5F7FA; font-weight: normal; font-family: \'Malgun Gothic\', \'맑은 고딕\', \'Dotum\', \'돋움\', sans-serif; font-size: 11pt;">',
        html_content,
    )
    html_content = re.sub(
        r'<td([^>]*)>',
        r'<td\1 style="border: 1px solid #DDDDDD; padding: 4px 6px; text-align: center; vertical-align: middle; font-family: \'Malgun Gothic\', \'맑은 고딕\', \'Dotum\', \'돋움\', sans-serif; font-size: 11pt;">',
        html_content,
    )
    html_content = re.sub(
        r'<h1([^>]*)>',
        r'<h1\1 style="font-family: \'Malgun Gothic\', \'맑은 고딕\', \'Dotum\', \'돋움\', sans-serif; font-size: 14pt; font-weight: bold; color: #000000; margin: 15px 0 10px 0;">',
        html_content,
    )
    html_content = re.sub(
        r'<h2([^>]*)>',
        r'<h2\1 style="font-family: \'Malgun Gothic\', \'맑은 고딕\', \'Dotum\', \'돋움\', sans-serif; font-size: 13pt; font-weight: bold; color: #000000; margin: 15px 0 10px 0;">',
        html_content,
    )
    html_content = re.sub(
        r'<h3([^>]*)>',
        r'<h3\1 style="font-family: \'Malgun Gothic\', \'맑은 고딕\', \'Dotum\', \'돋움\', sans-serif; font-size: 11.5pt; font-weight: bold; color: #000000; border-bottom: 1px solid #000000; padding-bottom: 3px; margin: 10px 0 8px 0;">',
        html_content,
    )
    html_content = re.sub(
        r'<h4([^>]*)>',
        r'<h4\1 style="font-family: \'Malgun Gothic\', \'맑은 고딕\', \'Dotum\', \'돋움\', sans-serif; font-size: 11pt; font-weight: bold; color: #000000; margin: 10px 0 5px 0;">',
        html_content,
    )
    html_content = re.sub(
        r'<p([^>]*)>',
        r'<p\1 style="font-family: \'Malgun Gothic\', \'맑은 고딕\', \'Dotum\', \'돋움\', sans-serif; font-size: 10pt; margin: 5px 0; line-height: 1.5;">',
        html_content,
    )
    html_content = re.sub(
        r'<ul([^>]*)>',
        r'<ul\1 style="margin: 10px 0 10px 25px; font-family: \'Malgun Gothic\', \'맑은 고딕\', \'Dotum\', \'돋움\', sans-serif; font-size: 10pt;">',
        html_content,
    )
    html_content = re.sub(
        r'<ol([^>]*)>',
        r'<ol\1 style="margin: 10px 0 10px 25px; font-family: \'Malgun Gothic\', \'맑은 고딕\', \'Dotum\', \'돋움\', sans-serif; font-size: 10pt;">',
        html_content,
    )
    html_content = re.sub(
        r'<td([^>]*style="[^"]*)"([^>]*)>(-?\d+[\.%]?)',
        r'<td\1 text-align: right; padding-right: 4px;"\2>\3',
        html_content,
    )
    html_content = re.sub(
        r'<li([^>]*)>',
        r'<li\1 style="margin: 3px 0;">',
        html_content,
    )
    return html_content


def _extract_body_content(html: str) -> str:
    match = re.search(r'<body[^>]*>(.*?)</body>', html, re.DOTALL | re.IGNORECASE)
    return match.group(1) if match else html


def _sanitize_page_html(page_html: str) -> str:
    body_content = _extract_body_content(page_html)
    body_content = re.sub(r'<style[^>]*>.*?</style>', '', body_content, flags=re.DOTALL | re.IGNORECASE)
    body_content = re.sub(r'<script[^>]*>.*?</script>', '', body_content, flags=re.DOTALL | re.IGNORECASE)
    body_content = re.sub(r'<link[^>]*>', '', body_content)
    body_content = re.sub(r'<meta[^>]*>', '', body_content)
    body_content = _strip_chart_elements(body_content)
    body_content = _strip_placeholders(body_content)
    body_content = _strip_page_wrapper(body_content)
    body_content = _add_table_inline_styles(body_content)

    return body_content


def _build_final_html(pages: list[dict[str, str]], year: int, quarter: int) -> str:
    final_html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{year}년 {quarter}/4분기 지역경제동향</title>
    <style>
        body {{
            font-family: 'Malgun Gothic', '맑은 고딕', 'Dotum', '돋움', sans-serif;
            font-size: 10pt;
            line-height: 1.5;
            color: #000;
            background: #fff;
            padding: 0;
            margin: 0;
        }}
        .copy-btn {{
            position: fixed;
            top: 10px;
            right: 10px;
            background: #0066cc;
            color: white;
            padding: 12px 20px;
            border: none;
            border-radius: 5px;
            font-size: 12pt;
            cursor: pointer;
            z-index: 9999;
            box-shadow: 0 2px 10px rgba(0,0,0,0.3);
        }}
        .copy-btn:hover {{ background: #0055aa; }}
        @media print {{ .copy-btn {{ display: none; }} }}
    </style>
</head>
<body>
    <button class="copy-btn" onclick="copyAll()">📋 전체 복사 (클릭)</button>
    <div id="hwp-content">
'''

    for idx, page in enumerate(pages, 1):
        page_title = page.get('title', f'페이지 {idx}')
        page_html = page.get('html', '')
        body_content = _sanitize_page_html(page_html)

        if idx > 1:
            final_html += '\n<div style="height: 1em;"></div>\n'

        final_html += f'''
        <!-- 페이지 {idx}: {page_title} -->
{body_content}
'''

    final_html += '''
    </div>
    <script>
        function copyAll() {
            const content = document.getElementById('hwp-content');
            const range = document.createRange();
            range.selectNodeContents(content);
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
            try {
                document.execCommand('copy');
                alert('복사 완료!\n\n한글(HWP)에서 Ctrl+V로 붙여넣기 하세요.\n※ 표와 서식이 유지됩니다.');
            } catch (e) {
                alert('자동 복사 실패.\nCtrl+A로 전체 선택 후 Ctrl+C로 복사하세요.');
            }
            selection.removeAllRanges();
        }
        document.addEventListener('keydown', function(e) {
            if (e.ctrlKey && e.key === 'a') {
                e.preventDefault();
                copyAll();
            }
        });
    </script>
</body>
</html>
'''
    return final_html


def _resolve_period(excel_path: str, year: int | None, quarter: int | None) -> tuple[int, int]:
    if year is not None and quarter is not None:
        return year, quarter
    y, q = extract_year_quarter_from_excel(excel_path)
    if y is None or q is None:
        raise ValueError('연도/분기 정보를 확인할 수 없습니다. --year, --quarter를 지정하세요.')
    return y, q


def _generate_pages(excel_path: str, year: int, quarter: int):
    excel_file = get_excel_file(excel_path, use_data_only=True)
    sector_pages = []
    regional_pages = []
    summary_pages = []
    errors = []

    for report in SECTOR_REPORTS:
        report_id = report.get('id')
        report_name = report.get('name', report_id)
        html_content, error, _ = generate_report_html(
            excel_path,
            report,
            year,
            quarter,
            None,
            excel_file=excel_file,
        )
        if error or html_content is None:
            errors.append({'report_id': report_id, 'name': report_name, 'error': str(error)})
            continue
        sector_pages.append({'title': report_name, 'report_id': report_id, 'html': html_content})

    for region in REGIONAL_REPORTS:
        region_name = region.get('name', region.get('id', 'Unknown'))
        html_content, error = generate_regional_report_html(
            excel_path,
            region_name,
            is_reference=False,
            year=year,
            quarter=quarter,
            excel_file=excel_file,
        )
        if error or html_content is None:
            errors.append({'report_id': region.get('id', 'Unknown'), 'name': f'시도별-{region_name}', 'error': str(error)})
            continue
        regional_pages.append({'title': f'시도별-{region_name}', 'report_id': region.get('id', ''), 'html': html_content})

    for report in SUMMARY_REPORTS:
        report_id = report.get('id')
        report_name = report.get('name', report_id)
        html_content, error, _ = generate_report_html(
            excel_path,
            report,
            year,
            quarter,
            None,
            excel_file=excel_file,
        )
        if error or html_content is None:
            errors.append({'report_id': report_id, 'name': report_name, 'error': str(error)})
            continue
        summary_pages.append({'title': report_name, 'report_id': report_id, 'html': html_content})

    pages = summary_pages + sector_pages + regional_pages
    return pages, errors


def main() -> int:
    parser = argparse.ArgumentParser(description='전체 통합 보도자료 HTML 생성기')
    parser.add_argument('--excel', '-e', required=True, help='엑셀 파일 경로')
    parser.add_argument('--year', type=int, help='연도 (미지정 시 엑셀에서 추출)')
    parser.add_argument('--quarter', type=int, help='분기 (미지정 시 엑셀에서 추출)')
    parser.add_argument('--output', '-o', help='출력 HTML 경로 (미지정 시 exports 폴더)')
    args = parser.parse_args()

    excel_path = str(Path(args.excel).resolve())
    if not Path(excel_path).exists():
        print(f"[ERROR] 엑셀 파일을 찾을 수 없습니다: {excel_path}", file=sys.stderr)
        return 1

    try:
        year, quarter = _resolve_period(excel_path, args.year, args.quarter)
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return 1

    pages, errors = _generate_pages(excel_path, year, quarter)
    if not pages:
        print("[ERROR] 생성된 페이지가 없습니다.", file=sys.stderr)
        if errors:
            print(f"[ERROR] 실패 {len(errors)}건", file=sys.stderr)
        return 1

    final_html = _build_final_html(pages, year, quarter)

    if args.output:
        output_path = Path(args.output).resolve()
    else:
        output_path = EXPORT_FOLDER / f"지역경제동향_{year}년_{quarter}분기_통합.html"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(final_html, encoding='utf-8')

    print(f"✅ 통합 보도자료 생성 완료: {output_path}")
    if errors:
        print(f"⚠️ 일부 보고서 생성 실패: {len(errors)}건")
        for err in errors:
            print(f" - {err['name']}: {err['error']}")
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
