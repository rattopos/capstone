#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
템플릿 형식 데이터 테이블 HTML 생성기
- 실제 추출 데이터를 템플릿과 동일한 테이블 구조로 출력
- 현재분기 | 직전분기 | 작년동분기 구조
- 중복 컬럼 제거 (growth_rate 또는 change_rate만 사용)
"""

import sys
from datetime import datetime
from pathlib import Path
from typing import Any, cast

sys.path.insert(0, str(Path(__file__).parent))

from report_generator import ReportGenerator
from config.reports import SECTOR_REPORTS


class TemplateTableGenerator:
    """템플릿 형식의 데이터 테이블 생성"""
    
    def __init__(self, excel_path: str) -> None:
        self.generator = ReportGenerator(excel_path)
    
    def _fixed_growth_labels(self) -> list[str]:
        return ["2023. 3/4", "2024. 3/4", "2025. 2/4", "2025. 3/4"]
    
    def _fixed_index_labels(self) -> list[str]:
        return ["2024. 3/4", "2025. 3/4"]
    
    def _fixed_change_labels(self) -> list[str]:
        return ["2023. 3/4", "2024. 3/4", "2025. 2/4", "2025. 3/4"]
    
    def _fixed_rate_labels(self) -> list[str]:
        return ["2024. 3/4", "2025. 3/4", "15-29세"]
    
    def format_value(self, value: Any, decimals: int = 1) -> str:
        """숫자 포맷팅"""
        if value is None or value == '' or value == '-':
            return '-'
        try:
            v = float(value)
            return f"{v:.{decimals}f}"
        except:
            return str(value)
    
    def _render_header_cell(self, text: str, extra_class: str = "") -> str:
        class_attr = f" class=\"{extra_class}\"" if extra_class else ""
        return f"<th{class_attr}>{text}</th>"

    def _render_region_cells(self, row: dict[str, Any]) -> str:
        if row.get('group'):
            group = row.get('group')
            region = row.get('region', '')
            rowspan = row.get('rowspan', 1)
            return (
                f"<td class=\"region-group\" rowspan=\"{rowspan}\">{group}</td>"
                f"<td class=\"region-name\">{region}</td>"
            )
        region = row.get('region', '')
        if '전' in region and '국' in region:
            return f"<td colspan=\"2\">{region}</td>"
        return f"<td class=\"region-name\" colspan=\"2\">{region}</td>"

    def render_summary_table(self, report_id: str, report_name: str) -> str:
        """템플릿 구조와 동일한 증감률/지수 테이블 렌더링"""
        data: dict[str, Any] = self.generator.extract_data(report_id)
        summary_table = data.get('summary_table')
        if not summary_table:
            return f"<div class='report-section'><p>❌ {report_name}: summary_table 없음</p></div>"

        summary_table_dict = cast(dict[str, Any], summary_table)
        columns = cast(dict[str, Any], summary_table_dict.get('columns', {}))
        regions = cast(list[dict[str, Any]], summary_table_dict.get('regions', []))

        growth_cols = self._fixed_growth_labels()
        index_cols = self._fixed_index_labels()
        change_cols = self._fixed_change_labels() if columns.get('change_columns') else None
        rate_cols = self._fixed_rate_labels() if columns.get('rate_columns') else None

        # Case 1: Growth Rate & Index Table (e.g. Manufacturing, Service)
        if growth_cols and index_cols and not (change_cols and rate_cols):
            index_cols = index_cols[:2]
            html = f"""
<div class="report-section">
  <h2>{report_name}</h2>
  <div class="table-title">《 {report_name} 지수 및 증감률 》</div>
  <table class="data-table">
    <thead>
      <tr>
        <th rowspan="2">구분</th>
        <th rowspan="2"></th>
        <th colspan="4">전년동분기대비 증감률(%)</th>
        <th colspan="2" class="index-section">지수</th>
      </tr>
      <tr>
    """
            for col in growth_cols:
                html += self._render_header_cell(col)
            for col in index_cols:
                html += self._render_header_cell(col, "index-section")
            html += """
      </tr>
    </thead>
    <tbody>
    """
            for row in regions:
                html += "<tr>"
                html += self._render_region_cells(row)
                
                # Growth Rates
                growth_rates = list(row.get('growth_rates', []))
                growth_rates = (growth_rates + ['-'] * 4)[:4]
                for rate in growth_rates:
                    cell = self.format_value(rate)
                    html += f"<td>{cell}</td>"
                
                # Indices
                indices = list(row.get('indices', []))
                indices = (indices + ['-'] * 2)[:2]
                for idx in indices:
                    cell = self.format_value(idx)
                    html += f"<td class=\"index-section\">{cell}</td>"
                html += "</tr>"
            
            html += """
    </tbody>
  </table>
</div>
            """
            return html

        # Case 2: Change & Rate Table (e.g. Employment, Unemployment)
        if change_cols and rate_cols:
            rate_cols = rate_cols[:3]
            html = f"""
<div class="report-section">
  <h2>{report_name}</h2>
  <div class="table-title">《 {report_name} 및 증감 》</div>
  <table class="data-table">
    <thead>
      <tr>
        <th rowspan="2" colspan="2"></th>
        <th colspan="4">전년동분기대비 증감(%p)</th>
        <th colspan="3" class="rate-section">고용률(%)</th>
      </tr>
      <tr>
    """
            for col in change_cols:
                html += self._render_header_cell(col)
            for col in rate_cols:
                html += self._render_header_cell(col, "rate-section")
            html += """
      </tr>
    </thead>
    <tbody>
    """
            for row in regions:
                html += "<tr>"
                html += self._render_region_cells(row)
                
                # Changes
                changes = list(row.get('changes', []))
                changes = (changes + ['-'] * 4)[:4]
                for change in changes:
                    cell = self.format_value(change)
                    html += f"<td>{cell}</td>"
                
                # Rates
                rates = list(row.get('rates', []))
                rates = (rates + ['-'] * 3)[:3]
                for rate in rates:
                    cell = self.format_value(rate)
                    html += f"<td class=\"rate-section\">{cell}</td>"
                html += "</tr>"
            
            html += """
    </tbody>
  </table>
</div>
            """
            return html

        return f"<div class='report-section'><p>❌ {report_name}: 호환되는 표 형식이 아닙니다.</p></div>"

    def generate_full_html(self) -> str:
        """모든 섹션의 HTML 생성"""
        html = """
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>통합 데이터 테이블</title>
    <style>
        body { font-family: 'Malgun Gothic', sans-serif; margin: 20px; }
        .report-section { margin-bottom: 40px; border: 1px solid #ccc; padding: 20px; border-radius: 5px; }
        .table-title { font-size: 1.2em; font-weight: bold; margin-bottom: 10px; color: #333; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }
        th { background-color: #f2f2f2; }
        .region-group { font-weight: bold; background-color: #fafafa; }
        .region-name { text-align: left; }
        .index-section { background-color: #f9f9f9; }
        .rate-section { background-color: #eef; }
    </style>
</head>
<body>
    <h1>통합 데이터 테이블 (증감률/지수)</h1>
"""
        from config.reports import SECTOR_REPORTS, SUMMARY_REPORTS
        
        # 부문별
        for report in SECTOR_REPORTS:
             html += self.render_summary_table(report['id'], report['name'])
             
        html += "</body></html>"
        return html


def main():
    excel_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/분석표_25년 3분기_캡스톤(업데이트).xlsx"
    
    print("🚀 데이터 테이블 HTML 생성 시작...")
    print(f"   Excel: {Path(excel_path).name}")
    
    generator = TemplateTableGenerator(excel_path)
    html = generator.generate_full_html()
    
    output_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/exports/extracted_data_tables.html"
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"\n✅ HTML 생성 완료!")
    print(f"📄 저장위치: {output_path}")
    print(f"🌐 브라우저에서 열어보세요: file://{output_path}")


if __name__ == '__main__':
    main()
