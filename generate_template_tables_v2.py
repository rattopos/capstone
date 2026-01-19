#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
템플릿 형식 데이터 테이블 HTML 생성기 (v2)
- 실제 추출 데이터를 템플릿과 동일한 테이블 구조로 출력
- 헤더 형식을 요청대로 고정
"""

import sys
from datetime import datetime
from pathlib import Path
from typing import Any, cast

sys.path.insert(0, str(Path(__file__).parent))

from report_generator import ReportGenerator
from config.reports import SECTOR_REPORTS


class TemplateTableGenerator:
    def __init__(self, excel_path: str) -> None:
        self.generator = ReportGenerator(excel_path)

    def _fixed_growth_labels(self) -> list[str]:
        year_quarter = self.generator._year_quarter
        if not year_quarter:
            return ["{Y-2}. {Q}/4", "{Y-1}. {Q}/4", "{Y}. {Q-1}/4", "{Y}. {Q}/4"]
        year, quarter = year_quarter
        prev_q_year, prev_q = self._previous_quarter(year, quarter)
        return [
            f"{year-2}. {quarter}/4",
            f"{year-1}. {quarter}/4",
            f"{prev_q_year}. {prev_q}/4",
            f"{year}. {quarter}/4",
        ]

    def _fixed_index_labels(self) -> list[str]:
        year_quarter = self.generator._year_quarter
        if not year_quarter:
            return ["{Y-1}. {Q}/4", "{Y}. {Q}/4"]
        year, quarter = year_quarter
        return [f"{year-1}. {quarter}/4", f"{year}. {quarter}/4"]

    def _fixed_change_labels(self) -> list[str]:
        year_quarter = self.generator._year_quarter
        if not year_quarter:
            return ["{Y-2}. {Q}/4", "{Y-1}. {Q}/4", "{Y}. {Q-1}/4", "{Y}. {Q}/4"]
        year, quarter = year_quarter
        prev_q_year, prev_q = self._previous_quarter(year, quarter)
        return [
            f"{year-2}. {quarter}/4",
            f"{year-1}. {quarter}/4",
            f"{prev_q_year}. {prev_q}/4",
            f"{year}. {quarter}/4",
        ]

    def _age_label(self, report_id: str) -> str:
        if report_id == "employment":
            return "20-29세"
        if report_id == "unemployment":
            return "15-29세"
        return "15-29세"

    def _fixed_rate_labels(self, report_id: str) -> list[str]:
        year_quarter = self.generator._year_quarter
        age_label = self._age_label(report_id)
        if not year_quarter:
            return ["{Y-1}. {Q}/4", "{Y}. {Q}/4", age_label]
        year, quarter = year_quarter
        return [f"{year-1}. {quarter}/4", f"{year}. {quarter}/4", age_label]

    def _previous_quarter(self, year: int, quarter: int) -> tuple[int, int]:
        if quarter <= 1:
            return (year - 1, 4)
        return (year, quarter - 1)

    def format_value(self, value: Any, decimals: int = 1) -> str:
        if value is None or value == '' or value == '-':
            return '-'
        try:
            v = float(value)
            return f"{v:.{decimals}f}"
        except Exception:
            return str(value)

    def _to_float(self, value: Any) -> float | None:
        if value is None or value == '' or value == '-':
            return None
        try:
            return float(value)
        except Exception:
            return None

    def _compute_growth(self, current: float | None, previous: float | None) -> float | None:
        if current is None or previous is None:
            return None
        if previous == 0:
            return None
        return round((current - previous) / previous * 100, 1)

    def _first_numeric(self, row: dict[str, Any], keys: list[str]) -> float | None:
        for key in keys:
            if key in row:
                value = self._to_float(row.get(key))
                if value is not None:
                    return value
        return None

    def _enrich_historical_growth_rates(self, row: dict[str, Any]) -> None:
        slots = self._build_growth_slots(row)
        if any(v is not None for v in slots):
            row['computed_growth_rates'] = slots

    def _build_growth_slots(self, row: dict[str, Any]) -> list[float | None]:
        rate_keys = row.get('rate_quarterly_keys') or row.get('quarterly_keys')
        rate_values = row.get('rate_quarterly_values')
        growth_keys = row.get('quarterly_keys')
        growth_values = row.get('quarterly_growth_rates')
        if rate_keys and rate_values and self.generator._year_quarter:
            year, quarter = self.generator._year_quarter
            prev_q_year, prev_q = self._previous_quarter(year, quarter)
            target_keys = [
                f"{year-2} {quarter}/4",
                f"{year-1} {quarter}/4",
                f"{prev_q_year} {prev_q}/4",
                f"{year} {quarter}/4",
            ]
            mapping = {k: v for k, v in zip(rate_keys, rate_values)}
            mapped = [mapping.get(k) for k in target_keys]
            if any(v is not None for v in mapped):
                return mapped
        if growth_keys and growth_values and self.generator._year_quarter:
            year, quarter = self.generator._year_quarter
            prev_q_year, prev_q = self._previous_quarter(year, quarter)
            target_keys = [
                f"{year-2} {quarter}/4",
                f"{year-1} {quarter}/4",
                f"{prev_q_year} {prev_q}/4",
                f"{year} {quarter}/4",
            ]
            mapping = {k: v for k, v in zip(growth_keys, growth_values)}
            mapped = [mapping.get(k) for k in target_keys]
            if any(v is not None for v in mapped):
                return mapped
        current_value = self._first_numeric(row, ['value', 'current_value', 'index'])
        prev_value = self._first_numeric(row, ['prev_value', 'previous_year_value', 'previous_year_index'])
        prev_prev_value = self._first_numeric(row, ['prev_prev_value', 'previous_prev_value', 'two_years_ago_value'])
        prev_prev_prev_value = self._first_numeric(row, ['prev_prev_prev_value', 'three_years_ago_value'])

        indices = row.get('indices')
        if isinstance(indices, list):
            if current_value is None and len(indices) >= 2:
                current_value = self._to_float(indices[1])
            if prev_value is None and len(indices) >= 1:
                prev_value = self._to_float(indices[0])

        two_years_ago = self._compute_growth(prev_prev_value, prev_prev_prev_value)
        last_year = self._compute_growth(prev_value, prev_prev_value)
        previous_quarter = self._to_float(
            row.get('previous_quarter_growth')
            or row.get('prev_quarter_growth')
        )
        current = self._compute_growth(current_value, prev_value)

        return [two_years_ago, last_year, previous_quarter, current]

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
        rate_cols = self._fixed_rate_labels(report_id) if columns.get('rate_columns') else None

        if report_id in ("employment", "unemployment") and change_cols and rate_cols:
            html = f"""
<div class=\"report-section\">
  <h2>{report_name}</h2>
  <div class=\"table-title\">《 {report_name} 및 증감 》</div>
  <table class=\"data-table\">
    <thead>
      <tr>
        <th rowspan=\"2\" colspan=\"2\"></th>
        <th colspan=\"4\">전년동분기대비 증감(%p)</th>
        <th colspan=\"3\" class=\"rate-section\">{report_name}(%)</th>
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
                if row.get('group'):
                    group = row.get('group')
                    region = row.get('region', '')
                    rowspan = row.get('rowspan', 1)
                    html += f"<td class=\"region-group\" rowspan=\"{rowspan}\">{group}</td>"
                    html += f"<td class=\"region-name\">{region}</td>"
                else:
                    region = row.get('region', '')
                    if '전' in region and '국' in region:
                        html += f"<td colspan=\"2\">{region}</td>"
                    else:
                        html += f"<td class=\"region-name\">{region}</td>"
                changes = list(row.get('changes', []))
                changes = (changes + ['-'] * 4)[:4]
                for change in changes:
                    cell = self.format_value(change)
                    html += f"<td>{cell}</td>"
                rates = list(row.get('rates', []))
                rates = (rates + ['-'] * 3)[:3]
                youth_rate = row.get('youth_rate')
                if youth_rate not in (None, '', '-'):
                    rates[2] = youth_rate
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

        if growth_cols and index_cols:
            html = f"""
<div class=\"report-section\">
  <h2>{report_name}</h2>
  <div class=\"table-title\">《 {report_name} 지수 및 증감률 》</div>
  <table class=\"data-table\">
    <thead>
      <tr>
        <th rowspan=\"2\">구분</th>
        <th rowspan=\"2\"></th>
        <th colspan=\"4\">전년동분기대비 증감률(%)</th>
        <th colspan=\"2\" class=\"index-section\">지수</th>
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
                self._enrich_historical_growth_rates(row)
                html += "<tr>"
                html += self._render_region_cells(row)
                growth_rates = list(row.get('growth_rates', []))
                growth_rates = (growth_rates + ['-'] * 4)[:4]
                computed_growth = self._build_growth_slots(row)
                filled_rates: list[Any] = []
                for idx, rate in enumerate(growth_rates):
                    if rate in (None, '', '-') and idx < len(computed_growth):
                        filled_rates.append(computed_growth[idx])
                    else:
                        filled_rates.append(rate)
                for rate in filled_rates:
                    cell = self.format_value(rate)
                    html += f"<td>{cell}</td>"
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

        if change_cols and rate_cols:
            html = f"""
<div class=\"report-section\">
  <h2>{report_name}</h2>
  <div class=\"table-title\">《 {report_name} 및 증감 》</div>
  <table class=\"data-table\">
    <thead>
      <tr>
        <th rowspan=\"2\" colspan=\"2\"></th>
        <th colspan=\"4\">전년동분기대비 증감(%p)</th>
        <th colspan=\"3\" class=\"rate-section\">고용률(%)</th>
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
                if row.get('group'):
                    group = row.get('group')
                    region = row.get('region', '')
                    rowspan = row.get('rowspan', 1)
                    html += f"<td class=\"region-group\" rowspan=\"{rowspan}\">{group}</td>"
                    html += f"<td class=\"region-name\">{region}</td>"
                else:
                    region = row.get('region', '')
                    if '전' in region and '국' in region:
                        html += f"<td colspan=\"2\">{region}</td>"
                    else:
                        html += f"<td class=\"region-name\">{region}</td>"
                changes = list(row.get('changes', []))
                changes = (changes + ['-'] * 4)[:4]
                for change in changes:
                    cell = self.format_value(change)
                    html += f"<td>{cell}</td>"
                for rate in (list(row.get('rates', [])) + ['-'] * 3)[:3]:
                    cell = self.format_value(rate)
                    html += f"<td class=\"rate-section\">{cell}</td>"
                html += "</tr>"
            html += """
    </tbody>
  </table>
</div>
            """
            return html

        return f"<div class='report-section'><p>❌ {report_name}: 표 컬럼 없음</p></div>"

    def generate_full_html(self) -> str:
        test_reports = [(r['id'], r['name']) for r in SECTOR_REPORTS]

        html_header = """
<!DOCTYPE html>
<html>
<head>
    <meta charset=\"utf-8\">
    <title>지역경제동향 - 데이터 테이블</title>
    <style>
        body { font-family: '맑은 고딕', sans-serif; margin: 20px; background-color: #f5f5f5; }
        .page-title { font-size: 24pt; font-weight: bold; text-align: center; margin-bottom: 30px; color: #333; }
        .report-section { background-color: white; margin-bottom: 30px; padding: 20px; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
        .report-section h2 { font-size: 16pt; border-bottom: 2px solid #4CAF50; padding-bottom: 10px; margin-bottom: 15px; color: #333; }
        .data-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        .data-table th { background-color: #e6e0ec; border: 1px solid #999; padding: 8px; text-align: center; font-weight: normal; height: 25px; }
        .data-table td { border: 1px solid #ccc; padding: 8px; text-align: left; height: 22px; }
        .data-table tbody tr:nth-child(odd) { background-color: #f9f9f9; }
        .data-table tbody tr:hover { background-color: #f0f0f0; }
        .footer { text-align: center; color: #666; margin-top: 40px; font-size: 10pt; }
    </style>
</head>
<body>
    <div class=\"page-title\">🏢 지역경제동향 - 데이터 테이블</div>
    <div style=\"text-align: center; color: #666; margin-bottom: 20px;\">
        <p>"""
        year_quarter = self.generator._year_quarter
        if year_quarter:
            year, quarter = year_quarter
            html_header += f"{year}년 {quarter}분기 | 증감률 표"
        else:
            html_header += "{Y}년 {Q}분기 | 증감률 표"
        html_header += """</p>
    </div>
"""

        html_content = html_header
        for report_id, report_name in test_reports:
            try:
                html_content += self.render_summary_table(report_id, report_name)
            except Exception as e:
                html_content += f"<div class='report-section'><p>❌ {report_name}: {str(e)}</p></div>"

        html_footer = """
    <div class=\"footer\">
        <p>생성 일시: """ + datetime.now().strftime("%Y-%m-%d %H:%M:%S") + """</p>
        <p>자료: 지역경제동향 분석표 (2025년 3분기)</p>
    </div>
</body>
</html>
"""
        return html_content + html_footer


def main() -> None:
    excel_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/분석표_25년 3분기_캡스톤(업데이트).xlsx"
    generator = TemplateTableGenerator(excel_path)
    html = generator.generate_full_html()

    output_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/exports/extracted_data_tables.html"
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print("✅ HTML 생성 완료!")
    print(f"📄 저장위치: {output_path}")


if __name__ == '__main__':
    main()
