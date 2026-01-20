import os
from pathlib import Path
from jinja2 import Environment, FileSystemLoader
from services.summary_data import get_summary_overview_data, get_summary_table_data
from config.settings import TEMPLATES_DIR


# 실제 데이터 엑셀 파일 경로 및 연도/분기
EXCEL_PATH = "분석표_25년 3분기_캡스톤(업데이트).xlsx"
YEAR = 2025
QUARTER = 3
OUTPUT_FILE = "exports/통합_데이터표_보고서(실데이터).html"

# 템플릿 파일 목록
TEMPLATE_FILES = [
    f for f in os.listdir(TEMPLATES_DIR)
    if f.endswith("_template.html")
]

# Jinja2 환경 설정
env = Environment(
    loader=FileSystemLoader(str(TEMPLATES_DIR)),
    autoescape=True
)


from config.reports import SECTOR_REPORTS
import pandas as pd
from config.table_locations import load_table_locations

def get_data_for_template(template_name):
    # 템플릿명으로 해당 부문 config 찾기
    sector_config = None
    for config in SECTOR_REPORTS:
        if config.get('template') == template_name:
            sector_config = config
            break
    # 요약/지역경제동향/summary류는 기존 방식 유지
    if sector_config is None:
        if "regional_economy_by_region" in template_name:
            data = {"table": get_summary_table_data(EXCEL_PATH, year=YEAR, quarter=QUARTER), "report_info": {"year": YEAR, "quarter": QUARTER}}
        elif "summary" in template_name:
            data = get_summary_overview_data(EXCEL_PATH, year=YEAR, quarter=QUARTER)
            if "report_info" not in data:
                data["report_info"] = {"year": YEAR, "quarter": QUARTER}
        else:
            data = {"report_info": {"year": YEAR, "quarter": QUARTER}}
        return data

    # 표 위치 정보 적용 (table_locations 기반)
    table_locations = load_table_locations()
    # config의 aggregation_structure/aggregation_range를 보장
    agg_struct = sector_config.get('aggregation_structure', {})
    agg_range = sector_config.get('aggregation_range', None)
    sheet_name = agg_struct.get('sheet')
    # pandas로 범위 추출
    df = None
    if sheet_name:
        try:
            df_full = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name, header=None)
            if agg_range:
                from openpyxl.utils import column_index_from_string
                def _col_to_index(col_value):
                    if col_value is None:
                        return None
                    if isinstance(col_value, int):
                        return col_value
                    if isinstance(col_value, str) and col_value.strip():
                        return column_index_from_string(col_value.strip().upper()) - 1
                    return None
                row_start = max((agg_range.get('start_row', 1) - 1), 0)
                row_end = agg_range.get('end_row', len(df_full))
                col_start = _col_to_index(agg_range.get('start_col'))
                col_end = _col_to_index(agg_range.get('end_col'))
                if col_end is not None:
                    col_end += 1
                df = df_full.iloc[row_start:row_end, col_start:col_end].copy()
            else:
                df = df_full.copy()
        except Exception as e:
            print(f"[ERROR] 표 추출 실패: {template_name}, {sheet_name}, {agg_range}, {e}")
            df = None
    # 데이터 매핑 결과 dict
    data = {"report_info": {"year": YEAR, "quarter": QUARTER}}
    if df is not None:
        # 헤더 포함 여부에 따라 첫 행을 컬럼명으로 지정
        if sector_config.get('header_included') and not df.empty:
            try:
                df.columns = df.iloc[0].tolist()
                df = df.iloc[1:].reset_index(drop=True)
            except Exception as e:
                print(f"[WARNING] 헤더 변환 실패: {e}")
        data['table_df'] = df.to_dict(orient='records')
        data['table_df_columns'] = list(df.columns)
    return data

def main():
    all_tables = []
    for filename in sorted(TEMPLATE_FILES):
        template = env.get_template(filename)
        data = get_data_for_template(filename)
        # 템플릿에서 table 태그만 추출
        rendered = template.render(**data)
        # table만 추출
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(rendered, "html.parser")
        tables = soup.find_all("table")
        for table in tables:
            # 제목 추출
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
    <title>통합 데이터표 보고서(실데이터)</title>
    <style>
        table {{ border-collapse: collapse; margin-bottom: 32px; }}
        th, td {{ border: 1px solid #888; padding: 6px 12px; }}
        h2 {{ margin-top: 32px; }}
    </style>
</head>
<body>
    <h1>통합 데이터표 보고서(실데이터)</h1>
    {table_html}
</body>
</html>
"""
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"실데이터 통합 데이터표 보고서가 생성되었습니다: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
