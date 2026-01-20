#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
정의된 표이름만 추출하는 최소 코드
"""
import sys
import pandas as pd
import openpyxl
from config.settings import BASE_DIR
from config.table_locations import load_table_locations

EXCEL_PATH = str((BASE_DIR / "분석표_25년 3분기_캡스톤(업데이트).xlsx").resolve())
TABLE_LOCATIONS = load_table_locations()

def extract_table_by_location(table_name: str) -> pd.DataFrame | None:
    info = TABLE_LOCATIONS.get(table_name)
    if not info:
        print(f"[오류] table_locations에 '{table_name}' 정보가 없습니다.")
        return None
    file = info.get("file", EXCEL_PATH)
    sheet = info.get("sheet")
    range_dict = info.get("range_dict")
    header = info.get("header_included", True)
    if not (file and sheet and range_dict):
        print(f"[오류] '{table_name}'의 파일/시트/범위 정보가 부족합니다.")
        return None
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb[sheet]
    from_col = range_dict["start_col"]
    from_row = range_dict["start_row"]
    to_col = range_dict["end_col"]
    to_row = range_dict["end_row"]
    data = []
    for row in ws[f"{from_col}{from_row}":f"{to_col}{to_row}"]:
        data.append([cell.value for cell in row])
    wb.close()
    if not data:
        return None
    df = pd.DataFrame(data)
    if header and len(df) > 1:
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
    return df

def save_table_html(table_name: str, output_path: str = None) -> str:
    df = extract_table_by_location(table_name)
    if df is None:
        print(f"[오류] '{table_name}' 데이터 추출 실패")
        return ""
    template = TABLE_LOCATIONS.get(table_name, {}).get("template", "")
    html = f"""
    <h2>{table_name}</h2>\n<p><strong>템플릿:</strong> {template}</p>\n{df.to_html(index=False, border=1, justify='left')}\n"""
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)
    return html

def main():
    import argparse
    parser = argparse.ArgumentParser(description="정의된 표이름만 추출하여 HTML로 저장 (최소코드)")
    parser.add_argument("--table", type=str, required=True, help="data_table_locations.md에 정의된 표이름")
    parser.add_argument("--output", type=str, default=None, help="HTML 저장 경로")
    args = parser.parse_args()

    html = save_table_html(args.table, args.output)
    if html:
        print(f"[완료] '{args.table}' 표를 HTML로 저장했습니다.")
    else:
        print(f"[실패] '{args.table}' 표 추출에 실패했습니다.")

if __name__ == "__main__":
    main()
