#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
모든 정의된 표이름에 대해 HTML로 일괄 추출 (최소코드)
"""
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
    for table_name in TABLE_LOCATIONS.keys():
        output_path = f"{table_name}.html"
        print(f"[진행] {table_name} → {output_path}")
        save_table_html(table_name, output_path)
    print("[완료] 모든 표를 HTML로 저장했습니다.")

if __name__ == "__main__":
    main()
