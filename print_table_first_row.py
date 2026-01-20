#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
각 표의 실제 첫 행(칼럼명)과 첫 데이터 행을 출력 (범위 조정 참고용)
"""
import openpyxl
from config.settings import BASE_DIR
from config.table_locations import load_table_locations

EXCEL_PATH = str((BASE_DIR / "분석표_25년 3분기_캡스톤(업데이트).xlsx").resolve())
TABLE_LOCATIONS = load_table_locations()

for table_name, info in TABLE_LOCATIONS.items():
    file = info.get("file", EXCEL_PATH)
    sheet = info.get("sheet")
    range_dict = info.get("range_dict")
    if not (file and sheet and range_dict):
        print(f"[SKIP] {table_name}")
        continue
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
    print(f"\n[{table_name}] 범위: {from_col}{from_row}:{to_col}{to_row}")
    if data:
        print("  첫 행(칼럼):", data[0])
        if len(data) > 1:
            print("  두번째 행(샘플):", data[1])
        else:
            print("  데이터 없음")
    else:
        print("  데이터 없음")
