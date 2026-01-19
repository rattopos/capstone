#!/usr/bin/env python3
"""국내인구이동 컬럼 탐색 테스트 (빠른 검증)"""

import pandas as pd
from templates.base_generator import BaseGenerator

excel_path = 'uploads/분석표_25년_3분기_캡스톤업데이트_f1da33c3.xlsx'
sheet_name = 'I(순인구이동)집계'

print(f'\n=== {sheet_name} 시트 헤더 분석 ===')

# 헤더만 빠르게 읽기
df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=5)

print(f'시트 크기: {len(df)}행 × {len(df.columns)}열')
print('\n처음 5행 (헤더):')
for i in range(min(5, len(df))):
    print(f'\n행 {i}:')
    for j in range(min(45, len(df.columns))):
        val = df.iloc[i, j]
        if pd.notna(val):
            print(f'  열 {j}: {val}')

# BaseGenerator의 find_target_col_index 메서드 테스트
class QuickTester(BaseGenerator):
    def __init__(self):
        pass

tester = QuickTester()

print('\n\n=== 분기별 컬럼 찾기 ===')
for year, quarter in [(2025, 3), (2025, 2), (2025, 1), (2024, 4)]:
    col = tester.find_target_col_index(df, year, quarter, require_type_match=False, max_header_rows=5)
    print(f'{year}년 {quarter}분기: 열 {col}')
    if col is not None and col < len(df.columns):
        print(f'  헤더 값들:')
        for i in range(min(5, len(df))):
            val = df.iloc[i, col]
            if pd.notna(val):
                print(f'    행 {i}: {val}')

print('\n✅ 테스트 완료')
