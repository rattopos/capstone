#!/usr/bin/env python3
"""국내인구이동 데이터 추출 및 전국 합산 테스트"""

import pandas as pd
import sys

excel_path = 'uploads/분석표_25년_3분기_캡스톤업데이트_f1da33c3.xlsx'
sheet_name = 'I(순인구이동)집계'

print(f'=== {sheet_name} 데이터 추출 테스트 ===\n')

# 전체 데이터 읽기
df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

# 헤더 행: 2 (0-based)
# 데이터 시작: 3
# 지역 열: 4
# 연령 열: 7
# 2025 3/4: 25, 2025 2/4: 24, 2025 1/4: 23, 2024 4/4: 22

region_col = 4
age_col = 7
target_col = 25  # 2025 3/4
prev_q_col = 24   # 2025 2/4
prev_prev_col = 23  # 2025 1/4
prev_prev_prev_col = 22  # 2024 4/4

regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
           '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']

# 각 지역의 "합계" 행 찾기
table_data = []

for region_name in regions:
    # 지역 필터
    region_filter = df[df.iloc[:, region_col].astype(str).str.strip() == region_name]
    
    if region_filter.empty:
        continue
    
    # "합계" 행 찾기 (연령 컬럼에서)
    total_row = region_filter[region_filter.iloc[:, age_col].astype(str).str.strip() == '합계']
    
    if total_row.empty:
        # 첫 번째 행 사용
        total_row = region_filter.head(1)
    
    if total_row.empty:
        continue
    
    row = total_row.iloc[0]
    
    # 값 추출
    try:
        value = float(row.iloc[target_col])
        prev_value = float(row.iloc[prev_q_col]) if pd.notna(row.iloc[prev_q_col]) else None
        prev_prev_value = float(row.iloc[prev_prev_col]) if pd.notna(row.iloc[prev_prev_col]) else None
        prev_prev_prev_value = float(row.iloc[prev_prev_prev_col]) if pd.notna(row.iloc[prev_prev_prev_col]) else None
    except (ValueError, TypeError, IndexError):
        continue
    
    table_data.append({
        'region_name': region_name,
        'value': value,
        'prev_value': prev_value,
        'prev_prev_value': prev_prev_value,
        'prev_prev_prev_value': prev_prev_prev_value
    })
    
    print(f'{region_name}: {value:>8.1f} | {prev_value:>8.1f} | {prev_prev_value:>8.1f} | {prev_prev_prev_value:>8.1f}')

print(f'\n추출된 지역 수: {len(table_data)}/17')

# 전국 합산
def sum_field(key):
    values = [row.get(key) for row in table_data if row.get(key) is not None]
    return round(sum(values), 1) if values else None

nationwide_value = sum_field('value')
nationwide_prev = sum_field('prev_value')
nationwide_prev_prev = sum_field('prev_prev_value')
nationwide_prev_prev_prev = sum_field('prev_prev_prev_value')

print(f'\n📊 전국 (지역 합산):')
print(f'  - 2025 3/4: {nationwide_value}')
print(f'  - 2025 2/4: {nationwide_prev}')
print(f'  - 2025 1/4: {nationwide_prev_prev}')
print(f'  - 2024 4/4: {nationwide_prev_prev_prev}')

print('\n✅ 테스트 완료')
