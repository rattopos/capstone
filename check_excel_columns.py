#!/usr/bin/env python3
import pandas as pd

# 광공업 집계 시트 확인
sheet_name = 'A(광공업생산)집계'
df = pd.read_excel('/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/양식/분석표_25년 3분기_캡스톤(업데이트).xlsx', sheet_name=sheet_name, header=None)

print("=" * 100)
print(f"Sheet: {sheet_name}")
print(f"Total rows: {len(df)}, Total columns: {len(df.columns)}")
print("=" * 100)
print("\nFirst 5 rows, all columns (showing headers):")
for i in range(min(5, len(df))):
    row_vals = [str(df.iloc[i, j])[:20] if pd.notna(df.iloc[i, j]) else 'NaN' for j in range(len(df.columns))]
    print(f"\nRow {i}:")
    for j, val in enumerate(row_vals):
        if j >= 7 and j <= 27:  # 관심있는 범위만
            print(f"  Col {j}: {val}")
