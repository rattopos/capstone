#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
고용 관련 시트 구조 분석
"""
import sys
sys.path.insert(0, '/Users/topos/Desktop/capstone')

import pandas as pd

RAW_PATH = '/Users/topos/Desktop/capstone/uploads/기초자료_수집표_2025년_3분기캡스톤_56a228c1.xlsx'

xl = pd.ExcelFile(RAW_PATH)

print("=" * 70)
print("1. 고용 관련 시트 목록")
print("=" * 70)
employment_sheets = [s for s in xl.sheet_names if '고용' in s or '실업' in s]
print(f"고용 관련 시트: {employment_sheets}")

# 연령별고용률 시트 분석
print("\n" + "=" * 70)
print("2. 연령별고용률 시트 분석")
print("=" * 70)

if '연령별고용률' in xl.sheet_names:
    df = pd.read_excel(xl, sheet_name='연령별고용률', header=None)
    print(f"크기: {df.shape}")
    
    # 첫 5행 출력
    print(f"\n처음 5행:")
    for i in range(min(5, len(df))):
        row_data = []
        for j in range(min(10, len(df.columns))):
            val = df.iloc[i, j]
            if pd.notna(val):
                row_data.append(f"col{j}={val}")
        print(f"  row {i}: {', '.join(row_data)}")
    
    # 헤더 행 찾기
    print(f"\n헤더 행 (row 2):")
    for col_idx in range(len(df.columns)):
        val = df.iloc[2, col_idx] if 2 < len(df) else None
        if pd.notna(val) and str(val).strip():
            print(f"  col {col_idx}: {val}")
            if col_idx > 70:
                break
    
    # 서울 데이터 찾기
    print(f"\n서울 데이터 찾기:")
    for row_idx in range(len(df)):
        for col_idx in [0, 1, 2]:
            if col_idx < len(df.columns):
                val = df.iloc[row_idx, col_idx]
                if pd.notna(val) and '서울' in str(val):
                    print(f"  row {row_idx}: ", end='')
                    for j in range(min(10, len(df.columns))):
                        v = df.iloc[row_idx, j]
                        if pd.notna(v):
                            print(f"col{j}={v}, ", end='')
                    print()
                    break
        else:
            continue
        break

# 고용 시트 분석
print("\n" + "=" * 70)
print("3. 고용 시트 분석")
print("=" * 70)

if '고용' in xl.sheet_names:
    df = pd.read_excel(xl, sheet_name='고용', header=None)
    print(f"크기: {df.shape}")
    
    # 첫 5행 출력
    print(f"\n처음 5행:")
    for i in range(min(5, len(df))):
        row_data = []
        for j in range(min(10, len(df.columns))):
            val = df.iloc[i, j]
            if pd.notna(val):
                row_data.append(f"col{j}={val}")
        print(f"  row {i}: {', '.join(row_data)}")
    
    # 헤더 행에서 2025 컬럼 찾기
    print(f"\n2025년 관련 컬럼:")
    for col_idx in range(len(df.columns)):
        val = df.iloc[2, col_idx] if 2 < len(df) else None
        if pd.notna(val) and '2025' in str(val):
            print(f"  col {col_idx}: {val}")

    # 전국/서울 데이터 찾기
    print(f"\n전국/서울 데이터 찾기:")
    for row_idx in range(min(30, len(df))):
        for col_idx in [0, 1]:
            if col_idx < len(df.columns):
                val = df.iloc[row_idx, col_idx]
                if pd.notna(val) and ('전국' in str(val) or '서울' in str(val)):
                    print(f"  row {row_idx}: col0={df.iloc[row_idx, 0] if pd.notna(df.iloc[row_idx, 0]) else '-'}, col1={df.iloc[row_idx, 1] if len(df.columns) > 1 and pd.notna(df.iloc[row_idx, 1]) else '-'}, col2={df.iloc[row_idx, 2] if len(df.columns) > 2 and pd.notna(df.iloc[row_idx, 2]) else '-'}")
                    break
        else:
            continue

# 실업자 수 시트 분석
print("\n" + "=" * 70)
print("4. 실업자 수 시트 분석")
print("=" * 70)

if '실업자 수' in xl.sheet_names:
    df = pd.read_excel(xl, sheet_name='실업자 수', header=None)
    print(f"크기: {df.shape}")
    
    # 2025년 관련 컬럼 찾기
    print(f"\n2025년 관련 컬럼:")
    for col_idx in range(60, min(68, len(df.columns))):
        val = df.iloc[2, col_idx] if 2 < len(df) else None
        if pd.notna(val):
            print(f"  col {col_idx}: {val}")

# 소비(소매, 추가) 시트의 권역 데이터 문제 분석
print("\n" + "=" * 70)
print("5. 소비(소매, 추가) 시트 권역 데이터 분석")
print("=" * 70)

if '소비(소매, 추가)' in xl.sheet_names:
    df = pd.read_excel(xl, sheet_name='소비(소매, 추가)', header=None)
    
    # 권역 행 찾기
    print(f"\n권역 데이터 (수도권, 충청권 등):")
    for row_idx in range(len(df)):
        region = str(df.iloc[row_idx, 1]).strip() if pd.notna(df.iloc[row_idx, 1]) else ''
        if region in ['수도권', '충청권', '호남권', '대경권', '동남권', '강원제주']:
            level = df.iloc[row_idx, 2] if len(df.columns) > 2 and pd.notna(df.iloc[row_idx, 2]) else '-'
            col64 = df.iloc[row_idx, 64] if len(df.columns) > 64 and pd.notna(df.iloc[row_idx, 64]) else '-'
            col60 = df.iloc[row_idx, 60] if len(df.columns) > 60 and pd.notna(df.iloc[row_idx, 60]) else '-'
            print(f"  row {row_idx}: region={region}, level={level}, col60(prev)={col60}, col64(curr)={col64}")

print("\n완료!")
