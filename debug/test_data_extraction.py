#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
데이터 추출 디버그 스크립트
"""
import sys
sys.path.insert(0, '/Users/topos/Desktop/capstone')

import pandas as pd
from pathlib import Path
from services.summary_data import get_summary_table_data, get_summary_overview_data
from templates.raw_data_extractor import RawDataExtractor
from extractors import DataExtractor

# 기초자료 파일 경로
RAW_PATH = '/Users/topos/Desktop/capstone/uploads/기초자료_수집표_2025년_3분기캡스톤_56a228c1.xlsx'

def test_summary_table():
    """요약 테이블 데이터 테스트"""
    print("=" * 60)
    print("1. 요약 테이블 데이터 테스트 (get_summary_table_data)")
    print("=" * 60)
    
    # year=2025, quarter=3으로 호출
    data = get_summary_table_data(RAW_PATH, year=2025, quarter=3)
    
    print(f"전국 데이터: {data.get('nationwide', {})}")
    print()
    
    # 첫 번째 지역 그룹 출력
    groups = data.get('region_groups', [])
    if groups:
        print(f"첫 번째 그룹: {groups[0].get('name')}")
        for region in groups[0].get('regions', [])[:3]:
            print(f"  - {region.get('name')}: 광공업={region.get('mining_production')}, 서비스업={region.get('service_production')}, 고용률={region.get('employment')}")

def test_regional_extractor():
    """시도별 데이터 추출 테스트"""
    print("\n" + "=" * 60)
    print("2. 시도별 데이터 추출 테스트 (RawDataExtractor)")
    print("=" * 60)
    
    extractor = RawDataExtractor(RAW_PATH, 2025, 3)
    
    # 서울 데이터 추출 테스트
    regional_data = extractor.extract_regional_data('서울')
    if regional_data:
        print(f"서울 데이터:")
        print(f"  report_info: {regional_data.get('report_info', {})}")
        
        # 고용이동 데이터
        emp_mig = regional_data.get('employment_migration', {})
        emp_rate = emp_mig.get('employment_rate', {})
        print(f"  고용률 데이터:")
        print(f"    total_change: {emp_rate.get('total_change')}")
        print(f"    direction: {emp_rate.get('direction')}")
        print(f"    increase_age_groups: {emp_rate.get('increase_age_groups', [])}")
        print(f"    decrease_age_groups: {emp_rate.get('decrease_age_groups', [])}")
        
        # summary_table
        summary_table = regional_data.get('summary_table', {})
        print(f"  summary_table rows 수: {len(summary_table.get('rows', []))}")
        if summary_table.get('rows'):
            print(f"    첫 번째 행: {summary_table['rows'][0]}")
    else:
        print("서울 데이터 추출 실패!")

def test_employment_sheet():
    """고용률 시트 직접 확인"""
    print("\n" + "=" * 60)
    print("3. 고용률 시트 직접 확인")
    print("=" * 60)
    
    xl = pd.ExcelFile(RAW_PATH)
    print(f"시트 목록: {xl.sheet_names}")
    
    if '고용률' in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name='고용률', header=None)
        print(f"\n고용률 시트 크기: {df.shape}")
        print(f"컬럼 수: {len(df.columns)}")
        
        # 헤더 행 출력 (2행, 3행)
        print(f"\n헤더 행 (2행):")
        header_row = 2
        for col_idx in range(60, min(70, len(df.columns))):
            val = df.iloc[header_row, col_idx]
            if pd.notna(val):
                print(f"  col {col_idx}: {val}")
        
        # 데이터 행 샘플
        print(f"\n서울 데이터 찾기:")
        for row_idx in range(len(df)):
            region = str(df.iloc[row_idx, 1]).strip() if pd.notna(df.iloc[row_idx, 1]) else ''
            if '서울' in region:
                level = df.iloc[row_idx, 2] if len(df.columns) > 2 else None
                name = df.iloc[row_idx, 3] if len(df.columns) > 3 else None
                print(f"  row {row_idx}: region={region}, level={level}, name={name}")
                
                # 2025년 3분기 컬럼 (67번)과 전년동기 (63번)
                if len(df.columns) > 67:
                    curr = df.iloc[row_idx, 67]  # 2025_3Q
                    prev = df.iloc[row_idx, 63]  # 2024_3Q
                    print(f"    2024_3Q (col63): {prev}, 2025_3Q (col67): {curr}")
                    if pd.notna(curr) and pd.notna(prev):
                        diff = float(curr) - float(prev)
                        print(f"    차이: {diff:.1f}%p")
                break
    else:
        print("고용률 시트 없음!")
        
    if '연령별고용률' in xl.sheet_names:
        print("\n연령별고용률 시트 존재")
        df = pd.read_excel(xl, sheet_name='연령별고용률', header=None)
        print(f"크기: {df.shape}")
    else:
        print("\n연령별고용률 시트 없음")

def test_construction_data():
    """건설동향 데이터 테스트"""
    print("\n" + "=" * 60)
    print("4. 건설동향 데이터 테스트")
    print("=" * 60)
    
    extractor = DataExtractor(RAW_PATH, 2025, 3)
    data = extractor.extract_construction_report_data()
    
    print(f"national_summary: {data.get('national_summary', {})}")
    print(f"nationwide_data: {data.get('nationwide_data', {})}")
    print(f"top3_increase: {data.get('top3_increase_regions', [])}")
    print(f"top3_decrease: {data.get('top3_decrease_regions', [])}")

def test_raw_sheet_columns():
    """기초자료 시트별 컬럼 확인"""
    print("\n" + "=" * 60)
    print("5. 기초자료 시트별 컬럼 확인")
    print("=" * 60)
    
    xl = pd.ExcelFile(RAW_PATH)
    
    sheets_to_check = ['광공업생산', '서비스업생산', '소비(소매, 추가)', '고용률', '품목성질별 물가']
    
    for sheet_name in sheets_to_check:
        if sheet_name not in xl.sheet_names:
            print(f"\n{sheet_name}: 시트 없음")
            continue
        
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        print(f"\n{sheet_name}:")
        print(f"  컬럼 수: {len(df.columns)}")
        
        # 헤더 행에서 2025.3/4 컬럼 찾기
        header_row = 2
        found_cols = []
        for col_idx in range(len(df.columns)):
            val = df.iloc[header_row, col_idx]
            if pd.notna(val):
                val_str = str(val)
                if '2025' in val_str and ('3' in val_str or '/4' in val_str):
                    found_cols.append((col_idx, val_str))
        
        if found_cols:
            print(f"  2025년 3분기 관련 컬럼:")
            for col_idx, val in found_cols[-3:]:  # 마지막 3개만
                print(f"    col {col_idx}: {val}")
        
        # 전국 데이터 샘플
        print(f"  전국 데이터 확인:")
        for row_idx in range(min(20, len(df))):
            region = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ''
            region2 = str(df.iloc[row_idx, 1]).strip() if len(df.columns) > 1 and pd.notna(df.iloc[row_idx, 1]) else ''
            if '전국' in region or '전국' in region2:
                print(f"    row {row_idx}: col0={region}, col1={region2}")
                break

if __name__ == "__main__":
    test_summary_table()
    test_regional_extractor()
    test_employment_sheet()
    test_construction_data()
    test_raw_sheet_columns()
    print("\n" + "=" * 60)
    print("디버그 완료")
    print("=" * 60)
