#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
25년 3분기 분석표 엑셀 파일 구조 분석 스크립트

목적: 각 시트의 헤더 구조를 분석하여 동적 매핑 로직 개선에 활용
"""

import pandas as pd
import openpyxl
from pathlib import Path
import json

def analyze_excel_structure(excel_path):
    """엑셀 파일 구조 분석"""
    print("=" * 80)
    print(f"엑셀 파일 구조 분석: {Path(excel_path).name}")
    print("=" * 80)
    
    # openpyxl로 시트 목록 확인
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    sheet_names = wb.sheetnames
    
    print(f"\n📋 총 {len(sheet_names)}개 시트 발견\n")
    
    analysis_results = {}
    
    for idx, sheet_name in enumerate(sheet_names, 1):
        print(f"\n{'='*80}")
        print(f"[{idx}/{len(sheet_names)}] 시트: '{sheet_name}'")
        print(f"{'='*80}")
        
        try:
            # pandas로 시트 읽기 (헤더 없이)
            df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
            
            print(f"📊 크기: {df.shape[0]}행 × {df.shape[1]}열")
            
            # 헤더 행 찾기 (첫 10행 분석)
            print(f"\n🔍 헤더 분석 (첫 10행):")
            header_info = {}
            
            for row_idx in range(min(10, len(df))):
                row = df.iloc[row_idx]
                # 연도/분기 패턴 찾기
                year_quarter_cols = []
                region_cols = []
                industry_cols = []
                
                for col_idx, cell_value in enumerate(row):
                    if pd.notna(cell_value):
                        cell_str = str(cell_value)
                        
                        # 연도/분기 패턴
                        if any(year in cell_str for year in ['2023', '2024', '2025', '2026']):
                            year_quarter_cols.append((col_idx, cell_str))
                        
                        # 지역 관련
                        if any(keyword in cell_str for keyword in ['지역', '시도', '전국', '서울', '부산']):
                            region_cols.append((col_idx, cell_str))
                        
                        # 업종/산업 관련
                        if any(keyword in cell_str for keyword in ['업종', '산업', '품목', '공정']):
                            industry_cols.append((col_idx, cell_str))
                
                if year_quarter_cols or region_cols or industry_cols:
                    print(f"\n  행 {row_idx}:")
                    if year_quarter_cols:
                        print(f"    📅 연도/분기: {year_quarter_cols[:5]}")
                    if region_cols:
                        print(f"    🗺️  지역: {region_cols[:3]}")
                    if industry_cols:
                        print(f"    🏭 업종/산업: {industry_cols[:3]}")
                    
                    header_info[f"row_{row_idx}"] = {
                        "year_quarter": year_quarter_cols,
                        "region": region_cols,
                        "industry": industry_cols
                    }
            
            # 2025년 3분기 데이터 위치 찾기
            print(f"\n🎯 2025년 3분기 데이터 컬럼 찾기:")
            target_found = False
            for row_idx in range(min(10, len(df))):
                row = df.iloc[row_idx]
                for col_idx, cell_value in enumerate(row):
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip()
                        # 2025년 3분기 패턴
                        if '2025' in cell_str and ('3/4' in cell_str or '3분기' in cell_str):
                            print(f"  ✅ 발견! 행 {row_idx}, 컬럼 {col_idx}: '{cell_str}'")
                            target_found = True
            
            if not target_found:
                print(f"  ⚠️ 2025년 3분기 데이터 컬럼을 찾지 못했습니다.")
            
            # 샘플 데이터 (첫 5행)
            print(f"\n📝 샘플 데이터 (첫 5행, 첫 10열):")
            for row_idx in range(min(5, len(df))):
                row_data = [str(df.iloc[row_idx, col_idx])[:15] if pd.notna(df.iloc[row_idx, col_idx]) else 'NaN' 
                           for col_idx in range(min(10, len(df.columns)))]
                print(f"  행 {row_idx}: {row_data}")
            
            # 결과 저장
            analysis_results[sheet_name] = {
                "shape": df.shape,
                "header_info": header_info,
                "has_2025_q3": target_found
            }
            
        except Exception as e:
            print(f"❌ 시트 분석 실패: {e}")
            analysis_results[sheet_name] = {"error": str(e)}
    
    wb.close()
    
    # 결과 요약
    print(f"\n\n{'='*80}")
    print("📊 분석 결과 요약")
    print(f"{'='*80}")
    
    # 주요 보고서 시트 확인
    key_sheets = {
        "광공업생산": ["A(광공업생산)집계", "A 분석"],
        "서비스업생산": ["B(서비스업생산)집계", "B 분석"],
        "소비동향": ["C(소비)집계", "C 분석"],
        "건설수주": ["F'(건설)집계", "F'분석"],
        "수출": ["G(수출)집계", "G 분석"],
        "수입": ["H(수입)집계", "H 분석"],
        "물가": ["E(품목성질물가)집계", "E(품목성질물가)분석"],
        "고용률": ["D(고용률)집계", "D(고용률)분석"],
        "실업": ["D(실업)집계", "D(실업)분석"],
        "인구이동": ["I(순인구이동)집계"]
    }
    
    print("\n🔍 주요 보고서 시트 존재 여부:")
    for report_name, required_sheets in key_sheets.items():
        print(f"\n  {report_name}:")
        for sheet in required_sheets:
            exists = sheet in sheet_names
            status = "✅" if exists else "❌"
            print(f"    {status} {sheet}")
    
    # JSON으로 저장
    output_path = Path(excel_path).parent / "excel_structure_analysis.json"
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump({
            "file_name": Path(excel_path).name,
            "total_sheets": len(sheet_names),
            "sheet_names": sheet_names,
            "analysis_results": {k: {
                "shape": str(v.get("shape", "")),
                "has_2025_q3": v.get("has_2025_q3", False),
                "error": v.get("error")
            } for k, v in analysis_results.items()}
        }, f, ensure_ascii=False, indent=2)
    
    print(f"\n\n💾 분석 결과 저장: {output_path}")
    
    return analysis_results


if __name__ == "__main__":
    excel_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/uploads/분석표_25년_3분기_캡스톤업데이트_ee0197ea.xlsx"
    
    if Path(excel_path).exists():
        results = analyze_excel_structure(excel_path)
    else:
        print(f"❌ 파일을 찾을 수 없습니다: {excel_path}")
