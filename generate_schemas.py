#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
엑셀 파일을 분석하여 JSON Schema를 생성하는 스크립트
"""

import pandas as pd
import json
import re
from pathlib import Path
from typing import Dict, List, Any, Tuple

def clean_column_name(col: str) -> str:
    """컬럼명을 정리하고 영문 필드명으로 변환"""
    if pd.isna(col) or col == 'NaN' or str(col).strip() == '':
        return None
    
    col = str(col).strip()
    
    # 한글 컬럼명을 영문으로 변환하는 매핑
    korean_to_english = {
        '지역\n코드': 'region_code',
        '지역\n이름': 'region_name',
        '분류\n단계': 'category_level',
        '분류\n코드': 'category_code',
        '시군구 코드': 'city_code',
        '시군구 이름': 'city_name',
        '번호': 'number',
        '상품\n코드': 'product_code',
        '상품 이름': 'product_name',
        '공정 이름': 'process_name',
        '연령별 실업자 수(천 명)': 'unemployed_by_age',
        '시도별': 'sido',
        '연령계층별': 'age_group',
        '가중치': 'weight',
        '분류 이름': 'category_name',
        'Unnamed:': None  # Unnamed 컬럼은 None 반환
    }
    
    # 매핑에 있으면 변환
    for k, v in korean_to_english.items():
        if k in col:
            return v
    
    # 연도 패턴 (2010.0, 2010, 2025. 1 등)
    year_pattern = r'(\d{4})[.\s]*(\d{1,2})?[/\s]*(\d{1,2})?'
    match = re.match(year_pattern, col)
    if match:
        year = match.group(1)
        quarter = match.group(2) if match.group(2) else None
        month = match.group(3) if match.group(3) else None
        
        if month:
            return f'value_{year}_m{month.zfill(2)}'
        elif quarter:
            return f'value_{year}_q{quarter}'
        else:
            return f'value_{year}'
    
    # 그 외는 소문자로 변환하고 공백을 언더스코어로
    col_en = re.sub(r'[^\w\s]', '', col)
    col_en = re.sub(r'\s+', '_', col_en).lower()
    return col_en if col_en else None

def infer_data_type(series: pd.Series) -> str:
    """데이터 타입 추론"""
    # NaN이 아닌 값들만 확인
    non_null = series.dropna()
    
    if len(non_null) == 0:
        return 'null'
    
    # 숫자 타입 확인
    try:
        pd.to_numeric(non_null, errors='raise')
        # 정수인지 확인
        if non_null.apply(lambda x: isinstance(x, (int, float)) and float(x).is_integer()).all():
            return 'integer'
        return 'number'
    except:
        return 'string'

def analyze_sheet_structure(file_path: str, sheet_name: str) -> Dict[str, Any]:
    """시트 구조 분석"""
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # 헤더 행 찾기
    header_row = None
    for i in range(min(5, len(df))):
        row_values = df.iloc[i].astype(str).tolist()
        if any('지역' in str(v) or '코드' in str(v) or '이름' in str(v) or '분류' in str(v) 
               for v in row_values if pd.notna(v)):
            header_row = i
            break
    
    if header_row is None:
        # 첫 번째 행을 헤더로 사용
        header_row = 0
    
    # 데이터 시작 행 찾기
    data_start_row = header_row + 1
    for i in range(header_row + 1, min(header_row + 5, len(df))):
        row = df.iloc[i]
        if pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() not in ['', 'NaN', 'nan']:
            first_val = str(row.iloc[0]).strip()
            # 숫자나 지역코드 형태인지 확인
            if first_val.isdigit() or len(first_val) == 2:
                data_start_row = i
                break
    
    # 헤더 추출
    header = df.iloc[header_row].astype(str).tolist()
    
    # 실제 데이터 샘플 (처음 10행)
    sample_data = df.iloc[data_start_row:data_start_row+10] if data_start_row < len(df) else pd.DataFrame()
    
    return {
        'header_row': header_row,
        'data_start_row': data_start_row,
        'header': header,
        'sample_data': sample_data,
        'total_rows': len(df),
        'total_cols': len(df.columns)
    }

def generate_json_schema(file_path: str, sheet_name: str) -> Dict[str, Any]:
    """JSON Schema 생성"""
    structure = analyze_sheet_structure(file_path, sheet_name)
    
    # 실제 데이터 읽기
    df = pd.read_excel(file_path, sheet_name=sheet_name, 
                       header=structure['header_row'],
                       skiprows=structure['data_start_row'] - structure['header_row'] - 1)
    
    # 스키마 기본 구조
    schema = {
        "$schema": "http://json-schema.org/draft-07/schema#",
        "title": sheet_name,
        "description": f"스키마 정의: {sheet_name}",
        "type": "array",
        "items": {
            "type": "object",
            "properties": {},
            "required": []
        }
    }
    
    properties = {}
    required_fields = []
    
    # 각 컬럼 분석
    for col in df.columns:
        clean_name = clean_column_name(str(col))
        
        if clean_name is None:
            continue
        
        # 데이터 타입 추론
        dtype = infer_data_type(df[col])
        
        # 필드 정의
        field_def = {
            "type": dtype if dtype != 'null' else ["number", "string", "null"],
            "description": f"원본 컬럼명: {col}"
        }
        
        # 숫자 타입인 경우 추가 정보
        if dtype in ['number', 'integer']:
            non_null = pd.to_numeric(df[col].dropna(), errors='coerce')
            if len(non_null) > 0:
                field_def["minimum"] = float(non_null.min())
                field_def["maximum"] = float(non_null.max())
        
        properties[clean_name] = field_def
        
        # 필수 필드 판단 (null이 50% 미만인 경우)
        null_ratio = df[col].isna().sum() / len(df)
        if null_ratio < 0.5:
            required_fields.append(clean_name)
    
    schema["items"]["properties"] = properties
    schema["items"]["required"] = required_fields[:5]  # 필수 필드는 최대 5개만
    
    return schema

def main():
    file_path = '기초자료 수집표_2025년 2분기_캡스톤.xlsx'
    xl_file = pd.ExcelFile(file_path)
    
    all_schemas = {}
    schemas_dir = Path('schemas')
    schemas_dir.mkdir(exist_ok=True)
    
    print(f"총 {len(xl_file.sheet_names)}개 시트 분석 시작...")
    
    for sheet_name in xl_file.sheet_names:
        print(f"\n처리 중: {sheet_name}")
        try:
            schema = generate_json_schema(file_path, sheet_name)
            
            # 개별 파일 저장
            safe_name = re.sub(r'[^\w\s-]', '', sheet_name).strip().replace(' ', '_')
            schema_file = schemas_dir / f"{safe_name}.schema.json"
            
            with open(schema_file, 'w', encoding='utf-8') as f:
                json.dump(schema, f, ensure_ascii=False, indent=2)
            
            print(f"  ✓ 저장 완료: {schema_file}")
            
            all_schemas[sheet_name] = schema
            
        except Exception as e:
            print(f"  ✗ 오류 발생: {e}")
            import traceback
            traceback.print_exc()
    
    # 통합 스키마 파일 저장
    all_schemas_file = schemas_dir / "all_schemas.json"
    with open(all_schemas_file, 'w', encoding='utf-8') as f:
        json.dump(all_schemas, f, ensure_ascii=False, indent=2)
    
    print(f"\n✓ 모든 스키마 생성 완료!")
    print(f"  - 개별 파일: {len(xl_file.sheet_names)}개")
    print(f"  - 통합 파일: {all_schemas_file}")

if __name__ == "__main__":
    main()

