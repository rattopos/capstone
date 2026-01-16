#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
소비동향 보도자료 생성기
C 분석 시트에서 데이터를 추출하여 HTML 보도자료를 생성합니다.
"""

import pandas as pd
import json
from jinja2 import Template
from pathlib import Path
from templates.base_generator import BaseGenerator

# 업태명 매핑
BUSINESS_MAPPING = {
    '백화점': '백화점',
    '대형마트': '대형마트',
    '면세점': '면세점',
    '슈퍼마켓 및 잡화점': '슈퍼마켓·잡화점',
    '슈퍼마켓· 잡화점 및 편의점': '슈퍼마켓·잡화점·편의점',
    '편의점': '편의점',
    '승용차 및 연료 소매점': '승용차·연료소매점',
    '승용차 및 연료소매점': '승용차·연료소매점',
    '전문소매점': '전문소매점',
    '무점포 소매': '무점포소매'
}

# 지역명 매핑 (표 표시용)
REGION_DISPLAY_MAPPING = {
    '전국': '전 국',
    '서울': '서 울',
    '부산': '부 산',
    '대구': '대 구',
    '인천': '인 천',
    '광주': '광 주',
    '대전': '대 전',
    '울산': '울 산',
    '세종': '세 종',
    '경기': '경 기',
    '강원': '강 원',
    '충북': '충 북',
    '충남': '충 남',
    '전북': '전 북',
    '전남': '전 남',
    '경북': '경 북',
    '경남': '경 남',
    '제주': '제 주'
}

# 지역 그룹
REGION_GROUPS = {
    '경인': ['서울', '인천', '경기'],
    '충청': ['대전', '세종', '충북', '충남'],
    '호남': ['광주', '전북', '전남', '제주'],
    '동북': ['대구', '경북', '강원'],
    '동남': ['부산', '울산', '경남']
}

# 유효한 시도 목록
VALID_REGIONS = [
    '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
    '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주'
]


def safe_float(value, default=None):
    """안전한 float 변환 함수 (NaN, '-' 체크 포함)"""
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
        if isinstance(value, str):
            value = value.strip()
            if value == '-' or value == '':
                return default
        result = float(value)
        if pd.isna(result):
            return default
        return result
    except (ValueError, TypeError):
        return default


def find_sheet_with_fallback(xl, primary_sheets, fallback_sheets):
    """시트 찾기 - 기본 시트가 없으면 대체 시트 사용"""
    for sheet in primary_sheets:
        if sheet in xl.sheet_names:
            return sheet, False
    for sheet in fallback_sheets:
        if sheet in xl.sheet_names:
            print(f"[시트 대체] '{primary_sheets[0]}' → '{sheet}' (기초자료)")
            return sheet, True
    return None, False


def load_data(excel_path):
    """엑셀 파일에서 데이터 로드"""
    xl = pd.ExcelFile(excel_path)
    
    # 분석 시트 찾기
    analysis_sheet, use_raw = find_sheet_with_fallback(
        xl, 
        ['C 분석', 'C분석'],
        ['소비(소매, 추가)', '소비', '소매판매액지수']
    )
    
    if analysis_sheet:
        df_analysis = pd.read_excel(xl, sheet_name=analysis_sheet, header=None)
        # 분석 시트에 실제 데이터가 있는지 확인 (수식 미계산 체크)
        test_row = df_analysis[(df_analysis[3] == '전국') | (df_analysis[4] == '전국')]
        if test_row.empty or test_row.iloc[0].isna().sum() > 20:
            print(f"[소비동향] 분석 시트가 비어있음 → 집계 시트에서 직접 계산")
            use_aggregation_only = True
        else:
            use_aggregation_only = False
    else:
        raise ValueError(f"소비동향 분석 시트를 찾을 수 없습니다. 시트: {xl.sheet_names}")
    
    # 집계 시트 찾기
    agg_sheet, _ = find_sheet_with_fallback(
        xl,
        ['C(소비)집계', 'C 집계'],
        ['소비(소매, 추가)', '소비', '소매판매액지수']
    )
    
    if agg_sheet and agg_sheet != analysis_sheet:
        df_index = pd.read_excel(xl, sheet_name=agg_sheet, header=None)
    else:
        df_index = df_analysis.copy()
    
    # use_raw 정보를 첫번째 df에 속성으로 저장
    df_analysis.attrs['use_raw'] = use_raw
    df_analysis.attrs['use_aggregation_only'] = use_aggregation_only
    
    return df_analysis, df_index


def get_region_indices(df_analysis):
    """각 지역의 시작 인덱스 찾기"""
    region_indices = {}
    for i in range(len(df_analysis)):
        row = df_analysis.iloc[i]
        if row[7] == '총지수':
            region = row[3]
            if region in VALID_REGIONS or region == '전국':
                region_indices[region] = i
    return region_indices


def get_nationwide_data(df_analysis, df_index, generator: BaseGenerator, year: int = 2025, quarter: int = 2):
    """전국 데이터 추출
    
    Args:
        df_analysis: 분석 시트 DataFrame
        df_index: 집계 시트 DataFrame
        generator: BaseGenerator 인스턴스 (find_target_col_index 사용)
        year: 현재 연도
        quarter: 현재 분기
    """
    use_raw = df_analysis.attrs.get('use_raw', False)
    use_aggregation_only = df_analysis.attrs.get('use_aggregation_only', False)
    
    if use_raw:
        return _get_nationwide_from_raw_data(df_analysis, generator, year, quarter)
    
    # 집계 시트 기반으로 추출 (분석 시트가 비어있는 경우 포함)
    if use_aggregation_only:
        return _get_nationwide_from_aggregation(df_index, generator, year, quarter)
    
    # [Robust Dynamic Parsing System]
    # 헤더 행 찾기
    header_row_idx = 2
    if len(df_analysis) > header_row_idx:
        header_row = df_analysis.iloc[header_row_idx]
    else:
        header_row = df_analysis.iloc[0] if len(df_analysis) > 0 else pd.Series()
    
    # 동적으로 현재 분기 컬럼 찾기
    target_col = generator.find_target_col_index(header_row, year, quarter)
    
    # 분석 시트에서 전국 총지수 행
    try:
        nationwide_row = df_analysis.iloc[3]
        growth_rate = safe_float(nationwide_row[target_col], None)
        if growth_rate is None:
            raise ValueError(f"{year}년 {quarter}분기 증감률 데이터를 찾을 수 없습니다")
        growth_rate = round(growth_rate, 1)
    except (IndexError, KeyError, ValueError) as e:
        print(f"[소비동향] 분석 시트 데이터 읽기 실패: {e}")
        return _get_nationwide_from_aggregation(df_index, generator, year, quarter)
    
    # 집계 시트에서 전국 지수
    try:
        index_header_row = df_index.iloc[header_row_idx] if len(df_index) > header_row_idx else df_index.iloc[0]
        index_target_col = generator.find_target_col_index(index_header_row, year, quarter)
        
        index_row = df_index.iloc[3]
        sales_index = safe_float(index_row[index_target_col], None)
        if sales_index is None:
            sales_index = 100.0  # 기본값
    except (IndexError, KeyError):
        sales_index = 100.0
    
    # 전국 업태별 증감률
    businesses = []
    for i in range(4, 12):
        try:
            row = df_analysis.iloc[i]
            business_name = row[7]
            business_growth = safe_float(row[target_col], None)
            if business_growth is not None:
                businesses.append({
                    'name': BUSINESS_MAPPING.get(business_name, business_name),
                    'growth_rate': round(business_growth, 1)
                })
        except (IndexError, KeyError):
            continue
    
    # 감소율이 큰 순으로 정렬 (음수 중 절대값이 큰 것)
    negative_businesses = [b for b in businesses if b['growth_rate'] < 0]
    negative_businesses.sort(key=lambda x: x['growth_rate'])
    main_businesses = negative_businesses[:3]
    
    return {
        'sales_index': sales_index,
        'growth_rate': growth_rate,
        'main_businesses': main_businesses
    }


def _get_nationwide_from_aggregation(df_index, generator: BaseGenerator, year: int = 2025, quarter: int = 2):
    """집계 시트에서 전국 데이터 추출 (증감률 직접 계산)
    
    Args:
        df_index: 집계 시트 DataFrame
        generator: BaseGenerator 인스턴스 (find_target_col_index 사용)
        year: 현재 연도
        quarter: 현재 분기
    """
    # [Robust Dynamic Parsing System]
    # 헤더 행 찾기
    header_row_idx = 2
    if len(df_index) > header_row_idx:
        header_row = df_index.iloc[header_row_idx]
    else:
        header_row = df_index.iloc[0] if len(df_index) > 0 else pd.Series()
    
    # 동적으로 컬럼 찾기
    target_col = generator.find_target_col_index(header_row, year, quarter)
    prev_year_col = generator.find_target_col_index(header_row, year - 1, quarter)
    
    # 전국 총지수 행 찾기
    nationwide_rows = df_index[(df_index[2] == '전국') & (df_index[3].astype(str) == '0')]
    if nationwide_rows.empty:
        return {
            'sales_index': 100.0,
            'growth_rate': 0.0,
            'main_businesses': []
        }
    
    nationwide_total = nationwide_rows.iloc[0]
    
    # 당분기와 전년동분기 지수로 증감률 계산
    current_index = safe_float(nationwide_total[target_col], None)
    prev_year_index = safe_float(nationwide_total[prev_year_col], None)
    
    if current_index is None:
        current_index = 100.0
    if prev_year_index is None:
        prev_year_index = 100.0
    
    if prev_year_index and prev_year_index != 0:
        growth_rate = ((current_index - prev_year_index) / prev_year_index) * 100
    else:
        growth_rate = 0.0
    
    # 전국 업태별 데이터 (분류단계 1)
    nationwide_businesses = df_index[(df_index[2] == '전국') & (df_index[3].astype(str) == '1')]
    
    businesses = []
    for _, row in nationwide_businesses.iterrows():
        curr = safe_float(row[target_col], None)
        prev = safe_float(row[prev_year_col], None)
        
        if curr is not None and prev is not None and prev != 0:
            bus_growth = ((curr - prev) / prev) * 100
            businesses.append({
                'name': BUSINESS_MAPPING.get(str(row[6]) if pd.notna(row[6]) else '', str(row[6]) if pd.notna(row[6]) else ''),
                'growth_rate': round(bus_growth, 1)
            })
    
    # 감소율이 큰 순으로 정렬 (음수 중 절대값이 큰 것)
    negative_businesses = sorted([b for b in businesses if b['growth_rate'] < 0], 
                                key=lambda x: x['growth_rate'])[:3]
    
    return {
        'sales_index': current_index,
        'growth_rate': round(growth_rate, 1),
        'main_businesses': negative_businesses
    }


def _get_nationwide_from_raw_data(df, generator: BaseGenerator, year: int = 2025, quarter: int = 2):
    """기초자료 시트에서 전국 데이터 추출
    
    Args:
        df: 기초자료 시트 DataFrame
        generator: BaseGenerator 인스턴스 (find_target_col_index 사용)
        year: 현재 연도
        quarter: 현재 분기
    """
    # [Robust Dynamic Parsing System]
    # 헤더 행 찾기
    header_row_idx = 2
    for i in range(min(10, len(df))):
        row = df.iloc[i]
        row_str = ' '.join([str(v) for v in row.values[:10] if pd.notna(v)])
        if '지역' in row_str and ('2024' in row_str or '2025' in row_str):
            header_row_idx = i
            break
    
    header_row = df.iloc[header_row_idx] if header_row_idx < len(df) else df.iloc[0]
    
    # 동적으로 컬럼 찾기
    current_quarter_col = generator.find_target_col_index(header_row, year, quarter)
    prev_year_col = generator.find_target_col_index(header_row, year - 1, quarter)
    
    # 전국 총지수 행 찾기 (분류단계 0)
    nationwide_row = None
    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i]
        region = str(row[1]).strip() if pd.notna(row[1]) else ''
        classification = str(row[2]).strip() if pd.notna(row[2]) else ''
        if region == '전국' and classification == '0':
            nationwide_row = row
            break
    
    if nationwide_row is None:
        return {'sales_index': 100.0, 'growth_rate': 0.0, 'main_businesses': []}
    
    # 증감률 계산
    current_val = safe_float(nationwide_row[current_quarter_col], 100)
    prev_val = safe_float(nationwide_row[prev_year_col], 100)
    if prev_val and prev_val != 0:
        growth_rate = ((current_val - prev_val) / prev_val) * 100
    else:
        growth_rate = 0.0
    
    # 업태별 데이터 추출 (분류단계 1)
    businesses = []
    for i in range(header_row + 1, len(df)):
        row = df.iloc[i]
        region = str(row[1]).strip() if pd.notna(row[1]) else ''
        classification = str(row[2]).strip() if pd.notna(row[2]) else ''
        if region == '전국' and classification == '1':
            current = safe_float(row[current_quarter_col], None)
            prev = safe_float(row[prev_year_col], None)
            if current is not None and prev is not None and prev != 0:
                bus_growth = ((current - prev) / prev) * 100
                bus_name = str(row[4]) if pd.notna(row[4]) else ''
                businesses.append({
                    'name': BUSINESS_MAPPING.get(bus_name, bus_name),
                    'growth_rate': round(bus_growth, 1)
                })
    
    # 감소율이 큰 순으로 정렬
    negative_businesses = sorted([b for b in businesses if b['growth_rate'] < 0], 
                                key=lambda x: x['growth_rate'])[:3]
    
    return {
        'sales_index': current_val,
        'growth_rate': round(growth_rate, 1),
        'main_businesses': negative_businesses
    }


def get_regional_data(df_analysis, df_index):
    """시도별 데이터 추출"""
    use_raw = df_analysis.attrs.get('use_raw', False)
    use_aggregation_only = df_analysis.attrs.get('use_aggregation_only', False)
    
    # 집계 시트 기반으로 추출 (분석 시트가 비어있는 경우 포함)
    if use_aggregation_only:
        return _get_regional_from_aggregation(df_index)
    
    if use_raw:
        return _get_regional_from_raw_data(df_analysis)
    
    region_indices = get_region_indices(df_analysis)
    regions = []
    
    if not region_indices:
        # 지역 인덱스를 찾을 수 없는 경우 집계 시트에서 계산
        return _get_regional_from_aggregation(df_index)
    
    for region, start_idx in region_indices.items():
        if region == '전국':
            continue
            
        # 총지수 행에서 증감률 (컬럼 20: 2025.2/4)
        total_row = df_analysis.iloc[start_idx]
        try:
            growth_rate = round(float(total_row[20]), 1)
            growth_2023_2 = round(float(total_row[12]), 1)
            growth_2024_2 = round(float(total_row[16]), 1)
            growth_2025_1 = round(float(total_row[19]), 1)
        except:
            continue
        
        # 집계 시트에서 지수
        idx_row = df_index[df_index[2] == region]
        if not idx_row.empty:
            index_2024 = idx_row.iloc[0][21]
            index_2025 = idx_row.iloc[0][24]
        else:
            index_2024 = 0
            index_2025 = 0
        
        # 업태별 증감률
        businesses = []
        # 각 지역마다 업태 수가 다르므로 다음 총지수까지 찾기
        next_region_idx = len(df_analysis)
        for other_region, other_idx in region_indices.items():
            if other_idx > start_idx and other_idx < next_region_idx:
                next_region_idx = other_idx
        
        for i in range(start_idx + 1, min(start_idx + 10, next_region_idx)):
            if i >= len(df_analysis):
                break
            row = df_analysis.iloc[i]
            if row[4] != 1:  # 분류단계가 1이 아니면 스킵
                continue
            business_name = row[7]
            try:
                business_growth = float(row[20])
            except:
                continue
            
            businesses.append({
                'name': BUSINESS_MAPPING.get(business_name, business_name),
                'growth_rate': round(business_growth, 1)
            })
        
        # 증가 지역: 양수 증감률 순으로 정렬
        # 감소 지역: 음수 증감률 순으로 정렬 (절대값 큰 순)
        if growth_rate >= 0:
            positive_businesses = [b for b in businesses if b['growth_rate'] > 0]
            positive_businesses.sort(key=lambda x: x['growth_rate'], reverse=True)
            top_businesses = positive_businesses[:3]
        else:
            negative_businesses = [b for b in businesses if b['growth_rate'] < 0]
            negative_businesses.sort(key=lambda x: x['growth_rate'])
            top_businesses = negative_businesses[:3]
        
        regions.append({
            'region': region,
            'growth_rate': growth_rate,
            'growth_2023_2': growth_2023_2,
            'growth_2024_2': growth_2024_2,
            'growth_2025_1': growth_2025_1,
            'index_2024': index_2024,
            'index_2025': index_2025,
            'top_businesses': top_businesses,
            'all_businesses': businesses
        })
    
    # 증가/감소 지역 분류
    # 0.0인 지역은 완전히 제외 (None도 제외)
    increase_regions = sorted(
        [r for r in regions if r.get('growth_rate') is not None and r['growth_rate'] > 0],
        key=lambda x: x['growth_rate'],
        reverse=True
    )
    decrease_regions = sorted(
        [r for r in regions if r.get('growth_rate') is not None and r['growth_rate'] < 0],
        key=lambda x: x['growth_rate']
    )
    
    return {
        'increase_regions': increase_regions,
        'decrease_regions': decrease_regions,
        'all_regions': regions
    }


def _get_regional_from_aggregation(df_index):
    """집계 시트에서 시도별 데이터 추출"""
    # 집계 시트 구조: 2=지역이름, 3=분류단계, 5=업태코드, 6=업태종류
    # 데이터 컬럼: 16=2023.2/4, 20=2024.2/4, 23=2025.1/4, 24=2025.2/4
    individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                          '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
    
    regions = []
    
    for region in individual_regions:
        # 해당 지역 총지수 (분류단계 0)
        region_total = df_index[(df_index[2] == region) & (df_index[3].astype(str) == '0')]
        if region_total.empty:
            continue
        region_total = region_total.iloc[0]
        
        # 증감률 계산
        current = safe_float(region_total[24], 100)  # 2025.2/4
        prev = safe_float(region_total[20], 100)  # 2024.2/4
        prev_2023_2 = safe_float(region_total[16], 100)  # 2023.2/4
        prev_2024_2 = prev
        prev_2025_1 = safe_float(region_total[23], 100)  # 2025.1/4
        
        if prev and prev != 0:
            growth_rate = ((current - prev) / prev) * 100
        else:
            growth_rate = 0.0
        
        # 전년동분기비 계산
        growth_2023_2 = ((prev_2023_2 - safe_float(region_total[12], 100)) / safe_float(region_total[12], 100) * 100) if safe_float(region_total[12], 0) != 0 else 0.0
        growth_2024_2 = ((prev_2024_2 - prev_2023_2) / prev_2023_2 * 100) if prev_2023_2 != 0 else 0.0
        growth_2025_1 = ((prev_2025_1 - safe_float(region_total[19], 100)) / safe_float(region_total[19], 100) * 100) if safe_float(region_total[19], 0) != 0 else 0.0
        
        # 해당 지역 업태별 데이터 (분류단계 1)
        region_businesses = df_index[(df_index[2] == region) & (df_index[3].astype(str) == '1')]
        
        businesses = []
        for _, row in region_businesses.iterrows():
            curr = safe_float(row[24], None)
            prev_ind = safe_float(row[20], None)
            
            if curr is not None and prev_ind is not None and prev_ind != 0:
                bus_growth = ((curr - prev_ind) / prev_ind) * 100
                businesses.append({
                    'name': BUSINESS_MAPPING.get(str(row[6]) if pd.notna(row[6]) else '', str(row[6]) if pd.notna(row[6]) else ''),
                    'growth_rate': round(bus_growth, 1)
                })
        
        # 증가 지역: 양수 증감률 순으로 정렬
        # 감소 지역: 음수 증감률 순으로 정렬
        if growth_rate >= 0:
            sorted_bus = sorted([b for b in businesses if b['growth_rate'] > 0], 
                              key=lambda x: x['growth_rate'], reverse=True)
        else:
            sorted_bus = sorted([b for b in businesses if b['growth_rate'] < 0], 
                              key=lambda x: x['growth_rate'])
        
        regions.append({
            'region': region,
            'growth_rate': round(growth_rate, 1),
            'growth_2023_2': round(growth_2023_2, 1),
            'growth_2024_2': round(growth_2024_2, 1),
            'growth_2025_1': round(growth_2025_1, 1),
            'index_2024': prev,
            'index_2025': current,
            'top_businesses': sorted_bus[:3],
            'all_businesses': businesses
        })
    
    # 증가/감소 지역 분류
    # 0.0인 지역은 완전히 제외 (None도 제외)
    increase_regions = sorted(
        [r for r in regions if r.get('growth_rate') is not None and r['growth_rate'] > 0],
        key=lambda x: x['growth_rate'],
        reverse=True
    )
    decrease_regions = sorted(
        [r for r in regions if r.get('growth_rate') is not None and r['growth_rate'] < 0],
        key=lambda x: x['growth_rate']
    )
    
    return {
        'increase_regions': increase_regions,
        'decrease_regions': decrease_regions,
        'all_regions': regions
    }


def _get_regional_from_raw_data(df):
    """기초자료 시트에서 시도별 데이터 추출"""
    individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                          '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
    
    # 헤더 행 및 컬럼 인덱스 찾기
    header_row = 2
    current_quarter_col = None
    prev_year_col = None
    
    for i in range(min(10, len(df))):
        row = df.iloc[i]
        row_str = ' '.join([str(v) for v in row.values[:10] if pd.notna(v)])
        if '지역' in row_str and ('2024' in row_str or '2025' in row_str):
            header_row = i
            break
    
    header = df.iloc[header_row] if header_row < len(df) else df.iloc[0]
    
    for col_idx in range(len(header) - 1, 4, -1):
        col_val = str(header[col_idx]) if pd.notna(header[col_idx]) else ''
        if '2025' in col_val and ('2/4' in col_val or '2' in col_val):
            current_quarter_col = col_idx
        if '2024' in col_val and ('2/4' in col_val or '2' in col_val):
            prev_year_col = col_idx
        if current_quarter_col and prev_year_col:
            break
    
    if current_quarter_col is None:
        current_quarter_col = len(header) - 2
    if prev_year_col is None:
        prev_year_col = current_quarter_col - 4
    
    regions = []
    
    for region in individual_regions:
        # 해당 지역 총지수 (분류단계 0) 찾기
        region_row = None
        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            r_name = str(row[1]).strip() if pd.notna(row[1]) else ''
            classification = str(row[2]).strip() if pd.notna(row[2]) else ''
            if r_name == region and classification == '0':
                region_row = row
                break
        
        if region_row is None:
            continue
        
        # 증감률 계산
        current_val = safe_float(region_row[current_quarter_col], None)
        prev_val = safe_float(region_row[prev_year_col], None)
        
        if current_val is not None and prev_val is not None and prev_val != 0:
            growth_rate = ((current_val - prev_val) / prev_val) * 100
        else:
            growth_rate = 0.0
        
        # 해당 지역의 업태별 데이터 (분류단계 1)
        businesses = []
        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            r_name = str(row[1]).strip() if pd.notna(row[1]) else ''
            classification = str(row[2]).strip() if pd.notna(row[2]) else ''
            if r_name == region and classification == '1':
                current = safe_float(row[current_quarter_col], None)
                prev = safe_float(row[prev_year_col], None)
                if current is not None and prev is not None and prev != 0:
                    bus_growth = ((current - prev) / prev) * 100
                    businesses.append({
                        'name': BUSINESS_MAPPING.get(str(row[4]) if pd.notna(row[4]) else '', str(row[4]) if pd.notna(row[4]) else ''),
                        'growth_rate': round(bus_growth, 1)
                    })
        
        # 기여도 순 정렬
        if growth_rate >= 0:
            sorted_bus = sorted([b for b in businesses if b['growth_rate'] > 0], 
                              key=lambda x: x['growth_rate'], reverse=True)
        else:
            sorted_bus = sorted([b for b in businesses if b['growth_rate'] < 0], 
                              key=lambda x: x['growth_rate'])
        
        regions.append({
            'region': region,
            'growth_rate': round(growth_rate, 1),
            'growth_2023_2': 0.0,  # 기초자료에서는 계산 불가
            'growth_2024_2': 0.0,
            'growth_2025_1': 0.0,
            'index_2024': prev_val or 0,
            'index_2025': current_val or 0,
            'top_businesses': sorted_bus[:3],
            'all_businesses': businesses
        })
    
    # 증가/감소 지역 분류
    # 0.0인 지역은 완전히 제외 (None도 제외)
    increase_regions = sorted(
        [r for r in regions if r.get('growth_rate') is not None and r['growth_rate'] > 0],
        key=lambda x: x['growth_rate'],
        reverse=True
    )
    decrease_regions = sorted(
        [r for r in regions if r.get('growth_rate') is not None and r['growth_rate'] < 0],
        key=lambda x: x['growth_rate']
    )
    
    return {
        'increase_regions': increase_regions,
        'decrease_regions': decrease_regions,
        'all_regions': regions
    }


def get_growth_rates_table(df_analysis, df_index):
    """표에 들어갈 증감률 및 지수 데이터 생성"""
    use_aggregation_only = df_analysis.attrs.get('use_aggregation_only', False)
    region_indices = get_region_indices(df_analysis)
    
    # 분석 시트가 비어있거나 지역 인덱스를 찾을 수 없으면 집계 시트에서 직접 계산
    if use_aggregation_only or len(region_indices) <= 1:
        return _get_table_data_from_aggregation(df_index)
    
    table_data = []
    
    # 전국
    try:
        nationwide_row = df_analysis.iloc[3]
        nationwide_idx = df_index.iloc[3]
        table_data.append({
            'group': None,
            'rowspan': None,
            'region': REGION_DISPLAY_MAPPING['전국'],
            'growth_rates': [
                round(float(nationwide_row[12]), 1),  # 2023 2/4
                round(float(nationwide_row[16]), 1),  # 2024 2/4
                round(float(nationwide_row[19]), 1),  # 2025 1/4
                round(float(nationwide_row[20]), 1),  # 2025 2/4
            ],
            'indices': [
                nationwide_idx[21],  # 2024 2/4
                nationwide_idx[24],  # 2025 2/4
            ]
        })
    except:
        return _get_table_data_from_aggregation(df_index)
    
    # 지역별 그룹
    for group_name, group_regions in REGION_GROUPS.items():
        for i, region in enumerate(group_regions):
            if region not in region_indices:
                continue
                
            start_idx = region_indices[region]
            row = df_analysis.iloc[start_idx]
            idx_row = df_index[df_index[2] == region]
            
            if idx_row.empty:
                continue
                
            idx_row = idx_row.iloc[0]
            
            try:
                entry = {
                    'region': REGION_DISPLAY_MAPPING.get(region, region),
                    'growth_rates': [
                        round(float(row[12]), 1),  # 2023 2/4
                        round(float(row[16]), 1),  # 2024 2/4
                        round(float(row[19]), 1),  # 2025 1/4
                        round(float(row[20]), 1),  # 2025 2/4
                    ],
                    'indices': [
                        idx_row[21],  # 2024 2/4
                        idx_row[24],  # 2025 2/4
                    ]
                }
            except:
                continue
            
            if i == 0:
                entry['group'] = group_name
                entry['rowspan'] = len(group_regions)
            else:
                entry['group'] = None
                entry['rowspan'] = None
                
            table_data.append(entry)
    
    return table_data


def _get_table_data_from_aggregation(df_index):
    """집계 시트에서 테이블 데이터 추출"""
    # 집계 시트 구조: 2=지역이름, 3=분류단계, 5=업태코드, 6=업태종류
    # 데이터 컬럼: 12=2022.2/4, 16=2023.2/4, 19=2024.1/4, 20=2024.2/4, 23=2025.1/4, 24=2025.2/4
    
    table_data = []
    all_regions = ['전국', '서울', '인천', '경기', '대전', '세종', '충북', '충남', 
                   '광주', '전북', '전남', '제주', '대구', '경북', '강원', '부산', '울산', '경남']
    
    for region in all_regions:
        # 해당 지역 총지수 (분류단계 0)
        region_total = df_index[(df_index[2] == region) & (df_index[3].astype(str) == '0')]
        if region_total.empty:
            continue
        
        row = region_total.iloc[0]
        
        # 지수 추출
        idx_2024_2 = safe_float(row[20], 0)  # 2024.2/4 지수
        idx_2025_2 = safe_float(row[24], 0)  # 2025.2/4 지수
        idx_2023_2 = safe_float(row[16], 0)  # 2023.2/4 지수
        idx_2022_2 = safe_float(row[12], 0)  # 2022.2/4 지수
        idx_2025_1 = safe_float(row[23], 0)  # 2025.1/4 지수
        idx_2024_1 = safe_float(row[19], 0)  # 2024.1/4 지수
        
        # 전년동분기비 증감률 계산
        growth_2023_2 = ((idx_2023_2 - idx_2022_2) / idx_2022_2 * 100) if idx_2022_2 and idx_2022_2 != 0 else 0.0
        growth_2024_2 = ((idx_2024_2 - idx_2023_2) / idx_2023_2 * 100) if idx_2023_2 and idx_2023_2 != 0 else 0.0
        growth_2025_1 = ((idx_2025_1 - idx_2024_1) / idx_2024_1 * 100) if idx_2024_1 and idx_2024_1 != 0 else 0.0
        growth_2025_2 = ((idx_2025_2 - idx_2024_2) / idx_2024_2 * 100) if idx_2024_2 and idx_2024_2 != 0 else 0.0
        
        table_data.append({
            'region': REGION_DISPLAY_MAPPING.get(region, region),
            'growth_rates': [
                round(growth_2023_2, 1),
                round(growth_2024_2, 1),
                round(growth_2025_1, 1),
                round(growth_2025_2, 1)
            ],
            'indices': [
                round(idx_2024_2, 1),
                round(idx_2025_2, 1)
            ]
        })
    
    # 그룹 정보 추가
    result_data = []
    
    # 전국 먼저
    nationwide = next((r for r in table_data if r['region'] == '전 국'), None)
    if nationwide:
        nationwide['group'] = None
        nationwide['rowspan'] = None
        result_data.append(nationwide)
    
    # 지역 그룹별로 추가
    for group_name, group_regions in REGION_GROUPS.items():
        for i, region in enumerate(group_regions):
            region_data = next((r for r in table_data if r['region'] == REGION_DISPLAY_MAPPING.get(region, region)), None)
            if region_data:
                if i == 0:
                    region_data['group'] = group_name
                    region_data['rowspan'] = len(group_regions)
                else:
                    region_data['group'] = None
                    region_data['rowspan'] = None
                result_data.append(region_data)
    
    return result_data


def get_summary_box_data(regional_data):
    """요약 박스 데이터 생성"""
    # 감소 지역 상위 3개
    top3_decrease = regional_data['decrease_regions'][:3]
    
    main_regions = []
    for r in top3_decrease:
        # 첫 번째 감소 업태
        main_business = r['top_businesses'][0]['name'] if r['top_businesses'] else ''
        main_regions.append({
            'region': r['region'],
            'main_business': main_business
        })
    
    return {
        'main_decrease_regions': main_regions,
        'region_count': len(regional_data['decrease_regions'])
    }


def generate_report(excel_path, template_path, output_path, raw_excel_path=None, year=None, quarter=None):
    """보도자료 생성
    
    Args:
        excel_path: 분석표 엑셀 파일 경로
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        raw_excel_path: 기초자료 엑셀 파일 경로 (선택사항, 향후 기초자료 직접 추출 지원 예정)
        year: 현재 연도 (선택사항)
        quarter: 현재 분기 (선택사항)
    """
    # 기초자료 직접 추출은 현재 사용하지 않음 (분석표만 사용)
    # if raw_excel_path and year and quarter:
    #     from raw_data_extractor import RawDataExtractor
    #     extractor = RawDataExtractor(raw_excel_path, year, quarter)
    #     # 기초자료에서 소비동향 데이터 직접 추출
    #     # return extract_from_raw_data(extractor, ...)
    # 데이터 로드
    df_analysis, df_index = load_data(excel_path)
    
    # 데이터 추출
    # [Robust Dynamic Parsing System]
    # BaseGenerator 인스턴스 생성 (동적 컬럼 탐색용)
    if year is None or quarter is None:
        from utils.excel_utils import extract_year_quarter_from_excel
        extracted_year, extracted_quarter = extract_year_quarter_from_excel(excel_path)
        year = year or extracted_year or 2025
        quarter = quarter or extracted_quarter or 2
    
    # 임시 BaseGenerator 인스턴스 생성
    class TempGenerator(BaseGenerator):
        pass
    
    generator = TempGenerator(excel_path, year, quarter)
    
    nationwide_data = get_nationwide_data(df_analysis, df_index, generator, year, quarter)
    regional_data = get_regional_data(df_analysis, df_index, generator, year, quarter)
    summary_box = get_summary_box_data(regional_data)
    table_data = get_growth_rates_table(df_analysis, df_index, generator, year, quarter)
    
    # Top 3 증가/감소 지역
    top3_increase = []
    for r in regional_data['increase_regions'][:3]:
        top3_increase.append({
            'region': r['region'],
            'growth_rate': r['growth_rate'],
            'businesses': r['top_businesses']
        })
    
    top3_decrease = []
    for r in regional_data['decrease_regions'][:3]:
        top3_decrease.append({
            'region': r['region'],
            'growth_rate': r['growth_rate'],
            'businesses': r['top_businesses']
        })
    
    # 증가/감소 업태 텍스트 생성
    increase_businesses = set()
    for r in regional_data['increase_regions'][:3]:
        for bus in r['top_businesses'][:2]:
            increase_businesses.add(bus['name'])
    increase_businesses_text = ', '.join(list(increase_businesses)[:3])
    
    decrease_businesses = set()
    for r in regional_data['decrease_regions'][:3]:
        for bus in r['top_businesses'][:2]:
            decrease_businesses.add(bus['name'])
    decrease_businesses_text = ', '.join(list(decrease_businesses)[:4])
    
    # 템플릿 데이터
    template_data = {
        'report_info': {
            'year': year if year else 2025,
            'quarter': quarter if quarter else 2,
            'data_source': '국가데이터처 국가통계포털(KOSIS), 소비동향조사'
        },
        'summary_box': summary_box,
        'nationwide_data': nationwide_data,
        'regional_data': regional_data,
        'top3_increase_regions': top3_increase,
        'top3_decrease_regions': top3_decrease,
        'increase_businesses_text': increase_businesses_text,
        'decrease_businesses_text': decrease_businesses_text,
        'summary_table': {
            'base_year': 2020,
            'columns': {
                'growth_rate_columns': ['2023.2/4', '2024.2/4', '2025.1/4', '2025.2/4p'],
                'index_columns': ['2024.2/4', '2025.2/4p']
            },
            'regions': table_data
        }
    }
    
    # JSON 데이터 저장
    data_path = Path(output_path).parent / 'consumption_data.json'
    with open(data_path, 'w', encoding='utf-8') as f:
        json.dump(template_data, f, ensure_ascii=False, indent=2, default=str)
    
    # 템플릿 렌더링
    with open(template_path, 'r', encoding='utf-8') as f:
        template = Template(f.read())
    
    html_output = template.render(**template_data)
    
    # HTML 저장
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_output)
    
    print(f"보도자료 생성 완료: {output_path}")
    print(f"데이터 파일 저장: {data_path}")
    
    return template_data


if __name__ == '__main__':
    base_path = Path(__file__).parent.parent
    excel_path = base_path / '분석표_25년 2분기_캡스톤.xlsx'
    template_path = Path(__file__).parent / 'consumption_template.html'
    output_path = Path(__file__).parent / 'consumption_output.html'
    
    data = generate_report(excel_path, template_path, output_path)
    
    # 검증용 출력
    print("\n=== 전국 데이터 ===")
    print(f"판매지수: {data['nationwide_data']['sales_index']}")
    print(f"증감률: {data['nationwide_data']['growth_rate']}%")
    print(f"주요 업태: {data['nationwide_data']['main_businesses']}")
    
    print("\n=== 증가 지역 Top 3 ===")
    for r in data['top3_increase_regions']:
        print(f"{r['region']}({r['growth_rate']}%): {r['businesses']}")
    
    print("\n=== 감소 지역 Top 3 ===")
    for r in data['top3_decrease_regions']:
        print(f"{r['region']}({r['growth_rate']}%): {r['businesses']}")

