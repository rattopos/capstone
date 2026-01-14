#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
서비스업생산 보도자료 생성기
B 분석 시트에서 데이터를 추출하여 HTML 보도자료를 생성합니다.
"""

import pandas as pd
import json
from jinja2 import Template
from pathlib import Path

# 업종명 매핑
INDUSTRY_MAPPING = {
    '수도, 하수 및 폐기물 처리, 원료 재생업': '수도·하수',
    '도매 및 소매업': '도소매',
    '운수 및 창고업': '운수·창고',
    '숙박 및 음식점업': '숙박·음식점',
    '정보통신업': '정보통신',
    '금융 및 보험업': '금융·보험',
    '부동산업': '부동산',
    '전문, 과학 및 기술 서비스업': '전문·과학·기술',
    '사업시설관리, 사업지원 및 임대 서비스업': '사업시설관리·사업지원·임대',
    '교육 서비스업': '교육',
    '보건업 및 사회복지 서비스업': '보건·복지',
    '예술, 스포츠 및 여가관련 서비스업': '예술·스포츠·여가',
    '협회 및 단체, 수리  및 기타 개인 서비스업': '협회·수리·개인서비스'
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
        ['B 분석', 'B분석'],
        ['서비스업생산', '서비스업생산지수']
    )
    
    if analysis_sheet:
        df_analysis = pd.read_excel(xl, sheet_name=analysis_sheet, header=None)
        # 분석 시트에 실제 데이터가 있는지 확인 (수식 미계산 체크)
        test_row = df_analysis[(df_analysis[3] == '전국') | (df_analysis[4] == '전국')]
        if test_row.empty or test_row.iloc[0].isna().sum() > 20:
            print(f"[서비스업생산] 분석 시트가 비어있음 → 집계 시트에서 직접 계산")
            use_aggregation_only = True
        else:
            use_aggregation_only = False
    else:
        raise ValueError(f"서비스업생산 분석 시트를 찾을 수 없습니다. 시트: {xl.sheet_names}")
    
    # 집계 시트 찾기
    agg_sheet, _ = find_sheet_with_fallback(
        xl,
        ['B(서비스업생산)집계', 'B 집계'],
        ['서비스업생산', '서비스업생산지수']
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
    
    # 시도 목록
    VALID_REGIONS = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                     '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
    
    for i in range(len(df_analysis)):
        try:
            row = df_analysis.iloc[i]
            col7_val = str(row[7]).strip() if pd.notna(row[7]) else ''
            col3_val = str(row[3]).strip() if pd.notna(row[3]) else ''
            
            if col7_val == '총지수' and col3_val in VALID_REGIONS:
                region_indices[col3_val] = i
        except (IndexError, KeyError):
            continue
    
    return region_indices


def get_nationwide_data(df_analysis, df_index):
    """전국 데이터 추출"""
    use_raw = df_analysis.attrs.get('use_raw', False)
    use_aggregation_only = df_analysis.attrs.get('use_aggregation_only', False)
    
    if use_raw:
        return _get_nationwide_from_raw_data(df_analysis)
    
    # 집계 시트 기반으로 추출 (분석 시트가 비어있는 경우 포함)
    if use_aggregation_only:
        return _get_nationwide_from_aggregation(df_index)
    
    # 시트 구조 동적 감지
    if len(df_analysis) < 4:
        # 데이터가 부족한 경우 집계 시트에서 계산
        return _get_nationwide_from_aggregation(df_index)
    
    # 전국 총지수 행 찾기 (컬럼 3이 '전국'이고 컬럼 7이 '총지수'인 행)
    nationwide_row = None
    nationwide_idx = None
    for i in range(len(df_analysis)):
        row = df_analysis.iloc[i]
        if pd.notna(row[3]) and str(row[3]).strip() == '전국':
            if pd.notna(row[7]) and str(row[7]).strip() == '총지수':
                nationwide_row = row
                nationwide_idx = i
                break
    
    if nationwide_row is None:
        # 집계 시트에서 계산
        return _get_nationwide_from_aggregation(df_index)
    
    # 분석 시트에서 증감률 읽기 (컬럼 20: 2025.2/4)
    growth_rate = safe_float(nationwide_row[20], 0)
    growth_rate = round(growth_rate, 1) if growth_rate else 0.0
    
    # 집계 시트에서 전국 지수
    try:
        if len(df_index) > 3:
            index_row = df_index.iloc[3]
            production_index = safe_float(index_row[25], 100)  # 2025.2/4p
        else:
            production_index = 100.0
    except (IndexError, KeyError):
        production_index = 100.0
    
    # 전국 주요 업종 (기여도 기준 상위 3개 - 양수만)
    industries = []
    start_idx = nationwide_idx + 1 if nationwide_idx is not None else 4
    end_idx = min(start_idx + 13, len(df_analysis))
    
    for i in range(start_idx, end_idx):
        try:
            row = df_analysis.iloc[i]
            industry_name = row[7] if pd.notna(row[7]) else ''
            industry_growth = safe_float(row[20], 0)
            if industry_growth is not None and str(industry_name).strip() != '총지수':
                industries.append({
                    'name': INDUSTRY_MAPPING.get(str(industry_name).strip(), str(industry_name).strip()),
                    'growth_rate': round(industry_growth, 1)
                })
        except (IndexError, KeyError):
            continue
    
    # 양수 증가율 중 상위 3개 (보건·복지, 금융·보험, 운수·창고 순)
    positive_industries = [i for i in industries if i['growth_rate'] > 0]
    positive_industries.sort(key=lambda x: x['growth_rate'], reverse=True)
    main_industries = positive_industries[:3]
    
    return {
        'production_index': production_index,
        'growth_rate': growth_rate,
        'main_industries': main_industries
    }


def _get_nationwide_from_aggregation(df_index):
    """집계 시트에서 전국 데이터 추출 (증감률 직접 계산)"""
    # 집계 시트 구조: 3=지역이름, 4=분류단계, 6=산업코드, 7=산업이름
    # 데이터 컬럼: 21=2024.2/4, 25=2025.2/4
    
    # 전국 총지수 행 찾기 (산업코드 'E~S' 또는 분류단계 0)
    nationwide_rows = df_index[(df_index[3] == '전국') & (df_index[6] == 'E~S')]
    if nationwide_rows.empty:
        # 대체: 분류단계 0인 전국 행
        nationwide_rows = df_index[(df_index[3] == '전국') & (df_index[4].astype(str) == '0')]
    if nationwide_rows.empty:
        return {
            'production_index': 100.0,
            'growth_rate': 0.0,
            'main_industries': []
        }
    
    nationwide_total = nationwide_rows.iloc[0]
    
    # 당분기(2025.2/4)와 전년동분기(2024.2/4) 지수로 증감률 계산
    current_index = safe_float(nationwide_total[25], 100)  # 2025.2/4
    prev_year_index = safe_float(nationwide_total[21], 100)  # 2024.2/4
    
    if prev_year_index and prev_year_index != 0:
        growth_rate = ((current_index - prev_year_index) / prev_year_index) * 100
    else:
        growth_rate = 0.0
    
    # 전국 중분류 업종별 데이터 (분류단계 1)
    nationwide_industries = df_index[(df_index[3] == '전국') & (df_index[4].astype(str) == '1')]
    
    industries = []
    for _, row in nationwide_industries.iterrows():
        curr = safe_float(row[25], None)
        prev = safe_float(row[21], None)
        
        if curr is not None and prev is not None and prev != 0:
            ind_growth = ((curr - prev) / prev) * 100
            industries.append({
                'name': INDUSTRY_MAPPING.get(str(row[7]) if pd.notna(row[7]) else '', str(row[7]) if pd.notna(row[7]) else ''),
                'growth_rate': round(ind_growth, 1)
            })
    
    # 양수 증가율 중 상위 3개
    positive_industries = sorted([i for i in industries if i['growth_rate'] > 0], 
                                key=lambda x: x['growth_rate'], reverse=True)[:3]
    
    return {
        'production_index': current_index,
        'growth_rate': round(growth_rate, 1),
        'main_industries': positive_industries
    }


def _get_nationwide_from_raw_data(df):
    """기초자료 시트에서 전국 데이터 추출"""
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
    
    # 전국 총지수 행 찾기 (분류단계 0)
    nationwide_row = None
    for i in range(header_row + 1, len(df)):
        row = df.iloc[i]
        region = str(row[1]).strip() if pd.notna(row[1]) else ''
        classification = str(row[2]).strip() if pd.notna(row[2]) else ''
        if region == '전국' and classification == '0':
            nationwide_row = row
            break
    
    if nationwide_row is None:
        return {'production_index': 100.0, 'growth_rate': 0.0, 'main_industries': []}
    
    # 증감률 계산
    current_val = safe_float(nationwide_row[current_quarter_col], 100)
    prev_val = safe_float(nationwide_row[prev_year_col], 100)
    if prev_val and prev_val != 0:
        growth_rate = ((current_val - prev_val) / prev_val) * 100
    else:
        growth_rate = 0.0
    
    # 업종별 데이터 추출 (분류단계 1 또는 2)
    industries = []
    for i in range(header_row + 1, len(df)):
        row = df.iloc[i]
        region = str(row[1]).strip() if pd.notna(row[1]) else ''
        classification = str(row[2]).strip() if pd.notna(row[2]) else ''
        if region == '전국' and classification in ['1', '2']:
            current = safe_float(row[current_quarter_col], None)
            prev = safe_float(row[prev_year_col], None)
            if current is not None and prev is not None and prev != 0:
                ind_growth = ((current - prev) / prev) * 100
                ind_name = str(row[4]) if pd.notna(row[4]) else ''
                industries.append({
                    'name': INDUSTRY_MAPPING.get(ind_name, ind_name),
                    'growth_rate': round(ind_growth, 1)
                })
    
    positive_industries = sorted([i for i in industries if i['growth_rate'] > 0], 
                                key=lambda x: x['growth_rate'], reverse=True)[:3]
    
    return {
        'production_index': current_val,
        'growth_rate': round(growth_rate, 1),
        'main_industries': positive_industries
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
        
        try:
            # 총지수 행에서 증감률 (컬럼 20: 2025.2/4)
            total_row = df_analysis.iloc[start_idx]
            growth_rate_val = safe_float(total_row[20], 0)
            growth_rate = round(growth_rate_val, 1) if growth_rate_val else 0.0
            
            # 집계 시트에서 지수
            idx_row = df_index[df_index[3] == region]
            if not idx_row.empty:
                index_2024 = safe_float(idx_row.iloc[0][21], 0)
                index_2025 = safe_float(idx_row.iloc[0][25], 0)
            else:
                index_2024 = 0
                index_2025 = 0
            
            # 업종별 기여도 (컬럼 26)
            industries = []
            for i in range(start_idx + 1, min(start_idx + 14, len(df_analysis))):
                row = df_analysis.iloc[i]
                classification = str(row[4]).strip() if pd.notna(row[4]) else ''
                if classification != '1':  # 분류단계가 1이 아니면 스킵
                    continue
                industry_name = row[7] if pd.notna(row[7]) else ''
                industry_growth = safe_float(row[20], 0)
                contribution = safe_float(row[26], 0)
                
                if contribution is not None:
                    industries.append({
                        'name': INDUSTRY_MAPPING.get(str(industry_name).strip(), str(industry_name).strip()),
                        'growth_rate': round(industry_growth, 1) if industry_growth else 0.0,
                        'contribution': contribution
                    })
            
            # 증가 지역: 양수 기여도 순으로 정렬
            # 감소 지역: 음수 기여도 순으로 정렬 (절대값 큰 순)
            if growth_rate >= 0:
                positive_industries = [i for i in industries if i['contribution'] > 0]
                positive_industries.sort(key=lambda x: x['contribution'], reverse=True)
                top_industries = positive_industries[:3]
            else:
                negative_industries = [i for i in industries if i['contribution'] < 0]
                negative_industries.sort(key=lambda x: x['contribution'])
                top_industries = negative_industries[:3]
            
            regions.append({
                'region': region,
                'growth_rate': growth_rate,
                'index_2024': index_2024,
                'index_2025': index_2025,
                'top_industries': top_industries,
                'all_industries': industries
            })
        except (IndexError, KeyError) as e:
            print(f"[WARNING] 지역 데이터 추출 실패 ({region}): {e}")
            continue
    
    # 증가/감소 지역 분류
    increase_regions = sorted(
        [r for r in regions if r['growth_rate'] > 0],
        key=lambda x: x['growth_rate'],
        reverse=True
    )
    decrease_regions = sorted(
        [r for r in regions if r['growth_rate'] < 0],
        key=lambda x: x['growth_rate']
    )
    
    return {
        'increase_regions': increase_regions,
        'decrease_regions': decrease_regions,
        'all_regions': regions
    }


def _get_regional_from_aggregation(df_index):
    """집계 시트에서 시도별 데이터 추출"""
    individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                          '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
    
    # 전국 전년동분기 지수 (기여도 계산용)
    # 집계 시트 구조: 3=지역이름, 4=분류단계, 6=산업코드('E~S'=총지수), 7=산업이름
    nationwide_rows = df_index[(df_index[3] == '전국') & (df_index[6] == 'E~S')]
    if nationwide_rows.empty:
        nationwide_rows = df_index[(df_index[3] == '전국') & (df_index[4].astype(str) == '0')]
    nationwide_prev = safe_float(nationwide_rows.iloc[0][21], 100) if not nationwide_rows.empty else 100
    
    regions = []
    
    for region in individual_regions:
        # 해당 지역 총지수 (산업코드 'E~S' 또는 분류단계 0)
        region_total = df_index[(df_index[3] == region) & (df_index[6] == 'E~S')]
        if region_total.empty:
            region_total = df_index[(df_index[3] == region) & (df_index[4].astype(str) == '0')]
        if region_total.empty:
            continue
        region_total = region_total.iloc[0]
        
        # 증감률 계산
        current = safe_float(region_total[25], 100)  # 2025.2/4
        prev = safe_float(region_total[21], 100)  # 2024.2/4
        
        if prev and prev != 0:
            growth_rate = ((current - prev) / prev) * 100
        else:
            growth_rate = 0.0
        
        # 해당 지역 업종별 데이터 (분류단계 1)
        region_industries = df_index[(df_index[3] == region) & (df_index[4].astype(str) == '1')]
        
        industries = []
        for _, row in region_industries.iterrows():
            curr = safe_float(row[25], None)
            prev_ind = safe_float(row[21], None)
            
            if curr is not None and prev_ind is not None and prev_ind != 0:
                ind_growth = ((curr - prev_ind) / prev_ind) * 100
                # 기여도 = (당기 - 전기) / 전국전기 * 100
                contribution = (curr - prev_ind) / nationwide_prev * 100 if nationwide_prev else 0
                industries.append({
                    'name': INDUSTRY_MAPPING.get(str(row[7]) if pd.notna(row[7]) else '', str(row[7]) if pd.notna(row[7]) else ''),
                    'growth_rate': round(ind_growth, 1),
                    'contribution': round(contribution, 6)
                })
        
        # 기여도 순 정렬
        if growth_rate >= 0:
            sorted_ind = sorted(industries, key=lambda x: x['contribution'], reverse=True)
        else:
            sorted_ind = sorted(industries, key=lambda x: x['contribution'])
        
        regions.append({
            'region': region,
            'growth_rate': round(growth_rate, 1),
            'index_2024': prev,
            'index_2025': current,
            'top_industries': sorted_ind[:3],
            'all_industries': industries
        })
    
    # 증가/감소 지역 분류
    increase_regions = sorted(
        [r for r in regions if r['growth_rate'] > 0],
        key=lambda x: x['growth_rate'],
        reverse=True
    )
    decrease_regions = sorted(
        [r for r in regions if r['growth_rate'] < 0],
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
        
        # 해당 지역의 업종별 데이터 (분류단계 1 또는 2)
        industries = []
        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            r_name = str(row[1]).strip() if pd.notna(row[1]) else ''
            classification = str(row[2]).strip() if pd.notna(row[2]) else ''
            if r_name == region and classification in ['1', '2']:
                current = safe_float(row[current_quarter_col], None)
                prev = safe_float(row[prev_year_col], None)
                if current is not None and prev is not None and prev != 0:
                    ind_growth = ((current - prev) / prev) * 100
                    industries.append({
                        'name': INDUSTRY_MAPPING.get(str(row[4]) if pd.notna(row[4]) else '', str(row[4]) if pd.notna(row[4]) else ''),
                        'growth_rate': round(ind_growth, 1),
                        'contribution': round(ind_growth * 0.1, 6)
                    })
        
        # 기여도 순 정렬
        if growth_rate >= 0:
            sorted_ind = sorted(industries, key=lambda x: x['contribution'], reverse=True)
        else:
            sorted_ind = sorted(industries, key=lambda x: x['contribution'])
        
        regions.append({
            'region': region,
            'growth_rate': round(growth_rate, 1),
            'index_2024': prev_val or 0,
            'index_2025': current_val or 0,
            'top_industries': sorted_ind[:3],
            'all_industries': industries
        })
    
    # 증가/감소 지역 분류
    increase_regions = sorted(
        [r for r in regions if r['growth_rate'] > 0],
        key=lambda x: x['growth_rate'],
        reverse=True
    )
    decrease_regions = sorted(
        [r for r in regions if r['growth_rate'] < 0],
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
    
    # 전국 데이터 찾기
    nationwide_idx = region_indices.get('전국', 3)
    try:
        nationwide_row = df_analysis.iloc[nationwide_idx]
        nationwide_idx_row = df_index.iloc[nationwide_idx] if len(df_index) > nationwide_idx else None
        
        table_data.append({
            'group': None,
            'rowspan': None,
            'region': REGION_DISPLAY_MAPPING['전국'],
            'growth_rates': [
                round(safe_float(nationwide_row[12], 0), 1),  # 2023 2/4
                round(safe_float(nationwide_row[16], 0), 1),  # 2024 2/4
                round(safe_float(nationwide_row[19], 0), 1),  # 2025 1/4
                round(safe_float(nationwide_row[20], 0), 1),  # 2025 2/4
            ],
            'indices': [
                safe_float(nationwide_idx_row[21], 0) if nationwide_idx_row is not None else 0,  # 2024 2/4
                safe_float(nationwide_idx_row[25], 0) if nationwide_idx_row is not None else 0,  # 2025 2/4
            ]
        })
    except (IndexError, KeyError):
        # 전국 데이터를 찾을 수 없는 경우 기본값 사용
        table_data.append({
            'group': None,
            'rowspan': None,
            'region': REGION_DISPLAY_MAPPING['전국'],
            'growth_rates': [0.0, 0.0, 0.0, 0.0],
            'indices': [0.0, 0.0]
        })
    
    # 지역별 그룹
    for group_name, group_regions in REGION_GROUPS.items():
        for i, region in enumerate(group_regions):
            if region not in region_indices:
                continue
            
            try:
                start_idx = region_indices[region]
                row = df_analysis.iloc[start_idx]
                idx_row = df_index[df_index[3] == region]
                
                if idx_row.empty:
                    continue
                    
                idx_row = idx_row.iloc[0]
                
                entry = {
                    'region': REGION_DISPLAY_MAPPING.get(region, region),
                    'growth_rates': [
                        round(safe_float(row[12], 0), 1),  # 2023 2/4
                        round(safe_float(row[16], 0), 1),  # 2024 2/4
                        round(safe_float(row[19], 0), 1),  # 2025 1/4
                        round(safe_float(row[20], 0), 1),  # 2025 2/4
                    ],
                    'indices': [
                        safe_float(idx_row[21], 0),  # 2024 2/4
                        safe_float(idx_row[25], 0),  # 2025 2/4
                    ]
                }
                
                if i == 0:
                    entry['group'] = group_name
                    entry['rowspan'] = len(group_regions)
                else:
                    entry['group'] = None
                    entry['rowspan'] = None
                    
                table_data.append(entry)
            except (IndexError, KeyError) as e:
                print(f"[WARNING] 지역 테이블 데이터 추출 실패 ({region}): {e}")
                continue
    
    return table_data


def _get_table_data_from_aggregation(df_index):
    """집계 시트에서 테이블 데이터 추출"""
    # 집계 시트 구조: 3=지역이름, 4=분류단계, 6=산업코드('E~S'=총지수), 7=산업이름
    # 데이터 컬럼: 13=2023.2/4, 17=2024.2/4, 20=2025.1/4, 21=2024.2/4 지수, 24=2025.1/4 지수, 25=2025.2/4 지수
    
    table_data = []
    all_regions = ['전국', '서울', '인천', '경기', '대전', '세종', '충북', '충남', 
                   '광주', '전북', '전남', '제주', '대구', '경북', '강원', '부산', '울산', '경남']
    
    for region in all_regions:
        # 해당 지역 총지수 (산업코드 'E~S' 또는 분류단계 0)
        region_total = df_index[(df_index[3] == region) & (df_index[6] == 'E~S')]
        if region_total.empty:
            region_total = df_index[(df_index[3] == region) & (df_index[4].astype(str) == '0')]
        if region_total.empty:
            continue
        
        row = region_total.iloc[0]
        
        # 지수 추출
        idx_2024_2 = safe_float(row[21], 0)  # 2024.2/4 지수
        idx_2025_2 = safe_float(row[25], 0)  # 2025.2/4 지수
        idx_2023_2 = safe_float(row[17], 0)  # 2023.2/4 지수
        idx_2025_1 = safe_float(row[24], 0)  # 2025.1/4 지수
        
        # 전년동분기비 증감률 계산
        # 2023.2/4 증감률 = (2023.2/4 지수 - 2022.2/4 지수) / 2022.2/4 지수 * 100
        idx_2022_2 = safe_float(row[13], 0)  # 2022.2/4 지수
        growth_2023_2 = ((idx_2023_2 - idx_2022_2) / idx_2022_2 * 100) if idx_2022_2 and idx_2022_2 != 0 else 0.0
        
        # 2024.2/4 증감률
        growth_2024_2 = ((idx_2024_2 - idx_2023_2) / idx_2023_2 * 100) if idx_2023_2 and idx_2023_2 != 0 else 0.0
        
        # 2025.1/4 증감률 (전년동분기대비)
        idx_2024_1 = safe_float(row[20], 0)  # 2024.1/4 지수
        growth_2025_1 = ((idx_2025_1 - idx_2024_1) / idx_2024_1 * 100) if idx_2024_1 and idx_2024_1 != 0 else 0.0
        
        # 2025.2/4 증감률
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
    top3 = regional_data['increase_regions'][:3]
    
    main_regions = []
    for r in top3:
        industries = [ind['name'] for ind in r['top_industries'][:2]]
        main_regions.append({
            'region': r['region'],
            'industries': industries
        })
    
    return {
        'main_increase_regions': main_regions,
        'region_count': len(regional_data['increase_regions'])
    }


def generate_report_data(excel_path, raw_excel_path=None, year=None, quarter=None):
    """미리보기용 보도자료 데이터 생성
    
    Args:
        excel_path: 분석표 엑셀 파일 경로
        raw_excel_path: 기초자료 엑셀 파일 경로 (선택사항)
        year: 현재 연도 (선택사항)
        quarter: 현재 분기 (선택사항)
    
    Returns:
        dict: 템플릿에 전달할 데이터
    """
    # 데이터 로드
    df_analysis, df_index = load_data(excel_path)
    
    # 데이터 추출
    nationwide_data = get_nationwide_data(df_analysis, df_index)
    regional_data = get_regional_data(df_analysis, df_index)
    summary_box = get_summary_box_data(regional_data)
    table_data = get_growth_rates_table(df_analysis, df_index)
    
    # Top 3 증가/감소 지역
    top3_increase = []
    for r in regional_data['increase_regions'][:3]:
        top3_increase.append({
            'region': r['region'],
            'growth_rate': r['growth_rate'],
            'industries': r['top_industries']
        })
    
    top3_decrease = []
    for r in regional_data['decrease_regions'][:3]:
        top3_decrease.append({
            'region': r['region'],
            'growth_rate': r['growth_rate'],
            'industries': r['top_industries']
        })
    
    # 감소/증가 업종 텍스트 생성
    decrease_industries = set()
    for r in regional_data['decrease_regions'][:3]:
        for ind in r['top_industries'][:2]:
            decrease_industries.add(ind['name'])
    decrease_industries_text = ', '.join(list(decrease_industries)[:4])
    
    increase_industries = set()
    for r in regional_data['increase_regions'][:3]:
        for ind in r['top_industries'][:2]:
            increase_industries.add(ind['name'])
    increase_industries_text = ', '.join(list(increase_industries)[:4])
    
    # 템플릿 데이터
    template_data = {
        'report_info': {
            'year': year,
            'quarter': quarter
        },
        'summary_box': summary_box,
        'nationwide_data': nationwide_data,
        'regional_data': regional_data,
        'top3_increase_regions': top3_increase,
        'top3_decrease_regions': top3_decrease,
        'decrease_industries_text': decrease_industries_text,
        'increase_industries_text': increase_industries_text,
        'summary_table': {
            'base_year': 2020,
            'columns': {
                'growth_rate_columns': ['2023.2/4', '2024.2/4', '2025.1/4', '2025.2/4p'],
                'index_columns': ['2024.2/4', '2025.2/4p']
            },
            'regions': table_data
        }
    }
    
    return template_data


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
    # 데이터 로드
    df_analysis, df_index = load_data(excel_path)
    
    # 데이터 추출
    nationwide_data = get_nationwide_data(df_analysis, df_index)
    regional_data = get_regional_data(df_analysis, df_index)
    summary_box = get_summary_box_data(regional_data)
    table_data = get_growth_rates_table(df_analysis, df_index)
    
    # Top 3 증가/감소 지역
    top3_increase = []
    for r in regional_data['increase_regions'][:3]:
        top3_increase.append({
            'region': r['region'],
            'growth_rate': r['growth_rate'],
            'industries': r['top_industries']
        })
    
    top3_decrease = []
    for r in regional_data['decrease_regions'][:3]:
        top3_decrease.append({
            'region': r['region'],
            'growth_rate': r['growth_rate'],
            'industries': r['top_industries']
        })
    
    # 감소/증가 업종 텍스트 생성
    decrease_industries = set()
    for r in regional_data['decrease_regions'][:3]:
        for ind in r['top_industries'][:2]:
            decrease_industries.add(ind['name'])
    decrease_industries_text = ', '.join(list(decrease_industries)[:4])
    
    increase_industries = set()
    for r in regional_data['increase_regions'][:3]:
        for ind in r['top_industries'][:2]:
            increase_industries.add(ind['name'])
    increase_industries_text = ', '.join(list(increase_industries)[:4])
    
    # 템플릿 데이터
    template_data = {
        'summary_box': summary_box,
        'nationwide_data': nationwide_data,
        'regional_data': regional_data,
        'top3_increase_regions': top3_increase,
        'top3_decrease_regions': top3_decrease,
        'decrease_industries_text': decrease_industries_text,
        'increase_industries_text': increase_industries_text,
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
    data_path = Path(output_path).parent / 'service_industry_data.json'
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
    template_path = Path(__file__).parent / 'service_industry_template.html'
    output_path = Path(__file__).parent / 'service_industry_output.html'
    
    data = generate_report(excel_path, template_path, output_path)
    
    # 검증용 출력
    print("\n=== 전국 데이터 ===")
    print(f"생산지수: {data['nationwide_data']['production_index']}")
    print(f"증감률: {data['nationwide_data']['growth_rate']}%")
    print(f"주요 업종: {data['nationwide_data']['main_industries']}")
    
    print("\n=== 증가 지역 Top 3 ===")
    for r in data['top3_increase_regions']:
        print(f"{r['region']}({r['growth_rate']}%): {r['industries']}")
    
    print("\n=== 감소 지역 Top 3 ===")
    for r in data['top3_decrease_regions']:
        print(f"{r['region']}({r['growth_rate']}%): {r['industries']}")

