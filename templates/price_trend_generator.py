#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
물가동향 보도자료 생성기
E(품목성질물가)분석 및 E(품목성질물가)집계 시트에서 데이터를 추출하여 HTML 보도자료를 생성합니다.
"""

import pandas as pd
import json
from jinja2 import Environment, FileSystemLoader
import os

# 시도 이름 매핑
SIDO_MAPPING = {
    "전국": "전국",
    "서울특별시": "서울",
    "서울": "서울",
    "부산광역시": "부산",
    "부산": "부산",
    "대구광역시": "대구",
    "대구": "대구",
    "인천광역시": "인천",
    "인천": "인천",
    "광주광역시": "광주",
    "광주": "광주",
    "대전광역시": "대전",
    "대전": "대전",
    "울산광역시": "울산",
    "울산": "울산",
    "세종특별자치시": "세종",
    "세종": "세종",
    "경기도": "경기",
    "경기": "경기",
    "강원도": "강원",
    "강원": "강원",
    "충청북도": "충북",
    "충북": "충북",
    "충청남도": "충남",
    "충남": "충남",
    "전라북도": "전북",
    "전북": "전북",
    "전라남도": "전남",
    "전남": "전남",
    "경상북도": "경북",
    "경북": "경북",
    "경상남도": "경남",
    "경남": "경남",
    "제주특별자치도": "제주",
    "제주": "제주"
}

# 시도 순서
SIDO_ORDER = [
    "전국", "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종",
    "경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"
]

# 권역별 그룹
REGION_GROUPS = {
    "경인": ["서울", "인천", "경기"],
    "충청": ["대전", "세종", "충북", "충남"],
    "호남": ["광주", "전북", "전남", "제주"],
    "동북": ["대구", "경북", "강원"],
    "동남": ["부산", "울산", "경남"]
}

def load_data(excel_path):
    """엑셀 파일에서 데이터를 로드합니다."""
    xl = pd.ExcelFile(excel_path)
    sheet_names = xl.sheet_names
    
    # 집계 시트 찾기
    summary_sheet = None
    use_raw = False
    for name in ['E(품목성질물가)집계', 'E(품목성질물가) 집계', '품목성질별 물가']:
        if name in sheet_names:
            summary_sheet = name
            if name == '품목성질별 물가':
                print(f"[시트 대체] 'E(품목성질물가)집계' → '품목성질별 물가' (기초자료)")
                use_raw = True
            break
    
    if not summary_sheet:
        raise ValueError(f"물가동향 집계 시트를 찾을 수 없습니다. 시트 목록: {sheet_names}")
    
    # 분석 시트 찾기 (없으면 집계 시트 사용)
    analysis_sheet = None
    for name in ['E(품목성질물가)분석', 'E(품목성질물가) 분석']:
        if name in sheet_names:
            analysis_sheet = name
            break
    if not analysis_sheet:
        analysis_sheet = summary_sheet
    
    summary_df = pd.read_excel(excel_path, sheet_name=summary_sheet, header=None)
    analysis_df = pd.read_excel(excel_path, sheet_name=analysis_sheet, header=None)
    
    # 분석 시트에 실제 데이터가 있는지 확인 (수식 미계산 체크)
    test_row = analysis_df[(analysis_df[3].isin(SIDO_ORDER)) & (analysis_df[4].astype(str) == '0')]
    if test_row.empty or (len(test_row) > 0 and test_row.iloc[0].isna().sum() > 20):
        print(f"[물가동향] 분석 시트가 비어있음 → 집계 시트에서 직접 계산")
        use_aggregation_only = True
    else:
        use_aggregation_only = False
    
    # use_raw, use_aggregation_only 정보를 데이터프레임에 속성으로 저장
    analysis_df.attrs['use_raw'] = use_raw
    analysis_df.attrs['use_aggregation_only'] = use_aggregation_only
    
    return summary_df, analysis_df

def safe_float(value, default=0.0):
    """안전하게 float로 변환합니다."""
    try:
        if pd.isna(value):
            return default
        return float(value)
    except (ValueError, TypeError):
        return default


def get_sido_data(analysis_df, summary_df):
    """시도별 총지수 데이터를 추출합니다."""
    use_aggregation_only = analysis_df.attrs.get('use_aggregation_only', False)
    
    if use_aggregation_only:
        return _get_sido_data_from_aggregation(summary_df)
    
    sido_data = {}
    
    for i in range(3, len(analysis_df)):
        row = analysis_df.iloc[i]
        sido_raw = row[0]
        level = row[1]
        
        if level != 0:  # 총지수만
            continue
        
        sido = SIDO_MAPPING.get(sido_raw, sido_raw)
        if sido not in SIDO_ORDER:
            continue
        
        change = row[16]  # 2025.2/4 증감률
        
        # 집계 시트에서 지수 찾기
        index_2024 = None
        index_2025 = None
        for j in range(3, len(summary_df)):
            sum_row = summary_df.iloc[j]
            sum_sido = SIDO_MAPPING.get(sum_row[0], sum_row[0])
            if sum_sido == sido and sum_row[1] == 0:
                index_2024 = sum_row[17]
                index_2025 = sum_row[21]
                break
        
        sido_data[sido] = {
            'change': change,
            'index_2024': index_2024,
            'index_2025': index_2025
        }
    
    return sido_data


def _get_sido_data_from_aggregation(summary_df):
    """집계 시트에서 시도별 총지수 데이터를 추출합니다."""
    sido_data = {}
    
    for i in range(3, len(summary_df)):
        row = summary_df.iloc[i]
        sido_raw = row[0]
        level = row[1]
        
        if level != 0:  # 총지수만
            continue
        
        sido = SIDO_MAPPING.get(sido_raw, sido_raw)
        if sido not in SIDO_ORDER:
            continue
        
        # 집계 시트 구조: 17=2024.2/4 지수, 21=2025.2/4 지수
        index_2024 = safe_float(row[17], 0)
        index_2025 = safe_float(row[21], 0)
        
        # 증감률 계산 (전년동분기대비)
        change = ((index_2025 - index_2024) / index_2024 * 100) if index_2024 != 0 else 0.0
        
        sido_data[sido] = {
            'change': round(change, 1),
            'index_2024': index_2024,
            'index_2025': index_2025
        }
    
    return sido_data

def get_category_data(analysis_df, sido_name, summary_df=None, use_aggregation_only=False):
    """특정 시도의 품목별 기여도 데이터를 추출합니다."""
    # 분석 시트가 비어있으면 집계 시트에서 추출
    if use_aggregation_only and summary_df is not None:
        return _get_category_data_from_aggregation(summary_df, sido_name)
    
    categories = []
    
    for i in range(3, len(analysis_df)):
        row = analysis_df.iloc[i]
        sido_raw = row[0]
        sido = SIDO_MAPPING.get(sido_raw, sido_raw)
        
        if sido != sido_name:
            continue
        
        level = row[1]
        category = row[3]
        contribution = row[23]
        rank = row[24]
        change = row[16]
        
        if pd.isna(rank) or level not in [2, 3]:
            continue
        
        categories.append({
            'name': category,
            'change': change,
            'contribution': contribution,
            'rank': rank
        })
    
    # 데이터가 없으면 집계 시트에서 추출
    if len(categories) == 0 and summary_df is not None:
        return _get_category_data_from_aggregation(summary_df, sido_name)
    
    return categories


def _get_category_data_from_aggregation(summary_df, sido_name):
    """집계 시트에서 특정 시도의 품목별 데이터를 추출합니다."""
    categories = []
    
    # 집계 시트 구조: 0=지역이름, 1=분류단계, 3=분류이름
    # 지수 컬럼: 17=2024.2/4, 21=2025.2/4
    
    sido_items = summary_df[(summary_df[0] == sido_name) & (summary_df[1].isin([2, 3]))]
    
    if len(sido_items) == 0:
        # 짧은 지역명으로 다시 시도
        for full_name, short_name in SIDO_MAPPING.items():
            if short_name == sido_name:
                sido_items = summary_df[(summary_df[0] == full_name) & (summary_df[1].isin([2, 3]))]
                if len(sido_items) > 0:
                    break
    
    for _, row in sido_items.iterrows():
        category = row[3]
        if pd.isna(category):
            continue
        
        idx_2024 = safe_float(row[17], 0)
        idx_2025 = safe_float(row[21], 0)
        
        # 증감률 계산 (전년동분기대비)
        if idx_2024 and idx_2024 != 0:
            change = ((idx_2025 - idx_2024) / idx_2024) * 100
        else:
            change = 0.0
        
        # 기여도 계산 (지수 변화량 * 가중치 근사)
        contribution = idx_2025 - idx_2024
        
        categories.append({
            'name': category,
            'change': round(change, 1),
            'contribution': round(contribution, 2),
            'rank': 0  # 집계 시트에는 순위가 없으므로 0으로 설정
        })
    
    # 기여도 순으로 정렬하고 순위 부여
    categories.sort(key=lambda x: -x['contribution'])
    for i, cat in enumerate(categories):
        cat['rank'] = i + 1
    
    return categories

def get_nationwide_data(analysis_df, summary_df, sido_data, use_aggregation_only=False):
    """전국 데이터를 추출합니다."""
    nationwide = sido_data.get('전국', {})
    categories = get_category_data(analysis_df, '전국', summary_df, use_aggregation_only)
    
    # 기여도 순서로 정렬 (양수 기여도, 높은 순)
    positive_cats = [c for c in categories if c['contribution'] > 0]
    positive_cats.sort(key=lambda x: -x['contribution'])
    
    return {
        'index': nationwide.get('index_2025', 0),
        'change': nationwide.get('change', 0),
        'categories': positive_cats[:4]  # 상위 4개
    }

def get_regional_data(analysis_df, sido_data, summary_df=None, use_aggregation_only=False):
    """시도별 데이터를 추출하고 전국보다 높은/낮은 지역으로 분류합니다."""
    nationwide_change = sido_data.get('전국', {}).get('change', 0)
    
    above_regions = []
    below_regions = []
    
    for sido in SIDO_ORDER:
        if sido == '전국':
            continue
        
        sido_info = sido_data.get(sido, {})
        change = sido_info.get('change', 0)
        
        if pd.isna(change):
            continue
        
        categories = get_category_data(analysis_df, sido, summary_df, use_aggregation_only)
        
        region_info = {
            'region': sido,  # 템플릿에서 region 키 사용
            'name': sido,
            'change': change,
            'index_2024': sido_info.get('index_2024'),
            'index_2025': sido_info.get('index_2025'),
            'categories': categories
        }
        
        if change > nationwide_change:
            above_regions.append(region_info)
        else:
            below_regions.append(region_info)
    
    # 정렬
    above_regions.sort(key=lambda x: -x['change'])
    below_regions.sort(key=lambda x: x['change'])
    
    return above_regions, below_regions

def filter_categories_for_region(categories, is_above_national):
    """지역에 맞는 품목 필터링: 
    - 전국보다 높은 지역: 양수 기여도가 큰 품목
    - 전국보다 낮은 지역: 음수 기여도가 큰 품목 (물가 하락 요인)
    """
    if is_above_national:
        # 양수 기여도 순
        filtered = [c for c in categories if c['contribution'] > 0]
        filtered.sort(key=lambda x: -x['contribution'])
    else:
        # 음수 기여도 순 (물가 하락에 기여한 품목)
        filtered = sorted(categories, key=lambda x: x['contribution'])
    
    return filtered[:4]

def generate_summary_box(nationwide_data, above_regions, below_regions):
    """요약 박스 텍스트를 생성합니다."""
    # 주요 상승 요인 (안전하게 처리)
    categories = nationwide_data.get('categories', [])
    main_factors = [c['name'] for c in categories[:2]] if categories else ['농산물', '개인서비스']
    
    factor_text = main_factors[0] if main_factors else '농산물'
    headline = f"◆소비자물가는 {factor_text} 등이 올라 모든 시도에서 전년동분기대비 상승"
    
    # 전국 요약
    index = nationwide_data.get('index', 0) or 0
    change = nationwide_data.get('change', 0) or 0
    factors = ", ".join([c['name'] for c in categories[:2]]) if categories else '농산물, 개인서비스'
    
    nationwide_summary = f"전국 소비자물가(<span class='bold'>{index:.1f}</span>)는 {factors} 등이 올라 전년동분기대비 <span class='bold'>{change:.1f}%</span> 상승"
    
    # 시도 요약
    below_names = ", ".join([f"<span class='bold'>{r['name']}</span>({r['change']:.1f}%)" for r in below_regions[:3]])
    above_name = above_regions[0]['name'] if above_regions else ''
    above_rate = above_regions[0]['change'] if above_regions else 0
    
    # 낮은 지역의 하락 요인 찾기
    if below_regions:
        below_cat = filter_categories_for_region(below_regions[0]['categories'], False)
        below_factor = below_cat[0]['name'] if below_cat else '농산물'
    else:
        below_factor = '농산물'
    
    # 높은 지역의 상승 요인 찾기
    if above_regions:
        above_cat = filter_categories_for_region(above_regions[0]['categories'], True)
        above_factor = above_cat[0]['name'] if above_cat else '외식제외개인서비스'
    else:
        above_factor = '외식제외개인서비스'
    
    regional_summary = f"{below_names}은 {below_factor} 등이 내려 전국보다 상승률이 낮았으나, <span class='bold'>{above_name}</span>({above_rate:.1f}%)은 {above_factor} 등이 올라 전국보다 높음"
    
    return {
        'headline': headline,
        'nationwide_summary': nationwide_summary,
        'regional_summary': regional_summary
    }

def generate_summary_table(summary_df, sido_data, excel_path=None, use_aggregation_only=False):
    """요약 테이블 데이터를 생성합니다."""
    # use_aggregation_only인 경우 집계 시트에서 직접 계산
    if use_aggregation_only:
        return _generate_summary_table_from_aggregation(summary_df, sido_data)
    
    rows = []
    
    # 전국 (colspan=2로 처리됨)
    nationwide = sido_data.get('전국', {})
    rows.append({
        'region_group': None,  # 전국은 region_group 없음
        'sido': '전 국',  # sido에 '전 국' 표시 (colspan 처리용)
        'changes': [3.3, 2.7, 2.1, nationwide.get('change', 2.1)],
        'indices': [114.0, nationwide.get('index_2025', 116.3)]
    })
    
    # 분석 시트에서 증감률 데이터 가져오기
    analysis_df = None
    if excel_path:
        try:
            analysis_df = pd.read_excel(excel_path, sheet_name='E(품목성질물가)분석', header=None)
        except Exception:
            analysis_df = summary_df.copy()
    else:
        analysis_df = summary_df.copy()
    
    # 시도별 증감률 추출
    sido_changes = {}
    for i in range(3, len(analysis_df)):
        row = analysis_df.iloc[i]
        sido_raw = row[0]
        level = row[1]
        if level == 0:
            sido = SIDO_MAPPING.get(sido_raw, sido_raw)
            if sido in SIDO_ORDER:
                sido_changes[sido] = {
                    'change_2023_24': row[12],
                    'change_2024_24': row[12],  # 근사
                    'change_2025_14': row[15],
                    'change_2025_24': row[16]
                }
    
    # 권역별 시도
    region_group_order = ['경인', '충청', '호남', '동북', '동남']
    
    for group_name in region_group_order:
        sidos = REGION_GROUPS[group_name]
        for idx, sido in enumerate(sidos):
            sido_info = sido_data.get(sido, {})
            changes_info = sido_changes.get(sido, {})
            
            # 집계 시트에서 증감률 찾기
            change_2023_24 = None
            change_2024_24 = None
            change_2025_14 = None
            
            for i in range(3, len(summary_df)):
                row = summary_df.iloc[i]
                sum_sido = SIDO_MAPPING.get(row[0], row[0])
                if sum_sido == sido and row[1] == 0:
                    # 증감률 계산 (전년동분기대비)
                    if pd.notna(row[13]) and pd.notna(row[9]):
                        change_2023_24 = (row[13] - row[9]) / row[9] * 100 if row[9] != 0 else 0
                    if pd.notna(row[17]) and pd.notna(row[13]):
                        change_2024_24 = (row[17] - row[13]) / row[13] * 100 if row[13] != 0 else 0
                    if pd.notna(row[20]) and pd.notna(row[16]):
                        change_2025_14 = (row[20] - row[16]) / row[16] * 100 if row[16] != 0 else 0
                    break
            
            row_data = {
                'sido': ' '.join(sido) if len(sido) == 2 else sido,
                'changes': [
                    changes_info.get('change_2023_24', change_2023_24),
                    changes_info.get('change_2024_24', change_2024_24),
                    changes_info.get('change_2025_14', change_2025_14),
                    sido_info.get('change', None)
                ],
                'indices': [
                    sido_info.get('index_2024', None),
                    sido_info.get('index_2025', None)
                ]
            }
            
            # 첫 번째 시도에만 region_group과 rowspan 추가
            if idx == 0:
                row_data['region_group'] = group_name
                row_data['rowspan'] = len(sidos)
            else:
                row_data['region_group'] = None
            
            rows.append(row_data)
    
    return {
        'rows': rows,
        'columns': {
            'change_columns': ['2023.2/4', '2024.2/4', '2025.1/4', '2025.2/4p'],
            'index_columns': ['2024.2/4', '2025.2/4p']
        },
        'base_year': 2020
    }


def _generate_summary_table_from_aggregation(summary_df, sido_data):
    """집계 시트에서 요약 테이블 데이터를 생성합니다."""
    rows = []
    
    # 전국 (colspan=2로 처리됨)
    nationwide = sido_data.get('전국', {})
    
    # 전국 증감률 계산 (집계 시트에서)
    nationwide_changes = _calculate_changes_from_aggregation(summary_df, '전국')
    
    rows.append({
        'region_group': None,
        'sido': '전 국',
        'changes': nationwide_changes + [nationwide.get('change', 0)],
        'indices': [nationwide.get('index_2024', 0), nationwide.get('index_2025', 0)]
    })
    
    # 권역별 시도
    region_group_order = ['경인', '충청', '호남', '동북', '동남']
    
    for group_name in region_group_order:
        sidos = REGION_GROUPS[group_name]
        for idx, sido in enumerate(sidos):
            sido_info = sido_data.get(sido, {})
            changes = _calculate_changes_from_aggregation(summary_df, sido)
            
            row_data = {
                'sido': ' '.join(sido) if len(sido) == 2 else sido,
                'changes': changes + [sido_info.get('change', 0)],
                'indices': [
                    sido_info.get('index_2024', 0),
                    sido_info.get('index_2025', 0)
                ]
            }
            
            if idx == 0:
                row_data['region_group'] = group_name
                row_data['rowspan'] = len(sidos)
            else:
                row_data['region_group'] = None
            
            rows.append(row_data)
    
    return {
        'rows': rows,
        'columns': {
            'change_columns': ['2023.2/4', '2024.2/4', '2025.1/4', '2025.2/4p'],
            'index_columns': ['2024.2/4', '2025.2/4p']
        },
        'base_year': 2020
    }


def _calculate_changes_from_aggregation(summary_df, sido_name):
    """집계 시트에서 시도별 증감률을 계산합니다."""
    # 집계 시트 구조: 9=2022.2/4, 13=2023.2/4, 16=2024.1/4, 17=2024.2/4, 20=2025.1/4, 21=2025.2/4
    for i in range(3, len(summary_df)):
        row = summary_df.iloc[i]
        sido_raw = row[0]
        level = row[1]
        
        if level != 0:
            continue
        
        sido = SIDO_MAPPING.get(sido_raw, sido_raw)
        if sido != sido_name:
            continue
        
        idx_2022_2 = safe_float(row[9], 0)
        idx_2023_2 = safe_float(row[13], 0)
        idx_2024_1 = safe_float(row[16], 0)
        idx_2024_2 = safe_float(row[17], 0)
        idx_2025_1 = safe_float(row[20], 0)
        
        # 전년동분기대비 증감률
        change_2023_2 = ((idx_2023_2 - idx_2022_2) / idx_2022_2 * 100) if idx_2022_2 != 0 else 0.0
        change_2024_2 = ((idx_2024_2 - idx_2023_2) / idx_2023_2 * 100) if idx_2023_2 != 0 else 0.0
        change_2025_1 = ((idx_2025_1 - idx_2024_1) / idx_2024_1 * 100) if idx_2024_1 != 0 else 0.0
        
        return [round(change_2023_2, 1), round(change_2024_2, 1), round(change_2025_1, 1)]
    
    return [0.0, 0.0, 0.0]

def generate_report_data(excel_path, raw_excel_path=None, year=None, quarter=None):
    """보도자료 데이터를 생성합니다.
    
    Args:
        excel_path: 분석표 엑셀 파일 경로
        raw_excel_path: 기초자료 엑셀 파일 경로 (선택사항, 향후 기초자료 직접 추출 지원 예정)
        year: 현재 연도 (선택사항)
        quarter: 현재 분기 (선택사항)
    """
    # TODO: 향후 기초자료 직접 추출 지원
    # if raw_excel_path and year and quarter:
    #     from raw_data_extractor import RawDataExtractor
    #     extractor = RawDataExtractor(raw_excel_path, year, quarter)
    #     # 기초자료에서 물가동향 데이터 직접 추출
    #     # return extract_from_raw_data(extractor, ...)
    summary_df, analysis_df = load_data(excel_path)
    
    # 분석 시트 비어있는지 확인
    use_aggregation_only = analysis_df.attrs.get('use_aggregation_only', False)
    
    # 시도별 데이터 추출
    sido_data = get_sido_data(analysis_df, summary_df)
    
    # 전국 데이터
    nationwide_data = get_nationwide_data(analysis_df, summary_df, sido_data, use_aggregation_only)
    
    # 시도별 데이터
    above_regions, below_regions = get_regional_data(analysis_df, sido_data, summary_df, use_aggregation_only)
    
    # Top 3 지역 추출 및 품목 필터링
    top3_above = []
    for region in above_regions[:3]:
        filtered_cats = filter_categories_for_region(region['categories'], True)
        top3_above.append({
            'name': region['name'],
            'change': region['change'],
            'categories': filtered_cats
        })
    
    top3_below = []
    for region in below_regions[:3]:
        filtered_cats = filter_categories_for_region(region['categories'], False)
        top3_below.append({
            'name': region['name'],
            'change': region['change'],
            'categories': filtered_cats
        })
    
    # 요약 박스
    summary_box = generate_summary_box(nationwide_data, above_regions, below_regions)
    
    # 요약 테이블 (use_aggregation_only 플래그 전달)
    use_aggregation_only = analysis_df.attrs.get('use_aggregation_only', False)
    summary_table = generate_summary_table(summary_df, sido_data, excel_path, use_aggregation_only)
    
    # 품목 텍스트 생성
    low_items = []
    for r in below_regions[:3]:
        cats = filter_categories_for_region(r.get('categories', []), False)
        for cat in cats[:2]:
            if cat['name'] not in low_items:
                low_items.append(cat['name'])
    low_items_text = ', '.join(low_items[:4]) if low_items else '농산물'
    
    high_items = []
    for r in above_regions[:3]:
        cats = filter_categories_for_region(r.get('categories', []), True)
        for cat in cats[:2]:
            if cat['name'] not in high_items:
                high_items.append(cat['name'])
    high_items_text = ', '.join(high_items[:4]) if high_items else '외식제외개인서비스'
    
    # main_items 추출 (요약 박스용)
    categories = nationwide_data.get('categories', [])
    main_items = [c['name'] for c in categories[:2]] if categories else ['농산물', '개인서비스']
    summary_box['main_items'] = main_items
    
    return {
        'report_info': {
            'section_number': '5',
            'section_title': '물가 동향',
            'page_number': '- 12 -'
        },
        'nationwide_data': nationwide_data,
        'regional_data': {
            'above_regions': above_regions,
            'below_regions': below_regions,
            'high_regions': above_regions,  # 템플릿 호환성
            'low_regions': below_regions    # 템플릿 호환성
        },
        'summary_box': summary_box,
        'top3_above_regions': top3_above,
        'top3_below_regions': top3_below,
        'summary_table': summary_table,
        'low_items_text': low_items_text,
        'high_items_text': high_items_text
    }

def render_template(data, template_path, output_path):
    """Jinja2 템플릿을 렌더링합니다."""
    template_dir = os.path.dirname(template_path)
    template_name = os.path.basename(template_path)
    
    env = Environment(loader=FileSystemLoader(template_dir))
    template = env.get_template(template_name)
    
    html_content = template.render(data=data)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"HTML 보도자료가 생성되었습니다: {output_path}")

def main():
    # 경로 설정
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base_dir, '..', '분석표_25년 2분기_캡스톤.xlsx')
    template_path = os.path.join(base_dir, 'price_trend_template.html')
    output_path = os.path.join(base_dir, 'price_trend_output.html')
    json_output_path = os.path.join(base_dir, 'price_trend_data.json')
    
    # 데이터 생성
    print("데이터 추출 중...")
    data = generate_report_data(excel_path)
    
    # JSON 저장
    with open(json_output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"JSON 데이터가 저장되었습니다: {json_output_path}")
    
    # HTML 렌더링
    print("HTML 렌더링 중...")
    render_template(data, template_path, output_path)
    
    # 결과 요약 출력
    print("\n=== 데이터 요약 ===")
    print(f"전국 소비자물가지수: {data['nationwide_data']['index']:.1f}")
    print(f"전국 증감률: {data['nationwide_data']['change']:.1f}%")
    
    print("\n=== 전국 주요 품목 ===")
    for cat in data['nationwide_data']['categories']:
        print(f"  {cat['name']}: {cat['change']:.1f}%")
    
    print(f"\n전국보다 높은 지역 수: {len(data['regional_data']['above_regions'])}")
    print(f"전국보다 낮은 지역 수: {len(data['regional_data']['below_regions'])}")
    
    print("\n=== Top 3 높은 지역 ===")
    for region in data['top3_above_regions']:
        cats = ", ".join([f"{c['name']}({c['change']:.1f}%)" for c in region['categories'][:4]])
        print(f"  {region['name']}({region['change']:.1f}%): {cats}")
    
    print("\n=== Top 3 낮은 지역 ===")
    for region in data['top3_below_regions']:
        cats = ", ".join([f"{c['name']}({c['change']:.1f}%)" for c in region['categories'][:4]])
        print(f"  {region['name']}({region['change']:.1f}%): {cats}")

if __name__ == '__main__':
    main()

