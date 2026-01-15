#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
고용률 보도자료 생성기
D(고용률)분석 시트에서 데이터를 추출하여 HTML 보도자료를 생성합니다.
"""

import pandas as pd
import json
from jinja2 import Template
from pathlib import Path

# 연령대명 매핑 (이미지 표기에 맞춤)
AGE_GROUP_MAPPING = {
    '계': '계',
    '15 - 29세': '20~29세',
    '30 - 39세': '30~39세',
    '40 - 49세': '40~49세',
    '50 - 59세': '50~59세',
    '60세이상': '60세이상',  # 수정: 70세이상이 아닌 60세이상으로 표시
    '70세이상': '70세이상'  # 엑셀에 70세이상 컬럼이 있는 경우를 대비
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


def load_data(excel_path):
    """엑셀 파일에서 데이터 로드"""
    xl = pd.ExcelFile(excel_path)
    sheet_names = xl.sheet_names
    
    # 분석 시트 찾기
    analysis_sheet = None
    use_raw = False
    for name in ['D(고용률)분석', 'D(고용률) 분석', '고용률']:
        if name in sheet_names:
            analysis_sheet = name
            if name == '고용률':
                print(f"[시트 대체] 'D(고용률)분석' → '고용률' (기초자료)")
                use_raw = True
            break
    
    if not analysis_sheet:
        raise ValueError(f"고용률 분석 시트를 찾을 수 없습니다. 시트 목록: {sheet_names}")
    
    # 집계 시트 찾기 (없으면 분석 시트 사용)
    index_sheet = None
    for name in ['D(고용률)집계', 'D(고용률) 집계']:
        if name in sheet_names:
            index_sheet = name
            break
    if not index_sheet:
        index_sheet = analysis_sheet
    
    df_analysis = pd.read_excel(xl, sheet_name=analysis_sheet, header=None)
    df_index = pd.read_excel(xl, sheet_name=index_sheet, header=None)
    
    # 분석 시트에 실제 데이터가 있는지 확인 (수식 미계산 체크)
    test_row = df_analysis[(df_analysis[0].isin(VALID_REGIONS + ['전국'])) | (df_analysis[1].isin(VALID_REGIONS + ['전국']))]
    if test_row.empty or (len(test_row) > 0 and test_row.iloc[0].isna().sum() > 20):
        print(f"[고용률] 분석 시트가 비어있음 → 집계 시트에서 직접 계산")
        use_aggregation_only = True
    else:
        use_aggregation_only = False
    
    # use_raw, use_aggregation_only 정보를 데이터프레임에 속성으로 저장
    df_analysis.attrs['use_raw'] = use_raw
    df_analysis.attrs['use_aggregation_only'] = use_aggregation_only
    
    return df_analysis, df_index


def get_region_indices(df_analysis):
    """각 지역의 시작 인덱스 찾기"""
    region_indices = {}
    for i in range(len(df_analysis)):
        row = df_analysis.iloc[i]
        if row[5] == '계':
            region = row[2]
            if region in VALID_REGIONS or region == '전국':
                region_indices[region] = i
    return region_indices


def get_nationwide_data(df_analysis, df_index):
    """전국 데이터 추출"""
    use_aggregation_only = df_analysis.attrs.get('use_aggregation_only', False)
    
    if use_aggregation_only:
        return _get_nationwide_data_from_aggregation(df_index)
    
    # 분석 시트에서 전국 계 행
    try:
        if len(df_analysis) <= 3:
            print(f"[고용률] 분석 시트 행 수 부족: {len(df_analysis)}")
            return _get_nationwide_data_from_aggregation(df_index)
    nationwide_row = df_analysis.iloc[3]
        change = safe_float(nationwide_row[18] if len(nationwide_row) > 18 else None, 0)
    change = round(change, 1) if change is not None else 0.0
    except Exception as e:
        print(f"[고용률] 전국 데이터 읽기 실패: {e}")
        return _get_nationwide_data_from_aggregation(df_index)
    
    # 집계 시트에서 전국 고용률
    try:
        if len(df_index) <= 3:
            print(f"[고용률] 집계 시트 행 수 부족: {len(df_index)}")
            employment_rate = 60.0
        else:
    index_row = df_index.iloc[3]
            employment_rate = safe_float(index_row[21] if len(index_row) > 21 else None, 60.0)  # 2025.2/4
    except Exception as e:
        print(f"[고용률] 집계 시트 데이터 읽기 실패: {e}")
        employment_rate = 60.0
    
    # 전국 연령별 증감
    age_groups = []
    try:
        for i in range(4, min(9, len(df_analysis))):
            try:
        row = df_analysis.iloc[i]
                age_name = row[5] if len(row) > 5 and pd.notna(row[5]) else ''
                age_change = safe_float(row[18] if len(row) > 18 else None, None)
        if age_change is not None:
            age_groups.append({
                'name': age_name,
                'display_name': AGE_GROUP_MAPPING.get(age_name, age_name),
                'change': round(age_change, 1)
            })
            except Exception as e:
                print(f"[고용률] 연령별 데이터 추출 실패 (행 {i}): {e}")
                continue
    except Exception as e:
        print(f"[고용률] 연령별 데이터 목록 추출 실패: {e}")
    
    # 양수 증감률 순으로 정렬
    positive_ages = [a for a in age_groups if a['change'] > 0]
    positive_ages.sort(key=lambda x: x['change'], reverse=True)
    
    # 상위 4개 (이미지에서는 4개 표시)
    main_age_groups = positive_ages[:4]
    
    return {
        'employment_rate': employment_rate,
        'change': change,
        'main_age_groups': main_age_groups,
        'top_age_groups': positive_ages[:4]
    }


def _get_nationwide_data_from_aggregation(df_index):
    """집계 시트에서 전국 데이터 추출"""
    # 전국 계 행 찾기
    nationwide_total = df_index[(df_index[1] == '전국') & (df_index[3] == '계')]
    
    if nationwide_total.empty:
        return {
            'employment_rate': 60.0,
            'change': 0.0,
            'main_age_groups': [],
            'top_age_groups': []
        }
    
    nrow = nationwide_total.iloc[0]
    
    # 고용률 및 증감 계산
    rate_2024_2 = safe_float(nrow[17], 0)
    rate_2025_2 = safe_float(nrow[21], 0)
    change = round(rate_2025_2 - rate_2024_2, 1)
    
    # 연령별 데이터 추출
    # 60세이상과 70세이상 모두 체크 (엑셀 구조에 따라 다를 수 있음)
    age_names = ['15 - 29세', '30 - 39세', '40 - 49세', '50 - 59세', '60세이상', '70세이상']
    age_groups = []
    
    for age_name in age_names:
        age_row = df_index[(df_index[1] == '전국') & (df_index[3] == age_name)]
        if age_row.empty:
            # 60세이상이 없으면 70세이상으로 시도
            if age_name == '60세이상':
                age_row = df_index[(df_index[1] == '전국') & (df_index[3] == '70세이상')]
                if not age_row.empty:
                    age_name = '70세이상'  # 실제 엑셀의 연령대명 사용
            else:
                continue
        
        arow = age_row.iloc[0]
        age_rate_2024 = safe_float(arow[17], 0)
        age_rate_2025 = safe_float(arow[21], 0)
        age_change = round(age_rate_2025 - age_rate_2024, 1)
        
        age_groups.append({
            'name': age_name,
            'display_name': AGE_GROUP_MAPPING.get(age_name, age_name),
            'change': age_change
        })
    
    # 양수 증감률 순으로 정렬
    positive_ages = [a for a in age_groups if a['change'] > 0]
    positive_ages.sort(key=lambda x: x['change'], reverse=True)
    
    return {
        'employment_rate': round(rate_2025_2, 1),
        'change': change,
        'main_age_groups': positive_ages[:4],
        'top_age_groups': positive_ages[:4]
    }


def get_regional_data(df_analysis, df_index):
    """시도별 데이터 추출"""
    use_aggregation_only = df_analysis.attrs.get('use_aggregation_only', False)
    
    if use_aggregation_only:
        return _get_regional_data_from_aggregation(df_index)
    
    region_indices = get_region_indices(df_analysis)
    regions = []
    
    for region, start_idx in region_indices.items():
        if region == '전국':
            continue
            
        # 계 행에서 증감
        total_row = df_analysis.iloc[start_idx]
        change = safe_float(total_row[18], None)
        change_2023_2 = safe_float(total_row[10], 0)
        change_2024_2 = safe_float(total_row[14], 0)
        change_2025_1 = safe_float(total_row[17], 0)
        
        if change is None:
            continue
        
        change = round(change, 1)
        change_2023_2 = round(change_2023_2, 1) if change_2023_2 else 0.0
        change_2024_2 = round(change_2024_2, 1) if change_2024_2 else 0.0
        change_2025_1 = round(change_2025_1, 1) if change_2025_1 else 0.0
        
        # 집계 시트에서 고용률
        idx_row = df_index[(df_index[1] == region) & (df_index[3] == '계')]
        if not idx_row.empty:
            rate_2024 = idx_row.iloc[0][17]
            rate_2025 = idx_row.iloc[0][21]
            # 20-29세 고용률
            age_row = df_index[(df_index[1] == region) & (df_index[3] == '15 - 29세')]
            if not age_row.empty:
                rate_20_29 = age_row.iloc[0][21]
            else:
                rate_20_29 = 0
        else:
            rate_2024 = 0
            rate_2025 = 0
            rate_20_29 = 0
        
        # 연령별 증감
        age_groups = []
        for i in range(start_idx + 1, min(start_idx + 6, len(df_analysis))):
            row = df_analysis.iloc[i]
            if row[5] == '계':
                break
            age_name = row[5]
            try:
                age_change = round(float(row[18]), 1)
            except:
                continue
            
            age_groups.append({
                'name': age_name,
                'display_name': AGE_GROUP_MAPPING.get(age_name, age_name),
                'change': age_change
            })
        
        # 증가 지역: 양수 증감률 순으로 정렬
        # 감소 지역: 음수 증감률 순으로 정렬 (절대값 큰 순)
        if change >= 0:
            sorted_ages = sorted([a for a in age_groups if a['change'] > 0], 
                               key=lambda x: x['change'], reverse=True)
        else:
            sorted_ages = sorted([a for a in age_groups if a['change'] < 0], 
                               key=lambda x: x['change'])
        
        regions.append({
            'region': region,
            'change': change,
            'change_2023_2': change_2023_2,
            'change_2024_2': change_2024_2,
            'change_2025_1': change_2025_1,
            'rate_2024': rate_2024,
            'rate_2025': rate_2025,
            'rate_20_29': rate_20_29,
            'top_age_groups': sorted_ages[:3],
            'all_age_groups': age_groups
        })
    
    # 증가/감소 지역 분류
    increase_regions = sorted(
        [r for r in regions if r['change'] > 0],
        key=lambda x: x['change'],
        reverse=True
    )
    decrease_regions = sorted(
        [r for r in regions if r['change'] < 0],
        key=lambda x: x['change']
    )
    
    return {
        'increase_regions': increase_regions,
        'decrease_regions': decrease_regions,
        'all_regions': regions
    }


def _get_regional_data_from_aggregation(df_index):
    """집계 시트에서 시도별 데이터 추출"""
    regions = []
    # 60세이상과 70세이상 모두 체크
    age_names = ['15 - 29세', '30 - 39세', '40 - 49세', '50 - 59세', '60세이상', '70세이상']
    
    for region in VALID_REGIONS:
        region_total = df_index[(df_index[1] == region) & (df_index[3] == '계')]
        if region_total.empty:
            continue
        
        rrow = region_total.iloc[0]
        
        # 고용률 값
        rate_2022_2 = safe_float(rrow[9], 0)
        rate_2023_2 = safe_float(rrow[13], 0)
        rate_2024_1 = safe_float(rrow[16], 0)
        rate_2024_2 = safe_float(rrow[17], 0)
        rate_2025_1 = safe_float(rrow[20], 0)
        rate_2025_2 = safe_float(rrow[21], 0)
        
        # 전년동분기대비 %p 변화
        change = round(rate_2025_2 - rate_2024_2, 1)
        change_2023_2 = round(rate_2023_2 - rate_2022_2, 1)
        change_2024_2 = round(rate_2024_2 - rate_2023_2, 1)
        change_2025_1 = round(rate_2025_1 - rate_2024_1, 1)
        
        # 20-29세 고용률
        age_row = df_index[(df_index[1] == region) & (df_index[3] == '15 - 29세')]
        rate_20_29 = safe_float(age_row.iloc[0][21], 0) if not age_row.empty else 0
        
        # 연령별 증감 계산
        age_groups = []
        for age_name in age_names:
            age_data = df_index[(df_index[1] == region) & (df_index[3] == age_name)]
            if age_data.empty:
                # 60세이상이 없으면 70세이상으로 시도
                if age_name == '60세이상':
                    age_data = df_index[(df_index[1] == region) & (df_index[3] == '70세이상')]
                    if not age_data.empty:
                        age_name = '70세이상'  # 실제 엑셀의 연령대명 사용
                    else:
                        continue
                else:
                    continue
            
            arow = age_data.iloc[0]
            age_rate_2024 = safe_float(arow[17], 0)
            age_rate_2025 = safe_float(arow[21], 0)
            age_change = round(age_rate_2025 - age_rate_2024, 1)
            
            age_groups.append({
                'name': age_name,
                'display_name': AGE_GROUP_MAPPING.get(age_name, age_name),
                'change': age_change
            })
        
        # 증가/감소에 따른 정렬
        if change >= 0:
            sorted_ages = sorted([a for a in age_groups if a['change'] > 0],
                               key=lambda x: x['change'], reverse=True)
        else:
            sorted_ages = sorted([a for a in age_groups if a['change'] < 0],
                               key=lambda x: x['change'])
        
        regions.append({
            'region': region,
            'change': change,
            'change_2023_2': change_2023_2,
            'change_2024_2': change_2024_2,
            'change_2025_1': change_2025_1,
            'rate_2024': round(rate_2024_2, 1),
            'rate_2025': round(rate_2025_2, 1),
            'rate_20_29': round(rate_20_29, 1),
            'top_age_groups': sorted_ages[:3],
            'all_age_groups': age_groups
        })
    
    # 증가/감소 지역 분류
    increase_regions = sorted(
        [r for r in regions if r['change'] > 0],
        key=lambda x: x['change'],
        reverse=True
    )
    decrease_regions = sorted(
        [r for r in regions if r['change'] < 0],
        key=lambda x: x['change']
    )
    
    return {
        'increase_regions': increase_regions,
        'decrease_regions': decrease_regions,
        'all_regions': regions
    }


def get_table_data(df_analysis, df_index):
    """표에 들어갈 데이터 생성"""
    use_aggregation_only = df_analysis.attrs.get('use_aggregation_only', False)
    
    if use_aggregation_only:
        return _get_table_data_from_aggregation(df_index)
    
    region_indices = get_region_indices(df_analysis)
    
    table_data = []
    
    # 전국
    nationwide_row = df_analysis.iloc[3]
    nationwide_idx = df_index.iloc[3]
    # 15-29세 행
    age_idx = df_index.iloc[4]
    
    table_data.append({
        'group': None,
        'rowspan': None,
        'region': REGION_DISPLAY_MAPPING['전국'],
        'changes': [
            round(float(nationwide_row[10]), 1),  # 2023 2/4
            round(float(nationwide_row[14]), 1),  # 2024 2/4
            round(float(nationwide_row[17]), 1),  # 2025 1/4
            round(float(nationwide_row[18]), 1),  # 2025 2/4
        ],
        'rates': [
            nationwide_idx[17],  # 2024 2/4
            nationwide_idx[21],  # 2025 2/4
            age_idx[21],  # 20-29세 2025 2/4
        ]
    })
    
    # 지역별 그룹
    for group_name, group_regions in REGION_GROUPS.items():
        for i, region in enumerate(group_regions):
            if region not in region_indices:
                continue
                
            start_idx = region_indices[region]
            row = df_analysis.iloc[start_idx]
            idx_row = df_index[(df_index[1] == region) & (df_index[3] == '계')]
            age_row = df_index[(df_index[1] == region) & (df_index[3] == '15 - 29세')]
            
            if idx_row.empty:
                continue
                
            idx_row = idx_row.iloc[0]
            age_rate = age_row.iloc[0][21] if not age_row.empty else 0
            
            try:
                entry = {
                    'region': REGION_DISPLAY_MAPPING.get(region, region),
                    'changes': [
                        round(float(row[10]), 1),  # 2023 2/4
                        round(float(row[14]), 1),  # 2024 2/4
                        round(float(row[17]), 1),  # 2025 1/4
                        round(float(row[18]), 1),  # 2025 2/4
                    ],
                    'rates': [
                        idx_row[17],  # 2024 2/4
                        idx_row[21],  # 2025 2/4
                        age_rate,  # 20-29세
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
    """집계 시트에서 표 데이터 생성"""
    # 집계 시트 구조: 1=지역명, 3=연령구분
    # 컬럼: 9=2022.2/4, 13=2023.2/4, 16=2024.1/4, 17=2024.2/4, 20=2025.1/4, 21=2025.2/4
    
    table_data = []
    
    # 전국 데이터
    nationwide_total = df_index[(df_index[1] == '전국') & (df_index[3] == '계')]
    nationwide_age = df_index[(df_index[1] == '전국') & (df_index[3] == '15 - 29세')]
    
    if not nationwide_total.empty:
        nrow = nationwide_total.iloc[0]
        age_row = nationwide_age.iloc[0] if not nationwide_age.empty else None
        
        # 증감률 계산 (전년동분기대비 %p 변화)
        rate_2022_2 = safe_float(nrow[9], 0)
        rate_2023_2 = safe_float(nrow[13], 0)
        rate_2024_1 = safe_float(nrow[16], 0)
        rate_2024_2 = safe_float(nrow[17], 0)
        rate_2025_1 = safe_float(nrow[20], 0)
        rate_2025_2 = safe_float(nrow[21], 0)
        
        change_2023_2 = rate_2023_2 - rate_2022_2
        change_2024_2 = rate_2024_2 - rate_2023_2
        change_2025_1 = rate_2025_1 - rate_2024_1
        change_2025_2 = rate_2025_2 - rate_2024_2
        
        age_rate = safe_float(age_row[21], 0) if age_row is not None else 0
        
        table_data.append({
            'group': None,
            'rowspan': None,
            'region': REGION_DISPLAY_MAPPING['전국'],
            'changes': [
                round(change_2023_2, 1),
                round(change_2024_2, 1),
                round(change_2025_1, 1),
                round(change_2025_2, 1),
            ],
            'rates': [
                round(rate_2024_2, 1),
                round(rate_2025_2, 1),
                round(age_rate, 1),
            ]
        })
    
    # 지역별 그룹
    for group_name, group_regions in REGION_GROUPS.items():
        for i, region in enumerate(group_regions):
            region_total = df_index[(df_index[1] == region) & (df_index[3] == '계')]
            region_age = df_index[(df_index[1] == region) & (df_index[3] == '15 - 29세')]
            
            if region_total.empty:
                continue
            
            rrow = region_total.iloc[0]
            age_row = region_age.iloc[0] if not region_age.empty else None
            
            # 증감률 계산
            rate_2022_2 = safe_float(rrow[9], 0)
            rate_2023_2 = safe_float(rrow[13], 0)
            rate_2024_1 = safe_float(rrow[16], 0)
            rate_2024_2 = safe_float(rrow[17], 0)
            rate_2025_1 = safe_float(rrow[20], 0)
            rate_2025_2 = safe_float(rrow[21], 0)
            
            change_2023_2 = rate_2023_2 - rate_2022_2
            change_2024_2 = rate_2024_2 - rate_2023_2
            change_2025_1 = rate_2025_1 - rate_2024_1
            change_2025_2 = rate_2025_2 - rate_2024_2
            
            age_rate = safe_float(age_row[21], 0) if age_row is not None else 0
            
            entry = {
                'region': REGION_DISPLAY_MAPPING.get(region, region),
                'changes': [
                    round(change_2023_2, 1),
                    round(change_2024_2, 1),
                    round(change_2025_1, 1),
                    round(change_2025_2, 1),
                ],
                'rates': [
                    round(rate_2024_2, 1),
                    round(rate_2025_2, 1),
                    round(age_rate, 1),
                ]
            }
            
            if i == 0:
                entry['group'] = group_name
                entry['rowspan'] = len(group_regions)
            else:
                entry['group'] = None
                entry['rowspan'] = None
            
            table_data.append(entry)
    
    return table_data


def get_summary_box_data(regional_data):
    """요약 박스 데이터 생성"""
    # 증가 지역 상위 3개
    top3 = regional_data['increase_regions'][:3]
    main_regions = [r['region'] for r in top3]
    
    return {
        'main_increase_regions': main_regions,
        'region_count': len(regional_data['increase_regions'])
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
    #     # 기초자료에서 고용률 데이터 직접 추출
    #     # return extract_from_raw_data(extractor, ...)
    # 데이터 로드
    df_analysis, df_index = load_data(excel_path)
    
    # 데이터 추출
    nationwide_data = get_nationwide_data(df_analysis, df_index)
    regional_data = get_regional_data(df_analysis, df_index)
    summary_box = get_summary_box_data(regional_data)
    table_data = get_table_data(df_analysis, df_index)
    
    # Top 3 증가/감소 지역
    top3_increase = []
    for r in regional_data['increase_regions'][:3]:
        top3_increase.append({
            'region': r['region'],
            'change': r['change'],
            'age_groups': r['top_age_groups']
        })
    
    top3_decrease = []
    for r in regional_data['decrease_regions'][:3]:
        top3_decrease.append({
            'region': r['region'],
            'change': r['change'],
            'age_groups': r['top_age_groups']
        })
    
    # 템플릿 데이터
    template_data = {
        'summary_box': summary_box,
        'nationwide_data': nationwide_data,
        'regional_data': regional_data,
        'top3_increase_regions': top3_increase,
        'top3_decrease_regions': top3_decrease,
        'summary_table': {
            'columns': {
                'change_columns': ['2023.2/4', '2024.2/4', '2025.1/4', '2025.2/4'],
                'rate_columns': ['2024.2/4', '2025.2/4', '20-29세']
            },
            'regions': table_data
        }
    }
    
    # JSON 데이터 저장
    data_path = Path(output_path).parent / 'employment_rate_data.json'
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
    template_path = Path(__file__).parent / 'employment_rate_template.html'
    output_path = Path(__file__).parent / 'employment_rate_output.html'
    
    data = generate_report(excel_path, template_path, output_path)
    
    # 검증용 출력
    print("\n=== 전국 데이터 ===")
    print(f"고용률: {data['nationwide_data']['employment_rate']}%")
    print(f"증감: {data['nationwide_data']['change']}%p")
    print(f"주요 연령: {data['nationwide_data']['main_age_groups']}")
    
    print("\n=== 증가 지역 Top 3 ===")
    for r in data['top3_increase_regions']:
        print(f"{r['region']}({r['change']}%p): {r['age_groups']}")
    
    print("\n=== 감소 지역 Top 3 ===")
    for r in data['top3_decrease_regions']:
        print(f"{r['region']}({r['change']}%p): {r['age_groups']}")

