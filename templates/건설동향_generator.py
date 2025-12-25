#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
건설동향 보고서 생성기
F'분석 시트에서 데이터를 추출하여 HTML 보고서를 생성합니다.
"""

import pandas as pd
import json
import re
from jinja2 import Template
from pathlib import Path

# 공종 매핑
CONSTRUCTION_TYPE_MAPPING = {
    '주거용건축': '주거용건축',
    '비주거용건축': '비주거용건축',
    '토목': '토목',
    '건축': '건축',
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
    """안전한 float 변환 함수 (NaN 체크 포함)"""
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
        result = float(value)
        if pd.isna(result):
            return default
        return result
    except (ValueError, TypeError):
        return default


def safe_round(value, decimals=1, default=None):
    """안전한 반올림 함수"""
    if value is None:
        return default
    try:
        result = safe_float(value, default)
        if result is None:
            return default
        return round(result, decimals)
    except (ValueError, TypeError):
        return default


def normalize_sheet_name(name):
    """시트 이름 정규화 (공백 제거, 따옴표 정규화)"""
    if not name:
        return ""
    # 공백 제거
    normalized = name.replace(" ", "").replace("　", "")
    # 따옴표 정규화 (작은따옴표, 큰따옴표 모두 통일)
    normalized = normalized.replace("'", "'").replace('"', "'")
    return normalized


def find_sheet_by_pattern(sheet_names, patterns, sheet_type="분석"):
    """패턴 매칭으로 시트 찾기"""
    # 1. 정확한 이름 매칭
    for pattern in patterns:
        normalized_pattern = normalize_sheet_name(pattern)
        for sheet_name in sheet_names:
            if normalize_sheet_name(sheet_name) == normalized_pattern:
                return sheet_name
    
    # 2. 부분 매칭 (패턴이 시트 이름에 포함되는지)
    for pattern in patterns:
        normalized_pattern = normalize_sheet_name(pattern)
        for sheet_name in sheet_names:
            normalized_sheet = normalize_sheet_name(sheet_name)
            if normalized_pattern in normalized_sheet or normalized_sheet in normalized_pattern:
                return sheet_name
    
    # 3. 키워드 매칭 (F와 분석이 모두 포함되는지)
    for sheet_name in sheet_names:
        normalized = normalize_sheet_name(sheet_name)
        # F'분석을 찾을 때
        if sheet_type == "분석":
            if ('F' in normalized or 'f' in normalized) and '분석' in normalized:
                return sheet_name
        # F'집계를 찾을 때
        elif sheet_type == "집계":
            if ('F' in normalized or 'f' in normalized) and '집계' in normalized and '건설' in normalized:
                return sheet_name
    
    return None


def load_data(excel_path):
    """엑셀 파일에서 데이터 로드"""
    xl = pd.ExcelFile(excel_path)
    
    # 디버깅: 모든 시트 이름 출력
    print(f"[DEBUG] 건설동향 - 사용 가능한 시트 목록: {xl.sheet_names}")
    
    # F'분석 시트 찾기
    possible_analysis_sheets = ["F'분석", "F' 분석", "F분석", "F 분석"]
    analysis_sheet_name = find_sheet_by_pattern(xl.sheet_names, possible_analysis_sheets, "분석")
    
    df_analysis = None
    if analysis_sheet_name:
        try:
            print(f"[DEBUG] 건설동향 분석 시트 발견: '{analysis_sheet_name}'")
            df_analysis = pd.read_excel(xl, sheet_name=analysis_sheet_name, header=None)
        except Exception as e:
            print(f"[ERROR] 건설동향 분석 시트 읽기 실패: {e}")
            raise ValueError(f"건설동향 분석 시트를 읽을 수 없습니다: '{analysis_sheet_name}'")
    else:
        raise ValueError(f"건설동향 분석 시트를 찾을 수 없습니다. 시트 목록: {xl.sheet_names}")
    
    # F'집계 시트 찾기
    possible_index_sheets = ["F'(건설)집계", "F'(건설) 집계", "F'집계", "F' 집계", "F(건설)집계", "F집계"]
    index_sheet_name = find_sheet_by_pattern(xl.sheet_names, possible_index_sheets, "집계")
    
    df_index = None
    if index_sheet_name:
        try:
            print(f"[DEBUG] 건설동향 집계 시트 발견: '{index_sheet_name}'")
            df_index = pd.read_excel(xl, sheet_name=index_sheet_name, header=None)
        except Exception as e:
            print(f"[WARNING] 건설동향 집계 시트 읽기 실패: {e}, 분석 시트를 사용합니다.")
            df_index = None
    
    # F'집계가 없으면 F'분석을 대신 사용
    if df_index is None:
        print(f"[DEBUG] 건설동향 집계 시트를 찾지 못해 분석 시트를 사용합니다.")
        df_index = df_analysis.copy()
    
    return df_analysis, df_index


def get_region_indices(df_analysis):
    """각 지역의 시작 인덱스 찾기 (분류단계 0인 행)"""
    region_indices = {}
    
    # 디버깅: 처음 몇 행의 구조 확인
    print(f"[DEBUG] 건설동향 - 데이터프레임 shape: {df_analysis.shape}")
    if len(df_analysis) > 0:
        print(f"[DEBUG] 건설동향 - 첫 5행 샘플:")
        for i in range(min(5, len(df_analysis))):
            row = df_analysis.iloc[i]
            print(f"  행 {i}: {[str(row[j])[:20] if pd.notna(row[j]) else 'NaN' for j in range(min(10, len(row)))]}")
    
    for i in range(len(df_analysis)):
        row = df_analysis.iloc[i]
        try:
            # 분류단계 컬럼 찾기 (여러 가능한 인덱스 시도)
            classification_step = None
            for col_idx in [3, 4, 5]:  # 일반적으로 3번이지만 유연하게
                try:
                    step_val = str(row[col_idx]).strip()
                    if step_val == '0':
                        classification_step = col_idx
                        break
                except:
                    continue
            
            if classification_step is not None:
                # 지역명 컬럼 찾기 (일반적으로 2번이지만 유연하게)
                region = None
                for col_idx in [2, 1, 0]:
                    try:
                        region_val = str(row[col_idx]).strip()
                        if region_val in VALID_REGIONS or region_val == '전국':
                            region = region_val
                            break
                    except:
                        continue
                
                if region:
                    if region not in region_indices:
                        region_indices[region] = i
                        print(f"[DEBUG] 건설동향 - 지역 '{region}' 발견 (행 {i})")
        except Exception as e:
            continue
    
    print(f"[DEBUG] 건설동향 - 발견된 지역: {list(region_indices.keys())}")
    return region_indices


def get_nationwide_data(df_analysis, df_index):
    """전국 데이터 추출"""
    # 분석 시트에서 전국 총지수 행 (분류단계 0)
    region_indices = get_region_indices(df_analysis)
    nationwide_idx = region_indices.get('전국', 3)
    
    if nationwide_idx >= len(df_analysis):
        print(f"[WARNING] 건설동향 - 전국 인덱스 {nationwide_idx}가 범위를 벗어남 (총 행수: {len(df_analysis)})")
        nationwide_idx = 3
    
    nationwide_row = df_analysis.iloc[nationwide_idx]
    
    # 컬럼 인덱스 확인 및 디버깅
    if len(nationwide_row) <= 19:
        print(f"[WARNING] 건설동향 - 전국 행의 컬럼 수가 부족함: {len(nationwide_row)} (최소 20개 필요)")
    
    growth_rate = safe_round(nationwide_row[19] if len(nationwide_row) > 19 else None, 1, 0.0)  # 2025.2/4p
    print(f"[DEBUG] 건설동향 - 전국 증감률: {growth_rate}%")
    
    # 집계 시트에서 전국 지수
    nationwide_idx_row = df_index[df_index[2] == '전국']
    if not nationwide_idx_row.empty:
        idx_row = nationwide_idx_row.iloc[0]
        if len(idx_row) > 22:
            construction_index = safe_float(idx_row[22], 0.0)  # 2025.2/4p
        else:
            print(f"[WARNING] 건설동향 - 집계 행의 컬럼 수가 부족함: {len(idx_row)} (최소 23개 필요)")
            construction_index = 0.0
    else:
        print(f"[WARNING] 건설동향 - 집계 시트에서 '전국' 행을 찾을 수 없음")
        construction_index = 0.0
    
    # 전국 공종별 증감률 (분류단계 1인 행들)
    construction_types = []
    for i in range(nationwide_idx + 1, min(nationwide_idx + 5, len(df_analysis))):
        row = df_analysis.iloc[i]
        if str(row[3]) != '1':  # 분류단계 1이 아니면 스킵
            continue
        type_name = row[5]  # 공종명
        type_growth = safe_round(row[19], 1)
        if type_growth is None:
            continue
        construction_types.append({
            'name': CONSTRUCTION_TYPE_MAPPING.get(type_name, type_name),
            'growth_rate': type_growth
        })
    
    # 증감 큰 순으로 정렬
    if growth_rate is not None and growth_rate < 0:
        # 감소 시 감소율이 큰 순
        negative_types = [t for t in construction_types if t['growth_rate'] is not None and t['growth_rate'] < 0]
        negative_types.sort(key=lambda x: x['growth_rate'] if x['growth_rate'] is not None else 0)
        main_types = negative_types[:3]
    else:
        # 증가 시 증가율이 큰 순
        positive_types = [t for t in construction_types if t['growth_rate'] is not None and t['growth_rate'] > 0]
        positive_types.sort(key=lambda x: x['growth_rate'] if x['growth_rate'] is not None else 0, reverse=True)
        main_types = positive_types[:3]
    
    return {
        'construction_index': construction_index if construction_index is not None else 0.0,
        'growth_rate': growth_rate if growth_rate is not None else 0.0,
        'main_types': main_types if main_types else construction_types[:3]
    }


def get_regional_data(df_analysis, df_index):
    """시도별 데이터 추출"""
    region_indices = get_region_indices(df_analysis)
    regions = []
    
    for region, start_idx in region_indices.items():
        if region == '전국':
            continue
            
        # 분류단계 0 행에서 증감률
        total_row = df_analysis.iloc[start_idx]
        growth_rate = safe_round(total_row[19], 1)  # 2025.2/4p
        growth_2023_2 = safe_round(total_row[11], 1)  # 2023 2/4
        growth_2024_2 = safe_round(total_row[15], 1)  # 2024 2/4
        growth_2025_1 = safe_round(total_row[18], 1)  # 2025 1/4
        
        # 모든 값이 None이면 스킵하지 않고 기본값 0.0 사용
        if growth_rate is None:
            growth_rate = 0.0
        if growth_2023_2 is None:
            growth_2023_2 = 0.0
        if growth_2024_2 is None:
            growth_2024_2 = 0.0
        if growth_2025_1 is None:
            growth_2025_1 = 0.0
        
        # 집계 시트에서 지수
        idx_row = df_index[df_index[2] == region]
        if not idx_row.empty:
            index_2024 = safe_float(idx_row.iloc[0][19], 0.0)
            index_2025 = safe_float(idx_row.iloc[0][22], 0.0)
        else:
            index_2024 = 0.0
            index_2025 = 0.0
        
        # 공종별 증감률
        construction_types = []
        next_region_idx = len(df_analysis)
        for other_region, other_idx in region_indices.items():
            if other_idx > start_idx and other_idx < next_region_idx:
                next_region_idx = other_idx
        
        for i in range(start_idx + 1, min(start_idx + 5, next_region_idx)):
            if i >= len(df_analysis):
                break
            row = df_analysis.iloc[i]
            if str(row[3]) != '1':
                continue
            type_name = row[5]
            type_growth = safe_round(row[19], 1)
            if type_growth is None:
                continue
            
            construction_types.append({
                'name': CONSTRUCTION_TYPE_MAPPING.get(type_name, type_name),
                'growth_rate': type_growth
            })
        
        # 증가/감소 지역에 따라 정렬
        if growth_rate >= 0:
            positive_types = [t for t in construction_types if t['growth_rate'] is not None and t['growth_rate'] > 0]
            positive_types.sort(key=lambda x: x['growth_rate'] if x['growth_rate'] is not None else 0, reverse=True)
            top_types = positive_types[:3]
        else:
            negative_types = [t for t in construction_types if t['growth_rate'] is not None and t['growth_rate'] < 0]
            negative_types.sort(key=lambda x: x['growth_rate'] if x['growth_rate'] is not None else 0)
            top_types = negative_types[:3]
        
        if not top_types:
            top_types = construction_types[:3]
        
        regions.append({
            'region': region,
            'growth_rate': growth_rate,
            'growth_2023_2': growth_2023_2,
            'growth_2024_2': growth_2024_2,
            'growth_2025_1': growth_2025_1,
            'index_2024': index_2024 if index_2024 is not None else 0.0,
            'index_2025': index_2025 if index_2025 is not None else 0.0,
            'top_types': top_types,
            'all_types': construction_types
        })
    
    # 증가/감소 지역 분류
    increase_regions = sorted(
        [r for r in regions if r['growth_rate'] is not None and r['growth_rate'] > 0],
        key=lambda x: x['growth_rate'] if x['growth_rate'] is not None else 0,
        reverse=True
    )
    decrease_regions = sorted(
        [r for r in regions if r['growth_rate'] is not None and r['growth_rate'] < 0],
        key=lambda x: x['growth_rate'] if x['growth_rate'] is not None else 0
    )
    
    return {
        'increase_regions': increase_regions,
        'decrease_regions': decrease_regions,
        'all_regions': regions
    }


def get_growth_rates_table(df_analysis, df_index):
    """표에 들어갈 증감률 및 지수 데이터 생성"""
    region_indices = get_region_indices(df_analysis)
    
    table_data = []
    
    # 전국
    nationwide_idx = region_indices.get('전국', 3)
    nationwide_row = df_analysis.iloc[nationwide_idx]
    nationwide_idx_row = df_index[df_index[2] == '전국']
    
    if not nationwide_idx_row.empty:
        nationwide_idx_row = nationwide_idx_row.iloc[0]
        table_data.append({
            'group': None,
            'rowspan': None,
            'region': REGION_DISPLAY_MAPPING['전국'],
            'growth_rates': [
                safe_round(nationwide_row[11], 1),  # 2023 2/4
                safe_round(nationwide_row[15], 1),  # 2024 2/4
                safe_round(nationwide_row[18], 1),  # 2025 1/4
                safe_round(nationwide_row[19], 1),  # 2025 2/4
            ],
            'indices': [
                safe_float(nationwide_idx_row[19], None),  # 2024 2/4
                safe_float(nationwide_idx_row[22], None),  # 2025 2/4
            ]
        })
    
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
            
            entry = {
                'region': REGION_DISPLAY_MAPPING.get(region, region),
                'growth_rates': [
                    safe_round(row[11], 1),  # 2023 2/4
                    safe_round(row[15], 1),  # 2024 2/4
                    safe_round(row[18], 1),  # 2025 1/4
                    safe_round(row[19], 1),  # 2025 2/4
                ],
                'indices': [
                    safe_float(idx_row[19], None),  # 2024 2/4
                    safe_float(idx_row[22], None),  # 2025 2/4
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
    # 증가/감소 지역 상위
    top3_increase = regional_data['increase_regions'][:3]
    top3_decrease = regional_data['decrease_regions'][:3]
    
    main_increase_regions = []
    for r in top3_increase:
        main_type = ''
        if r.get('top_types') and len(r['top_types']) > 0:
            main_type = r['top_types'][0].get('name', '')
        main_increase_regions.append({
            'region': r['region'],
            'main_type': main_type if main_type else 'nan'
        })
    
    main_decrease_regions = []
    for r in top3_decrease:
        main_type = ''
        if r.get('top_types') and len(r['top_types']) > 0:
            main_type = r['top_types'][0].get('name', '')
        main_decrease_regions.append({
            'region': r['region'],
            'main_type': main_type if main_type else 'nan'
        })
    
    return {
        'main_increase_regions': main_increase_regions,
        'main_decrease_regions': main_decrease_regions,
        'increase_count': len(regional_data['increase_regions']),
        'decrease_count': len(regional_data['decrease_regions'])
    }


def generate_report_data(excel_path):
    """보고서 데이터 생성 (app.py에서 호출)"""
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
            'types': r['top_types']
        })
    
    top3_decrease = []
    for r in regional_data['decrease_regions'][:3]:
        top3_decrease.append({
            'region': r['region'],
            'growth_rate': r['growth_rate'],
            'types': r['top_types']
        })
    
    # 증가/감소 공종 텍스트 생성
    increase_types = set()
    for r in regional_data['increase_regions'][:3]:
        for t in r['top_types'][:2]:
            type_name = t.get('name', '')
            if isinstance(type_name, str) and type_name:
                increase_types.add(type_name)
    increase_types_text = ', '.join([str(t) for t in list(increase_types)[:3]]) if increase_types else ''
    
    decrease_types = set()
    for r in regional_data['decrease_regions'][:3]:
        for t in r['top_types'][:2]:
            type_name = t.get('name', '')
            if isinstance(type_name, str) and type_name:
                decrease_types.add(type_name)
    decrease_types_text = ', '.join([str(t) for t in list(decrease_types)[:3]]) if decrease_types else ''
    
    # 템플릿 데이터
    return {
        'summary_box': summary_box,
        'nationwide_data': nationwide_data,
        'regional_data': regional_data,
        'top3_increase_regions': top3_increase,
        'top3_decrease_regions': top3_decrease,
        'increase_types_text': increase_types_text,
        'decrease_types_text': decrease_types_text,
        'summary_table': {
            'base_year': 2020,
            'columns': {
                'growth_rate_columns': ['2023.2/4', '2024.2/4', '2025.1/4', '2025.2/4p'],
                'index_columns': ['2024.2/4', '2025.2/4p']
            },
            'regions': table_data
        }
    }


def generate_report(excel_path, template_path, output_path):
    """보고서 생성"""
    template_data = generate_report_data(excel_path)
    
    # JSON 데이터 저장
    data_path = Path(output_path).parent / '건설동향_data.json'
    with open(data_path, 'w', encoding='utf-8') as f:
        json.dump(template_data, f, ensure_ascii=False, indent=2, default=str)
    
    # 템플릿 렌더링
    with open(template_path, 'r', encoding='utf-8') as f:
        template = Template(f.read())
    
    html_output = template.render(**template_data)
    
    # HTML 저장
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_output)
    
    print(f"보고서 생성 완료: {output_path}")
    print(f"데이터 파일 저장: {data_path}")
    
    return template_data


if __name__ == '__main__':
    base_path = Path(__file__).parent.parent
    excel_path = base_path / '분석표_25년 2분기_캡스톤.xlsx'
    template_path = Path(__file__).parent / '건설동향_template.html'
    output_path = Path(__file__).parent / '건설동향_output.html'
    
    data = generate_report(excel_path, template_path, output_path)
    
    # 검증용 출력
    print("\n=== 전국 데이터 ===")
    print(f"건설수주: {data['nationwide_data']['construction_index']}")
    print(f"증감률: {data['nationwide_data']['growth_rate']}%")
    print(f"주요 공종: {data['nationwide_data']['main_types']}")
    
    print("\n=== 증가 지역 Top 3 ===")
    for r in data['top3_increase_regions']:
        print(f"{r['region']}({r['growth_rate']}%): {r['types']}")
    
    print("\n=== 감소 지역 Top 3 ===")
    for r in data['top3_decrease_regions']:
        print(f"{r['region']}({r['growth_rate']}%): {r['types']}")

