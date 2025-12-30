#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
서비스업생산 보고서 생성기
B 분석 시트에서 데이터를 추출하여 HTML 보고서를 생성합니다.
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
    
    return df_analysis, df_index


def get_region_indices(df_analysis):
    """각 지역의 시작 인덱스 찾기"""
    region_indices = {}
    for i in range(len(df_analysis)):
        row = df_analysis.iloc[i]
        if row[7] == '총지수':
            region = row[3]
            region_indices[region] = i
    return region_indices


def get_nationwide_data(df_analysis, df_index):
    """전국 데이터 추출"""
    use_raw = df_analysis.attrs.get('use_raw', False)
    
    if use_raw:
        return _get_nationwide_from_raw_data(df_analysis)
    
    # 분석 시트에서 전국 총지수 행
    nationwide_row = df_analysis.iloc[3]
    growth_rate = safe_float(nationwide_row[20], 0)
    growth_rate = round(growth_rate, 1) if growth_rate else 0.0
    
    # 집계 시트에서 전국 지수
    index_row = df_index.iloc[3]
    production_index = safe_float(index_row[25], 100)  # 2025.2/4p
    
    # 전국 주요 업종 (기여도 기준 상위 3개 - 양수만)
    industries = []
    for i in range(4, 17):
        row = df_analysis.iloc[i]
        industry_name = row[7]
        industry_growth = safe_float(row[20], 0)
        if industry_growth is not None:
            industries.append({
                'name': INDUSTRY_MAPPING.get(industry_name, industry_name),
                'growth_rate': round(industry_growth, 1)
            })
    
    # 양수 증가율 중 상위 3개 (보건·복지, 금융·보험, 운수·창고 순)
    positive_industries = [i for i in industries if i['growth_rate'] > 0]
    positive_industries.sort(key=lambda x: x['growth_rate'], reverse=True)
    main_industries = positive_industries[:3]
    
    return {
        'production_index': production_index,
        'growth_rate': growth_rate,
        'main_industries': main_industries
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
    region_indices = get_region_indices(df_analysis)
    regions = []
    
    for region, start_idx in region_indices.items():
        if region == '전국':
            continue
            
        # 총지수 행에서 증감률
        total_row = df_analysis.iloc[start_idx]
        growth_rate = round(total_row[20], 1)
        
        # 집계 시트에서 지수
        idx_row = df_index[df_index[3] == region]
        if not idx_row.empty:
            index_2024 = idx_row.iloc[0][21]
            index_2025 = idx_row.iloc[0][25]
        else:
            index_2024 = 0
            index_2025 = 0
        
        # 업종별 기여도
        industries = []
        for i in range(start_idx + 1, start_idx + 14):
            if i >= len(df_analysis):
                break
            row = df_analysis.iloc[i]
            if row[4] != '1':  # 분류단계가 1이 아니면 스킵
                continue
            industry_name = row[7]
            industry_growth = row[20]
            contribution = row[26]
            
            if pd.notna(contribution):
                industries.append({
                    'name': INDUSTRY_MAPPING.get(industry_name, industry_name),
                    'growth_rate': round(industry_growth, 1),
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
    region_indices = get_region_indices(df_analysis)
    
    table_data = []
    
    # 전국
    nationwide_row = df_analysis.iloc[3]
    nationwide_idx = df_index.iloc[3]
    table_data.append({
        'group': None,
        'rowspan': None,
        'region': REGION_DISPLAY_MAPPING['전국'],
        'growth_rates': [
            round(nationwide_row[12], 1),  # 2023 2/4
            round(nationwide_row[16], 1),  # 2024 2/4
            round(nationwide_row[19], 1),  # 2025 1/4
            round(nationwide_row[20], 1),  # 2025 2/4
        ],
        'indices': [
            nationwide_idx[21],  # 2024 2/4
            nationwide_idx[25],  # 2025 2/4
        ]
    })
    
    # 지역별 그룹
    for group_name, group_regions in REGION_GROUPS.items():
        for i, region in enumerate(group_regions):
            if region not in region_indices:
                continue
                
            start_idx = region_indices[region]
            row = df_analysis.iloc[start_idx]
            idx_row = df_index[df_index[3] == region]
            
            if idx_row.empty:
                continue
                
            idx_row = idx_row.iloc[0]
            
            entry = {
                'region': REGION_DISPLAY_MAPPING.get(region, region),
                'growth_rates': [
                    round(row[12], 1),  # 2023 2/4
                    round(row[16], 1),  # 2024 2/4
                    round(row[19], 1),  # 2025 1/4
                    round(row[20], 1),  # 2025 2/4
                ],
                'indices': [
                    idx_row[21],  # 2024 2/4
                    idx_row[25],  # 2025 2/4
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


def generate_report(excel_path, template_path, output_path, raw_excel_path=None, year=None, quarter=None):
    """보고서 생성
    
    Args:
        excel_path: 분석표 엑셀 파일 경로
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        raw_excel_path: 기초자료 엑셀 파일 경로 (선택사항, 향후 기초자료 직접 추출 지원 예정)
        year: 현재 연도 (선택사항)
        quarter: 현재 분기 (선택사항)
    """
    # TODO: 향후 기초자료 직접 추출 지원
    # if raw_excel_path and year and quarter:
    #     from raw_data_extractor import RawDataExtractor
    #     extractor = RawDataExtractor(raw_excel_path, year, quarter)
    #     # 기초자료에서 서비스업생산 데이터 직접 추출
    #     # return extract_from_raw_data(extractor, ...)
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
    
    print(f"보고서 생성 완료: {output_path}")
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

