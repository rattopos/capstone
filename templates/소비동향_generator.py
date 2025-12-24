#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
소비동향 보고서 생성기
C 분석 시트에서 데이터를 추출하여 HTML 보고서를 생성합니다.
"""

import pandas as pd
import json
from jinja2 import Template
from pathlib import Path

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


def load_data(excel_path):
    """엑셀 파일에서 데이터 로드"""
    xl = pd.ExcelFile(excel_path)
    df_analysis = pd.read_excel(xl, sheet_name='C 분석', header=None)
    df_index = pd.read_excel(xl, sheet_name='C(소비)집계', header=None)
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


def get_nationwide_data(df_analysis, df_index):
    """전국 데이터 추출"""
    # 분석 시트에서 전국 총지수 행
    nationwide_row = df_analysis.iloc[3]
    growth_rate = round(float(nationwide_row[20]), 1)
    
    # 집계 시트에서 전국 지수
    index_row = df_index.iloc[3]
    sales_index = index_row[24]  # 2025.2/4p
    
    # 전국 업태별 증감률
    businesses = []
    for i in range(4, 12):
        row = df_analysis.iloc[i]
        business_name = row[7]
        business_growth = float(row[20])
        businesses.append({
            'name': BUSINESS_MAPPING.get(business_name, business_name),
            'growth_rate': round(business_growth, 1)
        })
    
    # 감소율이 큰 순으로 정렬 (음수 중 절대값이 큰 것)
    negative_businesses = [b for b in businesses if b['growth_rate'] < 0]
    negative_businesses.sort(key=lambda x: x['growth_rate'])
    main_businesses = negative_businesses[:3]
    
    return {
        'sales_index': sales_index,
        'growth_rate': growth_rate,
        'main_businesses': main_businesses
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


def generate_report(excel_path, template_path, output_path):
    """보고서 생성"""
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
    data_path = Path(output_path).parent / '소비동향_data.json'
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
    template_path = Path(__file__).parent / '소비동향_template.html'
    output_path = Path(__file__).parent / '소비동향_output.html'
    
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

