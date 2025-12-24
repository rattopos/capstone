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


def load_data(excel_path):
    """엑셀 파일에서 데이터 로드"""
    xl = pd.ExcelFile(excel_path)
    df_analysis = pd.read_excel(xl, sheet_name='B 분석', header=None)
    df_index = pd.read_excel(xl, sheet_name='B(서비스업생산)집계', header=None)
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
    # 분석 시트에서 전국 총지수 행
    nationwide_row = df_analysis.iloc[3]
    growth_rate = round(nationwide_row[20], 1)
    
    # 집계 시트에서 전국 지수
    index_row = df_index.iloc[3]
    production_index = index_row[25]  # 2025.2/4p
    
    # 전국 주요 업종 (기여도 기준 상위 3개 - 양수만)
    industries = []
    for i in range(4, 17):
        row = df_analysis.iloc[i]
        industry_name = row[7]
        industry_growth = row[20]
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
    data_path = Path(output_path).parent / '서비스업생산_data.json'
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
    template_path = Path(__file__).parent / '서비스업생산_template.html'
    output_path = Path(__file__).parent / '서비스업생산_output.html'
    
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

