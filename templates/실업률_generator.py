#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
실업률 보고서 생성기
D(실업)분석 및 D(실업)집계 시트에서 데이터를 추출하여 HTML 보고서를 생성합니다.
"""

import pandas as pd
import json
from jinja2 import Environment, FileSystemLoader
import os

# 연령대 이름 매핑
AGE_GROUP_MAPPING = {
    "15 - 29세": "15~29세",
    "30 - 59세": "30~59세",
    "60세이상": "60세이상"
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
    summary_df = pd.read_excel(excel_path, sheet_name='D(실업)집계', header=None)
    analysis_df = pd.read_excel(excel_path, sheet_name='D(실업)분석', header=None)
    return summary_df, analysis_df

def get_unemployment_rate_data(summary_df):
    """집계 시트에서 실업률 데이터를 추출합니다."""
    # 실업률 데이터는 79행부터 시작
    rate_data = {}
    
    for i in range(80, 152):
        row = summary_df.iloc[i]
        sido = row[0]
        age_group = row[1]
        
        if sido not in rate_data:
            rate_data[sido] = {}
        
        rate_data[sido][age_group] = {
            'rate_2023_24': row[11],
            'rate_2024_24': row[15],
            'rate_2025_14': row[18],
            'rate_2025_24': row[19],
            'change': row[19] - row[15]  # 2025.2/4 - 2024.2/4
        }
    
    return rate_data

def get_nationwide_data(rate_data):
    """전국 데이터를 추출합니다."""
    nationwide = rate_data.get('전국', {})
    total = nationwide.get('계', {})
    
    # 연령대별 증감 (절대값 기준 내림차순 정렬)
    age_groups = []
    for age_key in ['15 - 29세', '30 - 59세', '60세이상']:
        if age_key in nationwide:
            age_groups.append({
                'name': AGE_GROUP_MAPPING.get(age_key, age_key),
                'change': nationwide[age_key]['change']
            })
    
    # 변화가 큰 순서로 정렬 (전국 실업률 하락이므로 감소한 연령대 중심)
    # 음수 값이 크면(절대값 기준) 먼저 오도록 정렬
    age_groups.sort(key=lambda x: x['change'])
    
    return {
        'rate': total.get('rate_2025_24', 0),
        'change': total.get('change', 0),
        'age_groups': age_groups[:2]  # 상위 2개
    }

def get_regional_data(rate_data):
    """시도별 데이터를 추출하고 증가/감소 지역으로 분류합니다."""
    increase_regions = []
    decrease_regions = []
    
    for sido in SIDO_ORDER:
        if sido == '전국':
            continue
        
        sido_data = rate_data.get(sido, {})
        total = sido_data.get('계', {})
        
        if not total:
            continue
        
        change = total.get('change', 0)
        
        # 연령대별 증감 추출
        age_groups = []
        for age_key in ['15 - 29세', '30 - 59세', '60세이상']:
            if age_key in sido_data:
                age_groups.append({
                    'name': AGE_GROUP_MAPPING.get(age_key, age_key),
                    'change': sido_data[age_key]['change']
                })
        
        region_info = {
            'name': sido,
            'change': change,
            'rate': total.get('rate_2025_24', 0),
            'age_groups': age_groups
        }
        
        if change > 0:
            # 증가 지역: 양수 변화가 큰 연령대 순
            region_info['age_groups'] = sorted(age_groups, key=lambda x: -x['change'])
            increase_regions.append(region_info)
        elif change < 0:
            # 감소 지역: 음수 변화가 큰 연령대 순 (절대값 기준)
            region_info['age_groups'] = sorted(age_groups, key=lambda x: x['change'])
            decrease_regions.append(region_info)
    
    # 증가 지역: 증가율이 큰 순서로 정렬 (동일 증감률인 경우 이미지 순서: 광주, 세종, 경북)
    # 이미지 순서에 맞추기 위해 SIDO_ORDER 인덱스로 정렬
    increase_regions.sort(key=lambda x: (-round(x['change'], 1), SIDO_ORDER.index(x['name'])))
    
    # 감소 지역: 감소율이 큰 순서로 정렬 (절대값 기준)
    decrease_regions.sort(key=lambda x: (x['change'], SIDO_ORDER.index(x['name'])))
    
    return increase_regions, decrease_regions

def generate_summary_box(nationwide_data, increase_regions, decrease_regions):
    """요약 박스 텍스트를 생성합니다."""
    # 실업률 하락 = 긍정적 의미이므로 하락 지역이 헤드라인
    decrease_count = len(decrease_regions)
    decrease_names = ", ".join([r['name'] for r in decrease_regions[:3]])
    
    headline = f"실업률은 {decrease_names} 등 {decrease_count}개 시도에서 전년동분기대비 하락"
    
    # 전국 요약
    rate = nationwide_data['rate']
    change = nationwide_data['change']
    direction = "상승" if change > 0 else "하락"
    
    # 전국에서 하락한 연령대
    decreased_ages = [ag for ag in nationwide_data['age_groups'] if ag['change'] < 0]
    age_names = ", ".join([ag['name'] for ag in decreased_ages])
    
    nationwide_summary = f"전국 실업률은 <span class='bold'>{rate:.1f}%</span>로, {age_names} 연령대에서 실업률이 내려 전년동분기대비 <span class='bold'>{abs(change):.1f}%p {direction}</span>"
    
    # 시도 요약
    increase_names = ", ".join([f"<span class='bold'>{r['name']}</span>({r['change']:.1f}%p)" for r in increase_regions[:3]])
    decrease_names_detail = ", ".join([f"<span class='bold'>{r['name']}</span>({r['change']:.1f}%p)" for r in decrease_regions[:3]])
    
    regional_summary = f"{increase_names} 등의 실업률은 상승하였으나, {decrease_names_detail} 등의 실업률은 하락"
    
    return {
        'headline': headline,
        'nationwide_summary': nationwide_summary,
        'regional_summary': regional_summary
    }

def generate_summary_table(summary_df, rate_data):
    """요약 테이블 데이터를 생성합니다."""
    rows = []
    
    # 전국 행
    nationwide = rate_data.get('전국', {})
    total = nationwide.get('계', {})
    youth = nationwide.get('15 - 29세', {})
    
    rows.append({
        'region_group': '전 국',
        'sido': '',
        'changes': [
            summary_df.iloc[80, 11] - summary_df.iloc[80, 7],   # 2023.2/4 증감 
            summary_df.iloc[80, 15] - summary_df.iloc[80, 11],  # 2024.2/4 증감
            summary_df.iloc[80, 18] - summary_df.iloc[80, 17],  # 2025.1/4 증감
            total.get('change', 0)                               # 2025.2/4 증감
        ],
        'rates': [
            total.get('rate_2024_24', 0),
            total.get('rate_2025_24', 0)
        ],
        'youth_rate': youth.get('rate_2025_24', 0)
    })
    
    # 권역별 시도
    region_group_order = ['경인', '충청', '호남', '동북', '동남']
    
    for group_name in region_group_order:
        sidos = REGION_GROUPS[group_name]
        for idx, sido in enumerate(sidos):
            sido_data = rate_data.get(sido, {})
            total = sido_data.get('계', {})
            youth = sido_data.get('15 - 29세', {})
            
            # 해당 시도의 행 인덱스 찾기
            row_idx = None
            for i in range(80, 152, 4):
                if summary_df.iloc[i, 0] == sido:
                    row_idx = i
                    break
            
            if row_idx is None:
                continue
            
            # 증감 계산
            changes = [
                summary_df.iloc[row_idx, 11] - summary_df.iloc[row_idx, 7] if pd.notna(summary_df.iloc[row_idx, 11]) and pd.notna(summary_df.iloc[row_idx, 7]) else None,
                summary_df.iloc[row_idx, 15] - summary_df.iloc[row_idx, 11] if pd.notna(summary_df.iloc[row_idx, 15]) and pd.notna(summary_df.iloc[row_idx, 11]) else None,
                summary_df.iloc[row_idx, 18] - summary_df.iloc[row_idx, 17] if pd.notna(summary_df.iloc[row_idx, 18]) and pd.notna(summary_df.iloc[row_idx, 17]) else None,
                total.get('change', None)
            ]
            
            rows.append({
                'region_group': group_name if idx == 0 else '',
                'sido': sido.replace('', ' ') if len(sido) == 2 else sido,
                'changes': changes,
                'rates': [
                    total.get('rate_2024_24', None),
                    total.get('rate_2025_24', None)
                ],
                'youth_rate': youth.get('rate_2025_24', None)
            })
    
    return {'rows': rows}

def filter_age_groups(age_groups, is_increase_region):
    """연령대 필터링: 해당 방향의 연령대만 포함하고 상위 2개 선택"""
    if is_increase_region:
        # 증가 지역: 양수 변화만
        filtered = [ag for ag in age_groups if ag['change'] > 0]
        filtered.sort(key=lambda x: -x['change'])
    else:
        # 감소 지역: 음수 변화만
        filtered = [ag for ag in age_groups if ag['change'] < 0]
        filtered.sort(key=lambda x: x['change'])
    
    return filtered[:2] if filtered else age_groups[:2]

def generate_report_data(excel_path):
    """보고서 데이터를 생성합니다."""
    summary_df, analysis_df = load_data(excel_path)
    
    # 실업률 데이터 추출
    rate_data = get_unemployment_rate_data(summary_df)
    
    # 전국 데이터
    nationwide_data = get_nationwide_data(rate_data)
    
    # 시도별 데이터
    increase_regions, decrease_regions = get_regional_data(rate_data)
    
    # Top 3 지역 추출 및 연령대 필터링
    top3_increase = []
    for region in increase_regions[:3]:
        filtered_ages = filter_age_groups(region['age_groups'], is_increase_region=True)
        top3_increase.append({
            'name': region['name'],
            'change': region['change'],
            'age_groups': filtered_ages
        })
    
    top3_decrease = []
    for region in decrease_regions[:3]:
        filtered_ages = filter_age_groups(region['age_groups'], is_increase_region=False)
        top3_decrease.append({
            'name': region['name'],
            'change': region['change'],
            'age_groups': filtered_ages
        })
    
    # 요약 박스
    summary_box = generate_summary_box(nationwide_data, increase_regions, decrease_regions)
    
    # 요약 테이블
    summary_table = generate_summary_table(summary_df, rate_data)
    
    return {
        'report_info': {
            'section_number': '나',
            'section_title': '실업률',
            'page_number': '- 14 -'
        },
        'nationwide_data': nationwide_data,
        'regional_data': {
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions
        },
        'summary_box': summary_box,
        'top3_increase_regions': top3_increase,
        'top3_decrease_regions': top3_decrease,
        'summary_table': summary_table
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
    
    print(f"HTML 보고서가 생성되었습니다: {output_path}")

def main():
    # 경로 설정
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base_dir, '..', '분석표_25년 2분기_캡스톤.xlsx')
    template_path = os.path.join(base_dir, '실업률_template.html')
    output_path = os.path.join(base_dir, '실업률_output.html')
    json_output_path = os.path.join(base_dir, '실업률_data.json')
    
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
    print(f"전국 실업률: {data['nationwide_data']['rate']:.1f}%")
    print(f"전국 증감: {data['nationwide_data']['change']:.1f}%p")
    print(f"\n증가 지역 수: {len(data['regional_data']['increase_regions'])}")
    print(f"감소 지역 수: {len(data['regional_data']['decrease_regions'])}")
    
    print("\n=== Top 3 증가 지역 ===")
    for region in data['top3_increase_regions']:
        ages = ", ".join([f"{ag['name']}({ag['change']:.1f}%p)" for ag in region['age_groups']])
        print(f"  {region['name']}({region['change']:.1f}%p): {ages}")
    
    print("\n=== Top 3 감소 지역 ===")
    for region in data['top3_decrease_regions']:
        ages = ", ".join([f"{ag['name']}({ag['change']:.1f}%p)" for ag in region['age_groups']])
        print(f"  {region['name']}({region['change']:.1f}%p): {ages}")

if __name__ == '__main__':
    main()

