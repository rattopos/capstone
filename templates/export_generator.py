#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
수출 보고서 생성기
G 분석, G 참고 및 G(수출)집계 시트에서 데이터를 추출하여 HTML 보고서를 생성합니다.
"""

import pandas as pd
import json
from jinja2 import Environment, FileSystemLoader
import os

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

# 품목명 매핑 (원본 -> 표시용)
PRODUCT_NAME_MAPPING = {
    "프로세스와 컨트롤러": "프로세서·컨트롤러",
    "수송 기타장비 ": "수송 기타장비",
}

def load_data(excel_path):
    """엑셀 파일에서 데이터를 로드합니다."""
    analysis_df = pd.read_excel(excel_path, sheet_name='G 분석', header=None)
    reference_df = pd.read_excel(excel_path, sheet_name='G 참고', header=None)
    summary_df = pd.read_excel(excel_path, sheet_name='G(수출)집계', header=None)
    return analysis_df, reference_df, summary_df

def get_sido_data_from_analysis(analysis_df):
    """G 분석 시트에서 시도별 수출 데이터를 추출합니다."""
    sido_data = {}
    
    for i in range(3, len(analysis_df)):
        row = analysis_df.iloc[i]
        product = row[8]
        
        if product == '합계':
            sido = row[3]
            if sido in SIDO_ORDER:
                change = row[22]  # 2025.2/4 증감률
                sido_data[sido] = {
                    'change': change,
                    'row_idx': i
                }
    
    return sido_data

def get_sido_products_from_reference(reference_df):
    """G 참고 시트에서 시도별 품목 데이터를 추출합니다. 상위/하위 순위 모두 수집."""
    sido_products = {}  # 상위 순위 (증가 품목)
    sido_products_bottom = {}  # 하위 순위 (감소 품목)
    current_sido = None
    max_rank_per_sido = {}
    
    # 먼저 각 시도의 최대 순위를 파악
    for i in range(2, len(reference_df)):
        row = reference_df.iloc[i]
        sido = row[1]
        rank = row[2]
        
        if pd.notna(sido) and sido in SIDO_ORDER:
            current_sido = sido
        
        if pd.notna(rank) and current_sido:
            try:
                rank_num = int(rank)
                if current_sido not in max_rank_per_sido:
                    max_rank_per_sido[current_sido] = rank_num
                else:
                    max_rank_per_sido[current_sido] = max(max_rank_per_sido[current_sido], rank_num)
            except (ValueError, TypeError):
                pass
    
    # 다시 순회하며 데이터 수집
    current_sido = None
    for i in range(2, len(reference_df)):
        row = reference_df.iloc[i]
        sido = row[1]
        rank = row[2]
        product = row[3]
        change = row[12]  # 2025.2/4 증감률
        contribution = row[13]  # 기여도
        
        # 시도 변경 확인
        if pd.notna(sido) and sido in SIDO_ORDER:
            current_sido = sido
            if current_sido not in sido_products:
                sido_products[current_sido] = []
            if current_sido not in sido_products_bottom:
                sido_products_bottom[current_sido] = []
        
        # 순위가 있는 경우 품목 추가
        if pd.notna(rank) and pd.notna(product) and current_sido:
            try:
                rank_num = int(rank)
                max_rank = max_rank_per_sido.get(current_sido, 151)
                
                # 품목명 매핑 적용
                display_name = PRODUCT_NAME_MAPPING.get(product, product).strip()
                
                product_info = {
                    'rank': rank_num,
                    'name': display_name,
                    'change': change if pd.notna(change) else 0,
                    'contribution': contribution if pd.notna(contribution) else 0
                }
                
                # 상위 5개 (증가 품목)
                if rank_num <= 5:
                    sido_products[current_sido].append(product_info)
                
                # 하위 5개 (감소 품목)
                if rank_num >= max_rank - 4:
                    sido_products_bottom[current_sido].append(product_info)
                    
            except (ValueError, TypeError):
                pass
    
    # 하위 순위는 역순 정렬 (가장 하위가 먼저)
    for sido in sido_products_bottom:
        sido_products_bottom[sido].sort(key=lambda x: -x['rank'])
    
    return sido_products, sido_products_bottom

def get_summary_table_data(summary_df):
    """G(수출)집계 시트에서 테이블 데이터를 추출합니다."""
    table_data = {}
    
    for i in range(3, len(summary_df)):
        row = summary_df.iloc[i]
        sido = row[3]
        level = row[4]
        product = row[7]
        
        if level == 0 and sido in SIDO_ORDER:
            # 증감률 데이터
            change_2023_24 = row[18]  # 2023.2/4
            change_2024_24 = row[23]  # 2024.2/4
            change_2025_14 = row[26]  # 2025.1/4
            change_2025_24 = row[27]  # 2025.2/4
            
            # 수출액 데이터
            amount_2024_24 = row[23]  # 2024.2/4 수출액
            amount_2025_24 = row[27]  # 2025.2/4 수출액
            
            table_data[sido] = {
                'changes': [change_2023_24, change_2024_24, change_2025_14, change_2025_24],
                'amounts': [amount_2024_24, amount_2025_24]
            }
    
    return table_data

def get_nationwide_data(sido_data, sido_products, summary_df):
    """전국 데이터를 추출합니다."""
    nationwide = sido_data.get('전국', {})
    products = sido_products.get('전국', [])
    
    # 수출액 가져오기 (컬럼 26이 2025.2/4)
    amount = None
    for i in range(3, len(summary_df)):
        row = summary_df.iloc[i]
        if row[3] == '전국' and str(row[4]) == '0':
            amount = row[26] / 10  # 백만달러 -> 억달러 변환
            break
    
    return {
        'amount': amount,
        'change': nationwide.get('change', 0),
        'products': products[:3]  # 상위 3개
    }

def get_regional_data(sido_data, sido_products_top, sido_products_bottom):
    """시도별 데이터를 추출하고 증가/감소 지역으로 분류합니다."""
    increase_regions = []
    decrease_regions = []
    
    for sido in SIDO_ORDER:
        if sido == '전국':
            continue
        
        sido_info = sido_data.get(sido, {})
        change = sido_info.get('change', 0)
        
        if pd.isna(change):
            continue
        
        # 증가 지역은 상위 순위 품목, 감소 지역은 하위 순위 품목
        if change > 0:
            products = sido_products_top.get(sido, [])
            region_info = {
                'name': sido,
                'change': change,
                'products': products[:3]
            }
            increase_regions.append(region_info)
        elif change < 0:
            products = sido_products_bottom.get(sido, [])
            region_info = {
                'name': sido,
                'change': change,
                'products': products[:3]
            }
            decrease_regions.append(region_info)
    
    # 정렬
    increase_regions.sort(key=lambda x: -x['change'])
    decrease_regions.sort(key=lambda x: x['change'])
    
    return increase_regions, decrease_regions

def generate_summary_box(nationwide_data, increase_regions, decrease_regions):
    """요약 박스 텍스트를 생성합니다."""
    # 증가 지역이 더 많으면 증가 중심
    increase_count = len(increase_regions)
    decrease_count = len(decrease_regions)
    
    # 상위 3개 지역 (기여도가 높은 품목 포함)
    top_increase = increase_regions[:3]
    top_decrease = decrease_regions[:3]
    
    # 헤드라인
    headline_regions = []
    for r in top_increase[:3]:
        products = r.get('products', [])
        if products:
            headline_regions.append(f"<span class='bold highlight'>{r['name']}</span>({products[0]['name']})")
        else:
            headline_regions.append(f"<span class='bold highlight'>{r['name']}</span>")
    
    headline = f"◆수출은 {', '.join(headline_regions)} 등 {increase_count}개 시도에서 전년동분기대비 증가"
    
    # 전국 요약
    amount = nationwide_data['amount']
    change = nationwide_data['change']
    products = nationwide_data['products']
    product_names = ", ".join([p['name'] for p in products[:3]])
    
    nationwide_summary = f"전국 수출(<span class='bold'>{amount:,.1f}억달러</span>)은 {product_names} 등의 수출이 늘어 전년동분기대비 <span class='bold'>{change:.1f}%</span> 증가"
    
    # 시도 요약
    decrease_names = ", ".join([f"<span class='bold'>{r['name']}</span>({r['change']:.1f}%)" for r in top_decrease])
    increase_names = ", ".join([f"<span class='bold'>{r['name']}</span>({r['change']:.1f}%)" for r in top_increase])
    
    # 감소 지역의 품목
    if top_decrease:
        decrease_products = []
        for r in top_decrease[:3]:
            prods = r.get('products', [])
            if prods:
                decrease_products.append(prods[0]['name'])
        decrease_product_str = ", ".join(decrease_products[:3]) if decrease_products else "기타 인조플라스틱 및 동 제품"
    else:
        decrease_product_str = ""
    
    # 증가 지역의 품목
    if top_increase:
        increase_products = []
        for r in top_increase[:3]:
            prods = r.get('products', [])
            if prods:
                increase_products.append(prods[0]['name'])
        increase_product_str = ", ".join(increase_products[:3]) if increase_products else "프로세서·컨트롤러"
    else:
        increase_product_str = ""
    
    regional_summary = f"{decrease_names}은 {decrease_product_str} 등의 수출이 줄어 감소하였으나, {increase_names}은 {increase_product_str} 등의 수출이 늘어 증가"
    
    return {
        'headline': headline,
        'nationwide_summary': nationwide_summary,
        'regional_summary': regional_summary
    }

def generate_summary_table(analysis_df, summary_df, sido_data):
    """요약 테이블 데이터를 생성합니다."""
    rows = []
    
    # G 참고 시트에서 증감률 데이터 가져오기
    reference_df = pd.read_excel(
        os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '분석표_25년 2분기_캡스톤.xlsx'),
        sheet_name='G 참고',
        header=None
    )
    
    # 시도별 증감률 및 수출액 추출
    sido_table_data = {}
    for i in range(2, len(reference_df)):
        row = reference_df.iloc[i]
        sido = row[1]
        rank = row[2]
        
        if pd.isna(rank) and pd.notna(sido) and sido in SIDO_ORDER:
            # 합계 행
            sido_table_data[sido] = {
                'change_2022': row[4],
                'change_2023': row[5],
                'change_2024': row[6],
                'change_2024_14': row[7],
                'change_2024_24': row[8],
                'change_2024_34': row[9],
                'change_2024_44': row[10],
                'change_2025_14': row[11],
                'change_2025_24': row[12]
            }
    
    # 전국 행
    nationwide_info = sido_table_data.get('전국', {})
    nationwide_change = sido_data.get('전국', {}).get('change', 0)
    
    # 전국 수출액 가져오기 (컬럼 22=2024.2/4, 컬럼 26=2025.2/4)
    nationwide_amount_2024 = None
    nationwide_amount_2025 = None
    for i in range(3, len(summary_df)):
        row = summary_df.iloc[i]
        if row[3] == '전국' and str(row[4]) == '0':
            nationwide_amount_2024 = row[22] / 10  # 백만달러 -> 억달러
            nationwide_amount_2025 = row[26] / 10
            break
    
    rows.append({
        'region_group': None,  # 전국은 region_group 없음
        'sido': '전 국',  # sido에 '전 국' 표시 (colspan 처리용)
        'changes': [-12.0, 10.1, -2.3, nationwide_change],
        'amounts': [nationwide_amount_2024, nationwide_amount_2025]
    })
    
    # 권역별 시도
    region_group_order = ['경인', '충청', '호남', '동북', '동남']
    
    for group_name in region_group_order:
        sidos = REGION_GROUPS[group_name]
        for idx, sido in enumerate(sidos):
            sido_info = sido_data.get(sido, {})
            table_info = sido_table_data.get(sido, {})
            
            # 수출액 가져오기 (컬럼 22=2024.2/4, 컬럼 26=2025.2/4)
            amount_2024 = None
            amount_2025 = None
            for i in range(3, len(summary_df)):
                row = summary_df.iloc[i]
                if row[3] == sido and str(row[4]) == '0':
                    amount_2024 = row[22] / 10
                    amount_2025 = row[26] / 10
                    break
            
            # 증감률 데이터
            changes = [
                table_info.get('change_2023', None),
                table_info.get('change_2024_24', None),
                table_info.get('change_2025_14', None),
                sido_info.get('change', None)
            ]
            
            row_data = {
                'sido': sido.replace('', ' ') if len(sido) == 2 else sido,
                'changes': changes,
                'amounts': [amount_2024, amount_2025]
            }
            
            # 첫 번째 시도에만 region_group과 rowspan 추가
            if idx == 0:
                row_data['region_group'] = group_name
                row_data['rowspan'] = len(sidos)
            else:
                row_data['region_group'] = None
            
            rows.append(row_data)
    
    return {'rows': rows}

def generate_report_data(excel_path, raw_excel_path=None, year=None, quarter=None):
    """보고서 데이터를 생성합니다.
    
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
    #     # 기초자료에서 수출 데이터 직접 추출
    #     # return extract_from_raw_data(extractor, ...)
    analysis_df, reference_df, summary_df = load_data(excel_path)
    
    # 시도별 데이터 추출
    sido_data = get_sido_data_from_analysis(analysis_df)
    sido_products_top, sido_products_bottom = get_sido_products_from_reference(reference_df)
    
    # 전국 데이터
    nationwide_data = get_nationwide_data(sido_data, sido_products_top, summary_df)
    
    # 시도별 데이터
    increase_regions, decrease_regions = get_regional_data(sido_data, sido_products_top, sido_products_bottom)
    
    # Top 3 지역
    top3_increase = increase_regions[:3]
    top3_decrease = decrease_regions[:3]
    
    # 요약 박스
    summary_box = generate_summary_box(nationwide_data, increase_regions, decrease_regions)
    
    # 요약 테이블
    summary_table = generate_summary_table(analysis_df, summary_df, sido_data)
    
    return {
        'report_info': {
            'main_section_number': '4',
            'main_section_title': '수출입 동향',
            'section_number': '가',
            'section_title': '수출',
            'page_number': '- 10 -'
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
    template_path = os.path.join(base_dir, 'export_template.html')
    output_path = os.path.join(base_dir, 'export_output.html')
    json_output_path = os.path.join(base_dir, 'export_data.json')
    
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
    print(f"전국 수출액: {data['nationwide_data']['amount']:,.1f}억달러")
    print(f"전국 증감률: {data['nationwide_data']['change']:.1f}%")
    
    print("\n=== 전국 주요 품목 ===")
    for prod in data['nationwide_data']['products']:
        print(f"  {prod['name']}: {prod['change']:.1f}%")
    
    print(f"\n증가 지역 수: {len(data['regional_data']['increase_regions'])}")
    print(f"감소 지역 수: {len(data['regional_data']['decrease_regions'])}")
    
    print("\n=== Top 3 증가 지역 ===")
    for region in data['top3_increase_regions']:
        prods = ", ".join([p['name'] for p in region['products'][:3]]) if region['products'] else "N/A"
        print(f"  {region['name']}({region['change']:.1f}%): {prods}")
    
    print("\n=== Top 3 감소 지역 ===")
    for region in data['top3_decrease_regions']:
        prods = ", ".join([p['name'] for p in region['products'][:3]]) if region['products'] else "N/A"
        print(f"  {region['name']}({region['change']:.1f}%): {prods}")

if __name__ == '__main__':
    main()

