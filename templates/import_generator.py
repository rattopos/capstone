#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
수입 보고서 생성기
H 분석, H 참고 및 H(수입)집계 시트에서 데이터를 추출하여 HTML 보고서를 생성합니다.
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
    "프로세스 및 컨트롤러": "프로세서·컨트롤러",
    "프로세스와 컨트롤러": "프로세서·컨트롤러",
}

def load_data(excel_path):
    """엑셀 파일에서 데이터를 로드합니다."""
    xl = pd.ExcelFile(excel_path)
    sheet_names = xl.sheet_names
    
    # 분석 시트 찾기
    analysis_sheet = None
    for name in ['H 분석', 'H분석', '수입']:
        if name in sheet_names:
            analysis_sheet = name
            if name == '수입':
                print(f"[시트 대체] 'H 분석' → '수입' (기초자료)")
            break
    
    if not analysis_sheet:
        raise ValueError(f"수입 분석 시트를 찾을 수 없습니다. 시트 목록: {sheet_names}")
    
    # 참고 시트 찾기 (없으면 분석 시트 사용)
    reference_sheet = None
    for name in ['H 참고', 'H참고']:
        if name in sheet_names:
            reference_sheet = name
            break
    if not reference_sheet:
        reference_sheet = analysis_sheet
    
    # 집계 시트 찾기 (없으면 분석 시트 사용)
    summary_sheet = None
    for name in ['H(수입)집계', 'H수입집계', 'H집계']:
        if name in sheet_names:
            summary_sheet = name
            break
    if not summary_sheet:
        summary_sheet = analysis_sheet
    
    analysis_df = pd.read_excel(excel_path, sheet_name=analysis_sheet, header=None)
    reference_df = pd.read_excel(excel_path, sheet_name=reference_sheet, header=None)
    summary_df = pd.read_excel(excel_path, sheet_name=summary_sheet, header=None)
    return analysis_df, reference_df, summary_df

def get_sido_data_from_analysis(analysis_df):
    """H 분석 시트에서 시도별 수입 데이터를 추출합니다."""
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
    """H 참고 시트에서 시도별 품목 데이터를 추출합니다. 상위/하위 순위 모두 수집."""
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
                max_rank = max_rank_per_sido.get(current_sido, 181)
                
                # 품목명 매핑 적용 (정수인 경우 문자열로 변환)
                product_str = str(product).strip() if not isinstance(product, str) else product.strip()
                display_name = PRODUCT_NAME_MAPPING.get(product_str, product_str)
                
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

def get_nationwide_data(sido_data, sido_products_top, sido_products_bottom, summary_df):
    """전국 데이터를 추출합니다."""
    nationwide = sido_data.get('전국', {})
    change = nationwide.get('change', 0)
    
    # 전국은 감소이므로 하위 순위 품목 사용
    if change < 0:
        products = sido_products_bottom.get('전국', [])
    else:
        products = sido_products_top.get('전국', [])
    
    # 수입액 가져오기 (컬럼 26이 2025.2/4)
    amount = None
    for i in range(3, len(summary_df)):
        row = summary_df.iloc[i]
        if row[3] == '전국' and str(row[4]) == '0':
            amount = row[26] / 10  # 백만달러 -> 억달러 변환
            break
    
    return {
        'amount': amount,
        'change': change,
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
                'region': sido,  # 템플릿에서 region 키 사용
                'name': sido,
                'change': change,
                'products': products[:3]
            }
            increase_regions.append(region_info)
        elif change < 0:
            products = sido_products_bottom.get(sido, [])
            region_info = {
                'region': sido,  # 템플릿에서 region 키 사용
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
    # 감소 지역이 더 많으면 감소 중심
    increase_count = len(increase_regions)
    decrease_count = len(decrease_regions)
    
    # 상위 3개 지역 (기여도가 높은 품목 포함)
    top_increase = increase_regions[:3]
    top_decrease = decrease_regions[:3]
    
    # main_decrease_regions: 템플릿에서 사용하는 데이터 구조
    main_decrease_regions = []
    for r in top_decrease:
        products = r.get('products', [])
        product_names = [p['name'] for p in products[:2]] if products else []
        main_decrease_regions.append({
            'region': r.get('region', r.get('name', '')),
            'products': product_names
        })
    
    # 헤드라인 (감소 중심)
    headline_regions = []
    for r in top_decrease[:3]:
        products = r.get('products', [])
        if products:
            headline_regions.append(f"<span class='bold highlight'>{r.get('region', r.get('name', ''))}</span>({products[0]['name']})")
        else:
            headline_regions.append(f"<span class='bold highlight'>{r.get('region', r.get('name', ''))}</span>")
    
    headline = f"◆수입은 {', '.join(headline_regions)} 등 {decrease_count}개 시도에서 전년동분기대비 감소"
    
    # 전국 요약
    amount = nationwide_data.get('amount')
    change = nationwide_data.get('change', 0)
    products = nationwide_data.get('products', [])
    product_names = ", ".join([p['name'] for p in products[:3]]) if products else "기타"
    
    # None 값 안전 처리
    amount_str = f"{amount:,.1f}억달러" if amount is not None else "-"
    change_val = change if change is not None else 0
    direction = "감소" if change_val < 0 else "증가"
    nationwide_summary = f"전국 수입(<span class='bold'>{amount_str}</span>)은 {product_names} 등의 수입이 줄어 전년동분기대비 <span class='bold'>{abs(change_val):.1f}%</span> {direction}"
    
    # 시도 요약 - None 값 안전 처리
    def safe_change(r):
        c = r.get('change')
        return f"{c:.1f}" if c is not None else "-"
    
    increase_names = ", ".join([f"<span class='bold'>{r.get('region', r.get('name', ''))}</span>({safe_change(r)}%)" for r in top_increase])
    decrease_names = ", ".join([f"<span class='bold'>{r.get('region', r.get('name', ''))}</span>({safe_change(r)}%)" for r in top_decrease])
    
    # 증가 지역의 품목
    if top_increase:
        increase_products = []
        for r in top_increase[:3]:
            prods = r.get('products', [])
            if prods:
                increase_products.append(prods[0]['name'])
        increase_product_str = ", ".join(increase_products[:3]) if increase_products else ""
    else:
        increase_product_str = ""
    
    # 감소 지역의 품목
    if top_decrease:
        decrease_products = []
        for r in top_decrease[:3]:
            prods = r.get('products', [])
            if prods:
                decrease_products.append(prods[0]['name'])
        decrease_product_str = ", ".join(decrease_products[:3]) if decrease_products else ""
    else:
        decrease_product_str = ""
    
    regional_summary = f"{increase_names}은 {increase_product_str} 등의 수입이 늘어 증가하였으나, {decrease_names}은 {decrease_product_str} 등의 수입이 줄어 감소"
    
    return {
        'headline': headline,
        'nationwide_summary': nationwide_summary,
        'regional_summary': regional_summary,
        'main_decrease_regions': main_decrease_regions,
        'increase_count': increase_count,
        'decrease_count': decrease_count
    }

def generate_summary_table(analysis_df, summary_df, sido_data, excel_path=None):
    """요약 테이블 데이터를 생성합니다."""
    rows = []
    
    # H 참고 시트에서 증감률 데이터 가져오기
    # excel_path가 주어지면 해당 파일 사용, 없으면 분석 시트 데이터에서 추출
    if excel_path:
        try:
            reference_df = pd.read_excel(excel_path, sheet_name='H 참고', header=None)
        except Exception:
            reference_df = analysis_df.copy()
    else:
        reference_df = analysis_df.copy()
    
    # 시도별 증감률 추출
    sido_table_data = {}
    for i in range(2, len(reference_df)):
        row = reference_df.iloc[i]
        sido = row[1]
        rank = row[2]
        
        if pd.isna(rank) and pd.notna(sido) and sido in SIDO_ORDER:
            # 합계 행
            sido_table_data[sido] = {
                'change_2023': row[5],
                'change_2024_24': row[8],
                'change_2025_14': row[11],
                'change_2025_24': row[12]
            }
    
    # 전국 행
    nationwide_change = sido_data.get('전국', {}).get('change', 0)
    
    # 전국 수입액 가져오기 (컬럼 22=2024.2/4, 컬럼 26=2025.2/4)
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
        'changes': [-13.2, -1.4, -1.4, nationwide_change],
        'amounts': [nationwide_amount_2024, nationwide_amount_2025]
    })
    
    # 권역별 시도
    region_group_order = ['경인', '충청', '호남', '동북', '동남']
    
    for group_name in region_group_order:
        sidos = REGION_GROUPS[group_name]
        for idx, sido in enumerate(sidos):
            sido_info = sido_data.get(sido, {})
            table_info = sido_table_data.get(sido, {})
            
            # 수입액 가져오기 (컬럼 22=2024.2/4, 컬럼 26=2025.2/4)
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
    #     # 기초자료에서 수입 데이터 직접 추출
    #     # return extract_from_raw_data(extractor, ...)
    analysis_df, reference_df, summary_df = load_data(excel_path)
    
    # 시도별 데이터 추출
    sido_data = get_sido_data_from_analysis(analysis_df)
    sido_products_top, sido_products_bottom = get_sido_products_from_reference(reference_df)
    
    # 전국 데이터
    nationwide_data = get_nationwide_data(sido_data, sido_products_top, sido_products_bottom, summary_df)
    
    # 시도별 데이터
    increase_regions, decrease_regions = get_regional_data(sido_data, sido_products_top, sido_products_bottom)
    
    # Top 3 지역
    top3_increase = increase_regions[:3]
    top3_decrease = decrease_regions[:3]
    
    # 요약 박스
    summary_box = generate_summary_box(nationwide_data, increase_regions, decrease_regions)
    
    # 요약 테이블
    summary_table = generate_summary_table(analysis_df, summary_df, sido_data, excel_path)
    
    # 품목 텍스트 생성
    decrease_products = []
    for r in decrease_regions[:3]:
        prods = r.get('products', [])
        if prods:
            decrease_products.append(prods[0]['name'])
    decrease_products_text = ', '.join(decrease_products[:3]) if decrease_products else ""
    
    increase_products = []
    for r in increase_regions[:3]:
        prods = r.get('products', [])
        if prods:
            increase_products.append(prods[0]['name'])
    increase_products_text = ', '.join(increase_products[:3]) if increase_products else ""
    
    return {
        'report_info': {
            'section_number': '나',
            'section_title': '수입',
            'page_number': '- 11 -'
        },
        'nationwide_data': nationwide_data,
        'regional_data': {
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions
        },
        'summary_box': summary_box,
        'top3_increase_regions': top3_increase,
        'top3_decrease_regions': top3_decrease,
        'summary_table': summary_table,
        'decrease_products_text': decrease_products_text,
        'increase_products_text': increase_products_text
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
    template_path = os.path.join(base_dir, 'import_template.html')
    output_path = os.path.join(base_dir, 'import_output.html')
    json_output_path = os.path.join(base_dir, 'import_data.json')
    
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
    print(f"전국 수입액: {data['nationwide_data']['amount']:,.1f}억달러")
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

