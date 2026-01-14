#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
수출 보도자료 생성기
G 분석, G 참고 및 G(수출)집계 시트에서 데이터를 추출하여 HTML 보도자료를 생성합니다.
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
    xl = pd.ExcelFile(excel_path)
    sheet_names = xl.sheet_names
    
    # 분석 시트 찾기
    analysis_sheet = None
    use_raw = False
    for name in ['G 분석', 'G분석', '수출']:
        if name in sheet_names:
            analysis_sheet = name
            if name == '수출':
                print(f"[시트 대체] 'G 분석' → '수출' (기초자료)")
                use_raw = True
            break
    
    if not analysis_sheet:
        raise ValueError(f"수출 분석 시트를 찾을 수 없습니다. 시트 목록: {sheet_names}")
    
    analysis_df = pd.read_excel(excel_path, sheet_name=analysis_sheet, header=None)
    
    # 분석 시트에 실제 데이터가 있는지 확인 (수식 미계산 체크)
    test_row = analysis_df[(analysis_df[3].isin(SIDO_ORDER)) & (analysis_df[8] == '합계')]
    if test_row.empty or test_row.iloc[0].isna().sum() > 20:
        print(f"[수출] 분석 시트가 비어있음 → 집계 시트에서 직접 계산")
        use_aggregation_only = True
    else:
        use_aggregation_only = False
    
    # 참고 시트 찾기 (없으면 분석 시트 사용)
    reference_sheet = None
    for name in ['G 참고', 'G참고']:
        if name in sheet_names:
            reference_sheet = name
            break
    if not reference_sheet:
        reference_sheet = analysis_sheet
    
    # 집계 시트 찾기 (없으면 분석 시트 사용)
    summary_sheet = None
    for name in ['G(수출)집계', 'G수출집계', 'G집계']:
        if name in sheet_names:
            summary_sheet = name
            break
    if not summary_sheet:
        summary_sheet = analysis_sheet
    
    reference_df = pd.read_excel(excel_path, sheet_name=reference_sheet, header=None)
    summary_df = pd.read_excel(excel_path, sheet_name=summary_sheet, header=None)
    
    # use_raw, use_aggregation_only 정보를 데이터프레임에 속성으로 저장
    analysis_df.attrs['use_raw'] = use_raw
    analysis_df.attrs['use_aggregation_only'] = use_aggregation_only
    
    return analysis_df, reference_df, summary_df

def get_sido_data_from_analysis(analysis_df, summary_df=None):
    """G 분석 시트에서 시도별 수출 데이터를 추출합니다."""
    use_aggregation_only = analysis_df.attrs.get('use_aggregation_only', False)
    
    # 집계 시트 기반으로 추출 (분석 시트가 비어있는 경우 포함)
    if use_aggregation_only and summary_df is not None:
        return _get_sido_data_from_aggregation(summary_df)
    
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


def _get_sido_data_from_aggregation(summary_df):
    """집계 시트에서 시도별 수출 데이터 추출 (증감률 직접 계산)"""
    # 집계 시트 구조: 3=지역이름, 4=분류단계, 7=상품이름
    # 데이터 컬럼: 22=2024.2/4, 26=2025.2/4
    sido_data = {}
    
    for i in range(3, len(summary_df)):
        row = summary_df.iloc[i]
        sido = row[3]
        level = str(row[4]).strip() if pd.notna(row[4]) else ''
        
        if level == '0' and sido in SIDO_ORDER:
            # 당분기(2025.2/4)와 전년동분기(2024.2/4) 수출액으로 증감률 계산
            current = safe_float(row[26], 0)  # 2025.2/4
            prev = safe_float(row[22], 0)  # 2024.2/4
            
            if prev and prev != 0:
                change = ((current - prev) / prev) * 100
            else:
                change = 0.0
            
            sido_data[sido] = {
                'change': change,
                'row_idx': i
            }
    
    return sido_data


def safe_float(value, default=None):
    """안전한 float 변환 함수"""
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

def get_sido_products_from_reference(reference_df, summary_df=None, use_aggregation_only=False):
    """G 참고 시트에서 시도별 품목 데이터를 추출합니다. 상위/하위 순위 모두 수집."""
    # 분석 시트가 비어있으면 집계 시트에서 품목 데이터 추출
    if use_aggregation_only and summary_df is not None:
        return _get_sido_products_from_aggregation(summary_df)
    
    sido_products = {}  # 상위 순위 (증가 품목)
    sido_products_bottom = {}  # 하위 순위 (감소 품목)
    current_sido = None
    max_rank_per_sido = {}
    
    # 참고 시트에 데이터가 있는지 확인
    test_data = reference_df.iloc[3:10, 1:5].dropna(how='all')
    if len(test_data) == 0 or reference_df.iloc[3:10, 12].isna().all():
        # 참고 시트가 비어있으면 집계 시트 사용
        if summary_df is not None:
            return _get_sido_products_from_aggregation(summary_df)
    
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


def _get_sido_products_from_aggregation(summary_df):
    """집계 시트에서 시도별 품목 데이터를 추출합니다 (기여도 기준 순위)."""
    sido_products = {}  # 상위 순위 (증가 기여 품목)
    sido_products_bottom = {}  # 하위 순위 (감소 기여 품목)
    
    # 집계 시트 구조: 3=지역이름, 4=분류단계, 7=상품이름
    # 데이터 컬럼: 22=2024.2/4, 26=2025.2/4
    
    for sido in SIDO_ORDER:
        sido_products[sido] = []
        sido_products_bottom[sido] = []
        
        # 해당 시도의 소분류(2) 데이터만 추출
        sido_items = summary_df[(summary_df[3] == sido) & (summary_df[4].astype(str) == '2')]
        
        if len(sido_items) == 0:
            continue
        
        # 기여도(금액 변화량) 계산
        items_data = []
        for _, row in sido_items.iterrows():
            product = row[7]
            if pd.isna(product) or product == '합계':
                continue
            
            curr = safe_float(row[26], 0)  # 2025.2/4
            prev = safe_float(row[22], 0)  # 2024.2/4
            
            # 증감률 계산
            if prev and prev != 0:
                change = ((curr - prev) / prev) * 100
            else:
                change = 0.0
            
            # 기여도 = 금액 변화량 (절대값으로 순위 결정)
            contribution = curr - prev
            
            product_str = str(product).strip()
            display_name = PRODUCT_NAME_MAPPING.get(product_str, product_str)
            
            items_data.append({
                'name': display_name,
                'change': round(change, 1),
                'contribution': contribution
            })
        
        # 기여도 양수 (증가에 기여한 품목) - 내림차순
        positive_items = sorted([i for i in items_data if i['contribution'] > 0], 
                               key=lambda x: -x['contribution'])
        for i, item in enumerate(positive_items[:5]):
            item['rank'] = i + 1
            sido_products[sido].append(item)
        
        # 기여도 음수 (감소에 기여한 품목) - 오름차순 (가장 큰 감소가 먼저)
        negative_items = sorted([i for i in items_data if i['contribution'] < 0], 
                               key=lambda x: x['contribution'])
        for i, item in enumerate(negative_items[:5]):
            item['rank'] = i + 1
            sido_products_bottom[sido].append(item)
    
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
    # 증가 지역이 더 많으면 증가 중심
    increase_count = len(increase_regions)
    decrease_count = len(decrease_regions)
    
    # 상위 3개 지역 (기여도가 높은 품목 포함)
    top_increase = increase_regions[:3]
    top_decrease = decrease_regions[:3]
    
    # main_increase_regions: 템플릿에서 사용하는 데이터 구조
    main_increase_regions = []
    for r in top_increase:
        products = r.get('products', [])
        product_names = [p['name'] for p in products[:2]] if products else []
        main_increase_regions.append({
            'region': r.get('region', r.get('name', '')),
            'products': product_names
        })
    
    # 헤드라인
    headline_regions = []
    for r in top_increase[:3]:
        products = r.get('products', [])
        if products:
            headline_regions.append(f"<span class='bold highlight'>{r.get('region', r.get('name', ''))}</span>({products[0]['name']})")
        else:
            headline_regions.append(f"<span class='bold highlight'>{r.get('region', r.get('name', ''))}</span>")
    
    headline = f"◆수출은 {', '.join(headline_regions)} 등 {increase_count}개 시도에서 전년동분기대비 증가"
    
    # 전국 요약
    amount = nationwide_data.get('amount')
    change = nationwide_data.get('change', 0)
    products = nationwide_data.get('products', [])
    product_names = ", ".join([p['name'] for p in products[:3]]) if products else "기타"
    
    # None 값 안전 처리
    amount_str = f"{amount:,.1f}억달러" if amount is not None else "-"
    change_val = change if change is not None else 0
    direction = "증가" if change_val >= 0 else "감소"
    verb = "늘어" if change_val >= 0 else "줄어"
    
    nationwide_summary = f"전국 수출(<span class='bold'>{amount_str}</span>)은 {product_names} 등의 수출이 {verb} 전년동분기대비 <span class='bold'>{abs(change_val):.1f}%</span> {direction}"
    
    # 시도 요약 - None 값 안전 처리
    def safe_change(r):
        c = r.get('change')
        return f"{c:.1f}" if c is not None else "-"
    
    decrease_names = ", ".join([f"<span class='bold'>{r.get('region', r.get('name', ''))}</span>({safe_change(r)}%)" for r in top_decrease])
    increase_names = ", ".join([f"<span class='bold'>{r.get('region', r.get('name', ''))}</span>({safe_change(r)}%)" for r in top_increase])
    
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
    
    # 전국 방향에 따라 regional_summary 생성
    if change_val >= 0:
        # 전국이 증가일 때: "감소하였으나...증가"
        regional_summary = f"{decrease_names}은 {decrease_product_str} 등의 수출이 줄었으나, {increase_names}은 {increase_product_str} 등의 수출이 늘어 증가"
    else:
        # 전국이 감소일 때: "증가하였으나...감소"
        regional_summary = f"{increase_names}은 {increase_product_str} 등의 수출이 늘었으나, {decrease_names}은 {decrease_product_str} 등의 수출이 줄어 감소"
    
    return {
        'headline': headline,
        'nationwide_summary': nationwide_summary,
        'regional_summary': regional_summary,
        'main_increase_regions': main_increase_regions,
        'increase_count': increase_count,
        'decrease_count': decrease_count
    }

def generate_summary_table(analysis_df, summary_df, sido_data, excel_path=None):
    """요약 테이블 데이터를 생성합니다."""
    use_aggregation_only = analysis_df.attrs.get('use_aggregation_only', False)
    
    # 분석 시트가 비어있으면 집계 시트에서 직접 계산
    if use_aggregation_only:
        return _generate_summary_table_from_aggregation(summary_df, sido_data)
    
    rows = []
    
    # G 참고 시트에서 증감률 데이터 가져오기
    # excel_path가 주어지면 해당 파일 사용, 없으면 분석 시트 데이터에서 추출
    if excel_path:
        try:
            reference_df = pd.read_excel(excel_path, sheet_name='G 참고', header=None)
        except Exception:
            reference_df = analysis_df.copy()
    else:
        reference_df = analysis_df.copy()
    
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
                'sido': ' '.join(sido) if len(sido) == 2 else sido,
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


def _generate_summary_table_from_aggregation(summary_df, sido_data):
    """집계 시트에서 테이블 데이터 추출"""
    # 집계 시트 구조: 3=지역이름, 4=분류단계
    # 데이터 컬럼: 14=2022.2/4, 18=2023.2/4, 21=2024.1/4, 22=2024.2/4, 25=2025.1/4, 26=2025.2/4
    
    rows = []
    region_display = {
        '전국': '전 국', '서울': '서 울', '부산': '부 산', '대구': '대 구', '인천': '인 천',
        '광주': '광 주', '대전': '대 전', '울산': '울 산', '세종': '세 종', '경기': '경 기',
        '강원': '강 원', '충북': '충 북', '충남': '충 남', '전북': '전 북', '전남': '전 남',
        '경북': '경 북', '경남': '경 남', '제주': '제 주'
    }
    
    all_regions = ['전국', '서울', '인천', '경기', '대전', '세종', '충북', '충남', 
                   '광주', '전북', '전남', '제주', '대구', '경북', '강원', '부산', '울산', '경남']
    
    sido_table_data = {}
    for region in all_regions:
        region_total = summary_df[(summary_df[3] == region) & (summary_df[4].astype(str) == '0')]
        if region_total.empty:
            continue
        
        row = region_total.iloc[0]
        
        # 수출액 (백만달러 -> 억달러)
        amount_2024 = safe_float(row[22], 0) / 10
        amount_2025 = safe_float(row[26], 0) / 10
        
        # 지수 값
        idx_2022_2 = safe_float(row[14], 0)
        idx_2023_2 = safe_float(row[18], 0)
        idx_2024_1 = safe_float(row[21], 0)
        idx_2024_2 = safe_float(row[22], 0)
        idx_2025_1 = safe_float(row[25], 0)
        idx_2025_2 = safe_float(row[26], 0)
        
        # 전년동분기비 증감률 계산
        change_2023_2 = ((idx_2023_2 - idx_2022_2) / idx_2022_2 * 100) if idx_2022_2 and idx_2022_2 != 0 else 0.0
        change_2024_2 = ((idx_2024_2 - idx_2023_2) / idx_2023_2 * 100) if idx_2023_2 and idx_2023_2 != 0 else 0.0
        change_2025_1 = ((idx_2025_1 - idx_2024_1) / idx_2024_1 * 100) if idx_2024_1 and idx_2024_1 != 0 else 0.0
        change_2025_2 = ((idx_2025_2 - idx_2024_2) / idx_2024_2 * 100) if idx_2024_2 and idx_2024_2 != 0 else 0.0
        
        sido_table_data[region] = {
            'changes': [round(change_2023_2, 1), round(change_2024_2, 1), round(change_2025_1, 1), round(change_2025_2, 1)],
            'amounts': [round(amount_2024, 1), round(amount_2025, 1)]
        }
    
    # 전국 먼저
    if '전국' in sido_table_data:
        rows.append({
            'region_group': None,
            'sido': '전 국',
            'changes': sido_table_data['전국']['changes'],
            'amounts': sido_table_data['전국']['amounts']
        })
    
    # 권역별 시도
    region_group_order = ['경인', '충청', '호남', '동북', '동남']
    
    for group_name in region_group_order:
        sidos = REGION_GROUPS[group_name]
        for idx, sido in enumerate(sidos):
            if sido not in sido_table_data:
                continue
            
            row_data = {
                'sido': region_display.get(sido, sido),
                'changes': sido_table_data[sido]['changes'],
                'amounts': sido_table_data[sido]['amounts']
            }
            
            if idx == 0:
                row_data['region_group'] = group_name
                row_data['rowspan'] = len(sidos)
            else:
                row_data['region_group'] = None
            
            rows.append(row_data)
    
    return {'rows': rows}

def generate_report_data(excel_path, raw_excel_path=None, year=None, quarter=None):
    """보도자료 데이터를 생성합니다.
    
    ★ 핵심 원칙: 모든 나레이션 데이터는 테이블(summary_table)에서 가져옴
    - 테이블을 먼저 생성하고, 나레이션은 테이블 데이터를 참조
    - 데이터 불일치 원천 차단
    
    Args:
        excel_path: 분석표 엑셀 파일 경로
        raw_excel_path: 기초자료 엑셀 파일 경로 (선택사항)
        year: 현재 연도 (선택사항)
        quarter: 현재 분기 (선택사항)
    """
    analysis_df, reference_df, summary_df = load_data(excel_path)
    use_aggregation_only = analysis_df.attrs.get('use_aggregation_only', False)
    
    # ★ 1단계: 테이블 데이터를 먼저 생성 (이것이 Single Source of Truth)
    summary_table = _generate_summary_table_from_aggregation(summary_df, {})
    
    # ★ 2단계: 테이블에서 시도별 증감률 추출 (나레이션용)
    sido_data = {}
    table_change_map = {}  # 시도명 -> 증감률 매핑
    
    for row in summary_table['rows']:
        sido_raw = row['sido'].replace(' ', '')  # '서 울' -> '서울'
        if not sido_raw:
            # region_group이 '전 국'인 경우
            if row.get('region_group') == '전 국':
                sido_raw = '전국'
            else:
                continue
        
        # 테이블의 2025.2/4 증감률 (changes[3])
        change_value = row['changes'][3] if len(row['changes']) > 3 else 0
        table_change_map[sido_raw] = change_value
        
        sido_data[sido_raw] = {
            'change': change_value,
            'amount': row['amounts'][1] if len(row['amounts']) > 1 else 0  # 2025.2/4 수출액
        }
    
    # ★ 3단계: 품목 데이터 추출
    sido_products_top, sido_products_bottom = get_sido_products_from_reference(
        reference_df, summary_df, use_aggregation_only)
    
    # ★ 4단계: 테이블 증감률 기준으로 증가/감소 지역 분류
    increase_regions = []
    decrease_regions = []
    
    for sido in SIDO_ORDER:
        if sido == '전국':
            continue
        
        change = table_change_map.get(sido, 0)
        if change is None or pd.isna(change):
            continue
        
        if change > 0:
            products = sido_products_top.get(sido, [])
            increase_regions.append({
                'region': sido,
                'name': sido,
                'change': change,  # 테이블에서 가져온 값
                'products': products[:3]
            })
        elif change < 0:
            products = sido_products_bottom.get(sido, [])
            decrease_regions.append({
                'region': sido,
                'name': sido,
                'change': change,  # 테이블에서 가져온 값
                'products': products[:3]
            })
    
    # 정렬
    increase_regions.sort(key=lambda x: -x['change'])
    decrease_regions.sort(key=lambda x: x['change'])
    
    # Top 3 지역
    top3_increase = increase_regions[:3]
    top3_decrease = decrease_regions[:3]
    
    # ★ 5단계: 전국 데이터 (테이블에서 가져옴)
    nationwide_change = table_change_map.get('전국', 0)
    nationwide_amount = sido_data.get('전국', {}).get('amount', 0)
    nationwide_products = sido_products_top.get('전국', [])[:3]
    
    nationwide_data = {
        'amount': nationwide_amount,
        'change': nationwide_change,  # 테이블에서 가져온 값
        'products': nationwide_products
    }
    
    # ★ 6단계: 요약 박스 (테이블 데이터 기반)
    summary_box = _generate_summary_box_from_table(
        nationwide_data, increase_regions, decrease_regions)
    
    # 품목 텍스트 생성
    decrease_products = [r['products'][0]['name'] for r in decrease_regions[:3] if r.get('products')]
    increase_products = [r['products'][0]['name'] for r in increase_regions[:3] if r.get('products')]
    
    return {
        'report_info': {
            'year': year,
            'quarter': quarter,
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
        'summary_table': summary_table,
        'decrease_products_text': ', '.join(decrease_products),
        'increase_products_text': ', '.join(increase_products)
    }


def _generate_summary_box_from_table(nationwide_data, increase_regions, decrease_regions):
    """테이블 데이터 기반으로 요약 박스 생성"""
    increase_count = len(increase_regions)
    decrease_count = len(decrease_regions)
    
    top_increase = increase_regions[:3]
    top_decrease = decrease_regions[:3]
    
    # main_increase_regions
    main_increase_regions = []
    for r in top_increase:
        products = r.get('products', [])
        product_names = [p['name'] for p in products[:2]] if products else []
        main_increase_regions.append({
            'region': r['name'],
            'products': product_names
        })
    
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
    amount = nationwide_data.get('amount', 0)
    change = nationwide_data.get('change', 0)
    products = nationwide_data.get('products', [])
    product_names = ", ".join([p['name'] for p in products[:3]]) if products else "기타"
    
    amount_str = f"{amount:,.1f}억달러" if amount else "-"
    direction = "증가" if change >= 0 else "감소"
    verb = "늘어" if change >= 0 else "줄어"
    
    nationwide_summary = f"전국 수출(<span class='bold'>{amount_str}</span>)은 {product_names} 등의 수출이 {verb} 전년동분기대비 <span class='bold'>{abs(change):.1f}%</span> {direction}"
    
    # 시도 요약
    def safe_change(r):
        c = r.get('change', 0)
        return f"{c:.1f}" if c is not None else "-"
    
    decrease_names = ", ".join([f"<span class='bold'>{r['name']}</span>({safe_change(r)}%)" for r in top_decrease])
    increase_names = ", ".join([f"<span class='bold'>{r['name']}</span>({safe_change(r)}%)" for r in top_increase])
    
    # 품목
    decrease_products = [r['products'][0]['name'] for r in top_decrease if r.get('products')]
    increase_products = [r['products'][0]['name'] for r in top_increase if r.get('products')]
    
    decrease_product_str = ", ".join(decrease_products[:3]) if decrease_products else ""
    increase_product_str = ", ".join(increase_products[:3]) if increase_products else ""
    
    if change >= 0:
        regional_summary = f"{decrease_names}은 {decrease_product_str} 등의 수출이 줄어 감소하였으나, {increase_names}은 {increase_product_str} 등의 수출이 늘어 증가"
    else:
        regional_summary = f"{increase_names}은 {increase_product_str} 등의 수출이 늘었으나, {decrease_names}은 {decrease_product_str} 등의 수출이 줄어 감소"
    
    return {
        'headline': headline,
        'nationwide_summary': nationwide_summary,
        'regional_summary': regional_summary,
        'main_increase_regions': main_increase_regions,
        'increase_count': increase_count,
        'decrease_count': decrease_count
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

