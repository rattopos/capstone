# -*- coding: utf-8 -*-
"""
요약 보도자료 데이터 추출 서비스
"""

import pandas as pd


# =========================================================
# 동적 컬럼 인덱스 계산 헬퍼 함수들
# =========================================================

# 기초자료 시트별 기준 컬럼 설정 (2025년 2분기 기준)
# key: (시트명, 기준연도, 기준분기) -> 컬럼 인덱스
RAW_SHEET_BASE_COLS = {
    '광공업생산': {'base_year': 2025, 'base_quarter': 2, 'base_col': 64, 'region_col': 1, 'code_col': 4, 'total_code': 'BCD'},
    '서비스업생산': {'base_year': 2025, 'base_quarter': 2, 'base_col': 64, 'region_col': 1, 'code_col': 4, 'total_code': 'E~S'},
    '소비(소매, 추가)': {'base_year': 2025, 'base_quarter': 2, 'base_col': 63, 'region_col': 1, 'code_col': 4, 'total_code': '총지수'},
    '수출': {'base_year': 2025, 'base_quarter': 2, 'base_col': 68, 'region_col': 1, 'code_col': 5, 'total_code': '합계'},
    '수입': {'base_year': 2025, 'base_quarter': 2, 'base_col': 68, 'region_col': 1, 'code_col': 5, 'total_code': '합계'},
    '품목성질별 물가': {'base_year': 2025, 'base_quarter': 2, 'base_col': 56, 'region_col': 0, 'code_col': 3, 'total_code': '총지수'},
    '연령별고용률': {'base_year': 2025, 'base_quarter': 2, 'base_col': 66, 'region_col': 1, 'code_col': 3, 'total_code': '계'},
    '실업자 수': {'base_year': 2025, 'base_quarter': 2, 'base_col': 61, 'region_col': 0, 'code_col': 1, 'total_code': '계'},
    '시도 간 이동': {'base_year': 2025, 'base_quarter': 2, 'base_col': 80, 'region_col': 1, 'code_col': 1, 'total_code': None},
    # 건설 시트 - 열 구조가 다름 (2024 3Q = col64)
    '건설 (공표자료)': {'base_year': 2024, 'base_quarter': 3, 'base_col': 64, 'region_col': 1, 'code_col': 3, 'total_code': 0, 'name_col': 4},
}

# 집계 시트별 기준 컬럼 설정 (2025년 2분기 기준)
AGGREGATE_SHEET_BASE_COLS = {
    'A(광공업생산)집계': {'base_year': 2025, 'base_quarter': 2, 'base_col': 26, 'region_col': 4, 'code_col': 7, 'total_code': 'BCD'},
    'B(서비스업생산)집계': {'base_year': 2025, 'base_quarter': 2, 'base_col': 25, 'region_col': 3, 'code_col': 6, 'total_code': 'E~S'},
    'C(소비)집계': {'base_year': 2025, 'base_quarter': 2, 'base_col': 24, 'region_col': 2, 'code_col': 6, 'total_code': '총지수'},
    'G(수출)집계': {'base_year': 2025, 'base_quarter': 2, 'base_col': 26, 'region_col': 3, 'code_col': 4, 'total_code': '0'},
    'E(품목성질물가)집계': {'base_year': 2025, 'base_quarter': 2, 'base_col': 21, 'region_col': 0, 'code_col': 3, 'total_code': '총지수'},
    'D(고용률)집계': {'base_year': 2025, 'base_quarter': 2, 'base_col': 21, 'region_col': 1, 'code_col': 3, 'total_code': '계'},
}


def get_dynamic_col(sheet_name: str, target_year: int, target_quarter: int, is_aggregate: bool = False) -> int:
    """
    시트별 동적 컬럼 인덱스 계산
    
    Args:
        sheet_name: 시트 이름
        target_year: 대상 연도
        target_quarter: 대상 분기 (1-4)
        is_aggregate: 집계 시트 여부
        
    Returns:
        컬럼 인덱스
    """
    base_config = AGGREGATE_SHEET_BASE_COLS.get(sheet_name) if is_aggregate else RAW_SHEET_BASE_COLS.get(sheet_name)
    
    if not base_config:
        print(f"[경고] 알 수 없는 시트: {sheet_name}")
        return None
    
    base_year = base_config['base_year']
    base_quarter = base_config['base_quarter']
    base_col = base_config['base_col']
    
    # 분기 차이 계산 (분기 단위로)
    quarters_diff = (target_year - base_year) * 4 + (target_quarter - base_quarter)
    
    return base_col + quarters_diff


def get_dynamic_raw_sheet_config(sheet_name: str, year: int, quarter: int) -> dict:
    """
    기초자료 시트의 동적 설정 생성
    
    Args:
        sheet_name: 원본 시트 이름 (예: 'A 분석')
        year: 대상 연도
        quarter: 대상 분기
        
    Returns:
        동적으로 계산된 설정 딕셔너리
    """
    # 시트명 매핑
    sheet_mapping = {
        'A 분석': '광공업생산',
        'B 분석': '서비스업생산',
        'C 분석': '소비(소매, 추가)',
        'G 분석': '수출',
        'E(품목성질물가)분석': '품목성질별 물가',
        'D(고용률)분석': '연령별고용률',
    }
    
    raw_sheet = sheet_mapping.get(sheet_name)
    if not raw_sheet:
        return None
    
    base_config = RAW_SHEET_BASE_COLS.get(raw_sheet)
    if not base_config:
        return None
    
    curr_col = get_dynamic_col(raw_sheet, year, quarter, is_aggregate=False)
    prev_col = get_dynamic_col(raw_sheet, year - 1, quarter, is_aggregate=False)
    
    config = {
        'raw_sheet': raw_sheet,
        'region_col': base_config['region_col'],
        'code_col': base_config['code_col'],
        'total_code': base_config['total_code'],
        'curr_col': curr_col,
        'prev_col': prev_col,
    }
    
    # 고용률은 차이 계산
    if sheet_name == 'D(고용률)분석':
        config['calc_type'] = 'difference'
    
    return config


def get_dynamic_aggregate_sheet_config(sheet_name: str, year: int, quarter: int) -> dict:
    """
    집계 시트의 동적 설정 생성
    
    Args:
        sheet_name: 원본 시트 이름 (예: 'A 분석')
        year: 대상 연도
        quarter: 대상 분기
        
    Returns:
        동적으로 계산된 설정 딕셔너리
    """
    # 시트명 매핑
    sheet_mapping = {
        'A 분석': 'A(광공업생산)집계',
        'B 분석': 'B(서비스업생산)집계',
        'C 분석': 'C(소비)집계',
        'G 분석': 'G(수출)집계',
        'E(품목성질물가)분석': 'E(품목성질물가)집계',
        'D(고용률)분석': 'D(고용률)집계',
    }
    
    agg_sheet = sheet_mapping.get(sheet_name)
    if not agg_sheet:
        return None
    
    base_config = AGGREGATE_SHEET_BASE_COLS.get(agg_sheet)
    if not base_config:
        return None
    
    curr_col = get_dynamic_col(agg_sheet, year, quarter, is_aggregate=True)
    prev_col = get_dynamic_col(agg_sheet, year - 1, quarter, is_aggregate=True)
    
    config = {
        'aggregate_sheet': agg_sheet,
        'region_col': base_config['region_col'],
        'code_col': base_config['code_col'],
        'total_code': base_config['total_code'],
        'curr_col': curr_col,
        'prev_col': prev_col,
    }
    
    # 고용률은 차이 계산
    if sheet_name == 'D(고용률)분석':
        config['calc_type'] = 'difference'
    
    return config


# =========================================================


def safe_float(value, default=None):
    """안전한 float 변환 함수 (NaN, '-', 빈 문자열 체크 포함)"""
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
        if isinstance(value, str):
            value = value.strip()
            if value == '-' or value == '' or value.lower() in ['없음', 'nan', 'none']:
                return default
        result = float(value)
        if pd.isna(result):
            return default
        return result
    except (ValueError, TypeError):
        return default


# 지역명 정식 명칭 → 약칭 매핑
REGION_NAME_MAP = {
    '서울특별시': '서울', '부산광역시': '부산', '대구광역시': '대구',
    '인천광역시': '인천', '광주광역시': '광주', '대전광역시': '대전',
    '울산광역시': '울산', '세종특별자치시': '세종', '경기도': '경기',
    '강원도': '강원', '충청북도': '충북', '충청남도': '충남',
    '전라북도': '전북', '전라남도': '전남', '경상북도': '경북',
    '경상남도': '경남', '제주특별자치도': '제주',
    # 추가 변형 (강원특별자치도 등)
    '강원특별자치도': '강원', '전북특별자치도': '전북',
}


def normalize_region_name(name):
    """지역명을 약칭으로 정규화"""
    if not name:
        return name
    name = str(name).strip()
    return REGION_NAME_MAP.get(name, name)


def get_summary_overview_data(excel_path, year, quarter):
    """요약-지역경제동향 데이터 추출
    
    ★ 핵심 원칙: 모든 나레이션 데이터는 테이블(get_summary_table_data)에서 가져옴
    - 테이블 데이터가 Single Source of Truth
    - 데이터 불일치 원천 차단
    """
    try:
        # ★ 테이블 데이터를 먼저 가져옴 (Single Source of Truth)
        table_data = get_summary_table_data(excel_path, year, quarter)
        
        # ★ 테이블 데이터에서 나레이션용 구조로 변환
        return _convert_table_to_narration(table_data)
        
    except Exception as e:
        print(f"요약 데이터 추출 오류: {e}")
        import traceback
        traceback.print_exc()
        return _get_default_summary_data()


def _convert_table_to_narration(table_data):
    """테이블 데이터를 나레이션용 구조로 변환
    
    테이블의 각 지표별 값을 증가/감소 지역으로 분류
    """
    nationwide = table_data.get('nationwide', {})
    region_groups = table_data.get('region_groups', [])
    
    # 모든 지역 데이터를 flat list로 변환
    all_regions = []
    for group in region_groups:
        for region in group.get('regions', []):
            all_regions.append(region)
    
    def extract_sector_data(key, is_employment=False):
        """특정 지표의 나레이션 데이터 추출"""
        nationwide_val = nationwide.get(key, 0.0)
        
        increase_regions = []
        decrease_regions = []
        
        for region in all_regions:
            val = region.get(key, 0.0)
            if val is None:
                continue
            
            region_data = {'name': region['name'], 'value': val}
            
            if val > 0:
                increase_regions.append(region_data)
            elif val < 0:
                decrease_regions.append(region_data)
            else:
                # 0인 경우: 고용률은 상승도 하락도 아님, 나머지는 감소 쪽
                if not is_employment:
                    decrease_regions.append(region_data)
        
        # 정렬
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': round(nationwide_val, 1) if nationwide_val else 0.0,
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions)
        }
    
    return {
        'production': {
            'mining': extract_sector_data('mining_production'),
            'service': extract_sector_data('service_production')
        },
        'consumption': extract_sector_data('retail_sales'),
        'exports': extract_sector_data('exports'),
        'price': extract_sector_data('price'),
        'employment': extract_sector_data('employment', is_employment=True)
    }


def _extract_sector_summary(xl, sheet_name):
    """시트에서 요약 데이터 추출 (기초자료 또는 집계 시트에서 전년동기비 계산)"""
    try:
        regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 기초자료 시트 설정 (우선 사용)
        raw_config = {
            'A 분석': {
                'raw_sheet': '광공업생산',
                'region_col': 1, 'code_col': 4, 'total_code': 'BCD',
                'curr_col': 64, 'prev_col': 60,  # 2025 2/4p, 2024 2/4
            },
            'B 분석': {
                'raw_sheet': '서비스업생산',
                'region_col': 1, 'code_col': 4, 'total_code': 'E~S',
                'curr_col': 64, 'prev_col': 60,
            },
            'C 분석': {
                'raw_sheet': '소비(소매, 추가)',
                'region_col': 1, 'code_col': 4, 'total_code': '총지수',
                'curr_col': 63, 'prev_col': 59,
            },
            'G 분석': {
                'raw_sheet': '수출',
                'region_col': 1, 'code_col': 5, 'total_code': '합계',
                'curr_col': 68, 'prev_col': 64,
            },
            'E(품목성질물가)분석': {
                'raw_sheet': '품목성질별 물가',
                'region_col': 0, 'code_col': 3, 'total_code': '총지수',
                'curr_col': 56, 'prev_col': 52,
            },
            'D(고용률)분석': {
                'raw_sheet': '연령별고용률',
                'region_col': 1, 'code_col': 3, 'total_code': '계',
                'curr_col': 66, 'prev_col': 62,
                'calc_type': 'difference'  # 고용률은 %p
            },
        }
        
        # 집계 시트 설정 (fallback) - 실제 엑셀 열 구조에 맞게 수정
        aggregate_config = {
            'A 분석': {
                'aggregate_sheet': 'A(광공업생산)집계',
                'region_col': 4, 'code_col': 7, 'total_code': 'BCD',
                'curr_col': 26, 'prev_col': 22,
            },
            'B 분석': {
                'aggregate_sheet': 'B(서비스업생산)집계',
                'region_col': 3, 'code_col': 6, 'total_code': 'E~S',
                'curr_col': 25, 'prev_col': 21,
            },
            'C 분석': {
                'aggregate_sheet': 'C(소비)집계',
                'region_col': 2, 'code_col': 6, 'total_code': '총지수',
                'curr_col': 24, 'prev_col': 20,
            },
            'G 분석': {
                'aggregate_sheet': 'G(수출)집계',
                'region_col': 3, 'code_col': 4, 'total_code': '0',  # division_col 사용
                'curr_col': 26, 'prev_col': 22,  # 실제 열 위치
            },
            'E(품목성질물가)분석': {
                'aggregate_sheet': 'E(품목성질물가)집계',
                'region_col': 0, 'code_col': 3, 'total_code': '총지수',
                'curr_col': 21, 'prev_col': 17,
            },
            'D(고용률)분석': {
                'aggregate_sheet': 'D(고용률)집계',
                'region_col': 1, 'code_col': 3, 'total_code': '계',
                'curr_col': 21, 'prev_col': 17,
                'calc_type': 'difference'
            },
        }
        
        # 기초자료 시트 우선 시도
        config = raw_config.get(sheet_name)
        actual_sheet = None
        
        if config and config.get('raw_sheet') in xl.sheet_names:
            actual_sheet = config['raw_sheet']
        else:
            # 집계 시트 fallback
            config = aggregate_config.get(sheet_name)
            if config and config.get('aggregate_sheet') in xl.sheet_names:
                actual_sheet = config['aggregate_sheet']
        
        if not actual_sheet:
            print(f"시트 없음: {sheet_name}")
            return _get_default_sector_summary()
        
        df = pd.read_excel(xl, sheet_name=actual_sheet, header=None)
        
        region_col = config['region_col']
        code_col = config.get('code_col')
        division_col = config.get('division_col')
        total_code = config['total_code']
        curr_col = config['curr_col']
        prev_col = config['prev_col']
        calc_type = config.get('calc_type', 'growth_rate')
        
        increase_regions = []
        decrease_regions = []
        nationwide = 0.0
        
        for i, row in df.iterrows():
            try:
                region_raw = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
                region = normalize_region_name(region_raw)
                
                # 총지수 행 찾기
                is_total_row = False
                if code_col is not None:
                    code = str(row[code_col]).strip() if pd.notna(row[code_col]) else ''
                    is_total_row = (code == total_code)
                elif division_col is not None:
                    division = str(row[division_col]).strip() if pd.notna(row[division_col]) else ''
                    is_total_row = (division == total_code)
                
                if is_total_row:
                    # 전년동기비 계산
                    curr_val = safe_float(row[curr_col], 0)
                    prev_val = safe_float(row[prev_col], 0)
                    
                    # 계산 방식에 따라 증감률 또는 차이 계산
                    if calc_type == 'difference':
                        change = round(curr_val - prev_val, 1) if (curr_val is not None and prev_val is not None) else 0.0
                    else:  # growth_rate
                        if prev_val is not None and prev_val != 0:
                            change = round((curr_val - prev_val) / prev_val * 100, 1)
                        else:
                            change = 0.0
                    
                    if region == '전국':
                        nationwide = change
                    elif region in regions:
                        if change >= 0:
                            increase_regions.append({'name': region, 'value': change})
                        else:
                            decrease_regions.append({'name': region, 'value': change})
            except Exception as e:
                continue
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': round(nationwide, 1),
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions)
        }
    except Exception as e:
        print(f"{sheet_name} 데이터 추출 오류: {e}")
        import traceback
        traceback.print_exc()
        return _get_default_sector_summary()


def _extract_price_summary_from_aggregate(xl, regions):
    """E(품목성질물가)집계 시트에서 소비자물가 증감률 추출"""
    try:
        df = pd.read_excel(xl, sheet_name='E(품목성질물가)집계', header=None)
        
        # 열 구조: 0=지역이름, 1=분류단계, 2=가중치, 3=분류이름
        # 열 20=2024 2/4분기, 열 24=2025 2/4분기
        
        increase_regions = []
        decrease_regions = []
        nationwide = 0.0
        
        for i, row in df.iterrows():
            try:
                region_raw = str(row[0]).strip() if pd.notna(row[0]) else ''
                region = normalize_region_name(region_raw)  # 지역명 정규화
                division = str(row[1]).strip() if pd.notna(row[1]) else ''
                
                # 총지수 행 (division == '0')
                if division == '0':
                    # 2025 2/4분기 지수 (열 24)와 2024 2/4분기 지수 (열 20)
                    curr_val = safe_float(row[24], 0)
                    prev_val = safe_float(row[20], 0)
                    
                    # 전년동분기 대비 증감률 계산
                    if prev_val is not None and prev_val != 0:
                        change = round((curr_val - prev_val) / prev_val * 100, 1)
                    else:
                        change = 0.0
                    
                    if region == '전국':
                        nationwide = change
                    elif region in regions:
                        if change >= 0:
                            increase_regions.append({'name': region, 'value': change})
                        else:
                            decrease_regions.append({'name': region, 'value': change})
            except Exception as e:
                continue
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': round(nationwide, 1),
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions)
        }
    except Exception as e:
        print(f"물가 집계 데이터 추출 오류: {e}")
        import traceback
        traceback.print_exc()
        return _get_default_sector_summary()


def _extract_employment_summary_from_aggregate(xl, regions):
    """D(고용률)집계 시트에서 고용률 증감 추출"""
    try:
        df = pd.read_excel(xl, sheet_name='D(고용률)집계', header=None)
        
        # 열 구조: 0=지역코드, 1=지역이름, 2=분류단계, 3=산업이름
        # 열 24=2025 2/4분기 고용률, 열 20=2024 2/4분기 고용률
        
        increase_regions = []
        decrease_regions = []
        nationwide = 0.0
        
        for i, row in df.iterrows():
            try:
                region_raw = str(row[1]).strip() if pd.notna(row[1]) else ''
                region = normalize_region_name(region_raw)  # 지역명 정규화
                division = str(row[2]).strip() if pd.notna(row[2]) else ''
                industry = str(row[3]).strip() if pd.notna(row[3]) else ''
                
                # 총계 행 (division == '0' 또는 industry == '계')
                if division == '0' or industry == '계':
                    # 2025 2/4분기 고용률 (열 24)와 2024 2/4분기 고용률 (열 20)
                    curr_val = safe_float(row[24], 0)
                    prev_val = safe_float(row[20], 0)
                    
                    # 전년동분기 대비 증감 (고용률은 %p 단위)
                    change = round(curr_val - prev_val, 1) if (curr_val is not None and prev_val is not None) else 0.0
                    
                    if region == '전국':
                        nationwide = change
                    elif region in regions:
                        if change >= 0:
                            increase_regions.append({'name': region, 'value': change})
                        else:
                            decrease_regions.append({'name': region, 'value': change})
            except Exception as e:
                continue
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': round(nationwide, 1),
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions)
        }
    except Exception as e:
        print(f"고용률 집계 데이터 추출 오류: {e}")
        import traceback
        traceback.print_exc()
        return _get_default_sector_summary()


def _get_default_summary_data():
    """기본 요약 데이터"""
    default_sector = _get_default_sector_summary()
    return {
        'production': {
            'mining': default_sector,
            'service': default_sector
        },
        'consumption': default_sector,
        'exports': default_sector,
        'price': default_sector,
        'employment': default_sector
    }


def _get_default_sector_summary():
    """기본 부문 요약 데이터"""
    return {
        'nationwide': 0.0,
        'increase_regions': [{'name': '-', 'value': 0.0}],
        'decrease_regions': [{'name': '-', 'value': 0.0}],
        'increase_count': 0,
        'decrease_count': 0,
        'above_regions': [{'name': '-', 'value': 0.0}],
        'below_regions': [{'name': '-', 'value': 0.0}],
        'above_count': 0,
        'below_count': 0
    }


def get_summary_table_data(excel_path, year=None, quarter=None):
    """요약 테이블 데이터 (기초자료 또는 집계 시트에서 전년동기비 계산)
    
    Args:
        excel_path: 엑셀 파일 경로
        year: 대상 연도 (None이면 현재 날짜 기준)
        quarter: 대상 분기 (None이면 현재 날짜 기준)
    """
    try:
        print(f"[DEBUG] get_summary_table_data - excel_path: {excel_path}, year: {year}, quarter: {quarter}")
        xl = pd.ExcelFile(excel_path)
        print(f"[DEBUG] 시트 목록: {xl.sheet_names[:5]}...")
        all_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                       '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # year, quarter가 None이면 현재 날짜 기준
        if year is None or quarter is None:
            from datetime import datetime
            now = datetime.now()
            year = year or now.year
            quarter = quarter or ((now.month - 1) // 3) + 1
        
        print(f"[DEBUG] 사용할 연도/분기: {year}년 {quarter}분기")
        
        # 기초자료 시트 설정 (동적 컬럼 계산)
        raw_sheet_configs = {
            'mining_production': {
                'sheet': '광공업생산',
                'region_col': 1, 'code_col': 4, 'total_code': 'BCD',
                'curr_col': get_dynamic_col('광공업생산', year, quarter),
                'prev_col': get_dynamic_col('광공업생산', year - 1, quarter),
                'calc_type': 'growth_rate'
            },
            'service_production': {
                'sheet': '서비스업생산',
                'region_col': 1, 'code_col': 4, 'total_code': 'E~S',
                'curr_col': get_dynamic_col('서비스업생산', year, quarter),
                'prev_col': get_dynamic_col('서비스업생산', year - 1, quarter),
                'calc_type': 'growth_rate'
            },
            'retail_sales': {
                'sheet': '소비(소매, 추가)',
                'region_col': 1, 'code_col': 4, 'total_code': '총지수',
                'curr_col': get_dynamic_col('소비(소매, 추가)', year, quarter),
                'prev_col': get_dynamic_col('소비(소매, 추가)', year - 1, quarter),
                'calc_type': 'growth_rate'
            },
            'exports': {
                'sheet': '수출',
                'region_col': 1, 'code_col': 5, 'total_code': '합계',
                'curr_col': get_dynamic_col('수출', year, quarter),
                'prev_col': get_dynamic_col('수출', year - 1, quarter),
                'calc_type': 'growth_rate'
            },
            'price': {
                'sheet': '품목성질별 물가',
                'region_col': 0, 'code_col': 3, 'total_code': '총지수',
                'curr_col': get_dynamic_col('품목성질별 물가', year, quarter),
                'prev_col': get_dynamic_col('품목성질별 물가', year - 1, quarter),
                'calc_type': 'growth_rate'
            },
            'employment': {
                'sheet': '연령별고용률',
                'region_col': 1, 'code_col': 3, 'total_code': '계',
                'curr_col': get_dynamic_col('연령별고용률', year, quarter),
                'prev_col': get_dynamic_col('연령별고용률', year - 1, quarter),
                'calc_type': 'difference'  # 고용률은 %p
            },
        }
        
        # 집계 시트 설정 (동적 컬럼 계산)
        aggregate_sheet_configs = {
            'mining_production': {
                'sheet': 'A(광공업생산)집계',
                'region_col': 4, 'code_col': 7, 'total_code': 'BCD',
                'curr_col': get_dynamic_col('A(광공업생산)집계', year, quarter, is_aggregate=True),
                'prev_col': get_dynamic_col('A(광공업생산)집계', year - 1, quarter, is_aggregate=True),
                'calc_type': 'growth_rate'
            },
            'service_production': {
                'sheet': 'B(서비스업생산)집계',
                'region_col': 3, 'code_col': 6, 'total_code': 'E~S',
                'curr_col': get_dynamic_col('B(서비스업생산)집계', year, quarter, is_aggregate=True),
                'prev_col': get_dynamic_col('B(서비스업생산)집계', year - 1, quarter, is_aggregate=True),
                'calc_type': 'growth_rate'
            },
            'retail_sales': {
                'sheet': 'C(소비)집계',
                'region_col': 2, 'code_col': 6, 'total_code': '총지수',
                'curr_col': get_dynamic_col('C(소비)집계', year, quarter, is_aggregate=True),
                'prev_col': get_dynamic_col('C(소비)집계', year - 1, quarter, is_aggregate=True),
                'calc_type': 'growth_rate'
            },
            'exports': {
                'sheet': 'G(수출)집계',
                'region_col': 3, 'division_col': 4, 'total_code': '0',
                'curr_col': get_dynamic_col('G(수출)집계', year, quarter, is_aggregate=True),
                'prev_col': get_dynamic_col('G(수출)집계', year - 1, quarter, is_aggregate=True),
                'calc_type': 'growth_rate'
            },
            'price': {
                'sheet': 'E(품목성질물가)집계',
                'region_col': 0, 'code_col': 3, 'total_code': '총지수',
                'curr_col': get_dynamic_col('E(품목성질물가)집계', year, quarter, is_aggregate=True),
                'prev_col': get_dynamic_col('E(품목성질물가)집계', year - 1, quarter, is_aggregate=True),
                'calc_type': 'growth_rate'
            },
            'employment': {
                'sheet': 'D(고용률)집계',
                'region_col': 1, 'code_col': 3, 'total_code': '계',
                'curr_col': get_dynamic_col('D(고용률)집계', year, quarter, is_aggregate=True),
                'prev_col': get_dynamic_col('D(고용률)집계', year - 1, quarter, is_aggregate=True),
                'calc_type': 'difference'
            },
        }
        
        nationwide_data = {
            'mining_production': 0.0, 'service_production': 0.0, 'retail_sales': 0.0,
            'exports': 0.0, 'price': 0.0, 'employment': 0.0
        }
        
        region_data = {r: {'name': r, 'mining_production': 0.0, 'service_production': 0.0,
                          'retail_sales': 0.0, 'exports': 0.0, 'price': 0.0, 'employment': 0.0}
                      for r in all_regions}
        
        for key in raw_sheet_configs.keys():
            # 기초자료 시트 우선 시도
            config = raw_sheet_configs[key]
            sheet_name = config['sheet']
            
            if sheet_name not in xl.sheet_names:
                # 집계 시트 fallback
                config = aggregate_sheet_configs.get(key)
                if config:
                    sheet_name = config['sheet']
                    if sheet_name not in xl.sheet_names:
                        print(f"시트 없음: {sheet_name}")
                        continue
                else:
                    continue
            
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
                region_col = config['region_col']
                code_col = config.get('code_col')
                division_col = config.get('division_col')
                total_code = config['total_code']
                curr_col = config['curr_col']
                prev_col = config['prev_col']
                calc_type = config['calc_type']
                
                print(f"[TABLE] {key}: sheet={sheet_name}, curr_col={curr_col}, prev_col={prev_col}, total_code={total_code}")
                found_regions = []
                
                for i, row in df.iterrows():
                    try:
                        region_raw = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
                        region = normalize_region_name(region_raw)  # 지역명 정규화
                        
                        # 총지수 행 찾기
                        is_total_row = False
                        if code_col is not None:
                            code = str(row[code_col]).strip() if pd.notna(row[code_col]) else ''
                            is_total_row = (code == total_code)
                        elif division_col is not None:
                            division = str(row[division_col]).strip() if pd.notna(row[division_col]) else ''
                            is_total_row = (division == total_code)
                        
                        if is_total_row:
                            # 컬럼 인덱스 범위 확인
                            if curr_col is None or prev_col is None:
                                print(f"[DEBUG] {sheet_name}/{region}: curr_col={curr_col}, prev_col={prev_col} - 컬럼 없음")
                                continue
                            if curr_col >= len(row) or prev_col >= len(row):
                                print(f"[DEBUG] {sheet_name}/{region}: 컬럼 범위 초과 (curr_col={curr_col}, prev_col={prev_col}, row_len={len(row)})")
                                continue
                            
                            curr_val = safe_float(row[curr_col], 0)
                            prev_val = safe_float(row[prev_col], 0)
                            
                            # 계산 방식에 따라 증감률 또는 차이 계산
                            if calc_type == 'difference':
                                value = round(curr_val - prev_val, 1) if (curr_val is not None and prev_val is not None) else 0.0
                            else:  # growth_rate
                                if prev_val is not None and prev_val != 0:
                                    value = round((curr_val - prev_val) / prev_val * 100, 1)
                                else:
                                    value = 0.0
                            
                            # 디버그: 비정상 값 출력
                            if value == -100.0 or (value == 0.0 and key != 'employment'):
                                print(f"[DEBUG] {sheet_name}/{region}: curr={curr_val}, prev={prev_val} → {value}% (cols: {curr_col}, {prev_col})")
                            
                            if region == '전국':
                                nationwide_data[key] = value
                                found_regions.append('전국')
                            elif region in all_regions:
                                region_data[region][key] = value
                                found_regions.append(region)
                    except:
                        continue
                
                print(f"[TABLE] {key}: 찾은 지역 {len(found_regions)}개: {found_regions[:5]}...")
            except Exception as e:
                print(f"{sheet_name} 테이블 데이터 추출 오류: {e}")
                continue
        
        region_groups = [
            {'name': '경인', 'regions': [region_data['서울'], region_data['인천'], region_data['경기']]},
            {'name': '충청', 'regions': [region_data['대전'], region_data['세종'], region_data['충북'], region_data['충남']]},
            {'name': '호남', 'regions': [region_data['광주'], region_data['전북'], region_data['전남'], region_data['제주']]},
            {'name': '동북', 'regions': [region_data['대구'], region_data['경북'], region_data['강원']]},
            {'name': '동남', 'regions': [region_data['부산'], region_data['울산'], region_data['경남']]},
        ]
        
        return {
            'nationwide': nationwide_data,
            'region_groups': region_groups
        }
    except Exception as e:
        print(f"요약 테이블 데이터 오류: {e}")
        import traceback
        traceback.print_exc()
        return {'nationwide': {'mining_production': 0.0, 'service_production': 0.0, 'retail_sales': 0.0,
                              'exports': 0.0, 'price': 0.0, 'employment': 0.0}, 'region_groups': []}


def get_production_summary_data(excel_path, year, quarter):
    """요약-생산 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        mining = _extract_chart_data(xl, 'A 분석')
        service = _extract_chart_data(xl, 'B 분석')
        
        return {
            'mining_production': mining,
            'service_production': service
        }
    except Exception as e:
        print(f"생산 요약 데이터 오류: {e}")
        return {
            'mining_production': _get_default_chart_data(),
            'service_production': _get_default_chart_data()
        }


def get_consumption_construction_data(excel_path, year, quarter):
    """요약-소비건설 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        retail = _extract_chart_data(xl, 'C 분석')
        
        # 건설 데이터 추출 - year, quarter 파라미터 전달
        construction = _extract_construction_chart_data(xl, year, quarter)
        
        return {
            'retail_sales': retail,
            'construction': construction
        }
    except Exception as e:
        print(f"소비건설 요약 데이터 오류: {e}")
        return {
            'retail_sales': _get_default_chart_data(),
            'construction': _get_default_construction_data()
        }


def _extract_construction_chart_data(xl, year=2025, quarter=3):
    """건설수주액 차트 데이터 추출 - 건설 (공표자료) 시트 사용"""
    try:
        regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        nationwide = {'amount': 0, 'change': 0.0}
        increase_regions = []
        decrease_regions = []
        chart_data = []
        seen_regions = set()  # 중복 방지
        
        # 건설 (공표자료) 시트 사용
        sheet_name = '건설 (공표자료)'
        if sheet_name in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
            
            # 동적 열 계산 - 건설 시트는 2024 3Q = col64 기준
            curr_col = get_dynamic_col(sheet_name, year, quarter, is_aggregate=False)
            prev_col = get_dynamic_col(sheet_name, year - 1, quarter, is_aggregate=False)
            
            print(f"[건설] 동적 열 계산: curr_col={curr_col}, prev_col={prev_col} (year={year}, quarter={quarter})")
            
            for i, row in df.iterrows():
                try:
                    region = str(row[1]).strip() if pd.notna(row[1]) else ''
                    level = str(row[2]).strip() if pd.notna(row[2]) else ''
                    code = str(row[3]).strip() if pd.notna(row[3]) else ''
                    name = str(row[4]).strip() if pd.notna(row[4]) else ''
                    
                    # 이미 처리한 지역은 건너뛰기
                    if region in seen_regions:
                        continue
                    
                    # 총계 행 (level == '0', code == '0') - '계' 데이터
                    if level == '0' and code == '0':
                        curr_val = safe_float(row[curr_col], 0) if curr_col and curr_col < len(row) else 0
                        prev_val = safe_float(row[prev_col], 0) if prev_col and prev_col < len(row) else 0
                        
                        # 증감률 계산
                        if prev_val is not None and prev_val != 0:
                            change = round((curr_val - prev_val) / prev_val * 100, 1)
                        else:
                            change = 0.0
                        
                        # 금액 (백억원 단위로 변환 - 원본은 십억원 단위)
                        amount = int(round(curr_val / 10, 0))
                        amount_normalized = min(100, max(0, curr_val / 30))  # 최대 3000억원 기준
                        
                        if region == '전국':
                            if 'nationwide' not in seen_regions:
                                nationwide['amount'] = amount
                                nationwide['change'] = change
                                seen_regions.add('nationwide')
                                print(f"[건설] 전국: curr={curr_val}, prev={prev_val}, change={change}%")
                        elif region in regions:
                            seen_regions.add(region)
                            data = {
                                'name': region,
                                'value': change,
                                'amount': amount,
                                'amount_normalized': amount_normalized,
                                'change': change
                            }
                            
                            if change >= 0:
                                increase_regions.append(data)
                            else:
                                decrease_regions.append(data)
                            chart_data.append(data)
                except Exception as row_err:
                    continue
        else:
            print(f"[건설] '{sheet_name}' 시트를 찾을 수 없습니다.")
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': nationwide,
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'chart_data': chart_data[:18]
        }
    except Exception as e:
        print(f"건설 차트 데이터 추출 오류: {e}")
        import traceback
        traceback.print_exc()
        return _get_default_construction_data()


def _get_default_construction_data():
    """기본 건설 데이터"""
    return {
        'nationwide': {'amount': 0, 'change': 0.0},
        'increase_regions': [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
        'decrease_regions': [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
        'increase_count': 0, 'decrease_count': 0,
        'chart_data': []
    }


def get_trade_price_data(excel_path, year, quarter):
    """요약-수출물가 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        exports = _extract_chart_data(xl, 'G 분석', is_trade=True)
        price = _extract_chart_data(xl, 'E(품목성질물가)분석')
        
        return {
            'exports': exports,
            'price': price
        }
    except Exception as e:
        print(f"수출 데이터 추출 오류: {e}")
        return {
            'exports': _get_default_trade_data(),
            'price': _get_default_chart_data()
        }


def get_employment_population_data(excel_path, year, quarter):
    """요약-고용인구 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        employment = _extract_chart_data(xl, 'D(고용률)분석', is_employment=True)
        
        population = {
            'inflow_regions': [],
            'outflow_regions': [],
            'inflow_count': 0,
            'outflow_count': 0,
            'chart_data': []
        }
        try:
            df = pd.read_excel(xl, sheet_name='I(순인구이동)집계', header=None)
            regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                       '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
            
            # 시트 구조: col4=지역이름, col5=분류단계(0=합계), col25=2025 2/4분기, col21=2024 2/4분기
            # 합계(분류단계 0) 행만 추출
            processed_regions = set()
            region_data = {}  # 지역별 데이터 저장
            
            for i, row in df.iterrows():
                region = str(row[4]).strip() if pd.notna(row[4]) else ''
                division = str(row[5]).strip() if pd.notna(row[5]) else ''
                
                # 합계 행 (분류단계 0)만 처리, 중복 지역 방지
                if division == '0' and region in regions and region not in processed_regions:
                    try:
                        # 2025 2/4분기 데이터 (열 25)와 2024 2/4분기 데이터 (열 21)
                        curr_value = safe_float(row[25], 0)
                        prev_value = safe_float(row[21], 0)
                        value = int(curr_value) if curr_value is not None else 0
                        
                        # 전년동분기대비 증감률 계산 (천명 단위이므로 직접 비교)
                        if prev_value is not None and prev_value != 0:
                            change = round((curr_value - prev_value) / abs(prev_value) * 100, 1)
                        else:
                            change = 0.0
                        
                        processed_regions.add(region)
                        region_data[region] = {'value': value, 'change': change}
                        
                        if value > 0:
                            population['inflow_regions'].append({'name': region, 'value': value})
                        else:
                            population['outflow_regions'].append({'name': region, 'value': abs(value)})
                    except:
                        continue
            
            population['inflow_regions'].sort(key=lambda x: x['value'], reverse=True)
            population['outflow_regions'].sort(key=lambda x: x['value'], reverse=True)
            population['inflow_count'] = len(population['inflow_regions'])
            population['outflow_count'] = len(population['outflow_regions'])
            
            # chart_data 구성 - 지역 순서대로
            for region in regions:
                if region in region_data:
                    data = region_data[region]
                    population['chart_data'].append({
                        'name': region,
                        'value': data['value'],  # 순이동량 (천명)
                        'change': data['change']  # 전년동분기대비 증감률 (%)
                    })
                else:
                    population['chart_data'].append({
                        'name': region,
                        'value': 0,
                        'change': 0.0
                    })
                    
        except Exception as e:
            print(f"인구이동 데이터 오류: {e}")
            import traceback
            traceback.print_exc()
        
        return {
            'employment': employment,
            'population': population
        }
    except Exception as e:
        print(f"고용인구 요약 데이터 오류: {e}")
        return {
            'employment': _get_default_employment_data(),
            'population': {'inflow_regions': [], 'outflow_regions': [], 'inflow_count': 0, 
                          'outflow_count': 0, 'chart_data': []}
        }


def _extract_chart_data(xl, sheet_name, is_trade=False, is_employment=False):
    """차트용 데이터 추출 (시트별 열 설정 적용, 분석 시트 없으면 기초자료에서 직접 계산)"""
    try:
        regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 기초자료 시트 설정 (분석 시트 없을 때 fallback)
        raw_sheet_config = {
            'A 분석': {
                'raw_sheet': '광공업생산',
                'region_col': 1, 'code_col': 4, 'total_code': 'BCD',
                'curr_col': 64, 'prev_col': 60,  # 2025 2/4p, 2024 2/4
            },
            'B 분석': {
                'raw_sheet': '서비스업생산',
                'region_col': 1, 'code_col': 4, 'total_code': 'E~S',
                'curr_col': 64, 'prev_col': 60,
            },
            'C 분석': {
                'raw_sheet': '소비(소매, 추가)',
                'region_col': 1, 'code_col': 4, 'total_code': '총지수',
                'curr_col': 63, 'prev_col': 59,
            },
            'G 분석': {
                'raw_sheet': '수출',
                'region_col': 1, 'code_col': 5, 'total_code': '합계',
                'curr_col': 68, 'prev_col': 64,
                'is_amount': True
            },
            'E(품목성질물가)분석': {
                'raw_sheet': '품목성질별 물가',
                'region_col': 1, 'code_col': 5, 'total_code': '총지수',
                'curr_col': 56, 'prev_col': 52,
            },
            'D(고용률)분석': {
                'raw_sheet': '연령별고용률',
                'region_col': 1, 'code_col': 3, 'total_code': '계',
                'curr_col': 66, 'prev_col': 62,
                'calc_type': 'difference'  # 고용률은 %p
            },
        }
        
        # 시트별 설정 (분석 시트와 집계 시트 매핑) - 실제 엑셀 열 구조에 맞게 수정
        sheet_config = {
            'A 분석': {
                'region_col': 3, 'code_col': 6, 'total_code': 'BCD',
                'change_col': 21,  # 증감률
                'index_sheet': 'A(광공업생산)집계',
                'index_region_col': 4, 'index_code_col': 7, 'index_total_code': 'BCD',
                'index_value_col': 26  # 2025 2/4분기 지수
            },
            'B 분석': {
                'region_col': 3, 'code_col': 6, 'total_code': 'E~S',
                'change_col': 20,  # 증감률
                'index_sheet': 'B(서비스업생산)집계',
                'index_region_col': 3, 'index_code_col': 6, 'index_total_code': 'E~S',
                'index_value_col': 25  # 2025 2/4분기 지수
            },
            'C 분석': {
                'region_col': 3, 'division_col': 4, 'total_code': '0',
                'change_col': 20,  # 증감률
                'index_sheet': 'C(소비)집계',
                'index_region_col': 2, 'index_code_col': 6, 'index_total_code': '총지수',
                'index_value_col': 24  # 2025 2/4분기 지수
            },
            'G 분석': {
                'region_col': 3, 'division_col': 4, 'total_code': '0',
                'change_col': 22,  # 증감률
                'index_sheet': 'G(수출)집계',
                'index_region_col': 3, 'index_code_col': 7, 'index_total_code': '합계',
                'index_value_col': 56,  # 2025 2/4분기 수출액
                'is_amount': True  # 금액 단위 (억달러 변환)
            },
            'E(품목성질물가)분석': {
                'region_col': 0, 'division_col': 1, 'total_code': '0',
                'change_col': 16,  # 증감률
                'index_sheet': 'E(품목성질물가)집계',
                'index_region_col': 0, 'index_code_col': 3, 'index_total_code': '총지수',
                'index_value_col': 21  # 2025 2/4분기 지수
            },
            'D(고용률)분석': {
                'region_col': 2, 'division_col': 3, 'total_code': '0',
                'rate_sheet': 'D(고용률)집계',
                'rate_region_col': 1, 'rate_code_col': 3, 'rate_total_code': '계',
                'rate_value_col': 21,  # 2025 2/4분기 고용률
                'prev_rate_col': 17  # 2024 2/4분기 고용률 (증감 계산용)
            },
        }
        
        config = sheet_config.get(sheet_name, {})
        raw_config = raw_sheet_config.get(sheet_name, {})
        
        if not config and not raw_config:
            return _get_default_chart_data()
        
        # 분석 시트 존재 여부 확인
        use_raw = sheet_name not in xl.sheet_names
        
        if use_raw and raw_config.get('raw_sheet') in xl.sheet_names:
            # 기초자료 시트에서 직접 전년동기비 계산
            return _extract_chart_data_from_raw(xl, raw_config, regions, is_trade, is_employment)
        elif not use_raw:
            # 분석 시트 사용 - 먼저 유효한 데이터가 있는지 확인
            df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
            
            # 분석 시트의 증감률 열이 모두 비어있는지 확인
            change_col = config.get('change_col', 20)
            has_valid_change = False
            if change_col < len(df.columns):
                region_col = config['region_col']
                for _, row in df.iterrows():
                    region = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
                    if region in regions:
                        val = row[change_col] if change_col < len(row) else None
                        if pd.notna(val) and val != '-' and val != '없음':
                            try:
                                float(val)
                                has_valid_change = True
                                break
                            except (ValueError, TypeError):
                                pass
            
            # 분석 시트에 유효한 증감률이 없으면 기초자료/집계 시트로 fallback
            if not has_valid_change:
                if raw_config.get('raw_sheet') in xl.sheet_names:
                    print(f"[요약] {sheet_name} 분석 시트 비어있음 → 기초자료 시트에서 계산")
                    return _extract_chart_data_from_raw(xl, raw_config, regions, is_trade, is_employment)
                else:
                    # 집계 시트에서 추출 시도
                    aggregate_config = {
                        'A 분석': {
                            'aggregate_sheet': 'A(광공업생산)집계',
                            'region_col': 4, 'code_col': 7, 'total_code': 'BCD',
                            'curr_col': 26, 'prev_col': 22,
                        },
                        'B 분석': {
                            'aggregate_sheet': 'B(서비스업생산)집계',
                            'region_col': 3, 'code_col': 6, 'total_code': 'E~S',
                            'curr_col': 25, 'prev_col': 21,
                        },
                        'C 분석': {
                            'aggregate_sheet': 'C(소비)집계',
                            'region_col': 2, 'code_col': 6, 'total_code': '총지수',
                            'curr_col': 24, 'prev_col': 20,
                        },
                        'G 분석': {
                            'aggregate_sheet': 'G(수출)집계',
                            'region_col': 3, 'code_col': 4, 'total_code': '0',
                            'curr_col': 26, 'prev_col': 22,
                            'is_amount': True
                        },
                        'E(품목성질물가)분석': {
                            'aggregate_sheet': 'E(지출목적물가)집계',
                            'region_col': 2, 'code_col': 3, 'total_code': '0',
                            'curr_col': 21, 'prev_col': 17,
                        },
                    }
                    agg_config = aggregate_config.get(sheet_name)
                    if agg_config and agg_config.get('aggregate_sheet') in xl.sheet_names:
                        print(f"[요약] {sheet_name} 분석 시트 비어있음 → 집계 시트에서 계산")
                        return _extract_chart_data_from_aggregate(xl, agg_config, regions, is_trade)
        else:
            return _get_default_chart_data()
        
        nationwide = {'index': 100.0, 'change': 0.0, 'rate': 60.0, 'amount': 0}
        increase_regions = []
        decrease_regions = []
        chart_data = []
        region_changes = {}  # 지역별 증감률 저장
        
        region_col = config['region_col']
        code_col = config.get('code_col')
        division_col = config.get('division_col')
        total_code = config['total_code']
        change_col = config.get('change_col', 20)
        
        nationwide_change_set = False  # 전국 증감률이 설정되었는지 추적
        
        for i, row in df.iterrows():
            try:
                region = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
                
                # 총지수 행인지 확인
                is_total_row = False
                if code_col is not None:
                    code = str(row[code_col]).strip() if pd.notna(row[code_col]) else ''
                    is_total_row = (code == total_code)
                elif division_col is not None:
                    division = str(row[division_col]).strip() if pd.notna(row[division_col]) else ''
                    is_total_row = (division == total_code)
                
                if is_total_row:
                    # 유효한 숫자 값인지 확인
                    change_val = None
                    if change_col < len(row):
                        change_val = safe_float(row[change_col], None)
                        if change_val is not None:
                            change_val = round(change_val, 1)
                    
                    if region == '전국':
                        # 첫 번째 유효한 전국 값만 사용
                        if not nationwide_change_set and change_val is not None:
                            nationwide['change'] = change_val
                            nationwide_change_set = True
                    elif region in regions and change_val is not None:
                        # 첫 번째 유효한 지역 값만 사용
                        if region not in region_changes:
                            region_changes[region] = change_val
            except:
                continue
        
        # 집계 시트에서 지수/고용률 값 추출
        region_indices = {}
        
        if is_employment and 'rate_sheet' in config:
            # 고용률 집계 시트에서 값 추출
            try:
                df_rate = pd.read_excel(xl, sheet_name=config['rate_sheet'], header=None)
                rate_region_col = config['rate_region_col']
                rate_code_col = config.get('rate_code_col')
                rate_division_col = config.get('rate_division_col')
                rate_total_code = config['rate_total_code']
                rate_value_col = config['rate_value_col']
                prev_rate_col = config.get('prev_rate_col', rate_value_col - 4)
                
                for i, row in df_rate.iterrows():
                    try:
                        region_raw = str(row[rate_region_col]).strip() if pd.notna(row[rate_region_col]) else ''
                        region = normalize_region_name(region_raw)  # 지역명 정규화
                        
                        # 코드 컬럼 또는 division 컬럼으로 총계 행 확인
                        is_total = False
                        if rate_code_col is not None:
                            code = str(row[rate_code_col]).strip() if pd.notna(row[rate_code_col]) else ''
                            is_total = (code == rate_total_code)
                        elif rate_division_col is not None:
                            division = str(row[rate_division_col]).strip() if pd.notna(row[rate_division_col]) else ''
                            is_total = (division == rate_total_code)
                        
                        if is_total:
                            rate_val = safe_float(row[rate_value_col], 60.0)
                            prev_rate = safe_float(row[prev_rate_col], rate_val if rate_val is not None else 60.0)
                            change_val = round(rate_val - prev_rate, 1) if (rate_val is not None and prev_rate is not None) else 0.0
                            
                            if region == '전국':
                                nationwide['rate'] = round(rate_val, 1)
                                nationwide['index'] = round(rate_val, 1)
                                nationwide['change'] = change_val
                            elif region in regions:
                                region_indices[region] = round(rate_val, 1)
                                region_changes[region] = change_val
                    except:
                        continue
            except Exception as e:
                print(f"고용률 집계 시트 오류: {e}")
        
        elif 'index_sheet' in config:
            # 지수 집계 시트에서 값 추출
            try:
                df_index = pd.read_excel(xl, sheet_name=config['index_sheet'], header=None)
                idx_region_col = config['index_region_col']
                idx_code_col = config.get('index_code_col')
                idx_division_col = config.get('index_division_col')
                idx_total_code = config['index_total_code']
                idx_value_col = config['index_value_col']
                
                nationwide_index_set = False  # 전국 지수가 설정되었는지 추적
                
                for i, row in df_index.iterrows():
                    try:
                        region_raw = str(row[idx_region_col]).strip() if pd.notna(row[idx_region_col]) else ''
                        region = normalize_region_name(region_raw)  # 지역명 정규화
                        
                        is_total = False
                        if idx_code_col is not None:
                            code = str(row[idx_code_col]).strip() if pd.notna(row[idx_code_col]) else ''
                            is_total = (code == str(idx_total_code))
                        elif idx_division_col is not None:
                            division = str(row[idx_division_col]).strip() if pd.notna(row[idx_division_col]) else ''
                            is_total = (division == str(idx_total_code))
                        
                        if is_total:
                            # 유효한 숫자 값인지 확인
                            index_val = safe_float(row[idx_value_col], None)
                            if index_val is not None:
                                index_val = round(index_val, 1)
                            
                            if region == '전국':
                                # 첫 번째 유효한 전국 값만 사용
                                if not nationwide_index_set and index_val is not None:
                                    nationwide['index'] = index_val
                                    if is_trade:
                                        nationwide['amount'] = round(index_val, 0)
                                    nationwide_index_set = True
                            elif region in regions and index_val is not None:
                                # 첫 번째 유효한 지역 값만 사용
                                if region not in region_indices:
                                    region_indices[region] = index_val
                    except:
                        continue
            except Exception as e:
                print(f"지수 집계 시트 오류: {e}")
        
        # 수출액 특별 처리 (G 분석) - 금액을 억달러 단위로 변환
        if is_trade and config.get('is_amount'):
            try:
                # G(수출)집계 시트에서 수출액 가져오기
                if 'G(수출)집계' in xl.sheet_names:
                    df_export = pd.read_excel(xl, sheet_name='G(수출)집계', header=None)
                    for i, row in df_export.iterrows():
                        try:
                            region = str(row[3]).strip() if pd.notna(row[3]) else ''
                            division = str(row[4]).strip() if pd.notna(row[4]) else ''
                            if division == '0':
                                # 2025 2/4분기 수출액 (열 26, 백만달러 → 억달러 변환)
                                amount_val = safe_float(row[26], 0)
                                amount_val = amount_val if amount_val is not None else 0
                                amount_in_billion = round(amount_val / 100, 0)  # 백만달러 → 억달러
                                if region == '전국':
                                    nationwide['amount'] = amount_in_billion
                                    nationwide['index'] = amount_in_billion  # 차트용
                                elif region in regions:
                                    region_indices[region] = amount_in_billion
                        except:
                            continue
            except Exception as e:
                print(f"수출 집계 시트 오류: {e}")
        
        # 차트 데이터 구성
        for region in regions:
            change_val = region_changes.get(region, 0.0)
            index_val = region_indices.get(region, 100.0)
            
            data = {
                'name': region,
                'value': change_val,
                'index': index_val,
                'change': change_val,
                'rate': index_val
            }
            
            if is_trade:
                data['amount'] = index_val
                data['amount_normalized'] = min(100, max(0, index_val / 6))
            
            if change_val >= 0:
                increase_regions.append(data)
            else:
                decrease_regions.append(data)
            chart_data.append(data)
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': nationwide,
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions),
            'chart_data': chart_data[:18]
        }
    except Exception as e:
        print(f"{sheet_name} 차트 데이터 오류: {e}")
        import traceback
        traceback.print_exc()
        if is_trade:
            return _get_default_trade_data()
        elif is_employment:
            return _get_default_employment_data()
        return _get_default_chart_data()


def _extract_chart_data_from_raw(xl, config, regions, is_trade=False, is_employment=False):
    """기초자료 시트에서 직접 차트 데이터 추출 및 전년동기비 계산"""
    try:
        df = pd.read_excel(xl, sheet_name=config['raw_sheet'], header=None)
        
        region_col = config['region_col']
        code_col = config.get('code_col')
        total_code = config['total_code']
        curr_col = config['curr_col']
        prev_col = config['prev_col']
        calc_type = config.get('calc_type', 'growth_rate')
        is_amount = config.get('is_amount', False)
        
        nationwide = {'index': 100.0, 'change': 0.0, 'rate': 60.0, 'amount': 0}
        increase_regions = []
        decrease_regions = []
        chart_data = []
        
        for i, row in df.iterrows():
            try:
                region = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
                code = str(row[code_col]).strip() if code_col and pd.notna(row[code_col]) else ''
                
                if code != total_code:
                    continue
                
                # 현재 분기와 전년동기 값
                curr_val = safe_float(row[curr_col], 0)
                prev_val = safe_float(row[prev_col], 0)
                
                # 전년동기비 계산
                if calc_type == 'difference':
                    change = round(curr_val - prev_val, 1) if (curr_val is not None and prev_val is not None) else 0.0
                else:  # growth_rate
                    if prev_val is not None and prev_val != 0:
                        change = round((curr_val - prev_val) / prev_val * 100, 1)
                    else:
                        change = 0.0
                
                data = {
                    'name': region,
                    'value': change,
                    'index': round(curr_val, 1),
                    'change': change,
                    'rate': round(curr_val, 1) if is_employment else round(curr_val, 1)
                }
                
                if is_trade or is_amount:
                    # 수출액은 백만달러 → 억달러로 변환
                    amount = round(curr_val / 100, 1) if curr_val > 1000 else round(curr_val, 1)
                    data['amount'] = amount
                    data['amount_normalized'] = min(100, max(0, curr_val / 600))
                
                if region == '전국':
                    nationwide['index'] = round(curr_val, 1)
                    nationwide['change'] = change
                    nationwide['rate'] = round(curr_val, 1)
                    if is_trade or is_amount:
                        nationwide['amount'] = data.get('amount', 0)
                elif region in regions:
                    if change >= 0:
                        increase_regions.append(data)
                    else:
                        decrease_regions.append(data)
                    chart_data.append(data)
            except:
                continue
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': nationwide,
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions),
            'chart_data': chart_data[:18]
        }
    except Exception as e:
        print(f"기초자료 차트 데이터 오류: {e}")
        import traceback
        traceback.print_exc()
        return _get_default_chart_data()


def _extract_chart_data_from_aggregate(xl, config, regions, is_trade=False):
    """집계 시트에서 차트 데이터 추출 및 전년동기비 계산"""
    try:
        df = pd.read_excel(xl, sheet_name=config['aggregate_sheet'], header=None)
        
        region_col = config['region_col']
        code_col = config.get('code_col')
        total_code = config['total_code']
        curr_col = config['curr_col']
        prev_col = config['prev_col']
        is_amount = config.get('is_amount', False)
        
        nationwide = {'index': 100.0, 'change': 0.0, 'rate': 60.0, 'amount': 0}
        increase_regions = []
        decrease_regions = []
        chart_data = []
        
        for i, row in df.iterrows():
            try:
                region = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
                code = str(row[code_col]).strip() if code_col is not None and pd.notna(row[code_col]) else ''
                
                if code != total_code:
                    continue
                
                # 현재 분기와 전년동기 값
                curr_val = safe_float(row[curr_col], 0)
                prev_val = safe_float(row[prev_col], 0)
                
                # 전년동기비 계산
                if prev_val is not None and prev_val != 0:
                    change = round((curr_val - prev_val) / prev_val * 100, 1)
                else:
                    change = 0.0
                
                data = {
                    'name': region,
                    'value': change,
                    'index': round(curr_val, 1),
                    'change': change,
                    'rate': round(curr_val, 1)
                }
                
                if is_trade or is_amount:
                    # 금액 정규화
                    amount = round(curr_val / 100, 1) if curr_val > 1000 else round(curr_val, 1)
                    data['amount'] = amount
                    data['amount_normalized'] = min(100, max(0, curr_val / 600))
                
                if region == '전국':
                    nationwide['index'] = round(curr_val, 1)
                    nationwide['change'] = change
                    nationwide['rate'] = round(curr_val, 1)
                    if is_trade or is_amount:
                        nationwide['amount'] = data.get('amount', 0)
                elif region in regions:
                    if change >= 0:
                        increase_regions.append(data)
                    else:
                        decrease_regions.append(data)
                    chart_data.append(data)
            except:
                continue
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': nationwide,
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions),
            'chart_data': chart_data[:18]
        }
    except Exception as e:
        print(f"집계 시트 차트 데이터 오류: {e}")
        import traceback
        traceback.print_exc()
        return _get_default_chart_data()


def _get_default_chart_data():
    """기본 차트 데이터"""
    return {
        'nationwide': {'index': 100.0, 'change': 0.0, 'rate': 60.0, 'amount': 0},
        'increase_regions': [{'name': '-', 'value': 0.0, 'index': 100.0, 'change': 0.0, 'rate': 60.0}],
        'decrease_regions': [{'name': '-', 'value': 0.0, 'index': 100.0, 'change': 0.0, 'rate': 60.0}],
        'increase_count': 0, 'decrease_count': 0,
        'above_regions': [{'name': '-', 'value': 0.0}],
        'below_regions': [{'name': '-', 'value': 0.0}],
        'above_count': 0, 'below_count': 0,
        'chart_data': []
    }


def _get_default_trade_data():
    """기본 수출입 데이터"""
    return {
        'nationwide': {'amount': 0, 'change': 0.0},
        'increase_regions': [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
        'decrease_regions': [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
        'increase_count': 0, 'decrease_count': 0,
        'chart_data': []
    }


def _get_default_employment_data():
    """기본 고용 데이터"""
    return {
        'nationwide': {'rate': 60.0, 'change': 0.0},
        'increase_regions': [{'name': '-', 'value': 0.0, 'rate': 60.0, 'change': 0.0}],
        'decrease_regions': [{'name': '-', 'value': 0.0, 'rate': 60.0, 'change': 0.0}],
        'increase_count': 0, 'decrease_count': 0,
        'chart_data': []
    }

