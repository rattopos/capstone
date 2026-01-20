from __future__ import annotations

# -*- coding: utf-8 -*-
"""
보도자료 설정 및 상수 정의
"""

from typing import Any

REGIONAL_REPORTS: list[dict[str, Any]] = [
    {'id': 'region_seoul', 'name': '서울', 'full_name': '서울특별시', 'index': 1, 'icon': '🏙️'},
    {'id': 'region_busan', 'name': '부산', 'full_name': '부산광역시', 'index': 2, 'icon': '🌊'},
    {'id': 'region_daegu', 'name': '대구', 'full_name': '대구광역시', 'index': 3, 'icon': '🏛️'},
    {'id': 'region_incheon', 'name': '인천', 'full_name': '인천광역시', 'index': 4, 'icon': '✈️'},
    {'id': 'region_gwangju', 'name': '광주', 'full_name': '광주광역시', 'index': 5, 'icon': '🎨'},
    {'id': 'region_daejeon', 'name': '대전', 'full_name': '대전광역시', 'index': 6, 'icon': '🔬'},
    {'id': 'region_ulsan', 'name': '울산', 'full_name': '울산광역시', 'index': 7, 'icon': '🚗'},
    {'id': 'region_sejong', 'name': '세종', 'full_name': '세종특별자치시', 'index': 8, 'icon': '🏛️'},
    {'id': 'region_gyeonggi', 'name': '경기', 'full_name': '경기도', 'index': 9, 'icon': '🏘️'},
    {'id': 'region_gangwon', 'name': '강원', 'full_name': '강원특별자치도', 'index': 10, 'icon': '⛰️'},
    {'id': 'region_chungbuk', 'name': '충북', 'full_name': '충청북도', 'index': 11, 'icon': '🌾'},
    {'id': 'region_chungnam', 'name': '충남', 'full_name': '충청남도', 'index': 12, 'icon': '🌅'},
    {'id': 'region_jeonbuk', 'name': '전북', 'full_name': '전북특별자치도', 'index': 13, 'icon': '🌿'},
    {'id': 'region_jeonnam', 'name': '전남', 'full_name': '전라남도', 'index': 14, 'icon': '🍃'},
    {'id': 'region_gyeongbuk', 'name': '경북', 'full_name': '경상북도', 'index': 15, 'icon': '🏔️'},
    {'id': 'region_gyeongnam', 'name': '경남', 'full_name': '경상남도', 'index': 16, 'icon': '🌳'},
    {'id': 'region_jeju', 'name': '제주', 'full_name': '제주특별자치도', 'index': 17, 'icon': '🏝️'},
]

# 아래는 REGION_DISPLAY_MAPPING, REGION_GROUPS, VALID_REGIONS 등 통합 매핑 예시 (필요시 확장)
REGION_DISPLAY_MAPPING: dict[str, str] = {
    '서울': '서울특별시',
    '부산': '부산광역시',
    '대구': '대구광역시',
    '인천': '인천광역시',
    '광주': '광주광역시',
    '대전': '대전광역시',
    '울산': '울산광역시',
    '세종': '세종특별자치시',
    '경기': '경기도',
    '강원': '강원특별자치도',
    '충북': '충청북도',
    '충남': '충청남도',
    '전북': '전북특별자치도',
    '전남': '전라남도',
    '경북': '경상북도',
    '경남': '경상남도',
    '제주': '제주특별자치도',
}

REGION_GROUPS: dict[str, list[str]] = {
    '경인': ['서울', '인천', '경기'],
    '충청': ['대전', '세종', '충북', '충남'],
    '호남': ['광주', '전북', '전남', '제주'],
    '동북': ['대구', '경북', '강원'],
    '동남': ['부산', '울산', '경남'],
}

VALID_REGIONS: list[str] = [r['name'] for r in REGIONAL_REPORTS]

# ===== 요약 보도자료 목록 (요약만 포함) =====
# 주의: 표지, 일러두기, 목차, 인포그래픽, 차트, 통계표, GRDP는 고객사 요구사항 변경으로 더 이상 생성하지 않음
# 실무자는 표와 나레이션만 한글 문서에 복붙함
SUMMARY_REPORTS: list[dict[str, Any]] = [
    {
        'id': 'summary_overview',
        'name': '요약-지역경제동향',
        'sheet': 'multiple',
        'generator': None,
        'template': 'summary_regional_economy_template.html',
        'icon': '📈',
        'category': 'summary'
    },
    {
        'id': 'summary_production',
        'name': '요약-생산',
        'sheet': 'multiple',
        'generator': None,
        'template': 'summary_production_template.html',
        'icon': '🏭',
        'category': 'summary'
    },
    {
        'id': 'summary_consumption',
        'name': '요약-소비건설',
        'sheet': 'multiple',
        'generator': None,
        'template': 'summary_consumption_construction_template.html',
        'icon': '🛒',
        'category': 'summary'
    },
    {
        'id': 'summary_trade_price',
        'name': '요약-수출물가',
        'sheet': 'multiple',
        'generator': None,
        'template': 'summary_export_price_template.html',
        'icon': '📦',
        'category': 'summary'
    },
    {
        'id': 'summary_employment',
        'name': '요약-고용인구',
        'sheet': 'multiple',
        'generator': None,
        'template': 'summary_employment_template.html',
        'icon': '👔',
        'category': 'summary'
    },
]

# ===== 부문별 보도자료 순서 설정 =====
SECTOR_REPORTS: list[dict[str, Any]] = [
    {
        'id': 'manufacturing',
        'report_id': 'manufacturing',
        'name': '광공업생산',
        'sheet': 'A 분석',
        'generator': 'unified_generator.py',
        'template': 'mining_template.html',
        'icon': '🏭',
        'category': 'production',
        'class_name': 'MiningManufacturingGenerator',
        'name_mapping': {
            '의료, 정밀, 광학 기기 및 시계 제조업': '의료·정밀',
            '의료용 물질 및 의약품 제조업': '의약품',
            '기타 운송장비 제조업': '기타 운송장비',
            '기타 기계 및 장비 제조업': '기타기계장비',
            '전기장비 제조업': '전기장비',
            '자동차 및 트레일러 제조업': '자동차·트레일러',
            '전기, 가스, 증기 및 공기 조절 공급업': '전기·가스업',
            '전기업 및 가스업': '전기·가스업',
            '식료품 제조업': '식료품',
            '금속 가공제품 제조업; 기계 및 가구 제외': '금속가공제품',
            '1차 금속 제조업': '1차금속',
            '화학 물질 및 화학제품 제조업; 의약품 제외': '화학물질',
            '담배 제조업': '담배',
            '고무 및 플라스틱제품 제조업': '고무·플라스틱',
            '비금속 광물제품 제조업': '비금속광물',
            '섬유제품 제조업; 의복 제외': '섬유제품',
            '금속 광업': '금속광업',
            '산업용 기계 및 장비 수리업': '산업용기계',
            '펄프, 종이 및 종이제품 제조업': '펄프·종이',
            '인쇄 및 기록매체 복제업': '인쇄',
            '음료 제조업': '음료',
            '가구 제조업': '가구',
            '기타 제품 제조업': '기타제품',
            '가죽, 가방 및 신발 제조업': '가죽·신발',
            '의복, 의복액세서리 및 모피제품 제조업': '의복',
            '코크스, 연탄 및 석유정제품 제조업': '석유정제품',
            '목재 및 나무제품 제조업; 가구 제외': '목재제품',
            '비금속광물 광업; 연료용 제외': '비금속광물광업',
        },
        'aggregation_structure': {'total_code': 'BCD', 'sheet': 'A(광공업생산)집계'},
        'metadata_columns': ['region', 'classification', 'code', 'name']
    },
    {
        'id': 'service',
        'report_id': 'service',
        'name': '서비스업생산',
        'sheet': 'B 분석',
        'generator': 'unified_generator.py',
        'template': 'service_template.html',
        'icon': '🏢',
        'category': 'production',
        'class_name': 'ServiceIndustryGenerator',
        # B(서비스업생산)집계 시트의 업종(산업명)은 H열
        # 엑셀 H열(1-based) -> 0-based index 7
        'industry_name_col': 7,
        'name_mapping': {
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
        },
        'aggregation_structure': {'total_code': 'E~S', 'sheet': 'B(서비스업생산)집계'},
        'metadata_columns': ['region', 'classification', 'code', 'name']
    },
    {
        'id': 'consumption',
        'report_id': 'consumption',
        'name': '소비동향',
        'sheet': 'C 분석',
        'generator': 'unified_generator.py',
        'template': 'consumption_template.html',
        'icon': '🛒',
        'category': 'consumption',
        'class_name': 'ConsumptionGenerator',
        'name_mapping': {
            '백화점': '백화점',
            '대형마트': '대형마트',
            '면세점': '면세점',
            '슈퍼마켓 및 잡화점': '슈퍼마켓·잡화점',
            '슈퍼마켓· 잡화점 및 편의점': '슈퍼마켓·잡화점·편의점',
            '편의점': '편의점',
            '승용차 및 연료 소매점': '승용차·연료소매점',
            '전문소매점': '전문소매점',
            '무점포 소매': '무점포소매'
        },
        'aggregation_structure': {'total_code': 'A0', 'sheet': 'C(소비)집계'},
        'metadata_columns': ['region', 'classification', 'code', 'name'],
        'value_type': 'change_rate',  # 이미 계산된 증감률 직접 읽기
        'analysis_sheet': 'C 분석'
    },
    {
        'id': 'construction',
        'report_id': 'construction',
        'name': '건설동향',
        'sheet': "F'분석",
        'generator': 'unified_generator.py',
        'template': 'construction_template.html',
        'icon': '🏗️',
        'category': 'construction',
        'class_name': 'ConstructionGenerator',
        'name_mapping': {
            '건축': '건축',
            '토목': '토목',
            '주거용 건물': '주거용',
            '비주거용 건물': '비주거용',
        },
        'aggregation_structure': {'total_code': '0', 'sheet': "F'(건설)집계"},
        'metadata_columns': ['region', 'classification', 'code', 'name']
    },
    {
        'id': 'export',
        'report_id': 'export',
        'name': '수출',
        'sheet': 'G 분석',
        'generator': 'unified_generator.py',
        'template': 'export_template.html',
        'icon': '📦',
        'category': 'trade',
        'class_name': 'ExportGenerator',
        'name_mapping': {},
        'aggregation_structure': {'total_code': '합계', 'sheet': 'G(수출)집계'},
        'metadata_columns': ['region', 'classification', 'code', 'name'],
        'header_rows': 3  # 집계 시트 헤더 행 수 (데이터는 4행부터)
    },
    {
        'id': 'import',
        'report_id': 'import',
        'name': '수입',
        'sheet': 'H 분석',
        'generator': 'unified_generator.py',
        'template': 'import_template.html',
        'icon': '🚢',
        'category': 'trade',
        'class_name': 'ImportGenerator',
        'name_mapping': {},
        'aggregation_structure': {'total_code': '합계', 'sheet': 'H(수입)집계'},
        'metadata_columns': ['region', 'classification', 'code', 'name'],
        'header_rows': 3  # 집계 시트 헤더 행 수 (데이터는 4행부터)
    },
    {
        'id': 'price',
        'report_id': 'price',
        'name': '물가동향',
        'sheet': 'E(품목성질물가)분석',
        'generator': 'unified_generator.py',
        'template': 'price_template.html',
        'icon': '💰',
        'category': 'price',
        'class_name': 'PriceTrendGenerator',
        'name_mapping': {},
        'aggregation_structure': {'total_code': '00', 'sheet': 'E(지출목적물가)집계'},
        'metadata_columns': ['region', 'classification', 'code', 'name']
    },
    {
        'id': 'employment',
        'report_id': 'employment',
        'name': '고용률',
        'sheet': 'D(고용률)분석',
        'generator': 'unified_generator.py',
        'template': 'employment_template.html',
        'icon': '👔',
        'category': 'employment',
        'class_name': 'EmploymentRateGenerator',
        'name_mapping': {},
        'aggregation_structure': {'total_code': '계', 'sheet': 'D(고용률)집계'},
        'aggregation_range': {'start_row': 3, 'end_row': 111, 'start_col': 'A', 'end_col': 'V'},
        'metadata_columns': ['year', 'quarter', 'region'],
        'header_rows': 3  # 상단 2줄 설명 + 실제 헤더 1줄
    },
    {
        'id': 'unemployment',
        'report_id': 'unemployment',
        'name': '실업률',
        'sheet': 'D(실업)분석',
        'generator': 'unified_generator.py',
        'template': 'unemployment_template.html',
        'icon': '📉',
        'category': 'employment',
        'class_name': 'UnemploymentGenerator',
        'name_mapping': {},
        'aggregation_structure': {'total_code': '계', 'sheet': 'D(실업)집계'},
        'aggregation_range': {'start_row': 80, 'end_row': 152, 'start_col': 'A', 'end_col': 'T'},
        'metadata_columns': ['year', 'quarter', 'region'],
        'header_rows': 3  # 상단 2줄 설명 + 실제 헤더 1줄
    },
    {
        'id': 'migration',
        'report_id': 'migration',
        'name': '국내인구이동',
        'sheet': 'I(순인구이동)집계',  # 실제 Excel 시트명
        'generator': 'unified_generator.py',
        'template': 'migration_template.html',
        'icon': '👥',
        'category': 'population',
        'class_name': 'DomesticMigrationGenerator',
        'name_mapping': {},
        # 집계 시트의 합계 행은 연령별 컬럼에 '합계'로 표기됨
        'aggregation_structure': {'total_code': '합계', 'sheet': 'I(순인구이동)집계'},
        'metadata_columns': ['region', 'classification', 'code', 'name'],
        'require_industry_code': False,
        'has_nationwide': False  # 국내이동은 지역간 이동이므로 전국 합계(0)는 의미없음
    }
]

# 전체 보도자료 순서 (부문별 → 요약)
REPORT_ORDER = SECTOR_REPORTS + SUMMARY_REPORTS

# ===== 통계표 보도자료 목록 =====
# 주의: 고객사 요청으로 통계표 섹션 전체(통계표, GRDP, 부록)를 생성하지 않기로 결정됨
# 실무자는 요약, 부문별, 시도별의 표와 나레이션만 사용함
STATISTICS_REPORTS = []

# 시도별 보도자료 목록 (17개 시도 + 참고_GRDP)
REGIONAL_REPORTS = [
    {'id': 'region_seoul', 'name': '서울', 'full_name': '서울특별시', 'index': 1, 'icon': '🏙️'},
    {'id': 'region_busan', 'name': '부산', 'full_name': '부산광역시', 'index': 2, 'icon': '🌊'},
    {'id': 'region_daegu', 'name': '대구', 'full_name': '대구광역시', 'index': 3, 'icon': '🏛️'},
    {'id': 'region_incheon', 'name': '인천', 'full_name': '인천광역시', 'index': 4, 'icon': '✈️'},
    {'id': 'region_gwangju', 'name': '광주', 'full_name': '광주광역시', 'index': 5, 'icon': '🎨'},
    {'id': 'region_daejeon', 'name': '대전', 'full_name': '대전광역시', 'index': 6, 'icon': '🔬'},
    {'id': 'region_ulsan', 'name': '울산', 'full_name': '울산광역시', 'index': 7, 'icon': '🚗'},
    {'id': 'region_sejong', 'name': '세종', 'full_name': '세종특별자치시', 'index': 8, 'icon': '🏛️'},
    {'id': 'region_gyeonggi', 'name': '경기', 'full_name': '경기도', 'index': 9, 'icon': '🏘️'},
    {'id': 'region_gangwon', 'name': '강원', 'full_name': '강원특별자치도', 'index': 10, 'icon': '⛰️'},
    {'id': 'region_chungbuk', 'name': '충북', 'full_name': '충청북도', 'index': 11, 'icon': '🌾'},
    {'id': 'region_chungnam', 'name': '충남', 'full_name': '충청남도', 'index': 12, 'icon': '🌅'},
    {'id': 'region_jeonbuk', 'name': '전북', 'full_name': '전북특별자치도', 'index': 13, 'icon': '🌿'},
    {'id': 'region_jeonnam', 'name': '전남', 'full_name': '전라남도', 'index': 14, 'icon': '🍃'},
    {'id': 'region_gyeongbuk', 'name': '경북', 'full_name': '경상북도', 'index': 15, 'icon': '🏔️'},
    {'id': 'region_gyeongnam', 'name': '경남', 'full_name': '경상남도', 'index': 16, 'icon': '🌳'},
    {'id': 'region_jeju', 'name': '제주', 'full_name': '제주특별자치도', 'index': 17, 'icon': '🏝️'},
]

# ===== 페이지 수 설정 (목차 생성용) =====
# 주의: 목차를 생성하지 않으므로 이 설정은 더 이상 사용되지 않음
# 보존 목적으로만 유지 (향후 참고용)
PAGE_CONFIG = {
    # 페이지 번호 없는 섹션들 (표지, 일러두기, 목차, 인포그래픽)
    'pre_pages': 0,  # 이 섹션들은 페이지 번호가 없음
    
    # 요약 섹션 페이지 수 (1~5페이지)
    'summary': {
        'overview': 1,      # 요약-지역경제동향: 1페이지
        'production': 1,    # 요약-생산: 2페이지
        'consumption': 1,   # 요약-소비건설: 3페이지
        'trade_price': 1,   # 요약-수출물가: 4페이지
        'employment': 1,    # 요약-고용인구: 5페이지
    },
    
    # 부문별 섹션 페이지 수 (6~15페이지) - 정답 이미지 기준 각 1페이지
    # 목차 항목은 통합 표시: 생산(6), 소비(8), 건설(9), 수출입(10), 물가(12), 고용(13), 국내인구이동(15)
    'sector': {
        'manufacturing': 1,     # 광공업생산: 6페이지
        'service': 1,           # 서비스업생산: 7페이지
        'consumption': 1,       # 소비동향: 8페이지
        'construction': 1,      # 건설동향: 9페이지
        'export': 1,            # 수출: 10페이지
        'import': 1,            # 수입: 11페이지
        'price': 1,             # 물가동향: 12페이지
        'employment': 1,        # 고용률: 13페이지
        'unemployment': 1,      # 실업률: 14페이지
        'migration': 1,        # 국내인구이동: 15페이지
    },
    
    # 시도별 섹션 페이지 수 (16~49페이지) - 각 시도 2페이지
    'regional': 2,  # 각 시도별 페이지 수
    
    # 통계표 섹션 페이지 수 (52~페이지)
    # 주의: 통계표 목차는 더 이상 생성하지 않음
    'statistics': {
        'toc': 0,           # 통계표 목차 (생성하지 않음)
        'per_table': 1,     # 각 통계표당 페이지 수
        'count': 10,        # 통계표 개수 (광공업, 서비스업, 소매판매, 건설수주, 고용률, 실업률, 인구이동, 수출, 수입, 소비자물가)
    },
    
    # 부록 페이지 수
    'appendix': 1,
}

# ===== 목차용 항목 정의 (원본 이미지 기준) =====
# 부문별 7개 항목 (일부는 통합 표시)
TOC_SECTOR_ITEMS = [
    {'number': 1, 'name': '생산', 'start_from': 'manufacturing'},  # 광공업 시작 페이지
    {'number': 2, 'name': '소비', 'start_from': 'consumption'},
    {'number': 3, 'name': '건설', 'start_from': 'construction'},
    {'number': 4, 'name': '수출입', 'start_from': 'export'},       # 수출 시작 페이지
    {'number': 5, 'name': '물가', 'start_from': 'price'},
    {'number': 6, 'name': '고용', 'start_from': 'employment'},     # 고용률 시작 페이지
    {'number': 7, 'name': '국내 인구이동', 'start_from': 'migration'},
]

# 시도별 17개 항목 (원본 이미지 기준 - 띄어쓰기 없음)
TOC_REGION_ITEMS = [
    {'number': 1, 'name': '서울'},
    {'number': 2, 'name': '부산'},
    {'number': 3, 'name': '대구'},
    {'number': 4, 'name': '인천'},
    {'number': 5, 'name': '광주'},
    {'number': 6, 'name': '대전'},
    {'number': 7, 'name': '울산'},
    {'number': 8, 'name': '세종'},
    {'number': 9, 'name': '경기'},
    {'number': 10, 'name': '강원'},
    {'number': 11, 'name': '충북'},
    {'number': 12, 'name': '충남'},
    {'number': 13, 'name': '전북'},
    {'number': 14, 'name': '전남'},
    {'number': 15, 'name': '경북'},
    {'number': 16, 'name': '경남'},
    {'number': 17, 'name': '제주'},
]

