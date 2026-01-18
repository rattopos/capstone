# DEPRECATED: config/reports.py를 사용하세요.
# 기존 REPORT_CONFIGS, get_report_config 등은 reports.py로 통합되었습니다.

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
보고서 설정 파일

각 부문별 보고서의 차이점(시트명, 매핑, 템플릿 등)을 정의합니다.
통합 Generator는 이 설정을 기반으로 작동합니다.
"""

# 부문별 보고서 설정
REPORT_CONFIGS = {
    'mining': {
        'name': '광공업생산',
        'report_id': 'mining',
        'header_row': 1,  # 분석시트의 헤더 행 (0부터 시작)
        'sheets': {
            'analysis': ['A 분석', 'A분석'],
            'aggregation': ['A(광공업생산)집계', 'A 집계'],
            'fallback': ['광공업생산', '광공업생산지수']
        },
        'aggregation_structure': {
            'region_name_col': 4,
            'industry_code_col': 7,
            'total_code': 'BCD'
        },
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
        'template': 'by_type/mining_template.html',
        'metadata_columns': {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계'],
            'code': ['산업코드', 'code', '코드'],
            'name': ['산업명', '산업이름', 'industry']
        },
        'index_name': '생산지수',
        'item_name': '업종'  # "업종" vs "업태"
    },
    
    'service': {
        'name': '서비스업생산',
        'report_id': 'service',
        'sheets': {
            'analysis': ['B 분석', 'B분석'],
            'aggregation': ['B(서비스업생산)집계', 'B 집계'],
            'fallback': ['서비스업생산', '서비스업생산지수']
        },
        'aggregation_structure': {
            'region_name_col': 3,
            'industry_code_col': 6,
            'total_code': 'E~S'
        },
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
        'template': 'by_type/service_template.html',
        'metadata_columns': {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계'],
            'code': ['산업코드', 'code', '코드'],
            'name': ['산업명', '산업이름', 'industry']
        },
        'index_name': '생산지수',
        'item_name': '업종'
    },
    
    'consumption': {
        'name': '소비동향',
        'report_id': 'consumption',
        'sheets': {
            'analysis': ['C 분석', 'C분석'],
            'aggregation': ['C(소비)집계', 'C 집계'],
            'fallback': ['소비', '소매판매액지수']
        },
        'aggregation_structure': {
            'region_name_col': 2,
            'industry_code_col': 5,
            'total_code': 'A0'
        },
        'name_mapping': {
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
        },
        'template': 'by_type/consumption_template.html',
        'metadata_columns': {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계'],
            'code': ['업태코드', 'code', '코드'],
            'name': ['업태명', '업태종류', 'business']
        },
        'index_name': '판매지수',
        'item_name': '업태'
    },
    
    'construction': {
        'name': '건설수주',
        'report_id': 'construction',
        'sheets': {
            'analysis': ["F'분석", "F' 분석"],
            'aggregation': ["F'(건설)집계", "F' 집계"],
            'fallback': ['건설', '건설수주']
        },
        'aggregation_structure': {
            'region_name_col': 1,
            'industry_code_col': 3,
            'total_code': '0'
        },
        'name_mapping': {
            '건축': '건축',
            '토목': '토목',
            '주거용 건물': '주거용',
            '비주거용 건물': '비주거용',
        },
        'template': 'by_type/construction_template.html',
        'metadata_columns': {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계'],
            'code': ['분류코드', 'code', '코드'],
            'name': ['공정이름', '공정명', 'construction']
        },
        'index_name': '수주액',
        'item_name': '공종'
    },
    
    'export': {
        'name': '수출',
        'report_id': 'export',
        'sheets': {
            'analysis': ['G 분석', 'G분석'],
            'aggregation': ['G(수출)집계', 'G 집계'],
            'fallback': ['수출', '수출액']
        },
        'aggregation_structure': {
            'region_name_col': 3,
            'industry_code_col': 6,
            'total_code': '합계'
        },
        'name_mapping': {},  # 품목명은 그대로 사용
        'template': 'by_type/export_template.html',
        'metadata_columns': {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계'],
            'code': ['품목코드', 'code', '코드'],
            'name': ['품목명', '품목', 'item']
        },
        'index_name': '수출액',
        'item_name': '품목'
    },
    
    'import': {
        'name': '수입',
        'report_id': 'import',
        'sheets': {
            'analysis': ['H 분석', 'H분석'],
            'aggregation': ['H(수입)집계', 'H 집계'],
            'fallback': ['수입', '수입액']
        },
        'aggregation_structure': {
            'region_name_col': 3,
            'industry_code_col': 6,
            'total_code': '합계'
        },
        'name_mapping': {},
        'template': 'by_type/import_template.html',
        'metadata_columns': {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계'],
            'code': ['품목코드', 'code', '코드'],
            'name': ['품목명', '품목', 'item']
        },
        'index_name': '수입액',
        'item_name': '품목'
    },
    
    'price': {
        'name': '물가동향',
        'report_id': 'price',
        'sheets': {
            'analysis': ['E(지출목적물가) 분석', 'E 분석'],
            'aggregation': ['E(지출목적물가)집계', 'E 집계'],
            'fallback': ['물가', '소비자물가']
        },
        'aggregation_structure': {
            'region_name_col': 2,
            'industry_code_col': 5,
            'total_code': '00'
        },
        'name_mapping': {},
        'template': 'by_type/price_template.html',
        'metadata_columns': {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계'],
            'code': ['품목코드', 'code', '코드'],
            'name': ['품목명', '품목 이름', '품목', 'item', '지출목적']
        },
        'index_name': '물가지수',
        'item_name': '품목'
    },
    
    'employment': {
        'name': '고용률',
        'report_id': 'employment',
        'sheets': {
            'analysis': ['D(고용률)분석', 'D 분석'],
            'aggregation': ['D(고용률)집계', 'D 집계'],
            'fallback': ['고용', '고용률']
        },
        'aggregation_structure': {
            'region_name_col': 1,
            'industry_code_col': 3,
            'total_code': '계'
        },
        'name_mapping': {},
        'template': 'by_type/employment_template.html',
        'metadata_columns': {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계'],
            'code': ['산업코드', 'code', '코드'],
            'name': ['산업명', '산업 이름', 'industry']
        },
        'index_name': '고용률',
        'item_name': '업종'
    },
    
    'unemployment': {
        'name': '실업률',
        'report_id': 'unemployment',
        'sheets': {
            'analysis': ['D(실업)분석', 'D 분석'],
            'aggregation': ['D(실업)집계', 'D 집계'],
            'fallback': ['실업', '실업률']
        },
        'aggregation_structure': {
            'region_name_col': 0,
            'industry_code_col': 1,
            'total_code': '계'
        },
        'name_mapping': {},
        'template': 'by_type/unemployment_template.html',
        'metadata_columns': {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계'],
            'code': ['산업코드', 'code', '코드'],
            'name': ['산업명', '산업 이름', 'industry']
        },
        'index_name': '실업률',
        'item_name': '업종'
    },
    
    'migration': {
        'name': '국내인구이동',
        'report_id': 'migration',
        'sheets': {
            'analysis': ['I(순인구이동)분석', 'I 분석'],
            'aggregation': ['I(순인구이동)집계', 'I 집계'],
            'fallback': ['인구이동', '순인구이동']
        },
        'aggregation_structure': {
            'region_name_col': 4,           # 지역이름 컬럼 (서울, 부산, ...)
            'industry_code_col': 3,          # 지역코드 컬럼 (11, 21, 22, ...)
            'total_code': '순인구이동 수'   # 이동 유형: 유입/유출/순인구이동 중 '순인구이동 수' 선택
        },
        'name_mapping': {},
        'template': 'by_type/migration_template.html',
        'metadata_columns': {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계', '유형'],  # 이동 유형 (유입/유출/순이동)
            'code': ['지역코드', 'code', '코드'],  # 지역코드
            'name': ['분류', '유형', '단계']  # 이동 유형 (유입인구 수/유출인구 수/순인구이동 수)
        },
        'index_name': '순인구이동',
        'item_name': '연령대',  # 연령대별 분류
        'has_nationwide': False,  # 국내인구이동은 전국 통계 없음
        'require_analysis_sheet': False  # 집계시트만 사용
    },
    
    'regional_economy_by_region': {
        'name': '시도별 경제동향',
        'report_id': 'regional_economy_by_region',
        'sheets': {
            'analysis': None,  # 분석시트 없음, 집계시트만 사용
            'aggregation': [
                'A(광공업생산)집계', 'B(서비스업생산)집계', 'C(소비)집계',
                "F'(건설)집계", 'G(수출)집계', 'H(수입)집계', 
                'D(고용률)집계', 'E(지출목적물가)집계', 'I(순인구이동)집계'
            ],
            'fallback': []
        },
        'aggregation_structure': {
            'region_name_col': None,  # 동적으로 찾음
            'industry_code_col': None,  # 동적으로 찾음
            'total_code': None  # 동적으로 찾음
        },
        'name_mapping': {},
        'template': 'regional_economy_by_region_template.html',
        'metadata_columns': {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계'],
            'code': ['코드', 'code'],
            'name': ['산업명', '품목명', '분류', '유형', 'item', 'industry']
        },
        'index_name': '주요지표',
        'item_name': '부문',
        'has_nationwide': False,  # 시도별이므로 전국 통계 없음
        'require_analysis_sheet': False,  # 집계시트만 사용
        'is_regional_by_region': True  # 시도별 경제동향 플래그
    },
}

# 지역 표시명 (모든 부문 공통)
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

# 지역 그룹 (모든 부문 공통)
REGION_GROUPS = {
    '경인': ['서울', '인천', '경기'],
    '충청': ['대전', '세종', '충북', '충남'],
    '호남': ['광주', '전북', '전남', '제주'],
    '동북': ['대구', '경북', '강원'],
    '동남': ['부산', '울산', '경남']
}

# 유효한 지역 목록 (모든 부문 공통)
VALID_REGIONS = [
    '전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
    '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주'
]


def get_report_config(report_type: str) -> dict:
    """보고서 설정 가져오기"""
    if report_type not in REPORT_CONFIGS:
        available = ', '.join(REPORT_CONFIGS.keys())
        raise ValueError(f"Unknown report type: '{report_type}'. Available: {available}")
    return REPORT_CONFIGS[report_type]


def list_available_reports() -> list:
    """사용 가능한 보고서 목록"""
    return [
        {
            'type': key,
            'name': config['name'],
            'report_id': config['report_id']
        }
        for key, config in REPORT_CONFIGS.items()
    ]
