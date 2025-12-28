# -*- coding: utf-8 -*-
"""
보고서 설정 및 상수 정의
"""

# ===== 요약 보고서 목록 (표지-일러두기-목차-인포그래픽-요약 순서) =====
SUMMARY_REPORTS = [
    {
        'id': 'cover',
        'name': '표지',
        'sheet': None,
        'generator': None,
        'template': 'cover_template.html',
        'icon': '📑',
        'category': 'summary'
    },
    {
        'id': 'guide',
        'name': '일러두기',
        'sheet': None,
        'generator': None,
        'template': 'guide_template.html',
        'icon': '📖',
        'category': 'summary'
    },
    {
        'id': 'toc',
        'name': '목차',
        'sheet': None,
        'generator': None,
        'template': 'toc_template.html',
        'icon': '📋',
        'category': 'summary'
    },
    {
        'id': 'infographic',
        'name': '인포그래픽',
        'sheet': 'multiple',
        'generator': 'infographic_generator.py',
        'template': 'infographic_js_template.html',
        'icon': '📊',
        'category': 'summary'
    },
    {
        'id': 'summary_overview',
        'name': '요약-지역경제동향',
        'sheet': 'multiple',
        'generator': 'summary_regional_economy_generator.py',
        'template': 'summary_regional_economy_template.html',
        'icon': '📈',
        'category': 'summary'
    },
    {
        'id': 'summary_production',
        'name': '요약-생산',
        'sheet': 'multiple',
        'generator': 'summary_production_generator.py',
        'template': 'summary_production_template.html',
        'icon': '🏭',
        'category': 'summary'
    },
    {
        'id': 'summary_consumption',
        'name': '요약-소비건설',
        'sheet': 'multiple',
        'generator': 'summary_consumption_construction_generator.py',
        'template': 'summary_consumption_construction_template.html',
        'icon': '🛒',
        'category': 'summary'
    },
    {
        'id': 'summary_trade_price',
        'name': '요약-수출물가',
        'sheet': 'multiple',
        'generator': 'summary_export_price_generator.py',
        'template': 'summary_export_price_template.html',
        'icon': '📦',
        'category': 'summary'
    },
    {
        'id': 'summary_employment',
        'name': '요약-고용인구',
        'sheet': 'multiple',
        'generator': 'summary_employment_generator.py',
        'template': 'summary_employment_template.html',
        'icon': '👔',
        'category': 'summary'
    },
]

# ===== 부문별 보고서 순서 설정 =====
SECTOR_REPORTS = [
    {
        'id': 'manufacturing',
        'name': '광공업생산',
        'sheet': 'A 분석',
        'generator': 'mining_manufacturing_generator.py',
        'template': 'mining_manufacturing_template.html',
        'icon': '🏭',
        'category': 'production'
    },
    {
        'id': 'service',
        'name': '서비스업생산',
        'sheet': 'B 분석',
        'generator': 'service_industry_generator.py',
        'template': 'service_industry_template.html',
        'icon': '🏢',
        'category': 'production'
    },
    {
        'id': 'consumption',
        'name': '소비동향',
        'sheet': 'C 분석',
        'generator': 'consumption_generator.py',
        'template': 'consumption_template.html',
        'icon': '🛒',
        'category': 'consumption'
    },
    {
        'id': 'construction',
        'name': '건설동향',
        'sheet': "F'분석",
        'generator': 'construction_generator.py',
        'template': 'construction_template.html',
        'icon': '🏗️',
        'category': 'construction'
    },
    {
        'id': 'export',
        'name': '수출',
        'sheet': 'G 분석',
        'generator': 'export_generator.py',
        'template': 'export_template.html',
        'icon': '📦',
        'category': 'trade'
    },
    {
        'id': 'import',
        'name': '수입',
        'sheet': 'H 분석',
        'generator': 'import_generator.py',
        'template': 'import_template.html',
        'icon': '🚢',
        'category': 'trade'
    },
    {
        'id': 'price',
        'name': '물가동향',
        'sheet': 'E(품목성질물가)분석',
        'generator': 'price_trend_generator.py',
        'template': 'price_trend_template.html',
        'icon': '💰',
        'category': 'price'
    },
    {
        'id': 'employment',
        'name': '고용률',
        'sheet': 'D(고용률)분석',
        'generator': 'employment_rate_generator.py',
        'template': 'employment_rate_template.html',
        'icon': '👔',
        'category': 'employment'
    },
    {
        'id': 'unemployment',
        'name': '실업률',
        'sheet': 'D(실업)분석',
        'generator': 'unemployment_generator.py',
        'template': 'unemployment_template.html',
        'icon': '📉',
        'category': 'employment'
    },
    {
        'id': 'population',
        'name': '국내인구이동',
        'sheet': 'I(순인구이동)집계',
        'generator': 'domestic_migration_generator.py',
        'template': 'domestic_migration_template.html',
        'icon': '👥',
        'category': 'population'
    }
]

# 전체 보고서 순서 (요약 → 부문별)
REPORT_ORDER = SUMMARY_REPORTS + SECTOR_REPORTS

# ===== 통계표 보고서 목록 =====
STATISTICS_REPORTS = [
    {
        'id': 'stat_toc',
        'name': '통계표-목차',
        'table_name': None,
        'template': 'statistics_table_toc_template.html',
        'icon': '📋',
        'category': 'statistics'
    },
    {
        'id': 'stat_mining',
        'name': '통계표-광공업생산지수',
        'table_name': '광공업생산지수',
        'template': 'statistics_table_index_template.html',
        'icon': '🏭',
        'category': 'statistics'
    },
    {
        'id': 'stat_service',
        'name': '통계표-서비스업생산지수',
        'table_name': '서비스업생산지수',
        'template': 'statistics_table_index_template.html',
        'icon': '🏢',
        'category': 'statistics'
    },
    {
        'id': 'stat_retail',
        'name': '통계표-소매판매액지수',
        'table_name': '소매판매액지수',
        'template': 'statistics_table_index_template.html',
        'icon': '🛒',
        'category': 'statistics'
    },
    {
        'id': 'stat_construction',
        'name': '통계표-건설수주액',
        'table_name': '건설수주액',
        'template': 'statistics_table_index_template.html',
        'icon': '🏗️',
        'category': 'statistics'
    },
    {
        'id': 'stat_employment',
        'name': '통계표-고용률',
        'table_name': '고용률',
        'template': 'statistics_table_index_template.html',
        'icon': '👔',
        'category': 'statistics'
    },
    {
        'id': 'stat_unemployment',
        'name': '통계표-실업률',
        'table_name': '실업률',
        'template': 'statistics_table_index_template.html',
        'icon': '📉',
        'category': 'statistics'
    },
    {
        'id': 'stat_population',
        'name': '통계표-국내인구이동',
        'table_name': '국내인구이동',
        'template': 'statistics_table_index_template.html',
        'icon': '👥',
        'category': 'statistics'
    },
    {
        'id': 'stat_export',
        'name': '통계표-수출액',
        'table_name': '수출액',
        'template': 'statistics_table_index_template.html',
        'icon': '📦',
        'category': 'statistics'
    },
    {
        'id': 'stat_import',
        'name': '통계표-수입액',
        'table_name': '수입액',
        'template': 'statistics_table_index_template.html',
        'icon': '🚢',
        'category': 'statistics'
    },
    {
        'id': 'stat_price',
        'name': '통계표-소비자물가지수',
        'table_name': '소비자물가지수',
        'template': 'statistics_table_index_template.html',
        'icon': '💰',
        'category': 'statistics'
    },
    {
        'id': 'stat_grdp',
        'name': '통계표-참고-GRDP',
        'table_name': 'GRDP',
        'template': 'statistics_table_grdp_template.html',
        'icon': '📊',
        'category': 'statistics'
    },
    {
        'id': 'stat_appendix',
        'name': '부록-주요용어정의',
        'table_name': None,
        'template': 'statistics_table_appendix_template.html',
        'icon': '📖',
        'category': 'statistics'
    },
]

# 시도별 보고서 목록 (17개 시도 + 참고_GRDP)
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
    {'id': 'reference_grdp', 'name': '참고_GRDP', 'full_name': '분기 지역내총생산(GRDP)', 'index': 18, 'icon': '📊', 'is_reference': True},
]

