#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
데이터 추출기 패키지

기초자료에서 보도자료용 데이터를 추출하는 모듈화된 추출기들을 제공합니다.

사용 예시:
    from extractors import DataExtractor
    
    extractor = DataExtractor('path/to/excel.xlsx', 2025, 2)
    manufacturing_data = extractor.extract_manufacturing_report_data()
    service_data = extractor.extract_service_industry_report_data()
"""

from .config import (
    ALL_REGIONS,
    REGION_GROUPS,
    RAW_SHEET_MAPPING,
    RAW_SHEET_QUARTER_COLS,
    REPORT_CONFIG,
)

from .base import BaseExtractor
from .production import ProductionExtractor
from .consumption import ConsumptionConstructionExtractor
from .trade import TradeExtractor
from .price import PriceExtractor
from .employment import EmploymentPopulationExtractor
from .facade import DataExtractor

__all__ = [
    # 설정
    'ALL_REGIONS',
    'REGION_GROUPS',
    'RAW_SHEET_MAPPING',
    'RAW_SHEET_QUARTER_COLS',
    'REPORT_CONFIG',
    # 추출기
    'BaseExtractor',
    'ProductionExtractor',
    'ConsumptionConstructionExtractor',
    'TradeExtractor',
    'PriceExtractor',
    'EmploymentPopulationExtractor',
    'DataExtractor',  # Facade - 메인 진입점
]
