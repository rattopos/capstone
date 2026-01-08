#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
데이터 추출기 Facade

모든 보도자료 추출 기능을 하나의 인터페이스로 제공합니다.
기존 RawDataExtractor와 동일한 인터페이스를 유지하면서
내부적으로 모듈화된 추출기들을 사용합니다.
"""

from typing import Dict, Any, Optional
from pathlib import Path

from .base import BaseExtractor
from .production import ProductionExtractor
from .consumption import ConsumptionConstructionExtractor
from .trade import TradeExtractor
from .price import PriceExtractor
from .employment import EmploymentPopulationExtractor
from .config import ALL_REGIONS, RAW_SHEET_MAPPING, RAW_SHEET_QUARTER_COLS


class DataExtractor:
    """데이터 추출기 Facade
    
    모든 보도자료 데이터 추출을 위한 통합 인터페이스를 제공합니다.
    내부적으로 각 도메인별 추출기를 사용합니다.
    
    사용 예시:
        extractor = DataExtractor('path/to/excel.xlsx', 2025, 2)
        
        # 개별 보도자료 데이터 추출
        manufacturing = extractor.extract_mining_manufacturing_report_data()
        service = extractor.extract_service_industry_report_data()
        
        # 전체 보도자료 데이터 추출
        all_data = extractor.extract_all_report_data()
    """
    
    # 상수 노출 (하위 호환성)
    ALL_REGIONS = ALL_REGIONS
    RAW_SHEET_MAPPING = RAW_SHEET_MAPPING
    RAW_SHEET_QUARTER_COLS = RAW_SHEET_QUARTER_COLS
    
    def __init__(self, raw_excel_path: str, current_year: int, current_quarter: int):
        """
        Args:
            raw_excel_path: 기초자료 엑셀 파일 경로
            current_year: 현재 연도
            current_quarter: 현재 분기 (1-4)
        """
        self.raw_excel_path = Path(raw_excel_path)
        self.current_year = current_year
        self.current_quarter = current_quarter
        
        # 도메인별 추출기 초기화 (지연 로딩)
        self._production: Optional[ProductionExtractor] = None
        self._consumption: Optional[ConsumptionConstructionExtractor] = None
        self._trade: Optional[TradeExtractor] = None
        self._price: Optional[PriceExtractor] = None
        self._employment: Optional[EmploymentPopulationExtractor] = None
    
    # =========================================================================
    # 지연 로딩 프로퍼티
    # =========================================================================
    
    @property
    def production(self) -> ProductionExtractor:
        """생산 추출기 (광공업, 서비스업)"""
        if self._production is None:
            self._production = ProductionExtractor(
                str(self.raw_excel_path), self.current_year, self.current_quarter
            )
        return self._production
    
    @property
    def consumption(self) -> ConsumptionConstructionExtractor:
        """소비/건설 추출기"""
        if self._consumption is None:
            self._consumption = ConsumptionConstructionExtractor(
                str(self.raw_excel_path), self.current_year, self.current_quarter
            )
        return self._consumption
    
    @property
    def trade(self) -> TradeExtractor:
        """무역 추출기 (수출, 수입)"""
        if self._trade is None:
            self._trade = TradeExtractor(
                str(self.raw_excel_path), self.current_year, self.current_quarter
            )
        return self._trade
    
    @property
    def price(self) -> PriceExtractor:
        """물가 추출기"""
        if self._price is None:
            self._price = PriceExtractor(
                str(self.raw_excel_path), self.current_year, self.current_quarter
            )
        return self._price
    
    @property
    def employment(self) -> EmploymentPopulationExtractor:
        """고용/인구 추출기"""
        if self._employment is None:
            self._employment = EmploymentPopulationExtractor(
                str(self.raw_excel_path), self.current_year, self.current_quarter
            )
        return self._employment
    
    # =========================================================================
    # 보도자료 데이터 추출 메서드 (RawDataExtractor 호환 인터페이스)
    # =========================================================================
    
    def extract_mining_manufacturing_report_data(self) -> Dict[str, Any]:
        """광공업생산 보도자료 데이터 추출"""
        return self.production.extract_manufacturing_data()
    
    def extract_service_industry_report_data(self) -> Dict[str, Any]:
        """서비스업생산 보도자료 데이터 추출"""
        return self.production.extract_service_data()
    
    def extract_consumption_report_data(self) -> Dict[str, Any]:
        """소비동향 보도자료 데이터 추출"""
        return self.consumption.extract_consumption_data()
    
    def extract_construction_report_data(self) -> Dict[str, Any]:
        """건설동향 보도자료 데이터 추출"""
        return self.consumption.extract_construction_data()
    
    def extract_export_report_data(self) -> Dict[str, Any]:
        """수출 보도자료 데이터 추출"""
        return self.trade.extract_export_data()
    
    def extract_import_report_data(self) -> Dict[str, Any]:
        """수입 보도자료 데이터 추출"""
        return self.trade.extract_import_data()
    
    def extract_price_report_data(self) -> Dict[str, Any]:
        """물가동향 보도자료 데이터 추출"""
        return self.price.extract_price_data()
    
    def extract_employment_rate_report_data(self) -> Dict[str, Any]:
        """고용률 보도자료 데이터 추출"""
        return self.employment.extract_employment_rate_data()
    
    def extract_unemployment_report_data(self) -> Dict[str, Any]:
        """실업률 보도자료 데이터 추출"""
        return self.employment.extract_unemployment_data()
    
    def extract_population_migration_report_data(self) -> Dict[str, Any]:
        """국내인구이동 보도자료 데이터 추출"""
        return self.employment.extract_population_migration_data()
    
    # =========================================================================
    # 통합 메서드
    # =========================================================================
    
    def extract_all_report_data(self) -> Dict[str, Dict[str, Any]]:
        """모든 보도자료 데이터 추출"""
        return {
            'manufacturing': self.extract_mining_manufacturing_report_data(),
            'service': self.extract_service_industry_report_data(),
            'consumption': self.extract_consumption_report_data(),
            'construction': self.extract_construction_report_data(),
            'export': self.extract_export_report_data(),
            'import': self.extract_import_report_data(),
            'price': self.extract_price_report_data(),
            'employment': self.extract_employment_rate_report_data(),
            'unemployment': self.extract_unemployment_report_data(),
            'population': self.extract_population_migration_report_data(),
        }
    
    def extract_report_data(self, report_id: str) -> Optional[Dict[str, Any]]:
        """특정 보도자료 데이터 추출
        
        Args:
            report_id: 보도자료 ID
                - 'manufacturing': 광공업생산
                - 'service': 서비스업생산
                - 'consumption': 소비동향
                - 'construction': 건설동향
                - 'export': 수출
                - 'import': 수입
                - 'price': 물가동향
                - 'employment': 고용률
                - 'unemployment': 실업률
                - 'population': 국내인구이동
                
        Returns:
            보도자료 데이터 또는 None
        """
        extractors = {
            'manufacturing': self.extract_mining_manufacturing_report_data,
            'service': self.extract_service_industry_report_data,
            'consumption': self.extract_consumption_report_data,
            'construction': self.extract_construction_report_data,
            'export': self.extract_export_report_data,
            'import': self.extract_import_report_data,
            'price': self.extract_price_report_data,
            'employment': self.extract_employment_rate_report_data,
            'unemployment': self.extract_unemployment_report_data,
            'population': self.extract_population_migration_report_data,
        }
        
        extractor_func = extractors.get(report_id)
        if extractor_func:
            return extractor_func()
        return None


# 하위 호환성을 위한 별칭
RawDataExtractor = DataExtractor
