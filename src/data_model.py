"""
데이터 모델 모듈
엑셀 데이터를 구조화된 객체로 변환
"""

from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class RegionData:
    """지역별 데이터"""
    region_code: str
    region_name: str
    values: Dict[str, Any] = field(default_factory=dict)  # 연도/분기별 값
    categories: Dict[str, Any] = field(default_factory=dict)  # 카테고리별 데이터


@dataclass
class CategoryData:
    """카테고리별 데이터 (산업, 업태 등)"""
    category_code: str
    category_name: str
    values: Dict[str, Any] = field(default_factory=dict)
    subcategories: List['CategoryData'] = field(default_factory=list)


@dataclass
class SheetData:
    """시트 전체 데이터"""
    sheet_name: str
    title: str
    header_row: int = 3
    regions: Dict[str, RegionData] = field(default_factory=dict)  # 지역코드 -> RegionData
    categories: Dict[str, CategoryData] = field(default_factory=dict)  # 카테고리코드 -> CategoryData
    year_columns: Dict[int, int] = field(default_factory=dict)  # 연도 -> 컬럼 인덱스
    quarter_columns: Dict[str, int] = field(default_factory=dict)  # "2025_2" -> 컬럼 인덱스
    raw_data: List[List[Any]] = field(default_factory=list)  # 원본 데이터
    
    def get_region(self, region_name: str) -> Optional[RegionData]:
        """지역명으로 RegionData 찾기"""
        for region in self.regions.values():
            if region.region_name == region_name:
                return region
        return None
    
    def get_value(self, region_name: str, year: int, quarter: int, 
                  category_name: Optional[str] = None) -> Optional[Any]:
        """특정 지역, 연도, 분기의 값 가져오기"""
        region = self.get_region(region_name)
        if not region:
            return None
        
        period_key = f"{year}_{quarter}"
        if period_key in region.values:
            return region.values[period_key]
        
        # 카테고리별 값
        if category_name and category_name in region.categories:
            return region.categories[category_name].values.get(period_key)
        
        return None
    
    def get_growth_rate(self, region_name: str, year: int, quarter: int,
                       category_name: Optional[str] = None) -> Optional[float]:
        """전년 동분기 대비 증감률 계산"""
        current_value = self.get_value(region_name, year, quarter, category_name)
        if current_value is None:
            return None
        
        prev_year = year - 1
        prev_value = self.get_value(region_name, prev_year, quarter, category_name)
        
        if prev_value is None or prev_value == 0:
            return None
        
        try:
            current = float(current_value)
            prev = float(prev_value)
            if prev == 0:
                return None
            return ((current - prev) / prev) * 100
        except (ValueError, TypeError):
            return None


@dataclass
class DocumentData:
    """전체 문서 데이터"""
    year: int
    quarter: int
    sheets: Dict[str, SheetData] = field(default_factory=dict)  # 시트명 -> SheetData
    
    def get_sheet(self, sheet_name: str) -> Optional[SheetData]:
        """시트명으로 SheetData 찾기"""
        return self.sheets.get(sheet_name)
    
    def get_summary_data(self) -> Dict[str, Any]:
        """요약 데이터 생성"""
        summary = {
            'year': self.year,
            'quarter': self.quarter,
            'production': {},
            'consumption': {},
            'construction': {},
            'exports': {},
            'imports': {},
            'employment': {},
            'price': {},
            'population': {}
        }
        
        # 생산 데이터
        if '광공업생산' in self.sheets:
            sheet = self.sheets['광공업생산']
            national = sheet.get_region('전국')
            if national:
                period_key = f"{self.year}_{self.quarter}"
                summary['production']['mining_manufacturing'] = {
                    'value': national.values.get(period_key),
                    'growth_rate': sheet.get_growth_rate('전국', self.year, self.quarter)
                }
        
        if '서비스업생산' in self.sheets:
            sheet = self.sheets['서비스업생산']
            national = sheet.get_region('전국')
            if national:
                period_key = f"{self.year}_{self.quarter}"
                summary['production']['service'] = {
                    'value': national.values.get(period_key),
                    'growth_rate': sheet.get_growth_rate('전국', self.year, self.quarter)
                }
        
        # 소비 데이터
        if '소비(소매, 추가)' in self.sheets:
            sheet = self.sheets['소비(소매, 추가)']
            national = sheet.get_region('전국')
            if national:
                period_key = f"{self.year}_{self.quarter}"
                summary['consumption']['retail'] = {
                    'value': national.values.get(period_key),
                    'growth_rate': sheet.get_growth_rate('전국', self.year, self.quarter)
                }
        
        # 건설 데이터
        if '건설 (공표자료)' in self.sheets:
            sheet = self.sheets['건설 (공표자료)']
            national = sheet.get_region('전국')
            if national:
                period_key = f"{self.year}_{self.quarter}"
                summary['construction']['orders'] = {
                    'value': national.values.get(period_key),
                    'growth_rate': sheet.get_growth_rate('전국', self.year, self.quarter)
                }
        
        # 수출입 데이터
        if '수출' in self.sheets:
            sheet = self.sheets['수출']
            national = sheet.get_region('전국')
            if national:
                period_key = f"{self.year}_{self.quarter}"
                summary['exports']['total'] = {
                    'value': national.values.get(period_key),
                    'growth_rate': sheet.get_growth_rate('전국', self.year, self.quarter)
                }
        
        if '수입' in self.sheets:
            sheet = self.sheets['수입']
            national = sheet.get_region('전국')
            if national:
                period_key = f"{self.year}_{self.quarter}"
                summary['imports']['total'] = {
                    'value': national.values.get(period_key),
                    'growth_rate': sheet.get_growth_rate('전국', self.year, self.quarter)
                }
        
        return summary
    
    def get_region_data(self, region_name: str) -> Dict[str, Any]:
        """특정 시도의 종합 데이터"""
        region_data = {
            'region_name': region_name,
            'year': self.year,
            'quarter': self.quarter,
            'production': {},
            'consumption': {},
            'construction': {},
            'exports': {},
            'imports': {},
            'employment': {},
            'price': {},
            'population': {}
        }
        
        # 각 시트에서 해당 지역 데이터 추출
        for sheet_name, sheet in self.sheets.items():
            region = sheet.get_region(region_name)
            if region:
                period_key = f"{self.year}_{self.quarter}"
                value = region.values.get(period_key)
                growth_rate = sheet.get_growth_rate(region_name, self.year, self.quarter)
                
                if '생산' in sheet_name:
                    region_data['production'][sheet_name] = {
                        'value': value,
                        'growth_rate': growth_rate
                    }
                elif '소비' in sheet_name or '소매' in sheet_name:
                    region_data['consumption'][sheet_name] = {
                        'value': value,
                        'growth_rate': growth_rate
                    }
                elif '건설' in sheet_name:
                    region_data['construction'][sheet_name] = {
                        'value': value,
                        'growth_rate': growth_rate
                    }
                elif '수출' in sheet_name:
                    region_data['exports'][sheet_name] = {
                        'value': value,
                        'growth_rate': growth_rate
                    }
                elif '수입' in sheet_name:
                    region_data['imports'][sheet_name] = {
                        'value': value,
                        'growth_rate': growth_rate
                    }
                elif '고용' in sheet_name:
                    region_data['employment'][sheet_name] = {
                        'value': value,
                        'growth_rate': growth_rate
                    }
        
        return region_data

