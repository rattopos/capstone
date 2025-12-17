"""
데이터 모델 기반 템플릿 필러
객체 기반 매핑 시스템
"""

import re
from typing import Any, Dict, Optional
from .template_manager import TemplateManager
from .data_model import DocumentData, SheetData, RegionData
from .excel_to_model import ExcelToModelConverter
from .calculator import Calculator


class ModelBasedFiller:
    """데이터 모델 기반 템플릿 필러"""
    
    def __init__(self, template_manager: TemplateManager, excel_extractor):
        """
        모델 기반 필러 초기화
        
        Args:
            template_manager: 템플릿 관리자
            excel_extractor: 엑셀 추출기
        """
        self.template_manager = template_manager
        self.excel_extractor = excel_extractor
        self.converter = ExcelToModelConverter(excel_extractor)
        self.calculator = Calculator()
        self.document_data: Optional[DocumentData] = None
    
    def fill_template(self, year: int, quarter: int, 
                     sheet_names: Optional[list] = None) -> str:
        """
        템플릿을 데이터로 채우기
        
        Args:
            year: 연도
            quarter: 분기
            sheet_names: 사용할 시트 목록 (None이면 모든 시트)
            
        Returns:
            채워진 HTML 템플릿
        """
        # 엑셀 데이터를 모델로 변환
        self.document_data = self.converter.convert_document(year, quarter, sheet_names)
        
        # 템플릿 로드
        template_content = self.template_manager.template_content
        
        # 마커 찾기 및 치환
        markers = self.template_manager.parse_markers()
        
        for marker in markers:
            sheet_name = marker['sheet_name']
            key = marker['key']
            operation = marker.get('operation')
            
            # 값 추출
            value = self._get_value(sheet_name, key, year, quarter, operation)
            
            # 포맷팅
            formatted_value = self._format_value(value, operation)
            
            # 마커 치환
            marker_pattern = re.escape(marker['full_match'])
            template_content = re.sub(marker_pattern, formatted_value, template_content)
        
        return template_content
    
    def _get_value(self, sheet_name: str, key: str, year: int, quarter: int,
                   operation: Optional[str] = None) -> Any:
        """
        마커 키에서 값 추출
        
        Args:
            sheet_name: 시트명
            key: 마커 키 (예: "전국_증감률", "서울_2025_2")
            year: 연도
            quarter: 분기
            operation: 연산 (예: "증감률", "합계")
            
        Returns:
            값
        """
        if not self.document_data:
            return None
        
        # 특수 키 처리
        if sheet_name == '요약':
            return self._get_summary_value(key, year, quarter)
        
        # 시트 데이터 가져오기
        sheet_data = self.document_data.get_sheet(sheet_name)
        if not sheet_data:
            return None
        
        # 키 파싱
        parts = key.split('_')
        
        # 지역명 추출
        region_name = None
        if len(parts) > 0:
            # 첫 번째 부분이 지역명일 가능성
            potential_region = parts[0]
            region = sheet_data.get_region(potential_region)
            if region:
                region_name = potential_region
                parts = parts[1:]  # 지역명 제거
        
        # 연도/분기 추출
        period_key = f"{year}_{quarter}"
        
        # 키 패턴에 따른 처리
        if len(parts) == 0:
            # 단순 지역 값
            if region_name:
                return sheet_data.get_value(region_name, year, quarter)
        
        elif len(parts) == 1:
            # "증감률", "값" 등
            if parts[0] == '증감률' or parts[0] == 'growth_rate':
                if region_name:
                    return sheet_data.get_growth_rate(region_name, year, quarter)
            elif parts[0] in ['값', 'value']:
                if region_name:
                    return sheet_data.get_value(region_name, year, quarter)
            else:
                # 카테고리명일 가능성
                if region_name:
                    return sheet_data.get_value(region_name, year, quarter, parts[0])
        
        elif len(parts) == 2:
            # "2025_2", "반도체_증감률" 등
            try:
                # 연도_분기 형식
                key_year = int(parts[0])
                key_quarter = int(parts[1])
                if region_name:
                    return sheet_data.get_value(region_name, key_year, key_quarter)
            except ValueError:
                # 카테고리_연산 형식
                category = parts[0]
                op = parts[1]
                if op == '증감률' and region_name:
                    return sheet_data.get_growth_rate(region_name, year, quarter, category)
                elif region_name:
                    return sheet_data.get_value(region_name, year, quarter, category)
        
        elif len(parts) >= 3:
            # "서울_반도체_증감률", "전국_2025_2" 등
            # 첫 번째가 지역명이 아닐 수도 있음
            if region_name is None:
                # 다시 지역명 찾기 시도
                potential_region = parts[0]
                region = sheet_data.get_region(potential_region)
                if region:
                    region_name = potential_region
                    parts = parts[1:]
            
            if len(parts) >= 2:
                try:
                    # 연도_분기 형식
                    key_year = int(parts[0])
                    key_quarter = int(parts[1])
                    if region_name:
                        return sheet_data.get_value(region_name, key_year, key_quarter)
                except ValueError:
                    # 카테고리_연산 형식
                    category = '_'.join(parts[:-1])
                    op = parts[-1]
                    if op == '증감률' and region_name:
                        return sheet_data.get_growth_rate(region_name, year, quarter, category)
                    elif region_name:
                        return sheet_data.get_value(region_name, year, quarter, category)
        
        return None
    
    def _get_summary_value(self, key: str, year: int, quarter: int) -> Any:
        """요약 섹션 값 가져오기"""
        if not self.document_data:
            return None
        
        summary = self.document_data.get_summary_data()
        
        # 키 파싱
        parts = key.split('_')
        
        if len(parts) == 1:
            if parts[0] == '연도':
                return year
            elif parts[0] == '분기':
                return quarter
        
        elif len(parts) == 2:
            category = parts[0]  # "광공업생산", "서비스업생산" 등
            metric = parts[1]  # "증감률", "값" 등
            
            if category in summary:
                category_data = summary[category]
                if isinstance(category_data, dict):
                    # 첫 번째 항목의 값 반환
                    for item_key, item_value in category_data.items():
                        if isinstance(item_value, dict) and metric in item_value:
                            return item_value[metric]
        
        return None
    
    def _format_value(self, value: Any, operation: Optional[str] = None) -> str:
        """
        값 포맷팅
        
        Args:
            value: 값
            operation: 연산 타입
            
        Returns:
            포맷팅된 문자열
        """
        if value is None:
            return "N/A"
        
        # 숫자 포맷팅
        try:
            num_value = float(value)
            
            # 증감률/퍼센트
            if operation in ['증감률', 'growth_rate', 'percentage']:
                return f"{num_value:.1f}%"
            
            # 일반 숫자
            if abs(num_value) >= 1000:
                return f"{num_value:,.1f}"
            else:
                return f"{num_value:.1f}"
        
        except (ValueError, TypeError):
            return str(value) if value else "N/A"

