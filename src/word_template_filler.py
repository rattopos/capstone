"""
Word 템플릿 채우기 모듈
마커를 값으로 치환하고 포맷팅 처리
"""

import html
import re
import math
from typing import Any, Dict, Optional
from .word_template_manager import WordTemplateManager
from .excel_extractor import ExcelExtractor
from .calculator import Calculator
from .data_analyzer import DataAnalyzer
from .period_detector import PeriodDetector
from .flexible_mapper import FlexibleMapper
from .sheet_resolver import SheetResolver
from .template_filler import TemplateFiller  # 기존 로직 재사용


class WordTemplateFiller:
    """Word 템플릿에 데이터를 채우는 클래스"""
    
    def __init__(self, template_manager: WordTemplateManager, excel_extractor: ExcelExtractor):
        """
        Word 템플릿 필러 초기화
        
        Args:
            template_manager: Word 템플릿 관리자 인스턴스
            excel_extractor: 엑셀 추출기 인스턴스
        """
        self.template_manager = template_manager
        self.excel_extractor = excel_extractor
        
        # 기존 HTML 템플릿 필러를 래핑하여 로직 재사용
        # 임시로 HTML 템플릿 매니저를 생성 (실제로는 사용하지 않음)
        from .template_manager import TemplateManager
        temp_html_manager = TemplateManager("")  # 빈 경로
        self.html_filler = TemplateFiller(temp_html_manager, excel_extractor)
        
        self.calculator = Calculator()
        self.data_analyzer = DataAnalyzer(excel_extractor)
        self.period_detector = PeriodDetector(excel_extractor)
        self.flexible_mapper = FlexibleMapper(excel_extractor)
        self.sheet_resolver = None
        
        self._current_year = None
        self._current_quarter = None
        self._current_sheet_name = None
    
    def format_number(self, value: Any, use_comma: bool = True, decimal_places: int = 1) -> str:
        """숫자를 포맷팅합니다."""
        return self.html_filler.format_number(value, use_comma, decimal_places)
    
    def format_percentage(self, value: Any, decimal_places: int = 1, include_percent: bool = True) -> str:
        """퍼센트 값을 포맷팅합니다."""
        return self.html_filler.format_percentage(value, decimal_places, include_percent)
    
    def process_marker(self, marker_info: Dict[str, str]) -> str:
        """
        마커 정보를 처리하여 값을 추출하고 계산합니다.
        
        Args:
            marker_info: 마커 정보 딕셔너리
            
        Returns:
            처리된 값 (문자열)
        """
        # 기존 HTML 템플릿 필러의 로직 재사용
        marker_sheet_name = marker_info['sheet_name']
        cell_address = marker_info['cell_address']
        operation = marker_info.get('operation')
        
        # "완료체크" 시트명인 경우 템플릿 내용 기반으로 올바른 시트명 찾기
        if marker_sheet_name == '완료체크':
            if self.sheet_resolver is None:
                available_sheets = self.excel_extractor.get_sheet_names()
                self.sheet_resolver = SheetResolver(available_sheets)
            
            full_marker = marker_info.get('full_match', '')
            # Word 템플릿에서 텍스트 추출
            template_text = ""
            for para in self.template_manager.document.paragraphs:
                template_text += para.text + "\n"
            
            resolved_sheet = self.sheet_resolver.resolve_marker_in_template(
                template_text,
                full_marker
            )
            if resolved_sheet:
                marker_sheet_name = resolved_sheet
            else:
                marker_sheet_name = '건설 (공표자료)'
        
        # 기존 HTML 필러의 process_marker 로직 사용
        html_marker_info = {
            'sheet_name': marker_sheet_name,
            'cell_address': cell_address,
            'operation': operation,
            'full_match': marker_info.get('full_match', '')
        }
        
        # HTML 필러의 로직 사용 (내부적으로 동일한 엑셀 추출기 사용)
        self.html_filler._current_year = self._current_year
        self.html_filler._current_quarter = self._current_quarter
        self.html_filler._current_sheet_name = self._current_sheet_name
        
        return self.html_filler.process_marker(html_marker_info)
    
    def fill_template(self, sheet_name: str = None, year: int = None, quarter: int = None) -> None:
        """
        Word 템플릿의 모든 마커를 처리하여 완성된 템플릿을 만듭니다.
        
        Args:
            sheet_name: 시트 이름 (기본값: None, 마커에서 추출)
            year: 연도 (None이면 자동 감지)
            quarter: 분기 (None이면 자동 감지)
        """
        # 연도/분기가 지정되지 않으면 자동 감지
        # sheet_name이 None이면 마커에서 첫 번째 시트를 찾아서 사용
        if year is None or quarter is None:
            # 시트명이 없으면 마커에서 첫 번째 시트 찾기
            if sheet_name is None:
                # Word 템플릿 로드
                if self.template_manager.document is None:
                    self.template_manager.load_template()
                
                # 모든 마커 추출
                markers = self.template_manager.extract_markers()
                if markers:
                    # 첫 번째 마커의 시트명 사용
                    marker_sheet = markers[0]['sheet_name']
                    # 유연한 매핑으로 실제 시트명 찾기
                    actual_sheet = self.flexible_mapper.find_sheet_by_name(marker_sheet)
                    if actual_sheet:
                        sheet_name = actual_sheet
                    else:
                        # 매핑 실패 시 마커의 시트명 그대로 사용
                        sheet_name = marker_sheet
            
            if sheet_name:
                periods_info = self.period_detector.detect_available_periods(sheet_name)
                if year is None:
                    year = periods_info['default_year']
                if quarter is None:
                    quarter = periods_info['default_quarter']
        
        # 현재 처리 중인 연도/분기/시트명 저장
        self._current_year = year
        self._current_quarter = quarter
        self._current_sheet_name = sheet_name
        
        # HTML 필러에도 동일한 값 설정
        self.html_filler._current_year = year
        self.html_filler._current_quarter = quarter
        self.html_filler._current_sheet_name = sheet_name
        
        # Word 템플릿 로드
        if self.template_manager.document is None:
            self.template_manager.load_template()
        
        # 모든 마커 추출
        markers = self.template_manager.extract_markers()
        
        # 각 마커를 처리하여 치환
        for marker_info in markers:
            marker_str = marker_info['full_match']
            
            # 의미 기반 마커인 경우
            if marker_info.get('is_semantic', False):
                # 의미 기반 마커 해석기 사용
                from src.semantic_marker_resolver import SemanticMarkerResolver
                semantic_resolver = SemanticMarkerResolver(self.excel_extractor)
                
                resolved_value = semantic_resolver.resolve_semantic_marker(
                    marker_info['full_match'],
                    year=year,
                    quarter=quarter
                )
                
                if resolved_value is not None:
                    processed_value = self.html_filler._format_value(resolved_value)
                    if not processed_value or processed_value.startswith('[에러'):
                        processed_value = 'N/A'
                    self.template_manager.replace_marker(marker_info, processed_value)
                    continue
            
            # 기존 방식 (셀 주소 기반 또는 하위 호환)
            # sheet_name이 제공되면 마커의 시트명을 덮어쓰기
            if sheet_name:
                marker_info['sheet_name'] = sheet_name
            else:
                # 시트명이 제공되지 않으면 키워드 기반 의미 매칭 사용
                marker_sheet = marker_info.get('sheet_keyword') or marker_info.get('sheet_name')
                
                if marker_sheet:
                    # 키워드 기반 의미 매칭 시도
                    from src.semantic_sheet_matcher import SemanticSheetMatcher
                    semantic_matcher = SemanticSheetMatcher(self.excel_extractor)
                    semantic_sheet = semantic_matcher.find_sheet_by_semantic_keywords(marker_sheet)
                    
                    if semantic_sheet:
                        marker_info['sheet_name'] = semantic_sheet
                    else:
                        # 유연한 매핑으로 찾기
                        actual_sheet = self.flexible_mapper.find_sheet_by_name(marker_sheet)
                        if actual_sheet:
                            marker_info['sheet_name'] = actual_sheet
            
            # 마커 처리
            processed_value = self.process_marker(marker_info)
            
            # 값이 비어있거나 에러인 경우 N/A로 채움
            if not processed_value or processed_value.startswith('[에러'):
                processed_value = 'N/A'
            
            # Word 템플릿에서 마커를 값으로 치환
            self.template_manager.replace_marker(marker_info, processed_value)
    
    def save_filled_template(self, output_path: str) -> None:
        """
        채워진 템플릿을 파일로 저장합니다.
        
        Args:
            output_path: 저장할 파일 경로
        """
        self.template_manager.save_template(output_path)

