"""
템플릿 채우기 모듈 (슬림화 버전)
마커를 값으로 치환하고 포맷팅 처리
실제 마커 처리는 markers 모듈에 위임
"""

import html
import re
import math
from typing import Any, Dict, Optional

from .template_manager import TemplateManager
from .excel_extractor import ExcelExtractor
from .schema_loader import SchemaLoader
from ..utils.formatters import Formatter, safe_float, calculate_growth_rate
from ..utils.region_utils import (
    REGION_MAPPING, REGION_MAPPING_REVERSE,
    normalize_region_name, get_full_region_name, is_same_region
)
from ..utils.sheet_utils import SHEET_NAME_MAPPING, get_actual_sheet_name
from ..markers.dynamic_processor import DynamicMarkerProcessor
from ..markers.base import MarkerContext


class TemplateFiller:
    """템플릿에 데이터를 채우는 클래스 (슬림화 버전)"""
    
    def __init__(self, template_manager: TemplateManager, excel_extractor: ExcelExtractor, 
                 schema_loader: Optional[SchemaLoader] = None):
        """
        템플릿 필러 초기화
        
        Args:
            template_manager: 템플릿 관리자 인스턴스
            excel_extractor: 엑셀 추출기 인스턴스
            schema_loader: 스키마 로더 인스턴스 (기본값: 새로 생성)
        """
        self.template_manager = template_manager
        self.excel_extractor = excel_extractor
        self.schema_loader = schema_loader if schema_loader is not None else SchemaLoader()
        
        # 포맷터 초기화
        self.formatter = Formatter(self.schema_loader)
        
        # 분석기 및 파서는 지연 초기화
        self._calculator = None
        self._data_analyzer = None
        self._period_detector = None
        self._flexible_mapper = None
        self._dynamic_parser = None
        self._dynamic_marker_processor = None
        
        # 상태 변수
        self._current_year = None
        self._current_quarter = None
        self._current_sheet_name = None
        self._missing_value_overrides: Dict[str, float] = {}
        self._sheet_scale_cache: Dict[str, float] = {}
        self._current_analyzed_data: Dict[str, Any] = {}
    
    # === 지연 초기화 프로퍼티 ===
    
    @property
    def calculator(self):
        if self._calculator is None:
            from ..calculator import Calculator
            self._calculator = Calculator()
        return self._calculator
    
    @property
    def data_analyzer(self):
        if self._data_analyzer is None:
            from ..analyzers.data_analyzer import DataAnalyzer
            self._data_analyzer = DataAnalyzer(self.excel_extractor, self.schema_loader)
        return self._data_analyzer
    
    @property
    def period_detector(self):
        if self._period_detector is None:
            from ..analyzers.period_detector import PeriodDetector
            self._period_detector = PeriodDetector(self.excel_extractor)
        return self._period_detector
    
    @property
    def flexible_mapper(self):
        if self._flexible_mapper is None:
            from ..flexible_mapper import FlexibleMapper
            self._flexible_mapper = FlexibleMapper(self.excel_extractor)
        return self._flexible_mapper
    
    @property
    def dynamic_parser(self):
        if self._dynamic_parser is None:
            from ..analyzers.dynamic_sheet_parser import DynamicSheetParser
            self._dynamic_parser = DynamicSheetParser(self.excel_extractor, self.schema_loader)
        return self._dynamic_parser
    
    @property
    def dynamic_marker_processor(self):
        if self._dynamic_marker_processor is None:
            self._dynamic_marker_processor = DynamicMarkerProcessor(
                excel_extractor=self.excel_extractor,
                schema_loader=self.schema_loader,
                data_analyzer=self.data_analyzer,
                dynamic_parser=self.dynamic_parser,
                formatter=self.formatter
            )
        return self._dynamic_marker_processor
    
    # === 결측치 처리 ===
    
    def set_missing_value_overrides(self, overrides: Dict[str, float]) -> None:
        """사용자가 입력한 결측치 값을 설정합니다."""
        self._missing_value_overrides = overrides or {}
    
    def _get_missing_value_override(self, sheet_name: str, region: str, category: str, 
                                     year: int, quarter: int) -> Optional[float]:
        """사용자가 입력한 결측치 값을 가져옵니다."""
        key = f"{sheet_name}_{region}_{category}_{year}_{quarter}"
        return self._missing_value_overrides.get(key)
    
    def _detect_sheet_scale(self, sheet_name: str) -> float:
        """시트의 데이터 스케일을 감지합니다."""
        if sheet_name in self._sheet_scale_cache:
            return self._sheet_scale_cache[sheet_name]
        
        try:
            actual_sheet_name = get_actual_sheet_name(sheet_name)
            sheet = self.excel_extractor.get_sheet(actual_sheet_name)
            
            values = []
            for row in range(4, min(104, sheet.max_row + 1)):
                for col in range(7, min(sheet.max_column + 1, 20)):
                    cell = sheet.cell(row=row, column=col)
                    if cell.value is not None:
                        try:
                            val = float(cell.value)
                            if not math.isnan(val) and not math.isinf(val) and val > 0:
                                values.append(abs(val))
                        except (ValueError, TypeError):
                            continue
            
            if not values:
                self._sheet_scale_cache[sheet_name] = 1.0
                return 1.0
            
            values.sort()
            median_value = values[len(values) // 2]
            
            if median_value < 10:
                scale = 1.0
            elif median_value < 100:
                scale = 10.0
            elif median_value < 1000:
                scale = 100.0
            else:
                scale = 1000.0
            
            self._sheet_scale_cache[sheet_name] = scale
            return scale
        except Exception:
            self._sheet_scale_cache[sheet_name] = 1.0
            return 1.0
    
    # === 헬퍼 메서드 (하위 호환성) ===
    
    def _get_actual_sheet_name(self, sheet_name: str) -> str:
        """가상 시트명을 실제 시트명으로 변환합니다."""
        return get_actual_sheet_name(sheet_name)
    
    def _normalize_region_name(self, region_name: str) -> str:
        """지역명을 짧은 형식으로 정규화합니다."""
        return normalize_region_name(region_name)
    
    def _get_full_region_name(self, short_name: str) -> str:
        """짧은 지역명을 긴 형식으로 변환합니다."""
        return get_full_region_name(short_name)
    
    def _is_same_region(self, region1: str, region2: str) -> bool:
        """두 지역명이 같은 지역인지 확인합니다."""
        return is_same_region(region1, region2)
    
    def _handle_missing_value(self, value: Any, fallback: Any = None, sheet_name: str = None,
                               region: str = None, category: str = None) -> Any:
        """결측치를 처리합니다."""
        is_missing = False
        
        if value is None:
            is_missing = True
        elif isinstance(value, str):
            stripped = value.strip()
            if not stripped or stripped == '-':
                is_missing = True
        
        if is_missing:
            if sheet_name and region and self._current_year and self._current_quarter:
                override = self._get_missing_value_override(
                    sheet_name, region, category or '합계',
                    self._current_year, self._current_quarter
                )
                if override is not None:
                    return override
            
            if fallback is not None:
                return fallback
            
            if sheet_name:
                return self._detect_sheet_scale(sheet_name)
            
            return 1.0
        
        return value
    
    def _safe_float(self, value: Any, default: float = None) -> Optional[float]:
        """값을 안전하게 float로 변환합니다."""
        return safe_float(value, default)
    
    def _calculate_growth_rate(self, current: Any, prev: Any) -> Optional[float]:
        """증감률을 계산합니다."""
        return calculate_growth_rate(current, prev)
    
    # === 포맷팅 메서드 (하위 호환성) ===
    
    def format_number(self, value: Any, use_comma: bool = True, decimal_places: int = 1) -> str:
        """숫자를 포맷팅합니다."""
        return self.formatter.format_number(value, use_comma, decimal_places)
    
    def format_percentage(self, value: Any, decimal_places: int = 1, include_percent: bool = True) -> str:
        """퍼센트 값을 포맷팅합니다."""
        return self.formatter.format_percentage(value, decimal_places, include_percent)
    
    def get_growth_direction(self, value: Any, direction_type: str = "increase_decrease", 
                              expression_key: str = "rate") -> str:
        """증감률의 방향을 반환합니다."""
        return self.formatter.get_growth_direction(value, direction_type, expression_key)
    
    def get_production_change_expression(self, value: Any, direction_type: str = "increase_decrease") -> str:
        """생산/판매 변화 표현을 반환합니다."""
        return self.formatter.get_production_change_expression(value, direction_type)
    
    def format_growth_rate_abs(self, value: Any, decimal_places: int = 1, include_percent: bool = True) -> str:
        """증감률의 절대값을 포맷팅합니다."""
        return self.formatter.format_growth_rate_abs(value, decimal_places, include_percent)
    
    def format_growth_for_report(self, value: Any, decimal_places: int = 1, 
                                 direction_type: str = "increase_decrease") -> tuple:
        """보도자료용 증감률을 포맷팅합니다."""
        return self.formatter.format_growth_for_report(value, decimal_places, direction_type)
    
    def format_value_with_schema(self, value: Any, format_type: str = "percentage") -> str:
        """스키마 기반으로 값을 포맷팅합니다."""
        return self.formatter.format_value_with_schema(value, format_type)
    
    def escape_html(self, value: Any) -> str:
        """HTML 특수 문자를 이스케이프합니다."""
        return self.formatter.escape_html(value)
    
    def get_output_format(self, sheet_name: str = None) -> Optional[Dict]:
        """현재 시트의 출력 형식 스키마를 반환합니다."""
        if sheet_name is None:
            sheet_name = self._current_sheet_name
        if sheet_name is None:
            return None
        return self.schema_loader.get_output_format_for_sheet(sheet_name)
    
    # === 시트 설정 메서드 ===
    
    def _get_sheet_config(self, sheet_name: str) -> dict:
        """시트별 설정을 반환합니다."""
        actual_sheet_name = get_actual_sheet_name(sheet_name)
        return self.schema_loader.load_sheet_config(actual_sheet_name)
    
    def _get_quarter_columns(self, year: int, quarter: int, sheet_name: str = None) -> tuple:
        """연도와 분기에 해당하는 열 번호를 반환합니다."""
        if sheet_name:
            sheet_name = get_actual_sheet_name(sheet_name)
        
        if sheet_name:
            current_col = self.dynamic_parser.get_quarter_column(sheet_name, year, quarter)
            
            if not current_col:
                if quarter > 1:
                    current_col = self.dynamic_parser.get_quarter_column(sheet_name, year, quarter - 1)
                if not current_col:
                    current_col = self.dynamic_parser.get_quarter_column(sheet_name, year - 1, 4)
                if not current_col:
                    periods_info = self.period_detector.detect_available_periods(sheet_name)
                    max_year = periods_info.get('max_year')
                    max_quarter = periods_info.get('max_quarter')
                    if max_year and max_quarter:
                        current_col = self.dynamic_parser.get_quarter_column(sheet_name, max_year, max_quarter)
            
            if current_col:
                prev_col = self.dynamic_parser.get_quarter_column(sheet_name, year - 1, quarter)
                if not prev_col:
                    if quarter > 1:
                        prev_col = self.dynamic_parser.get_quarter_column(sheet_name, year - 1, quarter - 1)
                    if not prev_col:
                        prev_col = self.dynamic_parser.get_quarter_column(sheet_name, year - 2, 4)
                if prev_col:
                    return (current_col, prev_col)
        
        # 기존 로직 사용 (하위 호환성)
        config = self._get_sheet_config(sheet_name) if sheet_name else self.schema_loader.load_sheet_config('default')
        
        base_year = config['base_year']
        base_quarter = config['base_quarter']
        base_col = config['base_col']
        
        year_diff = year - base_year
        quarter_offset = (quarter - base_quarter)
        
        current_col = base_col + (year_diff * 4) + quarter_offset
        prev_col = current_col - 4
        
        return (current_col, prev_col)
    
    # === 마커 처리 ===
    
    def _process_dynamic_marker(self, sheet_name: str, key: str, year: int = 2025, quarter: int = 2) -> Optional[str]:
        """동적 마커를 처리합니다."""
        return self.dynamic_marker_processor.process(sheet_name, key, year, quarter)
    
    def process_marker(self, marker_info: Dict[str, str]) -> str:
        """
        마커를 처리하고 값을 반환합니다.
        
        Args:
            marker_info: 마커 정보 딕셔너리
            
        Returns:
            처리된 값 문자열
        """
        sheet_name = marker_info.get('sheet_name', '')
        cell_address = marker_info.get('cell_address', '')
        operation = marker_info.get('operation')
        
        year = self._current_year or 2025
        quarter = self._current_quarter or 2
        
        # 셀 주소가 알파벳+숫자 형식이 아니면 동적 마커로 처리
        if not re.match(r'^[A-Za-z]+\d+', cell_address):
            result = self._process_dynamic_marker(sheet_name, cell_address, year, quarter)
            if result and result != "N/A":
                return result
        
        # 일반 셀 참조 처리
        try:
            actual_sheet_name = get_actual_sheet_name(sheet_name)
            value = self.excel_extractor.extract_value(actual_sheet_name, cell_address)
            
            if operation:
                value = self._apply_operation(value, operation)
            
            return self._format_value(value, operation)
        except Exception as e:
            print(f"[WARNING] 마커 처리 중 오류 ({marker_info}): {str(e)}")
            return "N/A"
    
    def _apply_operation(self, raw_value: Any, operation: str) -> Any:
        """연산을 적용합니다."""
        if not operation:
            return raw_value
        
        operation = operation.lower().strip()
        
        if isinstance(raw_value, list):
            return self.calculator.calculate_from_cell_refs(operation, raw_value)
        
        return raw_value
    
    def _format_value(self, value: Any, operation: str = None) -> str:
        """값을 포맷팅합니다."""
        if value is None:
            return "N/A"
        
        if isinstance(value, str):
            stripped = value.strip()
            if not stripped or stripped == '-':
                return "N/A"
            return stripped
        
        if isinstance(value, (int, float)):
            if math.isnan(value) or math.isinf(value):
                return "N/A"
            
            if operation and operation.lower() in ['growth_rate', '증감률', '증가율']:
                return self.format_percentage(value, decimal_places=1, include_percent=False)
            
            return self.format_number(value)
        
        return str(value)
    
    # === 템플릿 채우기 ===
    
    def fill_template(self, sheet_name: str = None, year: int = None, quarter: int = None) -> str:
        """
        템플릿의 모든 마커를 채웁니다.
        
        Args:
            sheet_name: 기본 시트 이름 (마커에 시트가 명시되지 않은 경우 사용)
            year: 연도
            quarter: 분기
            
        Returns:
            채워진 HTML 문자열
        """
        self._current_year = year
        self._current_quarter = quarter
        self._current_sheet_name = sheet_name
        
        # 템플릿 로드
        content = self.template_manager.get_template_content()
        
        # CSS/Script 블록 임시 저장
        style_matches = []
        script_matches = []
        
        def style_replacer(match):
            style_matches.append(match.group(0))
            return f"__STYLE_PLACEHOLDER_{len(style_matches) - 1}__"
        
        def script_replacer(match):
            script_matches.append(match.group(0))
            return f"__SCRIPT_PLACEHOLDER_{len(script_matches) - 1}__"
        
        style_pattern = re.compile(r'<style[^>]*>.*?</style>', re.DOTALL | re.IGNORECASE)
        script_pattern = re.compile(r'<script[^>]*>.*?</script>', re.DOTALL | re.IGNORECASE)
        
        content = style_pattern.sub(style_replacer, content)
        content = script_pattern.sub(script_replacer, content)
        
        # 마커 추출 및 처리
        markers = self.template_manager.extract_markers()
        
        for marker in markers:
            full_match = marker['full_match']
            processed_value = self.process_marker(marker)
            content = content.replace(full_match, processed_value)
        
        # CSS/Script 블록 복원
        for i, style in enumerate(style_matches):
            content = content.replace(f"__STYLE_PLACEHOLDER_{i}__", style)
        
        for i, script in enumerate(script_matches):
            content = content.replace(f"__SCRIPT_PLACEHOLDER_{i}__", script)
        
        return content
    
    def fill_template_with_custom_format(self, format_func: callable = None) -> str:
        """사용자 정의 포맷팅 함수로 템플릿을 채웁니다."""
        if format_func is None:
            return self.fill_template()
        
        content = self.template_manager.get_template_content()
        markers = self.template_manager.extract_markers()
        
        for marker in markers:
            full_match = marker['full_match']
            raw_value = self.process_marker(marker)
            formatted_value = format_func(raw_value, marker)
            content = content.replace(full_match, formatted_value)
        
        return content

