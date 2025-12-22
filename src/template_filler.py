"""
템플릿 채우기 모듈
마커를 값으로 치환하고 포맷팅 처리
"""

import html
import re
import math
from typing import Any, Dict, Optional
from .template_manager import TemplateManager
from .excel_extractor import ExcelExtractor
from .calculator import Calculator
from .data_analyzer import DataAnalyzer
from .period_detector import PeriodDetector
from .flexible_mapper import FlexibleMapper
from .dynamic_sheet_parser import DynamicSheetParser
from .schema_loader import SchemaLoader


# 지역명 매핑 상수 (중복 제거)
REGION_MAPPING = {
    '서울특별시': '서울', '부산광역시': '부산', '대구광역시': '대구',
    '인천광역시': '인천', '광주광역시': '광주', '대전광역시': '대전',
    '울산광역시': '울산', '세종특별자치시': '세종', '경기도': '경기',
    '강원도': '강원', '충청북도': '충북', '충청남도': '충남',
    '전라북도': '전북', '전라남도': '전남', '경상북도': '경북',
    '경상남도': '경남', '제주특별자치도': '제주'
}

# 역방향 매핑 (짧은 이름 -> 긴 이름)
REGION_MAPPING_REVERSE = {v: k for k, v in REGION_MAPPING.items()}

# 시트명 매핑 (가상 시트명 -> 실제 시트명)
# 실업률 템플릿은 "실업자 수" 시트의 데이터를 사용
SHEET_NAME_MAPPING = {
    '실업률': '실업자 수'
}


class TemplateFiller:
    """템플릿에 데이터를 채우는 클래스"""
    
    def __init__(self, template_manager: TemplateManager, excel_extractor: ExcelExtractor, schema_loader: Optional[SchemaLoader] = None):
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
        self.calculator = Calculator()
        self.data_analyzer = DataAnalyzer(excel_extractor, self.schema_loader)
        self.period_detector = PeriodDetector(excel_extractor)
        self.flexible_mapper = FlexibleMapper(excel_extractor)
        self.dynamic_parser = DynamicSheetParser(excel_extractor, self.schema_loader)
        self._current_year = None  # 현재 처리 중인 연도
        self._current_quarter = None  # 현재 처리 중인 분기
        self._current_sheet_name = None  # 현재 처리 중인 시트명
        self._missing_value_overrides: Dict[str, float] = {}  # 사용자가 입력한 결측치 값
        self._sheet_scale_cache: Dict[str, float] = {}  # 시트별 스케일 캐시
    
    def set_missing_value_overrides(self, overrides: Dict[str, float]) -> None:
        """
        사용자가 입력한 결측치 값을 설정합니다.
        
        Args:
            overrides: 결측치 값 딕셔너리 (키: 시트_지역_카테고리_연도_분기, 값: 숫자)
        """
        self._missing_value_overrides = overrides or {}
    
    def _get_missing_value_override(self, sheet_name: str, region: str, category: str, year: int, quarter: int) -> Optional[float]:
        """
        사용자가 입력한 결측치 값을 가져옵니다.
        
        Args:
            sheet_name: 시트명
            region: 지역명
            category: 카테고리명
            year: 연도
            quarter: 분기
            
        Returns:
            사용자가 입력한 값 또는 None
        """
        key = f"{sheet_name}_{region}_{category}_{year}_{quarter}"
        return self._missing_value_overrides.get(key)
    
    def _detect_sheet_scale(self, sheet_name: str) -> float:
        """
        시트의 데이터 스케일을 감지합니다.
        - 한 자리 수 (1-9) → 1
        - 두 자리 수 (10-99) → 10
        - 세 자리 수 (100-999) → 100
        - 네 자리 수 이상 (1000+) → 1000
        
        Args:
            sheet_name: 시트 이름
            
        Returns:
            적절한 스케일 값
        """
        if sheet_name in self._sheet_scale_cache:
            return self._sheet_scale_cache[sheet_name]
        
        try:
            actual_sheet_name = self._get_actual_sheet_name(sheet_name)
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
    
    # ===== 헬퍼 메서드 (중복 코드 제거) =====
    
    def _get_actual_sheet_name(self, sheet_name: str) -> str:
        """가상 시트명을 실제 시트명으로 변환합니다.
        
        Args:
            sheet_name: 템플릿에서 사용하는 시트명
            
        Returns:
            실제 엑셀 시트명
        """
        return SHEET_NAME_MAPPING.get(sheet_name, sheet_name)
    
    def _normalize_region_name(self, region_name: str) -> str:
        """지역명을 짧은 형식으로 정규화합니다."""
        if not region_name:
            return ''
        region_name = str(region_name).strip()
        return REGION_MAPPING.get(region_name, region_name)
    
    def _get_full_region_name(self, short_name: str) -> str:
        """짧은 지역명을 긴 형식으로 변환합니다."""
        if not short_name:
            return ''
        short_name = str(short_name).strip()
        return REGION_MAPPING_REVERSE.get(short_name, short_name)
    
    def _is_same_region(self, region1: str, region2: str) -> bool:
        """두 지역명이 같은 지역인지 확인합니다."""
        if not region1 or not region2:
            return False
        norm1 = self._normalize_region_name(region1)
        norm2 = self._normalize_region_name(region2)
        return norm1 == norm2
    
    def _handle_missing_value(self, value: Any, fallback: Any = None, sheet_name: str = None,
                               region: str = None, category: str = None) -> Any:
        """
        결측치를 처리합니다. None, 빈 문자열, '-'를 fallback 값으로 대체합니다.
        사용자가 입력한 값이 있으면 우선 사용하고, 없으면 스마트 기본값을 추정합니다.
        
        Args:
            value: 처리할 값
            fallback: 대체할 값 (None이면 스마트 기본값 추정)
            sheet_name: 시트 이름 (결측치 오버라이드 확인용)
            region: 지역명 (결측치 오버라이드 확인용)
            category: 카테고리명 (결측치 오버라이드 확인용)
            
        Returns:
            처리된 값
        """
        is_missing = False
        
        if value is None:
            is_missing = True
        elif isinstance(value, str):
            stripped = value.strip()
            if not stripped or stripped == '-':
                is_missing = True
        
        if is_missing:
            # 우선순위 1: 사용자가 입력한 값
            if sheet_name and region and self._current_year and self._current_quarter:
                override = self._get_missing_value_override(
                    sheet_name, region, category or '합계',
                    self._current_year, self._current_quarter
                )
                if override is not None:
                    return override
            
            # 우선순위 2: fallback 값
            if fallback is not None:
                return fallback
            
            # 우선순위 3: 스마트 기본값 추정
            if sheet_name and region and self._current_year and self._current_quarter:
                smart_default = self._estimate_missing_value(
                    sheet_name, region, category,
                    self._current_year, self._current_quarter
                )
                if smart_default is not None:
                    return smart_default
            
            # 우선순위 4: 시트 스케일 기반 기본값
            if sheet_name:
                return self._detect_sheet_scale(sheet_name)
            
            return 1.0  # 최종 기본값
        
        return value
    
    def _estimate_missing_value(self, sheet_name: str, region: str, category: str,
                                  year: int, quarter: int) -> Optional[float]:
        """
        결측치의 스마트 기본값을 추정합니다.
        이전 분기 값, 전년 동분기 값, 또는 지역 평균을 기반으로 추정합니다.
        
        Args:
            sheet_name: 시트 이름
            region: 지역명
            category: 카테고리명
            year: 연도
            quarter: 분기
            
        Returns:
            추정된 값 또는 None
        """
        try:
            # 이전 분기 값 시도
            prev_quarter = quarter - 1
            prev_year = year
            if prev_quarter < 1:
                prev_quarter = 4
                prev_year = year - 1
            
            prev_value = self.dynamic_parser.get_quarter_value(
                sheet_name, region, prev_year, prev_quarter
            )
            if prev_value is not None and prev_value != 1.0:
                return prev_value
            
            # 전년 동분기 값 시도
            last_year_value = self.dynamic_parser.get_quarter_value(
                sheet_name, region, year - 1, quarter
            )
            if last_year_value is not None and last_year_value != 1.0:
                return last_year_value
            
            # 같은 시트의 전국 값 시도
            if region != '전국':
                national_value = self.dynamic_parser.get_quarter_value(
                    sheet_name, '전국', year, quarter
                )
                if national_value is not None and national_value != 1.0:
                    # 전국 값의 평균적인 비율로 추정
                    return national_value
            
        except Exception:
            pass
        
        return None
    
    def _safe_float(self, value: Any, default: float = None) -> Optional[float]:
        """값을 안전하게 float로 변환합니다."""
        if value is None:
            return default
        if isinstance(value, str):
            stripped = value.strip()
            if not stripped or stripped == '-':
                return default
            try:
                return float(stripped)
            except (ValueError, TypeError):
                return default
        try:
            result = float(value)
            if math.isnan(result) or math.isinf(result):
                return default
            return result
        except (ValueError, TypeError, OverflowError):
            return default
    
    def _calculate_growth_rate(self, current: Any, prev: Any) -> Optional[float]:
        """증감률을 계산합니다. ((current / prev) - 1) * 100"""
        current_val = self._safe_float(current)
        prev_val = self._safe_float(prev)
        
        if current_val is None or prev_val is None or prev_val == 0:
            return None
        
        try:
            growth = ((current_val / prev_val) - 1) * 100
            if math.isnan(growth) or math.isinf(growth):
                return None
            return growth
        except (ZeroDivisionError, OverflowError):
            return None
    
    # ===== 포맷팅 메서드 =====
    
    def format_number(self, value: Any, use_comma: bool = True, decimal_places: int = 1) -> str:
        """
        숫자를 포맷팅합니다.
        
        Args:
            value: 포맷팅할 값
            use_comma: 천 단위 구분 기호 사용 여부
            decimal_places: 소수점 자릿수 (기본값: 1)
            
        Returns:
            포맷팅된 문자열 (오류 시 "N/A")
        """
        try:
            # None이나 빈 값 체크
            if value is None:
                return "N/A"
            if isinstance(value, str) and not value.strip():
                return "N/A"
            
            # 숫자로 변환
            num = float(value)
            
            # NaN이나 Infinity 체크
            import math
            if math.isnan(num) or math.isinf(num):
                return "N/A"
            
            # 소수점 처리
            num = round(num, decimal_places)
            
            # 문자열로 변환 (항상 지정된 소수점 자릿수까지 표시)
            if decimal_places > 0:
                formatted = f"{num:.{decimal_places}f}"
            else:
                formatted = str(int(num))
            
            # 천 단위 구분
            if use_comma:
                parts = formatted.split('.')
                # 정수 부분에 천 단위 구분 적용 (음수 처리 포함)
                try:
                    integer_part = int(float(parts[0]))
                    # 천 단위 구분 적용 (음수도 올바르게 처리)
                    parts[0] = f"{integer_part:,}"
                except (ValueError, TypeError):
                    pass  # 변환 실패 시 원본 유지
                formatted = '.'.join(parts) if len(parts) > 1 else parts[0]
            
            return formatted
        except (ValueError, TypeError, OverflowError):
            # 숫자로 변환할 수 없으면 N/A 반환
            return "N/A"
    
    def format_percentage(self, value: Any, decimal_places: int = 1, include_percent: bool = True) -> str:
        """
        퍼센트 값을 포맷팅합니다.
        
        Args:
            value: 퍼센트 값 (예: 5.5는 5.5%를 의미)
            decimal_places: 소수점 자릿수 (기본값: 1)
            include_percent: % 기호 포함 여부
            
        Returns:
            포맷팅된 퍼센트 문자열 (예: "5.5%" 또는 "5.5", 오류 시 "N/A")
        """
        try:
            # None이나 빈 값 체크
            if value is None:
                return "N/A"
            if isinstance(value, str) and not value.strip():
                return "N/A"
            
            num = float(value)
            
            # NaN이나 Infinity 체크
            import math
            if math.isnan(num) or math.isinf(num):
                return "N/A"
            
            # 소수점 반올림
            num = round(num, decimal_places)
            
            # 항상 소수점 첫째자리까지 표시 (0이어도 0.0으로 표시)
            formatted = f"{num:.{decimal_places}f}"
            if include_percent:
                return f"{formatted}%"
            return formatted
        except (ValueError, TypeError, OverflowError):
            return "N/A"
    
    def get_growth_direction(self, value: Any, direction_type: str = "increase_decrease", 
                              expression_key: str = "rate") -> str:
        """
        증감률의 방향을 반환합니다 (보도자료용).
        스키마 기반으로 동작하며, 다양한 방향 표현을 지원합니다.
        
        Args:
            value: 증감률 값 (양수면 증가, 음수면 감소)
            direction_type: 방향 타입 ('increase_decrease', 'rise_fall')
            expression_key: 표현 키 ('rate', 'production', 'result', 'change')
            
        Returns:
            방향 문자열 (예: '증가', '감소', '상승', '하락') (오류 시 빈 문자열)
        """
        try:
            if value is None:
                return ""
            if isinstance(value, str) and not value.strip():
                return ""
            
            num = float(value)
            
            if math.isnan(num) or math.isinf(num):
                return ""
            
            # 스키마에서 방향 표현 가져오기
            if direction_type == "rise_fall":
                if num > 0:
                    return self.schema_loader.get_direction_expression("rise", expression_key) or "상승"
                elif num < 0:
                    return self.schema_loader.get_direction_expression("fall", expression_key) or "하강"
                else:
                    return "보합"
            else:  # increase_decrease (기본값)
                if num > 0:
                    return self.schema_loader.get_direction_expression("increase", expression_key) or "증가"
                elif num < 0:
                    return self.schema_loader.get_direction_expression("decrease", expression_key) or "감소"
                else:
                    return "동일"
        except (ValueError, TypeError, OverflowError):
            return ""
    
    def get_production_change_expression(self, value: Any, direction_type: str = "increase_decrease") -> str:
        """
        생산/판매 변화 표현을 반환합니다 (보도자료용).
        
        Args:
            value: 증감률 값
            direction_type: 방향 타입 ('increase_decrease', 'rise_fall')
            
        Returns:
            변화 표현 문자열 (예: '늘어', '줄어', '올라', '내려')
        """
        try:
            if value is None:
                return ""
            num = float(value)
            
            if math.isnan(num) or math.isinf(num):
                return ""
            
            if direction_type == "rise_fall":
                if num > 0:
                    return self.schema_loader.get_direction_expression("rise", "change") or "올라"
                elif num < 0:
                    return self.schema_loader.get_direction_expression("fall", "change") or "내려"
                else:
                    return "유지되어"
            else:
                if num > 0:
                    return self.schema_loader.get_direction_expression("increase", "production") or "늘어"
                elif num < 0:
                    return self.schema_loader.get_direction_expression("decrease", "production") or "줄어"
                else:
                    return "유지되어"
        except (ValueError, TypeError, OverflowError):
            return ""
    
    def format_growth_rate_abs(self, value: Any, decimal_places: int = 1, include_percent: bool = True) -> str:
        """
        증감률의 절대값을 포맷팅합니다 (보도자료용, 마이너스 기호 제거).
        
        Args:
            value: 증감률 값
            decimal_places: 소수점 자릿수 (기본값: 1)
            include_percent: % 기호 포함 여부
            
        Returns:
            마이너스 기호가 제거된 포맷팅된 증감률 문자열 (예: "3.5%")
        """
        try:
            if value is None:
                return "N/A"
            if isinstance(value, str) and not value.strip():
                return "N/A"
            
            num = float(value)
            
            import math
            if math.isnan(num) or math.isinf(num):
                return "N/A"
            
            # 절대값으로 변환 (마이너스 기호 제거)
            num = abs(num)
            
            # 소수점 반올림
            num = round(num, decimal_places)
            
            formatted = f"{num:.{decimal_places}f}"
            if include_percent:
                return f"{formatted}%"
            return formatted
        except (ValueError, TypeError, OverflowError):
            return "N/A"
    
    def format_growth_for_report(self, value: Any, decimal_places: int = 1, 
                                 direction_type: str = "increase_decrease") -> tuple:
        """
        보도자료용 증감률을 포맷팅합니다.
        
        Args:
            value: 증감률 값
            decimal_places: 소수점 자릿수 (기본값: 1)
            direction_type: 방향 타입 ('increase_decrease', 'rise_fall')
            
        Returns:
            (포맷팅된 수치, 방향) 튜플 (예: ("3.5%", "증가") 또는 ("2.1%", "감소"))
        """
        rate = self.format_growth_rate_abs(value, decimal_places, include_percent=True)
        direction = self.get_growth_direction(value, direction_type)
        return (rate, direction)
    
    def format_value_with_schema(self, value: Any, format_type: str = "percentage") -> str:
        """
        스키마 기반으로 값을 포맷팅합니다.
        
        Args:
            value: 포맷팅할 값
            format_type: 포맷 타입 ('percentage', 'percentage_point', 'count', 'population')
            
        Returns:
            포맷팅된 문자열
        """
        format_rule = self.schema_loader.get_format_rule(format_type)
        
        decimal_places = format_rule.get('decimal_places', 1)
        suffix = format_rule.get('suffix', '')
        use_comma = format_rule.get('use_comma', False)
        
        try:
            if value is None:
                return "N/A"
            
            num = float(value)
            if math.isnan(num) or math.isinf(num):
                return "N/A"
            
            # 소수점 반올림
            num = round(num, decimal_places)
            
            # 포맷팅
            if decimal_places > 0:
                formatted = f"{num:.{decimal_places}f}"
            else:
                formatted = str(int(num))
            
            # 천 단위 구분
            if use_comma:
                parts = formatted.split('.')
                try:
                    integer_part = int(float(parts[0]))
                    parts[0] = f"{integer_part:,}"
                except (ValueError, TypeError):
                    pass
                formatted = '.'.join(parts) if len(parts) > 1 else parts[0]
            
            return f"{formatted}{suffix}"
        except (ValueError, TypeError, OverflowError):
            return "N/A"
    
    def get_output_format(self, sheet_name: str = None) -> Optional[Dict]:
        """
        현재 시트의 출력 형식 스키마를 반환합니다.
        
        Args:
            sheet_name: 시트 이름 (None이면 현재 처리 중인 시트 사용)
            
        Returns:
            출력 형식 딕셔너리 또는 None
        """
        if sheet_name is None:
            sheet_name = self._current_sheet_name
        
        if sheet_name is None:
            return None
        
        return self.schema_loader.get_output_format_for_sheet(sheet_name)
    
    def escape_html(self, value: Any) -> str:
        """
        HTML 특수 문자를 이스케이프합니다.
        
        Args:
            value: 이스케이프할 값
            
        Returns:
            이스케이프된 문자열
        """
        return html.escape(str(value) if value is not None else "")
    
    def _analyze_data_if_needed(self, sheet_name: str, year: int = 2025, quarter: int = 2) -> None:
        """
        필요시 데이터를 분석하여 캐시에 저장
        
        Args:
            sheet_name: 시트 이름
            year: 연도
            quarter: 분기 (1-4)
        """
        # 시트별로 다른 캐시 키 사용 (시트가 다르면 다시 분석)
        cache_key = f"{sheet_name}_{year}_{quarter}"
        # 이미 분석된 데이터가 있으면 재분석하지 않음
        if hasattr(self, '_current_analyzed_data') and cache_key in self._current_analyzed_data:
            return
        
        try:
            # 캐시 없이 항상 새로 분석
            quarter_cols = self._get_quarter_columns(year, quarter, sheet_name)
            if quarter_cols and len(quarter_cols) == 2:
                quarter_data = {f"{year}_{quarter}/4": quarter_cols}
                analyzed = self.data_analyzer.analyze_quarter_data(sheet_name, quarter_data)
                # 캐시 대신 인스턴스 변수에 임시 저장 (시트별로 구분)
                if not hasattr(self, '_current_analyzed_data'):
                    self._current_analyzed_data = {}
                self._current_analyzed_data[cache_key] = analyzed.get(f"{year}_{quarter}/4", {})
            else:
                # quarter_cols가 없거나 잘못된 경우 빈 딕셔너리 저장
                if not hasattr(self, '_current_analyzed_data'):
                    self._current_analyzed_data = {}
                self._current_analyzed_data[cache_key] = {}
        except Exception as e:
            # 에러 발생 시 빈 딕셔너리 저장하여 계속 진행
            if not hasattr(self, '_current_analyzed_data'):
                self._current_analyzed_data = {}
            self._current_analyzed_data[cache_key] = {}
            print(f"[WARNING] 데이터 분석 중 오류 발생 (시트={sheet_name}, 연도={year}, 분기={quarter}): {str(e)}")
    
    def _get_sheet_config(self, sheet_name: str) -> dict:
        """
        시트별 설정을 반환합니다.
        
        Args:
            sheet_name: 시트 이름
            
        Returns:
            시트 설정 딕셔너리
        """
        # 가상 시트명을 실제 시트명으로 변환
        actual_sheet_name = self._get_actual_sheet_name(sheet_name)
        return self.schema_loader.load_sheet_config(actual_sheet_name)
    
    def _get_quarter_columns(self, year: int, quarter: int, sheet_name: str = None) -> tuple:
        """
        연도와 분기에 해당하는 열 번호를 반환합니다.
        헤더 기반으로 동적으로 찾습니다.
        
        Args:
            year: 연도
            quarter: 분기 (1-4)
            sheet_name: 시트 이름 (시트별로 기준점이 다를 수 있음)
            
        Returns:
            (현재 분기 열, 전년 동분기 열) 튜플
        """
        # 가상 시트명을 실제 시트명으로 변환
        if sheet_name:
            sheet_name = self._get_actual_sheet_name(sheet_name)
        if sheet_name:
            # 헤더 기반으로 분기 열 찾기
            current_col = self.dynamic_parser.get_quarter_column(sheet_name, year, quarter)
            
            # 찾지 못한 경우, 이전 분기나 최신 분기 시도
            if not current_col:
                # 이전 분기 시도
                if quarter > 1:
                    current_col = self.dynamic_parser.get_quarter_column(sheet_name, year, quarter - 1)
                # 여전히 없으면 이전 연도 마지막 분기 시도
                if not current_col:
                    current_col = self.dynamic_parser.get_quarter_column(sheet_name, year - 1, 4)
                # 여전히 없으면 period_detector로 최신 분기 찾기
                if not current_col:
                    periods_info = self.period_detector.detect_available_periods(sheet_name)
                    max_year = periods_info.get('max_year')
                    max_quarter = periods_info.get('max_quarter')
                    if max_year and max_quarter:
                        current_col = self.dynamic_parser.get_quarter_column(sheet_name, max_year, max_quarter)
            
            if current_col:
                # 전년 동분기 열 찾기
                prev_col = self.dynamic_parser.get_quarter_column(sheet_name, year - 1, quarter)
                # 전년 동분기를 찾지 못한 경우, 이전 분기나 최신 분기 시도
                if not prev_col:
                    if quarter > 1:
                        prev_col = self.dynamic_parser.get_quarter_column(sheet_name, year - 1, quarter - 1)
                    if not prev_col:
                        prev_col = self.dynamic_parser.get_quarter_column(sheet_name, year - 2, 4)
                if prev_col:
                    return (current_col, prev_col)
        
        # 헤더 기반으로 찾지 못한 경우, 기존 로직 사용 (하위 호환성)
        config = self._get_sheet_config(sheet_name) if sheet_name else self.schema_loader.load_sheet_config('default')
        
        base_year = config['base_year']
        base_quarter = config['base_quarter']
        base_col = config['base_col']
        
        # 연도 차이 계산
        year_diff = year - base_year
        
        # 분기 오프셋 계산
        quarter_offset = (quarter - base_quarter)
        
        # 현재 분기 열 계산
        current_col = base_col + (year_diff * 4) + quarter_offset
        
        # 전년 동분기 열 계산
        prev_col = current_col - 4
        
        return (current_col, prev_col)
    
    def _get_unemployment_rate_regions_data(self, year: int, quarter: int) -> list:
        """
        실업률 테이블에서 모든 지역의 증감률(백분율포인트)을 계산하여 반환합니다.
        
        Args:
            year: 연도
            quarter: 분기
            
        Returns:
            지역별 증감률 정보 리스트 [{'name': '서울', 'growth_rate': 0.2}, ...]
        """
        # 17개 시도 리스트 (전국 제외)
        region_list = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                       '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 실업자 수 시트에서 실업률 테이블 설정 가져오기
        actual_sheet_name = '실업자 수'
        sheet = self.excel_extractor.get_sheet(actual_sheet_name)
        if not sheet:
            return []
        
        sheet_config = self._get_sheet_config(actual_sheet_name)
        unemployment_table_config = sheet_config.get('unemployment_rate_table', {})
        
        if not unemployment_table_config.get('enabled', True):
            return []
        
        start_row = unemployment_table_config.get('start_row', 81)
        title_row = unemployment_table_config.get('title_row', 79)
        region_col = unemployment_table_config.get('region_column', 1)
        age_group_col = unemployment_table_config.get('age_group_column', 2)
        age_group_filter = unemployment_table_config.get('age_group_filter', '계')
        header_row_number = unemployment_table_config.get('header_row_number', 3)
        region_mapping = unemployment_table_config.get('region_mapping', {})
        
        # 역매핑 (긴 이름 -> 짧은 이름)
        reverse_mapping = {v: k for k, v in region_mapping.items()}
        
        # 헤더 행 찾기
        title_text = unemployment_table_config.get('title_text', '시도별 실업률(%)')
        header_row = None
        
        for row in range(1, min(200, sheet.max_row + 1)):
            for col in range(1, min(10, sheet.max_column + 1)):
                cell_val = sheet.cell(row=row, column=col).value
                if cell_val and title_text in str(cell_val):
                    for offset in [header_row_number, 1, 2, 3]:
                        test_row = row + offset
                        if test_row <= sheet.max_row:
                            for test_col in range(1, min(20, sheet.max_column + 1)):
                                test_val = sheet.cell(row=test_row, column=test_col).value
                                if test_val and ('분기' in str(test_val) or '/4' in str(test_val)):
                                    header_row = test_row
                                    break
                            if header_row:
                                break
                    break
            if header_row:
                break
        
        if header_row is None:
            header_row = title_row + header_row_number if title_row else header_row_number
        
        # 현재 분기와 전년 동분기 열 찾기
        current_col = None
        prev_col = None
        patterns_current = [f"{year}  {quarter}/4", f"{year} {quarter}/4", f"{year}년 {quarter}/4"]
        patterns_prev = [f"{year-1}  {quarter}/4", f"{year-1} {quarter}/4", f"{year-1}년 {quarter}/4"]
        
        for col in range(1, min(100, sheet.max_column + 1)):
            header_val = sheet.cell(row=header_row, column=col).value
            if header_val:
                header_str = str(header_val).strip()
                for pattern in patterns_current:
                    if pattern in header_str:
                        current_col = col
                        break
                for pattern in patterns_prev:
                    if pattern in header_str:
                        prev_col = col
                        break
        
        if not current_col or not prev_col:
            return []
        
        # 지역별 데이터 추출
        regions = []
        seen_regions = set()
        current_region_name = None
        
        for row in range(start_row, min(start_row + 500, sheet.max_row + 1)):
            cell_region = sheet.cell(row=row, column=region_col).value
            cell_age = sheet.cell(row=row, column=age_group_col).value
            
            # 지역 셀이 있으면 갱신
            if cell_region:
                region_str = str(cell_region).strip()
                # 역매핑으로 짧은 이름 가져오기
                current_region_name = reverse_mapping.get(region_str, region_str)
                # 매핑되지 않은 경우 직접 확인
                if current_region_name == region_str and region_str not in region_list:
                    for short_name in region_list:
                        if short_name in region_str:
                            current_region_name = short_name
                            break
            
            if not current_region_name or not cell_age:
                continue
            
            age_str = str(cell_age).strip()
            
            # '계' 연령대만, 전국 제외, 중복 제외
            if age_str != age_group_filter:
                continue
            if current_region_name == '전국' or current_region_name in seen_regions:
                continue
            if current_region_name not in region_list:
                continue
            
            seen_regions.add(current_region_name)
            
            # 현재 분기와 전년 동분기 값 가져오기
            current_val = sheet.cell(row=row, column=current_col).value
            prev_val = sheet.cell(row=row, column=prev_col).value
            
            try:
                current_float = float(current_val) if current_val is not None else None
                prev_float = float(prev_val) if prev_val is not None else None
            except (ValueError, TypeError):
                continue
            
            if current_float is not None and prev_float is not None:
                growth_rate = current_float - prev_float
                regions.append({
                    'name': current_region_name,
                    'growth_rate': round(growth_rate, 1)
                })
        
        return regions
    
    def _get_unemployment_rate_age_groups(self, region_name: str, year: int, quarter: int) -> list:
        """
        실업률 테이블에서 특정 지역의 연령별 증감률(백분율포인트)을 계산하여 반환합니다.
        
        Args:
            region_name: 지역 이름 (짧은 이름, 예: '서울', '전국')
            year: 연도
            quarter: 분기
            
        Returns:
            연령별 증감률 정보 리스트 [{'name': '30~59세', 'growth_rate': 0.5}, ...]
            증감률 절대값 기준 내림차순 정렬
        """
        actual_sheet_name = '실업자 수'
        sheet = self.excel_extractor.get_sheet(actual_sheet_name)
        if not sheet:
            return []
        
        # 실업자 수 시트의 스키마에서 실업률 테이블 설정 가져오기
        sheet_config = self.schema_loader.load_sheet_config('실업자 수')
        unemployment_table_config = sheet_config.get('unemployment_rate_table', {})
        
        if not unemployment_table_config.get('enabled', True):
            return []
        
        start_row = unemployment_table_config.get('start_row', 81)
        header_row = unemployment_table_config.get('header_row', 80)
        region_col = unemployment_table_config.get('region_column', 1)
        age_group_col = unemployment_table_config.get('age_group_column', 2)
        region_mapping = unemployment_table_config.get('region_mapping', {})
        quarter_header_pattern = unemployment_table_config.get('quarter_header_pattern', '{year}  {quarter}/4')
        
        # 지역명 매핑 (짧은 이름 -> 긴 이름)
        target_region = region_mapping.get(region_name, region_name)
        
        # 현재 분기와 전년 동분기 열 찾기
        current_header = quarter_header_pattern.format(year=year, quarter=quarter)
        prev_header = quarter_header_pattern.format(year=year-1, quarter=quarter)
        
        current_col = None
        prev_col = None
        
        for col in range(1, min(sheet.max_column + 1, 100)):
            cell_value = sheet.cell(row=header_row, column=col).value
            if cell_value:
                cell_str = str(cell_value).strip()
                if cell_str == current_header:
                    current_col = col
                elif cell_str == prev_header:
                    prev_col = col
        
        if not current_col or not prev_col:
            return []
        
        age_groups = []
        last_region = None
        in_target_region = False
        
        for row in range(start_row, min(start_row + 500, sheet.max_row + 1)):
            region_cell = sheet.cell(row=row, column=region_col).value
            age_group_cell = sheet.cell(row=row, column=age_group_col).value
            
            # 지역 셀이 있으면 갱신
            if region_cell:
                region_str = str(region_cell).strip()
                if region_str:
                    last_region = region_str
                    in_target_region = (last_region == target_region)
            
            if not in_target_region or not age_group_cell:
                continue
            
            age_group_str = str(age_group_cell).strip()
            
            # '계'는 총계이므로 제외
            if age_group_str == '계':
                continue
            
            current_value = sheet.cell(row=row, column=current_col).value
            prev_value = sheet.cell(row=row, column=prev_col).value
            
            try:
                current_float = float(current_value) if current_value is not None else None
                prev_float = float(prev_value) if prev_value is not None else None
            except (ValueError, TypeError):
                continue
            
            if current_float is not None and prev_float is not None:
                growth_rate = round(current_float - prev_float, 1)
                age_groups.append({
                    'name': age_group_str,
                    'growth_rate': growth_rate,
                    'current_value': current_float,
                    'prev_value': prev_float
                })
        
        # 증감률 절대값 기준 내림차순 정렬 (가장 큰 변화가 먼저)
        age_groups.sort(key=lambda x: abs(x.get('growth_rate', 0)), reverse=True)
        
        return age_groups
    
    def _get_categories_for_region(self, sheet_name: str, region_name: str,
                                    year: int, quarter: int, top_n: int = 3) -> list:
        """
        특정 지역의 산업/업태별 증감률을 계산하여 반환 (일반화된 함수)
        헤더 기반으로 동적으로 계산합니다.
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름
            year: 연도
            quarter: 분기
            top_n: 상위 개수
            
        Returns:
            산업/업태별 증감률 정보 리스트
        """
        # 헤더 기반으로 시트 구조 파싱
        structure = self.dynamic_parser.parse_sheet_structure(sheet_name)
        config = self._get_sheet_config(sheet_name)
        category_col = config.get('category_column', structure.get('category_column', 6))  # 시트 설정 우선, 기본값: 6
        current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
        
        sheet = self.excel_extractor.get_sheet(sheet_name)
        
        # 헤더 기반으로 지역 행 찾기
        region_row = self.dynamic_parser.find_region_row(sheet_name, region_name)
        region_growth_rate = None
        
        # find_region_row가 실패한 경우 직접 찾기
        if region_row is None:
            structure = self.dynamic_parser.parse_sheet_structure(sheet_name)
            region_col = structure['region_column']
            data_start_row = structure['data_start_row']
            classification_col = structure.get('classification_column')
            category_col = config.get('category_column', structure.get('category_column', 6))
            
            # 지역명 매핑 (모듈 상수 사용)
            region_mapping = REGION_MAPPING
            
            for row in range(data_start_row, min(data_start_row + 500, sheet.max_row + 1)):
                cell_region = sheet.cell(row=row, column=region_col)
                if not cell_region.value:
                    continue
                    
                cell_region_str = str(cell_region.value).strip()
                # 지역명 매칭 (정확히 일치하거나 매핑된 이름과 일치)
                is_matching_region = (cell_region_str == region_name or 
                                     cell_region_str == region_mapping.get(region_name, '') or
                                     region_name == region_mapping.get(cell_region_str, ''))
                
                if is_matching_region:
                    # 총지수 또는 계 확인
                    cell_category = sheet.cell(row=row, column=category_col)
                    is_total = False
                    if cell_category.value:
                        category_str = str(cell_category.value).strip()
                        if category_str == '총지수' or category_str == '계' or category_str == '   계':
                            is_total = True
                    
                    # 분류 단계가 0 또는 1인 행 찾기 (일부 시도는 분류 단계가 1)
                    if classification_col:
                        cell_class = sheet.cell(row=row, column=classification_col)
                        if cell_class.value is not None:
                            try:
                                classification_level = float(cell_class.value)
                                if is_total and classification_level <= 1:  # 0 또는 1
                                    region_row = row
                                    break
                            except (ValueError, TypeError):
                                pass
                    elif is_total:
                        # 분류 단계 열이 없으면 총지수/계만 확인
                        region_row = row
                        break
        
        # 여전히 None이면 data_start_row 사용
        if region_row is None:
            structure = self.dynamic_parser.parse_sheet_structure(sheet_name)
            region_col = structure['region_column']
            data_start_row = structure['data_start_row']
            
            # data_start_row에서 지역 확인
            cell_region = sheet.cell(row=data_start_row, column=region_col)
            if cell_region.value and str(cell_region.value).strip() == region_name:
                region_row = data_start_row
        
        if region_row:
            # 헤더 기반으로 증감률 계산
            region_growth_rate = self.dynamic_parser.calculate_growth_rate(sheet_name, region_name, year, quarter)
            
            # 헤더 기반으로 찾지 못한 경우, 기존 로직으로 계산 (하위 호환성)
            if region_growth_rate is None:
                current = sheet.cell(row=region_row, column=current_col).value
                prev = sheet.cell(row=region_row, column=prev_col).value
                
                # 가중치 확인 - 가중치 열이 없으면 원데이터로 비교
                weight_col = structure.get('weight_column')
                has_weight_column = weight_col is not None
                
                if has_weight_column:
                    weight = sheet.cell(row=region_row, column=weight_col).value
                    if weight is None or (isinstance(weight, str) and not weight.strip()):
                        has_weight_column = False
                    else:
                        try:
                            weight = float(weight)
                        except (ValueError, TypeError):
                            has_weight_column = False
                
                # 결측치 체크 - 결측치는 1로 채움
                if current is None or (isinstance(current, str) and (not current.strip() or current.strip() == '-')):
                    current = prev if prev is not None else 1.0
                if prev is None or (isinstance(prev, str) and (not prev.strip() or prev.strip() == '-')):
                    prev = current if current is not None else 1.0
                
                # 결측치 체크
                if current is not None and prev is not None:
                    # 빈 문자열 또는 "-" 체크
                    if isinstance(current, str):
                        if not current.strip() or current.strip() == '-':
                            current = prev if prev is not None else 1.0
                    if isinstance(prev, str):
                        if not prev.strip() or prev.strip() == '-':
                            prev = current if current is not None else 1.0
                    
                    try:
                        prev_num = float(prev)
                        current_num = float(current)
                        
                        # 값이 1이고 둘 다 1이면 스킵 (기본값)
                        if current_num == 1.0 and prev_num == 1.0:
                            region_growth_rate = None
                        elif prev_num != 0:
                            region_growth_rate = ((current_num / prev_num) - 1) * 100
                            # NaN이나 Infinity 체크
                            import math
                            if math.isnan(region_growth_rate) or math.isinf(region_growth_rate):
                                region_growth_rate = None
                    except (ValueError, TypeError, ZeroDivisionError, OverflowError):
                        region_growth_rate = None
        
        if region_row is None:
            return []
        
        categories = []
        
        # 연령별 인구이동 시트의 경우 특별 처리
        if sheet_name == '연령별 인구이동' or sheet_name == '시도 간 이동':
            return self._get_age_groups_for_region('연령별 인구이동', region_name, year, quarter, top_n)
        
        # 가중치 열 찾기 - 가중치 열이 없으면 원데이터로 비교
        weight_col = structure.get('weight_column')
        has_weight_column = weight_col is not None
        
        # 해당 지역의 산업/업태별 데이터 찾기
        for row in range(region_row + 1, min(region_row + 500, sheet.max_row + 1)):
            cell_b = sheet.cell(row=row, column=2)  # 지역 이름
            cell_c = sheet.cell(row=row, column=3)  # 분류 단계
            cell_category = sheet.cell(row=row, column=category_col)  # 산업/업태 이름
            
            # 같은 지역인지 확인 (지역명 매핑 고려)
            is_same_region = False
            cell_b_str = str(cell_b.value).strip() if cell_b.value else ''
            
            # 지역명 매핑 확인
            region_mapping = REGION_MAPPING
            
            mapped_region_name = region_mapping.get(region_name, region_name)
            mapped_cell_b = region_mapping.get(cell_b_str, cell_b_str)
            
            if (cell_b_str == region_name or 
                cell_b_str == mapped_region_name or
                mapped_cell_b == region_name or
                mapped_cell_b == mapped_region_name):
                is_same_region = True
            
            # 같은 지역이고 분류 단계가 1 이상인 것 (산업/업태)
            if is_same_region and cell_c.value:
                try:
                    classification_level = float(cell_c.value) if cell_c.value else 0
                except (ValueError, TypeError):
                    classification_level = 0
                
                if classification_level >= 1 and cell_category.value:
                    current = sheet.cell(row=row, column=current_col).value
                    prev = sheet.cell(row=row, column=prev_col).value
                    
                    # 가중치 확인 - 가중치 열이 없으면 원데이터로 비교
                    weight = 1.0
                    if has_weight_column:
                        weight_value = sheet.cell(row=row, column=weight_col).value
                        if weight_value is not None and not (isinstance(weight_value, str) and not weight_value.strip()):
                            try:
                                weight = float(weight_value)
                            except (ValueError, TypeError):
                                weight = 1.0
                    
                    # 결측치 체크 - 결측치는 1로 채움
                    if current is None or (isinstance(current, str) and (not current.strip() or current.strip() == '-')):
                        current = prev if prev is not None else 1.0
                    if prev is None or (isinstance(prev, str) and (not prev.strip() or prev.strip() == '-')):
                        prev = current if current is not None else 1.0
                    
                    # 빈 문자열 또는 "-" 체크
                    if isinstance(current, str):
                        if not current.strip() or current.strip() == '-':
                            current = prev if prev is not None else 1.0
                    if isinstance(prev, str):
                        if not prev.strip() or prev.strip() == '-':
                            prev = current if current is not None else 1.0
                    
                    try:
                        prev_num = float(prev)
                        current_num = float(current)
                        
                        # 값이 1이고 둘 다 1이면 스킵 (기본값)
                        if current_num == 1.0 and prev_num == 1.0:
                            continue
                        
                        if prev_num != 0:
                            growth_rate = ((current_num / prev_num) - 1) * 100
                            # NaN이나 Infinity 체크
                            import math
                            if not (math.isnan(growth_rate) or math.isinf(growth_rate)):
                                categories.append({
                                    'name': str(cell_category.value).strip(),
                                    'growth_rate': growth_rate,
                                    'row': row,
                                    'current': current,
                                    'prev': prev,
                                    'weight': weight
                                })
                    except (ValueError, TypeError, ZeroDivisionError, OverflowError):
                        continue
            else:
                # 다른 지역이 나오면 중단 (지역명 매핑 고려)
                if cell_b.value:
                    cell_b_str = str(cell_b.value).strip()
                    region_mapping = {
                        '서울특별시': '서울', '부산광역시': '부산', '대구광역시': '대구',
                        '인천광역시': '인천', '광주광역시': '광주', '대전광역시': '대전',
                        '울산광역시': '울산', '세종특별자치시': '세종', '경기도': '경기',
                        '강원도': '강원', '충청북도': '충북', '충청남도': '충남',
                        '전라북도': '전북', '전라남도': '전남', '경상북도': '경북',
                        '경상남도': '경남', '제주특별자치도': '제주'
                    }
                    mapped_region_name = region_mapping.get(region_name, region_name)
                    mapped_cell_b = region_mapping.get(cell_b_str, cell_b_str)
                    
                    if (cell_b_str != region_name and 
                        cell_b_str != mapped_region_name and
                        mapped_cell_b != region_name and
                        mapped_cell_b != mapped_region_name):
                        break
        
        # 전국인 경우: 시트별 설정에 따라 처리
        if region_name == '전국':
            national_priorities = config.get('national_priorities')
            if national_priorities:
                # 감소한 산업/업태만 필터링
                negative_categories = [c for c in categories if c['growth_rate'] < 0]
                
                # 우선순위에 따라 선택
                result = []
                for priority_name in national_priorities:
                    for cat in negative_categories:
                        if priority_name in cat['name'] and cat not in result:
                            result.append(cat)
                            break
                
                return result[:top_n]
            else:
                # 우선순위가 없으면 절대값 기준 정렬
                categories.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
                return categories[:top_n]
        
        # 지역별로 증가/감소에 따라 필터링
        if region_growth_rate is not None:
            if region_growth_rate > 0:
                # 증가한 지역: 증가한 산업/업태만 선택, 증가율이 큰 순서
                positive_categories = [c for c in categories if c['growth_rate'] > 0]
                positive_categories.sort(key=lambda x: x['growth_rate'], reverse=True)
                return positive_categories[:top_n]
            else:
                # 감소한 지역: 감소한 산업/업태만 선택, 지역별 우선순위 적용
                negative_categories = [c for c in categories if c['growth_rate'] < 0]
                
                # 시트별 지역 우선순위 가져오기
                region_priorities = config.get('region_priorities', {})
                priority_list = region_priorities.get(region_name, [])
                
                if priority_list:
                    # 우선순위에 따라 선택
                    result = []
                    for priority_keyword in priority_list:
                        for cat in negative_categories:
                            if priority_keyword in cat['name'] and cat not in result:
                                result.append(cat)
                                break
                    
                    # 우선순위에 없는 산업/업태는 절대값 기준으로 추가
                    remaining = [c for c in negative_categories if c not in result]
                    remaining.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
                    result.extend(remaining)
                    
                    return result[:top_n]
                else:
                    # 우선순위가 없으면 절대값 기준 정렬
                    negative_categories.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
                    return negative_categories[:top_n]
        
        # 기본: 증감률 절대값 기준 정렬
        categories.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
        return categories[:top_n]
    
    def _get_quarterly_growth_rate(self, sheet_name: str, region_name: str, quarter_key: str) -> Optional[float]:
        """
        특정 지역의 특정 분기 증감률을 계산합니다.
        헤더 기반으로 동적으로 계산합니다.
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름 (예: '전국', '서울')
            quarter_key: 분기 문자열 (예: '2023_3분기', '2024_1분기')
            
        Returns:
            증감률 (퍼센트) 또는 None
        """
        # 분기 키에서 연도와 분기 추출
        match = re.match(r'(\d{4})_(\d)분기', quarter_key)
        if not match:
            return None
        
        year = int(match.group(1))
        quarter = int(match.group(2))
        
        # 헤더 기반으로 증감률 계산 시도
        if sheet_name:
            growth_rate = self.dynamic_parser.calculate_growth_rate(sheet_name, region_name, year, quarter)
            if growth_rate is not None:
                return growth_rate
        
        # 헤더 기반으로 찾지 못한 경우, 기존 로직 사용
        # 열 번호 계산
        current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
        
        # 수출/수입 시트의 경우, 표에 표시되는 값은 이미 증감률이 계산된 값이므로
        # 직접 해당 열에서 값을 읽어야 함 (전년 동분기와 비교하는 것이 아님)
        # 하지만 템플릿에서는 "전년동분기비"라고 명시되어 있으므로, 
        # 엑셀의 해당 열이 이미 증감률인지 확인 필요
        # 일단 기존 로직 유지하되, 수출 시트의 경우 특별 처리
        
        # 지역의 총지수 행 찾기
        sheet = self.excel_extractor.get_sheet(sheet_name)
        config = self._get_sheet_config(sheet_name)
        category_col = config['category_column']
        
        # 가중치 열 찾기 - 가중치 열이 없으면 원데이터로 비교
        structure = self.dynamic_parser.parse_sheet_structure(sheet_name)
        weight_col = structure.get('weight_column')
        has_weight_column = weight_col is not None
        
        region_row = None
        
        for row in range(4, min(5000, sheet.max_row + 1)):
            cell_b = sheet.cell(row=row, column=2)  # 지역 이름
            cell_c = sheet.cell(row=row, column=3)  # 분류 단계
            cell_category = sheet.cell(row=row, column=category_col)  # 업태/산업 이름
            
            # 지역 이름이 일치하는지 확인
            if not cell_b.value:
                continue
                
            cell_b_str = str(cell_b.value).strip()
            if cell_b_str != region_name:
                continue
            
            # 총지수 또는 계 확인
            is_total = False
            if cell_category.value:
                category_str = str(cell_category.value).strip()
                if category_str == '총지수' or category_str == '계' or category_str == '   계':
                    is_total = True
            
            # 분류 단계 확인 (0 또는 1 모두 허용)
            classification_level = None
            if cell_c.value is not None:
                try:
                    classification_level = float(cell_c.value)
                except (ValueError, TypeError):
                    pass
            
            # 수출/수입 시트의 경우: 분류 단계가 없거나 0이거나, 총지수/계인 경우
            # 또는 category_col이 비어있거나 None인 경우도 허용 (수출 시트는 구조가 다를 수 있음)
            if sheet_name == '수출' or sheet_name == '수입':
                # 수출/수입 시트는 지역 이름만으로 찾기 (분류 단계나 카테고리 체크 완화)
                if is_total or (classification_level is None or classification_level == 0) or (cell_category.value is None or str(cell_category.value).strip() == ''):
                    region_row = row
                    break
            else:
                # 다른 시트는 총지수/계이고 분류 단계가 0 또는 1인 경우
                if is_total and (classification_level is None or classification_level <= 1):
                    region_row = row
                    break
        
        if region_row is None:
            return None
        
        # 수출/수입 시트의 경우, 엑셀에 이미 증감률이 계산되어 있을 수 있음
        # 표의 데이터를 보면 이미 증감률 값이므로, 해당 열에서 직접 값을 읽어야 함
        # 하지만 템플릿에서는 "전년동분기비"라고 명시되어 있으므로,
        # 엑셀의 해당 열이 이미 증감률인지 확인 필요
        # 일단 기존 로직대로 전년 동분기와 비교하여 계산
        
        # 가중치 열 찾기 - 가중치 열이 없으면 원데이터로 비교
        weight_col = structure.get('weight_column')
        has_weight_column = weight_col is not None
        
        # 현재 분기와 전년 동분기 값 가져오기
        current_value = sheet.cell(row=region_row, column=current_col).value
        prev_value = sheet.cell(row=region_row, column=prev_col).value
        
        # 결측치 체크 - 결측치는 1로 채움
        if current_value is None or (isinstance(current_value, str) and (not current_value.strip() or current_value.strip() == '-')):
            current_value = prev_value if prev_value is not None else 1.0
        if prev_value is None or (isinstance(prev_value, str) and (not prev_value.strip() or prev_value.strip() == '-')):
            prev_value = current_value if current_value is not None else 1.0
        
        # 결측치 체크
        if current_value is None or prev_value is None:
            return None
        
        # 수출/수입 시트의 경우, 값이 이미 증감률(%)일 수 있음
        # 음수나 작은 값(-100 ~ 100 범위)이면 이미 증감률로 간주
        if sheet_name == '수출' or sheet_name == '수입':
            try:
                current_num = float(current_value)
                # 이미 증감률인 경우 (일반적으로 -100 ~ 100 범위)
                if -100 <= current_num <= 100:
                    # 전년 동분기 값도 확인
                    try:
                        prev_num = float(prev_value)
                        # 전년 동분기도 증감률 범위면, 현재 값을 그대로 반환
                        if -100 <= prev_num <= 100:
                            return current_num
                    except (ValueError, TypeError):
                        pass
                    # 전년 동분기가 증감률이 아니면, 현재 값이 이미 증감률일 가능성
                    # 하지만 안전을 위해 계산도 시도
            except (ValueError, TypeError):
                pass
        
        # 0으로 나누기 방지
        try:
            prev_num = float(prev_value)
            if prev_num == 0:
                return None
            current_num = float(current_value)
        except (ValueError, TypeError):
            return None
        
        # 증감률 계산
        try:
            growth_rate = ((current_num / prev_num) - 1) * 100
            # NaN이나 Infinity 체크
            import math
            if math.isnan(growth_rate) or math.isinf(growth_rate):
                return None
            return growth_rate
        except (ZeroDivisionError, OverflowError):
            return None
    
    def _process_dynamic_marker(self, sheet_name: str, key: str, year: int = 2025, quarter: int = 2) -> Optional[str]:
        """
        동적 마커를 처리합니다.
        N/A가 발생할 경우 자동으로 추론 로직을 시도합니다.
        
        Args:
            sheet_name: 시트 이름
            key: 동적 키 (예: '상위시도1_이름', '상위시도1_증감률')
            year: 연도
            quarter: 분기
            
        Returns:
            처리된 값 또는 None
        """
        # 가상 시트명을 실제 시트명으로 변환 (실업률 -> 실업자 수)
        sheet_name = self._get_actual_sheet_name(sheet_name)
        
        self._analyze_data_if_needed(sheet_name, year, quarter)
        
        # 시트별로 다른 캐시 키 사용
        cache_key = f"{sheet_name}_{year}_{quarter}"
        if not hasattr(self, '_current_analyzed_data') or cache_key not in self._current_analyzed_data:
            # 캐시에 데이터가 없으면 빈 딕셔너리로 초기화하고 직접 계산 시도
            if not hasattr(self, '_current_analyzed_data'):
                self._current_analyzed_data = {}
            self._current_analyzed_data[cache_key] = {}
            # 직접 계산 시도 (추론 로직)
            inferred_value = self._try_infer_marker_value(sheet_name, key, year, quarter)
            if inferred_value is not None:
                return inferred_value
            return "N/A"
        
        data = self._current_analyzed_data[cache_key]
        
        # 전국 패턴
        if key == '전국_이름':
            if 'national_region' in data and data['national_region']:
                return data['national_region']['name']
            # 추론 시도
            inferred = self._try_infer_marker_value(sheet_name, key, year, quarter)
            if inferred is not None:
                return inferred
            return "N/A"
        elif key == '전국_증감률':
            # 실업률/고용률 시트인 경우 직접 처리 (아래 로직 사용)
            is_unemployment_sheet = ('실업' in sheet_name or '고용' in sheet_name)
            if not is_unemployment_sheet:
                # 일반 시트는 캐시에서 가져오기
                if 'national_region' in data and data['national_region']:
                    return self.format_percentage(data['national_region']['growth_rate'], decimal_places=1, include_percent=False)
                # 추론 시도
                inferred = self._try_infer_marker_value(sheet_name, key, year, quarter)
                if inferred is not None:
                    return inferred
                return "N/A"
            # 실업률/고용률 시트는 아래 로직으로 처리 (계속 진행)
        elif key == '전국_증감방향':
            # 전국 증감률에 따라 "증가" 또는 "감소" 반환
            if 'national_region' in data and data['national_region']:
                growth_rate = data['national_region']['growth_rate']
                return self.get_growth_direction(growth_rate)
            # 추론 시도: 전국 증감률 먼저 계산
            growth_rate = self.dynamic_parser.calculate_growth_rate(sheet_name, '전국', year, quarter)
            if growth_rate is not None:
                return self.get_growth_direction(growth_rate)
            return "N/A"
        elif key == '전국_변화표현':
            # 전국 증감률에 따라 "올라" 또는 "내려" 반환
            if 'national_region' in data and data['national_region']:
                growth_rate = data['national_region']['growth_rate']
                return self.get_production_change_expression(growth_rate, direction_type="rise_fall")
            # 추론 시도
            growth_rate = self.dynamic_parser.calculate_growth_rate(sheet_name, '전국', year, quarter)
            if growth_rate is not None:
                return self.get_production_change_expression(growth_rate, direction_type="rise_fall")
            return "N/A"
        elif key == '전국_방향':
            # 전국 증감률에 따라 "상승" 또는 "하강" 반환
            if 'national_region' in data and data['national_region']:
                growth_rate = data['national_region']['growth_rate']
                return self.get_growth_direction(growth_rate, direction_type="rise_fall", expression_key="rate")
            # 추론 시도
            growth_rate = self.dynamic_parser.calculate_growth_rate(sheet_name, '전국', year, quarter)
            if growth_rate is not None:
                return self.get_growth_direction(growth_rate, direction_type="rise_fall", expression_key="rate")
            return "N/A"
        
        # 전국 산업/업태 마커는 연령 패턴보다 먼저 처리 (패턴 충돌 방지)
        elif key.startswith('전국_업태') or key.startswith('전국_산업'):
            # 전국 산업/업태 마커 처리 (일반화)
            industry_match = re.match(r'전국_(업태|산업)(\d+)_(.+)', key)
            if industry_match:
                industry_type = industry_match.group(1)  # '업태' 또는 '산업'
                industry_idx = int(industry_match.group(2)) - 1
                industry_field = industry_match.group(3)
                
                # 물가 시트 여부 확인 (물가동향은 4개 산업이 필요)
                is_price_sheet = ('물가' in sheet_name)
                top_n = 4 if is_price_sheet else 3
                
                # 분석된 데이터의 national_region.top_industries 사용 (이미 계산되어 있음)
                categories = []
                if 'national_region' in data and data['national_region']:
                    categories = data['national_region'].get('top_industries', [])
                
                # top_industries가 비어있으면 _get_categories_for_region 호출
                if not categories:
                    categories = self._get_categories_for_region(sheet_name, '전국', year, quarter, top_n=top_n)
                    # 여전히 비어있으면 다른 방법 시도
                    if not categories:
                        current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                        if current_col and prev_col:
                            try:
                                industries_data = self.data_analyzer.get_top_industries_for_region(
                                    sheet_name, '전국', None, current_col, prev_col, top_n=top_n
                                )
                                if industries_data:
                                    categories = industries_data
                            except:
                                pass
                
                if categories and industry_idx < len(categories):
                    category = categories[industry_idx]
                    
                    if industry_field == '이름':
                        category_name = category['name']
                        config = self._get_sheet_config(sheet_name)
                        name_mapping = config.get('name_mapping', {})
                        mapped_name = name_mapping.get(category_name)
                        if not mapped_name:
                            # 부분 일치 확인
                            for key_map, value_map in name_mapping.items():
                                if key_map in category_name or category_name in key_map:
                                    mapped_name = value_map
                                    break
                        # 품목명 표시 매핑 적용 (엑셀 이름 -> 표시용 이름)
                        if not mapped_name:
                            item_display_mapping = self.schema_loader.get_name_mapping('item_display_mapping')
                            if item_display_mapping:
                                mapped_name = item_display_mapping.get(category_name)
                        return mapped_name if mapped_name else category_name
                    elif industry_field == '증감률':
                        return self.format_percentage(category['growth_rate'], decimal_places=1, include_percent=False)
                return "N/A"
            return "N/A"
        
        # 전국 연령별 증감률/증감pp 마커 처리 (실업률용)
        # 예: 전국_60세이상_증감률, 전국_30_59세_증감pp, 전국_15_29세_증감pp
        # 증감pp: 퍼센트포인트 차이 (현재값 - 이전값)
        national_age_match = re.match(r'^전국_(.+)_증감(?:률|pp)$', key)
        if national_age_match:
            age_group_key = national_age_match.group(1)
            
            # 실업률 시트 여부 확인
            is_unemployment_sheet = (sheet_name == '실업률' or sheet_name == '실업자 수')
            
            if is_unemployment_sheet:
                # 연령대 이름 매핑 (마커 이름 -> 엑셀 이름 리스트)
                # Excel에서 "30 - 59세" 형식을 사용할 수 있음
                age_mapping = {
                    '60세이상': ['60세이상', '60 세이상', '60세 이상'],
                    '30_59세': ['30~59세', '30 - 59세', '30-59세', '30 ~ 59세'],
                    '15_29세': ['15~29세', '15 - 29세', '15-29세', '15 ~ 29세']
                }
                target_ages = age_mapping.get(age_group_key, [age_group_key])
                
                # 전국의 연령별 데이터 가져오기
                age_groups = self._get_unemployment_rate_age_groups('전국', year, quarter)
                
                for age_group in age_groups:
                    age_name = age_group.get('name', '')
                    for target_age in target_ages:
                        if age_name == target_age or target_age in age_name or age_name in target_age:
                            return self.format_percentage(age_group.get('growth_rate', 0), decimal_places=1, include_percent=False)
                
            return "N/A"
        
        elif key == '강조_시도수':
            # 전국이 증가면 증가_시도수, 감소면 감소_시도수
            national_growth = None
            if 'national_region' in data and data['national_region']:
                national_growth = data['national_region']['growth_rate']
            else:
                # 추론 시도
                national_growth = self.dynamic_parser.calculate_growth_rate(sheet_name, '전국', year, quarter)
            
            if national_growth is not None:
                if national_growth >= 0:
                    return self._process_dynamic_marker(sheet_name, '증가_시도수', year, quarter)
                else:
                    return self._process_dynamic_marker(sheet_name, '감소_시도수', year, quarter)
            return "N/A"
        elif key == '강조_증감방향':
            # 전국이 증가면 "증가", 감소면 "감소"
            growth_rate = None
            if 'national_region' in data and data['national_region']:
                growth_rate = data['national_region']['growth_rate']
            else:
                # 추론 시도
                growth_rate = self.dynamic_parser.calculate_growth_rate(sheet_name, '전국', year, quarter)
            
            if growth_rate is not None:
                return self.get_growth_direction(growth_rate)
            return "N/A"
        elif key.startswith('강조시도'):
            # 강조시도N_이름, 강조시도N_증감률, 강조시도N_산업M_이름 등 처리
            # 전국이 증가면 상위시도, 감소면 하위시도
            emphasis_match = re.match(r'강조시도(\d+)_(.+)', key)
            if emphasis_match:
                idx = int(emphasis_match.group(1)) - 1  # 0-based
                field = emphasis_match.group(2)
                
                # 전국 증감률에 따라 상위/하위 결정
                national_growth = None
                if 'national_region' in data and data['national_region']:
                    national_growth = data['national_region']['growth_rate']
                else:
                    # 추론 시도
                    national_growth = self.dynamic_parser.calculate_growth_rate(sheet_name, '전국', year, quarter)
                
                if national_growth is not None:
                    if national_growth >= 0:
                        # 상위 시도 사용
                        regions = data.get('top_regions', [])
                        if not regions:
                            # 직접 계산
                            current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                            if current_col and prev_col:
                                regions_data = self.data_analyzer.get_regions_with_growth_rate(sheet_name, current_col, prev_col)
                                if regions_data:
                                    regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0), reverse=True)
                    else:
                        # 하위 시도 사용
                        regions = data.get('bottom_regions', [])
                        if not regions:
                            # 직접 계산
                            current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                            if current_col and prev_col:
                                regions_data = self.data_analyzer.get_regions_with_growth_rate(sheet_name, current_col, prev_col)
                                if regions_data:
                                    regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0))
                    
                    if regions and idx < len(regions):
                        region = regions[idx]
                        
                        if field == '이름':
                            return region['name']
                        elif field == '증감률':
                            return self.format_percentage(region['growth_rate'], decimal_places=1, include_percent=False)
                        elif field.startswith('업태') or field.startswith('산업'):
                            industry_match = re.match(r'(업태|산업)(\d+)_(.+)', field)
                            if industry_match:
                                industry_idx = int(industry_match.group(2)) - 1
                                industry_field = industry_match.group(3)
                                
                                # 분석된 데이터의 top_industries 사용 (이미 계산되어 있음)
                                categories = region.get('top_industries', [])
                                
                                # top_industries가 비어있으면 _get_categories_for_region 호출
                                if not categories:
                                    categories = self._get_categories_for_region(
                                        sheet_name, region.get('name', ''), year, quarter, top_n=3
                                    )
                                    # 여전히 비어있으면 직접 계산 시도
                                    if not categories:
                                        current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                                        if current_col and prev_col:
                                            try:
                                                region_row = None
                                                sheet = self.excel_extractor.get_sheet(sheet_name)
                                                structure = self.dynamic_parser.parse_sheet_structure(sheet_name)
                                                region_col = structure.get('region_column', 2)
                                                data_start_row = structure.get('data_start_row', 4)
                                                for r in range(data_start_row, min(data_start_row + 200, sheet.max_row + 1)):
                                                    cell = sheet.cell(row=r, column=region_col)
                                                    if cell.value and str(cell.value).strip() == region.get('name', ''):
                                                        region_row = r
                                                        break
                                                if region_row:
                                                    industries_data = self.data_analyzer.get_top_industries_for_region(
                                                        sheet_name, region.get('name', ''), region_row, current_col, prev_col, top_n=3
                                                    )
                                                    if industries_data:
                                                        categories = industries_data
                                            except:
                                                pass
                                
                                if categories and industry_idx < len(categories):
                                    category = categories[industry_idx]
                                    
                                    if industry_field == '이름':
                                        category_name = category['name']
                                        config = self._get_sheet_config(sheet_name)
                                        name_mapping = config.get('name_mapping', {})
                                        mapped_name = name_mapping.get(category_name)
                                        if not mapped_name:
                                            for key_map, value_map in name_mapping.items():
                                                if key_map in category_name or category_name in key_map:
                                                    mapped_name = value_map
                                                    break
                                        # 품목명 표시 매핑 적용 (엑셀 이름 -> 표시용 이름)
                                        if not mapped_name:
                                            item_display_mapping = self.schema_loader.get_name_mapping('item_display_mapping')
                                            if item_display_mapping:
                                                mapped_name = item_display_mapping.get(category_name)
                                        return mapped_name if mapped_name else category_name
                                    elif industry_field == '증감률':
                                        return self.format_percentage(category['growth_rate'], decimal_places=1, include_percent=False)
                            return "N/A"
                return "N/A"
            return "N/A"
        
        # 상위 시도 패턴
        top_match = re.match(r'상위시도(\d+)_(.+)', key)
        if top_match:
            idx = int(top_match.group(1)) - 1  # 0-based
            field = top_match.group(2)
            
            # 실업률 시트 여부 확인
            is_unemployment_sheet = (sheet_name == '실업률' or sheet_name == '실업자 수')
            # 물가 시트 여부 확인
            is_price_sheet = ('물가' in sheet_name)
            
            regions = []
            if is_unemployment_sheet:
                # 실업률 시트: 전용 함수로 백분율포인트 단위 증감률 계산 (캐시 무시)
                regions_data = self._get_unemployment_rate_regions_data(year, quarter)
                if regions_data:
                    # 상위시도 = 실업률이 가장 많이 상승한 지역 (양수가 큰 순서)
                    regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0), reverse=True)
            elif is_price_sheet:
                # 물가 시트: 항상 '품목성질별 물가' 시트 기준으로 상위시도 결정
                price_base_sheet = '품목성질별 물가'
                current_col, prev_col = self._get_quarter_columns(year, quarter, price_base_sheet)
                if current_col and prev_col:
                    regions_data = self.data_analyzer.get_regions_with_growth_rate(price_base_sheet, current_col, prev_col)
                    if regions_data:
                        regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0), reverse=True)
            else:
                # 일반 시트: 캐시 우선 사용
                regions = data.get('top_regions', [])
                if not regions:
                    current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                    if current_col and prev_col:
                        regions_data = self.data_analyzer.get_regions_with_growth_rate(sheet_name, current_col, prev_col)
                        if regions_data:
                            regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0), reverse=True)
            
            if regions and idx < len(regions):
                region = regions[idx]
                
                if field == '이름':
                    return region.get('name', '')
                elif field == '증감률' or field == '증감pp':
                    return self.format_percentage(region.get('growth_rate', 0), decimal_places=1, include_percent=False)
                elif field == '방향':
                    # 상위시도 증감률에 따라 "상승" 또는 "하강" 반환
                    return self.get_growth_direction(region['growth_rate'], direction_type="rise_fall", expression_key="rate")
                elif field == '변화표현':
                    # 상위시도 증감률에 따라 "올라" 또는 "내려" 반환
                    return self.get_production_change_expression(region['growth_rate'], direction_type="rise_fall")
                elif field.startswith('업태') or field.startswith('산업'):
                    # 업태1_이름, 업태1_증감률, 산업1_이름, 산업1_증감률 등 (일반화)
                    industry_match = re.match(r'(업태|산업)(\d+)_(.+)', field)
                    if industry_match:
                        industry_type = industry_match.group(1)  # '업태' 또는 '산업'
                        industry_idx = int(industry_match.group(2)) - 1
                        industry_field = industry_match.group(3)
                        
                        # 물가 시트의 경우: 해당 시도의 품목 데이터는 해당 시트에서 가져옴
                        # 분석된 데이터의 top_industries 사용 (이미 계산되어 있음)
                        categories = region.get('top_industries', [])
                        
                        # top_industries가 비어있으면 _get_categories_for_region 호출
                        if not categories:
                            # 물가 시트의 경우 현재 시트(지출목적별 물가 등)에서 품목 데이터 가져옴
                            categories = self._get_categories_for_region(
                                sheet_name, region['name'], year, quarter, top_n=4
                            )
                        
                        if categories and industry_idx < len(categories):
                            category = categories[industry_idx]
                            
                            if industry_field == '이름':
                                category_name = category['name']
                                config = self._get_sheet_config(sheet_name)
                                name_mapping = config.get('name_mapping', {})
                                mapped_name = name_mapping.get(category_name)
                                if not mapped_name:
                                    # 부분 일치 확인
                                    for key_map, value_map in name_mapping.items():
                                        if key_map in category_name or category_name in key_map:
                                            mapped_name = value_map
                                            break
                                # 품목명 표시 매핑 적용 (엑셀 이름 -> 표시용 이름)
                                if not mapped_name:
                                    item_display_mapping = self.schema_loader.get_name_mapping('item_display_mapping')
                                    if item_display_mapping:
                                        mapped_name = item_display_mapping.get(category_name)
                                return mapped_name if mapped_name else category_name
                            elif industry_field == '증감률':
                                return self.format_percentage(category['growth_rate'], decimal_places=1, include_percent=False)
                    return "N/A"
                elif field.startswith('연령'):
                    # 연령1_이름, 연령1_증감률, 연령2_있음? 등 (실업률 시트용)
                    age_match = re.match(r'연령(\d+)_(.+)', field)
                    if age_match:
                        age_idx = int(age_match.group(1)) - 1
                        age_field = age_match.group(2)
                        
                        # 실업률 시트인 경우 연령별 데이터 가져오기
                        if is_unemployment_sheet:
                            age_groups = self._get_unemployment_rate_age_groups(region['name'], year, quarter)
                            
                            if age_field == '있음?':
                                # 해당 연령대 존재 여부 확인 (조건부 출력용)
                                return '' if age_idx < len(age_groups) else 'N/A'
                            
                            if age_groups and age_idx < len(age_groups):
                                age_group = age_groups[age_idx]
                                if age_field == '이름':
                                    return age_group.get('name', '')
                                elif age_field in ('증감률', '증감pp'):
                                    return self.format_percentage(age_group.get('growth_rate', 0), decimal_places=1, include_percent=False)
                    return "N/A"
                return "N/A"
            return "N/A"
        
        # 하위 시도 패턴
        bottom_match = re.match(r'하위시도(\d+)_(.+)', key)
        if bottom_match:
            idx = int(bottom_match.group(1)) - 1  # 0-based
            field = bottom_match.group(2)
            
            # 실업률 시트 여부 확인
            is_unemployment_sheet = (sheet_name == '실업률' or sheet_name == '실업자 수')
            # 물가 시트 여부 확인
            is_price_sheet = ('물가' in sheet_name)
            
            regions = []
            if is_unemployment_sheet:
                # 실업률 시트: 전용 함수로 백분율포인트 단위 증감률 계산 (캐시 무시)
                regions_data = self._get_unemployment_rate_regions_data(year, quarter)
                if regions_data:
                    # 하위시도 = 실업률이 가장 많이 하락한 지역 (음수가 큰 순서 = 오름차순)
                    regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0))
            elif is_price_sheet:
                # 물가 시트: 항상 '품목성질별 물가' 시트 기준으로 하위시도 결정
                price_base_sheet = '품목성질별 물가'
                current_col, prev_col = self._get_quarter_columns(year, quarter, price_base_sheet)
                if current_col and prev_col:
                    regions_data = self.data_analyzer.get_regions_with_growth_rate(price_base_sheet, current_col, prev_col)
                    if regions_data:
                        regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0))
            else:
                # 일반 시트: 캐시 우선 사용
                regions = data.get('bottom_regions', [])
                if not regions:
                    current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                    if current_col and prev_col:
                        regions_data = self.data_analyzer.get_regions_with_growth_rate(sheet_name, current_col, prev_col)
                        if regions_data:
                            regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0))
            
            if regions and idx < len(regions):
                region = regions[idx]
                
                if field == '이름':
                    return region.get('name', '')
                elif field == '증감률' or field == '증감pp':
                    return self.format_percentage(region.get('growth_rate', 0), decimal_places=1, include_percent=False)
                elif field == '방향':
                    # 하위시도 증감률에 따라 "상승" 또는 "하강" 반환
                    return self.get_growth_direction(region['growth_rate'], direction_type="rise_fall", expression_key="rate")
                elif field == '변화표현':
                    # 하위시도 증감률에 따라 "올라" 또는 "내려" 반환
                    return self.get_production_change_expression(region['growth_rate'], direction_type="rise_fall")
                elif field.startswith('업태') or field.startswith('산업'):
                    # 업태1_이름, 업태1_증감률, 산업1_이름, 산업1_증감률 등 (일반화)
                    industry_match = re.match(r'(업태|산업)(\d+)_(.+)', field)
                    if industry_match:
                        industry_type = industry_match.group(1)  # '업태' 또는 '산업'
                        industry_idx = int(industry_match.group(2)) - 1
                        industry_field = industry_match.group(3)
                        
                        # 분석된 데이터의 top_industries 사용 (이미 계산되어 있음)
                        categories = region.get('top_industries', [])
                        
                        # top_industries가 비어있으면 _get_categories_for_region 호출
                        if not categories:
                            # 물가 시트의 경우 현재 시트에서 품목 데이터 가져옴
                            categories = self._get_categories_for_region(
                                sheet_name, region['name'], year, quarter, top_n=4
                            )
                        
                        if categories and industry_idx < len(categories):
                            category = categories[industry_idx]
                            
                            if industry_field == '이름':
                                category_name = category['name']
                                config = self._get_sheet_config(sheet_name)
                                name_mapping = config.get('name_mapping', {})
                                mapped_name = name_mapping.get(category_name)
                                if not mapped_name:
                                    # 부분 일치 확인
                                    for key_map, value_map in name_mapping.items():
                                        if key_map in category_name or category_name in key_map:
                                            mapped_name = value_map
                                            break
                                # 품목명 표시 매핑 적용 (엑셀 이름 -> 표시용 이름)
                                if not mapped_name:
                                    item_display_mapping = self.schema_loader.get_name_mapping('item_display_mapping')
                                    if item_display_mapping:
                                        mapped_name = item_display_mapping.get(category_name)
                                return mapped_name if mapped_name else category_name
                            elif industry_field == '증감률':
                                return self.format_percentage(category['growth_rate'], decimal_places=1, include_percent=False)
                            elif industry_field == '유입' or industry_field == '유출':
                                # 연령별 인구이동 시트의 경우, 연령 범위를 합산한 값을 반환
                                if sheet_name == '시도 간 이동' or sheet_name == '연령별 인구이동':
                                    return self._get_age_group_value(
                                        '연령별 인구이동', region['name'], category['name'], 
                                        industry_field, year, quarter
                                    )
                                # 일반적인 경우는 현재 값 반환
                                return str(int(category.get('current', 0))) if category.get('current') is not None else "N/A"
                    return "N/A"
                elif field.startswith('연령'):
                    # 연령1_이름, 연령1_증감률, 연령2_있음? 등 (실업률 시트용)
                    age_match = re.match(r'연령(\d+)_(.+)', field)
                    if age_match:
                        age_idx = int(age_match.group(1)) - 1
                        age_field = age_match.group(2)
                        
                        # 실업률 시트인 경우 연령별 데이터 가져오기
                        if is_unemployment_sheet:
                            age_groups = self._get_unemployment_rate_age_groups(region['name'], year, quarter)
                            
                            if age_field == '있음?':
                                # 해당 연령대 존재 여부 확인 (조건부 출력용)
                                return '' if age_idx < len(age_groups) else 'N/A'
                            
                            if age_groups and age_idx < len(age_groups):
                                age_group = age_groups[age_idx]
                                if age_field == '이름':
                                    return age_group.get('name', '')
                                elif age_field in ('증감률', '증감pp'):
                                    return self.format_percentage(age_group.get('growth_rate', 0), decimal_places=1, include_percent=False)
                    return "N/A"
                return "N/A"
            return "N/A"
        
        # 분기별 증감률 마커 처리 (예: 전국_2023_3분기_증감률)
        quarterly_match = re.match(r'(.+)_(\d{4})_(\d)분기_증감률', key)
        if quarterly_match:
            region_name = quarterly_match.group(1)
            year = quarterly_match.group(2)
            quarter_num = quarterly_match.group(3)
            quarter_key = f'{year}_{quarter_num}분기'
            
            growth_rate = self._get_quarterly_growth_rate(sheet_name, region_name, quarter_key)
            if growth_rate is not None:
                # 표 셀에는 % 기호 없이 표시
                return self.format_percentage(growth_rate, decimal_places=1, include_percent=False)
            # 추론 시도: 지역명 매핑 후 재시도
            region_mapping = REGION_MAPPING_REVERSE
            mapped_region = region_mapping.get(region_name, region_name)
            if mapped_region != region_name:
                growth_rate = self._get_quarterly_growth_rate(sheet_name, mapped_region, quarter_key)
                if growth_rate is not None:
                    return self.format_percentage(growth_rate, decimal_places=1, include_percent=False)
            return "N/A"
        
        # 지역별 분기별 증감률 마커 처리 (예: 전국_2023_3분기, 서울_2024_2분기)
        # 표 셀에 증감률 값을 % 없이 표시
        region_quarterly_match = re.match(r'^([가-힣]+)_(\d{4})_(\d)분기$', key)
        if region_quarterly_match:
            region_name = region_quarterly_match.group(1)
            target_year = int(region_quarterly_match.group(2))
            target_quarter = int(region_quarterly_match.group(3))
            quarter_key = f'{target_year}_{target_quarter}분기'
            
            growth_rate = self._get_quarterly_growth_rate(sheet_name, region_name, quarter_key)
            if growth_rate is not None:
                # 표 셀에는 % 기호 없이 표시
                return self.format_percentage(growth_rate, decimal_places=1, include_percent=False)
            # 추론 시도: 지역명 매핑 후 재시도
            region_mapping = REGION_MAPPING_REVERSE
            mapped_region = region_mapping.get(region_name, region_name)
            if mapped_region != region_name:
                growth_rate = self._get_quarterly_growth_rate(sheet_name, mapped_region, quarter_key)
                if growth_rate is not None:
                    return self.format_percentage(growth_rate, decimal_places=1, include_percent=False)
            return "N/A"
        
        # 분기 헤더 마커 처리 (예: 분기1_헤더, 2023_3분기)
        header_match = re.match(r'분기(\d+)_헤더', key)
        if header_match:
            quarter_idx = int(header_match.group(1)) - 1
            # 동적으로 분기 헤더 생성 (현재 연도/분기 기준으로 역순)
            headers = self.period_detector.get_quarter_headers(
                sheet_name, start_year=year, start_quarter=quarter, count=8
            )
            if 0 <= quarter_idx < len(headers):
                return headers[quarter_idx]
        
        # 연도_분기 헤더 마커 처리 (예: 2023_3분기)
        # 템플릿에서 헤더로 사용되는 경우 (예: {수출:2023_3분기})
        year_quarter_header_match = re.match(r'^(\d{4})_(\d)분기$', key)
        if year_quarter_header_match:
            target_year = int(year_quarter_header_match.group(1))
            target_quarter = int(year_quarter_header_match.group(2))
            # 헤더 형식: "2023.3" 또는 "2023 3/4" - 템플릿에 맞게 조정
            # 정답지 확인 필요하지만 일단 "2023.3" 형식 사용
            return f'{target_year}.{target_quarter}'
        
        # 지역별 연령대별 증감률/증감pp (예: 전국_60대이상_증감pp, 서울_30대59세_증감pp, 경북_15대29세_증감pp, 전국_30대_증감pp)
        # region_growth_match보다 먼저 처리해야 함
        # 증감pp: 퍼센트포인트 차이 (현재값 - 이전값)
        region_age_match = re.match(r'^([가-힣]+)_(\d+대이상|\d+대\d+세|\d+대)_증감(?:률|pp)$', key)
        if region_age_match:
            region_name = region_age_match.group(1)
            age_group = region_age_match.group(2)  # '60대이상', '30대59세', '15대29세'
            
            # 지역명 매핑
            region_mapping = REGION_MAPPING_REVERSE
            actual_region_name = region_mapping.get(region_name, region_name)
            
            # 연령대 매핑
            age_mapping = {
                '60대이상': ['60세이상', '60대이상', '60세 이상', '60세이상', '60 - 이상'],
                '30대59세': ['30~59세', '30-59세', '30대59세', '30세~59세', '30 - 59세', '30 ~ 59세'],
                '15대29세': ['15~29세', '15-29세', '15대29세', '15세~29세', '15 - 29세', '15 ~ 29세'],
                # 고용률용 연령대
                '30대': ['30 - 39세', '30~39세', '30-39세', '30대', '30대39세'],
                '40대': ['40 - 49세', '40~49세', '40-49세', '40대', '40대49세'],
                '50대': ['50 - 59세', '50~59세', '50-59세', '50대', '50대59세'],
            }
            
            search_ages = age_mapping.get(age_group, [age_group])
            
            # 현재 연도/분기의 증감률 계산
            current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
            sheet = self.excel_extractor.get_sheet(sheet_name)
            
            # 실업률/고용률 시트 구조 확인
            is_unemployment_sheet = ('실업' in sheet_name or sheet_name == '실업자 수')
            is_employment_rate_sheet = ('고용' in sheet_name or sheet_name == '고용' or '고용률' in sheet_name or sheet_name == '고용률')
            
            if is_unemployment_sheet:
                # 실업자 수 시트: 1열에 시도, 2열에 연령계층
                current_region = None
                for row in range(4, min(5000, sheet.max_row + 1)):
                    cell_a = sheet.cell(row=row, column=1)  # 시도
                    cell_b = sheet.cell(row=row, column=2)  # 연령계층
                    
                    if cell_a.value:
                        current_region = str(cell_a.value).strip()
                    
                    if cell_b.value and current_region:
                        age_str = str(cell_b.value).strip()
                        
                        if (current_region == actual_region_name or 
                            actual_region_name in current_region or 
                            current_region in actual_region_name):
                            for search_age in search_ages:
                                if search_age in age_str or age_str in search_age:
                                    current = sheet.cell(row=row, column=current_col).value
                                    prev = sheet.cell(row=row, column=prev_col).value
                                    
                                    if current is not None and prev is not None and prev != 0:
                                        growth_rate = ((current / prev) - 1) * 100
                                        return self.format_percentage(growth_rate, decimal_places=1)
            elif is_employment_rate_sheet:
                # 고용률 시트: 1열에 지역 코드, 2열에 지역 이름, 3열에 분류 단계, 4열에 연령대
                config = self._get_sheet_config(sheet_name)
                category_col = config['category_column']
                
                for row in range(4, min(5000, sheet.max_row + 1)):
                    cell_b = sheet.cell(row=row, column=2)  # 지역 이름
                    cell_c = sheet.cell(row=row, column=3)  # 분류 단계
                    cell_category = sheet.cell(row=row, column=category_col)  # 연령대
                    
                    if cell_b.value and cell_c.value and cell_category.value:
                        region_str = str(cell_b.value).strip()
                        class_str = str(cell_c.value).strip()
                        age_str = str(cell_category.value).strip()
                        
                        # 지역명 매칭 및 분류 단계가 1인 경우 (연령대별)
                        if (region_str == region_name and class_str == '1'):
                            for search_age in search_ages:
                                if search_age in age_str or age_str in search_age:
                                    current = sheet.cell(row=row, column=current_col).value
                                    prev = sheet.cell(row=row, column=prev_col).value
                                    
                                    if current is not None and prev is not None and prev != 0:
                                        growth_rate = ((current / prev) - 1) * 100
                                        return self.format_percentage(growth_rate, decimal_places=1)
            
            return "N/A"
        
        # 전국 증감률/증감pp 처리 (동적 연도/분기)
        # 증감pp: 퍼센트포인트 차이 (현재값 - 이전값)
        # 증감률: 증감률 (%) 계산
        if key == '전국_증감률' or key == '전국_증감pp':
            # 실업률/고용률 시트인지 확인
            is_unemployment_sheet = ('실업' in sheet_name or sheet_name == '실업자 수')
            is_employment_rate_sheet = ('고용' in sheet_name or sheet_name == '고용' or '고용률' in sheet_name or sheet_name == '고용률')
            
            if is_unemployment_sheet:
                # 실업률 증감 계산: 현재 실업률 - 전년 동분기 실업률 (%p 단위)
                # 실업률 표에서 직접 값 가져오기 (Row 81부터)
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                # 실업률 가져오기 헬퍼 함수
                def get_unemployment_rate(calc_year, calc_quarter, region='전국'):
                    """특정 연도/분기의 실업률을 실업률 표에서 가져옵니다."""
                    target_col, _ = self._get_quarter_columns(calc_year, calc_quarter, sheet_name)
                    
                    # 동적으로 찾지 못한 경우, Row 3에서 직접 헤더 찾기
                    if target_col is None:
                        header_pattern = f"{calc_year}  {calc_quarter}/4"
                        for col in range(1, min(100, sheet.max_column + 1)):
                            header_val = sheet.cell(row=3, column=col).value
                            if header_val and header_pattern in str(header_val):
                                target_col = col
                                break
                    
                    if target_col is None:
                        return None
                    
                    current_region = None
                    
                    for row in range(81, min(5000, sheet.max_row + 1)):
                        cell_a = sheet.cell(row=row, column=1)  # 시도
                        cell_b = sheet.cell(row=row, column=2)  # 연령계층
                        
                        if cell_a.value:
                            current_region = str(cell_a.value).strip()
                        
                        if cell_b.value and current_region:
                            region_str = current_region
                            age_str = str(cell_b.value).strip()
                            
                            if region_str == region and age_str == '계':
                                value = sheet.cell(row=row, column=target_col).value
                                if value is not None:
                                    try:
                                        return float(value)
                                    except (ValueError, TypeError):
                                        pass
                                break
                    return None
                
                # 현재 실업률과 전년 동분기 실업률 가져오기
                current_rate = get_unemployment_rate(year, quarter, '전국')
                prev_year = year - 1
                prev_rate = get_unemployment_rate(prev_year, quarter, '전국')
                
                if current_rate is not None and prev_rate is not None:
                    # 실업률 증감은 %p 단위 (퍼센트 포인트)
                    growth_rate_pp = current_rate - prev_rate
                    return self.format_percentage(growth_rate_pp, decimal_places=1, include_percent=False)
                
                return "N/A"
            elif is_employment_rate_sheet:
                # 고용 시트 구조: 1열에 지역 코드, 2열에 지역 이름, 3열에 분류 단계, 4열에 연령대/항목
                # 고용률 증감률 계산: 현재 고용률 - 전년 동분기 고용률 (%p 단위)
                current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                config = self._get_sheet_config(sheet_name)
                category_col = config['category_column']
                
                # 전국 고용률 계산 함수
                def calculate_employment_rate(region_name, col):
                    # 먼저 "계" 행에서 직접 값을 읽어봄
                    for row in range(4, min(5000, sheet.max_row + 1)):
                        cell_b = sheet.cell(row=row, column=2)
                        cell_c = sheet.cell(row=row, column=3)
                        cell_category = sheet.cell(row=row, column=category_col)
                        
                        if cell_b.value and cell_c.value and cell_category.value:
                            region_str = str(cell_b.value).strip()
                            class_str = str(cell_c.value).strip()
                            category_str = str(cell_category.value).strip()
                            
                            if (region_str == region_name and class_str == '0' and category_str == '계'):
                                value = sheet.cell(row=row, column=col).value
                                if value is not None:
                                    return float(value) if isinstance(value, (int, float)) else None
                    
                    # "계" 행에서 값을 찾지 못한 경우, 취업자수와 인구수로 계산
                    employment_count = None
                    population_count = None
                    
                    for row in range(4, min(5000, sheet.max_row + 1)):
                        cell_b = sheet.cell(row=row, column=2)
                        cell_c = sheet.cell(row=row, column=3)
                        cell_category = sheet.cell(row=row, column=category_col)
                        
                        if cell_b.value and cell_c.value and cell_category.value:
                            region_str = str(cell_b.value).strip()
                            class_str = str(cell_c.value).strip()
                            category_str = str(cell_category.value).strip()
                            
                            if (region_str == region_name and class_str == '0'):
                                value = sheet.cell(row=row, column=col).value
                                
                                if ('취업자' in category_str or '취업' in category_str) and '인구' not in category_str:
                                    if value is not None:
                                        employment_count = float(value) if isinstance(value, (int, float)) else None
                                
                                if ('15세' in category_str or '15세이상' in category_str or '15세 이상' in category_str) and '인구' in category_str:
                                    if value is not None:
                                        population_count = float(value) if isinstance(value, (int, float)) else None
                    
                    if employment_count is not None and population_count is not None and population_count != 0:
                        return (employment_count / population_count) * 100
                    return None
                
                # 전국 고용률 계산
                current_rate = calculate_employment_rate('전국', current_col)
                prev_rate = calculate_employment_rate('전국', prev_col)
                
                if current_rate is not None and prev_rate is not None:
                    # 고용률 증감은 %p 단위 (퍼센트 포인트)
                    growth_rate_pp = current_rate - prev_rate
                    return self.format_percentage(growth_rate_pp, decimal_places=1, include_percent=False)
                
                return "N/A"
            else:
                # 일반 시트 처리
                quarter_key = f'{year}_{quarter}분기'
                growth_rate = self._get_quarterly_growth_rate(sheet_name, '전국', quarter_key)
                if growth_rate is not None:
                    return self.format_percentage(growth_rate, decimal_places=1)
                return "N/A"
        
        # 지역별 증감률/증감pp 마커 처리 (예: 서울_증감률, 울산_증감pp)
        # 증감pp: 퍼센트포인트 차이 (현재값 - 이전값)
        region_growth_match = re.match(r'^([가-힣]+)_증감(?:률|pp)$', key)
        if region_growth_match:
            region_name = region_growth_match.group(1)
            
            # 실업률/고용률 시트인지 확인
            is_unemployment_sheet = ('실업' in sheet_name or sheet_name == '실업자 수')
            is_employment_rate_sheet = ('고용' in sheet_name or sheet_name == '고용' or '고용률' in sheet_name or sheet_name == '고용률')
            
            if is_unemployment_sheet:
                # 실업률 증감 계산: 현재 실업률 - 전년 동분기 실업률 (%p 단위)
                # 실업률 표에서 직접 값 가져오기 (Row 81부터)
                region_mapping = {
                    '서울': '서울특별시', '부산': '부산광역시', '대구': '대구광역시',
                    '인천': '인천광역시', '광주': '광주광역시', '대전': '대전광역시',
                    '울산': '울산광역시', '세종': '세종특별자치시', '경기': '경기도',
                    '강원': '강원도', '충북': '충청북도', '충남': '충청남도',
                    '전북': '전라북도', '전남': '전라남도', '경북': '경상북도',
                    '경남': '경상남도', '제주': '제주특별자치도',
                }
                actual_region_name = region_mapping.get(region_name, region_name)
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                # 실업률 가져오기 헬퍼 함수
                def get_unemployment_rate(calc_year, calc_quarter, region_name_val):
                    """특정 연도/분기의 실업률을 실업률 표에서 가져옵니다."""
                    target_col, _ = self._get_quarter_columns(calc_year, calc_quarter, sheet_name)
                    
                    # 동적으로 찾지 못한 경우, Row 3에서 직접 헤더 찾기
                    if target_col is None:
                        header_pattern = f"{calc_year}  {calc_quarter}/4"
                        for col in range(1, min(100, sheet.max_column + 1)):
                            header_val = sheet.cell(row=3, column=col).value
                            if header_val and header_pattern in str(header_val):
                                target_col = col
                                break
                    
                    if target_col is None:
                        return None
                    
                    current_region = None
                    
                    for row in range(81, min(5000, sheet.max_row + 1)):
                        cell_a = sheet.cell(row=row, column=1)  # 시도
                        cell_b = sheet.cell(row=row, column=2)  # 연령계층
                        
                        if cell_a.value:
                            current_region = str(cell_a.value).strip()
                        
                        if cell_b.value and current_region:
                            region_str = current_region
                            age_str = str(cell_b.value).strip()
                            
                            if (region_str == region_name_val or 
                                region_name_val in region_str or 
                                region_str in region_name_val):
                                if age_str == '계':
                                    value = sheet.cell(row=row, column=target_col).value
                                    if value is not None:
                                        try:
                                            return float(value)
                                        except (ValueError, TypeError):
                                            pass
                                    break
                    return None
                
                # 현재 실업률과 전년 동분기 실업률 가져오기
                current_rate = get_unemployment_rate(year, quarter, actual_region_name)
                prev_year = year - 1
                prev_rate = get_unemployment_rate(prev_year, quarter, actual_region_name)
                
                if current_rate is not None and prev_rate is not None:
                    # 실업률 증감은 %p 단위 (퍼센트 포인트)
                    growth_rate_pp = current_rate - prev_rate
                    return self.format_percentage(growth_rate_pp, decimal_places=1, include_percent=False)
                
                return "N/A"
            elif is_employment_rate_sheet:
                # 고용 시트 구조: 1열에 지역 코드, 2열에 지역 이름, 3열에 분류 단계, 4열에 연령대/항목
                # 고용률 증감률 계산: 현재 고용률 - 전년 동분기 고용률 (%p 단위)
                current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                config = self._get_sheet_config(sheet_name)
                category_col = config['category_column']
                
                # 지역별 고용률 계산 함수
                def calculate_employment_rate(region_name, col):
                    # 먼저 "계" 행에서 직접 값을 읽어봄
                    for row in range(4, min(5000, sheet.max_row + 1)):
                        cell_b = sheet.cell(row=row, column=2)
                        cell_c = sheet.cell(row=row, column=3)
                        cell_category = sheet.cell(row=row, column=category_col)
                        
                        if cell_b.value and cell_c.value and cell_category.value:
                            region_str = str(cell_b.value).strip()
                            class_str = str(cell_c.value).strip()
                            category_str = str(cell_category.value).strip()
                            
                            if (region_str == region_name and class_str == '0' and category_str == '계'):
                                value = sheet.cell(row=row, column=col).value
                                if value is not None:
                                    return float(value) if isinstance(value, (int, float)) else None
                    
                    # "계" 행에서 값을 찾지 못한 경우, 취업자수와 인구수로 계산
                    employment_count = None
                    population_count = None
                    
                    for row in range(4, min(5000, sheet.max_row + 1)):
                        cell_b = sheet.cell(row=row, column=2)
                        cell_c = sheet.cell(row=row, column=3)
                        cell_category = sheet.cell(row=row, column=category_col)
                        
                        if cell_b.value and cell_c.value and cell_category.value:
                            region_str = str(cell_b.value).strip()
                            class_str = str(cell_c.value).strip()
                            category_str = str(cell_category.value).strip()
                            
                            if (region_str == region_name and class_str == '0'):
                                value = sheet.cell(row=row, column=col).value
                                
                                if ('취업자' in category_str or '취업' in category_str) and '인구' not in category_str:
                                    if value is not None:
                                        employment_count = float(value) if isinstance(value, (int, float)) else None
                                
                                if ('15세' in category_str or '15세이상' in category_str or '15세 이상' in category_str) and '인구' in category_str:
                                    if value is not None:
                                        population_count = float(value) if isinstance(value, (int, float)) else None
                    
                    if employment_count is not None and population_count is not None and population_count != 0:
                        return (employment_count / population_count) * 100
                    return None
                
                # 지역별 고용률 계산
                current_rate = calculate_employment_rate(region_name, current_col)
                prev_rate = calculate_employment_rate(region_name, prev_col)
                
                if current_rate is not None and prev_rate is not None:
                    # 고용률 증감은 %p 단위 (퍼센트 포인트)
                    growth_rate_pp = current_rate - prev_rate
                    return self.format_percentage(growth_rate_pp, decimal_places=1, include_percent=False)
                
                return "N/A"
            else:
                # 일반 시트 처리
                quarter_key = f'{year}_{quarter}분기'
                growth_rate = self._get_quarterly_growth_rate(sheet_name, region_name, quarter_key)
                if growth_rate is not None:
                    return self.format_percentage(growth_rate, decimal_places=1)
                return "N/A"
        
        # 지역별 증감방향 마커 처리 (예: 서울_증감방향, 울산_증감방향)
        region_direction_match = re.match(r'^([가-힣]+)_증감방향$', key)
        if region_direction_match:
            region_name = region_direction_match.group(1)
            
            # 실업률/고용률 시트인지 확인
            is_unemployment_sheet = ('실업' in sheet_name or sheet_name == '실업자 수')
            is_employment_rate_sheet = ('고용' in sheet_name or '고용률' in sheet_name)
            
            if is_unemployment_sheet or is_employment_rate_sheet:
                # 실업률/고용률은 %p 단위이므로 증감pp를 가져와서 방향 결정
                growth_rate_str = self._process_dynamic_marker(sheet_name, f'{region_name}_증감pp', year, quarter)
                try:
                    growth_rate = float(growth_rate_str)
                    return self.get_growth_direction(growth_rate)
                except (ValueError, TypeError):
                    return "N/A"
            else:
                # 일반 시트 처리
                quarter_key = f'{year}_{quarter}분기'
                growth_rate = self._get_quarterly_growth_rate(sheet_name, region_name, quarter_key)
                if growth_rate is not None:
                    return self.get_growth_direction(growth_rate)
                return "N/A"
        
        # 전국 품목별 증감률 마커 처리 (예: 전국_메모리반도체_증감률, 전국_선박_증감률, 전국_중화학공업품_증감률)
        national_item_match = re.match(r'^전국_(.+)_증감률$', key)
        if national_item_match:
            item_name = national_item_match.group(1)
            # 품목 이름 매핑 (템플릿에서 사용하는 이름 -> 엑셀에서 찾을 이름)
            # 스키마에서 가져오기
            item_name_mapping = self.schema_loader.get_name_mapping('item_name_mapping')
            if item_name_mapping:
                search_names = item_name_mapping.get(item_name, [item_name])
            else:
                search_names = [item_name]
            if isinstance(search_names, str):
                search_names = [search_names]
            
            # 건설수주 시트의 경우 토목/건축은 특별 처리
            if sheet_name == '건설 (공표자료)':
                if item_name == '토목' or item_name == '건축':
                    # 전국의 토목/건축 카테고리 찾기
                    categories = self._get_categories_for_region(sheet_name, '전국', year, quarter, top_n=20)
                    for category in categories:
                        category_name = str(category['name']).strip()
                        if item_name in category_name or category_name == item_name:
                            return self.format_percentage(category['growth_rate'], decimal_places=1)
                    return "N/A"
            
            # 전국의 해당 품목 증감률 찾기
            categories = self._get_categories_for_region(sheet_name, '전국', year, quarter, top_n=50)
            for category in categories:
                category_name = str(category['name']).strip()
                # 매핑된 이름 중 하나라도 일치하는지 확인
                for search_name in search_names:
                    # 정확히 일치하거나 포함되는지 확인 (양방향)
                    # 공백과 특수문자 제거 후 비교
                    category_clean = category_name.replace(' ', '').replace('·', '').replace(',', '').replace('및', '')
                    search_clean = search_name.replace(' ', '').replace('·', '').replace(',', '').replace('및', '')
                    if (search_name in category_name or category_name in search_name or 
                        search_clean in category_clean or category_clean in search_clean):
                        return self.format_percentage(category['growth_rate'], decimal_places=1)
            return "N/A"
        
        # 지역별 품목별 증감률/증감pp 마커 처리 (예: 부산_외식제외개인서비스_증감pp, 제주_농산물_증감pp)
        # 증감pp: 퍼센트포인트 차이 (현재값 - 이전값)
        region_item_match = re.match(r'^([가-힣]+)_(.+)_증감(?:률|pp)$', key)
        if region_item_match:
            region_name = region_item_match.group(1)
            item_name = region_item_match.group(2)
            
            # 전국 품목별과 동일한 매핑 사용 (스키마에서 가져오기)
            item_name_mapping = self.schema_loader.get_name_mapping('item_name_mapping')
            if item_name_mapping:
                search_names = item_name_mapping.get(item_name, [item_name])
            else:
                search_names = [item_name]
            if isinstance(search_names, str):
                search_names = [search_names]
            
            # 건설수주 시트의 경우 토목/건축은 특별 처리
            if sheet_name == '건설 (공표자료)':
                if item_name == '토목' or item_name == '건축':
                    # 해당 지역의 토목/건축 카테고리 찾기
                    categories = self._get_categories_for_region(sheet_name, region_name, year, quarter, top_n=20)
                    for category in categories:
                        category_name = str(category['name']).strip()
                        if item_name in category_name or category_name == item_name:
                            return self.format_percentage(category['growth_rate'], decimal_places=1)
                    return "N/A"
            
            # 해당 지역의 품목 증감률 찾기
            categories = self._get_categories_for_region(sheet_name, region_name, year, quarter, top_n=50)
            for category in categories:
                category_name = str(category['name']).strip()
                # 매핑된 이름 중 하나라도 일치하는지 확인
                for search_name in search_names:
                    # 정확히 일치하거나 포함되는지 확인 (양방향)
                    # 공백과 특수문자 제거 후 비교
                    category_clean = category_name.replace(' ', '').replace('·', '').replace(',', '').replace('및', '')
                    search_clean = search_name.replace(' ', '').replace('·', '').replace(',', '').replace('및', '')
                    if (search_name in category_name or category_name in search_name or 
                        search_clean in category_clean or category_clean in search_clean):
                        return self.format_percentage(category['growth_rate'], decimal_places=1)
            return "N/A"
        
        # 지역별 순이동 마커 처리 (예: 서울_순이동)
        region_net_movement_match = re.match(r'^([가-힣]+)_순이동$', key)
        if region_net_movement_match:
            region_name = region_net_movement_match.group(1)
            # 시도 간 이동 시트에서 순이동 계산 (유입 - 유출)
            # 시도 간 이동 시트의 해당 지역 총지수 값이 순이동 (양수=유입, 음수=유출)
            current_col, _ = self._get_quarter_columns(year, quarter, '시도 간 이동')
            if current_col:
                sheet = self.excel_extractor.get_sheet('시도 간 이동')
                # 지역의 총지수 행 찾기
                for row in range(4, min(5000, sheet.max_row + 1)):
                    cell_region = sheet.cell(row=row, column=2).value
                    cell_category = sheet.cell(row=row, column=4).value
                    if (cell_region and str(cell_region).strip() == region_name and 
                        cell_category and str(cell_category).strip() == '계'):
                        value = sheet.cell(row=row, column=current_col).value
                        if value is not None:
                            try:
                                num_value = float(value)
                                return str(int(num_value))
                            except (ValueError, TypeError):
                                pass
            return "N/A"
        
        # 지역별 연령대별 유입/유출 마커 처리 (예: 서울_2024세_유입, 서울_2529세_유출)
        region_age_flow_match = re.match(r'^([가-힣]+)_(\d+세|\d+~\d+세|\d+대)_(유입|유출)$', key)
        if region_age_flow_match:
            region_name = region_age_flow_match.group(1)
            age_group_str = region_age_flow_match.group(2)  # '2024세', '2529세', '30대' 등
            direction = region_age_flow_match.group(3)  # '유입' 또는 '유출'
            
            # 연령 그룹 이름 매핑 (템플릿 형식 -> 스키마 형식)
            age_group_mapping = {
                '2024세': '20~24세',
                '2529세': '25~29세',
                '3034세': '30~34세',
                '5559세': '55~59세',
                '1529세': '15~29세',
                '3059세': '30~59세',
                '60세이상': '60세이상',
                '30대': '30~39세',
                '50대': '50~59세'
            }
            age_group_name = age_group_mapping.get(age_group_str, age_group_str)
            
            # 연령별 인구이동 시트에서 값 가져오기
            value = self._get_age_group_value('연령별 인구이동', region_name, age_group_name, direction, year, quarter)
            return value if value else "N/A"
        
        # 연도 헤더 마커 처리 (예: 2023, 2024, 2025)
        year_header_match = re.match(r'^(\d{4})$', key)
        if year_header_match:
            target_year = int(year_header_match.group(1))
            return str(target_year)
        
        # 연도/분기 값 처리
        if key == '연도':
            return str(year)
        elif key == '분기':
            return str(quarter)
        
        # 상승_시도수 처리 (고용률용)
        if key == '상승_시도수' or key == '상승시도수':
            sheet = self.excel_extractor.get_sheet(sheet_name)
            positive_count = 0
            seen_regions = set()
            
            # 연도/분기에 해당하는 열 번호 가져오기
            current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
            
            # 시트별 설정 가져오기
            config = self._get_sheet_config(sheet_name)
            category_col = config['category_column']
            
            for row in range(4, min(5000, sheet.max_row + 1)):
                cell_a = sheet.cell(row=row, column=1)  # 지역 코드
                cell_b = sheet.cell(row=row, column=2)  # 지역 이름
                cell_c = sheet.cell(row=row, column=3)  # 분류 단계
                cell_category = sheet.cell(row=row, column=category_col)  # 연령대/산업 이름
                
                # 총지수 또는 계 (분류 단계가 0인 경우)
                is_total = False
                if cell_category.value:
                    category_str = str(cell_category.value).strip()
                    if category_str == '총지수' or category_str == '계' or category_str == '   계':
                        is_total = True
                
                if cell_b.value and is_total:
                    # 시도 코드 확인: 2자리 숫자이고 00이 아닌 것
                    code_str = str(cell_a.value).strip() if cell_a.value else ''
                    is_sido = (len(code_str) == 2 and code_str.isdigit() and code_str != '00')
                    
                    region_name = str(cell_b.value).strip()
                    if is_sido and region_name not in seen_regions:
                        seen_regions.add(region_name)
                        current = sheet.cell(row=row, column=current_col).value
                        prev = sheet.cell(row=row, column=prev_col).value
                        
                        # 가중치 열 찾기 - 가중치 열이 없으면 원데이터로 비교
                        structure = self.dynamic_parser.parse_sheet_structure(sheet_name)
                        weight_col = structure.get('weight_column')
                        has_weight_column = weight_col is not None
                        
                        # 결측치 체크 - 결측치는 1로 채움
                        if current is None or (isinstance(current, str) and (not current.strip() or current.strip() == '-')):
                            current = prev if prev is not None else 1.0
                        if prev is None or (isinstance(prev, str) and (not prev.strip() or prev.strip() == '-')):
                            prev = current if current is not None else 1.0
                        
                        if current is not None and prev is not None and prev != 0:
                            # 값이 1이고 둘 다 1이면 스킵 (기본값)
                            try:
                                current_num = float(current)
                                prev_num = float(prev)
                                if current_num == 1.0 and prev_num == 1.0:
                                    continue
                                growth_rate = ((current_num / prev_num) - 1) * 100
                                if growth_rate > 0:
                                    positive_count += 1
                            except (ValueError, TypeError):
                                continue
            
            return str(positive_count)
        
        # 증가시도수 처리 (증가_시도수 형식도 지원)
        if key == '증가시도수' or key == '증가_시도수':
            # 시도만 카운트 (그룹 제외)
            sheet = self.excel_extractor.get_sheet(sheet_name)
            positive_count = 0
            
            # 연도/분기에 해당하는 열 번호 가져오기
            current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
            
            # 시트별 설정 가져오기
            config = self._get_sheet_config(sheet_name)
            category_col = config['category_column']
            
            for row in range(4, min(5000, sheet.max_row + 1)):
                cell_a = sheet.cell(row=row, column=1)  # 지역 코드
                cell_b = sheet.cell(row=row, column=2)  # 지역 이름
                cell_c = sheet.cell(row=row, column=3)  # 분류 단계
                cell_category = sheet.cell(row=row, column=category_col)  # 업태/산업 이름
                
                # 총지수, 계, 합계 인식 (분류 단계가 0인 경우)
                is_total = False
                if cell_category.value:
                    category_str = str(cell_category.value).strip()
                    if category_str in ['총지수', '계', '   계', '합계']:
                        is_total = True
                
                if cell_b.value and is_total:
                    # 시도 코드 확인: 2자리 숫자이고 00이 아닌 것
                    code_str = str(cell_a.value).strip() if cell_a.value else ''
                    is_sido = (len(code_str) == 2 and code_str.isdigit() and code_str != '00')
                    
                    if is_sido:
                        current = sheet.cell(row=row, column=current_col).value
                        prev = sheet.cell(row=row, column=prev_col).value
                        
                        if current is not None and prev is not None and prev != 0:
                            growth_rate = ((current / prev) - 1) * 100
                            if growth_rate > 0:
                                positive_count += 1
            
            return str(positive_count)
        elif key == '감소시도수' or key == '감소_시도수':
            # 감소한 시도 개수
            sheet = self.excel_extractor.get_sheet(sheet_name)
            negative_count = 0
            seen_regions = set()
            
            # 연도/분기에 해당하는 열 번호 가져오기
            current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
            
            # 시트별 설정 가져오기
            config = self._get_sheet_config(sheet_name)
            category_col = config['category_column']
            
            for row in range(4, min(5000, sheet.max_row + 1)):
                cell_a = sheet.cell(row=row, column=1)  # 지역 코드
                cell_b = sheet.cell(row=row, column=2)  # 지역 이름
                cell_c = sheet.cell(row=row, column=3)  # 분류 단계
                cell_category = sheet.cell(row=row, column=category_col)  # 업태/산업 이름
                
                # 총지수, 계, 합계 인식 (분류 단계가 0인 경우)
                is_total = False
                if cell_category.value:
                    category_str = str(cell_category.value).strip()
                    if category_str in ['총지수', '계', '   계', '합계']:
                        is_total = True
                
                if cell_b.value and is_total:
                    # 시도 코드 확인: 2자리 숫자이고 00이 아닌 것
                    code_str = str(cell_a.value).strip() if cell_a.value else ''
                    is_sido = (len(code_str) == 2 and code_str.isdigit() and code_str != '00')
                    
                    region_name = str(cell_b.value).strip()
                    if is_sido and region_name not in seen_regions:
                        seen_regions.add(region_name)
                        current = sheet.cell(row=row, column=current_col).value
                        prev = sheet.cell(row=row, column=prev_col).value
                        
                        if current is not None and prev is not None and prev != 0:
                            growth_rate = ((current / prev) - 1) * 100
                            if growth_rate < 0:
                                negative_count += 1
            
            return str(negative_count)
        elif key == '기준연도':
            return str(year)
        elif key == '기준분기':
            return f'{quarter}/4'
        
        # 하락_시도수 처리 (감소시도수와 동일)
        if key == '하락_시도수' or key == '하락시도수':
            # 실업률 시트 여부 확인
            is_unemployment_sheet = (sheet_name == '실업률' or sheet_name == '실업자 수')
            is_employment_rate_sheet = ('고용' in sheet_name or sheet_name == '고용' or '고용률' in sheet_name or sheet_name == '고용률')
            
            if is_unemployment_sheet:
                # 실업률 시트: 전용 함수로 백분율포인트 단위 증감률 계산
                regions_data = self._get_unemployment_rate_regions_data(year, quarter)
                # 하락 = 실업률이 내려간 지역 (증감률 < 0)
                negative_count = sum(1 for r in regions_data if r.get('growth_rate', 0) < 0)
                return str(negative_count)
            elif is_employment_rate_sheet:
                # 고용률 시트: 별도 로직 (필요시 추가)
                pass
            
            # 일반 시트 구조 (실업률/고용률 아닌 경우)
            if not is_unemployment_sheet:
                sheet = self.excel_extractor.get_sheet(sheet_name)
                negative_count = 0
                seen_regions = set()
                current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                
                config = self._get_sheet_config(sheet_name)
                category_col = config.get('category_column', 6)
                
                for row in range(4, min(5000, sheet.max_row + 1)):
                    cell_a = sheet.cell(row=row, column=1)  # 지역 코드
                    cell_b = sheet.cell(row=row, column=2)  # 지역 이름
                    cell_c = sheet.cell(row=row, column=3)  # 분류 단계
                    cell_category = sheet.cell(row=row, column=category_col)  # 업태/산업 이름
                    
                    # 총지수 또는 계 (분류 단계가 0인 경우)
                    is_total = False
                    if cell_category.value:
                        category_str = str(cell_category.value).strip()
                        if category_str == '총지수' or category_str == '계' or category_str == '   계':
                            is_total = True
                    
                    if cell_b.value and is_total:
                        # 시도 코드 확인: 2자리 숫자이고 00이 아닌 것
                        code_str = str(cell_a.value).strip() if cell_a.value else ''
                        is_sido = (len(code_str) == 2 and code_str.isdigit() and code_str != '00')
                        
                        region_name = str(cell_b.value).strip()
                        if is_sido and region_name not in seen_regions:
                            seen_regions.add(region_name)
                            current = sheet.cell(row=row, column=current_col).value
                            prev = sheet.cell(row=row, column=prev_col).value
                            
                            if current is not None and prev is not None and prev != 0:
                                growth_rate = ((current / prev) - 1) * 100
                                if growth_rate < 0:
                                    negative_count += 1
            
            return str(negative_count)
        
        # 실업률/고용률 관련 마커 처리
        # 지역별 실업률/고용률 값 (예: 전국_실업률, 서울_실업률, 전국_고용률)
        region_value_match = re.match(r'^([가-힣]+)_(실업률|고용률)$', key)
        if region_value_match:
            region_name = region_value_match.group(1)
            value_type = region_value_match.group(2)  # '실업률' 또는 '고용률'
            
            # 실업률 시트인지 고용률 시트인지 확인
            is_unemployment_sheet = ('실업' in sheet_name or sheet_name == '실업자 수')
            is_employment_rate_sheet = ('고용' in sheet_name or sheet_name == '고용' or '고용률' in sheet_name or sheet_name == '고용률')
            
            if is_unemployment_sheet:
                # 실업률 표 설정 가져오기
                sheet_config = self._get_sheet_config(sheet_name)
                unemployment_table_config = sheet_config.get('unemployment_rate_table', {})
                
                if not unemployment_table_config.get('enabled', True):
                    return "N/A"
                
                # 스키마에서 설정 가져오기
                start_row = unemployment_table_config.get('start_row', 81)
                title_row = unemployment_table_config.get('title_row', 79)
                region_col = unemployment_table_config.get('region_column', 1)
                age_group_col = unemployment_table_config.get('age_group_column', 2)
                age_group_filter = unemployment_table_config.get('age_group_filter', '계')
                header_row_number = unemployment_table_config.get('header_row_number', 3)
                region_mapping = unemployment_table_config.get('region_mapping', {})
                
                # 지역명 매핑
                actual_region_name = region_mapping.get(region_name, region_name)
                
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                # 헤더 행 찾기: "시도별 실업률(%)" 제목을 찾고 그 아래 헤더 행 찾기
                title_text = unemployment_table_config.get('title_text', '시도별 실업률(%)')
                header_row = None
                
                # 제목 행 찾기
                title_found = False
                for row in range(1, min(200, sheet.max_row + 1)):
                    for col in range(1, min(10, sheet.max_column + 1)):
                        cell_val = sheet.cell(row=row, column=col).value
                        if cell_val and title_text in str(cell_val):
                            title_found = True
                            # 제목 행 아래에서 헤더 행 찾기 (제목 행 + header_row_number 또는 제목 행 + 1, 2, 3 시도)
                            for offset in [header_row_number, 1, 2, 3]:
                                test_row = row + offset
                                if test_row <= sheet.max_row:
                                    # 헤더 행인지 확인 (연도/분기 패턴이 있는지 확인)
                                    for test_col in range(1, min(20, sheet.max_column + 1)):
                                        test_val = sheet.cell(row=test_row, column=test_col).value
                                        if test_val and (str(year) in str(test_val) or '분기' in str(test_val) or '/4' in str(test_val)):
                                            header_row = test_row
                                            break
                                    if header_row:
                                        break
                            break
                    if title_found and header_row:
                        break
                
                # 헤더 행을 찾지 못한 경우 설정값 사용
                if header_row is None:
                    header_row = title_row + header_row_number if title_row else header_row_number
                
                # 현재 연도/분기의 열 번호 가져오기
                current_col, _ = self._get_quarter_columns(year, quarter, sheet_name)
                
                # 동적으로 찾지 못한 경우, 헤더 행에서 직접 헤더 찾기
                if current_col is None:
                    # 여러 패턴 시도: "2025  2/4", "2025 2/4", "2025년 2/4", "2025 2분기" 등
                    patterns = [
                        f"{year}  {quarter}/4",  # 공백 2개
                        f"{year} {quarter}/4",    # 공백 1개
                        f"{year}년 {quarter}/4",
                        f"{year} {quarter}분기",
                        f"{year}/{quarter}",
                    ]
                    for col in range(1, min(100, sheet.max_column + 1)):
                        header_val = sheet.cell(row=header_row, column=col).value
                        if header_val:
                            header_str = str(header_val).strip()
                            for pattern in patterns:
                                if pattern in header_str:
                                    current_col = col
                                    break
                            if current_col:
                                break
                
                if current_col is None:
                    return "N/A"
                
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                # 실업률 표에서 직접 값 가져오기
                # 시도별로 여러 행이 있고, 각 시도마다 연령계층별로 4개 행이 있음 (15~29세, 30~59세, 60세이상, 계)
                # '계'는 하단에 있음
                current_region = None
                for row in range(start_row, min(5000, sheet.max_row + 1)):
                    cell_region = sheet.cell(row=row, column=region_col)  # 시도
                    cell_age = sheet.cell(row=row, column=age_group_col)  # 연령계층
                    
                    # 시도명이 있으면 현재 시도 업데이트
                    if cell_region.value:
                        region_str = str(cell_region.value).strip()
                        # 지역명 매핑 확인
                        if region_str in region_mapping.values() or region_str == actual_region_name:
                            current_region = region_str
                        # 역매핑도 확인 (예: "서울특별시" -> "서울")
                        for short_name, long_name in region_mapping.items():
                            if long_name == region_str:
                                current_region = short_name
                                break
                        if current_region is None:
                            # 직접 매칭 시도
                            if (region_str == actual_region_name or 
                                actual_region_name in region_str or 
                                region_str in actual_region_name):
                                current_region = region_str
                    
                    # 연령계층이 있고 현재 시도가 매칭되면
                    if cell_age.value and current_region:
                        age_str = str(cell_age.value).strip()
                        
                        # 지역명 최종 확인
                        region_match = False
                        if current_region == actual_region_name:
                            region_match = True
                        elif actual_region_name in current_region or current_region in actual_region_name:
                            region_match = True
                        else:
                            # 역매핑 확인
                            mapped_current = region_mapping.get(current_region, current_region)
                            if mapped_current == actual_region_name or actual_region_name == current_region:
                                region_match = True
                        
                        if region_match and age_str == age_group_filter:
                            value = sheet.cell(row=row, column=current_col).value
                            if value is not None:
                                try:
                                    num_value = float(value)
                                    return self.format_number(num_value, decimal_places=1)
                                except (ValueError, TypeError):
                                    pass
                
                return "N/A"
            elif is_employment_rate_sheet:
                # 고용 시트 구조: 1열에 지역 코드, 2열에 지역 이름, 3열에 분류 단계, 4열에 연령대/항목
                # 고용률 계산: (취업자수 ÷ 15세 이상 인구 수) × 100
                current_col, _ = self._get_quarter_columns(year, quarter, sheet_name)
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                config = self._get_sheet_config(sheet_name)
                category_col = config['category_column']
                
                # 먼저 "계" 행에서 직접 값을 읽어봄 (이미 계산된 고용률이 있을 수 있음)
                for row in range(4, min(5000, sheet.max_row + 1)):
                    cell_b = sheet.cell(row=row, column=2)  # 지역 이름
                    cell_c = sheet.cell(row=row, column=3)  # 분류 단계
                    cell_category = sheet.cell(row=row, column=category_col)  # 연령대/항목
                    
                    if cell_b.value and cell_c.value and cell_category.value:
                        region_str = str(cell_b.value).strip()
                        class_str = str(cell_c.value).strip()
                        category_str = str(cell_category.value).strip()
                        
                        # 지역명 매칭 및 분류 단계가 0이고 카테고리가 "계"인 경우
                        if (region_str == region_name and 
                            class_str == '0' and 
                            category_str == '계'):
                            value = sheet.cell(row=row, column=current_col).value
                            if value is not None:
                                return self.format_number(value, decimal_places=1)
                
                # "계" 행에서 값을 찾지 못한 경우, 취업자수와 인구수로 계산
                employment_count = None  # 취업자수
                population_count = None   # 15세 이상 인구 수
                
                for row in range(4, min(5000, sheet.max_row + 1)):
                    cell_b = sheet.cell(row=row, column=2)  # 지역 이름
                    cell_c = sheet.cell(row=row, column=3)  # 분류 단계
                    cell_category = sheet.cell(row=row, column=category_col)  # 항목
                    
                    if cell_b.value and cell_c.value and cell_category.value:
                        region_str = str(cell_b.value).strip()
                        class_str = str(cell_c.value).strip()
                        category_str = str(cell_category.value).strip()
                        
                        # 지역명 매칭 및 분류 단계가 0인 경우
                        if (region_str == region_name and class_str == '0'):
                            value = sheet.cell(row=row, column=current_col).value
                            
                            # 취업자수 찾기
                            if ('취업자' in category_str or '취업' in category_str) and '인구' not in category_str:
                                if value is not None:
                                    employment_count = float(value) if isinstance(value, (int, float)) else None
                            
                            # 15세 이상 인구 수 찾기
                            if ('15세' in category_str or '15세이상' in category_str or '15세 이상' in category_str) and '인구' in category_str:
                                if value is not None:
                                    population_count = float(value) if isinstance(value, (int, float)) else None
                
                # 고용률 계산: (취업자수 ÷ 15세 이상 인구 수) × 100
                if employment_count is not None and population_count is not None and population_count != 0:
                    employment_rate = (employment_count / population_count) * 100
                    return self.format_number(employment_rate, decimal_places=1)
                
                return "N/A"
            else:
                # 일반 시트 처리
                return "N/A"
        
        # 분기별 지역별 실업률/고용률 값 (예: 전국_실업률_2024_3분기, 서울_고용률_2025_2분기)
        region_quarter_value_match = re.match(r'^([가-힣]+)_(실업률|고용률)_(\d{4})_(\d)분기$', key)
        if region_quarter_value_match:
            region_name = region_quarter_value_match.group(1)
            value_type = region_quarter_value_match.group(2)
            target_year = int(region_quarter_value_match.group(3))
            target_quarter = int(region_quarter_value_match.group(4))
            
            # 실업률 시트인지 고용률 시트인지 확인
            # key에 "실업률"이 포함되어 있으면 실업률 시트로 간주
            is_unemployment_sheet = ('실업' in sheet_name or sheet_name == '실업자 수' or value_type == '실업률')
            is_employment_rate_sheet = ('고용' in sheet_name or sheet_name == '고용' or '고용률' in sheet_name or sheet_name == '고용률' or value_type == '고용률')
            
            if is_unemployment_sheet:
                # 실업률 표 설정 가져오기
                sheet_config = self._get_sheet_config(sheet_name)
                unemployment_table_config = sheet_config.get('unemployment_rate_table', {})
                
                if not unemployment_table_config.get('enabled', True):
                    return "N/A"
                
                # 스키마에서 설정 가져오기
                start_row = unemployment_table_config.get('start_row', 81)
                title_row = unemployment_table_config.get('title_row', 79)
                region_col = unemployment_table_config.get('region_column', 1)
                age_group_col = unemployment_table_config.get('age_group_column', 2)
                age_group_filter = unemployment_table_config.get('age_group_filter', '계')
                header_row_number = unemployment_table_config.get('header_row_number', 3)
                region_mapping = unemployment_table_config.get('region_mapping', {})
                
                # 지역명 매핑
                actual_region_name = region_mapping.get(region_name, region_name)
                
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                # 헤더 행 찾기: "시도별 실업률(%)" 제목을 찾고 그 아래 헤더 행 찾기
                title_text = unemployment_table_config.get('title_text', '시도별 실업률(%)')
                header_row = None
                
                # 제목 행 찾기
                title_found = False
                for row in range(1, min(200, sheet.max_row + 1)):
                    for col in range(1, min(10, sheet.max_column + 1)):
                        cell_val = sheet.cell(row=row, column=col).value
                        if cell_val and title_text in str(cell_val):
                            title_found = True
                            # 제목 행 아래에서 헤더 행 찾기 (제목 행 + header_row_number 또는 제목 행 + 1, 2, 3 시도)
                            for offset in [header_row_number, 1, 2, 3]:
                                test_row = row + offset
                                if test_row <= sheet.max_row:
                                    # 헤더 행인지 확인 (연도/분기 패턴이 있는지 확인)
                                    for test_col in range(1, min(20, sheet.max_column + 1)):
                                        test_val = sheet.cell(row=test_row, column=test_col).value
                                        if test_val and (str(target_year) in str(test_val) or '분기' in str(test_val) or '/4' in str(test_val)):
                                            header_row = test_row
                                            break
                                    if header_row:
                                        break
                            break
                    if title_found and header_row:
                        break
                
                # 헤더 행을 찾지 못한 경우 설정값 사용
                if header_row is None:
                    header_row = title_row + header_row_number if title_row else header_row_number
                
                # 해당 분기의 열 번호 가져오기
                target_col, _ = self._get_quarter_columns(target_year, target_quarter, sheet_name)
                
                # 동적으로 찾지 못한 경우, 헤더 행에서 직접 헤더 찾기
                if target_col is None:
                    # 여러 패턴 시도: "2025  2/4", "2025 2/4", "2025년 2/4", "2025 2분기" 등
                    patterns = [
                        f"{target_year}  {target_quarter}/4",  # 공백 2개
                        f"{target_year} {target_quarter}/4",    # 공백 1개
                        f"{target_year}년 {target_quarter}/4",
                        f"{target_year} {target_quarter}분기",
                        f"{target_year}/{target_quarter}",
                    ]
                    for col in range(1, min(100, sheet.max_column + 1)):
                        header_val = sheet.cell(row=header_row, column=col).value
                        if header_val:
                            header_str = str(header_val).strip()
                            for pattern in patterns:
                                if pattern in header_str:
                                    target_col = col
                                    break
                            if target_col:
                                break
                
                if target_col is None:
                    return "N/A"
                
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                # 실업률 표에서 직접 값 가져오기
                # 시도별로 여러 행이 있고, 각 시도마다 연령계층별로 4개 행이 있음 (15~29세, 30~59세, 60세이상, 계)
                # '계'는 하단에 있음
                current_region = None
                for row in range(start_row, min(5000, sheet.max_row + 1)):
                    cell_region = sheet.cell(row=row, column=region_col)  # 시도
                    cell_age = sheet.cell(row=row, column=age_group_col)  # 연령계층
                    
                    # 시도명이 있으면 현재 시도 업데이트
                    if cell_region.value:
                        region_str = str(cell_region.value).strip()
                        # 지역명 매핑 확인
                        if region_str in region_mapping.values() or region_str == actual_region_name:
                            current_region = region_str
                        # 역매핑도 확인 (예: "서울특별시" -> "서울")
                        for short_name, long_name in region_mapping.items():
                            if long_name == region_str:
                                current_region = short_name
                                break
                        if current_region is None:
                            # 직접 매칭 시도
                            if (region_str == actual_region_name or 
                                actual_region_name in region_str or 
                                region_str in actual_region_name):
                                current_region = region_str
                    
                    # 연령계층이 있고 현재 시도가 매칭되면
                    if cell_age.value and current_region:
                        age_str = str(cell_age.value).strip()
                        
                        # 지역명 최종 확인
                        region_match = False
                        if current_region == actual_region_name:
                            region_match = True
                        elif actual_region_name in current_region or current_region in actual_region_name:
                            region_match = True
                        else:
                            # 역매핑 확인
                            mapped_current = region_mapping.get(current_region, current_region)
                            if mapped_current == actual_region_name or actual_region_name == current_region:
                                region_match = True
                        
                        if region_match and age_str == age_group_filter:
                            value = sheet.cell(row=row, column=target_col).value
                            if value is not None:
                                try:
                                    num_value = float(value)
                                    return self.format_number(num_value, decimal_places=1)
                                except (ValueError, TypeError):
                                    pass
                
                return "N/A"
            elif is_employment_rate_sheet:
                # 고용 시트 구조: 1열에 지역 코드, 2열에 지역 이름, 3열에 분류 단계, 4열에 연령대/항목
                # 고용률 계산: (취업자수 ÷ 15세 이상 인구 수) × 100
                target_col, _ = self._get_quarter_columns(target_year, target_quarter, sheet_name)
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                config = self._get_sheet_config(sheet_name)
                category_col = config['category_column']
                
                # 먼저 "계" 행에서 직접 값을 읽어봄
                for row in range(4, min(5000, sheet.max_row + 1)):
                    cell_b = sheet.cell(row=row, column=2)
                    cell_c = sheet.cell(row=row, column=3)
                    cell_category = sheet.cell(row=row, column=category_col)
                    
                    if cell_b.value and cell_c.value and cell_category.value:
                        region_str = str(cell_b.value).strip()
                        class_str = str(cell_c.value).strip()
                        category_str = str(cell_category.value).strip()
                        
                        if (region_str == region_name and class_str == '0' and category_str == '계'):
                            value = sheet.cell(row=row, column=target_col).value
                            if value is not None:
                                return self.format_number(value, decimal_places=1)
                
                # "계" 행에서 값을 찾지 못한 경우, 취업자수와 인구수로 계산
                employment_count = None
                population_count = None
                
                for row in range(4, min(5000, sheet.max_row + 1)):
                    cell_b = sheet.cell(row=row, column=2)
                    cell_c = sheet.cell(row=row, column=3)
                    cell_category = sheet.cell(row=row, column=category_col)
                    
                    if cell_b.value and cell_c.value and cell_category.value:
                        region_str = str(cell_b.value).strip()
                        class_str = str(cell_c.value).strip()
                        category_str = str(cell_category.value).strip()
                        
                        if (region_str == region_name and class_str == '0'):
                            value = sheet.cell(row=row, column=target_col).value
                            
                            if ('취업자' in category_str or '취업' in category_str) and '인구' not in category_str:
                                if value is not None:
                                    employment_count = float(value) if isinstance(value, (int, float)) else None
                            
                            if ('15세' in category_str or '15세이상' in category_str or '15세 이상' in category_str) and '인구' in category_str:
                                if value is not None:
                                    population_count = float(value) if isinstance(value, (int, float)) else None
                
                if employment_count is not None and population_count is not None and population_count != 0:
                    employment_rate = (employment_count / population_count) * 100
                    return self.format_number(employment_rate, decimal_places=1)
                
                return "N/A"
            
            return "N/A"
        
        # 분기별 지역별 증감률/증감pp (예: 전국_증감_2024_3분기, 서울_증감pp_2025_2분기)
        # 증감pp: 퍼센트포인트 차이 (현재값 - 이전값)
        region_quarter_growth_match = re.match(r'^([가-힣]+)_증감(?:pp)?_(\d{4})_(\d)분기$', key)
        if region_quarter_growth_match:
            region_name = region_quarter_growth_match.group(1)
            target_year = int(region_quarter_growth_match.group(2))
            target_quarter = int(region_quarter_growth_match.group(3))
            
            # 실업률 시트인지 고용률 시트인지 확인
            is_unemployment_sheet = ('실업' in sheet_name or sheet_name == '실업자 수')
            is_employment_rate_sheet = ('고용' in sheet_name or sheet_name == '고용' or '고용률' in sheet_name or sheet_name == '고용률')
            
            if is_unemployment_sheet:
                # 실업률 증감 계산: 현재 실업률 - 전년 동분기 실업률 (%p 단위)
                # 실업률 표 설정 가져오기
                sheet_config = self._get_sheet_config(sheet_name)
                unemployment_table_config = sheet_config.get('unemployment_rate_table', {})
                
                if not unemployment_table_config.get('enabled', True):
                    return "N/A"
                
                # 스키마에서 설정 가져오기
                start_row = unemployment_table_config.get('start_row', 81)
                title_row = unemployment_table_config.get('title_row', 79)
                region_col = unemployment_table_config.get('region_column', 1)
                age_group_col = unemployment_table_config.get('age_group_column', 2)
                age_group_filter = unemployment_table_config.get('age_group_filter', '계')
                header_row_number = unemployment_table_config.get('header_row_number', 3)
                region_mapping = unemployment_table_config.get('region_mapping', {})
                
                # 지역명 매핑
                actual_region_name = region_mapping.get(region_name, region_name)
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                # 헤더 행 찾기: "시도별 실업률(%)" 제목을 찾고 그 아래 헤더 행 찾기
                title_text = unemployment_table_config.get('title_text', '시도별 실업률(%)')
                header_row = None
                
                # 제목 행 찾기
                title_found = False
                for row in range(1, min(200, sheet.max_row + 1)):
                    for col in range(1, min(10, sheet.max_column + 1)):
                        cell_val = sheet.cell(row=row, column=col).value
                        if cell_val and title_text in str(cell_val):
                            title_found = True
                            # 제목 행 아래에서 헤더 행 찾기 (제목 행 + header_row_number 또는 제목 행 + 1, 2, 3 시도)
                            for offset in [header_row_number, 1, 2, 3]:
                                test_row = row + offset
                                if test_row <= sheet.max_row:
                                    # 헤더 행인지 확인 (연도/분기 패턴이 있는지 확인)
                                    for test_col in range(1, min(20, sheet.max_column + 1)):
                                        test_val = sheet.cell(row=test_row, column=test_col).value
                                        if test_val and ('분기' in str(test_val) or '/4' in str(test_val)):
                                            header_row = test_row
                                            break
                                    if header_row:
                                        break
                            break
                    if title_found and header_row:
                        break
                
                # 헤더 행을 찾지 못한 경우 설정값 사용
                if header_row is None:
                    header_row = title_row + header_row_number if title_row else header_row_number
                
                # 실업률 가져오기 헬퍼 함수 (실업률 표에서 직접 읽기)
                def get_unemployment_rate(calc_year, calc_quarter, region_name_val):
                    """특정 연도/분기의 실업률을 실업률 표에서 가져옵니다."""
                    target_col, _ = self._get_quarter_columns(calc_year, calc_quarter, sheet_name)
                    
                    # 동적으로 찾지 못한 경우, 헤더 행에서 직접 헤더 찾기
                    if target_col is None:
                        # 여러 패턴 시도: "2025  2/4", "2025 2/4", "2025년 2/4", "2025 2분기" 등
                        patterns = [
                            f"{calc_year}  {calc_quarter}/4",  # 공백 2개
                            f"{calc_year} {calc_quarter}/4",    # 공백 1개
                            f"{calc_year}년 {calc_quarter}/4",
                            f"{calc_year} {calc_quarter}분기",
                            f"{calc_year}/{calc_quarter}",
                        ]
                        for col in range(1, min(100, sheet.max_column + 1)):
                            header_val = sheet.cell(row=header_row, column=col).value
                            if header_val:
                                header_str = str(header_val).strip()
                                for pattern in patterns:
                                    if pattern in header_str:
                                        target_col = col
                                        break
                                if target_col:
                                    break
                    
                    if target_col is None:
                        return None
                    
                    current_region = None
                    
                    for row in range(start_row, min(5000, sheet.max_row + 1)):
                        cell_region = sheet.cell(row=row, column=region_col)  # 시도
                        cell_age = sheet.cell(row=row, column=age_group_col)  # 연령계층
                        
                        # 시도명이 있으면 현재 시도 업데이트
                        if cell_region.value:
                            region_str = str(cell_region.value).strip()
                            # 지역명 매핑 확인
                            if region_str in region_mapping.values() or region_str == region_name_val:
                                current_region = region_str
                            # 역매핑도 확인 (예: "서울특별시" -> "서울")
                            for short_name, long_name in region_mapping.items():
                                if long_name == region_str:
                                    current_region = short_name
                                    break
                            if current_region is None:
                                # 직접 매칭 시도
                                if (region_str == region_name_val or 
                                    region_name_val in region_str or 
                                    region_str in region_name_val):
                                    current_region = region_str
                        
                        # 연령계층이 있고 현재 시도가 매칭되면
                        if cell_age.value and current_region:
                            age_str = str(cell_age.value).strip()
                            
                            # 지역명 최종 확인
                            region_match = False
                            if current_region == region_name_val:
                                region_match = True
                            elif region_name_val in current_region or current_region in region_name_val:
                                region_match = True
                            else:
                                # 역매핑 확인
                                mapped_current = region_mapping.get(current_region, current_region)
                                if mapped_current == region_name_val or region_name_val == current_region:
                                    region_match = True
                            
                            if region_match and age_str == age_group_filter:
                                value = sheet.cell(row=row, column=target_col).value
                                if value is not None:
                                    try:
                                        return float(value)
                                    except (ValueError, TypeError):
                                        pass
                                break
                    return None
                
                # 현재 실업률과 전년 동분기 실업률 가져오기
                current_rate = get_unemployment_rate(target_year, target_quarter, actual_region_name)
                prev_year = target_year - 1
                prev_rate = get_unemployment_rate(prev_year, target_quarter, actual_region_name)
                
                if current_rate is not None and prev_rate is not None:
                    # 실업률 증감은 %p 단위 (퍼센트 포인트)
                    growth_rate_pp = current_rate - prev_rate
                    return self.format_percentage(growth_rate_pp, decimal_places=1, include_percent=False)
                
                return "N/A"
            elif is_employment_rate_sheet:
                # 고용 시트 구조: 1열에 지역 코드, 2열에 지역 이름, 3열에 분류 단계, 4열에 연령대/항목
                # 고용률 증감률 계산: 현재 고용률 - 전년 동분기 고용률 (%p 단위)
                current_col, prev_col = self._get_quarter_columns(target_year, target_quarter, sheet_name)
                sheet = self.excel_extractor.get_sheet(sheet_name)
                
                config = self._get_sheet_config(sheet_name)
                category_col = config['category_column']
                
                # 지역별 고용률 계산 함수
                def calculate_employment_rate(region_name, col):
                    # 먼저 "계" 행에서 직접 값을 읽어봄
                    for row in range(4, min(5000, sheet.max_row + 1)):
                        cell_b = sheet.cell(row=row, column=2)
                        cell_c = sheet.cell(row=row, column=3)
                        cell_category = sheet.cell(row=row, column=category_col)
                        
                        if cell_b.value and cell_c.value and cell_category.value:
                            region_str = str(cell_b.value).strip()
                            class_str = str(cell_c.value).strip()
                            category_str = str(cell_category.value).strip()
                            
                            if (region_str == region_name and class_str == '0' and category_str == '계'):
                                value = sheet.cell(row=row, column=col).value
                                if value is not None:
                                    return float(value) if isinstance(value, (int, float)) else None
                    
                    # "계" 행에서 값을 찾지 못한 경우, 취업자수와 인구수로 계산
                    employment_count = None
                    population_count = None
                    
                    for row in range(4, min(5000, sheet.max_row + 1)):
                        cell_b = sheet.cell(row=row, column=2)
                        cell_c = sheet.cell(row=row, column=3)
                        cell_category = sheet.cell(row=row, column=category_col)
                        
                        if cell_b.value and cell_c.value and cell_category.value:
                            region_str = str(cell_b.value).strip()
                            class_str = str(cell_c.value).strip()
                            category_str = str(cell_category.value).strip()
                            
                            if (region_str == region_name and class_str == '0'):
                                value = sheet.cell(row=row, column=col).value
                                
                                if ('취업자' in category_str or '취업' in category_str) and '인구' not in category_str:
                                    if value is not None:
                                        employment_count = float(value) if isinstance(value, (int, float)) else None
                                
                                if ('15세' in category_str or '15세이상' in category_str or '15세 이상' in category_str) and '인구' in category_str:
                                    if value is not None:
                                        population_count = float(value) if isinstance(value, (int, float)) else None
                    
                    if employment_count is not None and population_count is not None and population_count != 0:
                        return (employment_count / population_count) * 100
                    return None
                
                # 지역별 고용률 계산
                current_rate = calculate_employment_rate(region_name, current_col)
                prev_rate = calculate_employment_rate(region_name, prev_col)
                
                if current_rate is not None and prev_rate is not None:
                    # 고용률 증감은 %p 단위 (퍼센트 포인트)
                    growth_rate_pp = current_rate - prev_rate
                    return self.format_percentage(growth_rate_pp, decimal_places=1, include_percent=False)
                
                return "N/A"
            
            return "N/A"
        
        # 마지막으로 추론 로직 시도
        inferred_value = self._try_infer_marker_value(sheet_name, key, year, quarter)
        if inferred_value is not None:
            return inferred_value
        
        return "N/A"
    
    def _try_infer_marker_value(self, sheet_name: str, key: str, year: int, quarter: int) -> Optional[str]:
        """
        N/A가 발생했을 때 값을 추론하기 위해 다양한 방법을 시도합니다.
        
        Args:
            sheet_name: 시트 이름
            key: 동적 키
            year: 연도
            quarter: 분기
            
        Returns:
            추론된 값 또는 None
        """
        try:
            # 전국_증감률 추론: 직접 계산 시도
            if key == '전국_증감률':
                growth_rate = self.dynamic_parser.calculate_growth_rate(sheet_name, '전국', year, quarter)
                if growth_rate is not None:
                    return self.format_percentage(growth_rate, decimal_places=1, include_percent=False)
                # 전국 지역명 매핑 시도
                for national_name in ['전국', '국', '전국계']:
                    growth_rate = self.dynamic_parser.calculate_growth_rate(sheet_name, national_name, year, quarter)
                    if growth_rate is not None:
                        return self.format_percentage(growth_rate, decimal_places=1, include_percent=False)
            
            # 전국_이름 추론: 전국 지역 찾기
            if key == '전국_이름':
                # 시트에서 전국 관련 행 찾기
                try:
                    sheet = self.excel_extractor.get_sheet(sheet_name)
                    structure = self.dynamic_parser.parse_sheet_structure(sheet_name)
                    region_col = structure.get('region_column', 2)
                    data_start_row = structure.get('data_start_row', 4)
                    
                    for row in range(data_start_row, min(data_start_row + 200, sheet.max_row + 1)):
                        cell_region = sheet.cell(row=row, column=region_col)
                        if cell_region.value:
                            region_str = str(cell_region.value).strip()
                            if region_str in ['전국', '국', '전국계', '계']:
                                return '전국'
                except:
                    pass
                # 기본값 반환
                return '전국'
            
            # 상위시도/하위시도 패턴 추론: 직접 계산 시도
            top_match = re.match(r'상위시도(\d+)_(.+)', key)
            if top_match:
                idx = int(top_match.group(1)) - 1
                field = top_match.group(2)
                # 직접 지역 리스트 가져오기
                current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                if current_col and prev_col:
                    try:
                        regions_data = self.data_analyzer.get_regions_with_growth_rate(sheet_name, current_col, prev_col)
                        if regions_data:
                            # 증감률 기준 정렬
                            sorted_regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0), reverse=True)
                            if idx < len(sorted_regions):
                                region = sorted_regions[idx]
                                if field == '이름':
                                    return region.get('name', '')
                                elif field == '증감률':
                                    return self.format_percentage(region.get('growth_rate', 0), decimal_places=1, include_percent=False)
                                elif field.startswith('업태') or field.startswith('산업'):
                                    # 산업/업태 패턴도 처리
                                    industry_match = re.match(r'(업태|산업)(\d+)_(.+)', field)
                                    if industry_match:
                                        industry_idx = int(industry_match.group(2)) - 1
                                        industry_field = industry_match.group(3)
                                        categories = self._get_categories_for_region(
                                            sheet_name, region.get('name', ''), year, quarter, top_n=3
                                        )
                                        if categories and industry_idx < len(categories):
                                            category = categories[industry_idx]
                                            if industry_field == '이름':
                                                return category.get('name', '')
                                            elif industry_field == '증감률':
                                                return self.format_percentage(category.get('growth_rate', 0), decimal_places=1, include_percent=False)
                    except Exception:
                        pass
            
            bottom_match = re.match(r'하위시도(\d+)_(.+)', key)
            if bottom_match:
                idx = int(bottom_match.group(1)) - 1
                field = bottom_match.group(2)
                # 직접 지역 리스트 가져오기
                current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                if current_col and prev_col:
                    try:
                        regions_data = self.data_analyzer.get_regions_with_growth_rate(sheet_name, current_col, prev_col)
                        if regions_data:
                            # 증감률 기준 정렬 (역순)
                            sorted_regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0))
                            if idx < len(sorted_regions):
                                region = sorted_regions[idx]
                                if field == '이름':
                                    return region.get('name', '')
                                elif field == '증감률':
                                    return self.format_percentage(region.get('growth_rate', 0), decimal_places=1, include_percent=False)
                                elif field.startswith('업태') or field.startswith('산업'):
                                    # 산업/업태 패턴도 처리
                                    industry_match = re.match(r'(업태|산업)(\d+)_(.+)', field)
                                    if industry_match:
                                        industry_idx = int(industry_match.group(2)) - 1
                                        industry_field = industry_match.group(3)
                                        categories = self._get_categories_for_region(
                                            sheet_name, region.get('name', ''), year, quarter, top_n=3
                                        )
                                        if categories and industry_idx < len(categories):
                                            category = categories[industry_idx]
                                            if industry_field == '이름':
                                                return category.get('name', '')
                                            elif industry_field == '증감률':
                                                return self.format_percentage(category.get('growth_rate', 0), decimal_places=1, include_percent=False)
                    except Exception:
                        pass
            
            # 강조시도 패턴 추론
            emphasis_match = re.match(r'강조시도(\d+)_(.+)', key)
            if emphasis_match:
                idx = int(emphasis_match.group(1)) - 1
                field = emphasis_match.group(2)
                # 전국 증감률 먼저 확인
                national_growth = self.dynamic_parser.calculate_growth_rate(sheet_name, '전국', year, quarter)
                if national_growth is not None:
                    current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
                    if current_col and prev_col:
                        try:
                            regions_data = self.data_analyzer.get_regions_with_growth_rate(sheet_name, current_col, prev_col)
                            if regions_data:
                                if national_growth >= 0:
                                    sorted_regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0), reverse=True)
                                else:
                                    sorted_regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0))
                                
                                if idx < len(sorted_regions):
                                    region = sorted_regions[idx]
                                    if field == '이름':
                                        return region.get('name', '')
                                    elif field == '증감률':
                                        return self.format_percentage(region.get('growth_rate', 0), decimal_places=1, include_percent=False)
                                    elif field.startswith('업태') or field.startswith('산업'):
                                        industry_match = re.match(r'(업태|산업)(\d+)_(.+)', field)
                                        if industry_match:
                                            industry_idx = int(industry_match.group(2)) - 1
                                            industry_field = industry_match.group(3)
                                            categories = self._get_categories_for_region(
                                                sheet_name, region.get('name', ''), year, quarter, top_n=3
                                            )
                                            if categories and industry_idx < len(categories):
                                                category = categories[industry_idx]
                                                if industry_field == '이름':
                                                    return category.get('name', '')
                                                elif industry_field == '증감률':
                                                    return self.format_percentage(category.get('growth_rate', 0), decimal_places=1, include_percent=False)
                        except Exception:
                            pass
            
            # 전국_업태/전국_산업 패턴 추론
            if key.startswith('전국_업태') or key.startswith('전국_산업'):
                industry_match = re.match(r'전국_(업태|산업)(\d+)_(.+)', key)
                if industry_match:
                    industry_idx = int(industry_match.group(2)) - 1
                    industry_field = industry_match.group(3)
                    try:
                        categories = self._get_categories_for_region(sheet_name, '전국', year, quarter, top_n=3)
                        if categories and industry_idx < len(categories):
                            category = categories[industry_idx]
                            if industry_field == '이름':
                                return category.get('name', '')
                            elif industry_field == '증감률':
                                return self.format_percentage(category.get('growth_rate', 0), decimal_places=1, include_percent=False)
                    except Exception:
                        pass
            
            # 분기별 증감률 추론: 직접 계산 시도
            quarterly_match = re.match(r'(.+)_(\d{4})_(\d)분기_증감률', key)
            if quarterly_match:
                region_name = quarterly_match.group(1)
                target_year = int(quarterly_match.group(2))
                target_quarter = int(quarterly_match.group(3))
                growth_rate = self._get_quarterly_growth_rate(sheet_name, region_name, f'{target_year}_{target_quarter}분기')
                if growth_rate is not None:
                    return self.format_percentage(growth_rate, decimal_places=1, include_percent=False)
            
            # 지역별 분기별 증감률 추론
            region_quarterly_match = re.match(r'^([가-힣]+)_(\d{4})_(\d)분기$', key)
            if region_quarterly_match:
                region_name = region_quarterly_match.group(1)
                target_year = int(region_quarterly_match.group(2))
                target_quarter = int(region_quarterly_match.group(3))
                growth_rate = self._get_quarterly_growth_rate(sheet_name, region_name, f'{target_year}_{target_quarter}분기')
                if growth_rate is not None:
                    return self.format_percentage(growth_rate, decimal_places=1, include_percent=False)
            
        except Exception as e:
            # 추론 중 오류 발생 시 None 반환 (기존 N/A 유지)
            pass
        
        return None
    
    def process_marker(self, marker_info: Dict[str, str]) -> str:
        """
        마커 정보를 처리하여 값을 추출하고 계산합니다.
        유연한 매핑을 사용하여 시트명과 컬럼명이 바뀌어도 자동으로 매핑합니다.
        
        Args:
            marker_info: 마커 정보 딕셔너리
                - 'sheet_name': 시트명
                - 'cell_address': 셀 주소 또는 범위 또는 헤더 기반 키
                - 'operation': 계산식 (선택적)
        
        Returns:
            처리된 값 (문자열)
        """
        marker_sheet_name = marker_info['sheet_name']
        cell_address = marker_info['cell_address']
        operation = marker_info.get('operation')
        
        try:
            # 동적 마커인지 확인 (셀 주소 형식이 아닌 경우)
            if not re.match(r'^[A-Z]+\d+', cell_address):
                # 실제 시트명 찾기 (유연한 매핑 사용)
                actual_sheet_name = self.flexible_mapper.find_sheet_by_name(marker_sheet_name)
                if not actual_sheet_name:
                    actual_sheet_name = marker_sheet_name
                
                # 동적 마커 처리 (현재 저장된 연도/분기 사용)
                current_year = self._current_year if self._current_year is not None else 2025
                current_quarter = self._current_quarter if self._current_quarter is not None else 2
                dynamic_value = self._process_dynamic_marker(actual_sheet_name, cell_address, current_year, current_quarter)
                if dynamic_value is not None and dynamic_value != "N/A":
                    return dynamic_value
                
                # 동적 마커가 실패하면 유연한 매핑 시도
                resolved = self.flexible_mapper.resolve_marker(marker_sheet_name, cell_address)
                if resolved:
                    actual_sheet, resolved_address = resolved
                    try:
                        raw_value = self.excel_extractor.extract_value(actual_sheet, resolved_address)
                        
                        # 결측치 체크
                        if raw_value is None:
                            return "N/A"
                        if isinstance(raw_value, str) and not raw_value.strip():
                            return "N/A"
                        if isinstance(raw_value, list):
                            filtered = [v for v in raw_value if v is not None and (not isinstance(v, str) or v.strip())]
                            if not filtered:
                                return "N/A"
                            raw_value = filtered
                        
                        # 계산 처리
                        if operation:
                            calculated = self._apply_operation(raw_value, operation)
                            return self._format_value(calculated, operation)
                        return self._format_value(raw_value)
                    except Exception:
                        return "N/A"
                
                # 동적 마커를 찾을 수 없으면 N/A 반환
                return "N/A"
            
            # 셀 주소 형식인 경우
            # 먼저 정확한 시트명으로 시도
            raw_value = None
            try:
                raw_value = self.excel_extractor.extract_value(marker_sheet_name, cell_address)
            except Exception:
                # 시트를 찾을 수 없으면 유연한 매핑 시도
                actual_sheet = self.flexible_mapper.find_sheet_by_name(marker_sheet_name)
                if actual_sheet:
                    try:
                        raw_value = self.excel_extractor.extract_value(actual_sheet, cell_address)
                    except Exception:
                        return "N/A"
                else:
                    return "N/A"
            
            # 결측치 체크 (None, 빈 문자열, 공백만 있는 값)
            if raw_value is None:
                return "N/A"
            if isinstance(raw_value, str) and not raw_value.strip():
                return "N/A"
            if isinstance(raw_value, list):
                # 리스트인 경우 None이나 빈 값 필터링
                filtered = [v for v in raw_value if v is not None and (not isinstance(v, str) or v.strip())]
                if not filtered:
                    return "N/A"
                raw_value = filtered
            
            # 계산 처리
            if operation:
                calculated = self._apply_operation(raw_value, operation)
            else:
                # 계산이 필요없는 경우
                if isinstance(raw_value, list):
                    # 범위인 경우 첫 번째 값 사용
                    calculated = raw_value[0] if raw_value else None
                else:
                    calculated = raw_value
            
            # 결과 포맷팅
            return self._format_value(calculated, operation)
                
        except Exception as e:
            # 에러 발생 시 N/A 반환
            return "N/A"
    
    def _apply_operation(self, raw_value: Any, operation: str) -> Any:
        """계산식을 적용합니다. 오류 발생 시 None 반환."""
        try:
            operation_lower = operation.lower().strip()
            
            # 결측치 체크
            if raw_value is None:
                return None
            
            # 범위인 경우 (리스트)
            if isinstance(raw_value, list):
                # 빈 리스트 체크
                if not raw_value:
                    return None
                
                # None 값 필터링
                filtered_values = [v for v in raw_value if v is not None and str(v).strip() != '']
                if not filtered_values:
                    return None
                
                # 증감률/증감액 계산의 경우 범위에서 첫 두 값 사용
                if operation_lower in ['growth_rate', '증감률', '증가율', 'growth_amount', '증감액', '증가액']:
                    if len(filtered_values) >= 2:
                        try:
                            return self.calculator.calculate(operation_lower, filtered_values[0], filtered_values[1])
                        except (ValueError, ZeroDivisionError, TypeError):
                            return None
                    else:
                        return None
                else:
                    # 다른 계산 (sum, average, max, min 등)
                    try:
                        return self.calculator.calculate_from_cell_refs(operation_lower, filtered_values)
                    except (ValueError, ZeroDivisionError, TypeError):
                        return None
            else:
                # 단일 값인 경우
                # 빈 문자열이나 공백만 있는 값 체크
                if isinstance(raw_value, str) and not raw_value.strip():
                    return None
                
                # 일부 계산식은 단일 값도 처리 (예: format)
                if operation_lower in ['format', '포맷']:
                    try:
                        return self.format_number(raw_value, decimal_places=1)
                    except (ValueError, TypeError):
                        return None
                # 계산식이 있지만 단일 값인 경우 그대로 사용
                return raw_value
        except Exception:
            # 모든 예외를 None으로 변환
            return None
    
    def _format_value(self, value: Any, operation: str = None) -> str:
        """값을 포맷팅합니다. 결측치나 오류는 N/A로 반환."""
        # None 체크
        if value is None:
            return "N/A"
        
        # 빈 문자열이나 공백만 있는 값 체크
        if isinstance(value, str):
            if not value.strip():
                return "N/A"
            # "N/A", "n/a", "null", "None" 등의 문자열도 N/A로 처리
            value_lower = value.strip().lower()
            if value_lower in ['n/a', 'na', 'null', 'none', '#n/a', '#na', '#null', '#ref!', '#value!', '#div/0!']:
                return "N/A"
        
        try:
            float_val = float(value)
            # NaN이나 Infinity 체크
            if math.isnan(float_val) or math.isinf(float_val):
                return "N/A"
            
            # 퍼센트 연산인 경우 퍼센트 포맷
            if operation and operation.lower() in ['growth_rate', '증감률', '증가율']:
                return self.format_percentage(float_val, decimal_places=1)
            # 일반 숫자 포맷팅 (천 단위 구분, 소수점 첫째자리까지)
            return self.format_number(float_val, use_comma=True, decimal_places=1)
        except (ValueError, TypeError, OverflowError):
            # 숫자가 아닌 경우 그대로 반환 (하지만 빈 문자열은 이미 N/A 처리됨)
            return str(value) if value else "N/A"
    
    def fill_template(self, sheet_name: str = None, year: int = None, quarter: int = None) -> str:
        """
        템플릿의 모든 마커를 처리하여 완성된 템플릿을 반환합니다.
        
        Args:
            sheet_name: 시트 이름 (기본값: None, 마커에서 추출)
            year: 연도 (None이면 자동 감지)
            quarter: 분기 (None이면 자동 감지)
        
        Returns:
            모든 마커가 값으로 치환된 HTML 템플릿
        """
        # 연도/분기가 지정되지 않으면 자동 감지
        if sheet_name and (year is None or quarter is None):
            periods_info = self.period_detector.detect_available_periods(sheet_name)
            if year is None:
                year = periods_info['default_year']
            if quarter is None:
                quarter = periods_info['default_quarter']
        
        # 현재 처리 중인 연도/분기/시트명 저장
        self._current_year = year
        self._current_quarter = quarter
        self._current_sheet_name = sheet_name
        # 템플릿 로드
        template_content = self.template_manager.get_template_content()
        
        # CSS와 스크립트 섹션을 제외하고 처리
        # <style>...</style> 및 <script>...</script> 태그 내부는 제외
        style_pattern = re.compile(r'<style[^>]*>.*?</style>', re.IGNORECASE | re.DOTALL)
        script_pattern = re.compile(r'<script[^>]*>.*?</script>', re.IGNORECASE | re.DOTALL)
        
        # CSS와 스크립트 섹션을 임시로 마스킹
        style_placeholders = {}
        script_placeholders = {}
        
        def style_replacer(match):
            placeholder = f"__STYLE_PLACEHOLDER_{len(style_placeholders)}__"
            style_placeholders[placeholder] = match.group(0)
            return placeholder
        
        def script_replacer(match):
            placeholder = f"__SCRIPT_PLACEHOLDER_{len(script_placeholders)}__"
            script_placeholders[placeholder] = match.group(0)
            return placeholder
        
        # CSS와 스크립트 섹션을 임시로 교체
        template_content = style_pattern.sub(style_replacer, template_content)
        template_content = script_pattern.sub(script_replacer, template_content)
        
        # 동적 마커 패턴: {시트명:동적키} (한글 시트명 포함)
        # 실제 마커만 매칭하도록 더 구체적으로 (한글이나 특정 패턴 포함)
        dynamic_marker_pattern = re.compile(r'\{([^:{}]+):([^:}]+)\}')
        
        # 동적 마커 찾기 및 치환
        for match in dynamic_marker_pattern.finditer(template_content):
            full_match = match.group(0)
            sheet_name = match.group(1)
            dynamic_key = match.group(2)
            
            # 셀 주소 형식이 아닌 경우 동적 마커로 처리
            # 하지만 CSS 속성명 같은 것들은 제외 (한글이나 숫자가 포함된 경우만)
            if not re.match(r'^[A-Z]+\d+', dynamic_key):
                # 실제 마커인지 확인 (한글, 숫자, 언더스코어 포함)
                if re.search(r'[가-힣0-9_]', dynamic_key):
                    try:
                        # 마커에서 추출한 시트명을 항상 사용 (파라미터의 sheet_name은 무시)
                        marker_sheet_name = match.group(1)
                        # 실제 시트명 찾기 (유연한 매핑 사용)
                        actual_sheet_name = self.flexible_mapper.find_sheet_by_name(marker_sheet_name)
                        if not actual_sheet_name:
                            actual_sheet_name = marker_sheet_name
                        # 현재 저장된 연도/분기 사용
                        current_year = year if year is not None else (self._current_year or 2025)
                        current_quarter = quarter if quarter is not None else (self._current_quarter or 2)
                        dynamic_value = self._process_dynamic_marker(actual_sheet_name, dynamic_key, current_year, current_quarter)
                        if dynamic_value is not None:
                            template_content = template_content.replace(full_match, dynamic_value)
                        else:
                            # 값이 없으면 N/A로 채움
                            template_content = template_content.replace(full_match, 'N/A')
                    except Exception:
                        # 에러 발생 시 N/A로 채움
                        template_content = template_content.replace(full_match, 'N/A')
        
        # 일반 마커 추출 및 처리
        markers = self.template_manager.extract_markers()
        
        # 각 마커를 처리하여 치환
        for marker_info in markers:
            marker_str = marker_info['full_match']
            # sheet_name이 제공되면 마커의 시트명을 덮어쓰기
            if sheet_name:
                marker_info['sheet_name'] = sheet_name
            else:
                # 시트명이 제공되지 않으면 유연한 매핑으로 찾기
                marker_sheet = marker_info['sheet_name']
                actual_sheet = self.flexible_mapper.find_sheet_by_name(marker_sheet)
                if actual_sheet:
                    marker_info['sheet_name'] = actual_sheet
            
            processed_value = self.process_marker(marker_info)
            
            # 값이 비어있거나 에러인 경우 N/A로 채움
            if not processed_value or processed_value.startswith('[에러'):
                processed_value = 'N/A'
            
            # 마커를 처리된 값으로 치환
            template_content = template_content.replace(marker_str, processed_value)
        
        # CSS와 스크립트 섹션 복원
        for placeholder, original in style_placeholders.items():
            template_content = template_content.replace(placeholder, original)
        for placeholder, original in script_placeholders.items():
            template_content = template_content.replace(placeholder, original)
        
        return template_content
    
    def fill_template_with_custom_format(self, format_func: callable = None) -> str:
        """
        커스텀 포맷팅 함수를 사용하여 템플릿을 채웁니다.
        
        Args:
            format_func: 포맷팅 함수 (marker_info, raw_value) -> str
        
        Returns:
            완성된 템플릿
        """
        template_content = self.template_manager.get_template_content()
        markers = self.template_manager.extract_markers()
        
        for marker_info in markers:
            marker_str = marker_info['full_match']
            
            try:
                raw_value = self.excel_extractor.extract_value(
                    marker_info['sheet_name'],
                    marker_info['cell_address']
                )
                
                if format_func:
                    processed_value = format_func(marker_info, raw_value)
                else:
                    processed_value = self.process_marker(marker_info)
                
                template_content = template_content.replace(marker_str, processed_value)
            except Exception as e:
                # 에러 발생 시 마커를 에러 메시지로 치환
                template_content = template_content.replace(
                    marker_str, 
                    f"[에러: {str(e)}]"
                )
        
        return template_content
    
    def _get_age_group_value(self, sheet_name: str, region_name: str, age_group_name: str, 
                             direction: str, year: int, quarter: int) -> str:
        """
        연령별 인구이동 시트에서 연령 범위를 합산한 값을 반환합니다.
        
        Args:
            sheet_name: 시트 이름 ('연령별 인구이동')
            region_name: 지역 이름
            age_group_name: 연령 그룹 이름 (예: '0~14세', '15~29세')
            direction: '유입' 또는 '유출'
            year: 연도
            quarter: 분기
            
        Returns:
            합산된 값 (문자열)
        """
        try:
            # 연령별 인구이동 시트 설정 가져오기
            config = self.schema_loader.load_sheet_config('연령별 인구이동')
            age_group_config = config.get('age_group_aggregation', {})
            groups = age_group_config.get('groups', [])
            
            # 해당 연령 그룹 찾기
            target_group = None
            for group in groups:
                if group['name'] == age_group_name:
                    target_group = group
                    break
            
            if not target_group:
                return "N/A"
            
            # 연령별 인구이동 시트 사용 (시도 간 이동 시트와 별도)
            age_sheet_name = '연령별 인구이동'
            current_col, prev_col = self._get_quarter_columns(year, quarter, age_sheet_name)
            if current_col is None:
                return "N/A"
            sheet = self.excel_extractor.get_sheet(age_sheet_name)
            
            # 지역명 매핑
            # 짧은 이름으로 정규화 (identity 매핑은 불필요)
            mapped_region_name = self._normalize_region_name(region_name)
            
            # 지역의 시작 행 찾기
            region_start_row = None
            for row in range(4, min(5000, sheet.max_row + 1)):
                cell_region = sheet.cell(row=row, column=2).value
                cell_classification = sheet.cell(row=row, column=3).value
                cell_category = sheet.cell(row=row, column=4).value
                
                if (cell_region and str(cell_region).strip() == mapped_region_name and 
                    cell_classification == 0 and cell_category == '계'):
                    region_start_row = row
                    break
            
            if not region_start_row:
                return "N/A"
            
            # 분류단계 1의 연령 범위 합산
            total_value = 0
            age_ranges = target_group.get('age_ranges', [])
            
            for row in range(region_start_row + 1, min(region_start_row + 50, sheet.max_row + 1)):
                cell_region = sheet.cell(row=row, column=2).value
                cell_classification = sheet.cell(row=row, column=3).value
                cell_age = sheet.cell(row=row, column=4).value
                
                # 같은 지역이고 분류단계 1인 경우
                if (cell_region and str(cell_region).strip() == mapped_region_name and 
                    cell_classification == 1 and cell_age):
                    age_str = str(cell_age).strip()
                    
                    # 연령 범위에 포함되는지 확인
                    if any(age_range in age_str or age_str in age_range for age_range in age_ranges):
                        value = sheet.cell(row=row, column=current_col).value
                        if value is not None:
                            try:
                                # 유입/유출에 따라 부호 처리
                                # 유입은 양수, 유출은 음수 (또는 반대일 수 있음)
                                num_value = float(value)
                                if direction == '유입':
                                    # 양수만 합산
                                    if num_value > 0:
                                        total_value += num_value
                                elif direction == '유출':
                                    # 음수만 합산 (절대값)
                                    if num_value < 0:
                                        total_value += abs(num_value)
                            except (ValueError, TypeError):
                                pass
                
                # 다른 지역이 나오면 중단
                if cell_region and str(cell_region).strip() != mapped_region_name:
                    break
            
            return str(int(total_value)) if total_value > 0 else "0"
            
        except Exception as e:
            return "N/A"
    
    def _get_age_groups_for_region(self, sheet_name: str, region_name: str, 
                                    year: int, quarter: int, top_n: int = 3) -> list:
        """
        연령별 인구이동 시트에서 연령 그룹별 데이터를 반환합니다.
        
        Args:
            sheet_name: 시트 이름 ('연령별 인구이동')
            region_name: 지역 이름
            year: 연도
            quarter: 분기
            top_n: 상위 개수
            
        Returns:
            연령 그룹별 정보 리스트
        """
        try:
            # 연령별 인구이동 시트 설정 가져오기
            config = self.schema_loader.load_sheet_config('연령별 인구이동')
            age_group_config = config.get('age_group_aggregation', {})
            groups = age_group_config.get('groups', [])
            
            # 연령별 인구이동 시트 사용 (시도 간 이동 시트와 별도)
            age_sheet_name = '연령별 인구이동'
            current_col, prev_col = self._get_quarter_columns(year, quarter, age_sheet_name)
            if current_col is None:
                return []
            sheet = self.excel_extractor.get_sheet(age_sheet_name)
            
            # 지역명 매핑
            # 짧은 이름으로 정규화 (identity 매핑은 불필요)
            mapped_region_name = self._normalize_region_name(region_name)
            
            # 지역의 시작 행 찾기
            region_start_row = None
            for row in range(4, min(5000, sheet.max_row + 1)):
                cell_region = sheet.cell(row=row, column=2).value
                cell_classification = sheet.cell(row=row, column=3).value
                cell_category = sheet.cell(row=row, column=4).value
                
                if (cell_region and str(cell_region).strip() == mapped_region_name and 
                    cell_classification == 0 and cell_category == '계'):
                    region_start_row = row
                    break
            
            if not region_start_row:
                return []
            
            # 각 연령 그룹별로 합산
            age_groups = []
            for group in groups:
                group_name = group['name']
                age_ranges = group.get('age_ranges', [])
                
                # 유입/유출 값 합산
                inflow_total = 0
                outflow_total = 0
                
                for row in range(region_start_row + 1, min(region_start_row + 50, sheet.max_row + 1)):
                    cell_region = sheet.cell(row=row, column=2).value
                    cell_classification = sheet.cell(row=row, column=3).value
                    cell_age = sheet.cell(row=row, column=4).value
                    
                    # 같은 지역이고 분류단계 1인 경우
                    if (cell_region and str(cell_region).strip() == mapped_region_name and 
                        cell_classification == 1 and cell_age):
                        age_str = str(cell_age).strip()
                        
                        # 연령 범위에 포함되는지 확인
                        if any(age_range in age_str or age_str in age_range for age_range in age_ranges):
                            value = sheet.cell(row=row, column=current_col).value
                            if value is not None:
                                try:
                                    num_value = float(value)
                                    if num_value > 0:
                                        inflow_total += num_value
                                    elif num_value < 0:
                                        outflow_total += abs(num_value)
                                except (ValueError, TypeError):
                                    pass
                    
                    # 다른 지역이 나오면 중단
                    if cell_region and str(cell_region).strip() != mapped_region_name:
                        break
                
                # 유입이 더 큰 그룹을 우선으로 정렬
                net_value = inflow_total - outflow_total
                age_groups.append({
                    'name': group_name,
                    'inflow': inflow_total,
                    'outflow': outflow_total,
                    'net_value': net_value,
                    'current': inflow_total,  # 유입값을 current로 사용
                    'prev': 0,  # 전년 동분기는 별도 계산 필요시 추가
                    'growth_rate': 0  # 증감률은 별도 계산 필요시 추가
                })
            
            # net_value 기준으로 정렬 (유입이 큰 순서)
            age_groups.sort(key=lambda x: x['net_value'], reverse=True)
            
            return age_groups[:top_n]
            
        except Exception as e:
            return []

