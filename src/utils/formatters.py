"""
포맷팅 유틸리티 모듈
숫자, 퍼센트, 증감률 등의 포맷팅 처리
"""

import html
import math
from typing import Any, Optional, Tuple


class Formatter:
    """데이터 포맷팅을 담당하는 클래스"""
    
    def __init__(self, schema_loader=None):
        """
        포맷터 초기화
        
        Args:
            schema_loader: 스키마 로더 인스턴스 (방향 표현 등에 사용)
        """
        self.schema_loader = schema_loader
    
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
            if value is None:
                return "N/A"
            if isinstance(value, str) and not value.strip():
                return "N/A"
            
            num = float(value)
            
            if math.isnan(num) or math.isinf(num):
                return "N/A"
            
            num = round(num, decimal_places)
            
            if decimal_places > 0:
                formatted = f"{num:.{decimal_places}f}"
            else:
                formatted = str(int(num))
            
            if use_comma:
                parts = formatted.split('.')
                try:
                    integer_part = int(float(parts[0]))
                    parts[0] = f"{integer_part:,}"
                except (ValueError, TypeError):
                    pass
                formatted = '.'.join(parts) if len(parts) > 1 else parts[0]
            
            return formatted
        except (ValueError, TypeError, OverflowError):
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
            if value is None:
                return "N/A"
            if isinstance(value, str) and not value.strip():
                return "N/A"
            
            num = float(value)
            
            if math.isnan(num) or math.isinf(num):
                return "N/A"
            
            num = round(num, decimal_places)
            
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
            
            if direction_type == "rise_fall":
                if num > 0:
                    if self.schema_loader:
                        return self.schema_loader.get_direction_expression("rise", expression_key) or "상승"
                    return "상승"
                elif num < 0:
                    if self.schema_loader:
                        return self.schema_loader.get_direction_expression("fall", expression_key) or "하강"
                    return "하강"
                else:
                    return "보합"
            else:  # increase_decrease (기본값)
                if num > 0:
                    if self.schema_loader:
                        return self.schema_loader.get_direction_expression("increase", expression_key) or "증가"
                    return "증가"
                elif num < 0:
                    if self.schema_loader:
                        return self.schema_loader.get_direction_expression("decrease", expression_key) or "감소"
                    return "감소"
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
                    if self.schema_loader:
                        return self.schema_loader.get_direction_expression("rise", "change") or "올라"
                    return "올라"
                elif num < 0:
                    if self.schema_loader:
                        return self.schema_loader.get_direction_expression("fall", "change") or "내려"
                    return "내려"
                else:
                    return "유지되어"
            else:
                if num > 0:
                    if self.schema_loader:
                        return self.schema_loader.get_direction_expression("increase", "production") or "늘어"
                    return "늘어"
                elif num < 0:
                    if self.schema_loader:
                        return self.schema_loader.get_direction_expression("decrease", "production") or "줄어"
                    return "줄어"
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
            
            if math.isnan(num) or math.isinf(num):
                return "N/A"
            
            num = abs(num)
            num = round(num, decimal_places)
            
            formatted = f"{num:.{decimal_places}f}"
            if include_percent:
                return f"{formatted}%"
            return formatted
        except (ValueError, TypeError, OverflowError):
            return "N/A"
    
    def format_growth_for_report(self, value: Any, decimal_places: int = 1, 
                                 direction_type: str = "increase_decrease") -> Tuple[str, str]:
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
        if not self.schema_loader:
            return self.format_percentage(value) if 'percent' in format_type else self.format_number(value)
        
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
            
            num = round(num, decimal_places)
            
            if decimal_places > 0:
                formatted = f"{num:.{decimal_places}f}"
            else:
                formatted = str(int(num))
            
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
    
    def escape_html(self, value: Any) -> str:
        """
        HTML 특수 문자를 이스케이프합니다.
        
        Args:
            value: 이스케이프할 값
            
        Returns:
            이스케이프된 문자열
        """
        return html.escape(str(value) if value is not None else "")


# 모듈 레벨 헬퍼 함수들 (하위 호환성)
def safe_float(value: Any, default: float = None) -> Optional[float]:
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


def calculate_growth_rate(current: Any, prev: Any) -> Optional[float]:
    """증감률을 계산합니다. ((current / prev) - 1) * 100"""
    current_val = safe_float(current)
    prev_val = safe_float(prev)
    
    if current_val is None or prev_val is None or prev_val == 0:
        return None
    
    try:
        growth = ((current_val / prev_val) - 1) * 100
        if math.isnan(growth) or math.isinf(growth):
            return None
        return growth
    except (ZeroDivisionError, OverflowError):
        return None

