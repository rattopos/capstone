"""
마커 처리 기본 클래스
마커 처리를 위한 컨텍스트와 기본 핸들러 인터페이스 정의
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import Any, Dict, Optional, List, TYPE_CHECKING

if TYPE_CHECKING:
    from ..utils.formatters import Formatter
    from ..schema_loader import SchemaLoader
    from ..excel_extractor import ExcelExtractor
    from ..data_analyzer import DataAnalyzer
    from ..dynamic_sheet_parser import DynamicSheetParser


@dataclass
class MarkerContext:
    """마커 처리에 필요한 컨텍스트 정보를 담는 데이터 클래스"""
    sheet_name: str
    key: str
    year: int
    quarter: int
    excel_extractor: 'ExcelExtractor'
    schema_loader: 'SchemaLoader'
    data_analyzer: 'DataAnalyzer'
    dynamic_parser: 'DynamicSheetParser'
    formatter: 'Formatter'
    analyzed_data: Dict[str, Any] = None
    
    def get_cache_key(self) -> str:
        """캐시 키를 반환합니다."""
        return f"{self.sheet_name}_{self.year}_{self.quarter}"
    
    def get_data(self) -> Dict[str, Any]:
        """분석된 데이터를 반환합니다."""
        if self.analyzed_data is None:
            return {}
        cache_key = self.get_cache_key()
        return self.analyzed_data.get(cache_key, {})


class MarkerHandler(ABC):
    """마커 핸들러 기본 클래스 (Chain of Responsibility 패턴)"""
    
    def __init__(self):
        self._next_handler: Optional['MarkerHandler'] = None
    
    def set_next(self, handler: 'MarkerHandler') -> 'MarkerHandler':
        """다음 핸들러를 설정합니다."""
        self._next_handler = handler
        return handler
    
    @abstractmethod
    def can_handle(self, ctx: MarkerContext) -> bool:
        """이 핸들러가 해당 마커를 처리할 수 있는지 확인합니다."""
        pass
    
    @abstractmethod
    def handle(self, ctx: MarkerContext) -> Optional[str]:
        """마커를 처리하고 결과를 반환합니다."""
        pass
    
    def process(self, ctx: MarkerContext) -> Optional[str]:
        """
        마커를 처리합니다. 처리할 수 없으면 다음 핸들러에 위임합니다.
        
        Args:
            ctx: 마커 컨텍스트
            
        Returns:
            처리된 값 또는 None
        """
        if self.can_handle(ctx):
            result = self.handle(ctx)
            if result is not None:
                return result
        
        if self._next_handler:
            return self._next_handler.process(ctx)
        
        return None


class MarkerProcessor:
    """마커 처리 총괄 클래스"""
    
    def __init__(self, handlers: List[MarkerHandler] = None):
        """
        마커 프로세서 초기화
        
        Args:
            handlers: 마커 핸들러 리스트 (체인으로 연결됨)
        """
        self._first_handler: Optional[MarkerHandler] = None
        
        if handlers:
            self._setup_chain(handlers)
    
    def _setup_chain(self, handlers: List[MarkerHandler]) -> None:
        """핸들러 체인을 설정합니다."""
        if not handlers:
            return
        
        self._first_handler = handlers[0]
        current = self._first_handler
        
        for handler in handlers[1:]:
            current.set_next(handler)
            current = handler
    
    def add_handler(self, handler: MarkerHandler) -> None:
        """핸들러를 체인 끝에 추가합니다."""
        if not self._first_handler:
            self._first_handler = handler
            return
        
        current = self._first_handler
        while current._next_handler:
            current = current._next_handler
        current.set_next(handler)
    
    def process(self, ctx: MarkerContext) -> str:
        """
        마커를 처리합니다.
        
        Args:
            ctx: 마커 컨텍스트
            
        Returns:
            처리된 값 또는 "N/A"
        """
        if self._first_handler:
            result = self._first_handler.process(ctx)
            if result is not None:
                return result
        
        return "N/A"

