"""
동적 마커 처리 총괄 프로세서
여러 핸들러를 조합하여 동적 마커를 처리
"""

from typing import Optional, Dict, Any, List, TYPE_CHECKING

from .base import MarkerContext, MarkerProcessor, MarkerHandler
from .national import NationalMarkerHandler
from .region_ranking import RegionRankingHandler
from .statistics import StatisticsHandler
from .unemployment import UnemploymentHandler
from .chart_data import ChartDataHandler

if TYPE_CHECKING:
    from ..utils.formatters import Formatter
    from ..schema_loader import SchemaLoader
    from ..excel_extractor import ExcelExtractor
    from ..data_analyzer import DataAnalyzer
    from ..dynamic_sheet_parser import DynamicSheetParser


class DynamicMarkerProcessor:
    """동적 마커 처리를 총괄하는 프로세서"""
    
    def __init__(
        self,
        excel_extractor: 'ExcelExtractor',
        schema_loader: 'SchemaLoader',
        data_analyzer: 'DataAnalyzer',
        dynamic_parser: 'DynamicSheetParser',
        formatter: 'Formatter'
    ):
        """
        동적 마커 프로세서 초기화
        
        Args:
            excel_extractor: 엑셀 추출기
            schema_loader: 스키마 로더
            data_analyzer: 데이터 분석기
            dynamic_parser: 동적 시트 파서
            formatter: 포맷터
        """
        self.excel_extractor = excel_extractor
        self.schema_loader = schema_loader
        self.data_analyzer = data_analyzer
        self.dynamic_parser = dynamic_parser
        self.formatter = formatter
        
        # 분석된 데이터 캐시
        self._analyzed_data: Dict[str, Dict[str, Any]] = {}
        
        # 핸들러 체인 설정 (순서 중요: 구체적인 핸들러를 먼저)
        handlers: List[MarkerHandler] = [
            ChartDataHandler(),         # 차트 데이터 (가장 먼저)
            UnemploymentHandler(),      # 실업률/고용률 전용
            NationalMarkerHandler(),    # 전국 관련
            RegionRankingHandler(),     # 지역 순위 관련
            StatisticsHandler(),        # 통계 관련
        ]
        
        self._processor = MarkerProcessor(handlers)
    
    def process(self, sheet_name: str, key: str, year: int, quarter: int) -> str:
        """
        동적 마커를 처리합니다.
        
        Args:
            sheet_name: 시트 이름
            key: 마커 키
            year: 연도
            quarter: 분기
            
        Returns:
            처리된 값 또는 "N/A"
        """
        from ..utils.sheet_utils import get_actual_sheet_name
        
        # 가상 시트명을 실제 시트명으로 변환
        actual_sheet_name = get_actual_sheet_name(sheet_name)
        
        # 데이터 분석 (캐시 사용)
        self._analyze_data_if_needed(actual_sheet_name, year, quarter)
        
        # 컨텍스트 생성
        ctx = MarkerContext(
            sheet_name=actual_sheet_name,
            key=key,
            year=year,
            quarter=quarter,
            excel_extractor=self.excel_extractor,
            schema_loader=self.schema_loader,
            data_analyzer=self.data_analyzer,
            dynamic_parser=self.dynamic_parser,
            formatter=self.formatter,
            analyzed_data=self._analyzed_data
        )
        
        # 핸들러 체인으로 처리
        result = self._processor.process(ctx)
        
        return result if result else "N/A"
    
    def _analyze_data_if_needed(self, sheet_name: str, year: int, quarter: int) -> None:
        """필요시 데이터를 분석하여 캐시에 저장합니다."""
        cache_key = f"{sheet_name}_{year}_{quarter}"
        
        if cache_key in self._analyzed_data:
            return
        
        try:
            # 분기 열 가져오기
            current_col = self.dynamic_parser.get_quarter_column(sheet_name, year, quarter)
            prev_col = self.dynamic_parser.get_quarter_column(sheet_name, year - 1, quarter)
            
            if current_col and prev_col:
                quarter_data = {f"{year}_{quarter}/4": (current_col, prev_col)}
                analyzed = self.data_analyzer.analyze_quarter_data(sheet_name, quarter_data)
                self._analyzed_data[cache_key] = analyzed.get(f"{year}_{quarter}/4", {})
            else:
                self._analyzed_data[cache_key] = {}
        except Exception as e:
            self._analyzed_data[cache_key] = {}
            print(f"[WARNING] 데이터 분석 중 오류 (시트={sheet_name}): {str(e)}")
    
    def clear_cache(self) -> None:
        """분석 데이터 캐시를 초기화합니다."""
        self._analyzed_data.clear()

