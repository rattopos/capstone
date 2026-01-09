"""
마커 처리 모듈
템플릿의 동적 마커를 처리하는 프로세서들
"""

from .base import MarkerProcessor, MarkerContext
from .dynamic_processor import DynamicMarkerProcessor
from .national import NationalMarkerHandler
from .region_ranking import RegionRankingHandler
from .statistics import StatisticsHandler
from .unemployment import UnemploymentHandler
from .chart_data import ChartDataHandler

__all__ = [
    'MarkerProcessor',
    'MarkerContext',
    'DynamicMarkerProcessor',
    'NationalMarkerHandler',
    'RegionRankingHandler',
    'StatisticsHandler',
    'UnemploymentHandler',
    'ChartDataHandler',
]

