"""
소스 모듈
지역경제동향 보도자료 자동생성 시스템

모듈 구조:
- core/: 핵심 모듈 (TemplateFiller, TemplateManager, ExcelExtractor, SchemaLoader)
- markers/: 마커 처리 모듈 (DynamicMarkerProcessor, 핸들러들)
- utils/: 유틸리티 모듈 (Formatter, region_utils, sheet_utils)
- analyzers/: 분석 모듈 (DataAnalyzer, PeriodDetector, DynamicSheetParser)
- generators/: 생성 모듈 (PDFGenerator, DOCXGenerator, TemplateGenerator)

하위 호환성을 위해 기존 모듈도 유지됩니다.
"""

# 하위 호환성을 위한 re-export
# 기존 코드에서 from src.template_filler import TemplateFiller 형식으로 사용 가능

# Core modules
from .template_manager import TemplateManager
from .excel_extractor import ExcelExtractor
from .schema_loader import SchemaLoader
from .template_filler import TemplateFiller

# Analyzers
from .data_analyzer import DataAnalyzer
from .period_detector import PeriodDetector
from .dynamic_sheet_parser import DynamicSheetParser
from .excel_header_parser import ExcelHeaderParser

# Generators
from .base_generator import BaseDocumentGenerator
from .pdf_generator import PDFGenerator
from .docx_generator import DOCXGenerator
from .template_generator import TemplateGenerator

# Utils
from .calculator import Calculator
from .flexible_mapper import FlexibleMapper

__all__ = [
    # Core
    'TemplateFiller',
    'TemplateManager',
    'ExcelExtractor',
    'SchemaLoader',
    # Analyzers
    'DataAnalyzer',
    'PeriodDetector',
    'DynamicSheetParser',
    'ExcelHeaderParser',
    # Generators
    'BaseDocumentGenerator',
    'PDFGenerator',
    'DOCXGenerator',
    'TemplateGenerator',
    # Utils
    'Calculator',
    'FlexibleMapper',
]
