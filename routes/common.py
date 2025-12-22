"""
공통 유틸리티
라우트에서 공유하는 헬퍼 함수 및 상수
"""

import math
import tempfile
from pathlib import Path
from typing import Tuple, Optional

from flask import current_app

# 허용된 파일 확장자
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'html'}
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'}

# 기본 엑셀 파일 경로
BASE_DIR = Path(__file__).parent.parent
DEFAULT_EXCEL_FILE = BASE_DIR / '기초자료 수집표_2025년 2분기_캡스톤.xlsx'


def allowed_file(filename: str) -> bool:
    """파일 확장자 검증"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def allowed_image_file(filename: str) -> bool:
    """이미지 파일 확장자 검증"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_IMAGE_EXTENSIONS


def get_excel_extractor(excel_path: Path):
    """엑셀 추출기를 새로 생성합니다."""
    from src.core.excel_extractor import ExcelExtractor
    
    excel_extractor = ExcelExtractor(str(excel_path.resolve()))
    excel_extractor.load_workbook()
    return excel_extractor


def get_template_manager(template_path: Path):
    """템플릿 매니저를 새로 생성합니다."""
    from src.core.template_manager import TemplateManager
    
    template_manager = TemplateManager(str(template_path.resolve()))
    template_manager.load_template()
    return template_manager


def get_schema_loader():
    """스키마 로더를 반환합니다."""
    from src.core.schema_loader import SchemaLoader
    return SchemaLoader()


def detect_sheet_scale(excel_extractor, sheet_name: str) -> float:
    """시트의 데이터 스케일을 감지합니다."""
    try:
        sheet = excel_extractor.get_sheet(sheet_name)
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
            return 1.0
        
        values.sort()
        median_value = values[len(values) // 2]
        
        if median_value < 10:
            return 1.0
        elif median_value < 100:
            return 10.0
        elif median_value < 1000:
            return 100.0
        else:
            return 1000.0
    except Exception:
        return 1.0


def get_template_for_sheet(sheet_name: str):
    """시트명에 해당하는 템플릿 정보를 반환합니다."""
    schema_loader = get_schema_loader()
    template_info = schema_loader.get_template_for_sheet(sheet_name)
    if template_info:
        return template_info
    
    template_mapping = schema_loader.load_template_mapping()
    return template_mapping.get('광공업생산', {'template': '광공업생산.html', 'display_name': '광공업생산'})

