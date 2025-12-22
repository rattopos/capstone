"""
공통 유틸리티
라우트에서 공유하는 헬퍼 함수 및 상수
"""

import math
import time
from dataclasses import dataclass, field, asdict
from hashlib import md5
from pathlib import Path
from typing import Tuple, Optional, List, Dict, Any

from flask import current_app, jsonify


# ============================================================
# 검증 관련 데이터 클래스
# ============================================================

@dataclass
class ValidationError:
    """검증 오류 정보"""
    code: str           # 예: "SHEET_NOT_FOUND"
    message: str        # 사용자 친화적 메시지
    detail: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class ValidationResult:
    """검증 결과"""
    success: bool
    errors: List[ValidationError] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    template_name: str = ""
    required_sheets: List[str] = field(default_factory=list)
    found_sheets: Dict[str, str] = field(default_factory=dict)
    missing_sheets: List[str] = field(default_factory=list)
    available_sheets: List[str] = field(default_factory=list)
    periods_info: Optional[Dict[str, Any]] = None
    period_valid: bool = True
    period_error: Optional[str] = None
    marker_count: int = 0
    data_completeness: Optional[Dict[str, Any]] = None
    
    def to_dict(self) -> Dict[str, Any]:
        result = {
            'success': self.success,
            'template_name': self.template_name,
            'required_sheets': self.required_sheets,
            'found_sheets': self.found_sheets,
            'missing_sheets': self.missing_sheets,
            'available_sheets': self.available_sheets,
            'periods_info': self.periods_info,
            'period_valid': self.period_valid,
            'period_error': self.period_error,
            'marker_count': self.marker_count,
        }
        
        if self.errors:
            result['errors'] = [e.to_dict() for e in self.errors]
        if self.warnings:
            result['warnings'] = self.warnings
        if self.data_completeness:
            result['data_completeness'] = self.data_completeness
            
        return result


# ============================================================
# 에러 코드 상수
# ============================================================

class ErrorCodes:
    """검증 에러 코드"""
    TEMPLATE_NOT_SELECTED = "TEMPLATE_NOT_SELECTED"
    TEMPLATE_NOT_FOUND = "TEMPLATE_NOT_FOUND"
    EXCEL_NOT_FOUND = "EXCEL_NOT_FOUND"
    INVALID_FILE_TYPE = "INVALID_FILE_TYPE"
    SHEET_NOT_FOUND = "SHEET_NOT_FOUND"
    PERIOD_NOT_FOUND = "PERIOD_NOT_FOUND"
    DATA_INCOMPLETE = "DATA_INCOMPLETE"
    INVALID_MARKER = "INVALID_MARKER"
    VALIDATION_ERROR = "VALIDATION_ERROR"


# ============================================================
# 검증 결과 캐시
# ============================================================

class ValidationCache:
    """검증 결과 캐싱"""
    _cache: Dict[str, Tuple[float, ValidationResult]] = {}
    TTL = 300  # 5분
    
    @classmethod
    def get_cache_key(cls, excel_path: str, template_name: str, year: int, quarter: int) -> str:
        """캐시 키 생성"""
        return md5(f"{excel_path}:{template_name}:{year}:{quarter}".encode()).hexdigest()
    
    @classmethod
    def get(cls, key: str) -> Optional[ValidationResult]:
        """캐시에서 결과 조회"""
        if key in cls._cache:
            timestamp, result = cls._cache[key]
            if time.time() - timestamp < cls.TTL:
                return result
            else:
                # 만료된 캐시 삭제
                del cls._cache[key]
        return None
    
    @classmethod
    def set(cls, key: str, result: ValidationResult) -> None:
        """캐시에 결과 저장"""
        cls._cache[key] = (time.time(), result)
    
    @classmethod
    def clear(cls) -> None:
        """캐시 초기화"""
        cls._cache.clear()
    
    @classmethod
    def cleanup_expired(cls) -> int:
        """만료된 캐시 정리"""
        now = time.time()
        expired_keys = [
            key for key, (timestamp, _) in cls._cache.items()
            if now - timestamp >= cls.TTL
        ]
        for key in expired_keys:
            del cls._cache[key]
        return len(expired_keys)


# ============================================================
# 에러 응답 헬퍼 함수
# ============================================================

def error_response(error: ValidationError, status_code: int = 400):
    """표준화된 에러 응답 생성"""
    return jsonify({
        'success': False,
        'error': error.to_dict()
    }), status_code


def validation_error_response(errors: List[ValidationError], status_code: int = 400):
    """여러 에러에 대한 응답 생성"""
    return jsonify({
        'success': False,
        'errors': [e.to_dict() for e in errors]
    }), status_code

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

