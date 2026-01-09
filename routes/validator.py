"""
템플릿 검증 클래스
TemplateValidator - 템플릿과 엑셀 데이터 검증을 담당
"""

from pathlib import Path
from typing import List, Dict, Any, Optional, Set

from .common import (
    ValidationError, ValidationResult, ErrorCodes,
    get_excel_extractor, get_template_manager
)


# 17개 시도 목록
REGIONS = [
    '전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
    '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주'
]


class TemplateValidator:
    """템플릿 검증을 담당하는 클래스"""
    
    def __init__(self, excel_path: Path, template_path: Path):
        self.excel_path = excel_path
        self.template_path = template_path
        self.errors: List[ValidationError] = []
        self.warnings: List[str] = []
        
        # Lazy loading을 위한 캐시
        self._excel_extractor = None
        self._template_manager = None
        self._markers = None
        self._flexible_mapper = None
    
    @property
    def excel_extractor(self):
        """엑셀 추출기 (lazy loading)"""
        if self._excel_extractor is None:
            self._excel_extractor = get_excel_extractor(self.excel_path)
        return self._excel_extractor
    
    @property
    def template_manager(self):
        """템플릿 매니저 (lazy loading)"""
        if self._template_manager is None:
            self._template_manager = get_template_manager(self.template_path)
        return self._template_manager
    
    @property
    def markers(self) -> List[Dict]:
        """마커 목록 (lazy loading)"""
        if self._markers is None:
            self._markers = self.template_manager.extract_markers()
        return self._markers
    
    @property
    def flexible_mapper(self):
        """유연한 매퍼 (lazy loading)"""
        if self._flexible_mapper is None:
            from src.flexible_mapper import FlexibleMapper
            self._flexible_mapper = FlexibleMapper(self.excel_extractor)
        return self._flexible_mapper
    
    def close(self):
        """리소스 정리"""
        if self._excel_extractor is not None:
            self._excel_extractor.close()
            self._excel_extractor = None
    
    def validate_all(self, year: Optional[int] = None, quarter: Optional[int] = None) -> ValidationResult:
        """
        전체 검증 수행
        
        Args:
            year: 연도 (선택)
            quarter: 분기 (선택)
            
        Returns:
            ValidationResult: 검증 결과
        """
        try:
            # 1. 파일 존재 검증
            file_result = self._validate_files()
            if not file_result['success']:
                return ValidationResult(
                    success=False,
                    errors=self.errors,
                    template_name=self.template_path.name
                )
            
            # 2. 시트 검증
            sheet_result = self._validate_sheets()
            
            # 3. 기간 검증 (연도/분기가 주어진 경우)
            period_result = self._validate_periods(year, quarter) if year and quarter else {
                'periods_info': None,
                'period_valid': True,
                'period_error': None
            }
            
            # 4. 마커 구문 검증
            marker_errors = self._validate_marker_syntax()
            if marker_errors:
                self.warnings.extend([f"마커 경고: {e.message}" for e in marker_errors])
            
            # 5. 데이터 완전성 검증 (시트가 있고 기간이 유효한 경우)
            data_completeness = None
            if sheet_result['found_sheets'] and period_result['period_valid'] and year and quarter:
                primary_sheet = list(sheet_result['found_sheets'].values())[0]
                data_completeness = self._validate_data_completeness(primary_sheet, year, quarter)
            
            # 결과 생성
            success = len(self.errors) == 0 and len(sheet_result['missing_sheets']) == 0
            
            return ValidationResult(
                success=success,
                errors=self.errors,
                warnings=self.warnings,
                template_name=self.template_path.name,
                required_sheets=sheet_result['required_sheets'],
                found_sheets=sheet_result['found_sheets'],
                missing_sheets=sheet_result['missing_sheets'],
                available_sheets=sheet_result['available_sheets'],
                periods_info=period_result.get('periods_info'),
                period_valid=period_result.get('period_valid', True),
                period_error=period_result.get('period_error'),
                marker_count=len(self.markers),
                data_completeness=data_completeness
            )
            
        except Exception as e:
            self.errors.append(ValidationError(
                code=ErrorCodes.VALIDATION_ERROR,
                message=f"검증 중 오류가 발생했습니다: {str(e)}",
                detail=str(e)
            ))
            return ValidationResult(
                success=False,
                errors=self.errors,
                template_name=self.template_path.name
            )
        finally:
            self.close()
    
    def _validate_files(self) -> Dict[str, Any]:
        """파일 존재 여부 검증"""
        result = {'success': True}
        
        if not self.excel_path.exists():
            self.errors.append(ValidationError(
                code=ErrorCodes.EXCEL_NOT_FOUND,
                message="엑셀 파일을 찾을 수 없습니다.",
                detail=str(self.excel_path)
            ))
            result['success'] = False
        
        if not self.template_path.exists():
            self.errors.append(ValidationError(
                code=ErrorCodes.TEMPLATE_NOT_FOUND,
                message="템플릿 파일을 찾을 수 없습니다.",
                detail=str(self.template_path)
            ))
            result['success'] = False
        
        return result
    
    def _validate_sheets(self) -> Dict[str, Any]:
        """시트 존재 여부 검증"""
        # 템플릿에서 필요한 시트 추출
        required_sheets: Set[str] = set()
        for marker in self.markers:
            sheet_name = marker.get('sheet_name', '').strip()
            if sheet_name:
                required_sheets.add(sheet_name)
        
        # 엑셀에서 사용 가능한 시트
        available_sheets = self.excel_extractor.get_sheet_names()
        
        # 시트 매핑
        found_sheets: Dict[str, str] = {}
        missing_sheets: List[str] = []
        
        for required_sheet in required_sheets:
            actual_sheet = self.flexible_mapper.find_sheet_by_name(required_sheet)
            if actual_sheet:
                found_sheets[required_sheet] = actual_sheet
            else:
                missing_sheets.append(required_sheet)
        
        # 누락된 시트가 있으면 경고 추가
        if missing_sheets:
            self.warnings.append(f"누락된 시트: {', '.join(missing_sheets)}")
        
        return {
            'required_sheets': list(required_sheets),
            'found_sheets': found_sheets,
            'missing_sheets': missing_sheets,
            'available_sheets': available_sheets
        }
    
    def _validate_periods(self, year: int, quarter: int) -> Dict[str, Any]:
        """연도/분기 데이터 검증"""
        from src.analyzers.period_detector import PeriodDetector
        
        result = {
            'periods_info': None,
            'period_valid': True,
            'period_error': None
        }
        
        try:
            sheet_names = self.excel_extractor.get_sheet_names()
            if not sheet_names:
                return result
            
            period_detector = PeriodDetector(self.excel_extractor)
            
            # 첫 번째 시트에서 사용 가능한 기간 감지
            result['periods_info'] = period_detector.detect_available_periods(sheet_names[0])
            
            # 기간 유효성 검증
            is_valid, error_msg = period_detector.validate_period(sheet_names[0], year, quarter)
            result['period_valid'] = is_valid
            
            if not is_valid:
                result['period_error'] = error_msg
                self.errors.append(ValidationError(
                    code=ErrorCodes.PERIOD_NOT_FOUND,
                    message=error_msg or f"{year}년 {quarter}분기 데이터를 찾을 수 없습니다.",
                    detail=f"year={year}, quarter={quarter}"
                ))
                
        except Exception as e:
            result['period_error'] = str(e)
            self.warnings.append(f"기간 검증 중 오류: {str(e)}")
        
        return result
    
    def _validate_marker_syntax(self) -> List[ValidationError]:
        """마커 구문 검증"""
        errors: List[ValidationError] = []
        
        for marker in self.markers:
            full_match = marker.get('full_match', '')
            
            # 시트명 누락 검사
            if not marker.get('sheet_name'):
                errors.append(ValidationError(
                    code=ErrorCodes.INVALID_MARKER,
                    message=f"시트명이 누락된 마커가 있습니다.",
                    detail=full_match
                ))
            
            # 필드명 누락 검사
            if not marker.get('field_name') and not marker.get('region'):
                errors.append(ValidationError(
                    code=ErrorCodes.INVALID_MARKER,
                    message=f"필드명 또는 지역이 누락된 마커가 있습니다.",
                    detail=full_match
                ))
        
        return errors
    
    def _validate_data_completeness(self, sheet_name: str, year: int, quarter: int) -> Dict[str, Any]:
        """데이터 완전성 검증"""
        result = {
            'complete': True,
            'missing_regions': [],
            'has_national_data': True,
            'coverage_percent': 100.0
        }
        
        try:
            # 전국 데이터 확인
            has_national = self._has_region_data(sheet_name, '전국', year, quarter)
            result['has_national_data'] = has_national
            
            if not has_national:
                result['missing_regions'].append('전국')
                self.warnings.append("전국 데이터가 누락되었습니다.")
            
            # 17개 시도 데이터 확인
            missing_regions = []
            for region in REGIONS[1:]:  # 전국 제외
                if not self._has_region_data(sheet_name, region, year, quarter):
                    missing_regions.append(region)
            
            if missing_regions:
                result['missing_regions'].extend(missing_regions)
                result['complete'] = False
                self.warnings.append(f"일부 지역 데이터 누락: {', '.join(missing_regions[:5])}" + 
                                    (f" 외 {len(missing_regions)-5}개" if len(missing_regions) > 5 else ""))
            
            # 커버리지 계산
            total_regions = len(REGIONS)
            found_regions = total_regions - len(result['missing_regions'])
            result['coverage_percent'] = round(found_regions / total_regions * 100, 1)
            
        except Exception as e:
            self.warnings.append(f"데이터 완전성 검증 중 오류: {str(e)}")
            result['error'] = str(e)
        
        return result
    
    def _has_region_data(self, sheet_name: str, region: str, year: int, quarter: int) -> bool:
        """특정 지역의 데이터 존재 여부 확인"""
        try:
            sheet = self.excel_extractor.get_sheet(sheet_name)
            if sheet is None:
                return False
            
            # 지역명을 찾아서 해당 행의 데이터 확인
            for row in range(1, min(sheet.max_row + 1, 50)):
                cell_value = sheet.cell(row=row, column=1).value
                if cell_value and region in str(cell_value):
                    # 해당 행에 데이터가 있는지 확인
                    for col in range(2, min(sheet.max_column + 1, 20)):
                        value = sheet.cell(row=row, column=col).value
                        if value is not None:
                            return True
            return False
        except Exception:
            return False

