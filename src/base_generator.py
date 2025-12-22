"""
문서 생성 베이스 모듈
PDF, DOCX 등 문서 생성기의 공통 로직을 제공
"""

from pathlib import Path
from typing import List, Tuple, Dict, Any, Optional
from functools import lru_cache

from src.template_manager import TemplateManager
from src.excel_extractor import ExcelExtractor
from src.template_filler import TemplateFiller
from src.period_detector import PeriodDetector
from src.flexible_mapper import FlexibleMapper
from src.schema_loader import SchemaLoader


# 10개 템플릿 순서 정의 (공통)
TEMPLATE_ORDER = [
    '광공업생산.html',
    '서비스업생산.html',
    '소매판매.html',
    '건설수주.html',
    '수출.html',
    '수입.html',
    '고용률.html',
    '실업률.html',
    '물가동향.html',
    '국내인구이동.html'
]


class BaseDocumentGenerator:
    """문서 생성 베이스 클래스"""
    
    def __init__(self, output_folder: str):
        """
        문서 생성기 초기화
        
        Args:
            output_folder: 출력 폴더 경로
        """
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(parents=True, exist_ok=True)
        self._schema_loader = None
        self._template_mapping_cache = None
    
    @property
    def schema_loader(self) -> SchemaLoader:
        """스키마 로더 (지연 로딩 및 캐싱)"""
        if self._schema_loader is None:
            self._schema_loader = SchemaLoader()
        return self._schema_loader
    
    @property
    def template_mapping(self) -> Dict:
        """템플릿 매핑 (캐싱)"""
        if self._template_mapping_cache is None:
            self._template_mapping_cache = self.schema_loader.load_template_mapping()
        return self._template_mapping_cache
    
    def validate_excel_file(self, excel_path: Path) -> Tuple[bool, str]:
        """
        엑셀 파일 유효성 검증
        
        Args:
            excel_path: 엑셀 파일 경로
            
        Returns:
            (성공 여부, 에러 메시지)
        """
        if not excel_path.exists():
            return False, f'엑셀 파일을 찾을 수 없습니다: {excel_path}'
        return True, ''
    
    def prepare_excel_extractor(
        self,
        excel_path: Path,
        year: int,
        quarter: int
    ) -> Tuple[Optional[ExcelExtractor], Optional[FlexibleMapper], str]:
        """
        엑셀 추출기 및 관련 객체 초기화
        
        Args:
            excel_path: 엑셀 파일 경로
            year: 연도
            quarter: 분기
            
        Returns:
            (ExcelExtractor, FlexibleMapper, 에러 메시지)
        """
        try:
            excel_extractor = ExcelExtractor(str(excel_path))
            excel_extractor.load_workbook()
            
            sheet_names = excel_extractor.get_sheet_names()
            primary_sheet = sheet_names[0] if sheet_names else None
            
            if not primary_sheet:
                excel_extractor.close()
                return None, None, '엑셀 파일에 시트가 없습니다.'
            
            # 연도 및 분기 유효성 검증
            period_detector = PeriodDetector(excel_extractor)
            is_valid, error_msg = period_detector.validate_period(primary_sheet, year, quarter)
            
            if not is_valid:
                excel_extractor.close()
                return None, None, error_msg
            
            flexible_mapper = FlexibleMapper(excel_extractor)
            
            return excel_extractor, flexible_mapper, ''
            
        except Exception as e:
            return None, None, f'엑셀 파일 처리 중 오류: {str(e)}'
    
    def find_actual_sheet_name(
        self,
        template_name: str,
        required_sheets: set,
        flexible_mapper: FlexibleMapper
    ) -> Optional[str]:
        """
        템플릿에 필요한 실제 시트 이름 찾기
        
        Args:
            template_name: 템플릿 파일명
            required_sheets: 필요한 시트 이름 집합
            flexible_mapper: FlexibleMapper 객체
            
        Returns:
            실제 시트 이름 또는 None
        """
        # template_mapping에서 템플릿 파일명으로 시트 이름 찾기
        for sheet_name, info in self.template_mapping.items():
            if info.get('template') == template_name:
                actual_sheet = flexible_mapper.find_sheet_by_name(sheet_name)
                if actual_sheet:
                    return actual_sheet
        
        # 템플릿 매핑에서 찾지 못한 경우, 마커에서 추출한 시트 이름 사용
        for sheet_name in required_sheets:
            actual_sheet = flexible_mapper.find_sheet_by_name(sheet_name)
            if actual_sheet:
                return actual_sheet
        
        return None
    
    def process_templates(
        self,
        excel_extractor: ExcelExtractor,
        flexible_mapper: FlexibleMapper,
        year: int,
        quarter: int,
        templates_dir: str = 'templates'
    ) -> Tuple[List[Tuple[str, str, str]], List[str]]:
        """
        모든 템플릿 처리
        
        Args:
            excel_extractor: ExcelExtractor 객체
            flexible_mapper: FlexibleMapper 객체
            year: 연도
            quarter: 분기
            templates_dir: 템플릿 디렉토리 경로
            
        Returns:
            (처리된 템플릿 리스트, 에러 리스트)
        """
        filled_templates = []
        errors = []
        templates_dir_path = Path(templates_dir)
        
        for template_name in TEMPLATE_ORDER:
            try:
                result = self._process_single_template(
                    template_name=template_name,
                    templates_dir_path=templates_dir_path,
                    excel_extractor=excel_extractor,
                    flexible_mapper=flexible_mapper,
                    year=year,
                    quarter=quarter
                )
                
                if result is None:
                    errors.append(f'{template_name}: 템플릿 처리 실패')
                elif isinstance(result, str):
                    errors.append(result)
                else:
                    filled_templates.append(result)
                    
            except Exception as e:
                errors.append(f'{template_name}: {str(e)}')
        
        return filled_templates, errors
    
    def _process_single_template(
        self,
        template_name: str,
        templates_dir_path: Path,
        excel_extractor: ExcelExtractor,
        flexible_mapper: FlexibleMapper,
        year: int,
        quarter: int
    ) -> Optional[Tuple[str, str, str]]:
        """
        단일 템플릿 처리
        
        Args:
            template_name: 템플릿 파일명
            templates_dir_path: 템플릿 디렉토리 경로
            excel_extractor: ExcelExtractor 객체
            flexible_mapper: FlexibleMapper 객체
            year: 연도
            quarter: 분기
            
        Returns:
            (템플릿명, HTML 내용, 스타일) 또는 에러 문자열 또는 None
        """
        template_path = templates_dir_path / template_name
        
        if not template_path.exists():
            return f'{template_name}: 템플릿 파일을 찾을 수 없습니다'
        
        # 템플릿 관리자 초기화
        template_manager = TemplateManager(str(template_path))
        template_manager.load_template()
        
        # 템플릿에서 필요한 시트 목록 추출
        markers = template_manager.extract_markers()
        required_sheets = {
            marker.get('sheet_name', '').strip()
            for marker in markers
            if marker.get('sheet_name', '').strip()
        }
        
        if not required_sheets:
            return f'{template_name}: 필요한 시트를 찾을 수 없습니다'
        
        # 실제 시트 이름 찾기
        actual_sheet = self.find_actual_sheet_name(
            template_name, required_sheets, flexible_mapper
        )
        
        if not actual_sheet:
            return f'{template_name}: 필요한 시트를 찾을 수 없습니다'
        
        # 템플릿 필러 초기화 및 처리
        template_filler = TemplateFiller(
            template_manager, excel_extractor, self.schema_loader
        )
        
        filled_template = template_filler.fill_template(
            sheet_name=actual_sheet,
            year=year,
            quarter=quarter
        )
        
        # HTML에서 body와 style 추출
        return self._extract_template_content(template_name, filled_template)
    
    def _extract_template_content(
        self,
        template_name: str,
        filled_template: str
    ) -> Tuple[str, str, str]:
        """
        HTML에서 body와 style 내용 추출
        
        Args:
            template_name: 템플릿 이름
            filled_template: 채워진 HTML 템플릿
            
        Returns:
            (템플릿명, body 내용, style 내용)
        """
        from bs4 import BeautifulSoup
        
        try:
            soup = BeautifulSoup(filled_template, 'html.parser')
            body = soup.find('body')
            style = soup.find('style')
            
            # body 내용 추출
            if body:
                try:
                    template_content = body.decode_contents()
                except:
                    template_content = ''.join(str(child) for child in body.children)
            else:
                template_content = filled_template
            
            # style 내용 추출
            template_style = style.string if style and style.string else ''
            
            return (template_name, template_content, template_style)
            
        except Exception:
            return (template_name, filled_template, '')

