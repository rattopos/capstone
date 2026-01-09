"""
템플릿 매칭 모듈
저장된 템플릿을 선택하고 해당 매핑 로직을 적용
"""

from typing import Dict, Any, Optional
from .template_storage import TemplateStorage
from .template_manager import TemplateManager
from .template_filler import TemplateFiller
from .excel_extractor import ExcelExtractor


class TemplateMatcher:
    """저장된 템플릿을 매칭하고 매핑 로직을 적용하는 클래스"""
    
    def __init__(self, template_storage: Optional[TemplateStorage] = None):
        """
        템플릿 매처 초기화
        
        Args:
            template_storage: 템플릿 저장소 인스턴스
        """
        self.template_storage = template_storage or TemplateStorage()
    
    def match_and_fill(
        self,
        template_name: str,
        excel_extractor: ExcelExtractor,
        year: int,
        quarter: int
    ) -> Optional[str]:
        """
        저장된 템플릿을 찾아서 데이터를 채웁니다.
        
        Args:
            template_name: 템플릿 이름
            excel_extractor: 엑셀 추출기 인스턴스
            year: 연도
            quarter: 분기
            
        Returns:
            완성된 HTML 템플릿 또는 None
        """
        # 템플릿 로드
        template_data = self.template_storage.load_template(template_name)
        if not template_data:
            return None
        
        template_html = template_data['template_html']
        mapping_data = template_data['mapping_data']
        
        # 템플릿 관리자 초기화
        template_manager = TemplateManager(template_path=None)
        template_manager.template_content = template_html
        
        # 매핑 로직 적용
        filled_template = self._apply_mapping_logic(
            template_html,
            mapping_data,
            excel_extractor,
            year,
            quarter
        )
        
        return filled_template
    
    def _apply_mapping_logic(
        self,
        template_html: str,
        mapping_data: Dict[str, Any],
        excel_extractor: ExcelExtractor,
        year: int,
        quarter: int
    ) -> str:
        """매핑 로직을 적용하여 템플릿 채우기"""
        from .template_filler import TemplateFiller
        
        # 템플릿 필러 초기화
        template_manager = TemplateManager(template_path=None)
        template_manager.template_content = template_html
        
        template_filler = TemplateFiller(template_manager, excel_extractor)
        
        # 매핑 데이터에서 시트명 추출
        sheet_name = mapping_data.get('sheet_name', '')
        if not sheet_name:
            # 매핑에서 첫 번째 시트명 찾기
            mappings = mapping_data.get('mappings', {})
            for mapping in mappings.values():
                if isinstance(mapping, dict) and 'sheet_name' in mapping:
                    sheet_name = mapping['sheet_name']
                    break
        
        # 템플릿 채우기
        filled_template = template_filler.fill_template(
            sheet_name=sheet_name,
            year=year,
            quarter=quarter
        )
        
        return filled_template

