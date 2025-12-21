"""
DOCX 생성 모듈
10개 템플릿을 순서대로 처리하여 DOCX로 생성
"""

import os
from pathlib import Path
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from src.template_manager import TemplateManager
from src.excel_extractor import ExcelExtractor
from src.template_filler import TemplateFiller
from src.period_detector import PeriodDetector
from src.flexible_mapper import FlexibleMapper
from src.schema_loader import SchemaLoader

# 10개 템플릿 순서 정의
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


class DOCXGenerator:
    """DOCX 생성 클래스"""
    
    def __init__(self, output_folder):
        """
        DOCX 생성기 초기화
        
        Args:
            output_folder: 출력 폴더 경로
        """
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(parents=True, exist_ok=True)
    
    def check_docx_generator_available(self):
        """
        DOCX 생성 라이브러리 사용 가능 여부 확인
        
        Returns:
            tuple: (사용 가능 여부, 에러 메시지)
        """
        try:
            from docx import Document
            return True, None
        except ImportError:
            return False, (
                'DOCX 생성 라이브러리가 설치되지 않았습니다. 다음을 설치해주세요:\n'
                'pip install python-docx'
            )
    
    def html_to_docx_paragraph(self, doc, html_content):
        """
        HTML 내용을 DOCX 단락으로 변환
        
        Args:
            doc: Document 객체
            html_content: HTML 문자열
        """
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # 스타일 태그 제거 (DOCX에서는 스타일이 다르게 적용됨)
        for style in soup.find_all(['style', 'script']):
            style.decompose()
        
        # body 내용 추출
        body = soup.find('body')
        if not body:
            body = soup
        
        # 각 요소를 순회하며 DOCX로 변환
        for element in body.children:
            if not hasattr(element, 'name'):
                continue
                
            tag_name = element.name.lower()
            
            if tag_name == 'h1':
                doc.add_heading(element.get_text(strip=True), level=1)
            elif tag_name == 'h2':
                doc.add_heading(element.get_text(strip=True), level=2)
            elif tag_name == 'h3':
                doc.add_heading(element.get_text(strip=True), level=3)
            elif tag_name == 'h4':
                doc.add_heading(element.get_text(strip=True), level=4)
            elif tag_name == 'p':
                text = element.get_text(strip=True)
                if text:
                    doc.add_paragraph(text)
            elif tag_name == 'div':
                # div 내부의 클래스 확인
                classes = element.get('class', [])
                if 'content-text' in classes or 'key-section' in classes:
                    text = element.get_text(strip=True)
                    if text:
                        p = doc.add_paragraph(text)
                else:
                    # div 내부 요소를 재귀적으로 처리
                    for child in element.children:
                        if hasattr(child, 'name'):
                            child_tag = child.name.lower()
                            if child_tag == 'table':
                                self.html_table_to_docx(doc, child)
                            elif child_tag in ['h1', 'h2', 'h3', 'h4']:
                                level = int(child_tag[1])
                                doc.add_heading(child.get_text(strip=True), level=level)
                            elif child_tag == 'p':
                                text = child.get_text(strip=True)
                                if text:
                                    doc.add_paragraph(text)
                            else:
                                text = child.get_text(strip=True)
                                if text:
                                    doc.add_paragraph(text)
            elif tag_name == 'table':
                self.html_table_to_docx(doc, element)
            elif tag_name in ['br', 'hr']:
                doc.add_paragraph()
            elif tag_name == 'ul':
                for li in element.find_all('li', recursive=False):
                    text = li.get_text(strip=True)
                    if text:
                        doc.add_paragraph(text, style='List Bullet')
            elif tag_name == 'ol':
                for li in element.find_all('li', recursive=False):
                    text = li.get_text(strip=True)
                    if text:
                        doc.add_paragraph(text, style='List Number')
            else:
                # 기타 요소는 텍스트만 추출
                text = element.get_text(strip=True)
                if text and len(text.strip()) > 0:
                    doc.add_paragraph(text)
    
    def html_table_to_docx(self, doc, table_element):
        """
        HTML 테이블을 DOCX 테이블로 변환
        
        Args:
            doc: Document 객체
            table_element: BeautifulSoup Table 요소
        """
        rows = table_element.find_all('tr')
        if not rows:
            return
        
        # 테이블 생성
        table = doc.add_table(rows=len(rows), cols=0)
        table.style = 'Light Grid Accent 1'
        
        # 첫 번째 행으로 열 수 결정
        first_row = rows[0]
        cols = len(first_row.find_all(['th', 'td']))
        
        # 테이블 재생성 (올바른 열 수로)
        doc.tables[-1]._element.getparent().remove(doc.tables[-1]._element)
        table = doc.add_table(rows=len(rows), cols=cols)
        table.style = 'Light Grid Accent 1'
        
        for row_idx, row in enumerate(rows):
            cells = row.find_all(['th', 'td'])
            for col_idx, cell in enumerate(cells):
                if col_idx < cols:
                    cell_text = cell.get_text(strip=True)
                    table.rows[row_idx].cells[col_idx].text = cell_text
                    
                    # 헤더 셀 스타일 적용
                    if cell.name == 'th':
                        cell_para = table.rows[row_idx].cells[col_idx].paragraphs[0]
                        cell_para.runs[0].bold = True
    
    def generate_docx(
        self,
        excel_path,
        year,
        quarter,
        templates_dir='templates'
    ):
        """
        여러 템플릿을 처리하여 하나의 DOCX 파일로 생성
        
        Args:
            excel_path: 엑셀 파일 경로
            year: 연도
            quarter: 분기
            templates_dir: 템플릿 디렉토리 경로
            
        Returns:
            tuple: (성공 여부, 결과 dict 또는 에러 메시지)
        """
        # DOCX 생성 라이브러리 확인
        is_available, error_msg = self.check_docx_generator_available()
        if not is_available:
            return False, error_msg
        
        try:
            # 파일 경로 검증
            excel_path = Path(excel_path)
            if not excel_path.exists():
                return False, f'엑셀 파일을 찾을 수 없습니다: {excel_path}'
            
            # 엑셀 추출기 초기화
            excel_extractor = ExcelExtractor(str(excel_path))
            excel_extractor.load_workbook()
            
            # 사용 가능한 시트 목록 가져오기
            sheet_names = excel_extractor.get_sheet_names()
            
            # 첫 번째 시트를 기본 시트로 사용 (연도/분기 감지용)
            primary_sheet = sheet_names[0] if sheet_names else None
            if not primary_sheet:
                excel_extractor.close()
                return False, '엑셀 파일에 시트가 없습니다.'
            
            # 연도 및 분기 자동 감지
            period_detector = PeriodDetector(excel_extractor)
            periods_info = period_detector.detect_available_periods(primary_sheet)
            
            # 유효성 검증
            is_valid, error_msg = period_detector.validate_period(primary_sheet, year, quarter)
            if not is_valid:
                excel_extractor.close()
                return False, error_msg
            
            # 유연한 매퍼 초기화
            flexible_mapper = FlexibleMapper(excel_extractor)
            
            # 스키마 로더 초기화
            schema_loader = SchemaLoader()
            
            # DOCX 문서 생성
            doc = Document()
            
            # 문서 제목 추가
            title = doc.add_heading(f'{year}년 {quarter}분기 지역경제동향 보도자료', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 각 템플릿 처리
            processed_count = 0
            errors = []
            templates_dir_path = Path(templates_dir)
            
            for template_name in TEMPLATE_ORDER:
                try:
                    template_path = templates_dir_path / template_name
                    
                    if not template_path.exists():
                        errors.append(f'템플릿 파일을 찾을 수 없습니다: {template_name}')
                        continue
                    
                    # 템플릿 관리자 초기화
                    template_manager = TemplateManager(str(template_path))
                    template_manager.load_template()
                    
                    # 템플릿에서 필요한 시트 목록 추출
                    markers = template_manager.extract_markers()
                    required_sheets = set()
                    for marker in markers:
                        sheet_name = marker.get('sheet_name', '').strip()
                        if sheet_name:
                            required_sheets.add(sheet_name)
                    
                    if not required_sheets:
                        errors.append(f'{template_name}: 필요한 시트를 찾을 수 없습니다.')
                        continue
                    
                    # 템플릿 매핑에서 실제 시트 이름 찾기
                    template_mapping = schema_loader.load_template_mapping()
                    actual_sheet_name = None
                    
                    # template_mapping에서 템플릿 파일명으로 시트 이름 찾기
                    for sheet_name, info in template_mapping.items():
                        if info.get('template') == template_name:
                            actual_sheet_name = sheet_name
                            break
                    
                    # 템플릿 매핑에서 찾지 못한 경우, 마커에서 추출한 첫 번째 시트 이름 사용
                    if not actual_sheet_name:
                        if required_sheets:
                            actual_sheet_name = list(required_sheets)[0]
                        else:
                            errors.append(f'{template_name}: 필요한 시트를 찾을 수 없습니다.')
                            continue
                    
                    # 실제 시트가 존재하는지 확인
                    actual_sheet = flexible_mapper.find_sheet_by_name(actual_sheet_name)
                    if not actual_sheet:
                        errors.append(f'{template_name}: 필요한 시트를 찾을 수 없습니다: {actual_sheet_name}')
                        continue
                    
                    # 템플릿 필러 초기화 및 처리
                    template_filler = TemplateFiller(template_manager, excel_extractor, schema_loader)
                    
                    filled_template = template_filler.fill_template(
                        sheet_name=actual_sheet,
                        year=year,
                        quarter=quarter
                    )
                    
                    # HTML 파싱
                    soup = BeautifulSoup(filled_template, 'html.parser')
                    
                    # 페이지 구분선 추가 (첫 번째 템플릿이 아닌 경우)
                    if processed_count > 0:
                        doc.add_page_break()
                    
                    # 섹션 제목 추가 (템플릿 이름에서 .html 제거)
                    section_title = template_name.replace('.html', '')
                    doc.add_heading(section_title, level=1)
                    
                    # HTML 내용을 DOCX로 변환
                    self.html_to_docx_paragraph(doc, filled_template)
                    
                    processed_count += 1
                    
                except Exception as e:
                    errors.append(f'{template_name}: {str(e)}')
                    continue
            
            # 엑셀 파일 닫기
            excel_extractor.close()
            
            if processed_count == 0:
                return False, f'처리된 템플릿이 없습니다. 오류: {"; ".join(errors)}'
            
            # DOCX 파일 저장
            docx_path = self.output_folder / f"{year}년_{quarter}분기_지역경제동향_보도자료_전체.docx"
            doc.save(str(docx_path))
            
            # 결과 반환
            result = {
                'success': True,
                'output_filename': docx_path.name,
                'output_path': str(docx_path),
                'message': f'{processed_count}개 템플릿이 성공적으로 DOCX로 생성되었습니다.',
                'processed_templates': processed_count,
                'total_templates': len(TEMPLATE_ORDER)
            }
            
            if errors:
                result['warnings'] = errors
            
            return True, result
            
        except Exception as e:
            return False, f'서버 오류가 발생했습니다: {str(e)}'

