"""
DOCX 생성 모듈
10개 템플릿을 순서대로 처리하여 DOCX로 생성
"""

from pathlib import Path
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from src.base_generator import BaseDocumentGenerator, TEMPLATE_ORDER


class DOCXGenerator(BaseDocumentGenerator):
    """DOCX 생성 클래스"""
    
    def check_docx_generator_available(self):
        """DOCX 생성 라이브러리 사용 가능 여부 확인"""
        try:
            from docx import Document
            return True, None
        except ImportError:
            return False, (
                'DOCX 생성 라이브러리가 설치되지 않았습니다. 다음을 설치해주세요:\n'
                'pip install python-docx'
            )
    
    def html_to_docx_paragraph(self, doc, html_content):
        """HTML 내용을 DOCX 단락으로 변환"""
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # 스타일 태그 제거
        for style in soup.find_all(['style', 'script']):
            style.decompose()
        
        body = soup.find('body') or soup
        
        for element in body.children:
            self._process_element(doc, element)
    
    def _process_element(self, doc, element):
        """단일 HTML 요소를 DOCX로 변환"""
        if not hasattr(element, 'name') or element.name is None:
            return
        
        tag_name = element.name.lower()
        text = element.get_text(strip=True)
        
        # 헤딩 처리
        if tag_name in ('h1', 'h2', 'h3', 'h4'):
            level = int(tag_name[1])
            if text:
                doc.add_heading(text, level=level)
        # 문단 처리
        elif tag_name == 'p':
            if text:
                doc.add_paragraph(text)
        # div 처리
        elif tag_name == 'div':
            classes = element.get('class', [])
            if 'content-text' in classes or 'key-section' in classes:
                if text:
                    doc.add_paragraph(text)
            else:
                for child in element.children:
                    self._process_element(doc, child)
        # 테이블 처리
        elif tag_name == 'table':
            self._html_table_to_docx(doc, element)
        # 줄바꿈
        elif tag_name in ('br', 'hr'):
            doc.add_paragraph()
        # 리스트 처리
        elif tag_name == 'ul':
            for li in element.find_all('li', recursive=False):
                li_text = li.get_text(strip=True)
                if li_text:
                    doc.add_paragraph(li_text, style='List Bullet')
        elif tag_name == 'ol':
            for li in element.find_all('li', recursive=False):
                li_text = li.get_text(strip=True)
                if li_text:
                    doc.add_paragraph(li_text, style='List Number')
        # 기타 요소
        elif text:
            doc.add_paragraph(text)
    
    def _html_table_to_docx(self, doc, table_element):
        """HTML 테이블을 DOCX 테이블로 변환"""
        rows = table_element.find_all('tr')
        if not rows:
            return
        
        # 첫 번째 행으로 열 수 결정
        first_row_cells = rows[0].find_all(['th', 'td'])
        cols = len(first_row_cells)
        
        if cols == 0:
            return
        
        # 테이블 생성
        table = doc.add_table(rows=len(rows), cols=cols)
        table.style = 'Light Grid Accent 1'
        
        for row_idx, row in enumerate(rows):
            cells = row.find_all(['th', 'td'])
            for col_idx, cell in enumerate(cells):
                if col_idx < cols:
                    cell_text = cell.get_text(strip=True)
                    table.rows[row_idx].cells[col_idx].text = cell_text
                    
                    # 헤더 셀 볼드 처리
                    if cell.name == 'th':
                        para = table.rows[row_idx].cells[col_idx].paragraphs[0]
                        if para.runs:
                            para.runs[0].bold = True
    
    def generate_docx(self, excel_path, year, quarter, templates_dir='templates'):
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
        
        excel_path = Path(excel_path)
        
        # 엑셀 파일 유효성 검증
        is_valid, error_msg = self.validate_excel_file(excel_path)
        if not is_valid:
            return False, error_msg
        
        # 엑셀 추출기 준비
        excel_extractor, flexible_mapper, error_msg = self.prepare_excel_extractor(
            excel_path, year, quarter
        )
        if excel_extractor is None:
            return False, error_msg
        
        try:
            # DOCX 문서 생성
            doc = Document()
            
            # 문서 제목
            title = doc.add_heading(f'{year}년 {quarter}분기 지역경제동향 보도자료', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # 템플릿 처리
            processed_count = 0
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
                    
                    if result is None or isinstance(result, str):
                        errors.append(result or f'{template_name}: 템플릿 처리 실패')
                        continue
                    
                    template_name_out, template_html, _ = result
                    
                    # 페이지 구분선 추가
                    if processed_count > 0:
                        doc.add_page_break()
                    
                    # 섹션 제목
                    section_title = template_name_out.replace('.html', '')
                    doc.add_heading(section_title, level=1)
                    
                    # HTML 내용을 DOCX로 변환
                    self.html_to_docx_paragraph(doc, template_html)
                    
                    processed_count += 1
                    
                except Exception as e:
                    errors.append(f'{template_name}: {str(e)}')
            
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
        finally:
            excel_extractor.close()
