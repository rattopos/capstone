"""
PDF 생성 모듈
10개 템플릿을 순서대로 처리하여 PDF로 생성
"""

from pathlib import Path
from .base_generator import BaseDocumentGenerator, TEMPLATE_ORDER

# PDF 생성 라이브러리 선택 (playwright 우선)
PDF_GENERATOR = None
try:
    from playwright.sync_api import sync_playwright
    PDF_GENERATOR = 'playwright'
except ImportError:
    try:
        from weasyprint import HTML
        PDF_GENERATOR = 'weasyprint'
    except ImportError:
        PDF_GENERATOR = None


# PDF 스타일 (상수로 분리하여 재사용)
PDF_STYLES = '''
                @page {
                    size: A4;
                    margin: 0.8cm;
                }
                html, body {
                    width: 100%;
                    height: 100%;
                }
                body {
                    font-family: "Malgun Gothic", "맑은 고딕", sans-serif;
                    margin: 0;
                    padding: 0;
                }
                .page-break {
                    page-break-after: always;
                    min-height: 100vh;
                    display: flex;
                    flex-direction: column;
                    padding: 0;
                    margin: 0;
                }
                .page-break:last-child {
                    page-break-after: auto;
                }
                .page-break .template-content {
                    padding: 10px 15px !important;
                    margin: 0 !important;
                    max-width: 100% !important;
                    width: 100% !important;
                    box-sizing: border-box !important;
                    font-size: 10pt !important;
                    line-height: 1.3 !important;
                    display: flex;
                    flex-direction: column;
                    height: 100%;
                }
.page-break .template-content > * { font-size: inherit !important; }
                .page-break .section-title {
                    font-size: 16pt !important;
                    font-weight: bold !important;
                    margin: 6px 0 4px 0 !important;
                    line-height: 1.2 !important;
                }
                .page-break .subsection-title {
                    font-size: 14pt !important;
                    font-weight: bold !important;
                    margin: 5px 0 3px 0 !important;
                    line-height: 1.2 !important;
                }
                .page-break .content-text {
                    font-size: 12pt !important;
                    margin-bottom: 4px !important;
                    line-height: 1.4 !important;
                    text-align: justify !important;
                }
                .page-break .key-section {
                    font-size: 12pt !important;
                    font-weight: bold !important;
                    margin: 6px 0 3px 0 !important;
                }
                .page-break .key-item {
                    font-size: 10pt !important;
                    margin-bottom: 2px !important;
                    line-height: 1.3 !important;
                }
                .page-break .table-title {
                    font-size: 12pt !important;
                    font-weight: bold !important;
                    margin: 6px 0 2px 0 !important;
                    text-align: center !important;
                }
                .page-break .table-subtitle {
                    font-size: 10pt !important;
                    margin-bottom: 3px !important;
                    text-align: center !important;
                }
                .page-break table {
                    width: 100% !important;
                    max-width: 100% !important;
                    margin: 5px 0 8px 0 !important;
                    font-size: 10pt !important;
                    border-collapse: collapse !important;
                    border: 1px solid #000 !important;
                    table-layout: fixed !important;
                }
                .page-break th {
                    background-color: #f5f5f5 !important;
                    font-weight: bold !important;
                    padding: 3px 2px !important;
                    font-size: 10pt !important;
                    line-height: 1.2 !important;
                    text-align: center !important;
                    border: 1px solid #000 !important;
                    vertical-align: middle !important;
                    word-wrap: break-word !important;
                    overflow: hidden !important;
                }
                .page-break td {
                    padding: 3px 2px !important;
                    font-size: 10pt !important;
                    line-height: 1.2 !important;
                    text-align: center !important;
                    border: 1px solid #000 !important;
                    vertical-align: middle !important;
                    word-wrap: break-word !important;
                    overflow: hidden !important;
                }
                .page-break table th:first-child,
                .page-break table td:first-child {
                    width: 13% !important;
                    min-width: 13% !important;
                    max-width: 13% !important;
                }
                .page-break table th:not(:first-child),
                .page-break table td:not(:first-child) {
                    width: calc(87% / 8) !important;
                    min-width: calc(87% / 8) !important;
                    max-width: calc(87% / 8) !important;
                }
.page-break .source { font-size: 10pt !important; margin-top: 5px !important; }
.page-break .footnote { font-size: 10pt !important; margin-top: 5px !important; }
.page-break ul, .page-break ol { margin: 5px 0 !important; padding-left: 20px !important; }
.page-break li { font-size: 10pt !important; line-height: 1.4 !important; margin-bottom: 2px !important; }
'''


class PDFGenerator(BaseDocumentGenerator):
    """PDF 생성 클래스"""
    
    def check_pdf_generator_available(self):
        """PDF 생성 라이브러리 사용 가능 여부 확인"""
        if not PDF_GENERATOR:
            return False, (
                'PDF 생성 라이브러리가 설치되지 않았습니다. 다음 중 하나를 설치해주세요:\n'
                '1. playwright: pip install playwright && playwright install chromium\n'
                '2. weasyprint: pip install weasyprint'
            )
        return True, None
    
    def _build_combined_html(self, filled_templates):
        """템플릿들을 하나의 HTML로 합치기"""
        html_parts = [
            '<!DOCTYPE html><html><head><meta charset="utf-8"><style>',
            PDF_STYLES,
            '</style></head><body>'
        ]
        
        for template_name, template_html, _ in filled_templates:
            template_class = template_name.replace('.html', '').replace('(', '').replace(')', '').replace(' ', '-')
            html_parts.append(
                f'<div class="page-break template-{template_class}">'
                f'<div class="template-content">{template_html}</div></div>'
            )
        
        html_parts.append('</body></html>')
        return ''.join(html_parts)
    
    def _generate_pdf_playwright(self, combined_html, pdf_path):
        """Playwright를 사용하여 PDF 생성"""
        with sync_playwright() as p:
            browser = p.chromium.launch()
            try:
                page = browser.new_page()
                page.set_content(
                    combined_html,
                    wait_until='domcontentloaded',
                    timeout=60000
                )
                page.wait_for_timeout(500)
                page.pdf(
                    path=str(pdf_path),
                    format='A4',
                    margin={'top': '0.8cm', 'right': '0.8cm', 'bottom': '0.8cm', 'left': '0.8cm'},
                    print_background=True,
                    prefer_css_page_size=True
                )
            finally:
                browser.close()
                        
    def _generate_pdf_weasyprint(self, combined_html, pdf_path, templates_dir_path):
        """WeasyPrint를 사용하여 PDF 생성"""
        from weasyprint import HTML
        HTML(string=combined_html, base_url=str(templates_dir_path.absolute())).write_pdf(str(pdf_path))
    
    def generate_pdf(self, excel_path, year, quarter, templates_dir='templates'):
        """
        여러 템플릿을 처리하여 하나의 PDF 파일로 생성
        
        Args:
            excel_path: 엑셀 파일 경로
            year: 연도
            quarter: 분기
            templates_dir: 템플릿 디렉토리 경로
            
        Returns:
            tuple: (성공 여부, 결과 dict 또는 에러 메시지)
        """
        # PDF 생성 라이브러리 확인
        is_available, error_msg = self.check_pdf_generator_available()
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
            # 템플릿 처리
            filled_templates, errors = self.process_templates(
                excel_extractor, flexible_mapper, year, quarter, templates_dir
            )
            
            if not filled_templates:
                return False, f'처리된 템플릿이 없습니다. 오류: {"; ".join(errors)}'
            
            # HTML 합치기
            combined_html = self._build_combined_html(filled_templates)
            
            # PDF 생성
            pdf_path = self.output_folder / f"{year}년_{quarter}분기_지역경제동향_보도자료_전체.pdf"
            templates_dir_path = Path(templates_dir)
            
            try:
                if PDF_GENERATOR == 'playwright':
                    self._generate_pdf_playwright(combined_html, pdf_path)
                elif PDF_GENERATOR == 'weasyprint':
                    self._generate_pdf_weasyprint(combined_html, pdf_path, templates_dir_path)
                else:
                    return False, 'PDF 생성 라이브러리를 사용할 수 없습니다.'
            except Exception as e:
                return False, f'PDF 생성 중 오류가 발생했습니다: {str(e)}'
            
            # 결과 반환
            result = {
                'success': True,
                'output_filename': pdf_path.name,
                'output_path': str(pdf_path),
                'message': f'{len(filled_templates)}개 템플릿이 성공적으로 PDF로 생성되었습니다.',
                'processed_templates': len(filled_templates),
                'total_templates': len(TEMPLATE_ORDER)
            }
            
            if errors:
                result['warnings'] = errors
            
            return True, result
            
        except Exception as e:
            return False, f'서버 오류가 발생했습니다: {str(e)}'
        finally:
            excel_extractor.close()
