"""
PDF 생성 모듈
10개 템플릿을 순서대로 처리하여 PDF로 생성
"""

import os
from pathlib import Path
from bs4 import BeautifulSoup

from src.template_manager import TemplateManager
from src.excel_extractor import ExcelExtractor
from src.template_filler import TemplateFiller
from src.period_detector import PeriodDetector
from src.flexible_mapper import FlexibleMapper
from src.schema_loader import SchemaLoader

# PDF 생성 라이브러리 선택 (playwright 우선, 없으면 weasyprint)
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


class PDFGenerator:
    """PDF 생성 클래스"""
    
    def __init__(self, output_folder):
        """
        PDF 생성기 초기화
        
        Args:
            output_folder: 출력 폴더 경로
        """
        self.output_folder = Path(output_folder)
        self.output_folder.mkdir(parents=True, exist_ok=True)
    
    def check_pdf_generator_available(self):
        """
        PDF 생성 라이브러리 사용 가능 여부 확인
        
        Returns:
            tuple: (사용 가능 여부, 에러 메시지)
        """
        if not PDF_GENERATOR:
            return False, (
                'PDF 생성 라이브러리가 설치되지 않았습니다. 다음 중 하나를 설치해주세요:\n'
                '1. playwright: pip install playwright && playwright install chromium\n'
                '2. weasyprint: pip install weasyprint (macOS에서는 Homebrew로 시스템 라이브러리 설치 필요)'
            )
        return True, None
    
    def generate_pdf(
        self,
        excel_path,
        year,
        quarter,
        templates_dir='templates'
    ):
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
            
            # 각 템플릿 처리
            filled_templates = []
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
                        # 마커에서 추출한 시트 이름 중 첫 번째 사용
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
                    
                    # 템플릿 채우기 (시트 이름은 실제 엑셀 시트 이름 사용)
                    filled_template = template_filler.fill_template(
                        sheet_name=actual_sheet,
                        year=year,
                        quarter=quarter
                    )
                    
                    # HTML에서 body와 style 내용 추출 (완전한 HTML 문서인 경우)
                    try:
                        soup = BeautifulSoup(filled_template, 'html.parser')
                        body = soup.find('body')
                        style = soup.find('style')
                        
                        template_content = ''
                        # body 내용 추가 (스타일은 PDF 생성 시 전역 스타일로 처리)
                        if body:
                            # body 내용만 추출 (body 태그 제외)
                            body_content = ''.join(str(child) for child in body.children)
                            template_content = body_content
                        else:
                            # body가 없으면 전체 내용 사용
                            template_content = filled_template
                        
                        # 템플릿별 스타일 추출 (나중에 사용할 수 있도록)
                        template_style = ''
                        if style:
                            template_style = style.string or str(style)
                        
                        filled_templates.append((template_name, template_content, template_style))
                    except:
                        # 파싱 실패 시 원본 사용
                        filled_templates.append((template_name, filled_template, ''))
                    
                except Exception as e:
                    errors.append(f'{template_name}: {str(e)}')
                    continue
            
            # 엑셀 파일 닫기
            excel_extractor.close()
            
            if not filled_templates:
                return False, f'처리된 템플릿이 없습니다. 오류: {"; ".join(errors)}'
            
            # 모든 HTML을 하나로 합치기 (페이지 브레이크 추가)
            combined_html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
            combined_html += '''
                @page {
                    size: A4;
                    margin: 1cm;
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
                
                /* 템플릿별 공통 스타일 - 한 페이지에 맞추기 */
                .page-break .template-content {
                    padding: 15px 20px !important;
                    margin: 0 !important;
                    max-width: 100% !important;
                    width: 100% !important;
                    box-sizing: border-box !important;
                    font-size: 11px !important;
                    line-height: 1.4 !important;
                    display: flex;
                    flex-direction: column;
                    height: 100%;
                }
                
                .page-break .template-content > * {
                    font-size: inherit !important;
                }
                
                .page-break .section-title {
                    font-size: 13px !important;
                    font-weight: bold !important;
                    margin: 8px 0 5px 0 !important;
                    line-height: 1.3 !important;
                }
                
                .page-break .subsection-title {
                    font-size: 12px !important;
                    font-weight: bold !important;
                    margin: 6px 0 4px 0 !important;
                    line-height: 1.3 !important;
                }
                
                .page-break .content-text {
                    font-size: 11px !important;
                    margin-bottom: 5px !important;
                    line-height: 1.5 !important;
                    text-align: justify !important;
                }
                
                .page-break .key-section {
                    font-size: 11px !important;
                    font-weight: bold !important;
                    margin: 8px 0 4px 0 !important;
                }
                
                .page-break .key-item {
                    font-size: 10px !important;
                    margin-bottom: 3px !important;
                    line-height: 1.4 !important;
                }
                
                .page-break .table-title {
                    font-size: 11px !important;
                    font-weight: bold !important;
                    margin: 8px 0 3px 0 !important;
                    text-align: center !important;
                }
                
                .page-break .table-subtitle {
                    font-size: 10px !important;
                    margin-bottom: 5px !important;
                    text-align: center !important;
                }
                
                .page-break table {
                    width: 100% !important;
                    margin: 5px 0 8px 0 !important;
                    font-size: 9px !important;
                    border-collapse: collapse !important;
                    border: 1px solid #000 !important;
                    table-layout: fixed !important;
                }
                
                .page-break th {
                    background-color: #f5f5f5 !important;
                    font-weight: bold !important;
                    padding: 4px 3px !important;
                    font-size: 9px !important;
                    line-height: 1.2 !important;
                    text-align: center !important;
                    border: 1px solid #000 !important;
                    vertical-align: middle !important;
                }
                
                .page-break td {
                    padding: 4px 3px !important;
                    font-size: 9px !important;
                    line-height: 1.2 !important;
                    text-align: center !important;
                    border: 1px solid #000 !important;
                    vertical-align: middle !important;
                }
                
                .page-break .source {
                    font-size: 9px !important;
                    margin-top: 5px !important;
                }
                
                /* 리스트 스타일 */
                .page-break ul,
                .page-break ol {
                    margin: 5px 0 !important;
                    padding-left: 20px !important;
                }
                
                .page-break li {
                    font-size: 10px !important;
                    line-height: 1.4 !important;
                    margin-bottom: 2px !important;
                }
            '''
            combined_html += '</style></head><body>'
            
            for i, template_data in enumerate(filled_templates):
                # 템플릿 데이터가 튜플인지 확인
                if isinstance(template_data, tuple):
                    if len(template_data) == 3:
                        template_name, template_html, template_style = template_data
                    else:
                        template_name, template_html = template_data
                        template_style = ''
                else:
                    # 이전 형식 호환성 유지
                    template_name = TEMPLATE_ORDER[i] if i < len(TEMPLATE_ORDER) else f'template_{i}'
                    template_html = template_data
                    template_style = ''
                
                # 템플릿 클래스명 생성
                template_class = template_name.replace('.html', '').replace('(', '').replace(')', '').replace(' ', '-')
                
                # 각 템플릿을 div로 감싸고 페이지 브레이크 및 템플릿 클래스 추가
                combined_html += f'<div class="page-break template-{template_class}"><div class="template-content">{template_html}</div></div>'
            
            combined_html += '</body></html>'
            
            # PDF 생성
            pdf_path = self.output_folder / f"{year}년_{quarter}분기_지역경제동향_보도자료_전체.pdf"
            
            try:
                if PDF_GENERATOR == 'playwright':
                    # playwright를 사용하여 PDF 생성
                    with sync_playwright() as p:
                        browser = p.chromium.launch()
                        page = browser.new_page()
                        
                        # HTML 파일을 임시로 저장하여 로드
                        temp_html_path = self.output_folder / 'temp_combined.html'
                        with open(temp_html_path, 'w', encoding='utf-8') as f:
                            f.write(combined_html)
                        
                        # 파일 경로를 file:// URL로 변환
                        file_url = f"file://{temp_html_path.absolute()}"
                        page.goto(file_url)
                        
                        # PDF 생성
                        page.pdf(
                            path=str(pdf_path),
                            format='A4',
                            margin={'top': '1cm', 'right': '1cm', 'bottom': '1cm', 'left': '1cm'},
                            print_background=True,
                            prefer_css_page_size=True
                        )
                        
                        browser.close()
                        
                        # 임시 HTML 파일 삭제
                        if temp_html_path.exists():
                            temp_html_path.unlink()
                            
                elif PDF_GENERATOR == 'weasyprint':
                    # weasyprint를 사용하여 PDF 생성
                    from weasyprint import HTML
                    HTML(string=combined_html, base_url=str(templates_dir_path.absolute())).write_pdf(str(pdf_path))
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

