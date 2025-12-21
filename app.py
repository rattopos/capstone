"""
Flask 웹 애플리케이션
지역경제동향 보도자료 자동생성 웹 인터페이스
"""

import os
import sys
import json
from pathlib import Path
from flask import Flask, request, render_template, jsonify, send_file, send_from_directory

# Flask 템플릿 폴더 설정
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from werkzeug.utils import secure_filename
import tempfile
import shutil

from src.template_manager import TemplateManager
from src.excel_extractor import ExcelExtractor
from src.template_filler import TemplateFiller
from src.period_detector import PeriodDetector
from src.template_generator import TemplateGenerator
from src.excel_header_parser import ExcelHeaderParser
from src.flexible_mapper import FlexibleMapper
from src.pdf_generator import PDFGenerator
from src.schema_loader import SchemaLoader
from bs4 import BeautifulSoup

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

app = Flask(__name__, template_folder='flask_templates', static_folder='static')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['OUTPUT_FOLDER'] = tempfile.mkdtemp()

# 기본 엑셀 파일 경로
BASE_DIR = Path(__file__).parent
DEFAULT_EXCEL_FILE = BASE_DIR / '기초자료 수집표_2025년 2분기_캡스톤.xlsx'

# 허용된 파일 확장자
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'html'}
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'}

# 스키마 로더 초기화
schema_loader = SchemaLoader()


def allowed_file(filename):
    """파일 확장자 검증"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def allowed_image_file(filename):
    """이미지 파일 확장자 검증"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_IMAGE_EXTENSIONS


def get_template_for_sheet(sheet_name):
    """
    시트명에 해당하는 템플릿 정보를 반환합니다.
    
    Args:
        sheet_name: 엑셀 시트명
        
    Returns:
        dict: {'template': 템플릿 파일명, 'display_name': 표시용 이름}
        매핑이 없으면 기본값 반환
    """
    template_info = schema_loader.get_template_for_sheet(sheet_name)
    if template_info:
        return template_info
    
    # 기본값: 광공업생산 템플릿 사용
    template_mapping = schema_loader.load_template_mapping()
    return template_mapping.get('광공업생산', {'template': '광공업생산.html', 'display_name': '광공업생산'})


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')


@app.route('/api/templates', methods=['GET'])
def get_templates():
    """사용 가능한 템플릿 목록 반환 (각 템플릿이 필요한 시트 정보 포함)"""
    templates_dir = Path('templates')
    templates = []
    
    # templates 디렉토리의 모든 HTML 파일을 스캔
    if templates_dir.exists():
        # 모든 HTML 파일 찾기
        html_files = list(templates_dir.glob('*.html'))
        
        for template_path in html_files:
            template_name = template_path.name
            
            # 템플릿에서 필요한 시트 목록 추출
            try:
                template_manager = TemplateManager(str(template_path))
                template_manager.load_template()
                markers = template_manager.extract_markers()
                
                # 마커에서 시트명 추출 (중복 제거)
                required_sheets = set()
                for marker in markers:
                    sheet_name = marker.get('sheet_name', '').strip()
                    if sheet_name:
                        required_sheets.add(sheet_name)
                
                # display_name 찾기 (템플릿 매핑에서 먼저 찾고, 없으면 파일명 사용)
                display_name = template_name.replace('.html', '')
                template_mapping = schema_loader.load_template_mapping()
                for sheet_name, info in template_mapping.items():
                    if info['template'] == template_name:
                        display_name = info['display_name']
                        break
                
                templates.append({
                    'name': template_name,
                    'path': str(template_path),
                    'display_name': display_name,
                    'required_sheets': list(required_sheets)
                })
            except Exception as e:
                # 템플릿 파싱 실패 시 기본 정보만 반환
                display_name = template_name.replace('.html', '')
                template_mapping = schema_loader.load_template_mapping()
                for sheet_name, info in template_mapping.items():
                    if info['template'] == template_name:
                        display_name = info['display_name']
                        break
                
                templates.append({
                    'name': template_name,
                    'path': str(template_path),
                    'display_name': display_name,
                    'required_sheets': []
                })
    
    # 템플릿 이름으로 정렬
    templates.sort(key=lambda x: x['display_name'])
    
    return jsonify({'templates': templates})


@app.route('/api/template-sheets', methods=['POST'])
def get_template_sheets():
    """템플릿이 필요한 시트 목록 반환"""
    try:
        template_name = request.form.get('template_name', '')
        if not template_name:
            return jsonify({'error': '템플릿명이 필요합니다.'}), 400
        
        template_path = Path('templates') / template_name
        if not template_path.exists():
            return jsonify({'error': f'템플릿 파일을 찾을 수 없습니다: {template_name}'}), 404
        
        # 템플릿에서 필요한 시트 목록 추출
        template_manager = TemplateManager(str(template_path))
        template_manager.load_template()
        markers = template_manager.extract_markers()
        
        # 마커에서 시트명 추출 (중복 제거)
        required_sheets = set()
        for marker in markers:
            sheet_name = marker.get('sheet_name', '').strip()
            if sheet_name:
                required_sheets.add(sheet_name)
        
        return jsonify({
            'template_name': template_name,
            'required_sheets': list(required_sheets)
        })
    except Exception as e:
        return jsonify({
            'error': f'템플릿 분석 중 오류가 발생했습니다: {str(e)}'
        }), 500


@app.route('/api/process', methods=['POST'])
def process_template():
    """엑셀 파일과 템플릿을 처리하여 결과 생성"""
    try:
        # 엑셀 파일 처리: 업로드된 파일이 있으면 사용, 없으면 기본 파일 사용
        excel_file = request.files.get('excel_file')
        excel_path = None
        use_default_file = False
        
        if excel_file and excel_file.filename:
            # 업로드된 파일이 있는 경우
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다. (.xlsx, .xls만 가능)'}), 400
            
            # 엑셀 파일 저장
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            # 업로드 폴더가 존재하는지 확인하고 없으면 생성
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            try:
                upload_folder.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                return jsonify({'error': f'업로드 폴더 생성 실패: {str(e)}'}), 500
            
            excel_path = upload_folder / excel_filename
            
            # 파일 저장 시도
            try:
                excel_file.save(str(excel_path))
            except Exception as e:
                return jsonify({'error': f'파일 저장 중 오류가 발생했습니다: {str(e)}'}), 500
            
            # 파일이 제대로 저장되었는지 확인 (크기도 확인)
            if not excel_path.exists():
                return jsonify({'error': '파일 저장에 실패했습니다.'}), 400
            
            # 파일 크기 확인 (0바이트 파일 방지)
            if excel_path.stat().st_size == 0:
                return jsonify({'error': '저장된 파일이 비어있습니다.'}), 400
        else:
            # 기본 엑셀 파일 사용
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'error': f'기본 엑셀 파일을 찾을 수 없습니다: {DEFAULT_EXCEL_FILE.name}'}), 400
            
            excel_path = DEFAULT_EXCEL_FILE
            use_default_file = True
        
        # 템플릿명, 연도 및 분기 파라미터 가져오기
        template_name = request.form.get('template_name', '')
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        
        if not template_name:
            return jsonify({'error': '템플릿을 선택해주세요.'}), 400
        
        template_path = Path('templates') / template_name
        
        if not template_path.exists():
            return jsonify({'error': f'템플릿 파일을 찾을 수 없습니다: {template_name}'}), 404
        
        try:
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
                return jsonify({
                    'error': '템플릿에서 필요한 시트를 찾을 수 없습니다.'
                }), 400
            
            # 파일 경로 검증
            if not excel_path.exists():
                return jsonify({'error': f'엑셀 파일을 찾을 수 없습니다: {excel_path}'}), 400
            
            # 파일 크기 확인
            file_size = excel_path.stat().st_size
            if file_size == 0:
                return jsonify({'error': '엑셀 파일이 비어있습니다.'}), 400
            
            # 엑셀 추출기 초기화
            excel_extractor = ExcelExtractor(str(excel_path))
            excel_extractor.load_workbook()
            
            # 사용 가능한 시트 목록 가져오기
            sheet_names = excel_extractor.get_sheet_names()
            
            # 필요한 시트가 모두 존재하는지 확인
            missing_sheets = []
            actual_sheet_mapping = {}  # 템플릿 시트명 -> 실제 시트명 매핑
            
            flexible_mapper = FlexibleMapper(excel_extractor)
            for required_sheet in required_sheets:
                # 유연한 매핑으로 실제 시트 찾기
                actual_sheet = flexible_mapper.find_sheet_by_name(required_sheet)
                if actual_sheet:
                    actual_sheet_mapping[required_sheet] = actual_sheet
                else:
                    missing_sheets.append(required_sheet)
            
            if missing_sheets:
                return jsonify({
                    'error': f'필요한 시트를 찾을 수 없습니다: {", ".join(missing_sheets)}. 사용 가능한 시트: {", ".join(sheet_names)}'
                }), 400
            
            # 첫 번째 필요한 시트를 기본 시트로 사용 (연도/분기 감지용)
            primary_sheet = list(actual_sheet_mapping.values())[0]
            
            # 연도 및 분기 자동 감지 또는 사용자 입력값 사용
            period_detector = PeriodDetector(excel_extractor)
            periods_info = period_detector.detect_available_periods(primary_sheet)
            
            if year_str and quarter_str:
                # 사용자가 입력한 값 사용
                year = int(year_str)
                quarter = int(quarter_str)
                
                # 유효성 검증
                is_valid, error_msg = period_detector.validate_period(sheet_name, year, quarter)
                if not is_valid:
                    return jsonify({'error': error_msg}), 400
            else:
                # 자동 감지된 기본값 사용
                year = periods_info['default_year']
                quarter = periods_info['default_quarter']
            
            # 템플릿 필러 초기화 및 처리
            template_filler = TemplateFiller(template_manager, excel_extractor, schema_loader)
            
            # primary_sheet를 사용하여 연도/분기 감지 (템플릿은 자동으로 필요한 시트를 찾음)
            filled_template = template_filler.fill_template(
                sheet_name=primary_sheet,  # 연도/분기 감지용
                year=year, 
                quarter=quarter
            )
            
            # display_name 찾기
            display_name = template_name.replace('.html', '')
            template_mapping = schema_loader.load_template_mapping()
            for sheet_name, info in template_mapping.items():
                if info['template'] == template_name:
                    display_name = info['display_name']
                    break
            
            # 결과 저장
            # 파일명 형식: (연도)년_(분기)분기_지역경제동향_보도자료(템플릿명).html
            output_filename = f"{year}년_{quarter}분기_지역경제동향_보도자료({display_name}).html"
            output_path = Path(app.config['OUTPUT_FOLDER']) / output_filename
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(filled_template)
            
            # 엑셀 파일 닫기
            excel_extractor.close()
            
            # 결과 반환
            return jsonify({
                'success': True,
                'output_filename': output_filename,
                'sheet_names': sheet_names,
                'message': '보도자료가 성공적으로 생성되었습니다.'
            })
            
        except FileNotFoundError as e:
            return jsonify({
                'error': f'엑셀 파일을 찾을 수 없습니다: {str(e)}'
            }), 400
        except PermissionError as e:
            return jsonify({
                'error': f'엑셀 파일에 접근할 수 없습니다. 파일이 다른 프로그램에서 사용 중일 수 있습니다: {str(e)}'
            }), 400
        except Exception as e:
            # 더 자세한 오류 정보 제공
            error_details = str(e)
            if 'BadZipFile' in str(type(e)) or 'zipfile' in str(e).lower():
                error_details = '엑셀 파일이 손상되었거나 올바른 형식이 아닙니다. 파일을 다시 확인해주세요.'
            elif 'openpyxl' in str(e).lower():
                error_details = f'엑셀 파일을 읽는 중 오류가 발생했습니다: {str(e)}'
            
            return jsonify({
                'error': f'처리 중 오류가 발생했습니다: {error_details}'
            }), 500
            
        finally:
            # 임시 엑셀 파일 삭제 (기본 파일이 아닌 경우에만)
            if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
                try:
                    excel_path.unlink()
                except Exception:
                    pass  # 파일 삭제 실패는 무시
    
    except Exception as e:
        return jsonify({
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@app.route('/api/download/<filename>')
def download_file(filename):
    """생성된 파일 다운로드"""
    try:
        return send_from_directory(
            app.config['OUTPUT_FOLDER'],
            filename,
            as_attachment=True
        )
    except Exception as e:
        return jsonify({'error': f'파일을 찾을 수 없습니다: {str(e)}'}), 404


@app.route('/api/preview/<filename>')
def preview_file(filename):
    """생성된 파일 미리보기"""
    try:
        file_path = Path(app.config['OUTPUT_FOLDER']) / filename
        if file_path.exists():
            return send_file(str(file_path))
        return jsonify({'error': '파일을 찾을 수 없습니다.'}), 404
    except Exception as e:
        return jsonify({'error': f'오류가 발생했습니다: {str(e)}'}), 500


@app.route('/api/compare-answer', methods=['POST'])
def compare_answer():
    """생성된 결과와 정답 파일 비교"""
    try:
        template_name = request.form.get('template_name', '')
        if not template_name:
            return jsonify({'error': '템플릿명이 필요합니다.'}), 400
        
        # 정답 파일 경로 찾기
        correct_answer_dir = BASE_DIR / 'correct_answer'
        
        # 템플릿명에서 정답 파일명 찾기
        template_mapping = schema_loader.load_template_mapping()
        answer_filename = None
        
        # 템플릿 매핑에서 display_name 찾기
        for sheet_name, info in template_mapping.items():
            if info['template'] == template_name:
                # display_name을 파일명으로 변환
                display_name = info['display_name']
                # 한글 파일명으로 변환
                answer_filename = f"{display_name}.png"
                break
        
        if not answer_filename:
            # 기본값: 템플릿명에서 확장자 제거
            answer_filename = template_name.replace('.html', '.png')
        
        answer_path = correct_answer_dir / answer_filename
        
        if not answer_path.exists():
            return jsonify({
                'error': f'정답 파일을 찾을 수 없습니다: {answer_filename}',
                'answer_file': answer_filename
            }), 404
        
        # 생성된 파일 경로
        output_filename = request.form.get('output_filename', '')
        if not output_filename:
            return jsonify({'error': '출력 파일명이 필요합니다.'}), 400
        
        output_path = Path(app.config['OUTPUT_FOLDER']) / output_filename
        
        if not output_path.exists():
            return jsonify({'error': '생성된 파일을 찾을 수 없습니다.'}), 404
        
        return jsonify({
            'success': True,
            'answer_file': answer_filename,
            'answer_path': str(answer_path.relative_to(BASE_DIR)),
            'output_file': output_filename,
            'output_path': str(output_path.relative_to(BASE_DIR)),
            'message': '정답 파일과 비교할 준비가 되었습니다.'
        })
        
    except Exception as e:
        return jsonify({
            'error': f'비교 중 오류가 발생했습니다: {str(e)}'
        }), 500


@app.route('/api/answer-image/<filename>')
def get_answer_image(filename):
    """정답 이미지 파일 반환"""
    try:
        answer_path = BASE_DIR / 'correct_answer' / filename
        if answer_path.exists():
            return send_file(str(answer_path))
        return jsonify({'error': '정답 파일을 찾을 수 없습니다.'}), 404
    except Exception as e:
        return jsonify({'error': f'오류가 발생했습니다: {str(e)}'}), 500


@app.route('/api/check-default-file', methods=['GET'])
def check_default_file():
    """기본 엑셀 파일 존재 여부 확인"""
    try:
        exists = DEFAULT_EXCEL_FILE.exists()
        return jsonify({
            'exists': exists,
            'filename': DEFAULT_EXCEL_FILE.name,
            'message': '기본 파일을 찾을 수 있습니다.' if exists else f'기본 엑셀 파일을 찾을 수 없습니다: {DEFAULT_EXCEL_FILE.name}'
        })
    except Exception as e:
        return jsonify({
            'exists': False,
            'error': f'파일 확인 중 오류가 발생했습니다: {str(e)}'
        }), 500


@app.route('/api/validate', methods=['POST'])
def validate_files():
    """파일 유효성 검증"""
    try:
        # 엑셀 파일 처리: 업로드된 파일이 있으면 사용, 없으면 기본 파일 사용
        excel_file = request.files.get('excel_file')
        excel_path = None
        use_default_file = False
        
        if excel_file and excel_file.filename:
            # 업로드된 파일이 있는 경우
            if not allowed_file(excel_file.filename):
                return jsonify({'valid': False, 'error': '지원하지 않는 파일 형식입니다.'}), 400
            
            # 엑셀 파일 임시 저장 및 검증
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'valid': False, 'error': '파일명이 유효하지 않습니다.'}), 400
            
            # 업로드 폴더가 존재하는지 확인하고 없으면 생성
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            try:
                upload_folder.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                return jsonify({'valid': False, 'error': f'업로드 폴더 생성 실패: {str(e)}'}), 500
            
            excel_path = upload_folder / excel_filename
            
            # 파일 저장 시도
            try:
                excel_file.save(str(excel_path))
            except Exception as e:
                return jsonify({'valid': False, 'error': f'파일 저장 중 오류가 발생했습니다: {str(e)}'}), 500
            
            # 파일이 제대로 저장되었는지 확인 (크기도 확인)
            if not excel_path.exists():
                return jsonify({'valid': False, 'error': '파일 저장에 실패했습니다.'}), 400
            
            # 파일 크기 확인 (0바이트 파일 방지)
            if excel_path.stat().st_size == 0:
                return jsonify({'valid': False, 'error': '저장된 파일이 비어있습니다.'}), 400
        else:
            # 기본 엑셀 파일 사용
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'valid': False, 'error': f'기본 엑셀 파일을 찾을 수 없습니다: {DEFAULT_EXCEL_FILE.name}'}), 400
            
            excel_path = DEFAULT_EXCEL_FILE
            use_default_file = True
        
        try:
            # 파일 경로 검증
            if not excel_path.exists():
                return jsonify({
                    'valid': False,
                    'error': f'엑셀 파일을 찾을 수 없습니다: {excel_path}'
                }), 400
            
            # 파일 크기 확인
            file_size = excel_path.stat().st_size
            if file_size == 0:
                return jsonify({
                    'valid': False,
                    'error': '엑셀 파일이 비어있습니다.'
                }), 400
            
            excel_extractor = ExcelExtractor(str(excel_path))
            excel_extractor.load_workbook()
            sheet_names = excel_extractor.get_sheet_names()
            
            if not sheet_names:
                excel_extractor.close()
                return jsonify({
                    'valid': False,
                    'error': '엑셀 파일에 시트가 없습니다.'
                }), 400
            
            # 각 시트별로 사용 가능한 연도/분기 정보 수집
            period_detector = PeriodDetector(excel_extractor)
            sheets_info = {}
            for sheet_name in sheet_names:
                periods_info = period_detector.detect_available_periods(sheet_name)
                sheets_info[sheet_name] = {
                    'min_year': periods_info['min_year'],
                    'max_year': periods_info['max_year'],
                    'default_year': periods_info['default_year'],
                    'default_quarter': periods_info['default_quarter'],
                    'available_periods': periods_info['available_periods']
                }
            
            excel_extractor.close()
            
            return jsonify({
                'valid': True,
                'sheet_names': sheet_names,
                'sheets_info': sheets_info,
                'message': '파일이 유효합니다.'
            })
        except FileNotFoundError as e:
            return jsonify({
                'valid': False,
                'error': f'엑셀 파일을 찾을 수 없습니다: {str(e)}'
            }), 400
        except PermissionError as e:
            return jsonify({
                'valid': False,
                'error': f'엑셀 파일에 접근할 수 없습니다. 파일이 다른 프로그램에서 사용 중일 수 있습니다: {str(e)}'
            }), 400
        except Exception as e:
            error_details = str(e)
            # 더 자세한 오류 정보 제공
            if 'BadZipFile' in str(type(e)) or 'zipfile' in str(e).lower():
                error_details = '엑셀 파일이 손상되었거나 올바른 형식이 아닙니다. 파일을 다시 확인해주세요.'
            elif 'openpyxl' in str(e).lower():
                error_details = f'엑셀 파일을 읽는 중 오류가 발생했습니다: {str(e)}'
            
            return jsonify({
                'valid': False,
                'error': f'엑셀 파일을 읽을 수 없습니다: {error_details}'
            }), 400
        finally:
            # 임시 엑셀 파일 삭제 (기본 파일이 아닌 경우에만)
            if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
                try:
                    excel_path.unlink()
                except Exception:
                    pass  # 파일 삭제 실패는 무시
    
    except Exception as e:
        return jsonify({
            'valid': False,
            'error': f'검증 중 오류가 발생했습니다: {str(e)}'
        }), 500


@app.route('/api/create-template', methods=['POST'])
def create_template():
    """이미지와 엑셀 파일에서 템플릿 생성 (헤더 기반 마커)"""
    try:
        # 파일 검증
        if 'image_file' not in request.files:
            return jsonify({'error': '이미지 파일이 없습니다.'}), 400
        
        image_file = request.files['image_file']
        if image_file.filename == '':
            return jsonify({'error': '이미지 파일을 선택해주세요.'}), 400
        
        if not allowed_image_file(image_file.filename):
            return jsonify({'error': '지원하지 않는 이미지 형식입니다. (png, jpg, jpeg, gif, bmp, webp만 가능)'}), 400
        
        # 엑셀 파일 검증 (선택사항이지만 권장)
        excel_file = request.files.get('excel_file')
        excel_path = None
        header_parser = None
        
        if excel_file and excel_file.filename:
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 엑셀 파일 형식입니다. (.xlsx, .xls만 가능)'}), 400
            
            # 엑셀 파일 저장
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            # 업로드 폴더가 존재하는지 확인하고 없으면 생성
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
            
            # 파일이 제대로 저장되었는지 확인
            if not excel_path.exists():
                return jsonify({'error': '파일 저장에 실패했습니다.'}), 400
        
        # 템플릿 이름 가져오기
        template_name = request.form.get('template_name', '').strip()
        if not template_name:
            # 파일명에서 확장자 제거하여 템플릿 이름으로 사용
            template_name = Path(image_file.filename).stem
        
        # 시트명 가져오기
        sheet_name = request.form.get('sheet_name', '').strip()
        if not sheet_name and excel_path:
            # 엑셀 파일이 있으면 첫 번째 시트 사용
            try:
                extractor = ExcelExtractor(str(excel_path))
                extractor.load_workbook()
                sheet_names = extractor.get_sheet_names()
                if sheet_names:
                    sheet_name = sheet_names[0]
                extractor.close()
            except:
                pass
        
        if not sheet_name:
            sheet_name = '시트1'
        
        # 이미지 파일 저장
        image_filename = secure_filename(image_file.filename)
        if not image_filename:
            return jsonify({'error': '이미지 파일명이 유효하지 않습니다.'}), 400
        
        # 업로드 폴더가 존재하는지 확인하고 없으면 생성
        upload_folder = Path(app.config['UPLOAD_FOLDER'])
        upload_folder.mkdir(parents=True, exist_ok=True)
        
        image_path = upload_folder / image_filename
        image_file.save(str(image_path))
        
        # 파일이 제대로 저장되었는지 확인
        if not image_path.exists():
            return jsonify({'error': '이미지 파일 저장에 실패했습니다.'}), 400
        
        try:
            # 엑셀 헤더 파서 초기화 (엑셀 파일이 있는 경우)
            if excel_path and excel_path.exists():
                header_parser = ExcelHeaderParser(str(excel_path))
            
            # 템플릿 생성기 초기화
            generator = TemplateGenerator(use_easyocr=True)
            
            # HTML 템플릿 생성 (헤더 파서 전달)
            html_template = generator.generate_html_template_from_excel(
                str(image_path),
                template_name,
                sheet_name,
                header_parser
            )
            
            # 템플릿 저장
            templates_dir = Path('templates')
            templates_dir.mkdir(exist_ok=True)
            
            template_filename = secure_filename(template_name) + '.html'
            template_path = templates_dir / template_filename
            
            with open(template_path, 'w', encoding='utf-8') as f:
                f.write(html_template)
            
            # 임시 파일 삭제
            if image_path.exists():
                image_path.unlink()
            if excel_path and excel_path.exists():
                excel_path.unlink()
            if header_parser:
                header_parser.close()
            
            return jsonify({
                'success': True,
                'template_name': template_filename,
                'template_path': str(template_path),
                'message': f'템플릿 "{template_name}"이 성공적으로 생성되었습니다.'
            })
            
        except Exception as e:
            # 임시 파일 삭제
            if image_path.exists():
                image_path.unlink()
            if excel_path and excel_path.exists():
                excel_path.unlink()
            if header_parser:
                header_parser.close()
            return jsonify({
                'error': f'템플릿 생성 중 오류가 발생했습니다: {str(e)}'
            }), 500
    
    except Exception as e:
        return jsonify({
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@app.route('/api/generate-pdf', methods=['POST'])
def generate_pdf():
    """10개 템플릿을 순서대로 처리하여 10페이지 PDF 생성"""
    try:
        # 엑셀 파일 처리: 업로드된 파일이 있으면 사용, 없으면 기본 파일 사용
        excel_file = request.files.get('excel_file')
        excel_path = None
        use_default_file = False
        
        if excel_file and excel_file.filename:
            # 업로드된 파일이 있는 경우
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다. (.xlsx, .xls만 가능)'}), 400
            
            # 엑셀 파일 저장
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            # 업로드 폴더가 존재하는지 확인하고 없으면 생성
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
            
            # 파일이 제대로 저장되었는지 확인
            if not excel_path.exists():
                return jsonify({'error': '파일 저장에 실패했습니다.'}), 400
        else:
            # 기본 엑셀 파일 사용
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'error': f'기본 엑셀 파일을 찾을 수 없습니다: {DEFAULT_EXCEL_FILE.name}'}), 400
            
            excel_path = DEFAULT_EXCEL_FILE
            use_default_file = True
        
        # 연도 및 분기 파라미터 가져오기
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        
        if not year_str or not quarter_str:
            return jsonify({'error': '연도와 분기를 입력해주세요.'}), 400
        
        year = int(year_str)
        quarter = int(quarter_str)
        
        # PDF 생성기 초기화
        pdf_generator = PDFGenerator(app.config['OUTPUT_FOLDER'])
        
        # PDF 생성 라이브러리 확인
        is_available, error_msg = pdf_generator.check_pdf_generator_available()
        if not is_available:
            return jsonify({'error': error_msg}), 500
        
        # PDF 생성
        success, result = pdf_generator.generate_pdf(
            excel_path=str(excel_path),
            year=year,
            quarter=quarter,
            templates_dir='templates'
        )
        
        # 임시 엑셀 파일 삭제 (기본 파일이 아닌 경우에만)
        if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
            try:
                excel_path.unlink()
            except Exception:
                pass  # 파일 삭제 실패는 무시
        
        if success:
            return jsonify(result)
        else:
            return jsonify({'error': result}), 500
    
    except Exception as e:
        return jsonify({
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@app.errorhandler(413)
def request_entity_too_large(error):
    """파일 크기 제한 초과 에러 처리"""
    return jsonify({
        'error': '파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.'
    }), 413


if __name__ == '__main__':
    # 출력 폴더 생성
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    
    print("=" * 50)
    print("지역경제동향 보도자료 자동생성")
    print("=" * 50)
    print(f"웹 서버가 시작되었습니다.")
    print(f"브라우저에서 http://localhost:8000 을 열어주세요.")
    print(f"최대 파일 크기: 100MB")
    print("=" * 50)
    
    app.run(debug=True, host='0.0.0.0', port=8000)

