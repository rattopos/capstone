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

# 시트명과 템플릿 파일 매핑
SHEET_TEMPLATE_MAPPING = {
    '광공업생산': {
        'template': '광공업생산.html',
        'display_name': '광공업생산'
    },
    '서비스업생산': {
        'template': 'service_production.html',
        'display_name': '서비스업생산'
    },
    '소비(소매, 추가)': {
        'template': 'retail_sales.html',
        'display_name': '소비(소매, 추가)'
    },
    '고용': {
        'template': 'employment.html',
        'display_name': '고용'
    },
    '고용(kosis)': {
        'template': 'employment_kosis.html',
        'display_name': '고용(kosis)'
    },
    '고용률': {
        'template': 'employment_rate.html',
        'display_name': '고용률'
    },
    '실업자 수': {
        'template': 'unemployed.html',
        'display_name': '실업자 수'
    },
    '지출목적별 물가': {
        'template': 'price_by_purpose.html',
        'display_name': '지출목적별 물가'
    },
    '품목성질별 물가': {
        'template': 'price_by_item.html',
        'display_name': '품목성질별 물가'
    },
    '건설 (공표자료)': {
        'template': 'construction_orders.html',
        'display_name': '건설수주'
    },
    '수출': {
        'template': 'exports.html',
        'display_name': '수출'
    },
    '수입': {
        'template': 'imports.html',
        'display_name': '수입'
    },
    '연령별 인구이동': {
        'template': 'population_movement_by_age.html',
        'display_name': '연령별 인구이동'
    },
    '시도 간 이동': {
        'template': 'inter_sido_movement.html',
        'display_name': '시도 간 이동'
    },
    '시군구인구이동': {
        'template': 'population_movement_sigungu.html',
        'display_name': '시군구인구이동'
    }
}


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
    # 정확한 매칭 시도
    if sheet_name in SHEET_TEMPLATE_MAPPING:
        return SHEET_TEMPLATE_MAPPING[sheet_name]
    
    # 부분 매칭 시도 (키워드 기반)
    sheet_lower = sheet_name.lower()
    for key, value in SHEET_TEMPLATE_MAPPING.items():
        if key.lower() in sheet_lower or sheet_lower in key.lower():
            return value
    
    # 특수 케이스: 소비/소매 관련
    if '소비' in sheet_name or '소매' in sheet_name:
        return SHEET_TEMPLATE_MAPPING['소비(소매, 추가)']
    
    # 기본값: 광공업생산 템플릿 사용
    return SHEET_TEMPLATE_MAPPING['광공업생산']


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
                
                # display_name 찾기 (SHEET_TEMPLATE_MAPPING에서 먼저 찾고, 없으면 파일명 사용)
                display_name = template_name.replace('.html', '')
                for sheet_name, info in SHEET_TEMPLATE_MAPPING.items():
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
                for sheet_name, info in SHEET_TEMPLATE_MAPPING.items():
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
            template_filler = TemplateFiller(template_manager, excel_extractor)
            
            # primary_sheet를 사용하여 연도/분기 감지 (템플릿은 자동으로 필요한 시트를 찾음)
            filled_template = template_filler.fill_template(
                sheet_name=primary_sheet,  # 연도/분기 감지용
                year=year, 
                quarter=quarter
            )
            
            # display_name 찾기
            display_name = template_name.replace('.html', '')
            for sheet_name, info in SHEET_TEMPLATE_MAPPING.items():
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
            
        except Exception as e:
            return jsonify({
                'error': f'처리 중 오류가 발생했습니다: {str(e)}'
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
            upload_folder.mkdir(parents=True, exist_ok=True)
            
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
            
            # 파일이 제대로 저장되었는지 확인
            if not excel_path.exists():
                return jsonify({'valid': False, 'error': '파일 저장에 실패했습니다.'}), 400
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
            
            excel_extractor = ExcelExtractor(str(excel_path))
            excel_extractor.load_workbook()
            sheet_names = excel_extractor.get_sheet_names()
            
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
        except Exception as e:
            return jsonify({
                'valid': False,
                'error': f'엑셀 파일을 읽을 수 없습니다: {str(e)}'
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


@app.route('/api/generate-pdf', methods=['POST'])
def generate_pdf():
    """10개 템플릿을 순서대로 처리하여 10페이지 PDF 생성"""
    if not PDF_GENERATOR:
        return jsonify({
            'error': 'PDF 생성 라이브러리가 설치되지 않았습니다. 다음 중 하나를 설치해주세요:\n'
                     '1. playwright: pip install playwright && playwright install chromium\n'
                     '2. weasyprint: pip install weasyprint (macOS에서는 Homebrew로 시스템 라이브러리 설치 필요)'
        }), 500
    
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
        
        # 파일 경로 검증
        if not excel_path.exists():
            return jsonify({'error': f'엑셀 파일을 찾을 수 없습니다: {excel_path}'}), 400
        
        # 엑셀 추출기 초기화
        excel_extractor = ExcelExtractor(str(excel_path))
        excel_extractor.load_workbook()
            
            # 사용 가능한 시트 목록 가져오기
            sheet_names = excel_extractor.get_sheet_names()
            
            # 첫 번째 시트를 기본 시트로 사용 (연도/분기 감지용)
            primary_sheet = sheet_names[0] if sheet_names else None
            if not primary_sheet:
                excel_extractor.close()
                return jsonify({'error': '엑셀 파일에 시트가 없습니다.'}), 400
        
        # 연도 및 분기 자동 감지 또는 사용자 입력값 사용
        period_detector = PeriodDetector(excel_extractor)
        periods_info = period_detector.detect_available_periods(primary_sheet)
        
        if year_str and quarter_str:
            # 사용자가 입력한 값 사용
            year = int(year_str)
            quarter = int(quarter_str)
            
            # 유효성 검증
            is_valid, error_msg = period_detector.validate_period(primary_sheet, year, quarter)
            if not is_valid:
                excel_extractor.close()
                return jsonify({'error': error_msg}), 400
        else:
            # 자동 감지된 기본값 사용
            year = periods_info['default_year']
            quarter = periods_info['default_quarter']
        
        # 유연한 매퍼 초기화
        flexible_mapper = FlexibleMapper(excel_extractor)
        
        # 각 템플릿 처리
        filled_templates = []
        errors = []
        
        for template_name in TEMPLATE_ORDER:
            try:
                template_path = Path('templates') / template_name
                
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
                
                # 필요한 시트가 모두 존재하는지 확인
                missing_sheets = []
                actual_sheet_mapping = {}
                
                for required_sheet in required_sheets:
                    # 유연한 매핑으로 실제 시트 찾기
                    actual_sheet = flexible_mapper.find_sheet_by_name(required_sheet)
                    if actual_sheet:
                        actual_sheet_mapping[required_sheet] = actual_sheet
                    else:
                        missing_sheets.append(required_sheet)
                
                if missing_sheets:
                    errors.append(f'{template_name}: 필요한 시트를 찾을 수 없습니다: {", ".join(missing_sheets)}')
                    continue
                
                # 첫 번째 필요한 시트를 기본 시트로 사용
                primary_sheet_for_template = list(actual_sheet_mapping.values())[0]
                
                # 템플릿 필러 초기화 및 처리
                template_filler = TemplateFiller(template_manager, excel_extractor)
                
                filled_template = template_filler.fill_template(
                    sheet_name=primary_sheet_for_template,
                    year=year,
                    quarter=quarter
                )
                
                # HTML에서 body와 style 내용 추출 (완전한 HTML 문서인 경우)
                try:
                    soup = BeautifulSoup(filled_template, 'html.parser')
                    body = soup.find('body')
                    style = soup.find('style')
                    
                    template_content = ''
                    # 스타일 추가
                    if style:
                        template_content += f'<style>{style.string}</style>'
                    # body 내용 추가
                    if body:
                        # body 내용만 추출 (body 태그 제외)
                        body_content = ''.join(str(child) for child in body.children)
                        template_content += body_content
                    else:
                        # body가 없으면 전체 내용 사용
                        template_content = filled_template
                    
                    filled_templates.append(template_content)
                except:
                    # 파싱 실패 시 원본 사용
                    filled_templates.append(filled_template)
                
            except Exception as e:
                errors.append(f'{template_name}: {str(e)}')
                continue
        
        # 엑셀 파일 닫기
        excel_extractor.close()
        
        if not filled_templates:
            return jsonify({
                'error': f'처리된 템플릿이 없습니다. 오류: {"; ".join(errors)}'
            }), 400
        
        # 모든 HTML을 하나로 합치기 (페이지 브레이크 추가)
        combined_html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
        combined_html += '''
            @page {
                size: A4;
                margin: 2cm;
            }
            body {
                font-family: "Malgun Gothic", "맑은 고딕", sans-serif;
            }
            .page-break {
                page-break-after: always;
            }
            .page-break:last-child {
                page-break-after: auto;
            }
        '''
        combined_html += '</style></head><body>'
        
        for i, template_html in enumerate(filled_templates):
            # 각 템플릿을 div로 감싸고 페이지 브레이크 추가
            combined_html += f'<div class="page-break">{template_html}</div>'
        
        combined_html += '</body></html>'
        
        # PDF 생성
        pdf_path = Path(app.config['OUTPUT_FOLDER']) / f"{year}년_{quarter}분기_지역경제동향_보도자료_전체.pdf"
        
        try:
            if PDF_GENERATOR == 'playwright':
                # playwright를 사용하여 PDF 생성
                with sync_playwright() as p:
                    browser = p.chromium.launch()
                    page = browser.new_page()
                    
                    # HTML 파일을 임시로 저장하여 로드
                    temp_html_path = Path(app.config['OUTPUT_FOLDER']) / 'temp_combined.html'
                    with open(temp_html_path, 'w', encoding='utf-8') as f:
                        f.write(combined_html)
                    
                    # 파일 경로를 file:// URL로 변환
                    file_url = f"file://{temp_html_path.absolute()}"
                    page.goto(file_url)
                    
                    # PDF 생성
                    page.pdf(
                        path=str(pdf_path),
                        format='A4',
                        margin={'top': '2cm', 'right': '2cm', 'bottom': '2cm', 'left': '2cm'},
                        print_background=True
                    )
                    
                    browser.close()
                    
                    # 임시 HTML 파일 삭제
                    if temp_html_path.exists():
                        temp_html_path.unlink()
                        
            elif PDF_GENERATOR == 'weasyprint':
                # weasyprint를 사용하여 PDF 생성
                from weasyprint import HTML
                HTML(string=combined_html, base_url=str(Path('templates').absolute())).write_pdf(str(pdf_path))
            else:
                return jsonify({
                    'error': 'PDF 생성 라이브러리를 사용할 수 없습니다.'
                }), 500
                
        except Exception as e:
            # 임시 엑셀 파일 삭제 (기본 파일이 아닌 경우에만)
            if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
                try:
                    excel_path.unlink()
                except Exception:
                    pass  # 파일 삭제 실패는 무시
            return jsonify({
                'error': f'PDF 생성 중 오류가 발생했습니다: {str(e)}'
            }), 500
        
        # 결과 반환
        result = {
            'success': True,
            'output_filename': pdf_path.name,
            'message': f'{len(filled_templates)}개 템플릿이 성공적으로 PDF로 생성되었습니다.',
            'processed_templates': len(filled_templates),
            'total_templates': len(TEMPLATE_ORDER)
        }
        
        if errors:
            result['warnings'] = errors
        
        # 임시 엑셀 파일 삭제 (기본 파일이 아닌 경우에만)
        if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
            try:
                excel_path.unlink()
            except Exception:
                pass  # 파일 삭제 실패는 무시
        
        return jsonify(result)
    
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

