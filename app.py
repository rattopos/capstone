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

app = Flask(__name__, template_folder='flask_templates', static_folder='static')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['OUTPUT_FOLDER'] = tempfile.mkdtemp()

# 허용된 파일 확장자
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'html'}
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'}

# 시트명과 템플릿 파일 매핑
SHEET_TEMPLATE_MAPPING = {
    '광공업생산': {
        'template': 'mining_manufacturing_production.html',
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
    """사용 가능한 템플릿 목록 반환"""
    templates_dir = Path('templates')
    templates = []
    
    if templates_dir.exists():
        for file in templates_dir.glob('*.html'):
            templates.append({
                'name': file.name,
                'path': str(file)
            })
    
    return jsonify({'templates': templates})


@app.route('/api/process', methods=['POST'])
def process_template():
    """엑셀 파일과 템플릿을 처리하여 결과 생성"""
    try:
        # 파일 검증
        if 'excel_file' not in request.files:
            return jsonify({'error': '엑셀 파일이 없습니다.'}), 400
        
        excel_file = request.files['excel_file']
        if excel_file.filename == '':
            return jsonify({'error': '엑셀 파일을 선택해주세요.'}), 400
        
        if not allowed_file(excel_file.filename):
            return jsonify({'error': '지원하지 않는 파일 형식입니다. (.xlsx, .xls만 가능)'}), 400
        
        # 시트명, 연도 및 분기 파라미터 가져오기
        sheet_name = request.form.get('sheet_name', '')
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        
        if not sheet_name:
            return jsonify({'error': '시트명을 선택해주세요.'}), 400
        
        # 시트명에 따라 템플릿 자동 선택
        template_info = get_template_for_sheet(sheet_name)
        template_name = template_info['template']
        sheet_display_name = template_info['display_name']
        
        template_path = Path('templates') / template_name
        
        if not template_path.exists():
            return jsonify({'error': f'템플릿 파일을 찾을 수 없습니다: {template_name}'}), 404
        
        # 엑셀 파일 저장
        excel_filename = secure_filename(excel_file.filename)
        excel_path = Path(app.config['UPLOAD_FOLDER']) / excel_filename
        excel_file.save(str(excel_path))
        
        try:
            # 템플릿 관리자 초기화
            template_manager = TemplateManager(str(template_path))
            template_manager.load_template()
            
            # 엑셀 추출기 초기화
            excel_extractor = ExcelExtractor(str(excel_path))
            excel_extractor.load_workbook()
            
            # 사용 가능한 시트 목록 가져오기
            sheet_names = excel_extractor.get_sheet_names()
            
            # 선택한 시트가 존재하는지 확인
            if sheet_name not in sheet_names:
                return jsonify({
                    'error': f'시트 "{sheet_name}"을 찾을 수 없습니다. 사용 가능한 시트: {", ".join(sheet_names)}'
                }), 400
            
            # 연도 및 분기 자동 감지 또는 사용자 입력값 사용
            period_detector = PeriodDetector(excel_extractor)
            periods_info = period_detector.detect_available_periods(sheet_name)
            
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
            filled_template = template_filler.fill_template(
                sheet_name=sheet_name,
                year=year, 
                quarter=quarter
            )
            
            # 결과 저장
            # 파일명 형식: (연도)년_(분기)분기_지역경제동향_보도자료(시트명).html
            output_filename = f"{year}년_{quarter}분기_지역경제동향_보도자료({sheet_display_name}).html"
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
            # 임시 엑셀 파일 삭제
            if excel_path.exists():
                excel_path.unlink()
    
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


@app.route('/api/validate', methods=['POST'])
def validate_files():
    """파일 유효성 검증"""
    try:
        if 'excel_file' not in request.files:
            return jsonify({'valid': False, 'error': '엑셀 파일이 없습니다.'}), 400
        
        excel_file = request.files['excel_file']
        if excel_file.filename == '':
            return jsonify({'valid': False, 'error': '엑셀 파일을 선택해주세요.'}), 400
        
        if not allowed_file(excel_file.filename):
            return jsonify({'valid': False, 'error': '지원하지 않는 파일 형식입니다.'}), 400
        
        # 엑셀 파일 임시 저장 및 검증
        excel_filename = secure_filename(excel_file.filename)
        excel_path = Path(app.config['UPLOAD_FOLDER']) / excel_filename
        excel_file.save(str(excel_path))
        
        try:
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
        except Exception as e:
            return jsonify({
                'valid': False,
                'error': f'엑셀 파일을 읽을 수 없습니다: {str(e)}'
            }), 400
        finally:
            if excel_path.exists():
                excel_path.unlink()
    
    except Exception as e:
        return jsonify({
            'valid': False,
            'error': f'검증 중 오류가 발생했습니다: {str(e)}'
        }), 500


@app.route('/api/create-template', methods=['POST'])
def create_template():
    """이미지에서 템플릿 생성"""
    try:
        # 파일 검증
        if 'image_file' not in request.files:
            return jsonify({'error': '이미지 파일이 없습니다.'}), 400
        
        image_file = request.files['image_file']
        if image_file.filename == '':
            return jsonify({'error': '이미지 파일을 선택해주세요.'}), 400
        
        if not allowed_image_file(image_file.filename):
            return jsonify({'error': '지원하지 않는 이미지 형식입니다. (png, jpg, jpeg, gif, bmp, webp만 가능)'}), 400
        
        # 템플릿 이름 가져오기
        template_name = request.form.get('template_name', '').strip()
        if not template_name:
            # 파일명에서 확장자 제거하여 템플릿 이름으로 사용
            template_name = Path(image_file.filename).stem
        
        # 시트명 가져오기 (선택사항)
        sheet_name = request.form.get('sheet_name', '시트1').strip()
        if not sheet_name:
            sheet_name = '시트1'
        
        # 이미지 파일 저장
        image_filename = secure_filename(image_file.filename)
        image_path = Path(app.config['UPLOAD_FOLDER']) / image_filename
        image_file.save(str(image_path))
        
        try:
            # 템플릿 생성기 초기화
            generator = TemplateGenerator(use_easyocr=True)
            
            # HTML 템플릿 생성
            html_template = generator.generate_html_template(
                str(image_path),
                template_name,
                sheet_name
            )
            
            # 템플릿 저장
            templates_dir = Path('templates')
            templates_dir.mkdir(exist_ok=True)
            
            template_filename = secure_filename(template_name) + '.html'
            template_path = templates_dir / template_filename
            
            with open(template_path, 'w', encoding='utf-8') as f:
                f.write(html_template)
            
            # 임시 이미지 파일 삭제
            if image_path.exists():
                image_path.unlink()
            
            return jsonify({
                'success': True,
                'template_name': template_filename,
                'template_path': str(template_path),
                'message': f'템플릿 "{template_name}"이 성공적으로 생성되었습니다.'
            })
            
        except Exception as e:
            # 임시 이미지 파일 삭제
            if image_path.exists():
                image_path.unlink()
            return jsonify({
                'error': f'템플릿 생성 중 오류가 발생했습니다: {str(e)}'
            }), 500
    
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

