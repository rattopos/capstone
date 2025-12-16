"""
Flask 웹 애플리케이션
통계청 보도자료 자동 생성 시스템 웹 인터페이스
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

app = Flask(__name__, template_folder='flask_templates', static_folder='static')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['OUTPUT_FOLDER'] = tempfile.mkdtemp()

# 허용된 파일 확장자
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'html'}


def allowed_file(filename):
    """파일 확장자 검증"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


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
        
        # 템플릿 경로 확인
        template_name = request.form.get('template', 'mining_manufacturing_production.html')
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
            
            # 템플릿 필러 초기화 및 처리
            template_filler = TemplateFiller(template_manager, excel_extractor)
            filled_template = template_filler.fill_template()
            
            # 결과 저장
            output_filename = f"result_{Path(excel_filename).stem}.html"
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
            excel_extractor.close()
            
            return jsonify({
                'valid': True,
                'sheet_names': sheet_names,
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
    print("통계청 보도자료 자동 생성 시스템")
    print("=" * 50)
    print(f"웹 서버가 시작되었습니다.")
    print(f"브라우저에서 http://localhost:8000 을 열어주세요.")
    print(f"최대 파일 크기: 100MB")
    print("=" * 50)
    
    app.run(debug=True, host='0.0.0.0', port=8000)

