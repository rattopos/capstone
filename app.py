"""
Flask 웹 애플리케이션
통계청 보도자료 자동 생성 시스템 웹 인터페이스
"""

import os
import tempfile
import uuid
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify, flash, redirect, url_for
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge
import traceback

from src.template_manager import TemplateManager
from src.excel_extractor import ExcelExtractor
from src.template_filler import TemplateFiller
from src.config import Config

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['RESULT_FOLDER'] = os.path.join(tempfile.gettempdir(), 'capstone_results')

# 허용된 파일 확장자
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'html'}


def allowed_file(filename):
    """파일 확장자 검증"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def ensure_result_folder():
    """결과 폴더 생성"""
    result_folder = app.config['RESULT_FOLDER']
    os.makedirs(result_folder, exist_ok=True)
    return result_folder


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('web_interface.html')


@app.route('/get_sheets', methods=['POST'])
def get_sheets():
    """엑셀 파일의 시트 목록 반환"""
    try:
        if 'excel_file' not in request.files:
            return jsonify({'error': '엑셀 파일이 필요합니다.'}), 400
        
        excel_file = request.files['excel_file']
        if excel_file.filename == '':
            return jsonify({'error': '엑셀 파일을 선택해주세요.'}), 400
        
        if not allowed_file(excel_file.filename):
            return jsonify({'error': '엑셀 파일만 업로드 가능합니다 (.xlsx, .xls)'}), 400
        
        # 임시 파일 저장
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(excel_file.filename))
        excel_file.save(excel_path)
        
        try:
            # 엑셀 추출기 초기화
            excel_extractor = ExcelExtractor(excel_path)
            excel_extractor.load_workbook()
            
            # 시트 목록 가져오기
            sheet_names = excel_extractor.get_sheet_names()
            
            # 엑셀 파일 닫기
            excel_extractor.close()
            
            # 임시 파일 삭제
            if os.path.exists(excel_path):
                os.remove(excel_path)
            
            return jsonify({'sheets': sheet_names})
        except Exception as e:
            # 에러 발생 시 임시 파일 정리
            if os.path.exists(excel_path):
                os.remove(excel_path)
            return jsonify({'error': f'시트 목록을 가져오는 중 오류가 발생했습니다: {str(e)}'}), 500
    
    except Exception as e:
        return jsonify({'error': f'서버 오류: {str(e)}'}), 500


@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    """파일 크기 초과 오류 처리"""
    max_size_mb = app.config['MAX_CONTENT_LENGTH'] / (1024 * 1024)
    return jsonify({
        'error': f'업로드한 파일이 너무 큽니다. 최대 파일 크기는 {int(max_size_mb)}MB입니다. 파일 크기를 확인하고 다시 시도해주세요.'
    }), 413


@app.route('/process', methods=['POST'])
def process():
    """파일 처리 및 보도자료 생성"""
    try:
        # 파일 검증
        if 'excel_file' not in request.files:
            return jsonify({'error': '엑셀 파일이 필요합니다.'}), 400
        
        excel_file = request.files['excel_file']
        template_file = request.files.get('template_file')
        
        if excel_file.filename == '':
            return jsonify({'error': '엑셀 파일을 선택해주세요.'}), 400
        
        if not allowed_file(excel_file.filename):
            return jsonify({'error': '엑셀 파일만 업로드 가능합니다 (.xlsx, .xls)'}), 400
        
        # 연도/분기 검증
        year = request.form.get('year', type=int)
        quarter = request.form.get('quarter', type=int)
        
        config = None
        if year and quarter:
            try:
                config = Config(year, quarter)
            except ValueError as e:
                return jsonify({'error': str(e)}), 400
        
        # 임시 파일 저장
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(excel_file.filename))
        excel_file.save(excel_path)
        
        # 템플릿 파일 처리
        if template_file and template_file.filename:
            if not allowed_file(template_file.filename):
                return jsonify({'error': '템플릿 파일은 HTML만 가능합니다.'}), 400
            template_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(template_file.filename))
            template_file.save(template_path)
        else:
            # 기본 템플릿 사용
            default_template = Path('templates/dynamic_template.html')
            if not default_template.exists():
                return jsonify({'error': '기본 템플릿 파일을 찾을 수 없습니다.'}), 500
            template_path = str(default_template)
        
        # 처리 실행
        try:
            # 템플릿 관리자 초기화
            template_manager = TemplateManager(template_path)
            template_manager.load_template()
            
            # 엑셀 추출기 초기화
            excel_extractor = ExcelExtractor(excel_path)
            excel_extractor.load_workbook()
            
            # 시트명 결정 (첫 번째 시트 또는 지정된 시트)
            sheet_name = request.form.get('sheet_name', '').strip()
            if not sheet_name:
                all_sheets = excel_extractor.get_sheet_names()
                sheet_name = all_sheets[0] if all_sheets else None
            
            # 템플릿 채우기
            template_filler = TemplateFiller(template_manager, excel_extractor, config, sheet_name=sheet_name)
            filled_template = template_filler.fill_template()
            
            # 원본 엑셀 파일명 저장 (다운로드 파일명용)
            original_excel_name = os.path.basename(excel_path).rsplit('.', 1)[0]
            
            # 엑셀 파일 닫기
            excel_extractor.close()
            
            # 임시 엑셀 파일 삭제
            if os.path.exists(excel_path):
                os.remove(excel_path)
            
            # 결과 파일 저장 (고유 ID 사용)
            result_folder = ensure_result_folder()
            file_id = str(uuid.uuid4())
            output_filename = f"result_{file_id}.html"
            output_path = os.path.join(result_folder, output_filename)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(filled_template)
            
            # 미리보기 페이지로 리다이렉트
            return jsonify({
                'success': True,
                'file_id': file_id,
                'filename': f"보도자료_{original_excel_name}.html"
            })
            
        except Exception as e:
            # 에러 발생 시 임시 파일 정리
            if os.path.exists(excel_path):
                os.remove(excel_path)
            if template_file and template_file.filename and os.path.exists(template_path):
                os.remove(template_path)
            
            error_msg = str(e)
            if app.debug:
                error_msg += f"\n\n{traceback.format_exc()}"
            return jsonify({'error': f'처리 중 오류가 발생했습니다: {error_msg}'}), 500
    
    except Exception as e:
        error_msg = str(e)
        if app.debug:
            error_msg += f"\n\n{traceback.format_exc()}"
        return jsonify({'error': f'서버 오류: {error_msg}'}), 500


@app.route('/preview/<file_id>')
def preview(file_id):
    """결과 미리보기 페이지"""
    result_folder = app.config['RESULT_FOLDER']
    file_path = os.path.join(result_folder, f"result_{file_id}.html")
    
    if not os.path.exists(file_path):
        return render_template('error.html', error='파일을 찾을 수 없습니다.'), 404
    
    return render_template('preview.html', file_id=file_id)


@app.route('/view/<file_id>')
def view_file(file_id):
    """결과 HTML 파일 직접 보기"""
    result_folder = app.config['RESULT_FOLDER']
    file_path = os.path.join(result_folder, f"result_{file_id}.html")
    
    if not os.path.exists(file_path):
        return jsonify({'error': '파일을 찾을 수 없습니다.'}), 404
    
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    return content, 200, {'Content-Type': 'text/html; charset=utf-8'}


@app.route('/download/<file_id>')
def download(file_id):
    """결과 파일 다운로드"""
    result_folder = app.config['RESULT_FOLDER']
    file_path = os.path.join(result_folder, f"result_{file_id}.html")
    
    if not os.path.exists(file_path):
        return jsonify({'error': '파일을 찾을 수 없습니다.'}), 404
    
    # 다운로드 파일명
    download_name = f"보도자료_{file_id[:8]}.html"
    
    return send_file(
        file_path,
        as_attachment=True,
        download_name=download_name,
        mimetype='text/html'
    )


@app.route('/health')
def health():
    """헬스 체크 엔드포인트"""
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    # 웹 인터페이스 템플릿 디렉토리 확인
    template_dir = Path('templates')
    if not template_dir.exists():
        template_dir.mkdir(parents=True, exist_ok=True)
    
    # 개발 모드로 실행
    app.run(debug=True, host='0.0.0.0', port=8000)

