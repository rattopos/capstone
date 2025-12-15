"""
Flask 웹 애플리케이션
통계청 보도자료 자동 생성 시스템 웹 인터페이스
"""

import os
import tempfile
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify, flash
from werkzeug.utils import secure_filename
import traceback

from src.template_manager import TemplateManager
from src.excel_extractor import ExcelExtractor
from src.template_filler import TemplateFiller
from src.config import Config

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

# 허용된 파일 확장자
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'html'}


def allowed_file(filename):
    """파일 확장자 검증"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('web_interface.html')


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
            default_template = Path('templates/mining_manufacturing_production.html')
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
            
            # 템플릿 채우기
            template_filler = TemplateFiller(template_manager, excel_extractor, config)
            filled_template = template_filler.fill_template()
            
            # 결과 파일 저장
            output_filename = f"result_{os.path.basename(excel_path).rsplit('.', 1)[0]}.html"
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(filled_template)
            
            # 엑셀 파일 닫기
            excel_extractor.close()
            
            # 임시 엑셀 파일 삭제
            if os.path.exists(excel_path):
                os.remove(excel_path)
            
            # 결과 파일 반환
            return send_file(
                output_path,
                as_attachment=True,
                download_name=output_filename,
                mimetype='text/html'
            )
            
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
    app.run(debug=True, host='0.0.0.0', port=5000)

