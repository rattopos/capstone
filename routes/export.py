"""
내보내기 관련 라우트
/api/generate-pdf, /api/generate-docx, /api/download, /api/preview
"""

from pathlib import Path
from flask import Blueprint, request, jsonify, current_app, send_file, send_from_directory
from werkzeug.utils import secure_filename

from .common import allowed_file, get_excel_extractor, DEFAULT_EXCEL_FILE

export_bp = Blueprint('export', __name__, url_prefix='/api')


@export_bp.route('/generate-pdf', methods=['POST'])
def generate_pdf():
    """10개 템플릿을 순서대로 처리하여 10페이지 PDF 생성"""
    try:
        from src.generators.pdf_generator import PDFGenerator
        
        excel_file = request.files.get('excel_file')
        excel_path = None
        use_default_file = False
        
        if excel_file and excel_file.filename:
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다.'}), 400
            
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            upload_folder = Path(current_app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
            
            if not excel_path.exists():
                return jsonify({'error': '파일 저장에 실패했습니다.'}), 400
        else:
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'error': f'기본 엑셀 파일을 찾을 수 없습니다.'}), 400
            excel_path = DEFAULT_EXCEL_FILE
            use_default_file = True
        
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        
        if not year_str or not quarter_str:
            return jsonify({'error': '연도와 분기를 입력해주세요.'}), 400
        
        year = int(year_str)
        quarter = int(quarter_str)
        
        pdf_generator = PDFGenerator(current_app.config['OUTPUT_FOLDER'])
        
        is_available, error_msg = pdf_generator.check_pdf_generator_available()
        if not is_available:
            return jsonify({'error': error_msg}), 500
        
        success, result = pdf_generator.generate_pdf(
            excel_path=str(excel_path),
            year=year,
            quarter=quarter,
            templates_dir='templates'
        )
        
        if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
            try:
                excel_path.unlink()
            except Exception:
                pass
        
        if success:
            return jsonify(result)
        else:
            return jsonify({'error': result}), 500
    
    except Exception as e:
        return jsonify({'error': f'서버 오류가 발생했습니다: {str(e)}'}), 500


@export_bp.route('/generate-docx', methods=['POST'])
def generate_docx():
    """10개 템플릿을 순서대로 처리하여 DOCX 생성"""
    try:
        from src.generators.docx_generator import DOCXGenerator
        
        excel_file = request.files.get('excel_file')
        excel_path = None
        use_default_file = False
        
        if excel_file and excel_file.filename:
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다.'}), 400
            
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            upload_folder = Path(current_app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
            
            if not excel_path.exists():
                return jsonify({'error': '파일 저장에 실패했습니다.'}), 400
        else:
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'error': f'기본 엑셀 파일을 찾을 수 없습니다.'}), 400
            excel_path = DEFAULT_EXCEL_FILE
            use_default_file = True
        
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        
        if not year_str or not quarter_str:
            return jsonify({'error': '연도와 분기를 입력해주세요.'}), 400
        
        year = int(year_str)
        quarter = int(quarter_str)
        
        docx_generator = DOCXGenerator(current_app.config['OUTPUT_FOLDER'])
        
        is_available, error_msg = docx_generator.check_docx_generator_available()
        if not is_available:
            return jsonify({'error': error_msg}), 500
        
        success, result = docx_generator.generate_docx(
            excel_path=str(excel_path),
            year=year,
            quarter=quarter,
            templates_dir='templates'
        )
        
        if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
            try:
                excel_path.unlink()
            except Exception:
                pass
        
        if success:
            return jsonify(result)
        else:
            return jsonify({'error': result}), 500
    
    except Exception as e:
        return jsonify({'error': f'서버 오류가 발생했습니다: {str(e)}'}), 500


@export_bp.route('/download/<filename>')
def download_file(filename):
    """생성된 파일 다운로드"""
    try:
        output_folder = Path(current_app.config['OUTPUT_FOLDER'])
        file_path = output_folder / filename
        
        if not file_path.exists():
            return jsonify({'error': '파일을 찾을 수 없습니다.'}), 404
        
        return send_file(str(file_path), as_attachment=True)
    except Exception as e:
        return jsonify({'error': f'다운로드 중 오류: {str(e)}'}), 500


@export_bp.route('/preview/<filename>')
def preview_file(filename):
    """생성된 파일 미리보기 (PDF)"""
    try:
        output_folder = Path(current_app.config['OUTPUT_FOLDER'])
        file_path = output_folder / filename
        
        if not file_path.exists():
            return jsonify({'error': '파일을 찾을 수 없습니다.'}), 404
        
        return send_file(str(file_path), mimetype='application/pdf')
    except Exception as e:
        return jsonify({'error': f'미리보기 중 오류: {str(e)}'}), 500

