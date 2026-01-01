# -*- coding: utf-8 -*-
"""
메인 페이지 라우트
"""

from flask import Blueprint, render_template, send_file, make_response
from pathlib import Path
from urllib.parse import quote

from config.reports import REPORT_ORDER, REGIONAL_REPORTS
from config.settings import TEMPLATES_DIR, UPLOAD_FOLDER, DEBUG_FOLDER, EXPORT_FOLDER, BASE_DIR

main_bp = Blueprint('main', __name__)


def send_file_with_korean_filename(filepath, filename, mimetype=None):
    """한글 파일명을 지원하는 파일 다운로드 응답 생성 (RFC 5987)"""
    if mimetype:
        response = make_response(send_file(filepath, mimetype=mimetype))
    else:
        response = make_response(send_file(filepath))
    
    # RFC 5987 방식으로 한글 파일명 인코딩
    encoded_filename = quote(filename, safe='')
    
    # Content-Disposition 헤더 설정 (ASCII fallback + UTF-8 filename)
    ascii_filename = filename.encode('ascii', 'ignore').decode('ascii') or 'download'
    response.headers['Content-Disposition'] = (
        f"attachment; filename=\"{ascii_filename}\"; "
        f"filename*=UTF-8''{encoded_filename}"
    )
    
    return response


@main_bp.route('/')
def index():
    """메인 대시보드 페이지"""
    return render_template('dashboard.html', reports=REPORT_ORDER, regional_reports=REGIONAL_REPORTS)


@main_bp.route('/preview/infographic')
def preview_infographic():
    """인포그래픽 미리보기 (직접 접근용)"""
    output_path = TEMPLATES_DIR / 'infographic_output.html'
    if output_path.exists():
        return send_file(output_path)
    return "인포그래픽이 아직 생성되지 않았습니다.", 404


@main_bp.route('/uploads/<filename>')
def download_file(filename):
    """업로드된 파일 다운로드 (한글 파일명 지원)"""
    filepath = UPLOAD_FOLDER / filename
    if filepath.exists():
        return send_file_with_korean_filename(str(filepath), filename)
    return "파일을 찾을 수 없습니다.", 404


@main_bp.route('/view/<filename>')
def view_file(filename):
    """파일 직접 보기 (다운로드 없이) - uploads와 debug 폴더 모두 확인"""
    # uploads 폴더 먼저 확인
    filepath = UPLOAD_FOLDER / filename
    if filepath.exists():
        if filename.endswith('.html'):
            return send_file(str(filepath), mimetype='text/html')
        return send_file(str(filepath))
    
    # debug 폴더 확인
    debug_filepath = DEBUG_FOLDER / filename
    if debug_filepath.exists():
        if filename.endswith('.html'):
            return send_file(str(debug_filepath), mimetype='text/html')
        return send_file(str(debug_filepath))
    
    return "파일을 찾을 수 없습니다.", 404


@main_bp.route('/debug/<filename>')
def view_debug_file(filename):
    """디버그 파일 직접 보기"""
    filepath = DEBUG_FOLDER / filename
    if filepath.exists():
        if filename.endswith('.html'):
            return send_file(str(filepath), mimetype='text/html')
        return send_file(str(filepath))
    return "디버그 파일을 찾을 수 없습니다.", 404


@main_bp.route('/exports/<path:filepath>')
def view_export_file(filepath):
    """내보내기 파일 직접 보기 (한글 불러오기용 HTML)"""
    file_path = EXPORT_FOLDER / filepath
    if file_path.exists() and file_path.is_file():
        if filepath.endswith('.html'):
            return send_file(str(file_path), mimetype='text/html')
        elif filepath.endswith('.png') or filepath.endswith('.jpg') or filepath.endswith('.jpeg'):
            return send_file(str(file_path))
        return send_file(str(file_path))
    return "파일을 찾을 수 없습니다.", 404


@main_bp.route('/templates/<filename>')
def serve_template_file(filename):
    """templates 폴더의 정적 파일 제공 (이미지 등)"""
    filepath = TEMPLATES_DIR / filename
    if filepath.exists() and filepath.is_file():
        # 이미지 파일
        if filename.endswith('.png'):
            return send_file(str(filepath), mimetype='image/png')
        elif filename.endswith('.jpg') or filename.endswith('.jpeg'):
            return send_file(str(filepath), mimetype='image/jpeg')
        elif filename.endswith('.svg'):
            return send_file(str(filepath), mimetype='image/svg+xml')
        elif filename.endswith('.css'):
            return send_file(str(filepath), mimetype='text/css')
        elif filename.endswith('.js'):
            return send_file(str(filepath), mimetype='application/javascript')
        return send_file(str(filepath))
    return "파일을 찾을 수 없습니다.", 404


@main_bp.route('/download-export/<export_dir>')
def download_export_zip(export_dir):
    """내보내기 폴더를 ZIP으로 다운로드"""
    import zipfile
    import tempfile
    import shutil
    
    export_path = EXPORT_FOLDER / export_dir
    if not export_path.exists() or not export_path.is_dir():
        return "내보내기 폴더를 찾을 수 없습니다.", 404
    
    try:
        # 임시 ZIP 파일 생성
        temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
        temp_zip.close()
        
        with zipfile.ZipFile(temp_zip.name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in export_path.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(export_path)
                    zipf.write(str(file_path), arcname=str(arcname))
        
        return send_file_with_korean_filename(
            temp_zip.name,
            f'{export_dir}.zip',
            'application/zip'
        )
    except Exception as e:
        return f"ZIP 생성 오류: {str(e)}", 500

