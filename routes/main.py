# -*- coding: utf-8 -*-
"""
메인 페이지 라우트
"""

from flask import Blueprint, render_template, send_file, make_response
from pathlib import Path
from urllib.parse import quote

from config.reports import REPORT_ORDER, REGIONAL_REPORTS
from config.settings import TEMPLATES_DIR, UPLOAD_FOLDER

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
    """파일 직접 보기 (다운로드 없이)"""
    filepath = UPLOAD_FOLDER / filename
    if filepath.exists():
        return send_file(str(filepath))
    return "파일을 찾을 수 없습니다.", 404

