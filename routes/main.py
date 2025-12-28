# -*- coding: utf-8 -*-
"""
메인 페이지 라우트
"""

from flask import Blueprint, render_template, send_file, send_from_directory
from pathlib import Path

from config.reports import REPORT_ORDER, REGIONAL_REPORTS
from config.settings import TEMPLATES_DIR, UPLOAD_FOLDER

main_bp = Blueprint('main', __name__)


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
    """업로드된 파일 다운로드"""
    return send_from_directory(str(UPLOAD_FOLDER), filename, as_attachment=True)


@main_bp.route('/view/<filename>')
def view_file(filename):
    """파일 직접 보기 (다운로드 없이)"""
    return send_from_directory(str(UPLOAD_FOLDER), filename, as_attachment=False)

