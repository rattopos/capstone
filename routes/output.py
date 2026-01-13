# -*- coding: utf-8 -*-
"""
Output 폴더 정적 파일 서빙 라우트
"""

from flask import Blueprint, send_from_directory
from config.config import Config

output_bp = Blueprint('output', __name__, url_prefix='/output')


@output_bp.route('/<filename>')
def serve_output_file(filename):
    """output 폴더의 파일 제공"""
    return send_from_directory(str(Config.OUTPUT_FOLDER), filename)
