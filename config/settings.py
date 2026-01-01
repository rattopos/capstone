# -*- coding: utf-8 -*-
"""
애플리케이션 기본 설정
"""

from pathlib import Path

# 프로젝트 루트 설정
BASE_DIR = Path(__file__).parent.parent
TEMPLATES_DIR = BASE_DIR / 'templates'
UPLOAD_FOLDER = BASE_DIR / 'uploads'
DEBUG_FOLDER = BASE_DIR / 'debug'
EXPORT_FOLDER = BASE_DIR / 'exports'  # 한글 불러오기용 HTML 및 이미지 저장

# 폴더 생성
UPLOAD_FOLDER.mkdir(exist_ok=True)
DEBUG_FOLDER.mkdir(exist_ok=True)
EXPORT_FOLDER.mkdir(exist_ok=True)

# Flask 설정
SECRET_KEY = 'capstone_secret_key_2025'
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max

