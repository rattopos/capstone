# -*- coding: utf-8 -*-
"""
애플리케이션 기본 설정
"""

import os
import sys
from pathlib import Path

# 프로젝트 루트 설정
BASE_DIR = Path(__file__).parent.parent
TEMPLATES_DIR = BASE_DIR / 'templates'

# 경로 설정
APP_DATA_DIR = BASE_DIR
USER_DOCUMENTS_DIR = BASE_DIR

# 디렉토리 설정
UPLOAD_FOLDER = BASE_DIR / 'uploads'
DEBUG_FOLDER = BASE_DIR / 'debug'
EXPORT_FOLDER = BASE_DIR / 'exports'  # 한글 불러오기용 HTML 및 이미지 저장
TEMP_DIR = EXPORT_FOLDER / '_temp'
TEMP_OUTPUT_DIR = TEMP_DIR / 'output'
TEMP_REGIONAL_OUTPUT_DIR = TEMP_DIR / 'regional_output'
TEMP_CALCULATED_DIR = TEMP_DIR / 'calculated'

# 폴더 생성
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
DEBUG_FOLDER.mkdir(parents=True, exist_ok=True)
EXPORT_FOLDER.mkdir(parents=True, exist_ok=True)
TEMP_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
TEMP_REGIONAL_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
TEMP_CALCULATED_DIR.mkdir(parents=True, exist_ok=True)

# Flask 설정
SECRET_KEY = 'capstone_secret_key_2025'
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max

