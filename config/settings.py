# -*- coding: utf-8 -*-
"""
애플리케이션 기본 설정
"""

import os
from pathlib import Path

# 프로젝트 루트 설정
BASE_DIR = Path(__file__).parent.parent
TEMPLATES_DIR = BASE_DIR / 'templates'

# Vercel 환경 감지 및 디렉토리 설정
# Vercel 서버리스 환경에서는 /tmp 디렉토리만 쓰기 가능
if os.environ.get('VERCEL') or os.environ.get('VERCEL_ENV'):
    # Vercel 환경: /tmp 디렉토리 사용
    UPLOAD_FOLDER = Path('/tmp/uploads')
    DEBUG_FOLDER = Path('/tmp/debug')
    EXPORT_FOLDER = Path('/tmp/exports')
else:
    # 로컬 환경: 프로젝트 디렉토리 사용
    UPLOAD_FOLDER = BASE_DIR / 'uploads'
    DEBUG_FOLDER = BASE_DIR / 'debug'
    EXPORT_FOLDER = BASE_DIR / 'exports'  # 한글 불러오기용 HTML 및 이미지 저장

# 폴더 생성
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
DEBUG_FOLDER.mkdir(parents=True, exist_ok=True)
EXPORT_FOLDER.mkdir(parents=True, exist_ok=True)

# Flask 설정
SECRET_KEY = 'capstone_secret_key_2025'
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max

