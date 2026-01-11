# -*- coding: utf-8 -*-
"""
애플리케이션 기본 설정
"""

import os
import sys
from pathlib import Path


def get_base_path():
    """
    실행 환경에 따른 기본 경로 반환
    - PyInstaller 패키징 환경: sys._MEIPASS (임시 폴더)
    - 개발 환경: 소스 코드 디렉토리
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller로 패키징된 환경
        return Path(sys._MEIPASS)
    else:
        # 개발 환경
        return Path(__file__).parent.parent


def get_data_path():
    """
    사용자 데이터 저장 경로 반환
    - PyInstaller 환경: 실행 파일 옆 또는 사용자 홈 디렉토리
    - 개발 환경: 프로젝트 디렉토리
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller 환경: 실행 파일이 있는 디렉토리 사용
        return Path(sys.executable).parent
    else:
        # 개발 환경
        return Path(__file__).parent.parent


# 프로젝트 루트 설정 (정적 파일, 템플릿 등)
BASE_DIR = get_base_path()
TEMPLATES_DIR = BASE_DIR / 'templates'

# 사용자 데이터 디렉토리 (업로드, 내보내기 등)
DATA_DIR = get_data_path()

# PyInstaller 패키징 여부
IS_FROZEN = getattr(sys, 'frozen', False)

# Vercel 환경 감지 및 디렉토리 설정
# Vercel 서버리스 환경에서는 /tmp 디렉토리만 쓰기 가능
if os.environ.get('VERCEL') or os.environ.get('VERCEL_ENV'):
    # Vercel 환경: /tmp 디렉토리 사용
    UPLOAD_FOLDER = Path('/tmp/uploads')
    DEBUG_FOLDER = Path('/tmp/debug')
    EXPORT_FOLDER = Path('/tmp/exports')
elif IS_FROZEN:
    # PyInstaller 환경: 실행 파일 옆에 데이터 폴더 생성
    UPLOAD_FOLDER = DATA_DIR / 'uploads'
    DEBUG_FOLDER = DATA_DIR / 'debug'
    EXPORT_FOLDER = DATA_DIR / 'exports'
else:
    # 로컬 개발 환경: 프로젝트 디렉토리 사용
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

