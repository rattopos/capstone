#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vercel 서버리스 함수 래퍼
Flask 애플리케이션을 Vercel의 서버리스 함수로 실행합니다.
"""

import sys
import os
from pathlib import Path

# 프로젝트 루트를 Python 경로에 추가
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

# Vercel 환경에서는 /tmp 디렉토리를 사용하도록 설정
if os.environ.get('VERCEL'):
    # Vercel 환경 변수 설정
    os.environ['UPLOAD_FOLDER'] = '/tmp/uploads'
    os.environ['DEBUG_FOLDER'] = '/tmp/debug'
    os.environ['EXPORT_FOLDER'] = '/tmp/exports'
    
    # 임시 디렉토리 생성
    for folder in ['/tmp/uploads', '/tmp/debug', '/tmp/exports']:
        Path(folder).mkdir(parents=True, exist_ok=True)

# Flask 애플리케이션 import
from app import app

# Vercel Python 런타임은 WSGI 애플리케이션(app 객체)을 자동으로 감지합니다

