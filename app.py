#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
지역경제동향 보도자료 웹 애플리케이션
Flask 기반 대시보드로 기초자료 수집표를 업로드하고 보도자료를 생성합니다.

모듈 구조:
- config/     : 설정 및 상수 (보도자료 정의, 경로 설정)
- utils/      : 유틸리티 함수 (필터, 엑셀 처리, 데이터 처리)
- services/   : 비즈니스 로직 (보도자료 생성, GRDP 처리, 요약 데이터)
- routes/     : API 라우트 (메인, API, 미리보기)
"""

import sys
import webbrowser
import threading

from flask import Flask

from config.settings import BASE_DIR, SECRET_KEY, MAX_CONTENT_LENGTH, UPLOAD_FOLDER, IS_FROZEN
from utils.filters import register_filters
from routes import main_bp, api_bp, preview_bp, debug_bp


def create_app():
    """Flask 애플리케이션 팩토리"""
    app = Flask(
        __name__, 
            template_folder=str(BASE_DIR),
        static_folder=str(BASE_DIR)
    )
    
    # 설정
    app.secret_key = SECRET_KEY
    app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)
    app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH
    
    # Jinja2 커스텀 필터 등록
    register_filters(app)
    
    # Blueprint 등록
    app.register_blueprint(main_bp)
    app.register_blueprint(api_bp)
    app.register_blueprint(preview_bp)
    app.register_blueprint(debug_bp)
    
    return app


# 애플리케이션 인스턴스 생성
app = create_app()


def open_browser():
    """기본 브라우저에서 애플리케이션 열기"""
    webbrowser.open('http://localhost:5050')


if __name__ == '__main__':
    print("=" * 50)
    print("지역경제동향 보도자료 생성 시스템")
    print("=" * 50)
    print(f"서버 시작: http://localhost:5050")
    print("=" * 50)
    
    if IS_FROZEN:
        # PyInstaller 패키징 환경: Production 모드
        # 1.5초 후 브라우저 자동 실행 (서버 시작 대기)
        threading.Timer(1.5, open_browser).start()
        print("브라우저가 자동으로 열립니다...")
        print("종료하려면 이 창을 닫거나 Ctrl+C를 누르세요.")
        app.run(debug=False, host='127.0.0.1', port=5050, threaded=True)
    else:
        # 개발 환경: Debug 모드
        app.run(debug=True, host='0.0.0.0', port=5050)
