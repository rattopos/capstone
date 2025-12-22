"""
Flask 웹 애플리케이션 (모듈화 버전)
지역경제동향 보도자료 자동생성 웹 인터페이스
"""

import os
import sys
import tempfile
from pathlib import Path

from flask import Flask, render_template

# 모듈 경로 설정
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Blueprint import
from routes import templates_bp, processing_bp, export_bp, validation_bp

# Flask 앱 초기화
app = Flask(__name__, template_folder='flask_templates', static_folder='static')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['OUTPUT_FOLDER'] = tempfile.mkdtemp()

# Blueprint 등록
app.register_blueprint(templates_bp)
app.register_blueprint(processing_bp)
app.register_blueprint(export_bp)
app.register_blueprint(validation_bp)


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True, port=5000)

