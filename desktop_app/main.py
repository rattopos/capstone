# -*- coding: utf-8 -*-
"""
지역경제동향 보도자료 생성기 - 데스크톱 앱 진입점
"""

import sys
import os

# 상위 디렉토리를 경로에 추가 (기존 모듈 임포트용)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from desktop_app.main_window import main

if __name__ == "__main__":
    main()
