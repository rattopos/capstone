#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
지역경제동향 보도자료 생성 시스템 - Qt6 데스크톱 애플리케이션
"""

import sys
from pathlib import Path

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt

from qt_ui.main_window import MainWindow
# config.settings에서 폴더가 자동으로 생성됨


def main():
    """메인 함수"""
    # 고해상도 디스플레이 지원
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    
    app = QApplication(sys.argv)
    app.setApplicationName("지역경제동향 보도자료 생성기")
    app.setOrganizationName("국가데이터처")
    
    # 메인 윈도우 생성
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
