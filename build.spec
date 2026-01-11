# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 빌드 설정 파일
지역경제동향 보도자료 생성 시스템

빌드 명령어:
    pyinstaller build.spec

결과물:
    dist/지역경제동향_보도자료_생성기.exe (Windows)
"""

import os
import sys
from pathlib import Path

# 프로젝트 루트 디렉토리
project_root = Path(os.getcwd())

block_cipher = None

# 데이터 파일 수집 (정적 파일, 템플릿 등)
datas = [
    # HTML 템플릿
    ('templates', 'templates'),
    ('dashboard.html', '.'),
    
    # 참고 이미지 (CI, 로고 등)
    ('correct_answer/MODS_MI_2025', 'correct_answer/MODS_MI_2025'),
    ('correct_answer/MODS_MI_sub_2025', 'correct_answer/MODS_MI_sub_2025'),
    
    # 설정 파일
    ('config', 'config'),
    
    # 서비스 모듈
    ('services', 'services'),
    ('routes', 'routes'),
    ('utils', 'utils'),
    ('extractors', 'extractors'),
]

# 숨겨진 imports (동적으로 로드되는 모듈)
hiddenimports = [
    # Flask 관련
    'flask',
    'flask.json',
    'werkzeug',
    'werkzeug.serving',
    'werkzeug.debug',
    'jinja2',
    'markupsafe',
    
    # 데이터 처리
    'pandas',
    'pandas._libs',
    'pandas._libs.tslibs',
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.np_datetime',
    'pandas.io.formats.style',
    'numpy',
    'openpyxl',
    'openpyxl.workbook',
    'openpyxl.worksheet',
    'openpyxl.cell',
    
    # HTML/이미지 처리
    'beautifulsoup4',
    'bs4',
    'PIL',
    'PIL.Image',
    
    # 날짜 처리
    'dateutil',
    'dateutil.parser',
    'pytz',
    
    # 프로젝트 모듈
    'config',
    'config.settings',
    'config.reports',
    'services',
    'services.report_generator',
    'services.excel_processor',
    'services.summary_data',
    'services.grdp_service',
    'routes',
    'routes.main',
    'routes.api',
    'routes.preview',
    'routes.debug',
    'utils',
    'utils.filters',
    'utils.excel_utils',
    'utils.data_utils',
    'extractors',
    'extractors.base',
    'extractors.config',
    'extractors.production',
    'extractors.consumption',
    'extractors.trade',
    'extractors.price',
    'extractors.employment',
    'extractors.facade',
]

# 제외할 모듈 (빌드 크기 최적화)
excludes = [
    'playwright',
    'xlwings',
    'tkinter',
    'PyQt5',
    'PyQt6',
    'PySide2',
    'PySide6',
    'matplotlib',
    'scipy',
    'sklearn',
    'tensorflow',
    'torch',
    'IPython',
    'notebook',
    'jupyter',
    'pytest',
    'unittest',
]

a = Analysis(
    ['app.py'],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher,
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='지역경제동향_보도자료_생성기',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # UPX 압축 사용 (파일 크기 감소)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 모드 (콘솔 창 숨김)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='resources/icon.ico',  # 아이콘 파일 (있으면 활성화)
)
