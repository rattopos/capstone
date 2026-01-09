#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PyInstaller 빌드 스크립트
데스크톱 앱을 Windows 실행 파일로 빌드합니다.
"""

import subprocess
import sys
import os
from pathlib import Path

# 프로젝트 루트
PROJECT_ROOT = Path(__file__).parent
DESKTOP_APP = PROJECT_ROOT / "desktop_app"


def check_dependencies():
    """의존성 확인"""
    try:
        import PyInstaller
        print(f"✅ PyInstaller {PyInstaller.__version__}")
    except ImportError:
        print("❌ PyInstaller가 설치되어 있지 않습니다.")
        print("   pip install pyinstaller")
        return False
    
    try:
        from PyQt6 import QtCore
        print(f"✅ PyQt6 {QtCore.PYQT_VERSION_STR}")
    except ImportError:
        print("❌ PyQt6가 설치되어 있지 않습니다.")
        print("   pip install PyQt6 PyQt6-WebEngine")
        return False
    
    return True


def build_exe():
    """실행 파일 빌드"""
    print("\n🔨 빌드 시작...\n")
    
    # PyInstaller 옵션
    options = [
        "pyinstaller",
        "--name=지역경제동향_생성기",
        "--onefile",
        "--windowed",
        f"--add-data={DESKTOP_APP / 'config'}:config",
        f"--add-data={PROJECT_ROOT / 'utils' / '양식.hwpx'}:templates",
        f"--add-data={PROJECT_ROOT / 'templates'}:web_templates",
        "--hidden-import=PyQt6.QtWebEngineWidgets",
        "--hidden-import=lxml.etree",
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--clean",
        "--noconfirm",
        str(DESKTOP_APP / "main.py"),
    ]
    
    # 아이콘 파일이 있으면 추가
    icon_path = DESKTOP_APP / "resources" / "icon.ico"
    if icon_path.exists():
        options.insert(3, f"--icon={icon_path}")
    
    # 빌드 실행
    result = subprocess.run(options, cwd=PROJECT_ROOT)
    
    if result.returncode == 0:
        print("\n✅ 빌드 완료!")
        print(f"   출력 위치: {PROJECT_ROOT / 'dist' / '지역경제동향_생성기.exe'}")
    else:
        print("\n❌ 빌드 실패")
        return False
    
    return True


def create_spec_file():
    """PyInstaller .spec 파일 생성"""
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 스펙 파일 - 지역경제동향 보도자료 생성기

block_cipher = None

a = Analysis(
    ['{DESKTOP_APP / "main.py"}'],
    pathex=['{PROJECT_ROOT}'],
    binaries=[],
    datas=[
        ('{DESKTOP_APP / "config"}', 'config'),
        ('{PROJECT_ROOT / "utils" / "양식.hwpx"}', 'templates'),
        ('{PROJECT_ROOT / "templates"}', 'web_templates'),
    ],
    hiddenimports=[
        'PyQt6.QtWebEngineWidgets',
        'lxml.etree',
        'lxml._elementpath',
        'pandas',
        'openpyxl',
        'jinja2',
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='지역경제동향_생성기',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
'''
    
    spec_path = PROJECT_ROOT / "desktop_app.spec"
    with open(spec_path, 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print(f"✅ 스펙 파일 생성: {spec_path}")
    return spec_path


def main():
    """메인 함수"""
    print("=" * 50)
    print("🖥️  지역경제동향 보도자료 생성기 빌드")
    print("=" * 50)
    
    # 의존성 확인
    if not check_dependencies():
        print("\n의존성을 먼저 설치하세요:")
        print("  pip install -r desktop_requirements.txt")
        sys.exit(1)
    
    # 빌드 옵션
    if len(sys.argv) > 1:
        if sys.argv[1] == "--spec":
            # 스펙 파일만 생성
            create_spec_file()
            return
        elif sys.argv[1] == "--help":
            print("\n사용법:")
            print("  python build_desktop.py         # 실행 파일 빌드")
            print("  python build_desktop.py --spec  # 스펙 파일만 생성")
            return
    
    # 빌드 실행
    build_exe()


if __name__ == "__main__":
    main()
