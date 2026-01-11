#!/bin/bash
# -*- coding: utf-8 -*-
# 지역경제동향 보도자료 생성 시스템 - macOS/Linux 빌드 스크립트

set -e

echo "================================================"
echo "  지역경제동향 보도자료 생성 시스템 - 빌드 스크립트"
echo "================================================"
echo ""

# 현재 디렉토리 확인
echo "[1/5] 프로젝트 디렉토리 확인..."
if [ ! -f "app.py" ]; then
    echo "[오류] app.py를 찾을 수 없습니다."
    echo "       프로젝트 루트 디렉토리에서 실행하세요."
    exit 1
fi
echo "      현재 디렉토리: $(pwd)"
echo ""

# 가상환경 확인 및 활성화
echo "[2/5] 가상환경 확인..."
if [ -f "VENV/bin/activate" ]; then
    echo "      가상환경 활성화: VENV"
    source VENV/bin/activate
elif [ -f "venv/bin/activate" ]; then
    echo "      가상환경 활성화: venv"
    source venv/bin/activate
else
    echo "[경고] 가상환경을 찾을 수 없습니다. 시스템 Python 사용."
fi
echo ""

# 의존성 설치 확인
echo "[3/5] 빌드 의존성 설치..."
pip install -r requirements-build.txt -q
echo "      의존성 설치 완료"
echo ""

# 이전 빌드 정리
echo "[4/5] 이전 빌드 정리..."
if [ -d "dist" ]; then
    rm -rf dist
    echo "      dist 폴더 삭제됨"
fi
if [ -d "build" ]; then
    rm -rf build
    echo "      build 폴더 삭제됨"
fi
echo ""

# PyInstaller 빌드 실행
echo "[5/5] PyInstaller 빌드 시작..."
echo "      이 작업은 몇 분 정도 걸릴 수 있습니다..."
echo ""
pyinstaller build.spec

echo ""
echo "================================================"
echo "  빌드 완료!"
echo "================================================"
echo ""
echo "  실행 파일 위치:"
if [[ "$OSTYPE" == "darwin"* ]]; then
    echo "  dist/지역경제동향_보도자료_생성기 (macOS)"
else
    echo "  dist/지역경제동향_보도자료_생성기 (Linux)"
fi
echo ""
echo "  사용 방법:"
echo "  1. dist 폴더의 실행 파일을 원하는 위치로 복사"
echo "  2. 실행 권한 부여: chmod +x 지역경제동향_보도자료_생성기"
echo "  3. 실행: ./지역경제동향_보도자료_생성기"
echo "  4. 자동으로 브라우저가 열림"
echo ""

# 실행 여부 확인
read -p "지금 바로 실행하시겠습니까? (y/n): " run_app
if [[ "$run_app" =~ ^[Yy]$ ]]; then
    echo ""
    echo "애플리케이션을 시작합니다..."
    ./dist/지역경제동향_보도자료_생성기 &
fi
