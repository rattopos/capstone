@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ================================================
echo   지역경제동향 보도자료 생성 시스템 - 빌드 스크립트
echo ================================================
echo.

REM 현재 디렉토리 확인
echo [1/5] 프로젝트 디렉토리 확인...
if not exist "app.py" (
    echo [오류] app.py를 찾을 수 없습니다.
    echo        프로젝트 루트 디렉토리에서 실행하세요.
    pause
    exit /b 1
)
echo       현재 디렉토리: %cd%
echo.

REM 가상환경 확인 및 활성화
echo [2/5] 가상환경 확인...
if exist "VENV\Scripts\activate.bat" (
    echo       가상환경 활성화: VENV
    call VENV\Scripts\activate.bat
) else if exist "venv\Scripts\activate.bat" (
    echo       가상환경 활성화: venv
    call venv\Scripts\activate.bat
) else (
    echo [경고] 가상환경을 찾을 수 없습니다. 시스템 Python 사용.
)
echo.

REM 의존성 설치 확인
echo [3/5] 빌드 의존성 설치...
pip install -r requirements-build.txt -q
if errorlevel 1 (
    echo [오류] 의존성 설치 실패
    pause
    exit /b 1
)
echo       의존성 설치 완료
echo.

REM 이전 빌드 정리
echo [4/5] 이전 빌드 정리...
if exist "dist" (
    rmdir /s /q dist
    echo       dist 폴더 삭제됨
)
if exist "build" (
    rmdir /s /q build
    echo       build 폴더 삭제됨
)
echo.

REM PyInstaller 빌드 실행
echo [5/5] PyInstaller 빌드 시작...
echo       이 작업은 몇 분 정도 걸릴 수 있습니다...
echo.
pyinstaller build.spec

if errorlevel 1 (
    echo.
    echo ================================================
    echo [오류] 빌드 실패!
    echo ================================================
    pause
    exit /b 1
)

echo.
echo ================================================
echo   빌드 완료!
echo ================================================
echo.
echo   실행 파일 위치:
echo   dist\지역경제동향_보도자료_생성기.exe
echo.
echo   사용 방법:
echo   1. dist 폴더의 exe 파일을 원하는 위치로 복사
echo   2. 더블클릭하여 실행
echo   3. 자동으로 브라우저가 열림
echo.
echo   배포 시:
echo   - exe 파일만 배포하면 됩니다.
echo   - Python 설치 불필요
echo.

REM 실행 여부 확인
set /p run_app="지금 바로 실행하시겠습니까? (Y/N): "
if /i "%run_app%"=="Y" (
    echo.
    echo 애플리케이션을 시작합니다...
    start "" "dist\지역경제동향_보도자료_생성기.exe"
)

pause
