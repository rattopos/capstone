"""
웹 애플리케이션 데모 비디오 자동 생성 스크립트
Playwright를 사용하여 웹 애플리케이션의 주요 기능을 자동으로 실행하고 녹화합니다.
"""

import subprocess
import time
import sys
import os
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# 프로젝트 루트 디렉토리
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "demo_output"
OUTPUT_DIR.mkdir(exist_ok=True)

# Flask 서버 설정
FLASK_HOST = "localhost"
FLASK_PORT = 8000
FLASK_URL = f"http://{FLASK_HOST}:{FLASK_PORT}"

# 비디오 출력 경로
VIDEO_OUTPUT = OUTPUT_DIR / "demo_video.mp4"


def start_flask_server():
    """Flask 서버를 백그라운드로 시작"""
    print("🚀 Flask 서버를 시작하는 중...")
    process = subprocess.Popen(
        [sys.executable, "app.py"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        cwd=BASE_DIR
    )
    
    # 서버가 시작될 때까지 대기
    import requests
    max_retries = 30
    for i in range(max_retries):
        try:
            response = requests.get(FLASK_URL, timeout=2)
            if response.status_code == 200:
                print(f"✅ Flask 서버가 시작되었습니다: {FLASK_URL}")
                return process
        except:
            time.sleep(1)
            if i % 5 == 0:
                print(f"   서버 시작 대기 중... ({i+1}/{max_retries})")
    
    raise Exception("Flask 서버를 시작할 수 없습니다.")


def stop_flask_server(process):
    """Flask 서버 종료"""
    print("\n🛑 Flask 서버를 종료하는 중...")
    process.terminate()
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        process.kill()
    print("✅ Flask 서버가 종료되었습니다.")


def wait_for_element(page, selector, timeout=10000):
    """요소가 나타날 때까지 대기"""
    try:
        page.wait_for_selector(selector, timeout=timeout)
        return True
    except PlaywrightTimeoutError:
        return False


def create_demo_video():
    """데모 비디오 생성"""
    flask_process = None
    
    try:
        # Flask 서버 시작
        flask_process = start_flask_server()
        time.sleep(2)  # 서버 안정화 대기
        
        with sync_playwright() as p:
            print("\n🎬 브라우저를 시작하고 녹화를 시작합니다...")
            
            # 브라우저 시작 (headless=False로 실제 브라우저 표시)
            browser = p.chromium.launch(
                headless=False,
                args=['--start-maximized']
            )
            
            # 컨텍스트 생성 (비디오 녹화 포함)
            context = browser.new_context(
                viewport={'width': 1920, 'height': 1080},
                record_video_path=str(VIDEO_OUTPUT),
                record_video_size={'width': 1920, 'height': 1080}
            )
            
            page = context.new_page()
            
            # 1. 메인 페이지 접속
            print("\n📄 1단계: 메인 페이지 접속")
            page.goto(FLASK_URL, wait_until='networkidle')
            time.sleep(2)
            
            # 2. 템플릿 목록 로드 대기
            print("📋 2단계: 템플릿 목록 로드")
            if wait_for_element(page, '#templateSelect option:not([value=""])', timeout=15000):
                time.sleep(1)
            
            # 3. 템플릿 선택 (광공업생산)
            print("🎯 3단계: 템플릿 선택")
            page.select_option('#templateSelect', value='광공업생산.html')
            time.sleep(1.5)
            
            # 4. 보도자료 생성 버튼 클릭
            print("⚙️ 4단계: 보도자료 생성 시작")
            if wait_for_element(page, '#processBtn:not([disabled])', timeout=5000):
                page.click('#processBtn')
                time.sleep(1)
            
            # 5. 처리 완료 대기 (최대 60초)
            print("⏳ 5단계: 처리 완료 대기 중...")
            max_wait = 60
            waited = 0
            while waited < max_wait:
                # 결과 섹션이 나타나거나 에러가 발생했는지 확인
                result_visible = page.locator('#resultSection').is_visible()
                error_visible = page.locator('#errorSection').is_visible()
                
                if result_visible or error_visible:
                    print("✅ 처리 완료!")
                    break
                
                time.sleep(1)
                waited += 1
                if waited % 10 == 0:
                    print(f"   대기 중... ({waited}초)")
            
            time.sleep(2)
            
            # 6. 미리보기 버튼 클릭 (있는 경우)
            print("👁️ 6단계: 미리보기")
            if page.locator('#previewBtn').is_visible():
                page.click('#previewBtn')
                time.sleep(3)
                # 미리보기 닫기
                if page.locator('#closePreviewBtn').is_visible():
                    page.click('#closePreviewBtn')
                    time.sleep(1)
            
            # 7. PDF 탭으로 이동
            print("📄 7단계: PDF 생성 탭 확인")
            page.click('#pdfTabBtn')
            time.sleep(2)
            
            # 8. DOCX 탭으로 이동
            print("📝 8단계: DOCX 생성 탭 확인")
            page.click('#docxTabBtn')
            time.sleep(2)
            
            # 9. 다시 HTML 탭으로 돌아가기
            print("🔄 9단계: HTML 탭으로 복귀")
            page.click('#htmlTabBtn')
            time.sleep(2)
            
            # 마지막 대기 (비디오 마무리)
            print("\n🎬 녹화를 마무리하는 중...")
            time.sleep(3)
            
            # 브라우저 종료
            context.close()
            browser.close()
            
            print(f"\n✅ 데모 비디오가 생성되었습니다: {VIDEO_OUTPUT}")
            if VIDEO_OUTPUT.exists():
                print(f"   파일 크기: {VIDEO_OUTPUT.stat().st_size / (1024*1024):.2f} MB")
            
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Flask 서버 종료
        if flask_process:
            stop_flask_server(flask_process)


def create_advanced_demo():
    """고급 데모: 여러 템플릿 테스트"""
    flask_process = None
    
    try:
        flask_process = start_flask_server()
        time.sleep(2)
        
        with sync_playwright() as p:
            print("\n🎬 고급 데모 비디오 생성 중...")
            
            browser = p.chromium.launch(headless=False, args=['--start-maximized'])
            context = browser.new_context(
                viewport={'width': 1920, 'height': 1080},
                record_video_path=str(OUTPUT_DIR / "advanced_demo.mp4"),
                record_video_size={'width': 1920, 'height': 1080}
            )
            
            page = context.new_page()
            page.goto(FLASK_URL, wait_until='networkidle')
            time.sleep(2)
            
            # 여러 템플릿 테스트
            templates_to_test = ['광공업생산.html', '고용률.html', '수출.html']
            
            for i, template in enumerate(templates_to_test, 1):
                print(f"\n📋 템플릿 {i}/{len(templates_to_test)}: {template}")
                
                # 템플릿 선택
                if wait_for_element(page, f'#templateSelect option[value="{template}"]', timeout=5000):
                    page.select_option('#templateSelect', value=template)
                    time.sleep(1.5)
                    
                    # 생성 버튼 클릭
                    if wait_for_element(page, '#processBtn:not([disabled])', timeout=5000):
                        page.click('#processBtn')
                        
                        # 완료 대기
                        waited = 0
                        while waited < 60:
                            if page.locator('#resultSection').is_visible() or page.locator('#errorSection').is_visible():
                                break
                            time.sleep(1)
                            waited += 1
                        
                        time.sleep(2)
            
            time.sleep(3)
            context.close()
            browser.close()
            
            output_file = OUTPUT_DIR / "advanced_demo.mp4"
            print(f"\n✅ 고급 데모 비디오 생성 완료: {output_file}")
            if output_file.exists():
                print(f"   파일 크기: {output_file.stat().st_size / (1024*1024):.2f} MB")
            
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if flask_process:
            stop_flask_server(flask_process)


if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='웹 애플리케이션 데모 비디오 생성')
    parser.add_argument('--advanced', action='store_true', help='고급 데모 (여러 템플릿 테스트)')
    parser.add_argument('--headless', action='store_true', help='헤드리스 모드 (비디오만 녹화)')
    
    args = parser.parse_args()
    
    print("=" * 60)
    print("🎥 웹 애플리케이션 데모 비디오 생성기")
    print("=" * 60)
    
    if args.advanced:
        create_advanced_demo()
    else:
        create_demo_video()
    
    print("\n" + "=" * 60)
    print("✨ 완료!")
    print("=" * 60)

