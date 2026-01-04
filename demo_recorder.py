#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
지역경제동향 보도자료 생성 시스템 - 데모 영상 자동 녹화 스크립트

시연 계획에 따라 주요 기능을 순차적으로 시연하고 영상으로 녹화합니다.
SRT 자막 파일도 자동으로 생성됩니다.

사용법:
    1. 서버 실행: python app.py
    2. 데모 녹화: python demo_recorder.py

요구사항:
    - playwright 설치: pip install playwright && playwright install chromium
"""

import os
import time
import asyncio
from pathlib import Path
from datetime import datetime
from playwright.async_api import async_playwright


# ============================================================================
# 설정
# ============================================================================

# 서버 URL
SERVER_URL = "http://localhost:5050"

# 테스트 파일 경로 (분석표 엑셀 파일) - 메인 디렉토리 파일 사용
TEST_FILE = Path(__file__).parent / "분석표_25년 2분기_캡스톤.xlsx"

# GRDP 업로드 파일 경로 - 메인 디렉토리 파일 사용
GRDP_FILE = Path(__file__).parent / "2025년_2분기_실질_지역내총생산(잠정).xlsx"

# 출력 디렉토리
OUTPUT_DIR = Path(__file__).parent / "demo_videos"

# 영상 설정
VIDEO_WIDTH = 1920
VIDEO_HEIGHT = 1080

# 대기 시간 (초)
ACTION_DELAY = 1.5      # 일반 액션 사이 대기
SCENE_DELAY = 2.0       # 씬 전환 대기
LOAD_DELAY = 3.0        # 페이지 로딩 대기
PREVIEW_DELAY = 4.0     # 미리보기 로딩 대기


# ============================================================================
# 자막 데이터 (시연 순서별)
# ============================================================================

SUBTITLES = [
    # ========== 1. 도입부 (30초) ==========
    {
        "text": "지역경제동향 보도자료 생성 시스템\n국가데이터처 캡스톤 프로젝트",
        "duration": 5.0
    },
    {
        "text": "분석표 엑셀 파일을 업로드하면\n78페이지 보도자료를 자동으로 생성합니다",
        "duration": 5.0
    },
    {
        "text": "대시보드 레이아웃:\n좌측 사이드바, 우측 미리보기 영역",
        "duration": 4.0
    },
    
    # ========== 2. 파일 업로드 (1분) ==========
    {
        "text": "Step 1: 분석표 엑셀 파일 업로드\n드래그 앤 드롭 또는 클릭으로 업로드",
        "duration": 4.0
    },
    {
        "text": "파일 업로드 중...\n자동으로 연도/분기를 감지합니다",
        "duration": 5.0
    },
    {
        "text": "✅ 2025년 2분기 자동 감지 완료\n보도자료 항목 목록이 활성화되었습니다",
        "duration": 5.0
    },

    # ========== 2.5 GRDP 업로드 (중간 절차) ==========
    {
        "text": "Step 1.5: GRDP 파일 업로드\n'참고-GRDP' 페이지에 필요한 데이터를 결합합니다",
        "duration": 4.0
    },
    {
        "text": "✅ GRDP 추출 완료\n시도별/통계표의 GRDP 관련 내용이 활성화됩니다",
        "duration": 4.0
    },
    
    # ========== 3. 요약 탭 보도자료 미리보기 (1분 30초) ==========
    {
        "text": "Step 2: 요약 탭 보도자료 미리보기\n9개 항목, 9페이지",
        "duration": 3.0
    },
    {
        "text": "표지 - 자동 생성된 제목과 기관명\n보도자료의 첫 페이지",
        "duration": 5.0
    },
    {
        "text": "인포그래픽 - 시각화된 요약 정보\n지역별 경제 지표를 한눈에",
        "duration": 5.0
    },
    {
        "text": "요약-지역경제동향 - 자동 생성된 요약 문장\n규칙기반 자연어 생성으로 정확하고 일관된 표현",
        "duration": 5.0
    },
    {
        "text": "검토 완료 체크 - 작업 진행 상태 관리\n각 항목별 검토 상태를 추적합니다",
        "duration": 4.0
    },
    
    # ========== 4. 부문별 탭 보도자료 미리보기 (1분) ==========
    {
        "text": "Step 3: 부문별 탭 보도자료 미리보기\n10개 항목, 20페이지",
        "duration": 3.0
    },
    {
        "text": "광공업생산 - 표, 그래프, 해설문 자동 생성\n증감률과 기여도를 자동으로 계산",
        "duration": 5.0
    },
    {
        "text": "고용동향 - 고용률/실업률 데이터 및 분석문\n규칙기반 자연어 생성으로 정확한 수치 표현",
        "duration": 5.0
    },
    
    # ========== 5. 시도별 탭 보도자료 미리보기 (1분) ==========
    {
        "text": "Step 4: 시도별 탭 보도자료 미리보기\n17개 시도 + GRDP 참고, 36페이지",
        "duration": 3.0
    },
    {
        "text": "서울 - 서울 지역경제동향 미리보기\n각 시도별로 동일한 형식으로 생성",
        "duration": 4.0
    },
    {
        "text": "경기 - 다른 지역도 동일 형식으로 생성\nGRDP 데이터 연동 확인",
        "duration": 4.0
    },
    
    # ========== 6. 통계표 탭 미리보기 (30초) ==========
    {
        "text": "Step 5: 통계표 탭 미리보기\n13개 항목, 13페이지",
        "duration": 3.0
    },
    {
        "text": "광공업생산 통계표 - 자동 생성된 통계표\n엑셀 데이터를 표 형식으로 변환",
        "duration": 4.0
    },
    
    # ========== 7. 전체 생성 및 내보내기 (1분) ==========
    {
        "text": "Step 6: 전체 생성 및 내보내기\n50개 항목, 78페이지 일괄 생성",
        "duration": 3.0
    },
    {
        "text": "전체 생성 버튼 클릭\n모든 보도자료를 한 번에 생성합니다",
        "duration": 4.0
    },
    {
        "text": "생성 진행 상황 표시\n약 5분 소요 (기존 1주일 → 5분, 99.8% 단축)",
        "duration": 5.0
    },
    {
        "text": "내보내기 - HTML 파일 다운로드\n한글(HWP) 복사-붙여넣기용 HTML 생성",
        "duration": 4.0
    },
    
    # ========== 8. 한글 복사-붙여넣기 시연 (30초) ==========
    {
        "text": "Step 7: 한글(HWP) 복사-붙여넣기\n생성된 HTML에서 내용을 복사하여 한글에 붙여넣기",
        "duration": 5.0
    },
    
    # ========== 9. 마무리 (30초) ==========
    {
        "text": "지역경제동향 보도자료 생성 완료!",
        "duration": 3.0
    },
    {
        "text": "✓ 분석표 업로드 → 자동 연도/분기 감지\n✓ 50개 항목 78페이지 자동 생성\n✓ 규칙기반 정확한 수치와 일관된 표현\n✓ HTML 내보내기 → 한글 복사-붙여넣기",
        "duration": 8.0
    },
    {
        "text": "시간 절감 효과: 1주일 → 약 5분 (99.8% 단축)\n\n감사합니다!",
        "duration": 5.0
    }
]


# ============================================================================
# SRT 자막 생성기
# ============================================================================

class SRTGenerator:
    """SRT 자막 파일 생성기"""
    
    def __init__(self):
        self.entries = []
        self.current_time = 0.0
    
    def add_subtitle(self, text: str, duration: float):
        """자막 추가"""
        start_time = self.current_time
        end_time = start_time + duration
        
        self.entries.append({
            "index": len(self.entries) + 1,
            "start": start_time,
            "end": end_time,
            "text": text
        })
        
        self.current_time = end_time
    
    def add_gap(self, duration: float):
        """자막 없는 구간 추가"""
        self.current_time += duration
    
    @staticmethod
    def format_time(seconds: float) -> str:
        """초를 SRT 시간 형식으로 변환 (HH:MM:SS,mmm)"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        millis = int((seconds % 1) * 1000)
        return f"{hours:02d}:{minutes:02d}:{secs:02d},{millis:03d}"
    
    def generate(self) -> str:
        """SRT 파일 내용 생성"""
        lines = []
        for entry in self.entries:
            lines.append(str(entry["index"]))
            lines.append(f"{self.format_time(entry['start'])} --> {self.format_time(entry['end'])}")
            lines.append(entry["text"])
            lines.append("")  # 빈 줄
        return "\n".join(lines)
    
    def save(self, filepath: Path):
        """SRT 파일 저장"""
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(self.generate())
        print(f"[자막] SRT 파일 저장: {filepath}")


# ============================================================================
# 데모 녹화기
# ============================================================================

# 가상 마우스 커서 CSS/HTML
VIRTUAL_CURSOR_STYLE = """
<style id="virtual-cursor-style">
#virtual-cursor {
    position: fixed;
    width: 24px;
    height: 24px;
    pointer-events: none;
    z-index: 999999;
    transition: none;
}
#virtual-cursor .cursor-pointer {
    width: 0;
    height: 0;
    border-left: 12px solid #333;
    border-right: 12px solid transparent;
    border-bottom: 20px solid transparent;
    filter: drop-shadow(2px 2px 2px rgba(0,0,0,0.3));
    transform: rotate(-5deg);
}
#virtual-cursor .cursor-pointer::after {
    content: '';
    position: absolute;
    top: 2px;
    left: -10px;
    width: 0;
    height: 0;
    border-left: 10px solid white;
    border-right: 10px solid transparent;
    border-bottom: 16px solid transparent;
}
#virtual-cursor.clicking .cursor-pointer {
    transform: rotate(-5deg) scale(0.85);
}
#virtual-cursor .click-ripple {
    position: absolute;
    top: 0;
    left: 0;
    width: 30px;
    height: 30px;
    border-radius: 50%;
    background: rgba(59, 130, 246, 0.4);
    transform: scale(0);
    opacity: 0;
}
#virtual-cursor.clicking .click-ripple {
    animation: click-ripple 0.4s ease-out;
}
@keyframes click-ripple {
    0% { transform: scale(0); opacity: 1; }
    100% { transform: scale(2); opacity: 0; }
}
</style>
<div id="virtual-cursor">
    <div class="cursor-pointer"></div>
    <div class="click-ripple"></div>
</div>
"""

class DemoRecorder:
    """데모 영상 녹화기"""
    
    def __init__(self):
        self.page = None
        self.context = None
        self.browser = None
        self.srt = SRTGenerator()
        self.start_time = None
        self.subtitle_index = 0
        self.cursor_x = 100
        self.cursor_y = 100
    
    async def setup(self, playwright):
        """브라우저 및 녹화 설정"""
        OUTPUT_DIR.mkdir(exist_ok=True)
        
        # 브라우저 실행 (headless=False로 화면 표시)
        self.browser = await playwright.chromium.launch(
            headless=False,
            args=[
                f"--window-size={VIDEO_WIDTH},{VIDEO_HEIGHT}",
                "--disable-infobars",
                "--hide-scrollbars"
            ]
        )
        
        # 녹화 컨텍스트 생성
        self.context = await self.browser.new_context(
            viewport={"width": VIDEO_WIDTH, "height": VIDEO_HEIGHT},
            record_video_dir=str(OUTPUT_DIR),
            record_video_size={"width": VIDEO_WIDTH, "height": VIDEO_HEIGHT},
            locale="ko-KR"
        )
        
        self.page = await self.context.new_page()
        print(f"[녹화] 브라우저 설정 완료 ({VIDEO_WIDTH}x{VIDEO_HEIGHT})")
    
    async def inject_virtual_cursor(self):
        """가상 마우스 커서를 페이지에 주입"""
        await self.page.evaluate(f"""
            () => {{
                // 이미 존재하면 제거
                const existing = document.getElementById('virtual-cursor');
                if (existing) existing.remove();
                const existingStyle = document.getElementById('virtual-cursor-style');
                if (existingStyle) existingStyle.remove();
                
                // 새로 추가
                const wrapper = document.createElement('div');
                wrapper.innerHTML = `{VIRTUAL_CURSOR_STYLE}`;
                document.body.appendChild(wrapper.querySelector('style'));
                document.body.appendChild(wrapper.querySelector('#virtual-cursor'));
                
                // 초기 위치
                const cursor = document.getElementById('virtual-cursor');
                cursor.style.left = '100px';
                cursor.style.top = '100px';
            }}
        """)
        print("[커서] 가상 마우스 커서 주입 완료")
    
    async def move_cursor_to(self, x: float, y: float, duration: float = 0.5):
        """마우스 커서를 부드럽게 이동 (Bézier 곡선 애니메이션)"""
        steps = max(int(duration * 60), 10)  # 60fps 기준
        
        start_x, start_y = self.cursor_x, self.cursor_y
        
        for i in range(steps + 1):
            t = i / steps
            # ease-out 곡선 적용
            t = 1 - (1 - t) ** 3
            
            current_x = start_x + (x - start_x) * t
            current_y = start_y + (y - start_y) * t
            
            await self.page.evaluate(f"""
                () => {{
                    const cursor = document.getElementById('virtual-cursor');
                    if (cursor) {{
                        cursor.style.left = '{current_x}px';
                        cursor.style.top = '{current_y}px';
                    }}
                }}
            """)
            await asyncio.sleep(duration / steps)
        
        self.cursor_x, self.cursor_y = x, y
    
    async def click_animation(self):
        """클릭 애니메이션 표시"""
        await self.page.evaluate("""
            () => {
                const cursor = document.getElementById('virtual-cursor');
                if (cursor) {
                    cursor.classList.add('clicking');
                    setTimeout(() => cursor.classList.remove('clicking'), 400);
                }
            }
        """)
        await asyncio.sleep(0.15)
    
    async def move_to_element(self, selector: str, duration: float = 0.6):
        """요소 위치로 마우스 커서 이동"""
        try:
            box = await self.page.locator(selector).first.bounding_box()
            if box:
                # 요소 중앙으로 이동
                target_x = box['x'] + box['width'] / 2
                target_y = box['y'] + box['height'] / 2
                await self.move_cursor_to(target_x, target_y, duration)
                return True
        except Exception as e:
            print(f"[경고] 요소 위치 찾기 실패: {selector} - {e}")
        return False
    
    async def cleanup(self):
        """정리"""
        if self.page:
            await self.page.close()
        if self.context:
            await self.context.close()
        if self.browser:
            await self.browser.close()
    
    def next_subtitle(self):
        """다음 자막 추가"""
        if self.subtitle_index < len(SUBTITLES):
            sub = SUBTITLES[self.subtitle_index]
            self.srt.add_subtitle(sub["text"], sub["duration"])
            self.subtitle_index += 1
            print(f"[자막 {self.subtitle_index}/{len(SUBTITLES)}] {sub['text'][:40]}...")
    
    async def wait(self, seconds: float = ACTION_DELAY):
        """대기"""
        await asyncio.sleep(seconds)
    
    async def scene_transition(self):
        """씬 전환 대기"""
        await asyncio.sleep(SCENE_DELAY)

    async def wait_for_js_condition(self, js_condition: str, timeout_ms: int = 15000, poll_ms: int = 250):
        """JS 조건이 True가 될 때까지 폴링 대기 (실패 시 RuntimeError)"""
        start = time.time()
        while True:
            try:
                ok = await self.page.evaluate(f"() => Boolean({js_condition})")
            except Exception:
                ok = False

            if ok:
                return

            if (time.time() - start) * 1000 >= timeout_ms:
                raise RuntimeError(f"조건 대기 타임아웃: {js_condition}")
            await asyncio.sleep(poll_ms / 1000)

    async def require_visible(self, selector: str, description: str, timeout_ms: int = 8000):
        """요소가 보이는지 확인 (아니면 RuntimeError)"""
        el = self.page.locator(selector).first
        try:
            await el.wait_for(state="visible", timeout=timeout_ms)
        except Exception as e:
            raise RuntimeError(f"{description} 요소가 보이지 않습니다: {selector}") from e

    async def click_required(self, selector: str, description: str, timeout_ms: int = 8000):
        """요소 클릭 (아니면 RuntimeError) - 마우스 이동 애니메이션 포함"""
        await self.require_visible(selector, description, timeout_ms=timeout_ms)
        try:
            # 마우스를 요소로 이동
            await self.move_to_element(selector)
            await asyncio.sleep(0.2)
            # 클릭 애니메이션
            await self.click_animation()
            # 실제 클릭
            await self.page.locator(selector).first.click()
            print(f"[클릭] {description}")
        except Exception as e:
            raise RuntimeError(f"{description} 클릭 실패: {selector}") from e
    
    async def wait_for_element(self, selector: str, timeout: int = 10000):
        """요소가 나타날 때까지 대기"""
        try:
            await self.page.wait_for_selector(selector, timeout=timeout)
            return True
        except:
            print(f"[경고] 요소를 찾을 수 없음: {selector}")
            return False
    
    async def safe_click(self, selector: str, description: str = ""):
        """안전하게 클릭 - 마우스 이동 애니메이션 포함"""
        try:
            element = self.page.locator(selector).first
            if await element.is_visible(timeout=3000):
                # 마우스를 요소로 이동
                await self.move_to_element(selector)
                await asyncio.sleep(0.2)
                # 클릭 애니메이션
                await self.click_animation()
                # 실제 클릭
                await element.click()
                print(f"[클릭] {description or selector}")
                return True
            else:
                print(f"[경고] 요소가 보이지 않음: {description or selector}")
                return False
        except Exception as e:
            print(f"[경고] 클릭 실패: {description or selector} - {e}")
            return False
    
    # ========== 시연 시나리오 ==========
    
    async def scene_1_intro(self):
        """Scene 1: 도입부 (30초)"""
        print("\n[Scene 1] 도입부")
        
        # 메인 페이지 접속
        await self.page.goto(SERVER_URL)
        await self.page.wait_for_load_state("networkidle")
        await self.wait(1)
        
        # 가상 마우스 커서 주입
        await self.inject_virtual_cursor()
        await self.wait(1)
        
        self.next_subtitle()  # 지역경제동향 보도자료 생성 시스템
        # 화면 중앙 부근에서 마우스 움직임 시작
        await self.move_cursor_to(VIDEO_WIDTH / 2, VIDEO_HEIGHT / 2, 1.0)
        await self.wait(4)
        
        self.next_subtitle()  # 분석표 엑셀 파일을 업로드하면...
        # 사이드바 영역으로 마우스 이동
        await self.move_cursor_to(200, 300, 0.8)
        await self.wait(2)
        # 미리보기 영역으로 마우스 이동
        await self.move_cursor_to(VIDEO_WIDTH - 400, VIDEO_HEIGHT / 2, 0.8)
        await self.wait(2)
        
        self.next_subtitle()  # 대시보드 레이아웃
        await self.wait(4)
        
        await self.scene_transition()
    
    async def scene_2_upload(self):
        """Scene 2: 파일 업로드 (1분)"""
        print("\n[Scene 2] 파일 업로드")
        
        self.next_subtitle()  # Step 1: 분석표 엑셀 파일 업로드
        
        # 파일 업로드 영역으로 마우스 이동
        await self.move_to_element('.upload-area, #uploadArea, .file-upload', 0.8)
        await self.wait(2)
        
        # 파일 업로드 (성공해야 다음 단계 진행)
        if not TEST_FILE.exists():
            raise RuntimeError(f"분석표 파일을 찾을 수 없습니다: {TEST_FILE}")

        self.next_subtitle()  # 파일 업로드 중...
        
        # 클릭 애니메이션 (파일 선택 동작 시뮬레이션)
        await self.click_animation()
        await self.wait(0.5)

        # 메인 업로드 input은 #fileInput (dashboard.html 기준)
        await self.require_visible('#fileInput', "메인 파일 업로드 input(#fileInput)")
        await self.page.locator('#fileInput').set_input_files(str(TEST_FILE))
        print(f"[업로드] 분석표 업로드: {TEST_FILE.name}")

        # 업로드/처리 완료 조건: state.fileUploaded && state.fileType === 'analysis'
        self.next_subtitle()  # ✅ 2025년 2분기 자동 감지 완료
        await self.wait_for_js_condition("window.state && state.fileUploaded === true", timeout_ms=60000)
        await self.wait_for_js_condition("window.state && state.fileType === 'analysis'", timeout_ms=60000)
        await self.wait_for_js_condition("document.getElementById('periodValue') && !document.getElementById('periodValue').classList.contains('waiting')", timeout_ms=60000)
        
        # 감지 완료 결과 영역으로 마우스 이동
        await self.move_to_element('#periodValue, .period-value', 0.6)
        await self.wait(1.5)
        
        await self.scene_transition()

    async def scene_3_grdp_upload(self):
        """Scene 3: GRDP 업로드 (성공해야 다음 단계 진행)"""
        print("\n[Scene 3] GRDP 업로드")

        if not GRDP_FILE.exists():
            raise RuntimeError(f"GRDP 파일을 찾을 수 없습니다: {GRDP_FILE}")

        # GRDP 누락 안내가 뜨는 경우: '추가' 버튼이 있고 showGrdpModal() 호출 가능
        self.next_subtitle()  # Step 1.5: GRDP 파일 업로드
        
        # GRDP 추가 버튼 영역으로 마우스 이동 (있다면)
        try:
            await self.move_to_element('.grdp-btn, #grdpAddBtn, button:has-text("GRDP")', 0.6)
            await self.click_animation()
        except:
            pass
        await self.wait(0.5)

        await self.page.evaluate("() => { if (typeof showGrdpModal === 'function') showGrdpModal(); }")
        await self.wait_for_js_condition("document.getElementById('grdpModal') && document.getElementById('grdpModal').classList.contains('active')", timeout_ms=15000)
        
        # 모달 내 업로드 영역으로 마우스 이동
        await self.move_to_element('#grdpModalFileInput, .grdp-upload-area', 0.6)
        await self.wait(0.5)
        await self.click_animation()

        await self.require_visible('#grdpModalFileInput', "GRDP 업로드 input(#grdpModalFileInput)")
        await self.page.locator('#grdpModalFileInput').set_input_files(str(GRDP_FILE))
        print(f"[업로드] GRDP 업로드: {GRDP_FILE.name}")

        # 업로드 성공 조건: 상태 텍스트에 '✅' 및 '추출 완료' 포함 + 모달 닫힘 + grdpInfo 표시
        await self.wait_for_js_condition(
            "document.getElementById('grdpUploadStatus') && document.getElementById('grdpUploadStatus').textContent.includes('추출 완료') && document.getElementById('grdpUploadStatus').textContent.includes('✅')",
            timeout_ms=120000
        )
        await self.wait_for_js_condition(
            "!document.getElementById('grdpModal').classList.contains('active')",
            timeout_ms=30000
        )
        await self.wait_for_js_condition(
            "document.getElementById('grdpInfo') && document.getElementById('grdpInfo').style.display !== 'none'",
            timeout_ms=30000
        )
        await self.wait_for_js_condition(
            "document.getElementById('grdpNational') && document.getElementById('grdpNational').textContent.trim().length > 0",
            timeout_ms=30000
        )

        self.next_subtitle()  # ✅ GRDP 추출 완료
        # GRDP 정보 표시 영역으로 마우스 이동
        await self.move_to_element('#grdpInfo, .grdp-info', 0.6)
        await self.wait(4)

        await self.scene_transition()
    
    async def scene_3_summary_preview(self):
        """Scene 3: 요약 탭 보도자료 미리보기 (1분 30초)"""
        print("\n[Scene 3] 요약 탭 보도자료 미리보기")
        
        self.next_subtitle()  # Step 2: 요약 탭 보도자료 미리보기
        
        # 요약 탭 버튼으로 마우스 이동 및 클릭
        await self.move_to_element('.tab-btn[data-tab="summary"], .tab-item:has-text("요약"), button:has-text("요약")', 0.6)
        await self.wait(0.5)
        await self.click_animation()
        
        # 요약 탭으로 이동 (JavaScript 함수 호출)
        await self.page.evaluate("""
            if (typeof switchTab === 'function') {
                switchTab('summary');
            } else if (typeof selectGlobalReport === 'function') {
                // 요약 탭의 첫 번째 항목 찾기
                const summaryItems = document.querySelectorAll('.report-item');
                for (let i = 0; i < summaryItems.length; i++) {
                    const item = summaryItems[i];
                    if (item.textContent.includes('표지') || item.textContent.includes('요약')) {
                        selectGlobalReport(i);
                        break;
                    }
                }
            }
        """)
        await self.wait(2)
        
        # 표지 클릭 (미리보기 iframe이 실제로 채워져야 성공)
        self.next_subtitle()  # 표지
        await self.safe_click('.report-item:has-text("표지"), .report-item:has-text("표지")', "표지")
        await self.wait_for_js_condition("document.getElementById('previewIframe') && document.getElementById('previewIframe').style.display !== 'none' && (document.getElementById('previewIframe').srcdoc || '').length > 500", timeout_ms=60000)
        await self.wait(5)
        
        # 인포그래픽 클릭
        self.next_subtitle()  # 인포그래픽
        await self.safe_click('.report-item:has-text("인포그래픽")', "인포그래픽")
        await self.wait_for_js_condition("document.getElementById('previewIframe') && document.getElementById('previewIframe').style.display !== 'none' && (document.getElementById('previewIframe').srcdoc || '').length > 500", timeout_ms=60000)
        await self.wait(5)
        
        # 요약-지역경제동향 클릭
        self.next_subtitle()  # 요약-지역경제동향
        await self.safe_click('.report-item:has-text("요약-지역경제동향"), .report-item:has-text("지역경제동향")', "요약-지역경제동향")
        await self.wait_for_js_condition("document.getElementById('previewIframe') && document.getElementById('previewIframe').style.display !== 'none' && (document.getElementById('previewIframe').srcdoc || '').length > 500", timeout_ms=60000)
        await self.wait(5)
        
        # 검토 완료 체크
        self.next_subtitle()  # 검토 완료 체크
        await self.safe_click('#markReviewedBtn, button:has-text("검토완료")', "검토완료")
        await self.wait(4)
        
        await self.scene_transition()
    
    async def scene_4_sectoral_preview(self):
        """Scene 4: 부문별 탭 보도자료 미리보기 (1분)"""
        print("\n[Scene 4] 부문별 탭 보도자료 미리보기")
        
        self.next_subtitle()  # Step 3: 부문별 탭 보도자료 미리보기
        
        # 부문별 탭 버튼으로 마우스 이동 및 클릭
        await self.move_to_element('.tab-btn[data-tab="sectoral"], .tab-item:has-text("부문별"), button:has-text("부문별")', 0.6)
        await self.wait(0.5)
        await self.click_animation()
        
        # 부문별 탭으로 이동
        await self.page.evaluate("""
            if (typeof switchTab === 'function') {
                switchTab('sectoral');
            }
        """)
        await self.wait(2)
        
        # 광공업생산 클릭
        self.next_subtitle()  # 광공업생산
        await self.safe_click('.report-item:has-text("광공업생산")', "광공업생산")
        await self.wait_for_js_condition("document.getElementById('previewIframe') && document.getElementById('previewIframe').style.display !== 'none' && (document.getElementById('previewIframe').srcdoc || '').length > 500", timeout_ms=60000)
        await self.wait(5)
        
        # 고용동향 클릭 (고용률 또는 실업률)
        self.next_subtitle()  # 고용동향
        await self.safe_click('.report-item:has-text("고용률"), .report-item:has-text("실업률")', "고용동향")
        await self.wait_for_js_condition("document.getElementById('previewIframe') && document.getElementById('previewIframe').style.display !== 'none' && (document.getElementById('previewIframe').srcdoc || '').length > 500", timeout_ms=60000)
        await self.wait(5)
        
        await self.scene_transition()
    
    async def scene_5_regional_preview(self):
        """Scene 5: 시도별 탭 보도자료 미리보기 (1분)"""
        print("\n[Scene 5] 시도별 탭 보도자료 미리보기")
        
        self.next_subtitle()  # Step 4: 시도별 탭 보도자료 미리보기
        
        # 시도별 탭 버튼으로 마우스 이동 및 클릭
        await self.move_to_element('.tab-btn[data-tab="regional"], .tab-item:has-text("시도별"), button:has-text("시도별")', 0.6)
        await self.wait(0.5)
        await self.click_animation()
        
        # 시도별 탭으로 이동
        await self.page.evaluate("""
            if (typeof switchTab === 'function') {
                switchTab('regional');
            }
        """)
        await self.wait(2)
        
        # 서울 클릭
        self.next_subtitle()  # 서울
        await self.safe_click('.report-item:has-text("서울")', "서울")
        await self.wait_for_js_condition("document.getElementById('previewIframe') && document.getElementById('previewIframe').style.display !== 'none' && (document.getElementById('previewIframe').srcdoc || '').length > 500", timeout_ms=60000)
        await self.wait(4)
        
        # 경기 클릭
        self.next_subtitle()  # 경기
        await self.safe_click('.report-item:has-text("경기")', "경기")
        await self.wait_for_js_condition("document.getElementById('previewIframe') && document.getElementById('previewIframe').style.display !== 'none' && (document.getElementById('previewIframe').srcdoc || '').length > 500", timeout_ms=60000)
        await self.wait(4)
        
        await self.scene_transition()
    
    async def scene_6_statistics_preview(self):
        """Scene 6: 통계표 탭 미리보기 (30초)"""
        print("\n[Scene 6] 통계표 탭 미리보기")
        
        self.next_subtitle()  # Step 5: 통계표 탭 미리보기
        
        # 통계표 탭 버튼으로 마우스 이동 및 클릭
        await self.move_to_element('.tab-btn[data-tab="statistics"], .tab-item:has-text("통계표"), button:has-text("통계표")', 0.6)
        await self.wait(0.5)
        await self.click_animation()
        
        # 통계표 탭으로 이동
        await self.page.evaluate("""
            if (typeof switchTab === 'function') {
                switchTab('statistics');
            }
        """)
        await self.wait(2)
        
        # 광공업생산 통계표 클릭
        self.next_subtitle()  # 광공업생산 통계표
        await self.safe_click('.report-item:has-text("통계표-광공업생산"), .report-item:has-text("광공업생산지수")', "광공업생산 통계표")
        await self.wait_for_js_condition("document.getElementById('previewIframe') && document.getElementById('previewIframe').style.display !== 'none' && (document.getElementById('previewIframe').srcdoc || '').length > 500", timeout_ms=60000)
        await self.wait(4)
        
        await self.scene_transition()
    
    async def scene_7_generate_and_export(self):
        """Scene 7: 전체 생성 및 내보내기 (1분)"""
        print("\n[Scene 7] 전체 생성 및 내보내기")
        
        self.next_subtitle()  # Step 6: 전체 생성 및 내보내기
        
        # 전체 생성 버튼으로 마우스 이동
        await self.move_to_element('#generateAllBtn, button:has-text("전체 생성"), .generate-all-btn', 0.8)
        await self.wait(1)
        
        # 전체 미리보기 생성 (dashboard.html의 generateAllReports() 사용)
        self.next_subtitle()  # 전체 생성 버튼 클릭
        await self.click_animation()
        await self.page.evaluate("() => { if (typeof generateAllReports === 'function') generateAllReports(); }")

        # 성공 조건(엄격): allGenerated=true AND generationStats.completed == 전체 항목 수
        self.next_subtitle()  # 생성 진행 상황 표시
        await self.wait_for_js_condition("window.state && state.allGenerated === true", timeout_ms=12 * 60 * 1000)
        await self.wait_for_js_condition(
            """(() => {
                const total =
                  (state.summaryReports?.length || 0) +
                  (state.sectoralReports?.length || 0) +
                  (state.regionalReports?.length || 0) +
                  (state.statisticsReports?.length || 0);
                return state.generationStats?.completed === total && total > 0;
            })()""",
            timeout_ms=12 * 60 * 1000
        )

        # 내보내기: 파일 저장 다이얼로그(권한/사용자 입력)가 필요 없는 '프로젝트 폴더 저장' 버튼 사용
        self.next_subtitle()  # 내보내기 - HTML 파일 다운로드
        await self.click_required('#saveHtmlToProjectBtn', "💾 HTML 저장(프로젝트 폴더)")
        # 로딩 오버레이가 끝나야 성공
        await self.wait_for_js_condition(
            "document.getElementById('loadingOverlay') && document.getElementById('loadingOverlay').style.display === 'none'",
            timeout_ms=5 * 60 * 1000
        )
        
        await self.scene_transition()
    
    async def scene_8_hwp_copy_paste(self):
        """Scene 8: 한글 복사-붙여넣기 시연 (30초)"""
        print("\n[Scene 8] 한글 복사-붙여넣기 시연")
        
        self.next_subtitle()  # Step 7: 한글(HWP) 복사-붙여넣기
        await self.wait(5)
        
        await self.scene_transition()
    
    async def scene_9_finale(self):
        """Scene 9: 마무리 (30초)"""
        print("\n[Scene 9] 마무리")
        
        self.next_subtitle()  # 지역경제동향 보도자료 생성 완료!
        await self.wait(3)
        
        self.next_subtitle()  # ✓ 분석표 업로드 → 자동 연도/분기 감지...
        await self.wait(8)
        
        self.next_subtitle()  # 시간 절감 효과: 1주일 → 약 5분...
        await self.wait(5)
    
    async def record(self):
        """전체 데모 녹화"""
        print("=" * 60)
        print("데모 영상 녹화 시작")
        print("=" * 60)
        print(f"서버 URL: {SERVER_URL}")
        print(f"테스트 파일: {TEST_FILE}")
        print(f"출력 디렉토리: {OUTPUT_DIR}")
        print("=" * 60)
        
        self.start_time = time.time()
        
        # 각 씬 실행
        await self.scene_1_intro()
        await self.scene_2_upload()
        await self.scene_3_grdp_upload()
        await self.scene_3_summary_preview()
        await self.scene_4_sectoral_preview()
        await self.scene_5_regional_preview()
        await self.scene_6_statistics_preview()
        await self.scene_7_generate_and_export()
        await self.scene_8_hwp_copy_paste()
        await self.scene_9_finale()
        
        # 마지막 대기
        await self.wait(3)
        
        elapsed = time.time() - self.start_time
        print("=" * 60)
        print(f"녹화 완료! 총 {elapsed:.1f}초 ({elapsed/60:.1f}분)")
        print("=" * 60)


async def main():
    """메인 함수"""
    # 출력 디렉토리 생성
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    async with async_playwright() as playwright:
        recorder = DemoRecorder()
        
        try:
            await recorder.setup(playwright)
            await recorder.record()
            
            # SRT 자막 저장
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            srt_path = OUTPUT_DIR / f"demo_subtitles_{timestamp}.srt"
            recorder.srt.save(srt_path)
            
            print(f"\n[완료] 영상 파일 위치: {OUTPUT_DIR}")
            print(f"[완료] 자막 파일: {srt_path}")
            print("\n[참고] Playwright가 생성한 영상 파일은 .webm 형식입니다.")
            print("      필요시 FFmpeg로 MP4로 변환할 수 있습니다.")
            
        except Exception as e:
            print(f"[오류] 녹화 중 오류 발생: {e}")
            import traceback
            traceback.print_exc()
        finally:
            await recorder.cleanup()


if __name__ == "__main__":
    # 서버 실행 확인 안내
    print("\n" + "=" * 60)
    print("📹 지역경제동향 보도자료 생성 시스템 - 데모 녹화")
    print("=" * 60)
    print("\n⚠️  녹화 전 확인사항:")
    print(f"  1. 서버가 실행 중인지 확인: {SERVER_URL}")
    print(f"  2. 테스트 파일 존재 확인: {TEST_FILE}")
    print("\n서버 실행 명령: python app.py")
    print("=" * 60)
    
    # 사용자 확인
    try:
        input("\n준비가 되면 Enter를 눌러 녹화를 시작하세요...")
    except EOFError:
        pass  # 비대화형 환경에서는 바로 시작
    
    # 녹화 시작
    asyncio.run(main())

