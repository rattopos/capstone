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

# 테스트 파일 경로 (분석표 엑셀 파일)
TEST_FILE = Path(__file__).parent / "기초자료 수집표_2025년 2분기_캡스톤_보완.xlsx"

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

class DemoRecorder:
    """데모 영상 녹화기"""
    
    def __init__(self):
        self.page = None
        self.context = None
        self.browser = None
        self.srt = SRTGenerator()
        self.start_time = None
        self.subtitle_index = 0
    
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
    
    async def wait_for_element(self, selector: str, timeout: int = 10000):
        """요소가 나타날 때까지 대기"""
        try:
            await self.page.wait_for_selector(selector, timeout=timeout)
            return True
        except:
            print(f"[경고] 요소를 찾을 수 없음: {selector}")
            return False
    
    async def safe_click(self, selector: str, description: str = ""):
        """안전하게 클릭"""
        try:
            element = self.page.locator(selector).first
            if await element.is_visible(timeout=3000):
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
        await self.wait(2)
        
        self.next_subtitle()  # 지역경제동향 보도자료 생성 시스템
        await self.wait(5)
        
        self.next_subtitle()  # 분석표 엑셀 파일을 업로드하면...
        await self.wait(5)
        
        self.next_subtitle()  # 대시보드 레이아웃
        await self.wait(4)
        
        await self.scene_transition()
    
    async def scene_2_upload(self):
        """Scene 2: 파일 업로드 (1분)"""
        print("\n[Scene 2] 파일 업로드")
        
        self.next_subtitle()  # Step 1: 분석표 엑셀 파일 업로드
        await self.wait(4)
        
        # 파일 업로드
        if TEST_FILE.exists():
            self.next_subtitle()  # 파일 업로드 중...
            
            # 파일 input 찾기
            file_input = self.page.locator('input[type="file"]')
            if await file_input.is_visible():
                await file_input.set_input_files(str(TEST_FILE))
                print(f"[업로드] 파일 업로드: {TEST_FILE.name}")
                await self.wait(5)
            else:
                print("[경고] 파일 input을 찾을 수 없음")
                self.srt.add_gap(5)
                self.subtitle_index += 1
            
            self.next_subtitle()  # ✅ 2025년 2분기 자동 감지 완료
            
            # 업로드 완료 대기 (연도/분기 자동 감지)
            await self.wait_for_element('.period-value:not(.waiting)', timeout=15000)
            await self.wait(5)
        else:
            print(f"[경고] 테스트 파일을 찾을 수 없음: {TEST_FILE}")
            self.srt.add_gap(14)
            self.subtitle_index += 2
        
        await self.scene_transition()
    
    async def scene_3_summary_preview(self):
        """Scene 3: 요약 탭 보도자료 미리보기 (1분 30초)"""
        print("\n[Scene 3] 요약 탭 보도자료 미리보기")
        
        self.next_subtitle()  # Step 2: 요약 탭 보도자료 미리보기
        await self.wait(3)
        
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
        
        # 표지 클릭
        self.next_subtitle()  # 표지
        await self.safe_click('.report-item:has-text("표지"), .report-item:has-text("표지")', "표지")
        await self.wait(PREVIEW_DELAY)
        await self.wait(5)
        
        # 인포그래픽 클릭
        self.next_subtitle()  # 인포그래픽
        await self.safe_click('.report-item:has-text("인포그래픽")', "인포그래픽")
        await self.wait(PREVIEW_DELAY)
        await self.wait(5)
        
        # 요약-지역경제동향 클릭
        self.next_subtitle()  # 요약-지역경제동향
        await self.safe_click('.report-item:has-text("요약-지역경제동향"), .report-item:has-text("지역경제동향")', "요약-지역경제동향")
        await self.wait(PREVIEW_DELAY)
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
        await self.wait(3)
        
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
        await self.wait(PREVIEW_DELAY)
        await self.wait(5)
        
        # 고용동향 클릭 (고용률 또는 실업률)
        self.next_subtitle()  # 고용동향
        await self.safe_click('.report-item:has-text("고용률"), .report-item:has-text("실업률")', "고용동향")
        await self.wait(PREVIEW_DELAY)
        await self.wait(5)
        
        await self.scene_transition()
    
    async def scene_5_regional_preview(self):
        """Scene 5: 시도별 탭 보도자료 미리보기 (1분)"""
        print("\n[Scene 5] 시도별 탭 보도자료 미리보기")
        
        self.next_subtitle()  # Step 4: 시도별 탭 보도자료 미리보기
        await self.wait(3)
        
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
        await self.wait(PREVIEW_DELAY)
        await self.wait(4)
        
        # 경기 클릭
        self.next_subtitle()  # 경기
        await self.safe_click('.report-item:has-text("경기")', "경기")
        await self.wait(PREVIEW_DELAY)
        await self.wait(4)
        
        await self.scene_transition()
    
    async def scene_6_statistics_preview(self):
        """Scene 6: 통계표 탭 미리보기 (30초)"""
        print("\n[Scene 6] 통계표 탭 미리보기")
        
        self.next_subtitle()  # Step 5: 통계표 탭 미리보기
        await self.wait(3)
        
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
        await self.wait(PREVIEW_DELAY)
        await self.wait(4)
        
        await self.scene_transition()
    
    async def scene_7_generate_and_export(self):
        """Scene 7: 전체 생성 및 내보내기 (1분)"""
        print("\n[Scene 7] 전체 생성 및 내보내기")
        
        self.next_subtitle()  # Step 6: 전체 생성 및 내보내기
        await self.wait(3)
        
        # 전체 생성 버튼 클릭
        self.next_subtitle()  # 전체 생성 버튼 클릭
        generate_btn = self.page.locator('#generateAllBtn, button:has-text("전체 생성"), button:has-text("일괄 생성")')
        if await generate_btn.first.is_visible():
            await generate_btn.first.click()
            print("[클릭] 전체 생성 버튼")
            await self.wait(4)
        else:
            print("[경고] 전체 생성 버튼을 찾을 수 없음")
            self.srt.add_gap(4)
            self.subtitle_index += 1
        
        self.next_subtitle()  # 생성 진행 상황 표시
        # 생성 진행 대기 (최대 30초)
        await self.wait(5)
        
        # 내보내기 버튼 클릭
        self.next_subtitle()  # 내보내기
        export_btn = self.page.locator('#exportBtn, button:has-text("내보내기"), button:has-text("다운로드")')
        if await export_btn.first.is_visible():
            await export_btn.first.click()
            print("[클릭] 내보내기 버튼")
            await self.wait(4)
        else:
            print("[경고] 내보내기 버튼을 찾을 수 없음")
            self.srt.add_gap(4)
            self.subtitle_index += 1
        
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

