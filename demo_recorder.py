#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
지역경제동향 보도자료 생성 시스템 - 데모 영상 자동 녹화 스크립트

Playwright를 사용하여 전체 기능을 시연하고 영상으로 녹화합니다.
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

# 테스트 파일 경로
TEST_FILE = Path(__file__).parent / "기초자료 수집표_2025년 2분기_캡스톤_보완.xlsx"

# 출력 디렉토리
OUTPUT_DIR = Path(__file__).parent / "demo_videos"

# 영상 설정
VIDEO_WIDTH = 1920
VIDEO_HEIGHT = 1080

# 대기 시간 (초) - 각 액션 사이의 대기 시간
ACTION_DELAY = 1.5  # 일반 액션
SCENE_DELAY = 2.0   # 씬 전환

# ============================================================================
# 자막 데이터 (Scene별)
# ============================================================================

SUBTITLES = [
    # ========== Scene 1: 시스템 소개 ==========
    {
        "text": "지역경제동향 보도자료 생성 시스템\n국가데이터처 캡스톤 프로젝트",
        "duration": 5.0
    },
    {
        "text": "[핵심 요구사항]\n보도자료 생성 자동화 → 시간/인력 절약",
        "duration": 5.0
    },
    
    # ========== Scene 2: 기초자료 업로드 ==========
    {
        "text": "[요구사항 반영] 기초자료 → 분석표 자동화\nStep 1: 기초자료 수집표 업로드",
        "duration": 3.0
    },
    {
        "text": "드래그 앤 드롭으로 간편한 파일 업로드",
        "duration": 5.0
    },
    {
        "text": "연도/분기 자동 감지 → 2025년 2분기",
        "duration": 4.0
    },
    
    # ========== Scene 3: 가중치 설정 ==========
    {
        "text": "[기술적 차별화 #1] 가중치 조절 기능\n결측치 대체값 설정",
        "duration": 3.0
    },
    {
        "text": "광공업/서비스업 가중치 개별 설정 가능",
        "duration": 5.0
    },
    
    # ========== Scene 4: 담당자 설정 ==========
    {
        "text": "[기술적 차별화 #2] 담당자 정보 설정\n보도자료에 자동 반영",
        "duration": 3.0
    },
    {
        "text": "배포일시, 배포부서, 담당자 정보 입력",
        "duration": 5.0
    },
    
    # ========== Scene 5: 분석표 다운로드 ==========
    {
        "text": "[요구사항 반영] 기초자료 → 분석표 자동 변환\nStep 4: 분석표 다운로드",
        "duration": 3.0
    },
    {
        "text": "수식 계산 포함된 분석표 엑셀 생성",
        "duration": 4.0
    },
    
    # ========== Scene 6: GRDP 설정 ==========
    {
        "text": "Step 5: GRDP 데이터 결합\nKOSIS 데이터 연동",
        "duration": 3.0
    },
    {
        "text": "GRDP 파일 업로드 또는 기본값 사용",
        "duration": 5.0
    },
    
    # ========== Scene 7: 보도자료 미리보기 ==========
    {
        "text": "[핵심 기능] 보도자료 미리보기\n실시간 렌더링",
        "duration": 3.0
    },
    {
        "text": "부문별 보도자료: 광공업, 서비스업, 고용률, 물가 등",
        "duration": 5.0
    },
    {
        "text": "시도별 보도자료: 17개 시도 경제동향",
        "duration": 5.0
    },
    {
        "text": "[기술적 차별화 #3] 인포그래픽\n지역별 지도 시각화",
        "duration": 5.0
    },
    {
        "text": "[기술적 차별화 #4] 차트 크기 조절\n슬라이더로 실시간 조정",
        "duration": 4.0
    },
    
    # ========== Scene 8: 검토 기능 ==========
    {
        "text": "[기술적 차별화 #5] 검토 기능\n작업 진행 상태 관리",
        "duration": 3.0
    },
    {
        "text": "검토완료 체크로 진행률 한눈에 파악",
        "duration": 4.0
    },
    
    # ========== Scene 9: 편집 기능 ==========
    {
        "text": "[기술적 차별화 #6] 편집 기능\n보도자료 내용 직접 수정",
        "duration": 3.0
    },
    {
        "text": "미리보기 화면에서 바로 편집 가능",
        "duration": 5.0
    },
    
    # ========== Scene 10: 내보내기 ==========
    {
        "text": "[핵심 기능] 보도자료 내보내기\n다양한 출력 형식 지원",
        "duration": 3.0
    },
    {
        "text": "PDF용 HTML / 한글 복붙용 HTML\n즉시 활용 가능",
        "duration": 5.0
    },
    
    # ========== 마무리 ==========
    {
        "text": "지역경제동향 보도자료 생성 완료!\n\n✓ 기초자료 → 분석표 자동화\n✓ 보도자료 생성 시간 단축",
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
            print(f"[자막 {self.subtitle_index}] {sub['text'][:30]}...")
    
    async def wait(self, seconds: float = ACTION_DELAY):
        """대기 (영상에서 액션 확인용)"""
        await asyncio.sleep(seconds)
    
    async def scene_transition(self):
        """씬 전환 대기"""
        await asyncio.sleep(SCENE_DELAY)
    
    # ========== 데모 시나리오 ==========
    
    async def scene_1_intro(self):
        """Scene 1: 시스템 소개"""
        print("\n[Scene 1] 시스템 소개")
        
        # 메인 페이지 접속
        await self.page.goto(SERVER_URL)
        await self.page.wait_for_load_state("networkidle")
        await self.wait(2)
        
        self.next_subtitle()  # 지역경제동향 보도자료 생성 시스템
        await self.wait(5)
        
        self.next_subtitle()  # 보도자료 생성 자동화
        await self.wait(5)
        
        await self.scene_transition()
    
    async def scene_2_upload(self):
        """Scene 2: 기초자료 업로드"""
        print("\n[Scene 2] 기초자료 업로드")
        
        self.next_subtitle()  # Step 1: 기초자료 수집표 업로드
        await self.wait(3)
        
        # 파일 업로드
        if TEST_FILE.exists():
            self.next_subtitle()  # 파일을 드래그하거나 클릭하여 업로드
            
            # 파일 input 찾기
            file_input = self.page.locator('input[type="file"]')
            await file_input.set_input_files(str(TEST_FILE))
            
            await self.wait(5)
            
            self.next_subtitle()  # 연도/분기 자동 감지
            await self.wait(4)
        else:
            print(f"[경고] 테스트 파일을 찾을 수 없음: {TEST_FILE}")
            self.srt.add_gap(9)  # 자막 건너뛰기
            self.subtitle_index += 2
        
        await self.scene_transition()
    
    async def scene_3_weight_settings(self):
        """Scene 3: 가중치 설정"""
        print("\n[Scene 3] 가중치 설정")
        
        self.next_subtitle()  # Step 2: 가중치 결측치 설정
        await self.wait(3)
        
        # 가중치 설정 버튼 클릭
        weight_btn = self.page.locator('#weightInfoBtn')
        if await weight_btn.is_visible():
            await weight_btn.click()
            await self.wait(1)
            
            self.next_subtitle()  # 광공업/서비스업 가중치 기본값 설정
            
            # 값 입력 (기본값 유지 또는 변경)
            mining_input = self.page.locator('#miningDefaultWeight')
            if await mining_input.is_visible():
                await mining_input.fill("1.0")
                await self.wait(1)
            
            service_input = self.page.locator('#serviceDefaultWeight')
            if await service_input.is_visible():
                await service_input.fill("1.0")
                await self.wait(1)
            
            # 저장 버튼 클릭
            save_btn = self.page.locator('button:has-text("저장")')
            if await save_btn.first.is_visible():
                await save_btn.first.click()
            
            await self.wait(3)
        else:
            self.srt.add_gap(5)
            self.subtitle_index += 1
        
        await self.scene_transition()
    
    async def scene_4_contact_info(self):
        """Scene 4: 담당자 정보 설정"""
        print("\n[Scene 4] 담당자 정보 설정")
        
        self.next_subtitle()  # Step 3: 담당자 정보 설정
        await self.wait(3)
        
        # 담당자 설정 버튼 클릭
        contact_btn = self.page.locator('#contactInfoBtn')
        if await contact_btn.is_visible():
            await contact_btn.click()
            await self.wait(1)
            
            self.next_subtitle()  # 배포일시, 배포부서, 담당자 정보 입력
            
            # 정보 입력
            dept_input = self.page.locator('#releaseDepartment')
            if await dept_input.is_visible():
                await dept_input.fill("국가데이터처 통계분석과")
                await self.wait(0.5)
            
            person_input = self.page.locator('#releasePerson')
            if await person_input.is_visible():
                await person_input.fill("김담당")
                await self.wait(0.5)
            
            # 저장
            save_btn = self.page.locator('button:has-text("저장")')
            if await save_btn.first.is_visible():
                await save_btn.first.click()
            
            await self.wait(3)
        else:
            self.srt.add_gap(5)
            self.subtitle_index += 1
        
        await self.scene_transition()
    
    async def scene_5_download_analysis(self):
        """Scene 5: 분석표 다운로드"""
        print("\n[Scene 5] 분석표 다운로드")
        
        self.next_subtitle()  # Step 4: 분석표 자동 생성
        await self.wait(3)
        
        # 분석표 다운로드 버튼 클릭
        download_btn = self.page.locator('#downloadAnalysisBtn')
        if await download_btn.is_visible() and await download_btn.is_enabled():
            self.next_subtitle()  # 기초자료 → 분석표 자동 변환
            await download_btn.click()
            await self.wait(4)
        else:
            self.srt.add_gap(4)
            self.subtitle_index += 1
        
        await self.scene_transition()
    
    async def scene_6_grdp_settings(self):
        """Scene 6: GRDP 설정"""
        print("\n[Scene 6] GRDP 설정")
        
        self.next_subtitle()  # Step 5: GRDP 데이터 설정
        await self.wait(3)
        
        # GRDP 모달이 자동으로 열리거나, 버튼 클릭
        grdp_modal = self.page.locator('#grdpModal')
        
        # GRDP 설정 버튼 찾기 시도
        grdp_btn = self.page.locator('button:has-text("GRDP")')
        if await grdp_btn.first.is_visible():
            await grdp_btn.first.click()
            await self.wait(1)
        
        self.next_subtitle()  # KOSIS GRDP 파일 업로드 또는 기본값 사용
        
        # 기본값 사용 버튼 클릭
        default_btn = self.page.locator('button:has-text("기본값")')
        if await default_btn.first.is_visible():
            await default_btn.first.click()
            await self.wait(3)
        
        await self.wait(2)
        await self.scene_transition()
    
    async def scene_7_preview(self):
        """Scene 7: 보도자료 미리보기"""
        print("\n[Scene 7] 보도자료 미리보기")
        
        self.next_subtitle()  # Step 6: 보도자료 미리보기
        await self.wait(3)
        
        # 부문별 탭 클릭
        sectoral_tab = self.page.locator('[data-tab="sectoral"]')
        if await sectoral_tab.is_visible():
            await sectoral_tab.click()
            await self.wait(1)
            
            self.next_subtitle()  # 부문별 보도자료
            
            # 첫 번째 보도자료 클릭
            first_report = self.page.locator('.report-list .report-item').first
            if await first_report.is_visible():
                await first_report.click()
                await self.wait(5)
        else:
            self.srt.add_gap(5)
            self.subtitle_index += 1
        
        # 시도별 탭 클릭
        regional_tab = self.page.locator('[data-tab="regional"]')
        if await regional_tab.is_visible():
            await regional_tab.click()
            await self.wait(1)
            
            self.next_subtitle()  # 시도별 보도자료
            
            # 서울 클릭
            seoul_report = self.page.locator('.report-item:has-text("서울")')
            if await seoul_report.first.is_visible():
                await seoul_report.first.click()
                await self.wait(5)
        else:
            self.srt.add_gap(5)
            self.subtitle_index += 1
        
        # 인포그래픽/요약 탭
        summary_tab = self.page.locator('[data-tab="summary"]')
        if await summary_tab.is_visible():
            await summary_tab.click()
            await self.wait(1)
            
            self.next_subtitle()  # 인포그래픽
            
            # 인포그래픽 항목 찾기
            infographic = self.page.locator('.report-item:has-text("인포그래픽")')
            if await infographic.first.is_visible():
                await infographic.first.click()
                await self.wait(5)
        else:
            self.srt.add_gap(5)
            self.subtitle_index += 1
        
        # 차트 크기 조절 (슬라이더가 있다면)
        self.next_subtitle()  # 차트 크기 조절 기능
        chart_slider = self.page.locator('input[type="range"]')
        if await chart_slider.first.is_visible():
            await chart_slider.first.fill("80")
            await self.wait(2)
            await chart_slider.first.fill("100")
            await self.wait(2)
        else:
            await self.wait(4)
        
        await self.scene_transition()
    
    async def scene_8_review(self):
        """Scene 8: 검토 기능"""
        print("\n[Scene 8] 검토 기능")
        
        self.next_subtitle()  # Step 7: 검토 기능
        await self.wait(3)
        
        # 검토완료 버튼 클릭
        review_btn = self.page.locator('#markReviewedBtn')
        if await review_btn.is_visible():
            self.next_subtitle()  # 검토완료 버튼
            await review_btn.click()
            await self.wait(4)
        else:
            self.srt.add_gap(4)
            self.subtitle_index += 1
        
        await self.scene_transition()
    
    async def scene_9_edit(self):
        """Scene 9: 편집 기능"""
        print("\n[Scene 9] 편집 기능")
        
        self.next_subtitle()  # Step 8: 편집 기능
        await self.wait(3)
        
        # 편집 버튼 클릭
        edit_btn = self.page.locator('#editBtn')
        if await edit_btn.is_visible():
            self.next_subtitle()  # 보도자료 내용 직접 수정 가능
            await edit_btn.click()
            await self.wait(2)
            
            # 편집 영역에 내용 수정 시뮬레이션
            edit_area = self.page.locator('#editableContent, .editable-content, [contenteditable="true"]')
            if await edit_area.first.is_visible():
                await edit_area.first.click()
                await self.wait(1)
            
            # 저장 또는 취소
            cancel_btn = self.page.locator('#cancelEditBtn')
            if await cancel_btn.is_visible():
                await cancel_btn.click()
            
            await self.wait(2)
        else:
            self.srt.add_gap(5)
            self.subtitle_index += 1
        
        await self.scene_transition()
    
    async def scene_10_export(self):
        """Scene 10: 내보내기"""
        print("\n[Scene 10] 내보내기")
        
        self.next_subtitle()  # Step 9: 보도자료 내보내기
        await self.wait(3)
        
        self.next_subtitle()  # PDF용 HTML / 한글 복붙용 HTML
        
        # PDF용 내보내기 버튼
        export_btn = self.page.locator('#exportBtn')
        if await export_btn.is_visible():
            await export_btn.click()
            await self.wait(3)
        
        # 한글 복붙용 내보내기 버튼
        hwp_btn = self.page.locator('#exportHwpBtn')
        if await hwp_btn.is_visible():
            await hwp_btn.click()
            await self.wait(2)
        
        await self.scene_transition()
    
    async def scene_finale(self):
        """마무리"""
        print("\n[Finale] 마무리")
        
        self.next_subtitle()  # 지역경제동향 보도자료 생성 완료!
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
        await self.scene_3_weight_settings()
        await self.scene_4_contact_info()
        await self.scene_5_download_analysis()
        await self.scene_6_grdp_settings()
        await self.scene_7_preview()
        await self.scene_8_review()
        await self.scene_9_edit()
        await self.scene_10_export()
        await self.scene_finale()
        
        # 마지막 대기
        await self.wait(3)
        
        elapsed = time.time() - self.start_time
        print("=" * 60)
        print(f"녹화 완료! 총 {elapsed:.1f}초")
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

