"""
스크린샷에서 템플릿 생성 스크립트
correct_answer 폴더의 모든 스크린샷을 분석하여 템플릿 생성
"""

import sys
from pathlib import Path
from src.screenshot_template_generator import ScreenshotTemplateGenerator

def main():
    """메인 함수"""
    # 경로 설정
    base_dir = Path(__file__).parent
    excel_path = base_dir / "기초자료 수집표_2025년 2분기_캡스톤.xlsx"
    screenshots_dir = base_dir / "correct_answer"
    templates_dir = base_dir / "templates"
    
    # 디렉토리 확인
    if not excel_path.exists():
        print(f"에러: 엑셀 파일을 찾을 수 없습니다: {excel_path}")
        sys.exit(1)
    
    if not screenshots_dir.exists():
        print(f"에러: 스크린샷 폴더를 찾을 수 없습니다: {screenshots_dir}")
        sys.exit(1)
    
    # 템플릿 디렉토리 생성
    templates_dir.mkdir(exist_ok=True)
    
    # 스크린샷 파일 목록 가져오기
    screenshot_files = sorted(screenshots_dir.glob("*.png"))
    
    if not screenshot_files:
        print(f"경고: 스크린샷 파일을 찾을 수 없습니다: {screenshots_dir}")
        sys.exit(1)
    
    print(f"엑셀 파일: {excel_path}")
    print(f"스크린샷 개수: {len(screenshot_files)}")
    print("=" * 50)
    
    # 템플릿 생성기 초기화
    try:
        generator = ScreenshotTemplateGenerator(str(excel_path), use_easyocr=True)
    except Exception as e:
        print(f"에러: 템플릿 생성기 초기화 실패: {e}")
        sys.exit(1)
    
    # 각 스크린샷 처리
    success_count = 0
    error_count = 0
    
    for i, screenshot_path in enumerate(screenshot_files, 1):
        print(f"\n[{i}/{len(screenshot_files)}] 처리 중: {screenshot_path.name}")
        
        try:
            # 템플릿 이름 생성 (파일명에서 확장자 제거)
            template_name = screenshot_path.stem
            
            # 템플릿 생성
            html_template = generator.generate_template_from_screenshot(
                str(screenshot_path),
                template_name
            )
            
            # 템플릿 저장
            template_path = templates_dir / f"{template_name}.html"
            with open(template_path, 'w', encoding='utf-8') as f:
                f.write(html_template)
            
            print(f"  ✓ 템플릿 생성 완료: {template_path.name}")
            success_count += 1
            
        except Exception as e:
            print(f"  ✗ 오류 발생: {e}")
            import traceback
            traceback.print_exc()
            error_count += 1
    
    # 리소스 정리
    generator.close()
    
    # 결과 요약
    print("\n" + "=" * 50)
    print(f"완료: 성공 {success_count}개, 실패 {error_count}개")
    print(f"템플릿 저장 위치: {templates_dir}")
    print("=" * 50)

if __name__ == '__main__':
    main()

