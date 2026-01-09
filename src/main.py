"""
메인 애플리케이션
전체 워크플로우 오케스트레이션 및 CLI 인터페이스
"""

import argparse
import sys
from pathlib import Path
from typing import Optional

from .template_manager import TemplateManager
from .excel_extractor import ExcelExtractor
from .template_filler import TemplateFiller
from .config import Config
from .template_generator import TemplateGenerator


def main():
    """메인 함수 - CLI 인터페이스"""
    parser = argparse.ArgumentParser(
        description='통계청 보도자료 자동 생성 시스템',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  python -m src.main --excel data/data.xlsx --output output/result.html
  python -m src.main --template templates/template.html --excel data/data.xlsx --output output/result.html
  
마커 형식:
  {시트명:셀주소}              : 단일 셀 값
  {시트명:셀주소:계산식}        : 계산식 적용
  예: {시트1:A1}, {시트1:A1:A5:sum}, {시트1:A1:A2:증감률}
        """
    )
    
    parser.add_argument(
        '--template', '-t',
        type=str,
        default='templates/dynamic_template.html',
        help='HTML 템플릿 파일 경로 (기본값: templates/dynamic_template.html)'
    )
    
    parser.add_argument(
        '--excel', '-e',
        type=str,
        required=True,
        help='엑셀 데이터 파일 경로'
    )
    
    parser.add_argument(
        '--output', '-o',
        type=str,
        required=True,
        help='출력 파일 경로'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='상세한 로그 출력'
    )
    
    parser.add_argument(
        '--year',
        type=int,
        default=None,
        help='분석할 연도 (예: 2025)'
    )
    
    parser.add_argument(
        '--quarter',
        type=int,
        default=None,
        help='분석할 분기 (1-4)'
    )
    
    parser.add_argument(
        '--all-sheets',
        action='store_true',
        default=True,
        help='모든 시트 처리 (기본값: True)'
    )
    
    parser.add_argument(
        '--sheet',
        type=str,
        default=None,
        help='특정 시트만 처리 (--all-sheets와 상호 배타적)'
    )
    
    parser.add_argument(
        '--generate-template',
        action='store_true',
        help='템플릿 생성 모드 (스크린샷에서 템플릿 생성)'
    )
    
    parser.add_argument(
        '--screenshot',
        type=str,
        default=None,
        help='템플릿 생성에 사용할 스크린샷 이미지 경로'
    )
    
    args = parser.parse_args()
    
    # 파일 경로 검증
    template_path = Path(args.template)
    excel_path = Path(args.excel)
    output_path = Path(args.output)
    
    if not template_path.exists():
        print(f"에러: 템플릿 파일을 찾을 수 없습니다: {template_path}", file=sys.stderr)
        sys.exit(1)
    
    if not excel_path.exists():
        print(f"에러: 엑셀 파일을 찾을 수 없습니다: {excel_path}", file=sys.stderr)
        sys.exit(1)
    
    # 출력 디렉토리 생성
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    # 템플릿 생성 모드
    if args.generate_template:
        if not args.screenshot:
            print("에러: 템플릿 생성 모드에서는 --screenshot 옵션이 필요합니다.", file=sys.stderr)
            sys.exit(1)
        
        screenshot_path = Path(args.screenshot)
        if not screenshot_path.exists():
            print(f"에러: 스크린샷 파일을 찾을 수 없습니다: {screenshot_path}", file=sys.stderr)
            sys.exit(1)
        
        # 시트명 결정
        sheet_name = args.sheet if args.sheet else '광공업생산'
        
        try:
            if args.verbose:
                print(f"스크린샷에서 템플릿 생성 중: {screenshot_path}")
                print(f"시트명: {sheet_name}")
            
            generator = TemplateGenerator()
            base_template = template_path if template_path.exists() else None
            template_html = generator.generate_from_screenshot(
                str(screenshot_path),
                sheet_name,
                str(output_path),
                str(base_template) if base_template else None
            )
            
            print(f"완료! 템플릿이 생성되었습니다: {output_path}")
            sys.exit(0)
        except Exception as e:
            print(f"에러 발생: {str(e)}", file=sys.stderr)
            if args.verbose:
                import traceback
                traceback.print_exc()
            sys.exit(1)
    
    # Config 설정
    config = None
    if args.year is not None and args.quarter is not None:
        try:
            config = Config(args.year, args.quarter)
            if args.verbose:
                print(f"설정: {config.year}년 {config.quarter}분기")
        except ValueError as e:
            print(f"에러: {e}", file=sys.stderr)
            sys.exit(1)
    elif args.year is not None or args.quarter is not None:
        print("에러: --year와 --quarter는 함께 지정해야 합니다.", file=sys.stderr)
        sys.exit(1)
    
    try:
        # 템플릿 관리자 초기화
        if args.verbose:
            print(f"템플릿 로딩 중: {template_path}")
        template_manager = TemplateManager(str(template_path))
        template_manager.load_template()
        
        # 마커 추출 및 표시
        markers = template_manager.extract_markers()
        if args.verbose:
            print(f"발견된 마커 수: {len(markers)}")
            for i, marker in enumerate(markers, 1):
                print(f"  {i}. {marker['full_match']}")
        
        # 엑셀 추출기 초기화
        if args.verbose:
            print(f"엑셀 파일 로딩 중: {excel_path}")
        excel_extractor = ExcelExtractor(str(excel_path))
        excel_extractor.load_workbook()
        
        # 사용 가능한 시트 표시
        all_sheet_names = excel_extractor.get_sheet_names()
        if args.verbose:
            print(f"사용 가능한 시트: {', '.join(all_sheet_names)}")
        
        # 처리할 시트 결정
        if args.sheet:
            if args.sheet not in all_sheet_names:
                print(f"에러: 시트 '{args.sheet}'를 찾을 수 없습니다.", file=sys.stderr)
                print(f"사용 가능한 시트: {', '.join(all_sheet_names)}", file=sys.stderr)
                sys.exit(1)
            sheets_to_process = [args.sheet]
        else:
            # 모든 시트 처리
            sheets_to_process = all_sheet_names
        
        # 각 시트별로 처리
        if len(sheets_to_process) == 1:
            # 단일 시트인 경우 기존 방식대로 처리
            sheet_name = sheets_to_process[0]
            if args.verbose:
                print(f"시트 처리 중: {sheet_name}")
            
            template_filler = TemplateFiller(template_manager, excel_extractor, config, sheet_name=sheet_name)
            filled_template = template_filler.fill_template()
            
            # 결과 저장
            if args.verbose:
                print(f"결과 저장 중: {output_path}")
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(filled_template)
            
            print(f"완료! 보도자료가 생성되었습니다: {output_path}")
        else:
            # 다중 시트인 경우 각 시트별로 별도 파일 생성
            if args.verbose:
                print(f"{len(sheets_to_process)}개 시트 처리 중...")
            
            for sheet_name in sheets_to_process:
                if args.verbose:
                    print(f"  시트 처리 중: {sheet_name}")
                
                # 각 시트별로 템플릿 필러 생성 (시트별로 다른 설정 사용)
                template_filler = TemplateFiller(template_manager, excel_extractor, config, sheet_name=sheet_name)
                filled_template = template_filler.fill_template()
                
                # 시트명을 포함한 출력 파일명 생성
                output_file = output_path.parent / f"{output_path.stem}_{sheet_name}{output_path.suffix}"
                
                if args.verbose:
                    print(f"  결과 저장 중: {output_file}")
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(filled_template)
            
            print(f"완료! {len(sheets_to_process)}개 시트의 보도자료가 생성되었습니다.")
        
        # 엑셀 파일 닫기
        excel_extractor.close()
        
    except Exception as e:
        print(f"에러 발생: {str(e)}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()

