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


def main():
    """메인 함수 - CLI 인터페이스"""
    parser = argparse.ArgumentParser(
        description='지역경제동향 보도자료 자동생성',
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
        default='templates/mining_manufacturing_production.html',
        help='HTML 템플릿 파일 경로 (기본값: templates/mining_manufacturing_production.html)'
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
        if args.verbose:
            sheet_names = excel_extractor.get_sheet_names()
            print(f"사용 가능한 시트: {', '.join(sheet_names)}")
        
        # 템플릿 필러 초기화 및 처리
        if args.verbose:
            print("템플릿 채우는 중...")
        template_filler = TemplateFiller(template_manager, excel_extractor)
        filled_template = template_filler.fill_template()
        
        # 결과 저장
        if args.verbose:
            print(f"결과 저장 중: {output_path}")
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(filled_template)
        
        print(f"완료! 보도자료가 생성되었습니다: {output_path}")
        
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

