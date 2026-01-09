"""
PDF를 HTML로 변환하는 CLI 인터페이스
"""

import argparse
import sys
from pathlib import Path

from .pdf_to_html import PDFToHTMLConverter


def main():
    """CLI 메인 함수"""
    parser = argparse.ArgumentParser(
        description='PDF 파일을 이미지로 변환하고 OCR을 통해 HTML로 생성',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  python -m src.pdf_to_html_cli --pdf "2025년+2분기+지역경제동향+보도자료.pdf" --output output.html
  python -m src.pdf_to_html_cli --pdf input.pdf --output output.html --dpi 300 --use-pytesseract
        """
    )
    
    parser.add_argument(
        '--pdf', '-p',
        type=str,
        required=True,
        help='PDF 파일 경로'
    )
    
    parser.add_argument(
        '--output', '-o',
        type=str,
        required=True,
        help='출력 HTML 파일 경로'
    )
    
    parser.add_argument(
        '--dpi', '-d',
        type=int,
        default=300,
        help='PDF를 이미지로 변환할 때의 DPI (기본값: 300)'
    )
    
    parser.add_argument(
        '--use-pytesseract',
        action='store_true',
        help='pytesseract 사용 (기본값: easyocr)'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='상세한 로그 출력'
    )
    
    args = parser.parse_args()
    
    # 파일 경로 검증
    pdf_path = Path(args.pdf)
    output_path = Path(args.output)
    
    if not pdf_path.exists():
        print(f"에러: PDF 파일을 찾을 수 없습니다: {pdf_path}", file=sys.stderr)
        sys.exit(1)
    
    # 출력 디렉토리 생성
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    try:
        if args.verbose:
            print(f"PDF 파일 로딩 중: {pdf_path}")
            print(f"출력 파일: {output_path}")
            print(f"DPI: {args.dpi}")
            print(f"OCR 엔진: {'pytesseract' if args.use_pytesseract else 'easyocr'}")
        
        # PDF 변환기 초기화
        converter = PDFToHTMLConverter(
            use_easyocr=not args.use_pytesseract,
            dpi=args.dpi
        )
        
        # PDF를 HTML로 변환
        html_content = converter.generate_html_from_pdf(
            str(pdf_path),
            str(output_path)
        )
        
        if args.verbose:
            print(f"\n변환 완료!")
            print(f"생성된 HTML 파일 크기: {len(html_content)} 문자")
            print(f"출력 경로: {output_path}")
        
        print(f"성공: HTML 파일이 생성되었습니다: {output_path}")
        
    except Exception as e:
        print(f"에러: {str(e)}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()

