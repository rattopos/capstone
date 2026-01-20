import sys
from services.excel_processor import preprocess_excel

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="엑셀 추출 정상 동작 테스트")
    parser.add_argument("--input", type=str, required=True, help="입력 엑셀 파일 경로")
    parser.add_argument("--output", type=str, default=None, help="출력 엑셀 파일 경로 (생략 시 입력 파일 덮어쓰기)")
    args = parser.parse_args()

    result_path, success, message = preprocess_excel(args.input, args.output)
    print(f"[결과] 파일: {result_path}\n성공 여부: {success}\n메시지: {message}")
    sys.exit(0 if success else 1)
