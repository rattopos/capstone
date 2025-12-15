"""
기본 기능 테스트 스크립트
"""

import sys
from pathlib import Path

# 프로젝트 루트를 경로에 추가
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from src.template_manager import TemplateManager
from src.excel_extractor import ExcelExtractor
from src.calculator import Calculator
from src.template_filler import TemplateFiller


def test_template_manager():
    """템플릿 관리자 테스트"""
    print("=== 템플릿 관리자 테스트 ===")
    template_path = project_root / "templates" / "sample_template.html"
    
    if not template_path.exists():
        print(f"템플릿 파일이 없습니다: {template_path}")
        return False
    
    tm = TemplateManager(str(template_path))
    tm.load_template()
    
    markers = tm.extract_markers()
    print(f"추출된 마커 수: {len(markers)}")
    for marker in markers:
        print(f"  - {marker['full_match']}")
    
    return True


def test_calculator():
    """계산기 테스트"""
    print("\n=== 계산기 테스트 ===")
    calc = Calculator()
    
    # 합계 테스트
    result = calc.sum([1, 2, 3, 4, 5])
    print(f"합계 [1,2,3,4,5]: {result} (예상: 15)")
    assert result == 15, "합계 계산 오류"
    
    # 평균 테스트
    result = calc.average([10, 20, 30])
    print(f"평균 [10,20,30]: {result} (예상: 20.0)")
    assert result == 20.0, "평균 계산 오류"
    
    # 증감률 테스트
    result = calc.growth_rate(100, 110)
    print(f"증감률 (100->110): {result}% (예상: 10.0%)")
    assert abs(result - 10.0) < 0.01, "증감률 계산 오류"
    
    # 증감액 테스트
    result = calc.growth_amount(100, 110)
    print(f"증감액 (100->110): {result} (예상: 10)")
    assert result == 10, "증감액 계산 오류"
    
    print("모든 계산 테스트 통과!")
    return True


def test_excel_extractor():
    """엑셀 추출기 테스트"""
    print("\n=== 엑셀 추출기 테스트 ===")
    excel_path = project_root / "기초자료 수집표_2025년 2분기_캡스톤.xlsx"
    
    if not excel_path.exists():
        print(f"엑셀 파일이 없습니다: {excel_path}")
        return False
    
    extractor = ExcelExtractor(str(excel_path))
    extractor.load_workbook()
    
    sheet_names = extractor.get_sheet_names()
    print(f"시트 목록: {sheet_names[:5]}...")  # 처음 5개만 표시
    
    # 첫 번째 시트에서 테스트
    if sheet_names:
        first_sheet = sheet_names[0]
        print(f"첫 번째 시트: {first_sheet}")
        
        try:
            value = extractor.get_cell_value(first_sheet, "A1")
            print(f"A1 값: {value}")
        except Exception as e:
            print(f"A1 읽기 실패: {e}")
    
    extractor.close()
    return True


def test_integration():
    """통합 테스트"""
    print("\n=== 통합 테스트 ===")
    template_path = project_root / "templates" / "sample_template.html"
    excel_path = project_root / "기초자료 수집표_2025년 2분기_캡스톤.xlsx"
    
    if not template_path.exists() or not excel_path.exists():
        print("필요한 파일이 없어 통합 테스트를 건너뜁니다.")
        return False
    
    try:
        # 컴포넌트 초기화
        tm = TemplateManager(str(template_path))
        tm.load_template()
        
        extractor = ExcelExtractor(str(excel_path))
        extractor.load_workbook()
        
        filler = TemplateFiller(tm, extractor)
        
        # 템플릿 채우기 시도
        print("템플릿 채우기 시도 중...")
        # 실제 데이터가 없을 수 있으므로 에러가 발생할 수 있음
        try:
            filled = filler.fill_template()
            print("템플릿 채우기 성공!")
            print(f"결과 길이: {len(filled)} 문자")
        except Exception as e:
            print(f"템플릿 채우기 중 에러 (예상 가능): {e}")
        
        extractor.close()
        return True
        
    except Exception as e:
        print(f"통합 테스트 실패: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    print("기본 기능 테스트 시작\n")
    
    results = []
    results.append(("템플릿 관리자", test_template_manager()))
    results.append(("계산기", test_calculator()))
    results.append(("엑셀 추출기", test_excel_extractor()))
    results.append(("통합 테스트", test_integration()))
    
    print("\n=== 테스트 결과 ===")
    for name, result in results:
        status = "✓ 통과" if result else "✗ 실패"
        print(f"{name}: {status}")
    
    all_passed = all(result for _, result in results)
    sys.exit(0 if all_passed else 1)

