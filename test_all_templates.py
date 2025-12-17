"""
모든 템플릿을 테스트하고 결측치를 확인하는 스크립트
"""
import sys
import re
from pathlib import Path
from typing import List, Dict, Set

# 프로젝트 루트를 경로에 추가
sys.path.insert(0, str(Path(__file__).parent))

from src.template_manager import TemplateManager
from src.excel_extractor import ExcelExtractor
from src.template_filler import TemplateFiller
from src.flexible_mapper import FlexibleMapper
from src.period_detector import PeriodDetector

# 기본 엑셀 파일 경로
BASE_DIR = Path(__file__).parent
DEFAULT_EXCEL_FILE = BASE_DIR / '기초자료 수집표_2025년 2분기_캡스톤.xlsx'

# 템플릿 목록
TEMPLATES = [
    '광공업생산.html',
    '서비스업생산.html',
    '소매판매.html',
    '고용률.html',
    '실업률.html',
    '물가동향.html',
    '건설수주.html',
    '수출.html',
    '수입.html',
    '국내인구이동.html',
]

# 템플릿명 -> 시트명 매핑 (키워드 기반 매칭을 위해)
TEMPLATE_SHEET_MAPPING = {
    '광공업생산.html': '광공업생산',
    '서비스업생산.html': '서비스업생산',
    '소매판매.html': '소비(소매, 추가)',
    '고용률.html': '고용률',
    '실업률.html': '실업률',  # 또는 '실업자 수'
    '물가동향.html': '지출목적별 물가',  # 또는 '품목성질별 물가'
    '건설수주.html': '건설 (공표자료)',
    '수출.html': '수출',
    '수입.html': '수입',
    '국내인구이동.html': '시도 간 이동',  # 또는 다른 인구이동 관련 시트
}


def find_unfilled_markers(html_content: str) -> List[str]:
    """HTML에서 채워지지 않은 마커를 찾습니다."""
    marker_pattern = re.compile(r'\{([^:{}]+):([^:}]+)(?::([^}]+))?\}')
    unfilled = []
    
    matches = marker_pattern.finditer(html_content)
    for match in matches:
        full_match = match.group(0)
        # N/A나 빈 값이 아닌 경우만 체크 (마커 자체가 남아있는 경우)
        if full_match in html_content:
            unfilled.append(full_match)
    
    return unfilled


def test_template(template_name: str, excel_path: Path, output_dir: Path) -> Dict:
    """단일 템플릿을 테스트합니다."""
    print(f"\n{'='*60}")
    print(f"테스트 중: {template_name}")
    print(f"{'='*60}")
    
    template_path = BASE_DIR / 'templates' / template_name
    
    if not template_path.exists():
        return {
            'template': template_name,
            'success': False,
            'error': f'템플릿 파일을 찾을 수 없습니다: {template_path}',
            'unfilled_markers': [],
            'sheet_mapping_issues': []
        }
    
    if not excel_path.exists():
        return {
            'template': template_name,
            'success': False,
            'error': f'엑셀 파일을 찾을 수 없습니다: {excel_path}',
            'unfilled_markers': [],
            'sheet_mapping_issues': []
        }
    
    try:
        # 템플릿 관리자 초기화
        template_manager = TemplateManager(str(template_path))
        template_manager.load_template()
        
        # 마커 추출
        markers = template_manager.extract_markers()
        print(f"발견된 마커 수: {len(markers)}")
        
        # 엑셀 추출기 초기화
        excel_extractor = ExcelExtractor(str(excel_path))
        excel_extractor.load_workbook()
        
        # 사용 가능한 시트 목록
        sheet_names = excel_extractor.get_sheet_names()
        print(f"사용 가능한 시트: {', '.join(sheet_names)}")
        
        # 필요한 시트 확인
        required_sheets = set()
        for marker in markers:
            sheet_name = marker.get('sheet_name', '').strip()
            if sheet_name:
                required_sheets.add(sheet_name)
        
        print(f"템플릿에서 요구하는 시트: {', '.join(required_sheets)}")
        
        # 유연한 매핑으로 시트 찾기
        flexible_mapper = FlexibleMapper(excel_extractor)
        actual_sheet_mapping = {}
        sheet_mapping_issues = []
        
        for required_sheet in required_sheets:
            actual_sheet = flexible_mapper.find_sheet_by_name(required_sheet)
            if actual_sheet:
                actual_sheet_mapping[required_sheet] = actual_sheet
                if required_sheet != actual_sheet:
                    print(f"  매핑: '{required_sheet}' -> '{actual_sheet}'")
            else:
                sheet_mapping_issues.append(f"시트를 찾을 수 없음: '{required_sheet}'")
                print(f"  ❌ 시트를 찾을 수 없음: '{required_sheet}'")
        
        if sheet_mapping_issues:
            excel_extractor.close()
            return {
                'template': template_name,
                'success': False,
                'error': '필요한 시트를 찾을 수 없습니다',
                'unfilled_markers': [],
                'sheet_mapping_issues': sheet_mapping_issues
            }
        
        # 첫 번째 시트로 연도/분기 감지
        primary_sheet = list(actual_sheet_mapping.values())[0]
        period_detector = PeriodDetector(excel_extractor)
        periods_info = period_detector.detect_available_periods(primary_sheet)
        
        year = periods_info['default_year']
        quarter = periods_info['default_quarter']
        print(f"사용할 연도/분기: {year}년 {quarter}분기")
        
        # 템플릿 필러 초기화 및 처리
        template_filler = TemplateFiller(template_manager, excel_extractor)
        filled_template = template_filler.fill_template(
            sheet_name=primary_sheet,
            year=year,
            quarter=quarter
        )
        
        # 결과 저장
        output_dir.mkdir(exist_ok=True, parents=True)
        output_filename = f"test_{template_name}"
        output_path = output_dir / output_filename
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(filled_template)
        
        print(f"생성된 파일: {output_path}")
        
        # 채워지지 않은 마커 확인
        unfilled_markers = find_unfilled_markers(filled_template)
        
        if unfilled_markers:
            print(f"\n⚠️  채워지지 않은 마커 {len(unfilled_markers)}개 발견:")
            for marker in unfilled_markers[:10]:  # 최대 10개만 표시
                print(f"  - {marker}")
            if len(unfilled_markers) > 10:
                print(f"  ... 외 {len(unfilled_markers) - 10}개")
        else:
            print("\n✅ 모든 마커가 성공적으로 채워졌습니다!")
        
        # 엑셀 파일 닫기
        excel_extractor.close()
        
        return {
            'template': template_name,
            'success': True,
            'error': None,
            'unfilled_markers': unfilled_markers,
            'sheet_mapping_issues': [],
            'output_path': str(output_path),
            'marker_count': len(markers),
            'unfilled_count': len(unfilled_markers)
        }
        
    except Exception as e:
        import traceback
        error_msg = f"오류 발생: {str(e)}\n{traceback.format_exc()}"
        print(f"\n❌ {error_msg}")
        return {
            'template': template_name,
            'success': False,
            'error': error_msg,
            'unfilled_markers': [],
            'sheet_mapping_issues': []
        }


def main():
    """모든 템플릿을 테스트합니다."""
    excel_path = DEFAULT_EXCEL_FILE
    output_dir = BASE_DIR / 'output' / 'test_results'
    
    if not excel_path.exists():
        print(f"❌ 엑셀 파일을 찾을 수 없습니다: {excel_path}")
        return
    
    print(f"엑셀 파일: {excel_path}")
    print(f"출력 디렉토리: {output_dir}")
    print(f"\n총 {len(TEMPLATES)}개 템플릿 테스트 시작...")
    
    results = []
    for template_name in TEMPLATES:
        result = test_template(template_name, excel_path, output_dir)
        results.append(result)
    
    # 결과 요약
    print(f"\n{'='*60}")
    print("테스트 결과 요약")
    print(f"{'='*60}")
    
    success_count = sum(1 for r in results if r['success'])
    total_markers = sum(r.get('marker_count', 0) for r in results if r['success'])
    total_unfilled = sum(r.get('unfilled_count', 0) for r in results if r['success'])
    
    print(f"\n성공: {success_count}/{len(TEMPLATES)}")
    print(f"총 마커 수: {total_markers}")
    print(f"채워지지 않은 마커: {total_unfilled}")
    
    print("\n상세 결과:")
    for result in results:
        status = "✅" if result['success'] and not result.get('unfilled_markers') else "⚠️" if result['success'] else "❌"
        print(f"{status} {result['template']}")
        if result.get('unfilled_count', 0) > 0:
            print(f"   채워지지 않은 마커: {result['unfilled_count']}개")
        if result.get('error'):
            print(f"   오류: {result['error']}")
        if result.get('sheet_mapping_issues'):
            for issue in result['sheet_mapping_issues']:
                print(f"   시트 매핑 문제: {issue}")


if __name__ == '__main__':
    main()

