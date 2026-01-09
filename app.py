"""
Flask 웹 애플리케이션
지역경제동향 보도자료 자동생성 웹 인터페이스
"""

import os
import sys
import json
import tempfile
from pathlib import Path

from flask import Flask, request, render_template, jsonify, send_file, send_from_directory
from werkzeug.utils import secure_filename

# 모듈 경로 설정
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.template_manager import TemplateManager
from src.excel_extractor import ExcelExtractor
from src.template_filler import TemplateFiller
from src.period_detector import PeriodDetector
from src.template_generator import TemplateGenerator
from src.excel_header_parser import ExcelHeaderParser
from src.flexible_mapper import FlexibleMapper
from src.pdf_generator import PDFGenerator
from src.docx_generator import DOCXGenerator
from src.schema_loader import SchemaLoader
from bs4 import BeautifulSoup

# Blueprint 임포트
from routes import validation_bp

app = Flask(__name__, template_folder='flask_templates', static_folder='static')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['OUTPUT_FOLDER'] = tempfile.mkdtemp()

# Blueprint 등록
app.register_blueprint(validation_bp)

# 기본 엑셀 파일 경로
BASE_DIR = Path(__file__).parent
DEFAULT_EXCEL_FILE = BASE_DIR / '기초자료 수집표_2025년 2분기_캡스톤.xlsx'

# 허용된 파일 확장자
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'html'}
ALLOWED_IMAGE_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp'}

# 스키마 로더 초기화
schema_loader = SchemaLoader()

# 성능 최적화: 템플릿 목록 캐시
_templates_cache = None
_templates_cache_mtime = None  # 템플릿 폴더 수정 시간

def get_excel_extractor(excel_path: Path, force_reload: bool = False) -> ExcelExtractor:
    """
    엑셀 추출기를 새로 생성합니다 (캐시 없이).
    
    Args:
        excel_path: 엑셀 파일 경로
        force_reload: 강제로 다시 로드할지 여부 (사용되지 않음, 항상 새로 생성)
        
    Returns:
        ExcelExtractor 인스턴스
    """
    # 항상 새로 생성
    excel_extractor = ExcelExtractor(str(excel_path.resolve()))
    excel_extractor.load_workbook()
    
    return excel_extractor


def get_template_manager(template_path: Path, force_reload: bool = False) -> TemplateManager:
    """
    템플릿 매니저를 새로 생성합니다 (캐시 없이).
    
    Args:
        template_path: 템플릿 파일 경로
        force_reload: 강제로 다시 로드할지 여부 (사용되지 않음, 항상 새로 생성)
        
    Returns:
        TemplateManager 인스턴스
    """
    # 항상 새로 생성
    template_manager = TemplateManager(str(template_path.resolve()))
    template_manager.load_template()
    
    return template_manager


def allowed_file(filename):
    """파일 확장자 검증"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def allowed_image_file(filename):
    """이미지 파일 확장자 검증"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_IMAGE_EXTENSIONS


def get_template_for_sheet(sheet_name):
    """
    시트명에 해당하는 템플릿 정보를 반환합니다.
    
    Args:
        sheet_name: 엑셀 시트명
        
    Returns:
        dict: {'template': 템플릿 파일명, 'display_name': 표시용 이름}
        매핑이 없으면 기본값 반환
    """
    template_info = schema_loader.get_template_for_sheet(sheet_name)
    if template_info:
        return template_info
    
    # 기본값: 광공업생산 템플릿 사용
    template_mapping = schema_loader.load_template_mapping()
    return template_mapping.get('광공업생산', {'template': '광공업생산.html', 'display_name': '광공업생산'})


@app.route('/')
def index():
    """메인 페이지"""
    return render_template('index.html')


@app.route('/api/templates', methods=['GET'])
def get_templates():
    """사용 가능한 템플릿 목록 반환 (각 템플릿이 필요한 시트 정보 포함)
    
    성능 최적화: 결과를 캐시하여 반복 호출 시 재파싱하지 않습니다.
    템플릿 폴더의 수정 시간이 변경되면 캐시를 무효화합니다.
    """
    global _templates_cache, _templates_cache_mtime
    
    templates_dir = Path('templates')
    
    # 캐시 유효성 확인
    if templates_dir.exists():
        current_mtime = templates_dir.stat().st_mtime
        # 개별 파일 수정 시간도 확인
        for f in templates_dir.glob('*.html'):
            file_mtime = f.stat().st_mtime
            if file_mtime > current_mtime:
                current_mtime = file_mtime
        
        if _templates_cache is not None and _templates_cache_mtime == current_mtime:
            return jsonify({'templates': _templates_cache})
    
    templates = []
    
    # templates 디렉토리의 모든 HTML 파일을 스캔
    if templates_dir.exists():
        # 모든 HTML 파일 찾기
        html_files = list(templates_dir.glob('*.html'))
        
        # 템플릿 매핑 한 번만 로드
        template_mapping = schema_loader.load_template_mapping()
        
        for template_path in html_files:
            template_name = template_path.name
            
            # 템플릿에서 필요한 시트 목록 추출
            try:
                template_manager = get_template_manager(template_path)
                markers = template_manager.extract_markers()
                
                # 마커에서 시트명 추출 (중복 제거)
                required_sheets = set()
                for marker in markers:
                    sheet_name = marker.get('sheet_name', '').strip()
                    if sheet_name:
                        required_sheets.add(sheet_name)
                
                # display_name 찾기 (템플릿 매핑에서 먼저 찾고, 없으면 파일명 사용)
                display_name = template_name.replace('.html', '')
                for sheet_name, info in template_mapping.items():
                    if info['template'] == template_name:
                        display_name = info['display_name']
                        break
                
                templates.append({
                    'name': template_name,
                    'path': str(template_path),
                    'display_name': display_name,
                    'required_sheets': list(required_sheets)
                })
            except Exception as e:
                # 템플릿 파싱 실패 시 기본 정보만 반환
                display_name = template_name.replace('.html', '')
                for sheet_name, info in template_mapping.items():
                    if info['template'] == template_name:
                        display_name = info['display_name']
                        break
                
                templates.append({
                    'name': template_name,
                    'path': str(template_path),
                    'display_name': display_name,
                    'required_sheets': []
                })
    
    # 서울 관련 템플릿 제외 (서울.html만 제외, 서울주요지표는 포함)
    excluded_templates = {'서울.html'}
    templates = [t for t in templates if t['name'] not in excluded_templates]
    
    # 템플릿 이름으로 정렬
    templates.sort(key=lambda x: x['display_name'])
    
    # 캐시에 저장
    _templates_cache = templates
    _templates_cache_mtime = current_mtime if templates_dir.exists() else None
    
    return jsonify({'templates': templates})


@app.route('/api/template-sheets', methods=['POST'])
def get_template_sheets():
    """템플릿이 필요한 시트 목록 반환"""
    try:
        template_name = request.form.get('template_name', '')
        if not template_name:
            return jsonify({'error': '템플릿명이 필요합니다.'}), 400
        
        template_path = Path('templates') / template_name
        if not template_path.exists():
            return jsonify({'error': f'템플릿 파일을 찾을 수 없습니다: {template_name}'}), 404
        
        # 템플릿에서 필요한 시트 목록 추출
        template_manager = get_template_manager(template_path)
        markers = template_manager.extract_markers()
        
        # 마커에서 시트명 추출 (중복 제거)
        required_sheets = set()
        for marker in markers:
            sheet_name = marker.get('sheet_name', '').strip()
            if sheet_name:
                required_sheets.add(sheet_name)
        
        return jsonify({
            'template_name': template_name,
            'required_sheets': list(required_sheets)
        })
    except Exception as e:
        return jsonify({
            'error': f'템플릿 분석 중 오류가 발생했습니다: {str(e)}'
        }), 500


def detect_sheet_scale(excel_extractor, sheet_name):
    """
    시트의 데이터 스케일을 감지하여 적절한 기본값 반환
    - 한 자리 수 (1-9) → 1
    - 두 자리 수 (10-99) → 10
    - 세 자리 수 (100-999) → 100
    - 네 자리 수 이상 (1000+) → 1000
    """
    import math
    try:
        sheet = excel_extractor.get_sheet(sheet_name)
        values = []
        
        # 데이터 영역에서 숫자 값 수집 (4행부터, 7열 이후)
        for row in range(4, min(104, sheet.max_row + 1)):
            for col in range(7, min(sheet.max_column + 1, 20)):
                cell = sheet.cell(row=row, column=col)
                if cell.value is not None:
                    try:
                        val = float(cell.value)
                        if not math.isnan(val) and not math.isinf(val) and val > 0:
                            values.append(abs(val))
                    except (ValueError, TypeError):
                        continue
        
        if not values:
            return 1.0
        
        # 중앙값 기준으로 스케일 결정
        values.sort()
        median_value = values[len(values) // 2]
        
        if median_value < 10:
            return 1.0
        elif median_value < 100:
            return 10.0
        elif median_value < 1000:
            return 100.0
        else:
            return 1000.0
    except Exception:
        return 1.0


def find_missing_values_in_sheet(excel_extractor, sheet_name, year, quarter):
    """시트에서 결측치 찾기 - 동적 열 탐색 사용"""
    import re
    missing_values = []
    
    try:
        sheet = excel_extractor.get_sheet(sheet_name)
        
        # 분기별 열 찾기
        target_col = None
        prev_col = None
        
        # 헤더에서 해당 분기 열 찾기 (1-5행 검사)
        for row in range(1, 6):
            for col in range(1, min(sheet.max_column + 1, 150)):
                cell = sheet.cell(row=row, column=col)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    # 현재 분기 열 찾기 (예: "2025 2/4" 또는 "2025 2/4p")
                    pattern = rf'{year}\s*{quarter}/4[pP]?'
                    if re.search(pattern, cell_str):
                        target_col = col
                    # 전년 동분기 열 찾기
                    prev_pattern = rf'{year-1}\s*{quarter}/4[pP]?'
                    if re.search(prev_pattern, cell_str):
                        prev_col = col
        
        if not target_col:
            return missing_values
        
        # 스케일 감지
        scale = detect_sheet_scale(excel_extractor, sheet_name)
        
        # 지역 열을 동적으로 찾기
        region_col = _find_region_column_in_sheet(sheet)
        
        # 카테고리 열을 동적으로 찾기
        category_col = _find_category_column_in_sheet(sheet)
        
        # 데이터 시작 행 찾기
        data_start_row = _find_data_start_row(sheet, region_col)
        
        # 데이터 영역 검사
        for row in range(data_start_row, min(sheet.max_row + 1, 500)):
            cell_value = sheet.cell(row=row, column=target_col).value
            region_value = sheet.cell(row=row, column=region_col).value
            category_value = sheet.cell(row=row, column=category_col).value if category_col else None
            
            # 결측치 확인 (None, 빈 문자열, '-')
            is_missing = (
                cell_value is None or 
                (isinstance(cell_value, str) and (not cell_value.strip() or cell_value.strip() == '-'))
            )
            
            if is_missing and region_value:
                region_str = str(region_value).strip()
                category_str = str(category_value).strip() if category_value else ''
                
                # 지역명이 있고, 총지수/계 등인 경우만 추가
                if region_str and category_str in ['총지수', '계', '   계', '합계', '']:
                    missing_values.append({
                        'sheet': sheet_name,
                        'region': region_str,
                        'category': category_str or '합계',
                        'year': year,
                        'quarter': quarter,
                        'row': row,
                        'col': target_col,
                        'default_value': scale,
                        'cell_address': _col_to_letter(target_col) + str(row)
                    })
        
        return missing_values
    except Exception as e:
        print(f"[ERROR] 결측치 확인 중 오류: {sheet_name}, {str(e)}")
        return missing_values


def _find_region_column_in_sheet(sheet) -> int:
    """시트에서 지역 열을 동적으로 찾습니다."""
    region_names = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', 
                   '세종', '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
    region_keywords = ['시도', '지역', 'region', 'area', '시·도']
    
    # 헤더에서 키워드로 찾기
    for row in range(1, 5):
        for col in range(1, min(10, sheet.max_column + 1)):
            cell = sheet.cell(row=row, column=col)
            if cell.value:
                header_text = str(cell.value).strip().lower()
                if any(keyword in header_text for keyword in region_keywords):
                    return col
    
    # 데이터에서 지역명이 포함된 열 찾기
    best_col = 2
    best_count = 0
    
    for col in range(1, min(10, sheet.max_column + 1)):
        region_count = 0
        for row in range(2, min(25, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=col)
            if cell.value:
                cell_str = str(cell.value).strip()
                if any(region in cell_str for region in region_names):
                    region_count += 1
        
        if region_count > best_count:
            best_count = region_count
            best_col = col
    
    return best_col


def _find_category_column_in_sheet(sheet) -> int:
    """시트에서 카테고리 열을 동적으로 찾습니다."""
    category_keywords = ['산업', '업태', '품목', 'category', 'industry', '업종', '분류명']
    category_values = ['총지수', '계', '합계', '반도체', '전자', '서비스', '제조업', '백화점']
    
    # 헤더에서 키워드로 찾기
    for row in range(1, 5):
        for col in range(3, min(12, sheet.max_column + 1)):
            cell = sheet.cell(row=row, column=col)
            if cell.value:
                header_text = str(cell.value).strip().lower()
                if any(keyword in header_text for keyword in category_keywords):
                    return col
    
    # 데이터에서 카테고리 값이 포함된 열 찾기
    best_col = 6
    best_count = 0
    
    for col in range(3, min(12, sheet.max_column + 1)):
        cat_count = 0
        for row in range(2, min(30, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=col)
            if cell.value:
                cell_str = str(cell.value).strip()
                if any(cat in cell_str for cat in category_values):
                    cat_count += 1
        
        if cat_count > best_count:
            best_count = cat_count
            best_col = col
    
    return best_col


def _find_data_start_row(sheet, region_col: int) -> int:
    """데이터 시작 행을 찾습니다."""
    region_names = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', 
                   '세종', '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
    
    for row in range(1, min(20, sheet.max_row + 1)):
        cell = sheet.cell(row=row, column=region_col)
        if cell.value:
            cell_str = str(cell.value).strip()
            if any(region in cell_str for region in region_names):
                return row
    
    return 4


def _col_to_letter(col: int) -> str:
    """열 번호를 열 문자로 변환"""
    result = ""
    while col > 0:
        col -= 1
        result = chr(65 + (col % 26)) + result
        col //= 26
    return result


def detect_structure_changes(excel_extractor, sheet_name: str) -> dict:
    """
    시트의 구조 변경을 감지합니다.
    
    Args:
        excel_extractor: 엑셀 추출기
        sheet_name: 시트 이름
        
    Returns:
        구조 분석 결과 및 경고 메시지
    """
    warnings = []
    structure_info = {}
    
    try:
        sheet = excel_extractor.get_sheet(sheet_name)
        
        # 지역 열 탐색
        region_col = _find_region_column_in_sheet(sheet)
        structure_info['region_column'] = region_col
        
        # 기존 기대값과 다르면 경고
        if region_col != 2:
            warnings.append(f"지역 열이 B열이 아닌 {_col_to_letter(region_col)}열에서 발견됨")
        
        # 카테고리 열 탐색
        category_col = _find_category_column_in_sheet(sheet)
        structure_info['category_column'] = category_col
        
        if category_col != 6:
            warnings.append(f"카테고리 열이 F열이 아닌 {_col_to_letter(category_col)}열에서 발견됨")
        
        # 데이터 시작 행 탐색
        data_start_row = _find_data_start_row(sheet, region_col)
        structure_info['data_start_row'] = data_start_row
        
        if data_start_row > 5:
            warnings.append(f"데이터 시작 행이 {data_start_row}행으로 예상보다 늦음")
        
        # 분기 열 확인
        quarter_cols_found = 0
        for row in range(1, 6):
            for col in range(1, min(sheet.max_column + 1, 150)):
                cell = sheet.cell(row=row, column=col)
                if cell.value:
                    import re
                    cell_str = str(cell.value).strip()
                    if re.search(r'\d{4}\s*\d/4[pP]?', cell_str):
                        quarter_cols_found += 1
        
        structure_info['quarter_columns_count'] = quarter_cols_found
        
        if quarter_cols_found == 0:
            warnings.append("분기별 데이터 열을 찾을 수 없음")
        elif quarter_cols_found < 4:
            warnings.append(f"분기별 데이터 열이 {quarter_cols_found}개만 발견됨 (최소 4개 권장)")
        
        # 지역 데이터 개수 확인
        region_count = 0
        region_names = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', 
                       '세종', '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        for row in range(data_start_row, min(data_start_row + 100, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=region_col)
            if cell.value:
                cell_str = str(cell.value).strip()
                if any(region in cell_str for region in region_names):
                    region_count += 1
        
        structure_info['region_count'] = region_count
        
        if region_count < 17:
            warnings.append(f"지역 데이터가 {region_count}개만 발견됨 (18개 지역 필요)")
        
    except Exception as e:
        warnings.append(f"구조 분석 중 오류: {str(e)}")
    
    return {
        'sheet_name': sheet_name,
        'structure': structure_info,
        'warnings': warnings,
        'has_warnings': len(warnings) > 0
    }


@app.route('/api/validate-template', methods=['POST'])
def validate_template():
    """템플릿의 마커를 검증하고 경고를 반환합니다."""
    try:
        excel_file = request.files.get('excel_file')
        excel_path = None
        
        if excel_file and excel_file.filename:
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다.'}), 400
            
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
        else:
            excel_path = DEFAULT_EXCEL_FILE
        
        template_name = request.form.get('template_name', '')
        
        if not template_name:
            return jsonify({'error': '템플릿을 선택해주세요.'}), 400
        
        # 템플릿 관리자 초기화
        template_path = Path('templates') / template_name
        if not template_path.exists():
            return jsonify({'error': f'템플릿 파일을 찾을 수 없습니다: {template_name}'}), 404
        
        template_manager = TemplateManager(str(template_path))
        template_manager.load_template()
        
        # 엑셀 시트 목록 가져오기
        excel_extractor = get_excel_extractor(excel_path)
        excel_sheets = excel_extractor.get_sheet_names()
        
        # 마커 검증
        validation_result = template_manager.validate_markers_against_excel(excel_sheets)
        
        # 마커 통계
        statistics = template_manager.get_marker_statistics()
        
        excel_extractor.close()
        
        # 경고 메시지 생성
        all_warnings = []
        if validation_result['missing_sheets']:
            for sheet in validation_result['missing_sheets']:
                all_warnings.append(f"시트 '{sheet}'를 엑셀 파일에서 찾을 수 없습니다.")
        
        all_warnings.extend(validation_result['warnings'])
        
        return jsonify({
            'success': True,
            'valid': validation_result['valid'],
            'missing_sheets': validation_result['missing_sheets'],
            'warnings': all_warnings,
            'has_warnings': len(all_warnings) > 0,
            'statistics': statistics,
            'sheet_summary': validation_result['sheet_summary']
        })
        
    except Exception as e:
        import traceback
        print(f"[ERROR] 템플릿 검증 중 오류: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': f'템플릿 검증 중 오류: {str(e)}'}), 500


@app.route('/api/check-structure', methods=['POST'])
def check_structure():
    """엑셀 파일의 구조를 확인하고 경고를 반환합니다."""
    try:
        excel_file = request.files.get('excel_file')
        excel_path = None
        
        if excel_file and excel_file.filename:
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다.'}), 400
            
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
        else:
            excel_path = DEFAULT_EXCEL_FILE
        
        template_name = request.form.get('template_name', '')
        
        if not template_name:
            return jsonify({'error': '템플릿을 선택해주세요.'}), 400
        
        # 엑셀 추출기 초기화
        excel_extractor = get_excel_extractor(excel_path)
        
        # 템플릿 관리자 초기화
        template_path = Path('templates') / template_name
        template_manager = TemplateManager(str(template_path))
        template_manager.load_template()
        
        # 템플릿에서 필요한 시트 목록 추출
        markers = template_manager.extract_markers()
        required_sheets = set()
        for marker in markers:
            sheet_name = marker.get('sheet_name', '').strip()
            if sheet_name:
                required_sheets.add(sheet_name)
        
        # 각 시트의 구조 확인
        all_results = []
        all_warnings = []
        flexible_mapper = FlexibleMapper(excel_extractor)
        
        for sheet_name in required_sheets:
            actual_sheet = flexible_mapper.find_sheet_by_name(sheet_name)
            if actual_sheet:
                result = detect_structure_changes(excel_extractor, actual_sheet)
                all_results.append(result)
                if result['has_warnings']:
                    all_warnings.extend([f"[{actual_sheet}] {w}" for w in result['warnings']])
            else:
                all_warnings.append(f"시트를 찾을 수 없음: {sheet_name}")
        
        excel_extractor.close()
        
        return jsonify({
            'success': True,
            'sheets_analyzed': len(all_results),
            'results': all_results,
            'all_warnings': all_warnings,
            'has_warnings': len(all_warnings) > 0
        })
        
    except Exception as e:
        import traceback
        print(f"[ERROR] 구조 확인 중 오류: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': f'구조 확인 중 오류: {str(e)}'}), 500


@app.route('/api/check-missing-values', methods=['POST'])
def check_missing_values():
    """결측치 확인 API - 스케일 기반 기본값 제공"""
    try:
        excel_file = request.files.get('excel_file')
        excel_path = None
        
        if excel_file and excel_file.filename:
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다.'}), 400
            
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
        else:
            excel_path = DEFAULT_EXCEL_FILE
        
        template_name = request.form.get('template_name', '')
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        
        if not template_name or not year_str or not quarter_str:
            return jsonify({'error': '템플릿, 연도, 분기를 모두 입력해주세요.'}), 400
        
        year = int(year_str)
        quarter = int(quarter_str)
        
        # 엑셀 추출기 초기화
        excel_extractor = get_excel_extractor(excel_path)
        
        # 템플릿 관리자 초기화
        template_path = Path('templates') / template_name
        template_manager = TemplateManager(str(template_path))
        template_manager.load_template()
        
        # 템플릿에서 필요한 시트 목록 추출
        markers = template_manager.extract_markers()
        required_sheets = set()
        for marker in markers:
            sheet_name = marker.get('sheet_name', '').strip()
            if sheet_name:
                required_sheets.add(sheet_name)
        
        # 각 시트에서 결측치 찾기
        all_missing_values = []
        flexible_mapper = FlexibleMapper(excel_extractor)
        
        for sheet_name in required_sheets:
            actual_sheet = flexible_mapper.find_sheet_by_name(sheet_name)
            if actual_sheet:
                missing = find_missing_values_in_sheet(excel_extractor, actual_sheet, year, quarter)
                all_missing_values.extend(missing)
        
        excel_extractor.close()
        
        return jsonify({
            'success': True,
            'missing_values': all_missing_values,
            'has_missing': len(all_missing_values) > 0,
            'count': len(all_missing_values)
        })
    except Exception as e:
        import traceback
        print(f"[ERROR] 결측치 확인 중 오류: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': f'결측치 확인 중 오류: {str(e)}'}), 500


@app.route('/api/process', methods=['POST'])
def process_template():
    """엑셀 파일과 템플릿을 처리하여 결과 생성"""
    try:
        # 엑셀 파일 처리: 업로드된 파일이 있으면 사용, 없으면 기본 파일 사용
        excel_file = request.files.get('excel_file')
        excel_path = None
        use_default_file = False
        
        if excel_file and excel_file.filename:
            # 업로드된 파일이 있는 경우
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다. (.xlsx, .xls만 가능)'}), 400
            
            # 엑셀 파일 저장
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            # 업로드 폴더가 존재하는지 확인하고 없으면 생성
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            try:
                upload_folder.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                return jsonify({'error': f'업로드 폴더 생성 실패: {str(e)}'}), 500
            
            excel_path = upload_folder / excel_filename
            
            # 파일 저장 시도
            try:
                excel_file.save(str(excel_path))
            except Exception as e:
                return jsonify({'error': f'파일 저장 중 오류가 발생했습니다: {str(e)}'}), 500
            
            # 파일이 제대로 저장되었는지 확인 (크기도 확인)
            if not excel_path.exists():
                return jsonify({'error': '파일 저장에 실패했습니다.'}), 400
            
            # 파일 크기 확인 (0바이트 파일 방지)
            if excel_path.stat().st_size == 0:
                return jsonify({'error': '저장된 파일이 비어있습니다.'}), 400
        else:
            # 기본 엑셀 파일 사용
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'error': f'기본 엑셀 파일을 찾을 수 없습니다: {DEFAULT_EXCEL_FILE.name}'}), 400
            
            excel_path = DEFAULT_EXCEL_FILE
            use_default_file = True
        
        # 템플릿명, 연도 및 분기 파라미터 가져오기
        template_name = request.form.get('template_name', '')
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        missing_values_str = request.form.get('missing_values', '{}')
        
        # 사용자가 입력한 결측치 값 파싱
        try:
            import json
            user_missing_values = json.loads(missing_values_str) if missing_values_str else {}
        except json.JSONDecodeError:
            user_missing_values = {}
        
        if not template_name:
            return jsonify({'error': '템플릿을 선택해주세요.'}), 400
        
        template_path = Path('templates') / template_name
        
        if not template_path.exists():
            return jsonify({'error': f'템플릿 파일을 찾을 수 없습니다: {template_name}'}), 404
        
        try:
            # 템플릿 관리자 초기화
            template_manager = TemplateManager(str(template_path))
            template_manager.load_template()
            
            # 템플릿에서 필요한 시트 목록 추출
            markers = template_manager.extract_markers()
            required_sheets = set()
            for marker in markers:
                sheet_name = marker.get('sheet_name', '').strip()
                if sheet_name:
                    required_sheets.add(sheet_name)
            
            if not required_sheets:
                return jsonify({
                    'error': '템플릿에서 필요한 시트를 찾을 수 없습니다.'
                }), 400
            
            # 파일 경로 검증
            if not excel_path.exists():
                return jsonify({'error': f'엑셀 파일을 찾을 수 없습니다: {excel_path}'}), 400
            
            # 파일 크기 확인
            file_size = excel_path.stat().st_size
            if file_size == 0:
                return jsonify({'error': '엑셀 파일이 비어있습니다.'}), 400
            
            # 엑셀 추출기 초기화 (캐시 사용)
            excel_extractor = get_excel_extractor(excel_path)
            
            # 사용 가능한 시트 목록 가져오기
            sheet_names = excel_extractor.get_sheet_names()
            
            # 필요한 시트가 모두 존재하는지 확인
            missing_sheets = []
            actual_sheet_mapping = {}  # 템플릿 시트명 -> 실제 시트명 매핑
            
            flexible_mapper = FlexibleMapper(excel_extractor)
            for required_sheet in required_sheets:
                # 유연한 매핑으로 실제 시트 찾기
                actual_sheet = flexible_mapper.find_sheet_by_name(required_sheet)
                if actual_sheet:
                    actual_sheet_mapping[required_sheet] = actual_sheet
                else:
                    missing_sheets.append(required_sheet)
            
            if missing_sheets:
                return jsonify({
                    'error': f'필요한 시트를 찾을 수 없습니다: {", ".join(missing_sheets)}. 사용 가능한 시트: {", ".join(sheet_names)}'
                }), 400
            
            # 첫 번째 필요한 시트를 기본 시트로 사용 (연도/분기 감지용)
            primary_sheet = list(actual_sheet_mapping.values())[0]
            
            # 연도 및 분기 자동 감지 또는 사용자 입력값 사용
            period_detector = PeriodDetector(excel_extractor)
            periods_info = period_detector.detect_available_periods(primary_sheet)
            
            if year_str and quarter_str:
                # 사용자가 입력한 값 사용
                year = int(year_str)
                quarter = int(quarter_str)
                
                # 유효성 검증
                is_valid, error_msg = period_detector.validate_period(primary_sheet, year, quarter)
                if not is_valid:
                    return jsonify({'error': error_msg}), 400
            else:
                # 자동 감지된 기본값 사용
                year = periods_info['default_year']
                quarter = periods_info['default_quarter']
            
            # 템플릿 필러 초기화 및 처리
            template_filler = TemplateFiller(template_manager, excel_extractor, schema_loader)
            
            # 사용자가 입력한 결측치 값 설정
            if user_missing_values:
                template_filler.set_missing_value_overrides(user_missing_values)
                print(f"[DEBUG] 사용자 결측치 값 설정: {len(user_missing_values)}개")
            
            # primary_sheet를 사용하여 연도/분기 감지 (템플릿은 자동으로 필요한 시트를 찾음)
            print(f"[DEBUG] 템플릿 채우기 시작: {template_name}, 연도={year}, 분기={quarter}, primary_sheet={primary_sheet}")
            try:
                filled_template = template_filler.fill_template(
                    sheet_name=primary_sheet,  # 연도/분기 감지용
                    year=year, 
                    quarter=quarter
                )
                print(f"[DEBUG] 템플릿 채우기 완료: {template_name}")
            except Exception as e:
                import traceback
                print(f"[ERROR] 템플릿 채우기 중 오류 발생: {str(e)}")
                print(f"[ERROR] 트레이스백:\n{traceback.format_exc()}")
                raise
            
            # display_name 찾기
            display_name = template_name.replace('.html', '')
            template_mapping = schema_loader.load_template_mapping()
            for sheet_name, info in template_mapping.items():
                if info['template'] == template_name:
                    display_name = info['display_name']
                    break
            
            # 결과 저장
            # 파일명 형식: (연도)년_(분기)분기_지역경제동향_보도자료(템플릿명).html
            output_filename = f"{year}년_{quarter}분기_지역경제동향_보도자료({display_name}).html"
            output_path = Path(app.config['OUTPUT_FOLDER']) / output_filename
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(filled_template)
            
            # 엑셀 파일 닫기
            excel_extractor.close()
            
            # 결과 반환
            return jsonify({
                'success': True,
                'output_filename': output_filename,
                'sheet_names': sheet_names,
                'message': '보도자료가 성공적으로 생성되었습니다.'
            })
            
        except FileNotFoundError as e:
            return jsonify({
                'error': f'엑셀 파일을 찾을 수 없습니다: {str(e)}'
            }), 400
        except PermissionError as e:
            return jsonify({
                'error': f'엑셀 파일에 접근할 수 없습니다. 파일이 다른 프로그램에서 사용 중일 수 있습니다: {str(e)}'
            }), 400
        except Exception as e:
            # 더 자세한 오류 정보 제공
            error_details = str(e)
            if 'BadZipFile' in str(type(e)) or 'zipfile' in str(e).lower():
                error_details = '엑셀 파일이 손상되었거나 올바른 형식이 아닙니다. 파일을 다시 확인해주세요.'
            elif 'openpyxl' in str(e).lower():
                error_details = f'엑셀 파일을 읽는 중 오류가 발생했습니다: {str(e)}'
            
            return jsonify({
                'error': f'처리 중 오류가 발생했습니다: {error_details}'
            }), 500
            
        finally:
            # 임시 엑셀 파일 삭제 (기본 파일이 아닌 경우에만)
            if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
                try:
                    excel_path.unlink()
                except Exception:
                    pass  # 파일 삭제 실패는 무시
    
    except Exception as e:
        return jsonify({
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@app.route('/api/download/<filename>')
def download_file(filename):
    """생성된 파일 다운로드"""
    try:
        return send_from_directory(
            app.config['OUTPUT_FOLDER'],
            filename,
            as_attachment=True
        )
    except Exception as e:
        return jsonify({'error': f'파일을 찾을 수 없습니다: {str(e)}'}), 404


@app.route('/api/preview/<filename>')
def preview_file(filename):
    """생성된 파일 미리보기"""
    try:
        file_path = Path(app.config['OUTPUT_FOLDER']) / filename
        if file_path.exists():
            return send_file(str(file_path))
        return jsonify({'error': '파일을 찾을 수 없습니다.'}), 404
    except Exception as e:
        return jsonify({'error': f'오류가 발생했습니다: {str(e)}'}), 500


# /api/check-default-file, /api/validate 엔드포인트는 routes/validation.py Blueprint로 이동


@app.route('/api/validate-excel', methods=['POST'])
def validate_excel_file():
    """엑셀 파일 유효성 검증 (시트 정보, 기간 정보 반환)"""
    try:
        # 엑셀 파일 처리: 업로드된 파일이 있으면 사용, 없으면 기본 파일 사용
        excel_file = request.files.get('excel_file')
        excel_path = None
        use_default_file = False
        
        if excel_file and excel_file.filename:
            # 업로드된 파일이 있는 경우
            if not allowed_file(excel_file.filename):
                return jsonify({'valid': False, 'error': '지원하지 않는 파일 형식입니다.'}), 400
            
            # 엑셀 파일 임시 저장 및 검증
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'valid': False, 'error': '파일명이 유효하지 않습니다.'}), 400
            
            # 업로드 폴더가 존재하는지 확인하고 없으면 생성
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            try:
                upload_folder.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                return jsonify({'valid': False, 'error': f'업로드 폴더 생성 실패: {str(e)}'}), 500
            
            excel_path = upload_folder / excel_filename
            
            # 파일 저장 시도
            try:
                excel_file.save(str(excel_path))
            except Exception as e:
                return jsonify({'valid': False, 'error': f'파일 저장 중 오류가 발생했습니다: {str(e)}'}), 500
            
            # 파일이 제대로 저장되었는지 확인 (크기도 확인)
            if not excel_path.exists():
                return jsonify({'valid': False, 'error': '파일 저장에 실패했습니다.'}), 400
            
            # 파일 크기 확인 (0바이트 파일 방지)
            if excel_path.stat().st_size == 0:
                return jsonify({'valid': False, 'error': '저장된 파일이 비어있습니다.'}), 400
        else:
            # 기본 엑셀 파일 사용
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'valid': False, 'error': f'기본 엑셀 파일을 찾을 수 없습니다: {DEFAULT_EXCEL_FILE.name}'}), 400
            
            excel_path = DEFAULT_EXCEL_FILE
            use_default_file = True
        
        try:
            # 파일 경로 검증
            if not excel_path.exists():
                return jsonify({
                    'valid': False,
                    'error': f'엑셀 파일을 찾을 수 없습니다: {excel_path}'
                }), 400
            
            # 파일 크기 확인
            file_size = excel_path.stat().st_size
            if file_size == 0:
                return jsonify({
                    'valid': False,
                    'error': '엑셀 파일이 비어있습니다.'
                }), 400
            
            excel_extractor = get_excel_extractor(excel_path)
            sheet_names = excel_extractor.get_sheet_names()
            
            if not sheet_names:
                excel_extractor.close()
                return jsonify({
                    'valid': False,
                    'error': '엑셀 파일에 시트가 없습니다.'
                }), 400
            
            # 성능 최적화: 첫 번째 시트에 대해서만 상세 정보 수집
            # 대부분의 시트가 동일한 연도/분기 범위를 가지므로,
            # 첫 번째 시트의 정보를 기본값으로 사용
            period_detector = PeriodDetector(excel_extractor)
            
            # 첫 번째 시트의 상세 정보 수집
            primary_sheet = sheet_names[0]
            periods_info = period_detector.detect_available_periods(primary_sheet)
            primary_sheet_info = {
                'min_year': periods_info['min_year'],
                'max_year': periods_info['max_year'],
                'default_year': periods_info['default_year'],
                'default_quarter': periods_info['default_quarter'],
                'available_periods': periods_info['available_periods']
            }
            
            # 나머지 시트는 기본값 사용 (성능 최적화)
            sheets_info = {}
            for sheet_name in sheet_names:
                if sheet_name == primary_sheet:
                    sheets_info[sheet_name] = primary_sheet_info
                else:
                    # 첫 번째 시트와 동일한 기본값 사용
                    sheets_info[sheet_name] = primary_sheet_info.copy()
            
            excel_extractor.close()
            
            return jsonify({
                'valid': True,
                'sheet_names': sheet_names,
                'sheets_info': sheets_info,
                'primary_sheet': primary_sheet,
                'message': '파일이 유효합니다.'
            })
        except FileNotFoundError as e:
            return jsonify({
                'valid': False,
                'error': f'엑셀 파일을 찾을 수 없습니다: {str(e)}'
            }), 400
        except PermissionError as e:
            return jsonify({
                'valid': False,
                'error': f'엑셀 파일에 접근할 수 없습니다. 파일이 다른 프로그램에서 사용 중일 수 있습니다: {str(e)}'
            }), 400
        except Exception as e:
            error_details = str(e)
            # 더 자세한 오류 정보 제공
            if 'BadZipFile' in str(type(e)) or 'zipfile' in str(e).lower():
                error_details = '엑셀 파일이 손상되었거나 올바른 형식이 아닙니다. 파일을 다시 확인해주세요.'
            elif 'openpyxl' in str(e).lower():
                error_details = f'엑셀 파일을 읽는 중 오류가 발생했습니다: {str(e)}'
            
            return jsonify({
                'valid': False,
                'error': f'엑셀 파일을 읽을 수 없습니다: {error_details}'
            }), 400
        finally:
            # 임시 엑셀 파일 삭제 (기본 파일이 아닌 경우에만)
            if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
                try:
                    excel_path.unlink()
                except Exception:
                    pass  # 파일 삭제 실패는 무시
    
    except Exception as e:
        return jsonify({
            'valid': False,
            'error': f'검증 중 오류가 발생했습니다: {str(e)}'
        }), 500


@app.route('/api/create-template', methods=['POST'])
def create_template():
    """이미지와 엑셀 파일에서 템플릿 생성 (헤더 기반 마커)"""
    try:
        # 파일 검증
        if 'image_file' not in request.files:
            return jsonify({'error': '이미지 파일이 없습니다.'}), 400
        
        image_file = request.files['image_file']
        if image_file.filename == '':
            return jsonify({'error': '이미지 파일을 선택해주세요.'}), 400
        
        if not allowed_image_file(image_file.filename):
            return jsonify({'error': '지원하지 않는 이미지 형식입니다. (png, jpg, jpeg, gif, bmp, webp만 가능)'}), 400
        
        # 엑셀 파일 검증 (선택사항이지만 권장)
        excel_file = request.files.get('excel_file')
        excel_path = None
        header_parser = None
        
        if excel_file and excel_file.filename:
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 엑셀 파일 형식입니다. (.xlsx, .xls만 가능)'}), 400
            
            # 엑셀 파일 저장
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            # 업로드 폴더가 존재하는지 확인하고 없으면 생성
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
            
            # 파일이 제대로 저장되었는지 확인
            if not excel_path.exists():
                return jsonify({'error': '파일 저장에 실패했습니다.'}), 400
        
        # 템플릿 이름 가져오기
        template_name = request.form.get('template_name', '').strip()
        if not template_name:
            # 파일명에서 확장자 제거하여 템플릿 이름으로 사용
            template_name = Path(image_file.filename).stem
        
        # 시트명 가져오기
        sheet_name = request.form.get('sheet_name', '').strip()
        if not sheet_name and excel_path:
            # 엑셀 파일이 있으면 첫 번째 시트 사용
            try:
                extractor = get_excel_extractor(excel_path)
                sheet_names = extractor.get_sheet_names()
                if sheet_names:
                    sheet_name = sheet_names[0]
            except:
                pass
        
        if not sheet_name:
            sheet_name = '시트1'
        
        # 이미지 파일 저장
        image_filename = secure_filename(image_file.filename)
        if not image_filename:
            return jsonify({'error': '이미지 파일명이 유효하지 않습니다.'}), 400
        
        # 업로드 폴더가 존재하는지 확인하고 없으면 생성
        upload_folder = Path(app.config['UPLOAD_FOLDER'])
        upload_folder.mkdir(parents=True, exist_ok=True)
        
        image_path = upload_folder / image_filename
        image_file.save(str(image_path))
        
        # 파일이 제대로 저장되었는지 확인
        if not image_path.exists():
            return jsonify({'error': '이미지 파일 저장에 실패했습니다.'}), 400
        
        try:
            # 엑셀 헤더 파서 초기화 (엑셀 파일이 있는 경우)
            if excel_path and excel_path.exists():
                header_parser = ExcelHeaderParser(str(excel_path))
            
            # 템플릿 생성기 초기화
            generator = TemplateGenerator(use_easyocr=True)
            
            # HTML 템플릿 생성 (헤더 파서 전달)
            html_template = generator.generate_html_template_from_excel(
                str(image_path),
                template_name,
                sheet_name,
                header_parser
            )
            
            # 템플릿 저장
            templates_dir = Path('templates')
            templates_dir.mkdir(exist_ok=True)
            
            template_filename = secure_filename(template_name) + '.html'
            template_path = templates_dir / template_filename
            
            with open(template_path, 'w', encoding='utf-8') as f:
                f.write(html_template)
            
            # 임시 파일 삭제
            if image_path.exists():
                image_path.unlink()
            if excel_path and excel_path.exists():
                excel_path.unlink()
            if header_parser:
                header_parser.close()
            
            return jsonify({
                'success': True,
                'template_name': template_filename,
                'template_path': str(template_path),
                'message': f'템플릿 "{template_name}"이 성공적으로 생성되었습니다.'
            })
            
        except Exception as e:
            # 임시 파일 삭제
            if image_path.exists():
                image_path.unlink()
            if excel_path and excel_path.exists():
                excel_path.unlink()
            if header_parser:
                header_parser.close()
            return jsonify({
                'error': f'템플릿 생성 중 오류가 발생했습니다: {str(e)}'
            }), 500
    
    except Exception as e:
        return jsonify({
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@app.route('/api/generate-pdf', methods=['POST'])
def generate_pdf():
    """10개 템플릿을 순서대로 처리하여 10페이지 PDF 생성"""
    try:
        # 엑셀 파일 처리: 업로드된 파일이 있으면 사용, 없으면 기본 파일 사용
        excel_file = request.files.get('excel_file')
        excel_path = None
        use_default_file = False
        
        if excel_file and excel_file.filename:
            # 업로드된 파일이 있는 경우
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다. (.xlsx, .xls만 가능)'}), 400
            
            # 엑셀 파일 저장
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            # 업로드 폴더가 존재하는지 확인하고 없으면 생성
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
            
            # 파일이 제대로 저장되었는지 확인
            if not excel_path.exists():
                return jsonify({'error': '파일 저장에 실패했습니다.'}), 400
        else:
            # 기본 엑셀 파일 사용
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'error': f'기본 엑셀 파일을 찾을 수 없습니다: {DEFAULT_EXCEL_FILE.name}'}), 400
            
            excel_path = DEFAULT_EXCEL_FILE
            use_default_file = True
        
        # 연도 및 분기 파라미터 가져오기
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        
        if not year_str or not quarter_str:
            return jsonify({'error': '연도와 분기를 입력해주세요.'}), 400
        
        year = int(year_str)
        quarter = int(quarter_str)
        
        # PDF 생성기 초기화
        pdf_generator = PDFGenerator(app.config['OUTPUT_FOLDER'])
        
        # PDF 생성 라이브러리 확인
        is_available, error_msg = pdf_generator.check_pdf_generator_available()
        if not is_available:
            return jsonify({'error': error_msg}), 500
        
        # PDF 생성
        success, result = pdf_generator.generate_pdf(
            excel_path=str(excel_path),
            year=year,
            quarter=quarter,
            templates_dir='templates'
        )
        
        # 임시 엑셀 파일 삭제 (기본 파일이 아닌 경우에만)
        if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
            try:
                excel_path.unlink()
            except Exception:
                pass  # 파일 삭제 실패는 무시
        
        if success:
            return jsonify(result)
        else:
            return jsonify({'error': result}), 500
    
    except Exception as e:
        return jsonify({
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@app.route('/api/generate-docx', methods=['POST'])
def generate_docx():
    """10개 템플릿을 순서대로 처리하여 DOCX 생성"""
    try:
        # 엑셀 파일 처리: 업로드된 파일이 있으면 사용, 없으면 기본 파일 사용
        excel_file = request.files.get('excel_file')
        excel_path = None
        use_default_file = False
        
        if excel_file and excel_file.filename:
            # 업로드된 파일이 있는 경우
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다. (.xlsx, .xls만 가능)'}), 400
            
            # 엑셀 파일 저장
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            # 업로드 폴더가 존재하는지 확인하고 없으면 생성
            upload_folder = Path(app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
            
            # 파일이 제대로 저장되었는지 확인
            if not excel_path.exists():
                return jsonify({'error': '파일 저장에 실패했습니다.'}), 400
        else:
            # 기본 엑셀 파일 사용
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'error': f'기본 엑셀 파일을 찾을 수 없습니다: {DEFAULT_EXCEL_FILE.name}'}), 400
            
            excel_path = DEFAULT_EXCEL_FILE
            use_default_file = True
        
        # 연도 및 분기 파라미터 가져오기
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        missing_values_str = request.form.get('missing_values', '{}')
        
        if not year_str or not quarter_str:
            return jsonify({'error': '연도와 분기를 입력해주세요.'}), 400
        
        year = int(year_str)
        quarter = int(quarter_str)
        
        # 결측치 값 파싱
        try:
            user_missing_values = json.loads(missing_values_str) if missing_values_str else {}
        except json.JSONDecodeError:
            user_missing_values = {}
        
        # DOCX 생성기 초기화
        docx_generator = DOCXGenerator(app.config['OUTPUT_FOLDER'])
        
        # DOCX 생성 라이브러리 확인
        is_available, error_msg = docx_generator.check_docx_generator_available()
        if not is_available:
            return jsonify({'error': error_msg}), 500
        
        # DOCX 생성
        success, result = docx_generator.generate_docx(
            excel_path=str(excel_path),
            year=year,
            quarter=quarter,
            templates_dir='templates',
            missing_values=user_missing_values if user_missing_values else None
        )
        
        # 임시 엑셀 파일 삭제 (기본 파일이 아닌 경우에만)
        if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
            try:
                excel_path.unlink()
            except Exception:
                pass  # 파일 삭제 실패는 무시
        
        if success:
            return jsonify(result)
        else:
            return jsonify({'error': result}), 500
    
    except Exception as e:
        return jsonify({
            'error': f'서버 오류가 발생했습니다: {str(e)}'
        }), 500


@app.errorhandler(413)
def request_entity_too_large(error):
    """파일 크기 제한 초과 에러 처리"""
    return jsonify({
        'error': '파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.'
    }), 413


if __name__ == '__main__':
    # 출력 폴더 생성
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    
    print("=" * 50)
    print("지역경제동향 보도자료 자동생성")
    print("=" * 50)
    print(f"웹 서버가 시작되었습니다.")
    print(f"브라우저에서 http://localhost:8000 을 열어주세요.")
    print(f"최대 파일 크기: 100MB")
    print("=" * 50)
    
    app.run(debug=True, host='0.0.0.0', port=8000)

