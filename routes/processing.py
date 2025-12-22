"""
처리 관련 라우트
/api/process, /api/check-missing-values
"""

import json
import re
from pathlib import Path
from flask import Blueprint, request, jsonify, current_app
from werkzeug.utils import secure_filename

from .common import (
    allowed_file, get_excel_extractor, get_template_manager,
    get_schema_loader, detect_sheet_scale, DEFAULT_EXCEL_FILE
)

processing_bp = Blueprint('processing', __name__, url_prefix='/api')


def find_missing_values_in_sheet(excel_extractor, sheet_name, year, quarter):
    """시트에서 결측치 찾기"""
    missing_values = []
    
    try:
        sheet = excel_extractor.get_sheet(sheet_name)
        
        # 분기별 열 찾기
        target_col = None
        prev_col = None
        
        for col in range(1, min(150, sheet.max_column + 1)):
            cell_value = sheet.cell(row=3, column=col).value
            if cell_value:
                cell_str = str(cell_value).strip()
                target_pattern = f"{year}  {quarter}/4"
                prev_pattern = f"{year-1}  {quarter}/4"
                
                if target_pattern in cell_str or f"{year} {quarter}/4" in cell_str:
                    target_col = col
                if prev_pattern in cell_str or f"{year-1} {quarter}/4" in cell_str:
                    prev_col = col
        
        if not target_col:
            return []
        
        for row in range(4, min(200, sheet.max_row + 1)):
            region_cell = sheet.cell(row=row, column=2)
            category_cell = sheet.cell(row=row, column=6) if sheet.max_column >= 6 else None
            
            region = str(region_cell.value).strip() if region_cell.value else ''
            category = str(category_cell.value).strip() if category_cell and category_cell.value else '합계'
            
            if not region or region == '전국':
                continue
            
            current_cell = sheet.cell(row=row, column=target_col)
            current_value = current_cell.value
            
            is_missing = False
            if current_value is None:
                is_missing = True
            elif isinstance(current_value, str):
                stripped = current_value.strip()
                if not stripped or stripped == '-':
                    is_missing = True
            
            if is_missing:
                default_value = detect_sheet_scale(excel_extractor, sheet_name)
                
                missing_values.append({
                    'sheet': sheet_name,
                    'region': region,
                    'category': category,
                    'year': year,
                    'quarter': quarter,
                    'row': row,
                    'col': target_col,
                    'default_value': default_value
                })
    except Exception as e:
        print(f"[WARNING] 결측치 검색 중 오류 ({sheet_name}): {str(e)}")
    
    return missing_values


@processing_bp.route('/check-missing-values', methods=['POST'])
def check_missing_values():
    """결측치 확인"""
    try:
        excel_file = request.files.get('excel_file')
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        template_name = request.form.get('template_name', '')
        
        if not year_str or not quarter_str:
            return jsonify({'error': '연도와 분기를 입력해주세요.'}), 400
        
        year = int(year_str)
        quarter = int(quarter_str)
        
        # 엑셀 파일 처리
        if excel_file and excel_file.filename:
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다.'}), 400
            
            excel_filename = secure_filename(excel_file.filename)
            upload_folder = Path(current_app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
        else:
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'error': f'기본 엑셀 파일을 찾을 수 없습니다.'}), 400
            excel_path = DEFAULT_EXCEL_FILE
        
        excel_extractor = get_excel_extractor(excel_path)
        
        # 템플릿에서 필요한 시트 목록 추출
        template_path = Path('templates') / template_name
        if not template_path.exists():
            return jsonify({'error': f'템플릿 파일을 찾을 수 없습니다.'}), 404
        
        template_manager = get_template_manager(template_path)
        markers = template_manager.extract_markers()
        required_sheets = set()
        for marker in markers:
            sheet_name = marker.get('sheet_name', '').strip()
            if sheet_name:
                required_sheets.add(sheet_name)
        
        # 각 시트에서 결측치 찾기
        from src.flexible_mapper import FlexibleMapper
        
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


@processing_bp.route('/process', methods=['POST'])
def process_template():
    """엑셀 파일과 템플릿을 처리하여 결과 생성"""
    try:
        # 엑셀 파일 처리
        excel_file = request.files.get('excel_file')
        excel_path = None
        use_default_file = False
        
        if excel_file and excel_file.filename:
            if not allowed_file(excel_file.filename):
                return jsonify({'error': '지원하지 않는 파일 형식입니다.'}), 400
            
            excel_filename = secure_filename(excel_file.filename)
            if not excel_filename:
                return jsonify({'error': '파일명이 유효하지 않습니다.'}), 400
            
            upload_folder = Path(current_app.config['UPLOAD_FOLDER'])
            upload_folder.mkdir(parents=True, exist_ok=True)
            
            excel_path = upload_folder / excel_filename
            excel_file.save(str(excel_path))
            
            if not excel_path.exists() or excel_path.stat().st_size == 0:
                return jsonify({'error': '파일 저장에 실패했습니다.'}), 400
        else:
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'error': f'기본 엑셀 파일을 찾을 수 없습니다.'}), 400
            excel_path = DEFAULT_EXCEL_FILE
            use_default_file = True
        
        # 파라미터 가져오기
        template_name = request.form.get('template_name', '')
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        missing_values_str = request.form.get('missing_values', '{}')
        
        try:
            user_missing_values = json.loads(missing_values_str) if missing_values_str else {}
        except json.JSONDecodeError:
            user_missing_values = {}
        
        if not template_name:
            return jsonify({'error': '템플릿을 선택해주세요.'}), 400
        
        template_path = Path('templates') / template_name
        if not template_path.exists():
            return jsonify({'error': f'템플릿 파일을 찾을 수 없습니다.'}), 404
        
        # 템플릿 처리
        from src.core.template_filler import TemplateFiller
        from src.flexible_mapper import FlexibleMapper
        from src.analyzers.period_detector import PeriodDetector
        
        template_manager = get_template_manager(template_path)
        markers = template_manager.extract_markers()
        
        required_sheets = set()
        for marker in markers:
            sheet_name = marker.get('sheet_name', '').strip()
            if sheet_name:
                required_sheets.add(sheet_name)
        
        if not required_sheets:
            return jsonify({'error': '템플릿에서 필요한 시트를 찾을 수 없습니다.'}), 400
        
        excel_extractor = get_excel_extractor(excel_path)
        sheet_names = excel_extractor.get_sheet_names()
        
        # 연도 및 분기 검증
        year = int(year_str) if year_str else None
        quarter = int(quarter_str) if quarter_str else None
        
        if year and quarter:
            period_detector = PeriodDetector(excel_extractor)
            primary_sheet = sheet_names[0] if sheet_names else None
            
            if primary_sheet:
                is_valid, error_msg = period_detector.validate_period(primary_sheet, year, quarter)
                if not is_valid:
                    excel_extractor.close()
                    return jsonify({'error': error_msg}), 400
        else:
            period_detector = PeriodDetector(excel_extractor)
            primary_sheet = sheet_names[0] if sheet_names else None
            
            if primary_sheet:
                periods_info = period_detector.detect_available_periods(primary_sheet)
                year = periods_info.get('default_year', 2025)
                quarter = periods_info.get('default_quarter', 2)
        
        # 시트 매핑
        flexible_mapper = FlexibleMapper(excel_extractor)
        actual_sheet_mapping = {}
        
        for required_sheet in required_sheets:
            actual_sheet = flexible_mapper.find_sheet_by_name(required_sheet)
            if actual_sheet:
                actual_sheet_mapping[required_sheet] = actual_sheet
        
        # 템플릿 필러 초기화
        schema_loader = get_schema_loader()
        template_filler = TemplateFiller(template_manager, excel_extractor, schema_loader)
        
        if user_missing_values:
            template_filler.set_missing_value_overrides(user_missing_values)
        
        # 첫 번째 시트로 처리
        primary_sheet = list(actual_sheet_mapping.values())[0] if actual_sheet_mapping else None
        
        filled_template = template_filler.fill_template(
            sheet_name=primary_sheet,
            year=year,
            quarter=quarter
        )
        
        excel_extractor.close()
        
        # 결과 반환
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(filled_template, 'html.parser')
        body = soup.find('body')
        style = soup.find('style')
        
        return jsonify({
            'success': True,
            'html': str(body) if body else filled_template,
            'style': str(style) if style else '',
            'year': year,
            'quarter': quarter
        })
        
    except Exception as e:
        import traceback
        print(f"[ERROR] 템플릿 처리 중 오류: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': f'처리 중 오류: {str(e)}'}), 500

