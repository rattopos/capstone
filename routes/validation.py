"""
검증 관련 라우트
/api/validate, /api/test-all-templates, /api/check-default-file
"""

import re
from pathlib import Path
from collections import defaultdict
from flask import Blueprint, request, jsonify, current_app
from werkzeug.utils import secure_filename

from .common import (
    allowed_file, get_excel_extractor, get_template_manager,
    get_schema_loader, DEFAULT_EXCEL_FILE
)

validation_bp = Blueprint('validation', __name__, url_prefix='/api')


@validation_bp.route('/check-default-file', methods=['GET'])
def check_default_file():
    """기본 엑셀 파일 존재 여부 확인"""
    if DEFAULT_EXCEL_FILE.exists():
        return jsonify({
            'exists': True,
            'filename': DEFAULT_EXCEL_FILE.name
        })
    else:
        return jsonify({
            'exists': False,
            'message': '기본 엑셀 파일이 없습니다. 파일을 업로드해주세요.'
        })


@validation_bp.route('/validate', methods=['POST'])
def validate_template():
    """템플릿 유효성 검증"""
    try:
        from src.flexible_mapper import FlexibleMapper
        from src.analyzers.period_detector import PeriodDetector
        
        excel_file = request.files.get('excel_file')
        template_name = request.form.get('template_name', '')
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        
        if not template_name:
            return jsonify({'error': '템플릿을 선택해주세요.'}), 400
        
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
        
        excel_extractor = get_excel_extractor(excel_path)
        sheet_names = excel_extractor.get_sheet_names()
        
        flexible_mapper = FlexibleMapper(excel_extractor)
        
        missing_sheets = []
        found_sheets = {}
        for required_sheet in required_sheets:
            actual_sheet = flexible_mapper.find_sheet_by_name(required_sheet)
            if actual_sheet:
                found_sheets[required_sheet] = actual_sheet
            else:
                missing_sheets.append(required_sheet)
        
        # 연도/분기 검증
        year = int(year_str) if year_str else None
        quarter = int(quarter_str) if quarter_str else None
        
        periods_info = None
        period_valid = True
        period_error = None
        
        if found_sheets:
            primary_sheet = list(found_sheets.values())[0]
            period_detector = PeriodDetector(excel_extractor)
            periods_info = period_detector.detect_available_periods(primary_sheet)
            
            if year and quarter:
                period_valid, period_error = period_detector.validate_period(primary_sheet, year, quarter)
        
        excel_extractor.close()
        
        return jsonify({
            'success': True,
            'template_name': template_name,
            'required_sheets': list(required_sheets),
            'found_sheets': found_sheets,
            'missing_sheets': missing_sheets,
            'available_sheets': sheet_names,
            'periods_info': periods_info,
            'period_valid': period_valid,
            'period_error': period_error,
            'marker_count': len(markers)
        })
        
    except Exception as e:
        import traceback
        print(f"[ERROR] 검증 중 오류: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': f'검증 중 오류: {str(e)}'}), 500


@validation_bp.route('/test-all-templates', methods=['POST'])
def test_all_templates():
    """모든 템플릿을 실행하고 N/A 값을 분석"""
    try:
        from src.core.template_filler import TemplateFiller
        from src.flexible_mapper import FlexibleMapper
        from src.analyzers.period_detector import PeriodDetector
        
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
            
            if not excel_path.exists():
                return jsonify({'error': '파일 저장에 실패했습니다.'}), 400
        else:
            if not DEFAULT_EXCEL_FILE.exists():
                return jsonify({'error': f'기본 엑셀 파일을 찾을 수 없습니다.'}), 400
            excel_path = DEFAULT_EXCEL_FILE
            use_default_file = True
        
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        
        if not year_str or not quarter_str:
            return jsonify({'error': '연도와 분기를 입력해주세요.'}), 400
        
        year = int(year_str)
        quarter = int(quarter_str)
        
        templates_dir = Path('templates')
        if not templates_dir.exists():
            return jsonify({'error': '템플릿 디렉토리를 찾을 수 없습니다.'}), 404
        
        template_files = list(templates_dir.glob('*.html'))
        if not template_files:
            return jsonify({'error': '템플릿 파일이 없습니다.'}), 404
        
        results = []
        excel_extractor = get_excel_extractor(excel_path)
        schema_loader = get_schema_loader()
        
        try:
            sheet_names = excel_extractor.get_sheet_names()
            
            period_detector = PeriodDetector(excel_extractor)
            if sheet_names:
                is_valid, period_error = period_detector.validate_period(sheet_names[0], year, quarter)
                if not is_valid:
                    excel_extractor.close()
                    return jsonify({'error': period_error}), 400
            
            flexible_mapper = FlexibleMapper(excel_extractor)
            
            for template_path in template_files:
                template_name = template_path.name
                
                if template_name in {'서울.html', '서울주요지표.html'}:
                    continue
                
                try:
                    template_manager = get_template_manager(template_path)
                    markers = template_manager.extract_markers()
                    
                    required_sheets = set()
                    for marker in markers:
                        sheet_name = marker.get('sheet_name', '').strip()
                        if sheet_name:
                            required_sheets.add(sheet_name)
                    
                    if not required_sheets:
                        results.append({
                            'template': template_name,
                            'status': 'error',
                            'error': '필요한 시트를 찾을 수 없습니다.',
                            'na_count': 0,
                            'na_markers': []
                        })
                        continue
                    
                    actual_sheet = None
                    for sheet_name in required_sheets:
                        actual_sheet = flexible_mapper.find_sheet_by_name(sheet_name)
                        if actual_sheet:
                            break
                    
                    if not actual_sheet:
                        results.append({
                            'template': template_name,
                            'status': 'error',
                            'error': f'필요한 시트를 찾을 수 없습니다: {required_sheets}',
                            'na_count': 0,
                            'na_markers': []
                        })
                        continue
                    
                    # 템플릿 필러로 처리
                    template_filler = TemplateFiller(template_manager, excel_extractor, schema_loader)
                    filled_template = template_filler.fill_template(
                        sheet_name=actual_sheet,
                        year=year,
                        quarter=quarter
                    )
                    
                    # N/A 값 분석
                    na_pattern = re.compile(r'\bN/A\b')
                    na_matches = na_pattern.findall(filled_template)
                    na_count = len(na_matches)
                    
                    # N/A가 포함된 마커 찾기
                    na_markers = []
                    for marker in markers:
                        full_match = marker.get('full_match', '')
                        if full_match not in filled_template and 'N/A' in filled_template:
                            na_markers.append(full_match)
                    
                    results.append({
                        'template': template_name,
                        'status': 'success' if na_count == 0 else 'warning',
                        'na_count': na_count,
                        'na_markers': na_markers[:10],
                        'total_markers': len(markers)
                    })
                    
                except Exception as e:
                    results.append({
                        'template': template_name,
                        'status': 'error',
                        'error': str(e),
                        'na_count': 0,
                        'na_markers': []
                    })
            
            excel_extractor.close()
            
            if not use_default_file and excel_path and excel_path.exists() and excel_path != DEFAULT_EXCEL_FILE:
                try:
                    excel_path.unlink()
                except Exception:
                    pass
            
            return jsonify({
                'success': True,
                'results': results,
                'total_templates': len(results),
                'year': year,
                'quarter': quarter
            })
            
        except Exception as e:
            excel_extractor.close()
            raise e
    
    except Exception as e:
        import traceback
        print(f"[ERROR] 테스트 중 오류: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': f'서버 오류가 발생했습니다: {str(e)}'}), 500

