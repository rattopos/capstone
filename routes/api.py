# -*- coding: utf-8 -*-
"""
API 라우트
"""

import json
import base64
import re
from pathlib import Path
from urllib.parse import quote

from flask import Blueprint, request, jsonify, session, send_file, make_response
from werkzeug.utils import secure_filename
import unicodedata
import uuid

from config.settings import BASE_DIR, TEMPLATES_DIR, UPLOAD_FOLDER, EXPORT_FOLDER
from config.reports import REPORT_ORDER, REGIONAL_REPORTS, SUMMARY_REPORTS, STATISTICS_REPORTS
from utils.excel_utils import extract_year_quarter_from_excel, extract_year_quarter_from_raw, detect_file_type
from services.report_generator import (
    generate_report_html,
    generate_regional_report_html,
    generate_statistics_report_html,
    generate_individual_statistics_html
)
from services.grdp_service import (
    get_kosis_grdp_download_info, 
    parse_kosis_grdp_file,
    get_default_grdp_data,
    save_extracted_contributions
)
from services.excel_processor import preprocess_excel, check_available_methods, get_recommended_method
from data_converter import DataConverter
import openpyxl


def safe_filename(filename):
    """한글을 보존하면서 안전한 파일명 생성
    
    - 한글, 영문, 숫자, 언더스코어, 하이픈, 점 허용
    - 위험한 문자 제거
    - 파일명 충돌 방지를 위해 UUID 추가
    """
    # 파일명과 확장자 분리
    if '.' in filename:
        name, ext = filename.rsplit('.', 1)
        ext = '.' + ext.lower()
    else:
        name = filename
        ext = ''
    
    # 유니코드 정규화
    name = unicodedata.normalize('NFC', name)
    
    # 허용할 문자만 유지 (한글, 영문, 숫자, 언더스코어, 하이픈, 공백)
    safe_chars = []
    for char in name:
        if char.isalnum() or char in ('_', '-', ' ', '년', '분기'):
            safe_chars.append(char)
        elif '\uAC00' <= char <= '\uD7A3':  # 한글 완성형
            safe_chars.append(char)
        elif '\u3131' <= char <= '\u3163':  # 한글 자모
            safe_chars.append(char)
    
    name = ''.join(safe_chars).strip()
    
    # 공백을 언더스코어로
    name = name.replace(' ', '_')
    
    # 빈 파일명 방지
    if not name:
        name = 'upload'
    
    # 파일명 충돌 방지를 위해 짧은 UUID 추가
    short_uuid = str(uuid.uuid4())[:8]
    
    return f"{name}_{short_uuid}{ext}"


def send_file_with_korean_filename(filepath, filename, mimetype):
    """한글 파일명을 지원하는 파일 다운로드 응답 생성 (RFC 5987)"""
    response = make_response(send_file(filepath, mimetype=mimetype))
    
    # RFC 5987 방식으로 한글 파일명 인코딩
    encoded_filename = quote(filename, safe='')
    
    # Content-Disposition 헤더 설정 (ASCII fallback + UTF-8 filename)
    ascii_filename = filename.encode('ascii', 'ignore').decode('ascii') or 'download'
    response.headers['Content-Disposition'] = (
        f"attachment; filename=\"{ascii_filename}\"; "
        f"filename*=UTF-8''{encoded_filename}"
    )
    
    return response


api_bp = Blueprint('api', __name__, url_prefix='/api')


def cleanup_upload_folder(keep_current_files=True, cleanup_excel_only=True):
    """업로드 폴더 정리 (현재 세션 파일 제외)
    
    Args:
        keep_current_files: True면 현재 세션에서 사용 중인 파일은 보존
        cleanup_excel_only: True면 엑셀 파일만 정리 (HTML 등은 보존)
    """
    try:
        # 현재 세션에서 사용 중인 파일 목록
        protected_files = set()
        if keep_current_files:
            excel_path = session.get('excel_path')
            raw_excel_path = session.get('raw_excel_path')
            
            if excel_path:
                protected_files.add(Path(excel_path).name)
            if raw_excel_path:
                protected_files.add(Path(raw_excel_path).name)
        
        # 업로드 폴더의 모든 파일 확인
        deleted_count = 0
        for file_path in UPLOAD_FOLDER.glob('*'):
            if file_path.is_file():
                # 정리 대상인지 확인
                should_delete = False
                
                # 현재 세션 파일이 아닌 경우
                if file_path.name not in protected_files:
                    # 엑셀 파일만 정리하는 경우
                    if cleanup_excel_only:
                        if file_path.suffix.lower() in ['.xlsx', '.xls']:
                            should_delete = True
                    else:
                        # 디버그 파일이 아닌 경우 모두 삭제
                        if '디버그' not in file_path.name:
                            should_delete = True
                
                if should_delete:
                    try:
                        file_path.unlink()
                        deleted_count += 1
                        print(f"[정리] 파일 삭제: {file_path.name}")
                    except Exception as e:
                        print(f"[경고] 파일 삭제 실패 ({file_path.name}): {e}")
        
        if deleted_count > 0:
            print(f"[정리] 업로드 폴더 정리 완료: {deleted_count}개 파일 삭제")
        
        return deleted_count
        
    except Exception as e:
        print(f"[경고] 업로드 폴더 정리 중 오류: {e}")
        return 0


def _calculate_analysis_sheets(excel_path: str, preserve_formulas: bool = True):
    """분석 시트의 수식 계산 및 검증 (수식 보존 옵션)
    
    분석 시트의 수식은 집계 시트를 참조합니다.
    수식을 유지하면서 계산 결과를 로그로 출력합니다.
    
    Args:
        excel_path: 분석표 파일 경로
        preserve_formulas: True면 수식 유지 (엑셀에서 계산), False면 값으로 대체
    """
    # 분석 시트 → 집계 시트 매핑
    analysis_aggregate_mapping = {
        'A 분석': 'A(광공업생산)집계',
        'B 분석': 'B(서비스업생산)집계',
        'C 분석': 'C(소비)집계',
        'D(고용률)분석': 'D(고용률)집계',
        'D(실업)분석': 'D(실업)집계',
        'E(지출목적물가) 분석': 'E(지출목적물가)집계',
        'E(품목성질물가)분석': 'E(품목성질물가)집계',
        "F'분석": "F'(건설)집계",
        'G 분석': 'G(수출)집계',
        'H 분석': 'H(수입)집계',
    }
    
    wb = openpyxl.load_workbook(excel_path, data_only=False)
    
    calculated_count = 0
    formula_count = 0
    
    for analysis_sheet, aggregate_sheet in analysis_aggregate_mapping.items():
        if analysis_sheet not in wb.sheetnames:
            continue
        if aggregate_sheet not in wb.sheetnames:
            continue
        
        ws_analysis = wb[analysis_sheet]
        ws_aggregate = wb[aggregate_sheet]
        
        # 집계 시트를 dict로 캐싱 (빠른 조회용)
        aggregate_data = {}
        for row in ws_aggregate.iter_rows(min_row=1, max_row=ws_aggregate.max_row):
            for cell in row:
                if cell.value is not None:
                    aggregate_data[(cell.row, cell.column)] = cell.value
        
        # 분석 시트의 수식 셀 처리
        for row in ws_analysis.iter_rows(min_row=1, max_row=ws_analysis.max_row):
            for cell in row:
                if cell.value is None:
                    continue
                    
                val = str(cell.value)
                
                # 수식인 경우 (=로 시작)
                if val.startswith('='):
                    formula_count += 1
                    
                    # 집계 시트 참조 파싱: ='시트이름'!셀주소
                    import re
                    match = re.match(r"^='?([^'!]+)'?!([A-Z]+)(\d+)$", val)
                    if match:
                        ref_sheet = match.group(1)
                        ref_col_letter = match.group(2)
                        ref_row = int(match.group(3))
                        
                        # 열 문자를 숫자로 변환 (A=1, B=2, ...)
                        ref_col = 0
                        for i, c in enumerate(reversed(ref_col_letter)):
                            ref_col += (ord(c) - ord('A') + 1) * (26 ** i)
                        
                        # 집계 시트에서 값 가져오기
                        ref_value = aggregate_data.get((ref_row, ref_col))
                        if ref_value is not None:
                            calculated_count += 1
                            
                            if preserve_formulas:
                                # 수식 유지 (엑셀에서 열면 자동 계산)
                                # 계산 결과는 로그로만 출력
                                pass
                            else:
                                # 수식을 계산된 값으로 대체
                                cell.value = ref_value
    
    wb.save(excel_path)
    wb.close()
    
    if preserve_formulas:
        print(f"[분석표] 수식 보존 완료: {excel_path}")
        print(f"  → 총 {formula_count}개 수식 유지, {calculated_count}개 참조 값 확인")
        print(f"  → 엑셀에서 열면 수식이 자동으로 계산됩니다.")
    else:
        print(f"[분석표] 수식 계산 완료: {excel_path}")
        print(f"  → {calculated_count}개 셀 값으로 변환")


# ============================================================================
# 트랙 1: 기초자료 직접 처리 (주 워크플로우)
# ============================================================================

def _handle_raw_data_upload(filepath: Path, filename: str):
    """[트랙 1] 기초자료 수집표 → 직접 보도자료 생성
    
    기초자료에서 RawDataExtractor를 사용하여 직접 데이터를 추출합니다.
    분석표를 거치지 않으므로 처리 속도가 빠릅니다.
    """
    print(f"\n{'='*50}")
    print(f"[트랙 1] 기초자료 직접 처리: {filename}")
    print(f"{'='*50}")
    
    try:
        from templates.raw_data_extractor import RawDataExtractor
        from utils.excel_utils import extract_year_quarter_from_raw
        
        # 연도/분기 추출
        year, quarter = extract_year_quarter_from_raw(str(filepath))
        print(f"[정보] 감지된 연도/분기: {year}년 {quarter}분기")
        
        # RawDataExtractor 준비
        extractor = RawDataExtractor(str(filepath), year, quarter)
        
        # GRDP 데이터 추출 시도
        has_grdp = False
        grdp_data = None
        grdp_sheet_found = None
        needs_review = True
        
        try:
            converter = DataConverter(str(filepath))
            grdp_data = converter.extract_grdp_data()
            if grdp_data and not grdp_data.get('national_summary', {}).get('placeholder', True):
                has_grdp = True
                needs_review = False
                print(f"[GRDP] 추출 성공 - 전국: {grdp_data['national_summary']['growth_rate']}%")
                save_extracted_contributions(grdp_data)
        except Exception as e:
            print(f"[GRDP] 추출 실패: {e}")
        
        if grdp_data is None:
            # 기본값 사용 안 함 - GRDP 없으면 N/A로 표시
            has_grdp = False
            needs_review = True
            print(f"[GRDP] 데이터 없음 - 별도 파일 업로드 필요")
        
        # 가중치 검증
        weight_validation = validate_weights_in_raw_data(str(filepath))
        if weight_validation['has_missing_weights']:
            print(f"[가중치 경고] 누락: {weight_validation['total_missing']}건")
        
        print(f"[결과] 보도자료 생성 준비 완료")
        
    except Exception as e:
        import traceback
        print(f"[오류] 기초자료 처리 실패: {e}")
        traceback.print_exc()
        return jsonify({
            'success': False,
            'error': f'기초자료 처리 중 오류가 발생했습니다: {str(e)}'
        })
    
    # 세션에 저장 (트랙 1 전용)
    session['raw_excel_path'] = str(filepath)
    session['excel_path'] = str(filepath)
    session['year'] = year
    session['quarter'] = quarter
    session['file_type'] = 'raw_direct'
    
    if grdp_data:
        session['grdp_data'] = grdp_data
        grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
        try:
            with open(grdp_json_path, 'w', encoding='utf-8') as f:
                json.dump(grdp_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[경고] GRDP JSON 저장 실패: {e}")
    
    try:
        session['raw_file_mtime'] = Path(filepath).stat().st_mtime
    except OSError:
        pass
    
    return jsonify({
        'success': True,
        'filename': filename,
        'file_type': 'raw_direct',
        'year': year,
        'quarter': quarter,
        'reports': REPORT_ORDER,
        'regional_reports': REGIONAL_REPORTS,
        'conversion_info': None,
        'has_grdp': has_grdp,
        'grdp_sheet': grdp_sheet_found,
        'needs_grdp': not has_grdp,
        'needs_review': needs_review,
        'weight_validation': weight_validation,
        'preprocessing': {
            'success': True,
            'message': '기초자료 직접 처리',
            'method': 'raw_direct'
        }
    })


# ============================================================================
# 트랙 2: 분석표 처리 (레거시/폴백)
# ============================================================================

def _handle_analysis_upload(filepath: Path, filename: str):
    """[트랙 2] 분석표 업로드 → 보도자료 생성 (레거시/폴백)
    
    이미 생성된 분석표를 업로드하여 보도자료를 생성합니다.
    기존 워크플로우와의 호환성을 위해 유지됩니다.
    """
    print(f"\n{'='*50}")
    print(f"[트랙 2] 분석표 업로드 (레거시): {filename}")
    print(f"{'='*50}")
    
    # 수식 계산 전처리
    print(f"[전처리] 엑셀 수식 계산 시작...")
    processed_path, preprocess_success, preprocess_msg = preprocess_excel(str(filepath))
    
    if preprocess_success:
        print(f"[전처리] 성공: {preprocess_msg}")
        filepath = Path(processed_path)
    else:
        print(f"[전처리] {preprocess_msg} - fallback 사용")
    
    # 연도/분기 추출
    year, quarter = extract_year_quarter_from_excel(str(filepath))
    
    # GRDP 시트 확인
    has_grdp = False
    grdp_data = None
    grdp_sheet_found = None
    
    try:
        grdp_sheet_names = ['I GRDP', 'GRDP', 'grdp', 'I(GRDP)', '분기 GRDP']
        wb = openpyxl.load_workbook(str(filepath), read_only=True, data_only=True)
        
        for sheet_name in grdp_sheet_names:
            if sheet_name in wb.sheetnames:
                has_grdp = True
                grdp_sheet_found = sheet_name
                print(f"[GRDP] 시트 발견: {sheet_name}")
                
                grdp_data = _extract_grdp_from_analysis_sheet(wb[sheet_name], year, quarter)
                if grdp_data:
                    print(f"[GRDP] 추출 성공 - 전국: {grdp_data.get('national_summary', {}).get('growth_rate', 0)}%")
                break
        wb.close()
    except Exception as e:
        print(f"[경고] GRDP 시트 확인 실패: {e}")
    
    # 세션에 저장 (트랙 2 전용)
    session['analysis_excel_path'] = str(filepath)
    session['excel_path'] = str(filepath)
    session['year'] = year
    session['quarter'] = quarter
    session['file_type'] = 'analysis'
    
    try:
        session['excel_file_mtime'] = Path(filepath).stat().st_mtime
    except OSError:
        pass
    
    if grdp_data:
        session['grdp_data'] = grdp_data
        grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
        try:
            with open(grdp_json_path, 'w', encoding='utf-8') as f:
                json.dump(grdp_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"[경고] GRDP JSON 저장 실패: {e}")
    
    print(f"[결과] GRDP {'있음' if has_grdp else '없음'}")
    
    return jsonify({
        'success': True,
        'filename': filename,
        'file_type': 'analysis',
        'year': year,
        'quarter': quarter,
        'reports': REPORT_ORDER,
        'regional_reports': REGIONAL_REPORTS,
        'needs_grdp': not has_grdp,
        'has_grdp': has_grdp,
        'grdp_sheet': grdp_sheet_found,
        'conversion_info': None,
        'preprocessing': {
            'success': preprocess_success,
            'message': preprocess_msg,
            'method': get_recommended_method()
        }
    })


# ============================================================================
# 업로드 라우터 (트랙 분기)
# ============================================================================

@api_bp.route('/upload', methods=['POST'])
def upload_excel():
    """엑셀 파일 업로드 - 트랙 자동 분기
    
    파일 유형을 감지하여 적절한 트랙으로 분기합니다:
    - 트랙 1 (raw_direct): 기초자료 → 직접 보도자료 생성
    - 트랙 2 (analysis): 분석표 → 보도자료 생성 (레거시)
    """
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '파일이 없습니다'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '파일이 선택되지 않았습니다'})
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'error': '엑셀 파일만 업로드 가능합니다'})
    
    # 새 파일 업로드 전 이전 파일 정리
    cleanup_upload_folder(keep_current_files=False)
    
    # 파일 저장
    filename = safe_filename(file.filename)
    filepath = Path(UPLOAD_FOLDER) / filename
    file.save(str(filepath))
    
    saved_size = filepath.stat().st_size
    print(f"[업로드] 파일 저장: {filename} ({saved_size:,} bytes)")
    
    # 파일 유형 감지 및 트랙 분기
    file_type = detect_file_type(str(filepath))
    
    if file_type == 'analysis':
        # 트랙 2: 분석표 처리 (레거시)
        return _handle_analysis_upload(filepath, filename)
    else:
        # 트랙 1: 기초자료 직접 처리 (기본)
        return _handle_raw_data_upload(filepath, filename)


def _extract_grdp_from_analysis_sheet(ws, year, quarter):
    """분석표의 GRDP 시트에서 데이터 추출"""
    import pandas as pd
    
    try:
        # 시트 데이터를 DataFrame으로 변환
        data = []
        for row in ws.iter_rows(values_only=True):
            data.append(row)
        
        if not data:
            return None
        
        df = pd.DataFrame(data)
        
        regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        region_groups = {
            '서울': '경인', '인천': '경인', '경기': '경인',
            '대전': '충청', '세종': '충청', '충북': '충청', '충남': '충청',
            '광주': '호남', '전북': '호남', '전남': '호남', '제주': '호남',
            '대구': '동북', '경북': '동북', '강원': '동북',
            '부산': '동남', '울산': '동남', '경남': '동남'
        }
        
        regional_data = []
        national_growth = 0.0
        top_region = {'name': '-', 'growth_rate': 0.0}
        
        # 지역별 성장률 추출
        for i, row in df.iterrows():
            for j, val in enumerate(row):
                if pd.notna(val) and str(val).strip() in regions:
                    region_name = str(val).strip()
                    growth_rate = 0.0
                    
                    # 다음 컬럼에서 성장률 찾기
                    for k in range(j+1, min(j+10, len(row))):
                        try:
                            growth_rate = float(row.iloc[k])
                            break
                        except:
                            continue
                    
                    if region_name == '전국':
                        national_growth = growth_rate
                    else:
                        regional_data.append({
                            'region': region_name,
                            'region_group': region_groups.get(region_name, ''),
                            'growth_rate': growth_rate,
                            'manufacturing': 0.0,
                            'construction': 0.0,
                            'service': 0.0,
                            'other': 0.0
                        })
                        
                        if growth_rate > top_region['growth_rate']:
                            top_region = {'name': region_name, 'growth_rate': growth_rate}
        
        if not regional_data and national_growth == 0.0:
            return None
        
        return {
            'report_info': {'year': year, 'quarter': quarter, 'page_number': ''},
            'national_summary': {
                'growth_rate': national_growth,
                'direction': '증가' if national_growth > 0 else '감소',
                'contributions': {'manufacturing': 0.0, 'construction': 0.0, 'service': 0.0, 'other': 0.0}
            },
            'top_region': {
                'name': top_region['name'],
                'growth_rate': top_region['growth_rate'],
                'contributions': {'manufacturing': 0.0, 'construction': 0.0, 'service': 0.0, 'other': 0.0}
            },
            'regional_data': regional_data,
            'source': 'analysis_sheet'
        }
        
    except Exception as e:
        print(f"[GRDP] 시트 데이터 추출 오류: {e}")
        return None


@api_bp.route('/upload-grdp', methods=['POST'])
def upload_grdp_file():
    """KOSIS GRDP 파일 업로드 및 파싱 + 분석표에 GRDP 시트 추가"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '파일이 없습니다.'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '파일이 선택되지 않았습니다.'}), 400
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'error': '엑셀 파일만 업로드 가능합니다.'}), 400
    
    filename = safe_filename(file.filename)
    if 'grdp' not in filename.lower() and 'GRDP' not in filename:
        filename = f"grdp_{filename}"
    
    filepath = UPLOAD_FOLDER / filename
    file.save(str(filepath))
    print(f"[GRDP 업로드] 파일 저장 완료: {filename}")
    
    year = session.get('year', 2025)
    quarter = session.get('quarter', 2)
    
    grdp_data = parse_kosis_grdp_file(str(filepath), year, quarter)
    
    if grdp_data:
        session['grdp_data'] = grdp_data
        grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
        with open(grdp_json_path, 'w', encoding='utf-8') as f:
            json.dump(grdp_data, f, ensure_ascii=False, indent=2)
        
        # 분석표에 GRDP 시트 추가 (분석표가 업로드된 경우)
        analysis_path = session.get('excel_path')
        grdp_sheet_added = False
        
        if analysis_path and Path(analysis_path).exists():
            try:
                grdp_sheet_added = _add_grdp_sheet_to_analysis(analysis_path, str(filepath), year, quarter)
                if grdp_sheet_added:
                    print(f"[GRDP] 분석표에 GRDP 시트 추가 완료: {analysis_path}")
            except Exception as e:
                print(f"[GRDP] 분석표에 GRDP 시트 추가 실패: {e}")
        
        return jsonify({
            'success': True,
            'message': 'GRDP 데이터가 성공적으로 업로드되었습니다.',
            'national_growth_rate': grdp_data.get('national_summary', {}).get('growth_rate', 0),
            'top_region': grdp_data.get('top_region', {}).get('name', '-'),
            'grdp_sheet_added': grdp_sheet_added
        })
    else:
        return jsonify({
            'success': False,
            'error': 'GRDP 데이터를 파싱할 수 없습니다. 올바른 KOSIS GRDP 파일인지 확인하세요.'
        }), 400


def _add_grdp_sheet_to_analysis(analysis_path: str, grdp_file_path: str, year: int, quarter: int) -> bool:
    """분석표에 GRDP 시트 추가 (KOSIS 파일에서 시트 복사)"""
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    
    try:
        # GRDP 파일에서 데이터 읽기
        grdp_df = pd.read_excel(grdp_file_path, header=None)
        
        # 분석표 열기
        wb = load_workbook(analysis_path)
        
        # 기존 GRDP 시트가 있으면 삭제
        grdp_sheet_names = ['I GRDP', 'GRDP', '분기 GRDP']
        for sheet_name in grdp_sheet_names:
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
        
        # 새 GRDP 시트 생성
        ws = wb.create_sheet('I GRDP')
        
        # 데이터 복사
        for r_idx, row in enumerate(dataframe_to_rows(grdp_df, index=False, header=False), 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
        
        # 저장
        wb.save(analysis_path)
        wb.close()
        
        print(f"[GRDP] 'I GRDP' 시트 추가 완료 ({len(grdp_df)}행)")
        return True
        
    except Exception as e:
        import traceback
        print(f"[GRDP] 시트 추가 오류: {e}")
        traceback.print_exc()
        return False


@api_bp.route('/use-default-grdp', methods=['POST'])
def use_default_grdp():
    """GRDP 파일이 없을 때 N/A로 진행 (기본값 사용 안 함)"""
    
    year = session.get('year', 2025)
    quarter = session.get('quarter', 2)
    
    # N/A 데이터 생성 (모든 값이 None)
    grdp_data = _generate_na_grdp_data(year, quarter)
    
    # 세션에 저장
    session['grdp_data'] = grdp_data
    
    # JSON 파일로도 저장
    grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
    with open(grdp_json_path, 'w', encoding='utf-8') as f:
        json.dump(grdp_data, f, ensure_ascii=False, indent=2)
    
    print(f"[GRDP] N/A로 진행 - {year}년 {quarter}분기 (별도 GRDP 파일 업로드 필요)")
    
    return jsonify({
        'success': True,
        'message': 'GRDP 데이터가 없습니다. 참고_GRDP 페이지는 N/A로 표시됩니다. KOSIS에서 데이터를 다운로드하여 업로드해주세요.',
        'is_na': True,
        'is_placeholder': True,
        'national_growth_rate': None,
        'needs_grdp_upload': True
    })


def _generate_na_grdp_data(year, quarter):
    """N/A GRDP 데이터 생성 (모든 값이 None)"""
    regions = ['전국', '서울', '인천', '경기', '대전', '세종', '충북', '충남',
               '광주', '전북', '전남', '제주', '대구', '경북', '강원', '부산', '울산', '경남']
    
    region_groups = {
        '서울': '경인', '인천': '경인', '경기': '경인',
        '대전': '충청', '세종': '충청', '충북': '충청', '충남': '충청',
        '광주': '호남', '전북': '호남', '전남': '호남', '제주': '호남',
        '대구': '동북', '경북': '동북', '강원': '동북',
        '부산': '동남', '울산': '동남', '경남': '동남'
    }
    
    regional_data = []
    for region in regions:
        regional_data.append({
            'region': region,
            'region_group': region_groups.get(region, ''),
            'growth_rate': None,  # N/A
            'manufacturing': None,
            'construction': None,
            'service': None,
            'other': None,
            'is_na': True,
            'placeholder': True
        })
    
    return {
        'report_info': {
            'year': year,
            'quarter': quarter,
        },
        'national_summary': {
            'growth_rate': None,  # N/A
            'contributions': {
                'manufacturing': None,
                'construction': None,
                'service': None,
                'other': None,
            },
            'is_na': True,
            'placeholder': True
        },
        'regional_data': regional_data,
        'is_na': True,
        'needs_grdp_upload': True
    }


def _add_placeholder_grdp_sheet(analysis_path: str, grdp_data: dict) -> bool:
    """분석표에 플레이스홀더 GRDP 시트 추가"""
    from openpyxl import load_workbook
    
    try:
        wb = load_workbook(analysis_path)
        
        # 기존 GRDP 시트가 있으면 삭제
        grdp_sheet_names = ['I GRDP', 'GRDP', '분기 GRDP']
        for sheet_name in grdp_sheet_names:
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
        
        # 새 GRDP 시트 생성
        ws = wb.create_sheet('I GRDP')
        
        # 헤더 행
        year = grdp_data.get('report_info', {}).get('year', 2025)
        quarter = grdp_data.get('report_info', {}).get('quarter', 2)
        
        ws['A1'] = '지역'
        ws['B1'] = f'{year}년 {quarter}분기 성장률(%)'
        ws['C1'] = '제조업'
        ws['D1'] = '건설업'
        ws['E1'] = '서비스업'
        ws['F1'] = '기타'
        
        # 전국 데이터
        ws['A2'] = '전국'
        ws['B2'] = grdp_data.get('national_summary', {}).get('growth_rate', 0.0)
        
        # 지역별 데이터
        regional_data = grdp_data.get('regional_data', [])
        for i, region in enumerate(regional_data, start=3):
            ws[f'A{i}'] = region.get('region', '')
            ws[f'B{i}'] = region.get('growth_rate', 0.0)
            ws[f'C{i}'] = region.get('manufacturing', 0.0)
            ws[f'D{i}'] = region.get('construction', 0.0)
            ws[f'E{i}'] = region.get('service', 0.0)
            ws[f'F{i}'] = region.get('other', 0.0)
        
        wb.save(analysis_path)
        wb.close()
        
        print(f"[GRDP] 플레이스홀더 'I GRDP' 시트 추가 완료")
        return True
        
    except Exception as e:
        import traceback
        print(f"[GRDP] 플레이스홀더 시트 추가 오류: {e}")
        traceback.print_exc()
        return False


# ============================================================================
# 레거시 API: 분석표 관련 (트랙 2 전용)
# 주의: 이 API들은 레거시 워크플로우 지원을 위해 유지됩니다.
# 새로운 기능 개발 시 트랙 1 (raw_direct)을 사용하세요.
# ============================================================================

@api_bp.route('/download-analysis', methods=['GET'])
def download_analysis():
    """[레거시] 분석표 다운로드 (다운로드 시점에 생성 + 수식 계산)
    
    주의: 이 API는 레거시 워크플로우 지원을 위해 유지됩니다.
    트랙 1 (raw_direct)에서는 사용되지 않습니다.
    """
    import time
    import zipfile
    
    raw_excel_path = session.get('raw_excel_path')
    
    if not raw_excel_path or not Path(raw_excel_path).exists():
        return jsonify({'success': False, 'error': '기초자료 파일을 찾을 수 없습니다. 먼저 기초자료를 업로드해주세요.'}), 404
    
    try:
        converter = DataConverter(str(raw_excel_path))
        analysis_output = str(UPLOAD_FOLDER / f"분석표_{converter.year}년_{converter.quarter}분기_자동생성.xlsx")
        
        # 이미 유효한 분석표가 있는지 확인 (세션에서 생성된 파일)
        download_path = session.get('download_analysis_path')
        raw_file_mtime = session.get('raw_file_mtime')  # 원본 파일 수정 시간
        need_regenerate = True
        
        if download_path and Path(download_path).exists():
            # 원본 파일이 변경되었는지 확인
            current_raw_mtime = Path(raw_excel_path).stat().st_mtime if Path(raw_excel_path).exists() else None
            file_changed = (raw_file_mtime is None or current_raw_mtime is None or 
                          abs(current_raw_mtime - raw_file_mtime) > 1.0)  # 1초 이상 차이
            
            if file_changed:
                print(f"[다운로드] 원본 파일이 변경되었습니다, 재생성 필요")
                need_regenerate = True
            else:
                # 기존 파일 유효성 검사
                try:
                    with zipfile.ZipFile(download_path, 'r') as zf:
                        # zip 파일이 유효한지 테스트
                        if zf.testzip() is None:
                            need_regenerate = False
                            analysis_output = download_path
                            print(f"[다운로드] 기존 분석표 재사용: {download_path}")
                except (zipfile.BadZipFile, EOFError):
                    print(f"[다운로드] 기존 파일 손상됨, 재생성 필요")
                    need_regenerate = True
        
        if need_regenerate:
            # 분석표 생성
            analysis_path = converter.convert_all(analysis_output, weight_settings=None)
            
            # 파일 저장 완료 대기 (파일 시스템 동기화)
            time.sleep(0.3)
            
            # 분석 시트 수식 계산 (집계 시트 값을 분석 시트로 복사)
            _calculate_analysis_sheets(analysis_path)
            
            # 세션에 저장 (원본 파일 수정 시간 포함)
            session['download_analysis_path'] = analysis_path
            try:
                session['raw_file_mtime'] = Path(raw_excel_path).stat().st_mtime
            except OSError:
                pass  # 파일 시간 확인 실패는 무시
        else:
            analysis_path = analysis_output
        
        filename = Path(analysis_path).name
        
        return send_file_with_korean_filename(
            analysis_path,
            filename,
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': f'분석표 생성 실패: {str(e)}'}), 500


@api_bp.route('/generate-analysis-with-weights', methods=['POST'])
def generate_analysis_with_weights():
    """[레거시] 가중치 설정을 포함하여 분석표 생성 + 다운로드
    
    주의: 이 API는 레거시 워크플로우 지원을 위해 유지됩니다.
    트랙 1 (raw_direct)에서는 사용되지 않습니다.
    """
    import time
    
    data = request.get_json()
    weight_settings = data.get('weight_settings', {})  # {mining: {mode, values}, service: {mode, values}}
    
    raw_excel_path = session.get('raw_excel_path')
    if not raw_excel_path or not Path(raw_excel_path).exists():
        return jsonify({'success': False, 'error': '기초자료 파일을 찾을 수 없습니다.'}), 404
    
    try:
        converter = DataConverter(str(raw_excel_path))
        
        # 분석표 생성 (가중치 설정 포함)
        analysis_output = str(UPLOAD_FOLDER / f"분석표_{converter.year}년_{converter.quarter}분기_자동생성.xlsx")
        analysis_path = converter.convert_all(analysis_output, weight_settings=weight_settings)
        
        # 파일 저장 완료 대기 (파일 시스템 동기화)
        time.sleep(0.3)
        
        # 분석 시트 수식 계산 (집계 시트 값을 분석 시트로 복사)
        _calculate_analysis_sheets(analysis_path)
        
        # 파일 무결성 확인
        import zipfile
        try:
            with zipfile.ZipFile(analysis_path, 'r') as zf:
                if zf.testzip() is not None:
                    raise Exception("생성된 파일이 손상되었습니다.")
        except zipfile.BadZipFile:
            raise Exception("생성된 파일이 손상되었습니다. 다시 시도해주세요.")
        
        session['download_analysis_path'] = analysis_path
        
        return jsonify({
            'success': True,
            'filename': Path(analysis_path).name,
            'message': '분석표가 성공적으로 생성되었습니다.'
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': f'분석표 생성 실패: {str(e)}'}), 500


@api_bp.route('/report-order', methods=['GET'])
def get_report_order():
    """현재 보도자료 순서 반환"""
    return jsonify({'reports': REPORT_ORDER, 'regional_reports': REGIONAL_REPORTS})


@api_bp.route('/report-order', methods=['POST'])
def update_report_order():
    """보도자료 순서 업데이트"""
    from config import reports as reports_module
    data = request.get_json()
    new_order = data.get('order', [])
    
    if new_order:
        order_map = {r['id']: idx for idx, r in enumerate(new_order)}
        reports_module.REPORT_ORDER = sorted(reports_module.REPORT_ORDER, key=lambda x: order_map.get(x['id'], 999))
    
    return jsonify({'success': True, 'reports': reports_module.REPORT_ORDER})


@api_bp.route('/export-final', methods=['POST'])
def export_final_document():
    """모든 보도자료를 HTML 문서로 합치기 (standalone 옵션 지원)"""
    try:
        data = request.get_json()
        pages = data.get('pages', [])
        year = data.get('year', session.get('year', 2025))
        quarter = data.get('quarter', session.get('quarter', 2))
        standalone = data.get('standalone', False)  # 완전한 standalone HTML 여부
        
        if not pages:
            return jsonify({'success': False, 'error': '페이지 데이터가 없습니다.'})
        
        # 모든 페이지의 스타일 수집
        all_styles = set()
        
        # standalone 모드: 외부 의존성 없는 완전한 HTML
        # 일반 모드: 외부 폰트 사용
        font_style = ''
        if not standalone:
            font_style = "@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap');"
        
        final_html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{year}년 {quarter}/4분기 지역경제동향</title>
    <style>
        {font_style}
        
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        html, body {{
            width: 210mm;
            background: white;
        }}
        
        body {{
            font-family: 'Noto Sans KR', '맑은 고딕', 'Malgun Gothic', '나눔고딕', 'NanumGothic', sans-serif;
        }}
        
        /* PDF 출력용 페이지 스타일 */
        .pdf-page {{
            width: 210mm;
            min-height: 297mm;
            max-height: 297mm;
            padding: 12mm 15mm 15mm 15mm;
            margin: 0 auto 5mm auto;
            background: white;
            position: relative;
            overflow: hidden;
            page-break-after: always;
            page-break-inside: avoid;
        }}
        
        .pdf-page:last-child {{
            page-break-after: auto;
            margin-bottom: 0;
        }}
        
        .pdf-page-content {{
            width: 100%;
            height: calc(297mm - 32mm);
            overflow: hidden;
        }}
        
        .pdf-page-content > * {{
            max-width: 100%;
        }}
        
        /* 페이지 번호 */
        .pdf-page-number {{
            position: absolute;
            bottom: 8mm;
            left: 0;
            right: 0;
            text-align: center;
            font-size: 9pt;
            color: #333;
        }}
        
        /* 변환된 차트 이미지 스타일 */
        img[data-converted-from="canvas"] {{
            max-width: 100%;
            height: auto;
            display: block;
        }}
        
        /* 화면 미리보기용 */
        @media screen {{
            body {{
                background: #f0f0f0;
                padding: 20px;
            }}
            
            .pdf-page {{
                box-shadow: 0 2px 10px rgba(0,0,0,0.15);
                border: 1px solid #ddd;
            }}
        }}
        
        /* 인쇄/PDF 저장용 */
        @media print {{
            html, body {{
                width: 210mm;
                background: white !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
            
            body {{
                padding: 0;
                margin: 0;
            }}
            
            .pdf-page {{
                width: 210mm;
                height: 297mm;
                min-height: 297mm;
                max-height: 297mm;
                padding: 12mm 15mm 15mm 15mm;
                margin: 0;
                box-shadow: none;
                border: none;
                page-break-after: always;
                page-break-inside: avoid;
            }}
            
            .pdf-page:last-child {{
                page-break-after: auto;
            }}
            
            /* 이미지 색상 유지 */
            img {{
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
        }}
        
        @page {{
            size: A4 portrait;
            margin: 0;
        }}
        
        /* 표 스타일 공통 */
        table {{
            border-collapse: collapse;
            width: 100%;
        }}
        
        th, td {{
            border: 1px solid #333;
            padding: 4px 6px;
            font-size: 9pt;
            text-align: center;
        }}
        
        th {{
            background: #f5f5f5;
            font-weight: 600;
        }}
        
        /* 차트 이미지 크기 조정 */
        .chart-container, .chart-wrapper {{
            max-width: 100%;
            overflow: hidden;
        }}
        
        .chart-container img, .chart-wrapper img {{
            max-width: 100%;
            height: auto;
        }}
        
        /* A4 최적화 - 이미지 리사이징 */
        img {{
            max-width: 100%;
            height: auto;
            object-fit: contain;
        }}
        
        /* A4 최적화 - 표 리사이징 */
        table {{
            width: 100%;
            table-layout: fixed;
            font-size: 8pt;
            word-wrap: break-word;
        }}
        
        th, td {{
            overflow: hidden;
            text-overflow: ellipsis;
            word-break: keep-all;
        }}
        
        /* A4 최적화 - 콘텐츠 영역 */
        .pdf-page-content {{
            width: 180mm;
            max-width: 180mm;
        }}
        
        .pdf-page-content img {{
            max-width: 180mm;
            max-height: 240mm;
        }}
        
        .pdf-page-content table {{
            max-width: 180mm;
        }}
        
        /* A4 최적화 - SVG 리사이징 */
        svg {{
            max-width: 100%;
            height: auto;
        }}
        
        /* 큰 요소 overflow 방지 */
        .pdf-page-content > * {{
            max-width: 100%;
            overflow: hidden;
        }}
        
        /* 그래프/차트 컨테이너 */
        .chart-area, .graph-container {{
            max-width: 100%;
            overflow: hidden;
        }}
    </style>
'''
        
        # 각 페이지에서 스타일 추출하여 추가
        for idx, page in enumerate(pages):
            page_html = page.get('html', '')
            if '<style' in page_html:
                style_matches = re.findall(r'<style[^>]*>(.*?)</style>', page_html, re.DOTALL)
                for style in style_matches:
                    # 중복 방지를 위해 hash 사용
                    style_hash = hash(style.strip())
                    if style_hash not in all_styles:
                        all_styles.add(style_hash)
                        # standalone 모드에서는 외부 폰트 import 제거
                        if standalone:
                            style = re.sub(r'@import\s+url\([^)]+\);?', '', style)
                        final_html += f'    <style>/* Page {idx+1} styles */\n{style}\n    </style>\n'
        
        final_html += '''</head>
<body>
'''
        
        for idx, page in enumerate(pages, 1):
            page_html = page.get('html', '')
            page_title = page.get('title', f'페이지 {idx}')
            
            # body 내용 추출
            body_content = page_html
            if '<body' in page_html.lower():
                body_match = re.search(r'<body[^>]*>(.*?)</body>', page_html, re.DOTALL | re.IGNORECASE)
                if body_match:
                    body_content = body_match.group(1)
            
            # 내용에서 style, script 태그 제거 (이미 head에 추가됨)
            body_content = re.sub(r'<style[^>]*>.*?</style>', '', body_content, flags=re.DOTALL)
            
            # standalone 모드에서는 script 태그도 제거 (Chart.js 등 불필요)
            if standalone:
                body_content = re.sub(r'<script[^>]*>.*?</script>', '', body_content, flags=re.DOTALL)
            
            # 페이지 래퍼 추가
            final_html += f'''
    <!-- Page {idx}: {page_title} -->
    <div class="pdf-page" data-page="{idx}" data-title="{page_title}">
        <div class="pdf-page-content">
{body_content}
        </div>
        <div class="pdf-page-number">- {idx} -</div>
    </div>
'''
        
        # standalone 모드에서는 최소한의 스크립트만 추가
        if standalone:
            final_html += '''
    <script>
        // 인쇄 전 준비
        window.onbeforeprint = function() {
            document.body.style.background = 'white';
        };
        console.log('이 HTML 파일은 오프라인에서도 사용 가능합니다.');
        console.log('PDF 저장: Ctrl+P (또는 Cmd+P) → "PDF로 저장" 선택');
    </script>
'''
        else:
            final_html += '''
    <script>
        // 인쇄 전 준비
        window.onbeforeprint = function() {
            document.body.style.background = 'white';
        };
        
        // Ctrl+P로 PDF 저장 안내
        console.log('PDF 저장: Ctrl+P (또는 Cmd+P) → "PDF로 저장" 선택');
    </script>
'''
        
        final_html += '''</body>
</html>
'''
        
        # 파일명 설정
        if standalone:
            output_filename = f'지역경제동향_{year}년_{quarter}분기.html'
        else:
            output_filename = f'지역경제동향_{year}년_{quarter}분기_PDF용.html'
        
        output_path = UPLOAD_FOLDER / output_filename
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        return jsonify({
            'success': True,
            'html': final_html,
            'filename': output_filename,
            'download_url': f'/uploads/{output_filename}',
            'total_pages': len(pages),
            'standalone': standalone
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


@api_bp.route('/export-xlsx', methods=['POST'])
def export_xlsx_document():
    """모든 보도자료를 XLSX 파일로 내보내기"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
        from openpyxl.utils.dataframe import dataframe_to_rows
        from openpyxl.chart import BarChart, LineChart, Reference
        from openpyxl.drawing.image import Image as XLImage
        from bs4 import BeautifulSoup
        import io
        import base64
        from PIL import Image as PILImage
        
        data = request.get_json()
        pages = data.get('pages', [])
        year = data.get('year', session.get('year', 2025))
        quarter = data.get('quarter', session.get('quarter', 2))
        
        if not pages:
            return jsonify({'success': False, 'error': '페이지 데이터가 없습니다.'})
        
        # 워크북 생성
        wb = Workbook()
        wb.remove(wb.active)  # 기본 시트 제거
        
        # 스타일 정의
        header_font = Font(bold=True, size=11)
        header_fill = PatternFill(start_color='E6E6E6', end_color='E6E6E6', fill_type='solid')
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        image_counter = 0
        
        for idx, page in enumerate(pages, 1):
            page_html = page.get('html', '')
            page_title = page.get('title', f'페이지{idx}')
            category = page.get('category', 'unknown')
            
            # 시트 이름 (최대 31자, 특수문자 제거)
            sheet_name = re.sub(r'[\\/*?:\[\]]', '', page_title)[:31]
            if sheet_name in [ws.title for ws in wb.worksheets]:
                sheet_name = f"{sheet_name[:28]}_{idx}"
            
            ws = wb.create_sheet(title=sheet_name)
            
            # HTML 파싱
            soup = BeautifulSoup(page_html, 'html.parser')
            
            current_row = 1
            
            # 제목 추가
            ws.cell(row=current_row, column=1, value=page_title)
            ws.cell(row=current_row, column=1).font = Font(bold=True, size=14)
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
            current_row += 2
            
            # 표 추출 및 변환
            tables = soup.find_all('table')
            for table in tables:
                # 표 제목 찾기 (이전 요소에서)
                table_title = None
                prev = table.find_previous(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'div'])
                if prev and prev.get_text(strip=True):
                    table_title = prev.get_text(strip=True)[:100]
                
                if table_title:
                    ws.cell(row=current_row, column=1, value=table_title)
                    ws.cell(row=current_row, column=1).font = Font(bold=True, size=11)
                    current_row += 1
                
                # 표 데이터 추출
                rows = table.find_all('tr')
                for row_idx, tr in enumerate(rows):
                    cells = tr.find_all(['th', 'td'])
                    for col_idx, cell in enumerate(cells, 1):
                        cell_text = cell.get_text(strip=True)
                        
                        # 숫자 변환 시도
                        try:
                            if '.' in cell_text:
                                cell_value = float(cell_text.replace(',', '').replace('%', ''))
                            elif cell_text.replace(',', '').replace('-', '').isdigit():
                                cell_value = int(cell_text.replace(',', ''))
                            else:
                                cell_value = cell_text
                        except:
                            cell_value = cell_text
                        
                        excel_cell = ws.cell(row=current_row, column=col_idx, value=cell_value)
                        excel_cell.border = thin_border
                        excel_cell.alignment = center_align
                        
                        # 헤더 스타일
                        if cell.name == 'th' or row_idx == 0:
                            excel_cell.font = header_font
                            excel_cell.fill = header_fill
                        
                        # colspan 처리
                        colspan = int(cell.get('colspan', 1))
                        if colspan > 1:
                            ws.merge_cells(
                                start_row=current_row, start_column=col_idx,
                                end_row=current_row, end_column=col_idx + colspan - 1
                            )
                    
                    current_row += 1
                
                current_row += 2  # 표 간 간격
            
            # 인포그래픽/이미지 처리 (base64 이미지)
            images = soup.find_all('img')
            for img in images:
                src = img.get('src', '')
                if src.startswith('data:image'):
                    try:
                        # base64 이미지 추출
                        match = re.match(r'data:image/([^;]+);base64,(.+)', src)
                        if match:
                            img_format = match.group(1)
                            img_data = base64.b64decode(match.group(2))
                            
                            # 이미지를 임시 파일로 저장
                            img_buffer = io.BytesIO(img_data)
                            pil_img = PILImage.open(img_buffer)
                            
                            # PNG로 변환
                            png_buffer = io.BytesIO()
                            pil_img.save(png_buffer, format='PNG')
                            png_buffer.seek(0)
                            
                            # 엑셀에 이미지 추가
                            xl_img = XLImage(png_buffer)
                            
                            # 이미지 크기 조정 (최대 너비 500px)
                            max_width = 500
                            if xl_img.width > max_width:
                                ratio = max_width / xl_img.width
                                xl_img.width = max_width
                                xl_img.height = int(xl_img.height * ratio)
                            
                            ws.add_image(xl_img, f'A{current_row}')
                            image_counter += 1
                            
                            # 이미지 높이만큼 행 건너뛰기
                            rows_needed = max(1, xl_img.height // 20)
                            current_row += rows_needed + 2
                    except Exception as e:
                        print(f"이미지 처리 오류: {e}")
                        continue
            
            # 열 너비 자동 조정
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column].width = adjusted_width
        
        # 파일 저장
        output_filename = f'지역경제동향_{year}년_{quarter}분기.xlsx'
        output_path = UPLOAD_FOLDER / output_filename
        
        wb.save(output_path)
        
        # 파일을 바이트로 읽어서 base64로 인코딩
        with open(output_path, 'rb') as f:
            xlsx_data = base64.b64encode(f.read()).decode('utf-8')
        
        return jsonify({
            'success': True,
            'filename': output_filename,
            'download_url': f'/uploads/{output_filename}',
            'total_pages': len(pages),
            'xlsx_data': xlsx_data,
            'image_count': image_counter
        })
        
    except ImportError as e:
        return jsonify({
            'success': False, 
            'error': f'필요한 라이브러리가 설치되지 않았습니다: {str(e)}. pip install openpyxl pillow beautifulsoup4'
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


@api_bp.route('/cleanup-uploads', methods=['POST'])
def cleanup_uploads():
    """업로드 폴더 정리 API (작업 완료 후 호출)"""
    try:
        deleted_count = cleanup_upload_folder(keep_current_files=True, cleanup_excel_only=True)
        return jsonify({
            'success': True,
            'deleted_count': deleted_count,
            'message': f'{deleted_count}개 파일이 정리되었습니다.'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })


def validate_weights_in_raw_data(filepath: str) -> dict:
    """기초자료 수집표에서 가중치 누락 여부 검증
    
    개선된 버전: 시트 내 다중 테이블 구조 지원
    - 빈 행으로 구분된 여러 테이블을 개별적으로 인식
    - 헤더에 "가중치" 열이 있는 테이블만 검증
    - "보조 지표" 등 가중치가 원래 없는 테이블은 제외
    
    Returns:
        {
            'has_missing_weights': bool,
            'missing_details': [
                {'sheet': '시트명', 'missing_count': int, 'total_count': int, 'missing_industries': [...]}
            ],
            'total_missing': int,
            'affected_reports': ['광공업생산', '서비스업생산']
        }
    """
    import pandas as pd
    
    result = {
        'has_missing_weights': False,
        'missing_details': [],
        'total_missing': 0,
        'affected_reports': []
    }
    
    def find_tables_in_sheet(df):
        """시트 내 테이블들의 범위와 헤더 정보를 찾음"""
        tables = []
        current_table = None
        
        for i in range(len(df)):
            row = df.iloc[i]
            
            # 빈 행 확인
            is_empty_row = all(pd.isna(row.iloc[j]) for j in range(min(6, len(row))))
            
            if is_empty_row:
                # 현재 테이블 종료
                if current_table is not None:
                    current_table['end_row'] = i - 1
                    tables.append(current_table)
                    current_table = None
                continue
            
            # 헤더 행 감지 (지역코드/지역이름 패턴)
            first_cell = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            
            # 헤더 패턴 감지: "지역" 포함, "코드" 포함
            is_header = False
            if '지역' in first_cell and ('코드' in first_cell or '이름' in first_cell):
                is_header = True
            elif first_cell in ['지역코드', '지역\n코드']:
                is_header = True
            
            if is_header:
                # 이전 테이블 종료
                if current_table is not None:
                    current_table['end_row'] = i - 1
                    tables.append(current_table)
                
                # 헤더에서 가중치 열 위치 찾기
                weight_col = None
                has_weight_header = False
                for j in range(min(10, len(row))):
                    cell_val = str(row.iloc[j]).strip() if pd.notna(row.iloc[j]) else ''
                    if '가중치' in cell_val:
                        weight_col = j
                        has_weight_header = True
                        break
                
                # 보조 지표 테이블 확인 (가중치 열 없음)
                is_auxiliary = False
                for j in range(min(10, len(row))):
                    cell_val = str(row.iloc[j]).strip() if pd.notna(row.iloc[j]) else ''
                    if '보조' in cell_val and '지표' in cell_val:
                        is_auxiliary = True
                        break
                
                # 새 테이블 시작
                current_table = {
                    'header_row': i,
                    'start_row': i + 1,
                    'end_row': len(df) - 1,  # 기본값
                    'weight_col': weight_col if has_weight_header else None,
                    'has_weight_header': has_weight_header,
                    'is_auxiliary': is_auxiliary
                }
        
        # 마지막 테이블 추가
        if current_table is not None:
            tables.append(current_table)
        
        return tables
    
    def get_column_indices_from_header(df, header_row):
        """헤더 행에서 열 인덱스 동적으로 찾기"""
        header = df.iloc[header_row]
        indices = {
            'region_col': None,
            'level_col': None,
            'weight_col': None,
            'code_col': None,
            'name_col': None
        }
        
        for j in range(min(15, len(header))):
            cell = str(header.iloc[j]).strip() if pd.notna(header.iloc[j]) else ''
            cell_lower = cell.replace('\n', '').replace(' ', '').lower()
            
            if '지역이름' in cell_lower or cell_lower == '지역\n이름' or '지역명' in cell:
                indices['region_col'] = j
            elif '분류단계' in cell_lower or '분류\n단계' in cell:
                indices['level_col'] = j
            elif '가중치' in cell_lower:
                indices['weight_col'] = j
            elif '산업코드' in cell_lower or '산업\n코드' in cell:
                indices['code_col'] = j
            elif '산업이름' in cell_lower or '산업 이름' in cell or ('산업' in cell and '이름' in cell):
                indices['name_col'] = j
        
        return indices
    
    try:
        xl = pd.ExcelFile(filepath)
        
        # 가중치 검증이 필요한 시트 목록
        weight_check_sheets = {
            '광공업생산': '광공업생산',
            '서비스업생산': '서비스업생산'
        }
        
        for sheet_name, affected_report in weight_check_sheets.items():
            if sheet_name not in xl.sheet_names:
                continue
            
            df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
            
            # 시트 내 테이블들 찾기
            tables = find_tables_in_sheet(df)
            
            print(f"[가중치 검증] {sheet_name}: {len(tables)}개 테이블 발견")
            for idx, table in enumerate(tables):
                print(f"  테이블 {idx+1}: 행 {table['header_row']+1}-{table['end_row']+1}, 가중치열={table['weight_col']}, 보조={table['is_auxiliary']}")
            
            missing_industries = []
            total_industries = 0
            
            for table in tables:
                # 보조 지표 테이블이거나 가중치 열이 없으면 건너뛰기
                if table['is_auxiliary'] or not table['has_weight_header']:
                    print(f"  → 테이블 건너뜀 (보조지표={table['is_auxiliary']}, 가중치헤더없음={not table['has_weight_header']})")
                    continue
                
                # 헤더에서 열 인덱스 동적으로 가져오기
                col_indices = get_column_indices_from_header(df, table['header_row'])
                
                weight_col = col_indices['weight_col'] or table['weight_col'] or 3
                region_col = col_indices['region_col'] or 1
                name_col = col_indices['name_col'] or 5
                level_col = col_indices['level_col'] or 2
                
                # 테이블 데이터 범위에서 검증
                for i in range(table['start_row'], table['end_row'] + 1):
                    row = df.iloc[i]
                    
                    # 지역 확인
                    region = str(row.iloc[region_col]).strip() if region_col < len(row) and pd.notna(row.iloc[region_col]) else ''
                    if not region or region in ['nan', 'NaN', '']:
                        continue
                    
                    # 산업 이름 확인
                    industry_name = str(row.iloc[name_col]).strip() if name_col < len(row) and pd.notna(row.iloc[name_col]) else ''
                    if not industry_name or industry_name in ['nan', 'NaN', '']:
                        continue
                    
                    # 분류단계
                    level = row.iloc[level_col] if level_col < len(row) else None
                    
                    total_industries += 1
                    
                    # 가중치 확인 (전국 데이터만)
                    if region == '전국':
                        weight = row.iloc[weight_col] if weight_col < len(row) else None
                        if pd.isna(weight):
                            missing_industries.append({
                                'row': i + 1,
                                'region': region,
                                'industry': industry_name,
                                'level': str(level) if pd.notna(level) else ''
                            })
            
            if missing_industries:
                result['missing_details'].append({
                    'sheet': sheet_name,
                    'missing_count': len(missing_industries),
                    'total_count': total_industries,
                    'missing_industries': missing_industries[:10]
                })
                result['affected_reports'].append(affected_report)
                result['total_missing'] += len(missing_industries)
        
        result['has_missing_weights'] = result['total_missing'] > 0
        
    except Exception as e:
        print(f"[가중치 검증] 오류: {e}")
        import traceback
        traceback.print_exc()
    
    return result


@api_bp.route('/export-hwp-import', methods=['POST'])
def export_hwp_import():
    """한글(HWP)에서 불러오기로 열 수 있는 HTML 문서 생성
    
    한글 프로그램에서 [파일 → 불러오기] 또는 [Ctrl+O]로 열 수 있는 HTML 형식입니다.
    - 완전한 인라인 스타일 적용 (한글 호환성 극대화)
    - 표 테두리, 글꼴, 여백 등이 한글에서 정확하게 렌더링됩니다.
    """
    try:
        data = request.get_json()
        pages = data.get('pages', [])
        year = data.get('year', session.get('year', 2025))
        quarter = data.get('quarter', session.get('quarter', 2))
        
        if not pages:
            return jsonify({'success': False, 'error': '페이지 데이터가 없습니다.'})
        
        # 한글에서 완벽하게 인식되는 HTML 구조
        final_html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>{year}년 {quarter}/4분기 지역경제동향</title>
</head>
<body style="font-family: '맑은 고딕', Malgun Gothic, sans-serif; font-size: 10pt; line-height: 160%; margin: 20mm 15mm; color: #000000;">

<h1 style="font-size: 18pt; font-weight: bold; text-align: center; margin-bottom: 20px; border-bottom: 2px solid #000; padding-bottom: 10px;">
{year}년 {quarter}/4분기 지역경제동향
</h1>

'''
        
        # 각 페이지 처리
        for idx, page in enumerate(pages, 1):
            page_html = page.get('html', '')
            page_title = page.get('title', f'페이지 {idx}')
            category = page.get('category', '')
            
            # body 내용 추출
            body_content = page_html
            if '<body' in page_html.lower():
                body_match = re.search(r'<body[^>]*>(.*?)</body>', page_html, re.DOTALL | re.IGNORECASE)
                if body_match:
                    body_content = body_match.group(1)
            
            # 불필요한 태그 제거
            body_content = re.sub(r'<style[^>]*>.*?</style>', '', body_content, flags=re.DOTALL)
            body_content = re.sub(r'<script[^>]*>.*?</script>', '', body_content, flags=re.DOTALL)
            body_content = re.sub(r'<link[^>]*/?>', '', body_content)
            body_content = re.sub(r'<meta[^>]*/?>', '', body_content)
            body_content = re.sub(r'<!DOCTYPE[^>]*>', '', body_content)
            body_content = re.sub(r'<html[^>]*>', '', body_content)
            body_content = re.sub(r'</html>', '', body_content)
            body_content = re.sub(r'<head[^>]*>.*?</head>', '', body_content, flags=re.DOTALL)
            
            # canvas, svg 태그 제거 (한글에서 지원 안됨)
            body_content = re.sub(r'<canvas[^>]*>.*?</canvas>', '', body_content, flags=re.DOTALL)
            body_content = re.sub(r'<canvas[^>]*/?>',  '', body_content)
            body_content = re.sub(r'<svg[^>]*>.*?</svg>', '', body_content, flags=re.DOTALL)
            
            # 한글 완벽 호환 인라인 스타일 적용
            body_content = _apply_hwp_inline_styles(body_content)
            
            # 카테고리 한글명
            category_names = {
                'summary': '요약',
                'sectoral': '부문별',
                'regional': '시도별',
                'statistics': '통계표'
            }
            category_name = category_names.get(category, '')
            
            # 페이지 구분
            page_break = 'page-break-after: always;' if idx < len(pages) else ''
            final_html += f'''
<!-- 페이지 {idx}: {page_title} -->
<div style="margin-bottom: 30px; {page_break}">
    <h2 style="font-size: 14pt; font-weight: bold; background-color: #e8e8e8; padding: 8px 12px; margin: 20px 0 15px 0; border-left: 4px solid #0066cc;">
        [{category_name}] {page_title}
    </h2>
    <div style="font-size: 10pt; line-height: 160%;">
{body_content}
    </div>
</div>
'''
        
        # 문서 끝
        final_html += '''
<p style="text-align: center; font-size: 9pt; color: #666; margin-top: 30px; border-top: 1px solid #ccc; padding-top: 15px;">
    ※ 이 문서는 지역경제동향 보도자료 시스템에서 생성되었습니다.
</p>

</body>
</html>
'''
        
        # 파일 저장
        output_filename = f'지역경제동향_{year}년_{quarter}분기_한글불러오기용.html'
        output_path = UPLOAD_FOLDER / output_filename
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        # exports 폴더에도 저장
        exports_path = EXPORT_FOLDER / output_filename
        EXPORT_FOLDER.mkdir(exist_ok=True)
        with open(exports_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        return jsonify({
            'success': True,
            'filename': output_filename,
            'download_url': f'/uploads/{output_filename}',
            'total_pages': len(pages),
            'message': f'한글 불러오기용 HTML이 생성되었습니다.\n\n사용법:\n1. 다운로드된 파일을 저장합니다\n2. 한글(HWP)에서 [파일 → 불러오기] 또는 Ctrl+O\n3. 파일 형식을 "HTML 문서"로 선택\n4. 저장된 파일을 엽니다'
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


def _apply_hwp_inline_styles(html_content):
    """한글 프로그램에서 완벽하게 렌더링되는 인라인 스타일 적용
    
    한글(HWP)의 HTML 불러오기 기능은 CSS 클래스를 무시하고
    인라인 스타일만 인식하므로, 모든 스타일을 인라인으로 변환합니다.
    
    주요 개선사항:
    - 모든 CSS 클래스를 인라인 스타일로 변환
    - 한글에서 지원하는 기본 글꼴 사용 (맑은 고딕)
    - 테두리, 배경색, 여백 등 명확하게 지정
    """
    
    # CSS 클래스를 인라인 스타일로 변환하는 매핑
    class_to_style = {
        'bold': 'font-weight: bold;',
        'title': 'text-align: center; font-size: 16pt; font-weight: bold; margin-bottom: 25px; border-bottom: 2px solid #1976D2; padding-bottom: 8px;',
        'blue': 'color: #1976D2;',
        'section-box': 'border: 1px solid #000000; padding: 12px 15px; margin-bottom: 15px;',
        'item': 'margin-bottom: 8px; padding-left: 20px; text-indent: -20px; line-height: 1.7;',
        'sub-item': 'margin-left: 20px; padding-left: 20px; text-indent: -20px; line-height: 1.7;',
        'gray-bg': 'background-color: #e8e8e8;',
        'light-gray': 'background-color: #f5f5f5;',
        'table-title': 'text-align: center; font-weight: bold; margin: 20px 0 10px 0;',
        'right-note': 'text-align: right; font-size: 9pt; margin-bottom: 3px;',
        'page-num': 'text-align: center; margin-top: 20px;',
        'center': 'text-align: center;',
        'right': 'text-align: right;',
        'no-border': 'border: none;',
        'increase': 'font-weight: bold; color: #c00000;',
        'decrease': 'font-weight: bold; color: #0000c0;',
        'highlight': 'font-weight: bold;',
        'placeholder': 'background-color: #fff3cd; border: 1px dashed #ffc107; padding: 2px 6px;',
        'section-title': 'font-weight: bold; font-size: 11pt; margin: 12px 0 8px 0;',
        'region-title': 'font-weight: bold; font-size: 13pt; margin-bottom: 12px;',
        'main-header': 'font-weight: bold; font-size: 14pt; text-align: center; padding: 6px 40px; background-color: #e0e0e0; margin-bottom: 18px;',
        'summary-table': 'border-collapse: collapse; width: 100%; font-size: 8pt;',
        'period-col': 'font-weight: bold; background-color: #f0f0f0;',
    }
    
    # CSS 클래스를 인라인 스타일로 변환
    def convert_class_to_inline(match):
        tag_name = match.group(1)
        attrs = match.group(2) if match.group(2) else ''
        
        # class 속성 추출
        class_match = re.search(r'class="([^"]*)"', attrs)
        if not class_match:
            return match.group(0)
        
        classes = class_match.group(1).split()
        styles = []
        
        for cls in classes:
            if cls in class_to_style:
                styles.append(class_to_style[cls])
        
        # 기존 style 속성 추출
        style_match = re.search(r'style="([^"]*)"', attrs)
        if style_match:
            styles.append(style_match.group(1))
        
        # class와 기존 style 제거
        attrs = re.sub(r'class="[^"]*"', '', attrs)
        attrs = re.sub(r'style="[^"]*"', '', attrs)
        
        if styles:
            style_str = ' '.join(styles)
            return f'<{tag_name}{attrs} style="{style_str}">'
        return f'<{tag_name}{attrs}>'
    
    # 모든 태그의 class를 인라인으로 변환
    html_content = re.sub(r'<(\w+)([^>]*)>', convert_class_to_inline, html_content)
    
    # 기존 style 속성 제거 후 새로운 인라인 스타일 적용
    
    # table 태그 처리
    def replace_table(match):
        attrs = match.group(1) if match.group(1) else ''
        # 기존 style이 있으면 유지하면서 추가
        if 'style=' not in attrs:
            return f'<table{attrs} style="border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 9pt; border: 2px solid #000000;">'
        return match.group(0)
    
    html_content = re.sub(r'<table([^>]*)>', replace_table, html_content)
    
    # th 태그 처리
    def replace_th(match):
        attrs = match.group(1) if match.group(1) else ''
        attrs = re.sub(r'style="[^"]*"', '', attrs)
        attrs = re.sub(r"style='[^']*'", '', attrs)
        return f'<th{attrs} style="border: 1px solid #000000; padding: 6px 8px; text-align: center; vertical-align: middle; background-color: #d9d9d9; font-weight: bold; font-size: 9pt;">'
    
    html_content = re.sub(r'<th([^>]*)>', replace_th, html_content)
    
    # td 태그 처리
    def replace_td(match):
        attrs = match.group(1) if match.group(1) else ''
        attrs = re.sub(r'style="[^"]*"', '', attrs)
        attrs = re.sub(r"style='[^']*'", '', attrs)
        return f'<td{attrs} style="border: 1px solid #000000; padding: 5px 6px; text-align: center; vertical-align: middle; font-size: 9pt;">'
    
    html_content = re.sub(r'<td([^>]*)>', replace_td, html_content)
    
    # tr 태그 처리 (테두리 확실히)
    def replace_tr(match):
        attrs = match.group(1) if match.group(1) else ''
        return f'<tr{attrs}>'
    
    html_content = re.sub(r'<tr([^>]*)>', replace_tr, html_content)
    
    # thead 태그 처리
    def replace_thead(match):
        attrs = match.group(1) if match.group(1) else ''
        return f'<thead{attrs} style="background-color: #d9d9d9; font-weight: bold;">'
    
    html_content = re.sub(r'<thead([^>]*)>', replace_thead, html_content)
    
    # h1~h4 태그 처리
    def replace_h1(match):
        attrs = match.group(1) if match.group(1) else ''
        attrs = re.sub(r'style="[^"]*"', '', attrs)
        return f'<h1{attrs} style="font-size: 18pt; font-weight: bold; margin: 15px 0 10px 0; color: #000000;">'
    
    def replace_h2(match):
        attrs = match.group(1) if match.group(1) else ''
        attrs = re.sub(r'style="[^"]*"', '', attrs)
        return f'<h2{attrs} style="font-size: 14pt; font-weight: bold; margin: 12px 0 8px 0; color: #000000;">'
    
    def replace_h3(match):
        attrs = match.group(1) if match.group(1) else ''
        attrs = re.sub(r'style="[^"]*"', '', attrs)
        return f'<h3{attrs} style="font-size: 12pt; font-weight: bold; margin: 10px 0 6px 0; color: #000000;">'
    
    def replace_h4(match):
        attrs = match.group(1) if match.group(1) else ''
        attrs = re.sub(r'style="[^"]*"', '', attrs)
        return f'<h4{attrs} style="font-size: 11pt; font-weight: bold; margin: 8px 0 5px 0; color: #000000;">'
    
    html_content = re.sub(r'<h1([^>]*)>', replace_h1, html_content)
    html_content = re.sub(r'<h2([^>]*)>', replace_h2, html_content)
    html_content = re.sub(r'<h3([^>]*)>', replace_h3, html_content)
    html_content = re.sub(r'<h4([^>]*)>', replace_h4, html_content)
    
    # p 태그 처리
    def replace_p(match):
        attrs = match.group(1) if match.group(1) else ''
        attrs = re.sub(r'style="[^"]*"', '', attrs)
        return f'<p{attrs} style="font-size: 10pt; margin: 5px 0; line-height: 160%; text-align: justify;">'
    
    html_content = re.sub(r'<p([^>]*)>', replace_p, html_content)
    
    # img 태그 처리 (Base64 이미지 지원)
    def replace_img(match):
        attrs = match.group(1) if match.group(1) else ''
        # src 속성 유지
        src_match = re.search(r'src="([^"]*)"', attrs)
        src = src_match.group(0) if src_match else ''
        # 다른 속성들
        other_attrs = re.sub(r'src="[^"]*"', '', attrs)
        other_attrs = re.sub(r'style="[^"]*"', '', other_attrs)
        return f'<img {src} {other_attrs} style="max-width: 100%; height: auto; display: block; margin: 10px auto;">'
    
    html_content = re.sub(r'<img([^>]*)/?>', replace_img, html_content)
    
    # ul, ol, li 태그 처리
    def replace_ul(match):
        attrs = match.group(1) if match.group(1) else ''
        return f'<ul{attrs} style="margin: 10px 0 10px 25px; padding: 0;">'
    
    def replace_ol(match):
        attrs = match.group(1) if match.group(1) else ''
        return f'<ol{attrs} style="margin: 10px 0 10px 25px; padding: 0;">'
    
    def replace_li(match):
        attrs = match.group(1) if match.group(1) else ''
        return f'<li{attrs} style="margin: 3px 0; line-height: 160%;">'
    
    html_content = re.sub(r'<ul([^>]*)>', replace_ul, html_content)
    html_content = re.sub(r'<ol([^>]*)>', replace_ol, html_content)
    html_content = re.sub(r'<li([^>]*)>', replace_li, html_content)
    
    # div 태그 처리 (기본 레이아웃)
    def replace_div(match):
        attrs = match.group(1) if match.group(1) else ''
        # 기존 style이 있으면 유지하면서 기본 스타일 추가
        if 'style=' in attrs:
            return match.group(0)  # 기존 스타일 유지
        return f'<div{attrs} style="margin: 5px 0;">'
    
    html_content = re.sub(r'<div([^>]*)>', replace_div, html_content)
    
    # span 태그 처리
    def replace_span(match):
        attrs = match.group(1) if match.group(1) else ''
        if 'style=' in attrs:
            return match.group(0)
        return f'<span{attrs}>'
    
    html_content = re.sub(r'<span([^>]*)>', replace_span, html_content)
    
    # strong, b 태그 처리
    html_content = re.sub(r'<strong([^>]*)>', r'<strong\1 style="font-weight: bold;">', html_content)
    html_content = re.sub(r'<b([^>]*)>', r'<b\1 style="font-weight: bold;">', html_content)
    
    # em, i 태그 처리
    html_content = re.sub(r'<em([^>]*)>', r'<em\1 style="font-style: italic;">', html_content)
    html_content = re.sub(r'<i([^>]*)>', r'<i\1 style="font-style: italic;">', html_content)
    
    return html_content


def _create_placeholder_image(image_path):
    """플레이스홀더 이미지 생성"""
    try:
        from PIL import Image, ImageDraw, ImageFont
        import os
        
        # 800x400 크기의 플레이스홀더 이미지 생성
        img = Image.new('RGB', (800, 400), color='#f5f5f5')
        draw = ImageDraw.Draw(img)
        
        # 테두리
        draw.rectangle([(0, 0), (799, 399)], outline='#666', width=2)
        
        # 텍스트
        try:
            # 한글 폰트 시도 (시스템 폰트)
            font = ImageFont.truetype('/System/Library/Fonts/AppleGothic.ttf', 24)
        except:
            try:
                font = ImageFont.truetype('/usr/share/fonts/truetype/nanum/NanumGothic.ttf', 24)
            except:
                font = ImageFont.load_default()
        
        text = '📊 차트 영역\n\n브라우저에서 HTML을 열어\nCanvas를 이미지로 변환하세요'
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        position = ((800 - text_width) // 2, (400 - text_height) // 2)
        
        draw.text(position, text, fill='#666', font=font, align='center')
        
        image_path.parent.mkdir(parents=True, exist_ok=True)
        img.save(str(image_path), 'PNG')
        return True
    except ImportError:
        # PIL이 없으면 기본 이미지 생성 스킵
        return False
    except Exception as e:
        print(f"[경고] 플레이스홀더 이미지 생성 실패: {e}")
        return False


@api_bp.route('/export-hwp-ready', methods=['POST'])
def export_hwp_ready():
    """한글(HWP) 복붙용 HTML 문서 생성 - 브라우저에서 열어 전체 복사 후 한글에 붙여넣기
    
    사용법:
    1. 생성된 HTML 파일을 브라우저에서 엽니다
    2. '전체 복사' 버튼 클릭 또는 Ctrl+A → Ctrl+C
    3. 한글(HWP)에서 Ctrl+V로 붙여넣기
    4. 표, 글꼴, 서식이 유지됩니다
    """
    try:
        data = request.get_json()
        pages = data.get('pages', [])
        year = data.get('year', session.get('year', 2025))
        quarter = data.get('quarter', session.get('quarter', 2))
        
        if not pages:
            return jsonify({'success': False, 'error': '페이지 데이터가 없습니다.'})
        
        # 한글 복붙에 최적화된 HTML 생성
        final_html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{year}년 {quarter}/4분기 지역경제동향 - 한글 복붙용</title>
    <style>
        /* 화면 표시용 (복사 시 인라인 스타일이 적용됨) */
        * {{ box-sizing: border-box; }}
        body {{
            font-family: '맑은 고딕', 'Malgun Gothic', sans-serif;
            font-size: 10pt;
            line-height: 1.6;
            color: #000;
            background: #f5f5f5;
            padding: 20px;
            margin: 0;
        }}
        
        #hwp-content {{
            max-width: 210mm;
            margin: 60px auto 20px auto;
            background: white;
            padding: 20mm;
            box-shadow: 0 2px 10px rgba(0,0,0,0.15);
        }}
        
        /* 도구 버튼 */
        .toolbar {{
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            background: linear-gradient(135deg, #1a5276 0%, #2980b9 100%);
            padding: 12px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.3);
            z-index: 9999;
        }}
        
        .toolbar-title {{
            color: white;
            font-size: 14pt;
            font-weight: bold;
        }}
        
        .toolbar-buttons {{
            display: flex;
            gap: 10px;
        }}
        
        .copy-btn {{
            background: #27ae60;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            font-size: 11pt;
            font-weight: bold;
            cursor: pointer;
            transition: background 0.2s;
        }}
        
        .copy-btn:hover {{ background: #219a52; }}
        
        .help-text {{
            color: rgba(255,255,255,0.9);
            font-size: 9pt;
            margin-left: 15px;
        }}
        
        @media print {{ 
            .toolbar {{ display: none; }}
            body {{ background: white; padding: 0; }}
            #hwp-content {{ box-shadow: none; margin: 0; }}
        }}
    </style>
</head>
<body>
    <div class="toolbar">
        <span class="toolbar-title">📋 한글 복붙용 문서</span>
        <div class="toolbar-buttons">
            <button class="copy-btn" onclick="copyAll()">📋 전체 복사</button>
            <button class="copy-btn" style="background: #3498db;" onclick="selectAll()">✓ 전체 선택</button>
            <span class="help-text">복사 후 한글(HWP)에서 Ctrl+V</span>
        </div>
    </div>
    
    <div id="hwp-content">
'''
        
        for idx, page in enumerate(pages, 1):
            page_html = page.get('html', '')
            page_title = page.get('title', f'페이지 {idx}')
            category = page.get('category', '')
            
            # body 내용 추출
            body_content = page_html
            if '<body' in page_html.lower():
                body_match = re.search(r'<body[^>]*>(.*?)</body>', page_html, re.DOTALL | re.IGNORECASE)
                if body_match:
                    body_content = body_match.group(1)
            
            # 불필요한 요소 제거
            body_content = re.sub(r'<style[^>]*>.*?</style>', '', body_content, flags=re.DOTALL)
            body_content = re.sub(r'<script[^>]*>.*?</script>', '', body_content, flags=re.DOTALL)
            body_content = re.sub(r'<link[^>]*/?>', '', body_content)
            body_content = re.sub(r'<meta[^>]*/?>', '', body_content)
            body_content = re.sub(r'<!DOCTYPE[^>]*>', '', body_content)
            body_content = re.sub(r'<html[^>]*>', '', body_content)
            body_content = re.sub(r'</html>', '', body_content)
            body_content = re.sub(r'<head[^>]*>.*?</head>', '', body_content, flags=re.DOTALL)
            
            # canvas, svg → 플레이스홀더
            chart_placeholder = '<p style="border: 2px dashed #888; padding: 20px; text-align: center; background: #f9f9f9; margin: 15px 0; color: #666;">📊 [차트 영역] 별도 이미지로 삽입하세요</p>'
            body_content = re.sub(r'<canvas[^>]*>.*?</canvas>', chart_placeholder, body_content, flags=re.DOTALL)
            body_content = re.sub(r'<canvas[^>]*/?>',  chart_placeholder, body_content)
            body_content = re.sub(r'<svg[^>]*>.*?</svg>', chart_placeholder, body_content, flags=re.DOTALL)
            
            # 인라인 스타일 적용 (한글 복붙 호환)
            body_content = _apply_hwp_inline_styles(body_content)
            
            # 카테고리 한글명
            category_names = {
                'summary': '요약',
                'sectoral': '부문별',
                'regional': '시도별',
                'statistics': '통계표'
            }
            category_name = category_names.get(category, '')
            
            # 페이지 구분
            final_html += f'''
        <!-- 페이지 {idx}: {page_title} -->
        <div style="margin-bottom: 25px; padding-bottom: 15px; border-bottom: 1px solid #ddd;">
            <h2 style="font-size: 13pt; font-weight: bold; color: #1a5276; margin: 0 0 12px 0; padding: 8px 12px; background-color: #eef4f8; border-left: 4px solid #2980b9;">
                {category_name} | {page_title}
            </h2>
            <div style="font-size: 10pt; line-height: 170%;">
{body_content}
            </div>
        </div>
'''
        
        # 문서 끝 및 스크립트
        final_html += f'''
        <p style="text-align: center; font-size: 9pt; color: #888; margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd;">
            {year}년 {quarter}/4분기 지역경제동향 | 총 {len(pages)}페이지
        </p>
    </div>
    
    <script>
        function selectAll() {{
            const content = document.getElementById('hwp-content');
            const range = document.createRange();
            range.selectNodeContents(content);
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
            alert('전체 선택 완료!\\n\\nCtrl+C로 복사한 후,\\n한글(HWP)에서 Ctrl+V로 붙여넣기 하세요.');
        }}
        
        function copyAll() {{
            const content = document.getElementById('hwp-content');
            const range = document.createRange();
            range.selectNodeContents(content);
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
            
            try {{
                const success = document.execCommand('copy');
                if (success) {{
                    alert('✅ 복사 완료!\\n\\n한글(HWP)에서 Ctrl+V로 붙여넣기 하세요.\\n\\n※ 표 테두리, 글꼴, 서식이 유지됩니다.');
                }} else {{
                    throw new Error('copy failed');
                }}
            }} catch (e) {{
                // Clipboard API 시도
                if (navigator.clipboard && window.ClipboardItem) {{
                    const html = content.innerHTML;
                    const blob = new Blob([html], {{ type: 'text/html' }});
                    navigator.clipboard.write([new ClipboardItem({{ 'text/html': blob }})]).then(() => {{
                        alert('✅ 복사 완료! (Clipboard API)\\n\\n한글(HWP)에서 Ctrl+V로 붙여넣기 하세요.');
                    }}).catch(() => {{
                        alert('⚠️ 자동 복사 실패\\n\\n1. Ctrl+A로 전체 선택\\n2. Ctrl+C로 복사\\n3. 한글에서 Ctrl+V');
                    }});
                }} else {{
                    alert('⚠️ 자동 복사 실패\\n\\n1. Ctrl+A로 전체 선택\\n2. Ctrl+C로 복사\\n3. 한글에서 Ctrl+V');
                }}
            }}
            
            selection.removeAllRanges();
        }}
        
        // 키보드 단축키
        document.addEventListener('keydown', function(e) {{
            if ((e.ctrlKey || e.metaKey) && e.key === 'a') {{
                e.preventDefault();
                selectAll();
            }}
        }});
        
        // 페이지 로드 시 안내
        window.onload = function() {{
            console.log('한글 복붙용 HTML 로드 완료');
            console.log('사용법: 전체 복사 버튼 클릭 → 한글에서 Ctrl+V');
        }};
    </script>
</body>
</html>
'''
        
        output_filename = f'지역경제동향_{year}년_{quarter}분기_한글복붙용.html'
        output_path = UPLOAD_FOLDER / output_filename
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        # exports 폴더에도 저장
        exports_path = EXPORT_FOLDER / output_filename
        EXPORT_FOLDER.mkdir(exist_ok=True)
        with open(exports_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        return jsonify({
            'success': True,
            'filename': output_filename,
            'view_url': f'/uploads/{output_filename}',
            'download_url': f'/uploads/{output_filename}',
            'total_pages': len(pages),
            'message': f'한글 복붙용 HTML이 생성되었습니다.\\n\\n사용법:\\n1. 다운로드된 파일을 브라우저에서 엽니다\\n2. "전체 복사" 버튼을 클릭합니다\\n3. 한글(HWP)에서 Ctrl+V로 붙여넣기\\n\\n※ 표 테두리, 글꼴, 서식이 유지됩니다.'
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


@api_bp.route('/save-html-to-project', methods=['POST'])
def save_html_to_project():
    """HTML 보도자료를 프로젝트 메인 디렉토리에 저장"""
    try:
        from datetime import datetime
        
        data = request.get_json()
        pages = data.get('pages', [])
        year = data.get('year', session.get('year', 2025))
        quarter = data.get('quarter', session.get('quarter', 2))
        
        if not pages:
            return jsonify({'success': False, 'error': '페이지 데이터가 없습니다.'})
        
        # 모든 페이지의 스타일 수집
        all_styles = set()
        
        final_html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{year}년 {quarter}/4분기 지역경제동향</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap');
        
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        html, body {{
            width: 210mm;
            background: white;
        }}
        
        body {{
            font-family: 'Noto Sans KR', '맑은 고딕', sans-serif;
        }}
        
        /* PDF 출력용 페이지 스타일 */
        .pdf-page {{
            width: 210mm;
            min-height: 297mm;
            max-height: 297mm;
            padding: 12mm 15mm 15mm 15mm;
            margin: 0 auto 5mm auto;
            background: white;
            position: relative;
            overflow: hidden;
            page-break-after: always;
            page-break-inside: avoid;
        }}
        
        .pdf-page:last-child {{
            page-break-after: auto;
            margin-bottom: 0;
        }}
        
        .pdf-page-content {{
            width: 100%;
            height: calc(297mm - 32mm);
            overflow: hidden;
        }}
        
        .pdf-page-content > * {{
            max-width: 100%;
        }}
        
        /* 페이지 번호 */
        .pdf-page-number {{
            position: absolute;
            bottom: 8mm;
            left: 0;
            right: 0;
            text-align: center;
            font-size: 9pt;
            color: #333;
        }}
        
        /* 화면 미리보기용 */
        @media screen {{
            body {{
                background: #f0f0f0;
                padding: 20px;
            }}
            
            .pdf-page {{
                box-shadow: 0 2px 10px rgba(0,0,0,0.15);
                border: 1px solid #ddd;
            }}
        }}
        
        /* 인쇄/PDF 저장용 */
        @media print {{
            html, body {{
                width: 210mm;
                background: white !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
            
            body {{
                padding: 0;
                margin: 0;
            }}
            
            .pdf-page {{
                width: 210mm;
                height: 297mm;
                min-height: 297mm;
                max-height: 297mm;
                padding: 12mm 15mm 15mm 15mm;
                margin: 0;
                box-shadow: none;
                border: none;
                page-break-after: always;
                page-break-inside: avoid;
            }}
            
            .pdf-page:last-child {{
                page-break-after: auto;
            }}
            
            /* 차트 색상 유지 */
            canvas {{
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
        }}
        
        @page {{
            size: A4 portrait;
            margin: 0;
        }}
        
        /* 표 스타일 공통 */
        table {{
            border-collapse: collapse;
            width: 100%;
        }}
        
        th, td {{
            border: 1px solid #333;
            padding: 4px 6px;
            font-size: 9pt;
            text-align: center;
        }}
        
        th {{
            background: #f5f5f5;
            font-weight: 600;
        }}
        
        /* 차트 크기 조정 */
        .chart-container, .chart-wrapper {{
            max-width: 100%;
        }}
        
        canvas {{
            max-width: 100% !important;
            height: auto !important;
        }}
    </style>
'''
        
        # 각 페이지에서 스타일 추출하여 추가
        for idx, page in enumerate(pages):
            page_html = page.get('html', '')
            if '<style' in page_html:
                style_matches = re.findall(r'<style[^>]*>(.*?)</style>', page_html, re.DOTALL)
                for style in style_matches:
                    # 중복 방지를 위해 hash 사용
                    style_hash = hash(style.strip())
                    if style_hash not in all_styles:
                        all_styles.add(style_hash)
                        final_html += f'    <style>/* Page {idx+1} styles */\n{style}\n    </style>\n'
        
        final_html += '''</head>
<body>
'''
        
        for idx, page in enumerate(pages, 1):
            page_html = page.get('html', '')
            page_title = page.get('title', f'페이지 {idx}')
            
            # body 내용 추출
            body_content = page_html
            if '<body' in page_html.lower():
                body_match = re.search(r'<body[^>]*>(.*?)</body>', page_html, re.DOTALL | re.IGNORECASE)
                if body_match:
                    body_content = body_match.group(1)
            
            # 내용에서 style 태그 제거 (이미 head에 추가됨)
            body_content = re.sub(r'<style[^>]*>.*?</style>', '', body_content, flags=re.DOTALL)
            
            # 페이지 래퍼 추가
            final_html += f'''
    <!-- Page {idx}: {page_title} -->
    <div class="pdf-page" data-page="{idx}" data-title="{page_title}">
        <div class="pdf-page-content">
{body_content}
        </div>
        <div class="pdf-page-number">- {idx} -</div>
    </div>
'''
        
        final_html += '''
    <script>
        // 인쇄 전 준비
        window.onbeforeprint = function() {
            document.body.style.background = 'white';
        };
        
        // Ctrl+P로 PDF 저장 안내
        console.log('PDF 저장: Ctrl+P (또는 Cmd+P) → "PDF로 저장" 선택');
    </script>
</body>
</html>
'''
        
        # 프로젝트 메인 디렉토리에 저장
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f'지역경제동향_{year}년_{quarter}분기_{timestamp}.html'
        output_path = BASE_DIR / output_filename
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        print(f"[HTML 저장] 프로젝트 메인 디렉토리에 저장됨: {output_path}")
        
        return jsonify({
            'success': True,
            'filename': output_filename,
            'path': str(output_path),
            'total_pages': len(pages),
            'message': f'HTML 파일이 프로젝트 메인 디렉토리에 저장되었습니다: {output_filename}'
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})

