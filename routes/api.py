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
from config.reports import REPORT_ORDER, REGIONAL_REPORTS, SUMMARY_REPORTS, STATISTICS_REPORTS
from utils.excel_utils import extract_year_quarter_from_excel, extract_year_quarter_from_raw, detect_file_type
from services.report_generator import (
    generate_report_html,
    generate_regional_report_html,
    generate_statistics_report_html,
    generate_individual_statistics_html
)
from services.grdp_service import get_kosis_grdp_download_info, parse_kosis_grdp_file
from services.excel_processor import preprocess_excel, check_available_methods, get_recommended_method
from data_converter import DataConverter
import openpyxl

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


@api_bp.route('/upload', methods=['POST'])
def upload_excel():
    """엑셀 파일 업로드
    
    프로세스 1: 기초자료 수집표 → 분석표 생성
    프로세스 2: 분석표 → GRDP 결합 → 지역경제동향 생성
    """
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '파일이 없습니다'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '파일이 선택되지 않았습니다'})
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'error': '엑셀 파일만 업로드 가능합니다'})
    
    # 새 파일 업로드 전 이전 파일 정리 (현재 세션 파일 제외)
    cleanup_upload_folder(keep_current_files=False)
    
    # 한글 파일명 보존하면서 안전한 파일명 생성
    filename = safe_filename(file.filename)
    filepath = Path(UPLOAD_FOLDER) / filename
    file.save(str(filepath))
    
    # 저장된 파일 크기 확인 (데이터 유실 방지)
    saved_size = filepath.stat().st_size
    print(f"[업로드] 파일 저장 완료: {filename} ({saved_size:,} bytes)")
    
    # 파일 유형 자동 감지
    file_type = detect_file_type(str(filepath))
    
    # ===== 프로세스 2: 분석표 업로드 → GRDP 결합 → 지역경제동향 생성 =====
    if file_type == 'analysis':
        print(f"\n{'='*50}")
        print(f"[프로세스 2] 분석표 업로드: {filename}")
        print(f"{'='*50}")
        
        # 수식 계산 전처리 (분석 시트의 수식을 계산)
        print(f"[전처리] 엑셀 수식 계산 시작...")
        processed_path, preprocess_success, preprocess_msg = preprocess_excel(str(filepath))
        
        if preprocess_success:
            print(f"[전처리] 성공: {preprocess_msg}")
            # 전처리된 파일 경로 사용
            filepath = Path(processed_path)
        else:
            print(f"[전처리] {preprocess_msg} - generator fallback 로직 사용")
        
        # 연도/분기 추출
        year, quarter = extract_year_quarter_from_excel(str(filepath))
        
        # GRDP 시트 존재 여부 확인 및 데이터 추출
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
                    
                    # GRDP 시트에서 데이터 추출
                    grdp_data = _extract_grdp_from_analysis_sheet(wb[sheet_name], year, quarter)
                    if grdp_data:
                        print(f"[GRDP] 데이터 추출 성공 - 전국: {grdp_data.get('national_summary', {}).get('growth_rate', 0)}%")
                    break
            wb.close()
        except Exception as e:
            print(f"[경고] GRDP 시트 확인 실패: {e}")
        
        # 세션에 저장 (파일 수정 시간 포함)
        session['excel_path'] = str(filepath)
        session['year'] = year
        session['quarter'] = quarter
        session['file_type'] = 'analysis'
        try:
            session['excel_file_mtime'] = Path(filepath).stat().st_mtime
        except OSError:
            pass  # 파일 시간 확인 실패는 무시
        
        if grdp_data:
            session['grdp_data'] = grdp_data
            # JSON 파일로도 저장
            grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
            try:
                with open(grdp_json_path, 'w', encoding='utf-8') as f:
                    json.dump(grdp_data, f, ensure_ascii=False, indent=2)
            except Exception as e:
                print(f"[경고] GRDP JSON 저장 실패: {e}")
        
        print(f"[결과] GRDP {'있음' if has_grdp else '없음'} → {'바로 보고서 생성' if has_grdp else 'GRDP 모달 표시'}")
        
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
    
    # ===== 프로세스 1: 기초자료 수집표 → 분석표 생성 =====
    print(f"\n{'='*50}")
    print(f"[프로세스 1] 기초자료 수집표 업로드: {filename}")
    print(f"{'='*50}")
    
    try:
        converter = DataConverter(str(filepath))
        year = converter.year
        quarter = converter.quarter
        
        # 분석표 파일명 (다운로드 시 생성)
        analysis_filename = f"분석표_{year}년_{quarter}분기_자동생성.xlsx"
        
        conversion_info = {
            'original_file': filename,
            'analysis_file': analysis_filename,
            'year': year,
            'quarter': quarter
        }
        
        print(f"[결과] 분석표 다운로드 준비 완료: {analysis_filename}")
        
    except Exception as e:
        import traceback
        print(f"[오류] 기초자료 처리 실패: {e}")
        traceback.print_exc()
        return jsonify({
            'success': False,
            'error': f'기초자료 처리 중 오류가 발생했습니다: {str(e)}'
        })
    
    # 세션에 저장 (파일 수정 시간 포함)
    session['raw_excel_path'] = str(filepath)
    session['year'] = year
    session['quarter'] = quarter
    session['file_type'] = 'raw'
    try:
        session['raw_file_mtime'] = Path(filepath).stat().st_mtime
    except OSError:
        pass  # 파일 시간 확인 실패는 무시
    
    return jsonify({
        'success': True,
        'filename': filename,
        'file_type': 'raw',
        'year': year,
        'quarter': quarter,
        'reports': REPORT_ORDER,
        'regional_reports': REGIONAL_REPORTS,
        'conversion_info': conversion_info
    })


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


@api_bp.route('/check-grdp', methods=['GET'])
def check_grdp_status():
    """GRDP 데이터 상태 확인"""
    grdp_data = session.get('grdp_data')
    grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
    
    if grdp_data:
        return jsonify({
            'success': True,
            'has_grdp': True,
            'source': grdp_data.get('source', 'session'),
            'national_growth_rate': grdp_data.get('national_summary', {}).get('growth_rate', 0)
        })
    elif grdp_json_path.exists():
        return jsonify({
            'success': True,
            'has_grdp': True,
            'source': 'json_file'
        })
    else:
        kosis_info = get_kosis_grdp_download_info()
        return jsonify({
            'success': True,
            'has_grdp': False,
            'kosis_info': kosis_info
        })


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
    """GRDP 파일이 없을 때 기본값(placeholder) 사용"""
    from services.grdp_service import get_default_grdp_data
    
    year = session.get('year', 2025)
    quarter = session.get('quarter', 2)
    
    # 기본 GRDP 데이터 생성 (placeholder)
    grdp_data = get_default_grdp_data(year, quarter)
    
    # 세션에 저장
    session['grdp_data'] = grdp_data
    
    # JSON 파일로도 저장
    grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
    with open(grdp_json_path, 'w', encoding='utf-8') as f:
        json.dump(grdp_data, f, ensure_ascii=False, indent=2)
    
    # 분석표에 플레이스홀더 GRDP 시트 추가 (분석표가 업로드된 경우)
    analysis_path = session.get('excel_path')
    grdp_sheet_added = False
    
    if analysis_path and Path(analysis_path).exists():
        try:
            grdp_sheet_added = _add_placeholder_grdp_sheet(analysis_path, grdp_data)
            if grdp_sheet_added:
                print(f"[GRDP] 분석표에 플레이스홀더 GRDP 시트 추가 완료")
        except Exception as e:
            print(f"[GRDP] 분석표에 GRDP 시트 추가 실패: {e}")
    
    print(f"[GRDP] 기본값 사용 - {year}년 {quarter}분기")
    
    return jsonify({
        'success': True,
        'message': 'GRDP 기본값이 설정되었습니다. 나중에 실제 데이터로 업데이트할 수 있습니다.',
        'is_placeholder': True,
        'national_growth_rate': 0.0,
        'kosis_info': grdp_data.get('kosis_info', {}),
        'grdp_sheet_added': grdp_sheet_added
    })


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


@api_bp.route('/download-analysis', methods=['GET'])
def download_analysis():
    """분석표 다운로드 (다운로드 시점에 생성 + 수식 계산)"""
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
    """가중치 설정을 포함하여 분석표 생성 + 다운로드"""
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
    """현재 보고서 순서 반환"""
    return jsonify({'reports': REPORT_ORDER, 'regional_reports': REGIONAL_REPORTS})


@api_bp.route('/report-order', methods=['POST'])
def update_report_order():
    """보고서 순서 업데이트"""
    from config import reports as reports_module
    data = request.get_json()
    new_order = data.get('order', [])
    
    if new_order:
        order_map = {r['id']: idx for idx, r in enumerate(new_order)}
        reports_module.REPORT_ORDER = sorted(reports_module.REPORT_ORDER, key=lambda x: order_map.get(x['id'], 999))
    
    return jsonify({'success': True, 'reports': reports_module.REPORT_ORDER})


@api_bp.route('/session-info', methods=['GET'])
def get_session_info():
    """현재 세션 정보 반환"""
    return jsonify({
        'excel_path': session.get('excel_path'),
        'year': session.get('year'),
        'quarter': session.get('quarter'),
        'has_file': bool(session.get('excel_path'))
    })


@api_bp.route('/generate-all', methods=['POST'])
def generate_all_reports():
    """모든 보고서 일괄 생성"""
    data = request.get_json()
    year = data.get('year', session.get('year', 2025))
    quarter = data.get('quarter', session.get('quarter', 2))
    all_custom_data = data.get('all_custom_data', {})
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    generated_reports = []
    errors = []
    
    for report_config in REPORT_ORDER:
        custom_data = all_custom_data.get(report_config['id'], {})
        raw_excel_path = session.get('raw_excel_path')
        
        html_content, error, _ = generate_report_html(
            excel_path, report_config, year, quarter, custom_data, raw_excel_path
        )
        
        if error:
            errors.append({'report_id': report_config['id'], 'error': error})
        else:
            output_path = TEMPLATES_DIR / f"{report_config['name']}_output.html"
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            generated_reports.append({
                'report_id': report_config['id'],
                'name': report_config['name'],
                'path': str(output_path)
            })
    
    return jsonify({
        'success': len(errors) == 0,
        'generated': generated_reports,
        'errors': errors
    })


@api_bp.route('/generate-all-regional', methods=['POST'])
def generate_all_regional_reports():
    """시도별 보고서 전체 생성"""
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    generated_reports = []
    errors = []
    
    output_dir = TEMPLATES_DIR / 'regional_output'
    output_dir.mkdir(exist_ok=True)
    
    for region_config in REGIONAL_REPORTS:
        html_content, error = generate_regional_report_html(excel_path, region_config['name'])
        
        if error:
            errors.append({'region_id': region_config['id'], 'error': error})
        else:
            output_path = output_dir / f"{region_config['name']}_output.html"
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            generated_reports.append({
                'region_id': region_config['id'],
                'name': region_config['name'],
                'path': str(output_path)
            })
    
    return jsonify({
        'success': len(errors) == 0,
        'generated': generated_reports,
        'errors': errors
    })


@api_bp.route('/export-final', methods=['POST'])
def export_final_document():
    """모든 보고서를 PDF 출력용 HTML 문서로 합치기"""
    try:
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
                import re
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
                import re
                body_match = re.search(r'<body[^>]*>(.*?)</body>', page_html, re.DOTALL | re.IGNORECASE)
                if body_match:
                    body_content = body_match.group(1)
            
            # 내용에서 style 태그 제거 (이미 head에 추가됨)
            import re
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
        
        output_filename = f'지역경제동향_{year}년_{quarter}분기_PDF용.html'
        output_path = UPLOAD_FOLDER / output_filename
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        return jsonify({
            'success': True,
            'html': final_html,
            'filename': output_filename,
            'download_url': f'/uploads/{output_filename}',
            'total_pages': len(pages)
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


@api_bp.route('/render-chart-image', methods=['POST'])
def render_chart_image():
    """차트/인포그래픽을 이미지로 렌더링"""
    try:
        data = request.get_json()
        image_data = data.get('image_data', '')
        filename = data.get('filename', 'chart.png')
        
        if not image_data:
            return jsonify({'success': False, 'error': '이미지 데이터가 없습니다.'})
        
        match = re.match(r'data:([^;]+);base64,(.+)', image_data)
        if match:
            mimetype = match.group(1)
            img_data = base64.b64decode(match.group(2))
            
            img_path = UPLOAD_FOLDER / filename
            with open(img_path, 'wb') as f:
                f.write(img_data)
            
            return jsonify({
                'success': True,
                'filename': filename,
                'path': str(img_path),
                'url': f'/uploads/{filename}'
            })
        else:
            return jsonify({'success': False, 'error': '잘못된 이미지 데이터 형식입니다.'})
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


@api_bp.route('/get-industry-weights', methods=['GET'])
def get_industry_weights():
    """기초자료에서 업종별 가중치 정보 추출"""
    import pandas as pd
    
    sheet_type = request.args.get('sheet_type', '광공업생산')
    raw_excel_path = session.get('raw_excel_path')
    
    if not raw_excel_path or not Path(raw_excel_path).exists():
        return jsonify({
            'success': False, 
            'error': '기초자료 파일을 찾을 수 없습니다. 먼저 파일을 업로드하세요.'
        })
    
    try:
        xl = pd.ExcelFile(raw_excel_path)
        
        # 시트 매핑
        sheet_mapping = {
            '광공업생산': '광공업생산',
            '서비스업생산': '서비스업생산'
        }
        
        sheet_name = sheet_mapping.get(sheet_type)
        if not sheet_name or sheet_name not in xl.sheet_names:
            return jsonify({
                'success': False,
                'error': f'시트를 찾을 수 없습니다: {sheet_type}'
            })
        
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        
        # 업종별 정보 추출 (열 구조에 따라 다름)
        industries = []
        
        if sheet_type == '광공업생산':
            # 광공업생산 시트: 열 4=업종명, 열 8=가중치 (또는 해당 열 확인 필요)
            name_col = 4  # 업종명 열
            weight_col = 8  # 가중치 열
            
            for i, row in df.iterrows():
                if i < 3:  # 헤더 행 건너뛰기
                    continue
                    
                name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ''
                if not name or name in ['nan', 'NaN', '업종이름', '업종명']:
                    continue
                    
                weight = None
                if weight_col < len(row) and pd.notna(row[weight_col]):
                    try:
                        weight = float(row[weight_col])
                    except (ValueError, TypeError):
                        pass
                
                industries.append({
                    'row': i + 1,
                    'name': name,
                    'weight': weight
                })
                
        elif sheet_type == '서비스업생산':
            # 서비스업생산 시트: 열 4=업종명, 열 8=가중치
            name_col = 4  # 업종명 열
            weight_col = 8  # 가중치 열
            
            for i, row in df.iterrows():
                if i < 3:  # 헤더 행 건너뛰기
                    continue
                    
                name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ''
                if not name or name in ['nan', 'NaN', '업종이름', '업종명']:
                    continue
                    
                weight = None
                if weight_col < len(row) and pd.notna(row[weight_col]):
                    try:
                        weight = float(row[weight_col])
                    except (ValueError, TypeError):
                        pass
                
                industries.append({
                    'row': i + 1,
                    'name': name,
                    'weight': weight
                })
        
        return jsonify({
            'success': True,
            'sheet_type': sheet_type,
            'industries': industries[:100]  # 최대 100개
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': f'업종 정보 추출 실패: {str(e)}'})


@api_bp.route('/export-hwp-import', methods=['POST'])
def export_hwp_import():
    """한글 프로그램 불러오기용 HTML 문서 생성 - 차트를 이미지로 변환"""
    try:
        import base64
        import hashlib
        from datetime import datetime
        
        data = request.get_json()
        pages = data.get('pages', [])
        year = data.get('year', session.get('year', 2025))
        quarter = data.get('quarter', session.get('quarter', 2))
        
        if not pages:
            return jsonify({'success': False, 'error': '페이지 데이터가 없습니다.'})
        
        # 출력 폴더 및 이미지 폴더 생성
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        export_dir_name = f'지역경제동향_{year}년_{quarter}분기_한글불러오기용_{timestamp}'
        export_dir = EXPORT_FOLDER / export_dir_name
        export_dir.mkdir(exist_ok=True, parents=True)
        images_dir = export_dir / 'images'
        images_dir.mkdir(exist_ok=True)
        
        # 한글 불러오기용 HTML 생성
        final_html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{year}년 {quarter}/4분기 지역경제동향 - 한글 불러오기용</title>
    <style>
        @page {{
            size: A4;
            margin: 15mm 20mm;
        }}
        
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: '맑은 고딕', 'Malgun Gothic', '돋움', 'Dotum', sans-serif;
            font-size: 10pt;
            line-height: 1.6;
            color: #000;
            background: #fff;
            width: 210mm;
            margin: 0 auto;
            padding: 0;
        }}
        
        /* A4 페이지 단위 */
        .hwp-page {{
            width: 210mm;
            min-height: 297mm;
            padding: 15mm 20mm;
            margin: 0;
            background: #fff;
            page-break-after: always;
            page-break-inside: avoid;
        }}
        
        .hwp-page:last-child {{
            page-break-after: auto;
        }}
        
        /* 표 스타일 (한글 호환) */
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 10px 0;
            font-size: 9pt;
            border: 1px solid #000;
        }}
        
        th, td {{
            border: 1px solid #000;
            padding: 5px 8px;
            text-align: center;
            vertical-align: middle;
        }}
        
        th {{
            background-color: #d9d9d9;
            font-weight: bold;
        }}
        
        /* 차트 이미지 스타일 */
        .chart-image {{
            display: block;
            max-width: 100%;
            height: auto;
            margin: 10px auto;
            border: 1px solid #ddd;
        }}
        
        /* 제목 스타일 */
        h1, h2, h3, h4 {{
            font-family: '맑은 고딕', 'Malgun Gothic', sans-serif;
            margin: 15px 0 10px 0;
        }}
        
        h1 {{ font-size: 16pt; font-weight: bold; }}
        h2 {{ font-size: 14pt; font-weight: bold; }}
        h3 {{ font-size: 12pt; font-weight: bold; }}
        h4 {{ font-size: 11pt; font-weight: bold; }}
        
        p {{
            margin: 5px 0;
            line-height: 1.6;
        }}
        
        /* 페이지 구분선 */
        .page-divider {{
            border-top: 2px solid #333;
            margin: 20px 0;
            padding-top: 20px;
        }}
    </style>
</head>
<body>
'''
        
        # 페이지 처리 (Canvas를 이미지 플레이스홀더로 대체)
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
            
            # 스타일, 스크립트, link, meta 태그 제거
            body_content = re.sub(r'<style[^>]*>.*?</style>', '', body_content, flags=re.DOTALL)
            body_content = re.sub(r'<script[^>]*>.*?</script>', '', body_content, flags=re.DOTALL)
            body_content = re.sub(r'<link[^>]*>', '', body_content)
            body_content = re.sub(r'<meta[^>]*>', '', body_content)
            
            # Canvas를 이미지 플레이스홀더로 대체
            # Canvas ID나 클래스를 기반으로 이미지 경로 생성
            canvas_pattern = r'<canvas[^>]*id=["\']([^"\']+)["\'][^>]*>.*?</canvas>'
            canvas_matches = list(re.finditer(canvas_pattern, body_content, re.DOTALL))
            
            for match in reversed(canvas_matches):  # 역순으로 처리하여 인덱스 유지
                canvas_id = match.group(1)
                # 이미지 파일명 생성 (페이지 번호 + 캔버스 ID)
                image_filename = f'page{idx}_{canvas_id}.png'
                image_path = f'images/{image_filename}'
                
                # Canvas를 <img> 태그로 대체
                img_tag = f'<img src="{image_path}" alt="차트: {canvas_id}" class="chart-image" style="max-width: 100%; height: auto;" />'
                body_content = body_content[:match.start()] + img_tag + body_content[match.end():]
            
            # 이름 없는 Canvas도 처리
            body_content = re.sub(
                r'<canvas[^>]*>.*?</canvas>',
                '<img src="images/placeholder_chart.png" alt="차트" class="chart-image" style="max-width: 100%; height: auto;" />',
                body_content,
                flags=re.DOTALL
            )
            body_content = re.sub(
                r'<canvas[^>]*/?>',
                '<img src="images/placeholder_chart.png" alt="차트" class="chart-image" style="max-width: 100%; height: auto;" />',
                body_content
            )
            
            # SVG도 이미지로 대체
            body_content = re.sub(
                r'<svg[^>]*>.*?</svg>',
                '<img src="images/placeholder_chart.png" alt="차트" class="chart-image" style="max-width: 100%; height: auto;" />',
                body_content,
                flags=re.DOTALL
            )
            
            # 표에 인라인 스타일 추가 (한글 호환성)
            body_content = _add_table_inline_styles(body_content)
            
            # 카테고리 한글명
            category_names = {
                'summary': '요약',
                'sectoral': '부문별',
                'regional': '시도별',
                'statistics': '통계표'
            }
            category_name = category_names.get(category, '')
            
            # 페이지 래퍼 추가
            final_html += f'''
    <!-- 페이지 {idx}: {page_title} -->
    <div class="hwp-page">
        <h2 style="font-size: 14pt; font-weight: bold; margin-bottom: 15px; padding: 8px 12px; background-color: #e8e8e8; border-left: 4px solid #0066cc;">
            [{category_name}] {page_title}
        </h2>
        {body_content}
        <div class="page-divider"></div>
        <p style="text-align: center; font-size: 9pt; color: #666; margin-top: 20px;">- {idx} / {len(pages)} -</p>
    </div>
'''
        
        # Canvas를 이미지로 변환하는 스크립트 추가 (Chart.js 포함)
        final_html += '''
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <script>
        // Chart.js가 로드될 때까지 대기
        let chartInstances = {};
        let conversionComplete = false;
        
        // Canvas를 이미지로 변환하는 함수
        function convertCanvasToImage() {
            if (conversionComplete) return;
            
            const canvases = document.querySelectorAll('canvas');
            if (canvases.length === 0) {
                console.log('변환할 Canvas가 없습니다.');
                // Canvas가 없으면 이미 모두 이미지로 변환된 것으로 간주
                conversionComplete = true;
                return;
            }
            
            console.log(`${canvases.length}개의 Canvas를 찾았습니다.`);
            let convertedCount = 0;
            
            canvases.forEach(function(canvas, index) {
                try {
                    // Canvas가 실제로 그려졌는지 확인
                    const ctx = canvas.getContext('2d');
                    const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
                    const isBlank = imageData.data.every(function(pixel) {
                        return pixel === 0; // 투명 또는 검은색만 있으면 빈 Canvas
                    });
                    
                    if (isBlank && canvas.width > 0 && canvas.height > 0) {
                        console.log(`Canvas ${canvas.id || index}가 비어있습니다. 잠시 후 다시 시도합니다.`);
                        return; // 아직 렌더링 중
                    }
                    
                    // Canvas를 이미지로 변환 (Base64)
                    const imageDataUrl = canvas.toDataURL('image/png', 1.0);
                    
                    // Base64를 Blob으로 변환
                    const blob = base64ToBlob(imageDataUrl);
                    
                    // 다운로드 링크 생성
                    const canvasId = canvas.id || `canvas_${index}`;
                    const filename = `${canvasId}.png`;
                    downloadBlob(blob, filename);
                    
                    // Canvas를 <img> 태그로 대체
                    const img = document.createElement('img');
                    img.src = imageDataUrl;
                    img.className = 'chart-image';
                    img.style.maxWidth = '100%';
                    img.style.height = 'auto';
                    img.alt = `차트: ${canvasId}`;
                    
                    // 기존 Canvas의 부모 요소 찾기
                    if (canvas.parentNode) {
                        canvas.parentNode.replaceChild(img, canvas);
                    }
                    
                    convertedCount++;
                    console.log(`✓ ${canvasId} 변환 완료`);
                } catch (e) {
                    console.error(`Canvas 변환 오류 (${canvas.id || index}):`, e);
                }
            });
            
            if (convertedCount > 0) {
                conversionComplete = true;
                // 모든 변환 완료 후 안내
                setTimeout(function() {
                    alert(`Canvas 차트 ${convertedCount}개 변환 완료!\\n\\n변환된 이미지가 자동으로 다운로드되었습니다.\\n다운로드 폴더에서 images 폴더로 이동한 후 한글에서 HTML을 불러오세요.`);
                }, 500);
            }
        }
        
        // Base64를 Blob으로 변환
        function base64ToBlob(base64) {
            const parts = base64.split(';base64,');
            const contentType = parts[0].split(':')[1];
            const raw = window.atob(parts[1]);
            const rawLength = raw.length;
            const uInt8Array = new Uint8Array(rawLength);
            
            for (let i = 0; i < rawLength; ++i) {
                uInt8Array[i] = raw.charCodeAt(i);
            }
            
            return new Blob([uInt8Array], { type: contentType });
        }
        
        // Blob 다운로드
        function downloadBlob(blob, filename) {
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }
        
        // Chart.js 로드 후 차트 렌더링 및 변환
        window.addEventListener('load', function() {
            // 차트가 있는 경우 렌더링 대기
            setTimeout(function() {
                // Chart.js로 렌더링된 차트가 있으면 추가 대기
                setTimeout(convertCanvasToImage, 3000);
            }, 1000);
            
            // 수동 변환 버튼 추가
            const btn = document.createElement('button');
            btn.textContent = '📸 Canvas를 이미지로 변환';
            btn.style.cssText = 'position: fixed; top: 10px; right: 10px; z-index: 10000; padding: 10px 20px; background: #0066cc; color: white; border: none; border-radius: 5px; cursor: pointer; font-size: 12pt; box-shadow: 0 2px 10px rgba(0,0,0,0.3);';
            btn.onclick = function() {
                conversionComplete = false;
                convertCanvasToImage();
            };
            document.body.appendChild(btn);
        });
    </script>
</body>
</html>
'''
        
        # HTML 파일 저장
        html_filename = f'지역경제동향_{year}년_{quarter}분기.html'
        html_path = export_dir / html_filename
        
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        # 플레이스홀더 이미지 생성 (차트가 없는 경우를 위해)
        placeholder_info = _create_placeholder_image(images_dir / 'placeholder_chart.png')
        
        # 사용 안내 파일 생성
        readme_path = export_dir / '사용안내.txt'
        with open(readme_path, 'w', encoding='utf-8') as f:
            f.write(f'''한글 프로그램 불러오기용 파일

생성일: {datetime.now().strftime('%Y년 %m월 %d일 %H:%M:%S')}
연도/분기: {year}년 {quarter}분기

[사용 방법]

1. 한글 프로그램에서:
   - 파일 → 불러오기 → HTML 파일 선택
   - "{html_filename}" 파일 선택
   
2. 차트 이미지 변환:
   - HTML 파일을 브라우저에서 열기
   - Canvas 차트가 자동으로 렌더링됨
   - 브라우저 개발자 도구에서 Canvas를 이미지로 저장
   - 저장한 이미지를 "images" 폴더에 해당 파일명으로 저장
   
3. 한글에서 다시 불러오기:
   - 이미지가 모두 준비된 후 한글에서 HTML 불러오기
   - 차트가 이미지로 표시됩니다

[폴더 구조]
{html_filename}  ← 한글에서 불러올 메인 파일
images/          ← 차트 이미지 폴더
  ├── page1_chart-manufacturing.png
  ├── page2_chart-service.png
  └── ...

[주의사항]
- HTML 파일과 images 폴더는 같은 위치에 있어야 합니다
- 파일을 이동할 때는 HTML 파일과 images 폴더를 함께 이동하세요
- 한글에서 불러온 후 레이아웃을 확인하고 필요시 조정하세요
''')
        
        return jsonify({
            'success': True,
            'html': final_html,
            'filename': html_filename,
            'export_dir': export_dir_name,
            'html_path': str(html_path.relative_to(BASE_DIR)),
            'view_url': f'/exports/{export_dir_name}/{html_filename}',
            'download_url': f'/download-export/{export_dir_name}',
            'total_pages': len(pages),
            'message': '한글 불러오기용 HTML이 생성되었습니다. 브라우저에서 열어 차트를 이미지로 변환한 후 한글에서 불러오세요.'
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


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
    """한글(HWP) 복붙용 HTML 문서 생성 - 인라인 스타일 최적화"""
    try:
        data = request.get_json()
        pages = data.get('pages', [])
        year = data.get('year', session.get('year', 2025))
        quarter = data.get('quarter', session.get('quarter', 2))
        
        if not pages:
            return jsonify({'success': False, 'error': '페이지 데이터가 없습니다.'})
        
        # 한글 복붙에 최적화된 HTML 생성 (인라인 스타일 사용)
        final_html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{year}년 {quarter}/4분기 지역경제동향 - 한글 복붙용</title>
    <style>
        /* 브라우저 미리보기용 스타일 (한글 복붙 시에는 인라인 스타일 적용됨) */
        body {{
            font-family: '맑은 고딕', 'Malgun Gothic', sans-serif;
            font-size: 10pt;
            line-height: 1.6;
            color: #000;
            background: #fff;
            padding: 20px;
            max-width: 210mm;
            margin: 0 auto;
        }}
        .copy-btn {{
            position: fixed;
            top: 10px;
            right: 10px;
            background: #0066cc;
            color: white;
            padding: 12px 20px;
            border: none;
            border-radius: 5px;
            font-size: 12pt;
            cursor: pointer;
            z-index: 9999;
            box-shadow: 0 2px 10px rgba(0,0,0,0.3);
        }}
        .copy-btn:hover {{ background: #0055aa; }}
        @media print {{ .copy-btn {{ display: none; }} }}
    </style>
</head>
<body>
    <button class="copy-btn" onclick="copyAll()">📋 전체 복사 (클릭)</button>
    
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
            
            # 한글 복붙에 불필요한 요소 제거
            body_content = re.sub(r'<style[^>]*>.*?</style>', '', body_content, flags=re.DOTALL)
            body_content = re.sub(r'<script[^>]*>.*?</script>', '', body_content, flags=re.DOTALL)
            body_content = re.sub(r'<link[^>]*>', '', body_content)
            body_content = re.sub(r'<meta[^>]*>', '', body_content)
            
            # canvas를 차트 플레이스홀더로 대체 (인라인 스타일)
            chart_placeholder = '<div style="border: 2px dashed #666; padding: 15px; text-align: center; background: #f5f5f5; margin: 10px 0;">📊 [차트 영역 - 별도 이미지 삽입]</div>'
            body_content = re.sub(r'<canvas[^>]*>.*?</canvas>', chart_placeholder, body_content, flags=re.DOTALL)
            body_content = re.sub(r'<canvas[^>]*/?>',  chart_placeholder, body_content)
            
            # SVG 제거 (복잡한 차트)
            body_content = re.sub(r'<svg[^>]*>.*?</svg>', chart_placeholder, body_content, flags=re.DOTALL)
            
            # 표에 인라인 border 스타일 추가 (한글에서 표 테두리 인식)
            body_content = _add_table_inline_styles(body_content)
            
            # 카테고리 한글명
            category_names = {
                'summary': '요약',
                'sectoral': '부문별',
                'regional': '시도별',
                'statistics': '통계표'
            }
            category_name = category_names.get(category, '')
            
            # 페이지 구분 (인라인 스타일로)
            final_html += f'''
        <!-- 페이지 {idx}: {page_title} -->
        <div style="margin-bottom: 30px; padding-bottom: 20px; border-bottom: 2px solid #333; page-break-after: always;">
            <h2 style="font-family: '맑은 고딕', sans-serif; font-size: 14pt; font-weight: bold; color: #1a1a1a; margin-bottom: 15px; padding: 8px 12px; background-color: #e8e8e8; border-left: 4px solid #0066cc;">
                [{category_name}] {page_title}
            </h2>
            <div style="font-family: '맑은 고딕', sans-serif; font-size: 10pt; line-height: 1.6;">
{body_content}
            </div>
            <p style="text-align: center; font-size: 9pt; color: #666; margin-top: 20px;">- {idx} / {len(pages)} -</p>
        </div>
'''
        
        final_html += '''
    </div>
    
    <script>
        function copyAll() {
            const content = document.getElementById('hwp-content');
            const range = document.createRange();
            range.selectNodeContents(content);
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
            
            try {
                document.execCommand('copy');
                alert('복사 완료!\\n\\n한글(HWP)에서 Ctrl+V로 붙여넣기 하세요.\\n※ 표와 서식이 유지됩니다.');
            } catch (e) {
                alert('자동 복사 실패.\\nCtrl+A로 전체 선택 후 Ctrl+C로 복사하세요.');
            }
            
            selection.removeAllRanges();
        }
        
        // 단축키 지원
        document.addEventListener('keydown', function(e) {
            if (e.ctrlKey && e.key === 'a') {
                e.preventDefault();
                copyAll();
            }
        });
    </script>
</body>
</html>
'''
        
        output_filename = f'지역경제동향_{year}년_{quarter}분기_한글복붙용.html'
        output_path = UPLOAD_FOLDER / output_filename
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        return jsonify({
            'success': True,
            'html': final_html,
            'filename': output_filename,
            'view_url': f'/view/{output_filename}',
            'download_url': f'/uploads/{output_filename}',
            'total_pages': len(pages)
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


def _add_table_inline_styles(html_content):
    """표에 인라인 스타일 추가 (한글 복붙 최적화)"""
    # table 태그에 인라인 스타일 추가
    html_content = re.sub(
        r'<table([^>]*)>',
        r'<table\1 style="border-collapse: collapse; width: 100%; margin: 10px 0; font-family: \'맑은 고딕\', sans-serif; font-size: 9pt;">',
        html_content
    )
    
    # th 태그에 인라인 스타일 추가
    html_content = re.sub(
        r'<th([^>]*)>',
        r'<th\1 style="border: 1px solid #000; padding: 5px 8px; text-align: center; vertical-align: middle; background-color: #d9d9d9; font-weight: bold;">',
        html_content
    )
    
    # td 태그에 인라인 스타일 추가
    html_content = re.sub(
        r'<td([^>]*)>',
        r'<td\1 style="border: 1px solid #000; padding: 5px 8px; text-align: center; vertical-align: middle;">',
        html_content
    )
    
    # 제목 태그들에 인라인 스타일 추가
    html_content = re.sub(
        r'<h1([^>]*)>',
        r'<h1\1 style="font-family: \'맑은 고딕\', sans-serif; font-size: 16pt; font-weight: bold; margin: 15px 0 10px 0;">',
        html_content
    )
    html_content = re.sub(
        r'<h2([^>]*)>',
        r'<h2\1 style="font-family: \'맑은 고딕\', sans-serif; font-size: 14pt; font-weight: bold; margin: 15px 0 10px 0;">',
        html_content
    )
    html_content = re.sub(
        r'<h3([^>]*)>',
        r'<h3\1 style="font-family: \'맑은 고딕\', sans-serif; font-size: 12pt; font-weight: bold; margin: 10px 0 8px 0;">',
        html_content
    )
    html_content = re.sub(
        r'<h4([^>]*)>',
        r'<h4\1 style="font-family: \'맑은 고딕\', sans-serif; font-size: 11pt; font-weight: bold; margin: 10px 0 5px 0;">',
        html_content
    )
    
    # p 태그에 스타일 추가
    html_content = re.sub(
        r'<p([^>]*)>',
        r'<p\1 style="font-family: \'맑은 고딕\', sans-serif; margin: 5px 0; line-height: 1.6;">',
        html_content
    )
    
    # ul, ol 태그에 스타일 추가
    html_content = re.sub(
        r'<ul([^>]*)>',
        r'<ul\1 style="margin: 10px 0 10px 25px; font-family: \'맑은 고딕\', sans-serif;">',
        html_content
    )
    html_content = re.sub(
        r'<ol([^>]*)>',
        r'<ol\1 style="margin: 10px 0 10px 25px; font-family: \'맑은 고딕\', sans-serif;">',
        html_content
    )
    html_content = re.sub(
        r'<li([^>]*)>',
        r'<li\1 style="margin: 3px 0;">',
        html_content
    )
    
    return html_content

