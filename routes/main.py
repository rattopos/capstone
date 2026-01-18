# -*- coding: utf-8 -*-
"""
메인 페이지 라우트
"""

from flask import Blueprint, render_template, send_file, make_response, session
from pathlib import Path
from urllib.parse import quote

from config.reports import REPORT_ORDER, REGIONAL_REPORTS
from config.settings import TEMPLATES_DIR, UPLOAD_FOLDER, DEBUG_FOLDER, EXPORT_FOLDER, BASE_DIR

main_bp = Blueprint('main', __name__)


def send_file_with_korean_filename(filepath, filename, mimetype=None):
    """한글 파일명을 지원하는 파일 다운로드 응답 생성 (RFC 5987)"""
    if mimetype:
        response = make_response(send_file(filepath, mimetype=mimetype))
    else:
        response = make_response(send_file(filepath))
    
    # RFC 5987 방식으로 한글 파일명 인코딩
    encoded_filename = quote(filename, safe='')
    
    # Content-Disposition 헤더 설정 (ASCII fallback + UTF-8 filename)
    ascii_filename = filename.encode('ascii', 'ignore').decode('ascii') or 'download'
    response.headers['Content-Disposition'] = (
        f"attachment; filename=\"{ascii_filename}\"; "
        f"filename*=UTF-8''{encoded_filename}"
    )
    
    return response


@main_bp.route('/')
def index():
    """메인 대시보드 페이지"""
    return render_template('dashboard.html', reports=REPORT_ORDER, regional_reports=REGIONAL_REPORTS)


@main_bp.route('/preview/<report_id>')
def preview_report(report_id):
    """보도자료 미리보기"""
    # report_id로부터 보도자료 이름 찾기
    report = next((r for r in REPORT_ORDER if r['id'] == report_id), None)
    if not report:
        return f"보도자료를 찾을 수 없습니다: {report_id}", 404
    
    report_name = report['name']
    
    # 가능한 파일명 패턴들
    possible_files = [
        TEMPLATES_DIR / f"{report_name}_preview.html",
        TEMPLATES_DIR / f"{report_name}_output.html",
    ]
    
    for file_path in possible_files:
        if file_path.exists():
            return send_file(str(file_path), mimetype='text/html')
    
    return f"보도자료가 아직 생성되지 않았습니다: {report_name}", 404


@main_bp.route('/download/<report_id>')
def download_report(report_id):
    """보도자료 다운로드 (안전한 처리)"""
    # report_id 검증
    if not report_id or not isinstance(report_id, str):
        return "유효하지 않은 보도자료 ID입니다", 400
    
    # report_id로부터 보도자료 이름 찾기 (안전한 검색)
    report = None
    is_regional = False
    
    try:
        # 먼저 REPORT_ORDER에서 찾기
        for r in REPORT_ORDER:
            if r and isinstance(r, dict) and r.get('id') == report_id:
                report = r
                break
        
        # REPORT_ORDER에서 못 찾으면 REGIONAL_REPORTS에서 찾기
        if not report:
            for r in REGIONAL_REPORTS:
                if r and isinstance(r, dict) and r.get('id') == report_id:
                    report = r
                    is_regional = True
                    break
    except Exception as e:
        print(f"[ERROR] 보도자료 검색 중 오류: {e}")
        return f"보도자료 검색 중 오류가 발생했습니다: {report_id}", 500
    
    if not report:
        return f"보도자료를 찾을 수 없습니다: {report_id}", 404
    
    report_name = report.get('name', '')
    if not report_name or not isinstance(report_name, str):
        return "보도자료 이름이 유효하지 않습니다", 500
    
    # 파일명 안전화 (위험한 문자 제거, 하이픈은 유지)
    report_name_safe = report_name.replace('/', '_').replace('\\', '_').replace('..', '_')
    # 하이픈은 파일명에 포함될 수 있으므로 유지 (예: '요약-고용인구')
    
    # 디버그: 파일 검색 정보 출력
    print(f"[다운로드] 보도자료 검색:")
    print(f"  - report_id: {report_id}")
    print(f"  - report_name: {report_name}")
    print(f"  - report_name_safe: {report_name_safe}")
    print(f"  - is_regional: {is_regional}")
    
    # 가능한 파일명 패턴들
    possible_files = []
    
    if is_regional:
        # 시도별 보고서: regional_output 폴더 확인
        possible_files = [
            TEMPLATES_DIR / 'regional_output' / f"{report_name_safe}_output.html",
            TEMPLATES_DIR / 'regional_output' / f"{report_name}_output.html",  # 원본 이름도 시도
            TEMPLATES_DIR / f"{report_name_safe}_output.html",
            TEMPLATES_DIR / f"{report_name_safe}_preview.html",
        ]
    else:
        # 일반 보고서: templates 폴더 직접 확인 (여러 패턴 시도)
        possible_files = [
            TEMPLATES_DIR / f"{report_name_safe}_output.html",
            TEMPLATES_DIR / f"{report_name}_output.html",  # 원본 이름도 시도
            TEMPLATES_DIR / f"{report_name_safe}_preview.html",
            TEMPLATES_DIR / f"{report_name}_preview.html",  # 원본 이름도 시도
        ]
    
    # 디버그: 검색할 파일 목록 출력
    print(f"[다운로드] 검색할 파일 목록:")
    for file_path in possible_files:
        exists = file_path.exists() if file_path else False
        is_file = file_path.is_file() if exists else False
        print(f"  - {file_path} (존재: {exists}, 파일: {is_file})")
    
    # 파일 찾기 및 다운로드 (안전한 처리)
    for file_path in possible_files:
        try:
            if file_path.exists() and file_path.is_file():
                filename = f"{report_name}.html"
                print(f"[다운로드] ✅ 파일 발견: {file_path}")
                print(f"[다운로드] 다운로드 파일명: {filename}")
                return send_file_with_korean_filename(str(file_path), filename, 'text/html')
        except Exception as e:
            print(f"[ERROR] 파일 다운로드 중 오류 ({file_path}): {e}")
            import traceback
            traceback.print_exc()
            continue  # 다음 파일 시도
    
    # 디버그: 파일을 찾지 못한 경우 실제 존재하는 파일 목록 출력
    print(f"[다운로드] ❌ 파일을 찾을 수 없습니다. TEMPLATES_DIR의 파일 목록:")
    try:
        if is_regional:
            regional_dir = TEMPLATES_DIR / 'regional_output'
            if regional_dir.exists():
                files = list(regional_dir.glob('*.html'))
                print(f"  - regional_output 폴더의 HTML 파일: {[f.name for f in files[:10]]}")
        else:
            files = list(TEMPLATES_DIR.glob('*_output.html'))
            print(f"  - templates 폴더의 *_output.html 파일: {[f.name for f in files[:10]]}")
    except Exception as e:
        print(f"  - 파일 목록 조회 중 오류: {e}")
    
    return f"보도자료가 아직 생성되지 않았습니다: {report_name}", 404


@main_bp.route('/preview/infographic')
def preview_infographic():
    """인포그래픽 미리보기 (직접 접근용)"""
    output_path = TEMPLATES_DIR / 'infographic_output.html'
    if output_path.exists():
        return send_file(output_path)
    return "인포그래픽이 아직 생성되지 않았습니다.", 404


@main_bp.route('/uploads/<filename>')
def download_file(filename):
    """업로드된 파일 다운로드 (한글 파일명 지원)"""
    filepath = UPLOAD_FOLDER / filename
    if filepath.exists():
        return send_file_with_korean_filename(str(filepath), filename)
    return "파일을 찾을 수 없습니다.", 404


@main_bp.route('/view/<filename>')
def view_file(filename):
    """파일 직접 보기 (다운로드 없이) - uploads와 debug 폴더 모두 확인"""
    # uploads 폴더 먼저 확인
    filepath = UPLOAD_FOLDER / filename
    if filepath.exists():
        if filename.endswith('.html'):
            return send_file(str(filepath), mimetype='text/html')
        return send_file(str(filepath))
    
    # debug 폴더 확인
    debug_filepath = DEBUG_FOLDER / filename
    if debug_filepath.exists():
        if filename.endswith('.html'):
            return send_file(str(debug_filepath), mimetype='text/html')
        return send_file(str(debug_filepath))
    
    return "파일을 찾을 수 없습니다.", 404


@main_bp.route('/debug/<filename>')
def view_debug_file(filename):
    """디버그 파일 직접 보기"""
    filepath = DEBUG_FOLDER / filename
    if filepath.exists():
        if filename.endswith('.html'):
            return send_file(str(filepath), mimetype='text/html')
        return send_file(str(filepath))
    return "디버그 파일을 찾을 수 없습니다.", 404


@main_bp.route('/exports/<path:filepath>')
def view_export_file(filepath):
    """내보내기 파일 직접 보기 (한글 불러오기용 HTML)"""
    file_path = EXPORT_FOLDER / filepath
    if file_path.exists() and file_path.is_file():
        if filepath.endswith('.html'):
            return send_file(str(file_path), mimetype='text/html')
        elif filepath.endswith('.png') or filepath.endswith('.jpg') or filepath.endswith('.jpeg'):
            return send_file(str(file_path))
        return send_file(str(file_path))
    return "파일을 찾을 수 없습니다.", 404


@main_bp.route('/templates/<filename>')
def serve_template_file(filename):
    """templates 폴더의 정적 파일 제공 (이미지 등)"""
    filepath = TEMPLATES_DIR / filename
    if filepath.exists() and filepath.is_file():
        # 이미지 파일
        if filename.endswith('.png'):
            return send_file(str(filepath), mimetype='image/png')
        elif filename.endswith('.jpg') or filename.endswith('.jpeg'):
            return send_file(str(filepath), mimetype='image/jpeg')
        elif filename.endswith('.svg'):
            return send_file(str(filepath), mimetype='image/svg+xml')
        elif filename.endswith('.css'):
            return send_file(str(filepath), mimetype='text/css')
        elif filename.endswith('.js'):
            return send_file(str(filepath), mimetype='application/javascript')
        return send_file(str(filepath))
    return "파일을 찾을 수 없습니다.", 404


@main_bp.route('/download-export/<export_dir>')
def download_export_zip(export_dir):
    """내보내기 폴더를 ZIP으로 다운로드"""
    import zipfile
    import tempfile
    import shutil
    
    export_path = EXPORT_FOLDER / export_dir
    if not export_path.exists() or not export_path.is_dir():
        return "내보내기 폴더를 찾을 수 없습니다.", 404
    
    try:
        # 임시 ZIP 파일 생성
        temp_zip = tempfile.NamedTemporaryFile(delete=False, suffix='.zip')
        temp_zip.close()
        
        with zipfile.ZipFile(temp_zip.name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in export_path.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(export_path)
                    zipf.write(str(file_path), arcname=str(arcname))
        
        return send_file_with_korean_filename(
            temp_zip.name,
            f'{export_dir}.zip',
            'application/zip'
        )
    except Exception as e:
        return f"ZIP 생성 오류: {str(e)}", 500

