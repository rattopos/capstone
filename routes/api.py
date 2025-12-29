# -*- coding: utf-8 -*-
"""
API 라우트
"""

import json
import base64
import re
from pathlib import Path

from flask import Blueprint, request, jsonify, session, send_file
from werkzeug.utils import secure_filename

from config.settings import BASE_DIR, TEMPLATES_DIR, UPLOAD_FOLDER
from config.reports import REPORT_ORDER, REGIONAL_REPORTS, SUMMARY_REPORTS, STATISTICS_REPORTS
from utils.excel_utils import extract_year_quarter_from_excel, extract_year_quarter_from_raw, detect_file_type
from services.report_generator import (
    generate_report_html,
    generate_regional_report_html,
    generate_statistics_report_html,
    generate_individual_statistics_html
)
from services.grdp_service import get_kosis_grdp_download_info, parse_kosis_grdp_file
from data_converter import DataConverter
import openpyxl

api_bp = Blueprint('api', __name__, url_prefix='/api')


def _calculate_analysis_sheets(excel_path: str):
    """분석 시트의 수식을 계산하여 값으로 채움 (집계 시트 참조)
    
    분석 시트의 수식은 집계 시트를 참조하므로, 집계 시트 값을 복사합니다.
    예: ='A(광공업생산)집계'!A4 → A(광공업생산)집계 시트의 A4 값 복사
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
        
        # 분석 시트의 수식 셀을 값으로 교체
        for row in ws_analysis.iter_rows(min_row=1, max_row=ws_analysis.max_row):
            for cell in row:
                if cell.value is None:
                    continue
                    
                val = str(cell.value)
                
                # 수식인 경우 (=로 시작)
                if val.startswith('='):
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
                            cell.value = ref_value
                    else:
                        # 다른 복잡한 수식은 0으로 처리 (나중에 확장 가능)
                        # 증감률 계산 수식 등은 별도 처리 필요
                        pass
    
    wb.save(excel_path)
    wb.close()
    print(f"[분석표] 분석 시트 수식 계산 완료: {excel_path}")


@api_bp.route('/upload', methods=['POST'])
def upload_excel():
    """엑셀 파일 업로드 (기초자료 수집표만 지원)"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '파일이 없습니다'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '파일이 선택되지 않았습니다'})
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'error': '엑셀 파일만 업로드 가능합니다'})
    
    filename = secure_filename(file.filename)
    filepath = Path(UPLOAD_FOLDER) / filename
    file.save(str(filepath))
    
    # 파일 유형 자동 감지
    file_type = detect_file_type(str(filepath))
    
    # 기초자료 수집표만 허용 (분석표는 더 이상 지원하지 않음)
    if file_type == 'analysis':
        filepath.unlink()  # 업로드된 파일 삭제
        return jsonify({
            'success': False, 
            'error': '분석표는 더 이상 지원하지 않습니다. 기초자료 수집표를 업로드해주세요.'
        })
    
    # 기초자료 수집표 처리 (분석표는 다운로드 시점에 생성)
    grdp_data = None
    conversion_info = None
    
    print(f"[업로드] 기초자료 수집표 감지")
    try:
        converter = DataConverter(str(filepath))
        
        # GRDP 데이터만 추출 (분석표는 다운로드 시 생성)
        grdp_data = converter.extract_grdp_data()
        
        grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
        with open(grdp_json_path, 'w', encoding='utf-8') as f:
            json.dump(grdp_data, f, ensure_ascii=False, indent=2)
        
        conversion_info = {
            'original_file': filename,
            'grdp_extracted': True,
            'national_growth_rate': grdp_data['national_summary']['growth_rate'],
            'top_region': grdp_data['top_region']['name'],
            'top_region_growth': grdp_data['top_region']['growth_rate']
        }
        
        print(f"[업로드] GRDP 추출 - 전국: {grdp_data['national_summary']['growth_rate']}%, 1위: {grdp_data['top_region']['name']}")
        
    except Exception as e:
        import traceback
        print(f"[오류] 기초자료 처리 실패: {e}")
        traceback.print_exc()
        return jsonify({
            'success': False,
            'error': f'기초자료 처리 중 오류가 발생했습니다: {str(e)}'
        })
    
    # 연도/분기 추출 (기초자료 파일에서)
    year, quarter = extract_year_quarter_from_raw(str(filepath))
    
    # 세션에 저장 (분석표는 다운로드 시 생성, 보고서는 기초자료에서 직접 추출)
    session['raw_excel_path'] = str(filepath)
    session['excel_path'] = str(filepath)  # 보고서 생성용 (기초자료 직접 사용)
    session['year'] = year
    session['quarter'] = quarter
    session['file_type'] = 'raw'
    
    if grdp_data:
        session['grdp_data'] = grdp_data
    
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
    """KOSIS GRDP 파일 업로드 및 파싱"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '파일이 없습니다.'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '파일이 선택되지 않았습니다.'}), 400
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'error': '엑셀 파일만 업로드 가능합니다.'}), 400
    
    filename = secure_filename(file.filename)
    if 'grdp' not in filename.lower() and 'GRDP' not in filename:
        filename = f"grdp_{filename}"
    
    filepath = UPLOAD_FOLDER / filename
    file.save(str(filepath))
    
    year = session.get('year', 2025)
    quarter = session.get('quarter', 2)
    
    grdp_data = parse_kosis_grdp_file(str(filepath), year, quarter)
    
    if grdp_data:
        session['grdp_data'] = grdp_data
        grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
        with open(grdp_json_path, 'w', encoding='utf-8') as f:
            json.dump(grdp_data, f, ensure_ascii=False, indent=2)
        
        return jsonify({
            'success': True,
            'message': 'GRDP 데이터가 성공적으로 업로드되었습니다.',
            'national_growth_rate': grdp_data.get('national_summary', {}).get('growth_rate', 0),
            'top_region': grdp_data.get('top_region', {}).get('name', '-')
        })
    else:
        return jsonify({
            'success': False,
            'error': 'GRDP 데이터를 파싱할 수 없습니다. 올바른 KOSIS GRDP 파일인지 확인하세요.'
        }), 400


@api_bp.route('/download-analysis', methods=['GET'])
def download_analysis():
    """분석표 다운로드 (다운로드 시점에 생성 + 수식 계산)"""
    raw_excel_path = session.get('raw_excel_path')
    
    if not raw_excel_path or not Path(raw_excel_path).exists():
        return jsonify({'success': False, 'error': '기초자료 파일을 찾을 수 없습니다. 먼저 기초자료를 업로드해주세요.'}), 404
    
    try:
        converter = DataConverter(str(raw_excel_path))
        
        # 분석표 생성
        analysis_output = str(UPLOAD_FOLDER / f"분석표_{converter.year}년_{converter.quarter}분기_자동생성.xlsx")
        analysis_path = converter.convert_all(analysis_output, weight_settings=None)
        
        # 분석 시트 수식 계산 (집계 시트 값을 분석 시트로 복사)
        _calculate_analysis_sheets(analysis_path)
        
        filename = Path(analysis_path).name
        
        return send_file(
            analysis_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': f'분석표 생성 실패: {str(e)}'}), 500


@api_bp.route('/generate-analysis-with-weights', methods=['POST'])
def generate_analysis_with_weights():
    """가중치 설정을 포함하여 분석표 생성 + 다운로드"""
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
        
        # 분석 시트 수식 계산 (집계 시트 값을 분석 시트로 복사)
        _calculate_analysis_sheets(analysis_path)
        
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


@api_bp.route('/export-hwp-ready', methods=['POST'])
def export_hwp_ready():
    """한글(HWP) 복붙용 HTML 문서 생성"""
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
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700&display=swap');
        
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: '맑은 고딕', 'Malgun Gothic', 'Noto Sans KR', sans-serif;
            font-size: 10pt;
            line-height: 1.6;
            color: #000;
            background: #fff;
            padding: 20px;
            max-width: 210mm;
            margin: 0 auto;
        }}
        
        /* 페이지 구분선 */
        .hwp-page {{
            margin-bottom: 30px;
            padding-bottom: 30px;
            border-bottom: 3px double #333;
            page-break-after: always;
        }}
        
        .hwp-page:last-child {{
            border-bottom: none;
            page-break-after: auto;
        }}
        
        /* 페이지 제목 */
        .hwp-page-title {{
            font-size: 14pt;
            font-weight: bold;
            color: #1a1a1a;
            margin-bottom: 15px;
            padding: 8px 12px;
            background: #f0f0f0;
            border-left: 4px solid #0066cc;
        }}
        
        /* 페이지 번호 */
        .hwp-page-number {{
            text-align: center;
            font-size: 9pt;
            color: #666;
            margin-top: 20px;
        }}
        
        /* 표 스타일 - 한글에서 잘 인식됨 */
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 10px 0;
            font-size: 9pt;
        }}
        
        th, td {{
            border: 1px solid #000;
            padding: 5px 8px;
            text-align: center;
            vertical-align: middle;
        }}
        
        th {{
            background-color: #e6e6e6;
            font-weight: bold;
        }}
        
        /* 제목 스타일 */
        h1, h2, h3, h4, h5, h6 {{
            margin: 15px 0 10px 0;
            color: #1a1a1a;
        }}
        
        h1 {{ font-size: 16pt; }}
        h2 {{ font-size: 14pt; }}
        h3 {{ font-size: 12pt; }}
        h4 {{ font-size: 11pt; }}
        
        /* 목록 */
        ul, ol {{
            margin: 10px 0 10px 25px;
        }}
        
        li {{
            margin: 5px 0;
        }}
        
        /* 강조 */
        strong, b {{
            font-weight: bold;
        }}
        
        /* 차트 대체 안내 */
        .chart-placeholder {{
            border: 2px dashed #999;
            padding: 20px;
            text-align: center;
            background: #fafafa;
            color: #666;
            margin: 15px 0;
        }}
        
        /* 숨김 처리 (복붙 시 불필요한 요소) */
        script, canvas, .chart-container canvas {{
            display: none !important;
        }}
        
        /* 복사 안내 */
        .copy-guide {{
            position: fixed;
            top: 10px;
            right: 10px;
            background: #0066cc;
            color: white;
            padding: 10px 15px;
            border-radius: 5px;
            font-size: 11pt;
            z-index: 9999;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
        }}
        
        .copy-guide:hover {{
            background: #0055aa;
        }}
        
        @media print {{
            .copy-guide {{
                display: none;
            }}
        }}
    </style>
</head>
<body>
    <div class="copy-guide" onclick="selectAll()">📋 전체 선택 (Ctrl+A) 후 복사 (Ctrl+C)</div>
    
    <div style="text-align: center; margin-bottom: 30px;">
        <h1 style="font-size: 18pt; margin-bottom: 5px;">{year}년 {quarter}/4분기 지역경제동향</h1>
        <p style="color: #666;">한글 복붙용 문서 - 복사 후 한글에 붙여넣기 하세요</p>
    </div>
'''
        
        for idx, page in enumerate(pages, 1):
            page_html = page.get('html', '')
            page_title = page.get('title', f'페이지 {idx}')
            category = page.get('category', '')
            
            # body 내용 추출
            body_content = page_html
            if '<body' in page_html.lower():
                import re
                body_match = re.search(r'<body[^>]*>(.*?)</body>', page_html, re.DOTALL | re.IGNORECASE)
                if body_match:
                    body_content = body_match.group(1)
            
            # 한글 복붙에 불필요한 요소 제거/변환
            import re
            
            # style 태그 제거
            body_content = re.sub(r'<style[^>]*>.*?</style>', '', body_content, flags=re.DOTALL)
            
            # script 태그 제거
            body_content = re.sub(r'<script[^>]*>.*?</script>', '', body_content, flags=re.DOTALL)
            
            # canvas를 차트 플레이스홀더로 대체
            body_content = re.sub(
                r'<canvas[^>]*>.*?</canvas>',
                '<div class="chart-placeholder">📊 [차트 영역 - 별도 이미지 삽입 필요]</div>',
                body_content,
                flags=re.DOTALL
            )
            
            # 빈 canvas 태그도 처리
            body_content = re.sub(
                r'<canvas[^>]*/?>',
                '<div class="chart-placeholder">📊 [차트 영역 - 별도 이미지 삽입 필요]</div>',
                body_content
            )
            
            # class 속성은 유지 (일부 스타일 적용 위해)
            # inline style은 유지
            
            # 카테고리 한글명
            category_names = {{
                'summary': '요약',
                'sectoral': '부문별',
                'regional': '시도별',
                'statistics': '통계표'
            }}
            category_name = category_names.get(category, '')
            
            final_html += f'''
    <div class="hwp-page" data-page="{idx}">
        <div class="hwp-page-title">[{category_name}] {page_title}</div>
        <div class="hwp-page-content">
{body_content}
        </div>
        <div class="hwp-page-number">- {idx} / {len(pages)} -</div>
    </div>
'''
        
        final_html += '''
    <script>
        function selectAll() {
            // 가이드 요소 제외하고 선택
            const guide = document.querySelector('.copy-guide');
            guide.style.display = 'none';
            
            const range = document.createRange();
            range.selectNodeContents(document.body);
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
            
            // 복사 후 가이드 다시 표시
            setTimeout(() => {
                guide.style.display = 'block';
            }, 100);
            
            alert('전체 선택되었습니다.\\nCtrl+C로 복사 후 한글에서 Ctrl+V로 붙여넣기 하세요.');
        }
        
        // Ctrl+A 시 전체 선택 함수 호출
        document.addEventListener('keydown', function(e) {
            if (e.ctrlKey && e.key === 'a') {
                e.preventDefault();
                selectAll();
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
            'view_url': f'/uploads/{output_filename}',
            'total_pages': len(pages)
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})

