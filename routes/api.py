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

api_bp = Blueprint('api', __name__, url_prefix='/api')


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
    
    # 기초자료 수집표 처리
    analysis_path = None
    grdp_data = None
    conversion_info = None
    
    print(f"[업로드] 기초자료 수집표 감지")
    try:
        converter = DataConverter(str(filepath))
        
        # 분석표 자동 생성
        analysis_output = str(UPLOAD_FOLDER / f"분석표_{converter.year}년_{converter.quarter}분기_자동생성.xlsx")
        analysis_path = converter.convert_all(analysis_output)
        print(f"[업로드] 분석표 자동 생성: {Path(analysis_path).name}")
        
        grdp_data = converter.extract_grdp_data()
        
        grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
        with open(grdp_json_path, 'w', encoding='utf-8') as f:
            json.dump(grdp_data, f, ensure_ascii=False, indent=2)
        
        session['download_analysis_path'] = analysis_path
        
        conversion_info = {
            'original_file': filename,
            'analysis_file': Path(analysis_path).name,
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
        # 분석표 생성 실패 시 기본 분석표 사용 시도
        original_analysis = BASE_DIR / '분석표_25년 2분기_캡스톤.xlsx'
        if original_analysis.exists():
            analysis_path = str(original_analysis)
        else:
            return jsonify({
                'success': False,
                'error': f'기초자료 처리 중 오류가 발생했습니다: {str(e)}'
            })
    
    # 연도/분기 추출 (기초자료 파일에서)
    year, quarter = extract_year_quarter_from_raw(str(filepath))
    
    # 세션에 저장
    session['excel_path'] = analysis_path
    session['raw_excel_path'] = str(filepath)
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
    """생성된 분석표 다운로드 (수식 유지 버전)"""
    excel_path = session.get('download_analysis_path') or session.get('excel_path')
    
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '분석표 파일을 찾을 수 없습니다.'}), 404
    
    filename = Path(excel_path).name
    
    return send_file(
        excel_path,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


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
    """모든 보고서를 하나의 HTML 문서로 합치기"""
    try:
        data = request.get_json()
        pages = data.get('pages', [])
        year = data.get('year', session.get('year', 2025))
        quarter = data.get('quarter', session.get('quarter', 2))
        
        if not pages:
            return jsonify({'success': False, 'error': '페이지 데이터가 없습니다.'})
        
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
        
        body {{
            font-family: 'Noto Sans KR', '맑은 고딕', sans-serif;
            background: white;
        }}
        
        .page {{
            width: 210mm;
            min-height: 297mm;
            padding: 15mm 20mm 25mm 20mm;
            margin: 0 auto;
            background: white;
            position: relative;
            page-break-after: always;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
        }}
        
        .page:last-child {{
            page-break-after: auto;
        }}
        
        .page-content {{
            width: 100%;
            min-height: calc(297mm - 40mm);
        }}
        
        .page-content > * {{
            max-width: 100%;
        }}
        
        .page-number {{
            position: absolute;
            bottom: 10mm;
            left: 0;
            right: 0;
            text-align: center;
            font-size: 10pt;
            color: #666;
        }}
        
        .page-content iframe {{
            border: none;
            width: 100%;
            min-height: 250mm;
        }}
        
        @media print {{
            body {{
                background: white;
            }}
            
            .page {{
                width: 210mm;
                min-height: 297mm;
                padding: 15mm 20mm 25mm 20mm;
                margin: 0;
                box-shadow: none;
                page-break-after: always;
            }}
            
            .page:last-child {{
                page-break-after: auto;
            }}
            
            .page-number {{
                position: absolute;
                bottom: 10mm;
            }}
        }}
        
        @page {{
            size: A4;
            margin: 0;
        }}
    </style>
</head>
<body>
'''
        
        for idx, page in enumerate(pages, 1):
            page_html = page.get('html', '')
            page_title = page.get('title', f'페이지 {idx}')
            
            body_content = page_html
            if '<body' in page_html:
                start = page_html.find('<body')
                start = page_html.find('>', start) + 1
                end = page_html.find('</body>')
                if end > start:
                    body_content = page_html[start:end]
            
            style_content = ''
            if '<style' in page_html:
                style_start = page_html.find('<style')
                style_end = page_html.find('</style>') + 8
                if style_end > style_start:
                    style_content = page_html[style_start:style_end]
            
            final_html += f'''
    <div class="page" data-page="{idx}">
        {style_content}
        <div class="page-content">
            {body_content}
        </div>
        <div class="page-number">{idx}</div>
    </div>
'''
        
        final_html += '''
</body>
</html>
'''
        
        output_filename = f'지역경제동향_{year}년_{quarter}분기.html'
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

