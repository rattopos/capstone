#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
지역경제동향 보고서 웹 애플리케이션
Flask 기반 대시보드로 분석표 엑셀을 업로드하고 보고서를 생성합니다.
"""

import os
import sys
import json
import importlib.util
from pathlib import Path
from flask import Flask, render_template, request, jsonify, session
from werkzeug.utils import secure_filename
import pandas as pd
from jinja2 import Template

# 프로젝트 루트 설정
BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / 'templates'
UPLOAD_FOLDER = BASE_DIR / 'uploads'

# 업로드 폴더 생성
UPLOAD_FOLDER.mkdir(exist_ok=True)

app = Flask(__name__, 
            template_folder=str(BASE_DIR),
            static_folder=str(BASE_DIR))
app.secret_key = 'capstone_secret_key_2025'
app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# 보고서 순서 설정 (유연하게 변경 가능)
REPORT_ORDER = [
    {
        'id': 'manufacturing',
        'name': '광공업생산',
        'sheet': 'A 분석',
        'generator': '광공업생산_generator.py',
        'template': '광공업생산_template.html',
        'icon': '🏭',
        'category': 'production'
    },
    {
        'id': 'service',
        'name': '서비스업생산',
        'sheet': 'B 분석',
        'generator': '서비스업생산_generator.py',
        'template': '서비스업생산_template.html',
        'icon': '🏢',
        'category': 'production'
    },
    {
        'id': 'consumption',
        'name': '소비동향',
        'sheet': 'C 분석',
        'generator': '소비동향_generator.py',
        'template': '소비동향_template.html',
        'icon': '🛒',
        'category': 'consumption'
    },
    {
        'id': 'employment',
        'name': '고용률',
        'sheet': 'D(고용률)분석',
        'generator': '고용률_generator.py',
        'template': '고용률_template.html',
        'icon': '👔',
        'category': 'employment'
    },
    {
        'id': 'unemployment',
        'name': '실업률',
        'sheet': 'D(실업)분석',
        'generator': '실업률_generator.py',
        'template': '실업률_template.html',
        'icon': '📉',
        'category': 'employment'
    },
    {
        'id': 'price',
        'name': '물가동향',
        'sheet': 'E(품목성질물가)분석',
        'generator': '물가동향_generator.py',
        'template': '물가동향_template.html',
        'icon': '💰',
        'category': 'price'
    },
    {
        'id': 'export',
        'name': '수출',
        'sheet': 'G 분석',
        'generator': '수출_generator.py',
        'template': '수출_template.html',
        'icon': '📦',
        'category': 'trade'
    },
    {
        'id': 'import',
        'name': '수입',
        'sheet': 'H 분석',
        'generator': '수입_generator.py',
        'template': '수입_template.html',
        'icon': '🚢',
        'category': 'trade'
    },
    {
        'id': 'population',
        'name': '국내인구이동',
        'sheet': 'I(순인구이동)집계',
        'generator': '국내인구이동_generator.py',
        'template': '국내인구이동_template.html',
        'icon': '👥',
        'category': 'population'
    }
]


def load_generator_module(generator_name):
    """동적으로 generator 모듈 로드"""
    generator_path = TEMPLATES_DIR / generator_name
    if not generator_path.exists():
        return None
    
    spec = importlib.util.spec_from_file_location(
        generator_name.replace('.py', ''),
        str(generator_path)
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def extract_year_quarter_from_excel(filepath):
    """엑셀 파일에서 최신 연도와 분기 추출"""
    try:
        xl = pd.ExcelFile(filepath)
        # A 분석 시트에서 연도/분기 정보 추출 시도
        df = pd.read_excel(xl, sheet_name='A 분석', header=None)
        
        # 일반적으로 컬럼 헤더에서 연도/분기 정보를 찾음
        # 예: "2025.2/4" 형태
        for row_idx in range(min(5, len(df))):
            for col_idx in range(len(df.columns)):
                cell = str(df.iloc[row_idx, col_idx])
                if '2025.2/4' in cell or '25.2/4' in cell:
                    return 2025, 2
                elif '2025.1/4' in cell or '25.1/4' in cell:
                    return 2025, 1
                elif '2024.4/4' in cell or '24.4/4' in cell:
                    return 2024, 4
        
        # 파일명에서 추출 시도
        filename = Path(filepath).stem
        if '25년' in filename and '2분기' in filename:
            return 2025, 2
        elif '25년' in filename and '1분기' in filename:
            return 2025, 1
        
        return 2025, 2  # 기본값
    except Exception as e:
        print(f"연도/분기 추출 오류: {e}")
        return 2025, 2


def check_missing_data(data, report_id):
    """보고서 생성에 필수적인 결측치만 확인"""
    missing_fields = []
    
    # 보고서별 필수 필드 정의
    REQUIRED_FIELDS = {
        'manufacturing': [
            'nationwide_data.production_index',
            'nationwide_data.growth_rate',
            'summary_box.region_count',
        ],
        'service': [
            'nationwide_data.production_index',
            'nationwide_data.growth_rate',
            'summary_box.region_count',
        ],
        'consumption': [
            'nationwide_data.index_value',
            'nationwide_data.growth_rate',
            'summary_box.region_count',
        ],
        'employment': [
            'nationwide_data.employment_rate',
            'nationwide_data.change',
            'summary_box.region_count',
        ],
        'unemployment': [
            'nationwide_data.unemployment_rate',
            'nationwide_data.change',
            'summary_box.region_count',
        ],
        'price': [
            'nationwide_data.price_index',
            'nationwide_data.change_rate',
            'summary_box.region_count',
        ],
        'export': [
            'nationwide_data.export_value',
            'nationwide_data.growth_rate',
            'summary_box.region_count',
        ],
        'import': [
            'nationwide_data.import_value',
            'nationwide_data.growth_rate',
            'summary_box.region_count',
        ],
        'population': [
            'nationwide_data.net_migration',
            'summary_box.region_count',
        ],
    }
    
    def get_nested_value(obj, path):
        """중첩된 경로에서 값 가져오기"""
        keys = path.replace('[', '.').replace(']', '').split('.')
        current = obj
        for key in keys:
            if current is None:
                return None
            if isinstance(current, dict):
                current = current.get(key)
            elif isinstance(current, list) and key.isdigit():
                idx = int(key)
                current = current[idx] if idx < len(current) else None
            else:
                return None
        return current
    
    def is_missing(value):
        """값이 결측치인지 확인"""
        if value is None:
            return True
        if value == '':
            return True
        if isinstance(value, float) and pd.isna(value):
            return True
        return False
    
    # 해당 보고서의 필수 필드만 확인
    required = REQUIRED_FIELDS.get(report_id, [])
    for field_path in required:
        value = get_nested_value(data, field_path)
        if is_missing(value):
            missing_fields.append(field_path)
    
    return missing_fields


def generate_report_html(excel_path, report_config, year, quarter, custom_data=None):
    """보고서 HTML 생성"""
    try:
        generator_name = report_config['generator']
        template_name = report_config['template']
        
        # Generator 모듈 로드
        module = load_generator_module(generator_name)
        if not module:
            return None, f"Generator 모듈을 찾을 수 없습니다: {generator_name}", []
        
        # Generator 클래스 찾기 및 실행
        generator_class = None
        for name in dir(module):
            obj = getattr(module, name)
            if isinstance(obj, type) and name.endswith('Generator'):
                generator_class = obj
                break
        
        # generate_report 함수 사용 (고용률 등)
        if hasattr(module, 'generate_report'):
            template_path = TEMPLATES_DIR / template_name
            output_path = TEMPLATES_DIR / f"{report_config['name']}_preview.html"
            
            # 데이터 추출
            if hasattr(module, 'load_data'):
                df_analysis, df_index = module.load_data(excel_path)
                data = {}
                
                if hasattr(module, 'get_nationwide_data'):
                    data['nationwide_data'] = module.get_nationwide_data(df_analysis, df_index)
                if hasattr(module, 'get_regional_data'):
                    data['regional_data'] = module.get_regional_data(df_analysis, df_index)
                if hasattr(module, 'get_summary_box_data'):
                    data['summary_box'] = module.get_summary_box_data(data.get('regional_data', {}))
                if hasattr(module, 'get_table_data'):
                    data['summary_table'] = {
                        'columns': {
                            'change_columns': [f'{year-2}.{quarter}/4', f'{year-1}.{quarter}/4', f'{year}.{quarter-1}/4' if quarter > 1 else f'{year-1}.4/4', f'{year}.{quarter}/4'],
                            'rate_columns': [f'{year-1}.{quarter}/4', f'{year}.{quarter}/4', '20-29세']
                        },
                        'regions': module.get_table_data(df_analysis, df_index)
                    }
                
                # Top3 regions - 양쪽 키 이름 모두 제공 (템플릿 호환성)
                if 'regional_data' in data:
                    top3_increase = []
                    for r in data['regional_data'].get('increase_regions', [])[:3]:
                        rate_value = r.get('change', r.get('growth_rate', 0))
                        items = r.get('top_age_groups', r.get('industries', r.get('top_industries', [])))
                        top3_increase.append({
                            'region': r.get('region', ''),
                            'change': rate_value,
                            'growth_rate': rate_value,  # 템플릿 호환
                            'age_groups': items,
                            'industries': items  # 템플릿 호환
                        })
                    
                    top3_decrease = []
                    for r in data['regional_data'].get('decrease_regions', [])[:3]:
                        rate_value = r.get('change', r.get('growth_rate', 0))
                        items = r.get('top_age_groups', r.get('industries', r.get('top_industries', [])))
                        top3_decrease.append({
                            'region': r.get('region', ''),
                            'change': rate_value,
                            'growth_rate': rate_value,  # 템플릿 호환
                            'age_groups': items,
                            'industries': items  # 템플릿 호환
                        })
                    
                    data['top3_increase_regions'] = top3_increase
                    data['top3_decrease_regions'] = top3_decrease
            else:
                data = module.generate_report(excel_path, template_path, output_path)
        elif generator_class:
            generator = generator_class(excel_path)
            data = generator.extract_all_data()
        else:
            return None, f"유효한 Generator를 찾을 수 없습니다: {generator_name}", []
        
        # 커스텀 데이터 병합 (사용자가 입력한 결측치)
        if custom_data:
            for key, value in custom_data.items():
                keys = key.split('.')
                obj = data
                for k in keys[:-1]:
                    if '[' in k:
                        name, idx = k.replace(']', '').split('[')
                        obj = obj[name][int(idx)]
                    else:
                        obj = obj[k]
                final_key = keys[-1]
                if '[' in final_key:
                    name, idx = final_key.replace(']', '').split('[')
                    obj[name][int(idx)] = value
                else:
                    obj[final_key] = value
        
        # 결측치 확인
        missing = check_missing_data(data, report_config['id'])
        
        # 템플릿 렌더링
        template_path = TEMPLATES_DIR / template_name
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        html_content = template.render(**data)
        
        return html_content, None, missing
        
    except Exception as e:
        import traceback
        error_msg = f"보고서 생성 오류: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        return None, error_msg, []


@app.route('/')
def index():
    """메인 대시보드 페이지"""
    return render_template('dashboard.html', reports=REPORT_ORDER)


@app.route('/api/upload', methods=['POST'])
def upload_excel():
    """엑셀 파일 업로드"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '파일이 없습니다'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '파일이 선택되지 않았습니다'})
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'error': '엑셀 파일만 업로드 가능합니다'})
    
    filename = secure_filename(file.filename)
    filepath = Path(app.config['UPLOAD_FOLDER']) / filename
    file.save(str(filepath))
    
    # 연도/분기 추출
    year, quarter = extract_year_quarter_from_excel(str(filepath))
    
    # 세션에 파일 경로 저장
    session['excel_path'] = str(filepath)
    session['year'] = year
    session['quarter'] = quarter
    
    return jsonify({
        'success': True,
        'filename': filename,
        'year': year,
        'quarter': quarter,
        'reports': REPORT_ORDER
    })


@app.route('/api/generate-preview', methods=['POST'])
def generate_preview():
    """미리보기 생성"""
    data = request.get_json()
    report_id = data.get('report_id')
    year = data.get('year', session.get('year', 2025))
    quarter = data.get('quarter', session.get('quarter', 2))
    custom_data = data.get('custom_data', {})
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    # 보고서 설정 찾기
    report_config = next((r for r in REPORT_ORDER if r['id'] == report_id), None)
    if not report_config:
        return jsonify({'success': False, 'error': f'보고서를 찾을 수 없습니다: {report_id}'})
    
    # HTML 생성
    html_content, error, missing_fields = generate_report_html(
        excel_path, report_config, year, quarter, custom_data
    )
    
    if error:
        return jsonify({'success': False, 'error': error})
    
    return jsonify({
        'success': True,
        'html': html_content,
        'missing_fields': missing_fields,
        'report_id': report_id,
        'report_name': report_config['name']
    })


@app.route('/api/generate-all', methods=['POST'])
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
        html_content, error, _ = generate_report_html(
            excel_path, report_config, year, quarter, custom_data
        )
        
        if error:
            errors.append({'report_id': report_config['id'], 'error': error})
        else:
            # 파일 저장
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


@app.route('/api/report-order', methods=['GET'])
def get_report_order():
    """현재 보고서 순서 반환"""
    return jsonify({'reports': REPORT_ORDER})


@app.route('/api/report-order', methods=['POST'])
def update_report_order():
    """보고서 순서 업데이트"""
    global REPORT_ORDER
    data = request.get_json()
    new_order = data.get('order', [])
    
    if new_order:
        # 새 순서로 재정렬
        order_map = {r['id']: idx for idx, r in enumerate(new_order)}
        REPORT_ORDER = sorted(REPORT_ORDER, key=lambda x: order_map.get(x['id'], 999))
    
    return jsonify({'success': True, 'reports': REPORT_ORDER})


@app.route('/api/session-info', methods=['GET'])
def get_session_info():
    """현재 세션 정보 반환"""
    return jsonify({
        'excel_path': session.get('excel_path'),
        'year': session.get('year'),
        'quarter': session.get('quarter'),
        'has_file': bool(session.get('excel_path'))
    })


if __name__ == '__main__':
    print("=" * 50)
    print("지역경제동향 보고서 생성 시스템")
    print("=" * 50)
    print(f"서버 시작: http://localhost:5050")
    print("=" * 50)
    app.run(debug=True, host='0.0.0.0', port=5050)

