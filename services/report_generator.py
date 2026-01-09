# -*- coding: utf-8 -*-
"""
보도자료 생성 서비스
"""

import importlib.util
import json
import inspect
import pandas as pd
from pathlib import Path
from jinja2 import Template

import re

from config.settings import TEMPLATES_DIR, BASE_DIR, UPLOAD_FOLDER
from utils.filters import is_missing, format_value
from utils.excel_utils import load_generator_module
from utils.data_utils import check_missing_data
from .grdp_service import (
    get_kosis_grdp_download_info,
    parse_kosis_grdp_file
)


def _remove_page_numbers(html_content: str) -> str:
    """HTML에서 페이지 번호 제거
    
    모든 .page-number 클래스를 가진 div 요소를 제거합니다.
    
    Args:
        html_content: 원본 HTML 문자열
        
    Returns:
        페이지 번호가 제거된 HTML 문자열
    """
    if not html_content:
        return html_content
    
    # .page-number div 제거 (단일 라인 및 멀티라인 모두 처리)
    html_content = re.sub(
        r'<div\s+class=["\']page-number["\'][^>]*>.*?</div>',
        '',
        html_content,
        flags=re.DOTALL | re.IGNORECASE
    )
    
    return html_content


def _extract_data_from_raw(raw_excel_path, report_id, year, quarter):
    """기초자료에서 직접 보도자료 데이터 추출 (모듈화된 extractors 패키지 사용)
    
    Args:
        raw_excel_path: 기초자료 엑셀 파일 경로
        report_id: 보도자료 ID
        year: 연도
        quarter: 분기
        
    Returns:
        추출된 데이터 딕셔너리 또는 None
    """
    try:
        from extractors import DataExtractor
        
        extractor = DataExtractor(raw_excel_path, year, quarter)
        data = extractor.extract_report_data(report_id)
        
        if data:
            # 연도/분기 정보 추가
            if 'report_info' not in data:
                data['report_info'] = {}
            data['report_info']['year'] = year
            data['report_info']['quarter'] = quarter
            
            # 템플릿 호환성을 위한 기본값 추가
            data = _ensure_template_compatibility(data, report_id)
            
        return data
        
    except Exception as e:
        import traceback
        print(f"[WARNING] 기초자료 직접 추출 실패 ({report_id}): {e}")
        traceback.print_exc()
        return None


def _generate_default_summary_table(data, report_id):
    """부문별 보도자료용 기본 summary_table 생성
    
    각 템플릿이 기대하는 구조에 맞게 기본값 생성
    """
    report_info = data.get('report_info', {})
    year = report_info.get('year', 2025)
    quarter = report_info.get('quarter', 2)
    
    # 기본 컬럼 생성
    prev_quarters = [
        f"{year-1}.{quarter}/4",
        f"{year}.{quarter-3 if quarter > 3 else quarter + 1}/4",
        f"{year}.{quarter-2 if quarter > 2 else quarter + 2}/4",
        f"{year}.{quarter-1 if quarter > 1 else quarter + 3}/4",
    ]
    current_quarter = f"{year}.{quarter}/4"
    
    columns = {
        'growth_rate_columns': [prev_quarters[0], prev_quarters[1], prev_quarters[2], current_quarter],
        'index_columns': [prev_quarters[3], current_quarter],
        # 실업률 템플릿용
        'change_columns': [f"{year-2}.{quarter}/4", f"{year-1}.{quarter}/4", f"{year}.{quarter-1 if quarter > 1 else quarter+3}/4", current_quarter],
        'rate_columns': [f"{year-1}.{quarter}/4", current_quarter],
        # 인구이동 템플릿용
        'quarter_columns': [f"{year-2}.{quarter}/4", f"{year-1}.{quarter}/4", f"{year}.{quarter-1 if quarter > 1 else quarter+3}/4", current_quarter],
        'amount_columns': [f"{year-1}.{quarter}/4", current_quarter],
    }
    
    # 지역별 데이터 생성
    regional_data = data.get('regional_data', {})
    all_regions = regional_data.get('all', [])
    national_summary = data.get('national_summary', {})
    
    # 지역 그룹화
    REGION_GROUPS = {
        '수도권': ['서울', '인천', '경기'],
        '동남권': ['부산', '울산', '경남'],
        '대경권': ['대구', '경북'],
        '호남권': ['광주', '전북', '전남'],
        '충청권': ['대전', '세종', '충북', '충남'],
        '강원제주': ['강원', '제주'],
    }
    
    region_display = {
        '전국': '전 국', '서울': '서 울', '부산': '부 산', '대구': '대 구', '인천': '인 천',
        '광주': '광 주', '대전': '대 전', '울산': '울 산', '세종': '세 종', '경기': '경 기',
        '강원': '강 원', '충북': '충 북', '충남': '충 남', '전북': '전 북', '전남': '전 남',
        '경북': '경 북', '경남': '경 남', '제주': '제 주'
    }
    
    regions = []
    
    # 전국 행 추가 (템플릿 형식: growth_rates, indices 배열)
    national_growth_rate = national_summary.get('growth_rate', None)
    regions.append({
        'region': '전 국',
        'group': None,
        'growth_rates': [national_growth_rate, national_growth_rate, national_growth_rate, national_growth_rate],
        'indices': [None, None],  # 원지수는 데이터에서 추출해야 함
    })
    
    # 지역 그룹별 행 추가
    region_data_map = {r.get('region'): r for r in all_regions}
    
    for group_name, group_regions in REGION_GROUPS.items():
        for i, region_name in enumerate(group_regions):
            region = region_data_map.get(region_name, {})
            growth_rate = region.get('growth_rate', None)
            
            row = {
                'region': region_display.get(region_name, region_name),
                'growth_rates': [growth_rate, growth_rate, growth_rate, growth_rate],
                'indices': [None, None],  # 원지수는 데이터에서 추출해야 함
            }
            if i == 0:
                row['group'] = group_name
                row['rowspan'] = len(group_regions)
            else:
                row['group'] = None
            regions.append(row)
    
    return {
        'base_year': 2020,
        'columns': columns,
        'regions': regions,
        'rows': regions,  # 일부 템플릿은 rows를 사용
        'title': '주요 경제지표',
        'nationwide': {
            'growth_rate': national_summary.get('growth_rate', None),
            'change': national_summary.get('growth_rate', None),
        },
    }


def _ensure_template_compatibility(data, report_id):
    """기초자료에서 추출한 데이터에 템플릿 호환성을 위한 기본값 추가
    
    템플릿에서 사용하는 키가 없으면 기본값을 추가합니다.
    데이터가 없는 경우 N/A로 표시할 수 있도록 None 값을 유지합니다.
    """
    if not data:
        return data
    
    # summary_box 기본값 (부문별 템플릿용)
    if 'summary_box' not in data:
        regional_data = data.get('regional_data', {})
        increase_regions = regional_data.get('increase_regions', [])
        decrease_regions = regional_data.get('decrease_regions', [])
        
        data['summary_box'] = {
            'main_increase_regions': increase_regions[:3] if increase_regions else [],
            'main_decrease_regions': decrease_regions[:3] if decrease_regions else [],
            'region_count': len(increase_regions),
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'main_items': [],  # 물가동향용
            'headline': '',  # 인구이동용
            'inflow_summary': '',
            'outflow_summary': '',
        }
    
    # nationwide_data 기본값 - 공통 속성
    national_summary = data.get('national_summary', {})
    growth_rate = national_summary.get('growth_rate', None)
    
    if 'nationwide_data' not in data:
        data['nationwide_data'] = {}
    
    nwd = data['nationwide_data']
    
    # 공통 속성 (결측치는 None 유지)
    nwd.setdefault('production_index', None)
    nwd.setdefault('sales_index', None)
    nwd.setdefault('index', None)
    nwd.setdefault('growth_rate', growth_rate)
    nwd.setdefault('change', growth_rate)
    nwd.setdefault('employment_rate', None)
    nwd.setdefault('rate', None)
    nwd.setdefault('main_industries', [])
    nwd.setdefault('main_businesses', [])
    nwd.setdefault('main_age_groups', [])
    nwd.setdefault('categories', [])
    
    # 수출/수입용 속성
    nwd.setdefault('amount', None)
    nwd.setdefault('increase_products', [])
    nwd.setdefault('decrease_products', [])
    nwd.setdefault('products', [])  # 전국 주요 품목
    
    # 건설동향용 속성
    nwd.setdefault('civil_growth', None)
    nwd.setdefault('building_growth', None)
    nwd.setdefault('top_types', [])
    
    # 고용률/실업률용 속성
    nwd.setdefault('top_age_groups', [])
    nwd.setdefault('age_groups', [])
    nwd.setdefault('main_age_groups', [])
    
    # 물가동향용 속성
    nwd.setdefault('low_regions', [])
    nwd.setdefault('high_regions', [])
    
    # summary_table 기본값 (모든 부문별에 필요)
    if 'summary_table' not in data:
        data['summary_table'] = _generate_default_summary_table(data, report_id)
    
    # 보도자료 유형별 추가 기본값
    if report_id == 'export' or report_id == 'import':
        # 수출/수입 전용 속성
        if 'increase_regions' not in data:
            data['increase_regions'] = data.get('regional_data', {}).get('increase_regions', [])
        if 'decrease_regions' not in data:
            data['decrease_regions'] = data.get('regional_data', {}).get('decrease_regions', [])
        
        # regional_data의 각 지역 항목에 'change' 속성 추가 (growth_rate 사용)
        regional = data.get('regional_data', {})
        for key in ['increase_regions', 'decrease_regions', 'all']:
            if key in regional:
                for item in regional[key]:
                    if 'change' not in item and 'growth_rate' in item:
                        item['change'] = item['growth_rate']
                    if 'products' not in item:
                        item['products'] = []
        
        # top3_increase_regions, top3_decrease_regions에도 적용
        for region_list in [data.get('top3_increase_regions', []), data.get('top3_decrease_regions', [])]:
            for item in region_list:
                if 'change' not in item and 'growth_rate' in item:
                    item['change'] = item['growth_rate']
                if 'products' not in item:
                    item['products'] = []
    
    if report_id == 'price':
        # 물가동향 전용 속성
        if 'above_regions' not in data:
            data['above_regions'] = data.get('regional_data', {}).get('increase_regions', [])
        if 'below_regions' not in data:
            data['below_regions'] = data.get('regional_data', {}).get('decrease_regions', [])
    
    if report_id == 'population':
        # 인구이동 전용 속성
        if 'inflow_regions' not in data:
            data['inflow_regions'] = data.get('regional_data', {}).get('increase_regions', [])
        if 'outflow_regions' not in data:
            data['outflow_regions'] = data.get('regional_data', {}).get('decrease_regions', [])
    
    if report_id in ('employment', 'unemployment'):
        # 고용률/실업률 지역 데이터에 필수 속성 추가
        regional = data.get('regional_data', {})
        for key in ['increase_regions', 'decrease_regions', 'all']:
            if key in regional:
                for item in regional[key]:
                    if 'change' not in item and 'growth_rate' in item:
                        item['change'] = item['growth_rate']
                    if 'age_groups' not in item:
                        item['age_groups'] = []
        
        # top3_increase_regions, top3_decrease_regions에도 적용
        for region_list in [data.get('top3_increase_regions', []), data.get('top3_decrease_regions', [])]:
            for item in region_list:
                if 'change' not in item and 'growth_rate' in item:
                    item['change'] = item['growth_rate']
                if 'age_groups' not in item:
                    item['age_groups'] = []
    
    return data


def _generate_from_schema(template_name, report_id, year, quarter, custom_data=None):
    """스키마 기본값으로 보도자료 생성 (일러두기 등 generator 없는 경우)"""
    try:
        # 스키마 파일에서 기본값 로드
        schema_path = TEMPLATES_DIR / f"{report_id}_schema.json"
        if not schema_path.exists():
            return None, f"스키마 파일을 찾을 수 없습니다: {schema_path}", []
        
        with open(schema_path, 'r', encoding='utf-8') as f:
            schema = json.load(f)
        
        # 기본값 추출 (example 필드)
        data = schema.get('example', {})
        
        # 연도/분기 정보 추가
        data['report_info'] = {'year': year, 'quarter': quarter}
        
        # 일러두기의 경우 담당자 정보에서 관세청 정보 업데이트
        if report_id == 'guide' and custom_data:
            contact_info = custom_data.get('contact_info', {})
            customs_dept = contact_info.get('customs_department', '관세청 정보데이터기획담당관')
            customs_phone = contact_info.get('customs_phone', '042-481-7845')
            
            # contacts 배열에서 수출입 항목 찾아서 업데이트
            if 'contacts' in data:
                for contact in data['contacts']:
                    if contact.get('category') == '수출입':
                        contact['department'] = customs_dept
                        contact['phone'] = customs_phone
                        break
        
        # 템플릿 렌더링
        template_path = TEMPLATES_DIR / template_name
        if not template_path.exists():
            return None, f"템플릿 파일을 찾을 수 없습니다: {template_path}", []
        
        with open(template_path, 'r', encoding='utf-8') as f:
            template_content = f.read()
        
        template = Template(template_content)
        html_content = template.render(**data)
        
        print(f"[DEBUG] 스키마 기반 보도자료 생성 완료: {report_id}")
        return _remove_page_numbers(html_content), None, []
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return None, f"스키마 기반 보도자료 생성 오류: {str(e)}", []


def generate_report_html(excel_path, report_config, year, quarter, custom_data=None, raw_excel_path=None, file_type=None):
    """보도자료 HTML 생성
    
    Args:
        excel_path: 엑셀 파일 경로 (분석표 또는 기초자료)
        report_config: 보도자료 설정
        year: 연도
        quarter: 분기
        custom_data: 사용자 정의 데이터
        raw_excel_path: 기초자료 파일 경로 (옵션)
        file_type: 파일 유형 ('raw_direct', 'analysis', 'raw_with_analysis')
    """
    try:
        # 파일 존재 및 접근 가능 여부 확인
        excel_path_obj = Path(excel_path)
        if not excel_path_obj.exists():
            error_msg = f"엑셀 파일을 찾을 수 없습니다: {excel_path}"
            print(f"[ERROR] {error_msg}")
            return None, error_msg, []
        
        if not excel_path_obj.is_file():
            error_msg = f"유효한 파일이 아닙니다: {excel_path}"
            print(f"[ERROR] {error_msg}")
            return None, error_msg, []
        
        generator_name = report_config['generator']
        template_name = report_config['template']
        report_name = report_config['name']
        report_id = report_config['id']
        
        print(f"\n[DEBUG] ========== {report_name} 보도자료 생성 시작 ==========")
        print(f"[DEBUG] Generator: {generator_name}")
        print(f"[DEBUG] Template: {template_name}")
        print(f"[DEBUG] File Type: {file_type}")
        # ★ 항상 수집표(기초자료)에서 데이터 추출 (분석표 사용 안함)
        if raw_excel_path:
            print(f"[DEBUG] 수집표(기초자료) 사용: {raw_excel_path}")
            raw_data = _extract_data_from_raw(raw_excel_path, report_id, year, quarter)
            if raw_data:
                print(f"[DEBUG] 수집표 직접 추출 성공: {list(raw_data.keys())}")
                # 템플릿 호환성을 위한 기본값 추가
                raw_data = _ensure_template_compatibility(raw_data, report_id)
                
                # 템플릿 렌더링
                template_path = TEMPLATES_DIR / template_name
                if template_path.exists():
                    with open(template_path, 'r', encoding='utf-8') as f:
                        template_content = f.read()
                    
                    template = Template(template_content)
                    template.globals['is_missing'] = is_missing
                    template.globals['format_value'] = format_value
                    
                    html_content = template.render(**raw_data)
                    missing_fields = check_missing_data(raw_data, report_id)
                    return _remove_page_numbers(html_content), None, missing_fields
            else:
                error_msg = f"수집표에서 데이터를 추출할 수 없습니다: {report_id}"
                print(f"[ERROR] {error_msg}")
                return None, error_msg, []
        
        # Generator가 None인 경우 (일러두기 등) 스키마에서 기본값 로드
        if generator_name is None:
            return _generate_from_schema(template_name, report_id, year, quarter, custom_data)
        
        # 수집표가 없고 generator가 있는 경우 에러 (분석표 사용 안함)
        error_msg = f"수집표(기초자료)가 제공되지 않았습니다. 분석표 기반 처리는 지원하지 않습니다."
        print(f"[ERROR] {error_msg}")
        return None, error_msg, []
        
    except Exception as e:
        import traceback
        error_msg = f"보도자료 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return None, error_msg, []


def generate_regional_report_html(excel_path, region_name, is_reference=False, 
                                   raw_excel_path=None, year=2025, quarter=2):
    """시도별 보도자료 HTML 생성
    
    Args:
        excel_path: 분석표 엑셀 파일 경로
        region_name: 지역명 (예: '서울', '부산' 등)
        is_reference: 참고_GRDP 여부
        raw_excel_path: 기초자료 엑셀 파일 경로 (옵션)
        year: 연도
        quarter: 분기
    """
    try:
        # 파일 존재 확인
        excel_path_obj = Path(excel_path)
        if not excel_path_obj.exists() or not excel_path_obj.is_file():
            # 분석표가 없으면 기초자료 경로 사용
            if raw_excel_path and Path(raw_excel_path).exists():
                excel_path = raw_excel_path
            else:
                error_msg = f"엑셀 파일을 찾을 수 없습니다: {excel_path}"
                print(f"[ERROR] {error_msg}")
                return None, error_msg
        
        if region_name == '참고_GRDP' or is_reference:
            return generate_grdp_reference_html(excel_path)
        
        # 기초자료에서 직접 추출 시도
        if raw_excel_path and Path(raw_excel_path).exists():
            try:
                from templates.raw_data_extractor import RawDataExtractor
                
                print(f"[시도별] 기초자료에서 직접 데이터 추출 시도: {region_name}")
                extractor = RawDataExtractor(raw_excel_path, year, quarter)
                regional_data = extractor.extract_regional_data(region_name)
                
                if regional_data:
                    # 업종별 상위/하위 데이터 추가
                    regional_data['manufacturing_industries'] = extractor.extract_regional_top_industries(
                        region_name, '광공업생산', 3)
                    regional_data['service_industries'] = extractor.extract_regional_top_industries(
                        region_name, '서비스업생산', 3)
                    
                    # 템플릿 렌더링
                    template_path = TEMPLATES_DIR / 'regional_template.html'
                    if template_path.exists():
                        with open(template_path, 'r', encoding='utf-8') as f:
                            from jinja2 import Template
                            template = Template(f.read())
                        
                        html_content = template.render(**regional_data)
                        print(f"[시도별] 기초자료 직접 추출 성공: {region_name}")
                        return _remove_page_numbers(html_content), None
            except Exception as e:
                import traceback
                error_msg = f"[시도별] 수집표에서 데이터 추출 실패: {e}"
                print(f"[ERROR] {error_msg}")
                traceback.print_exc()
                return None, error_msg
        
        # 수집표가 없으면 에러 (분석표 사용 안함)
        error_msg = f"수집표(기초자료)가 제공되지 않았습니다. 분석표 기반 처리는 지원하지 않습니다."
        print(f"[ERROR] {error_msg}")
        return None, error_msg
        
    except Exception as e:
        import traceback
        error_msg = f"시도별 보도자료 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return None, error_msg


def _generate_na_grdp_data(year, quarter):
    """N/A GRDP 데이터 생성 (모든 값이 None - GRDP 파일이 없을 때 사용)"""
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


def generate_grdp_reference_html(excel_path, session_data=None):
    """참고_GRDP 보도자료 HTML 생성"""
    try:
        from flask import session
        year = session.get('year', 2025) if session_data is None else session_data.get('year', 2025)
        quarter = session.get('quarter', 2) if session_data is None else session_data.get('quarter', 2)
        
        grdp_data = None
        
        # 1. 세션에서 추출된 GRDP 데이터 확인
        if session_data is None:
            if 'grdp_data' in session and session['grdp_data']:
                grdp_data = session['grdp_data']
                print(f"[GRDP] 세션에서 GRDP 데이터 로드 (전국 {grdp_data['national_summary']['growth_rate']}%)")
        else:
            grdp_data = session_data.get('grdp_data')
        
        # 2. 추출된 JSON 파일 확인
        if grdp_data is None:
            grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
            if grdp_json_path.exists():
                with open(grdp_json_path, 'r', encoding='utf-8') as f:
                    grdp_data = json.load(f)
                print(f"[GRDP] JSON 파일에서 GRDP 데이터 로드")
        
        # 3. 기초자료 수집표에서 직접 추출 시도
        if grdp_data is None and session_data is None:
            raw_path = session.get('raw_excel_path')
            if raw_path:
                try:
                    from data_converter import DataConverter
                    converter = DataConverter(raw_path)
                    grdp_data = converter.extract_grdp_data()
                    session['grdp_data'] = grdp_data
                    print(f"[GRDP] 기초자료에서 GRDP 데이터 추출")
                except Exception as e:
                    print(f"[GRDP] 기초자료 추출 실패: {e}")
        
        # 4. uploads 폴더에서 KOSIS GRDP 파일 확인
        if grdp_data is None:
            kosis_grdp_files = list(UPLOAD_FOLDER.glob('*grdp*.xlsx')) + list(UPLOAD_FOLDER.glob('*GRDP*.xlsx'))
            kosis_grdp_files += list(UPLOAD_FOLDER.glob('*지역내총생산*.xlsx'))
            
            for grdp_file in kosis_grdp_files:
                print(f"[GRDP] KOSIS GRDP 파일 발견: {grdp_file}")
                kosis_grdp_data = parse_kosis_grdp_file(str(grdp_file), year, quarter)
                if kosis_grdp_data:
                    grdp_data = kosis_grdp_data
                    if session_data is None:
                        session['grdp_data'] = grdp_data
                    grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
                    with open(grdp_json_path, 'w', encoding='utf-8') as f:
                        json.dump(grdp_data, f, ensure_ascii=False, indent=2)
                    print(f"[GRDP] KOSIS GRDP 파일에서 데이터 파싱 성공")
                    break
        
        # 5. 참고_GRDP Generator 로드 시도
        if grdp_data is None:
            grdp_generator_path = TEMPLATES_DIR / 'reference_grdp_generator.py'
            if grdp_generator_path.exists():
                spec = importlib.util.spec_from_file_location('reference_grdp_generator', str(grdp_generator_path))
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                
                if hasattr(module, 'generate_report_data'):
                    grdp_data = module.generate_report_data(excel_path, year, quarter, use_sample=True)
        
        # 6. GRDP 데이터 없으면 N/A로 표시 (기본값 사용 안 함)
        if grdp_data is None:
            grdp_data = _generate_na_grdp_data(year, quarter)
        
        # 7. 권역 그룹 정렬 및 플래그 추가 (is_group_start, group_size)
        if 'regional_data' in grdp_data:
            grdp_data['regional_data'] = _add_grdp_group_flags(grdp_data['regional_data'])
        
        # 연도/분기 정보 업데이트
        if 'report_info' not in grdp_data:
            grdp_data['report_info'] = {}
        grdp_data['report_info']['year'] = year
        grdp_data['report_info']['quarter'] = quarter
        
        # chart_config 기본값 추가 (누락된 경우)
        if 'chart_config' not in grdp_data:
            grdp_data['chart_config'] = {
                'y_axis': {
                    'min': -6,
                    'max': 8,
                    'step': 2
                }
            }
        
        # 템플릿 렌더링
        template_path = TEMPLATES_DIR / 'reference_grdp_template.html'
        if template_path.exists():
            with open(template_path, 'r', encoding='utf-8') as f:
                template = Template(f.read())
            html_content = template.render(**grdp_data)
        else:
            html_content = _generate_default_grdp_html(grdp_data)
        
        return _remove_page_numbers(html_content), None
        
    except Exception as e:
        import traceback
        error_msg = f"참고_GRDP 보도자료 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return None, error_msg


def _add_grdp_group_flags(regional_data):
    """GRDP 데이터에 권역 그룹 플래그 추가 및 순서 정렬"""
    # 권역별 지역 순서
    REGION_ORDER = [
        {"group": None, "region": "전국"},
        {"group": "경인", "region": "서울"},
        {"group": "경인", "region": "인천"},
        {"group": "경인", "region": "경기"},
        {"group": "충청", "region": "대전"},
        {"group": "충청", "region": "세종"},
        {"group": "충청", "region": "충북"},
        {"group": "충청", "region": "충남"},
        {"group": "호남", "region": "광주"},
        {"group": "호남", "region": "전북"},
        {"group": "호남", "region": "전남"},
        {"group": "호남", "region": "제주"},
        {"group": "동북", "region": "대구"},
        {"group": "동북", "region": "경북"},
        {"group": "동북", "region": "강원"},
        {"group": "동남", "region": "부산"},
        {"group": "동남", "region": "울산"},
        {"group": "동남", "region": "경남"},
    ]
    
    # region -> item 매핑
    region_map = {item.get('region'): item for item in regional_data}
    
    # 권역별 지역 수 계산
    group_counts = {}
    for r in REGION_ORDER:
        g = r["group"]
        if g:
            group_counts[g] = group_counts.get(g, 0) + 1
    
    # 정렬된 데이터 생성
    sorted_data = []
    prev_group = None
    
    for region_info in REGION_ORDER:
        region = region_info["region"]
        current_group = region_info["group"]
        
        # 기존 데이터에서 해당 지역 찾기
        item = region_map.get(region)
        if item is None:
            # 없으면 플레이스홀더 생성 (결측치는 None으로 표시)
            item = {
                'region': region,
                'growth_rate': None,
                'manufacturing': None,
                'construction': None,
                'service': None,
                'other': None,
                'placeholder': True
            }
        else:
            item = item.copy()  # 원본 수정 방지
        
        # 권역 그룹 정보 추가
        item['region_group'] = current_group
        
        # 그룹 시작 플래그
        is_group_start = (current_group is not None) and (current_group != prev_group)
        item['is_group_start'] = is_group_start
        item['group_size'] = group_counts.get(current_group, 0) if is_group_start else 0
        
        sorted_data.append(item)
        prev_group = current_group
    
    return sorted_data


def _generate_default_grdp_html(grdp_data):
    """기본 GRDP 참고자료 HTML 생성"""
    html = """
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>참고 - 분기 지역내총생산(GRDP)</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap');
        
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        body {
            font-family: 'Noto Sans KR', sans-serif;
            font-size: 10pt;
            line-height: 1.6;
            color: #000;
            background: #fff;
            padding: 20px 40px;
        }
        
        .report-container { max-width: 800px; margin: 0 auto; }
        
        h2 {
            font-size: 14pt;
            font-weight: bold;
            margin-bottom: 15px;
            border-bottom: 2px solid #000;
            padding-bottom: 5px;
        }
        
        .info-box {
            border: 1px dotted #666;
            padding: 15px;
            margin-bottom: 20px;
            background-color: #f9f9f9;
        }
        
        .info-box p {
            margin-bottom: 10px;
        }
        
        .data-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 9pt;
            margin-top: 20px;
        }
        
        .data-table th, .data-table td {
            border: 1px solid #000;
            padding: 4px 6px;
            text-align: center;
        }
        
        .data-table th {
            background-color: #e3f2fd;
            font-weight: 500;
        }
        
        .footnote {
            font-size: 8pt;
            color: #333;
            margin-top: 10px;
        }
    </style>
</head>
<body>
    <div class="report-container">
        <h2>〔참고〕 분기 지역내총생산(GRDP)</h2>
        
        <div class="info-box">
            <p><strong>■ 분기 지역내총생산(GRDP)이란?</strong></p>
            <p>일정 기간 동안에 일정 지역 내에서 새로이 창출된 최종생산물을 시장가격으로 평가한 가치의 합입니다.</p>
            <p>분기 GRDP는 시도별 경제성장 동향을 파악하는 주요 지표로 활용됩니다.</p>
        </div>
        
        <div class="info-box">
            <p><strong>■ 참고사항</strong></p>
            <p>· 현재 분기 GRDP 데이터는 별도 발표 자료를 참조하시기 바랍니다.</p>
            <p>· 본 보도자료에서는 분기 GRDP의 전년동기비 증감률을 시도별로 제공합니다.</p>
        </div>
        
        <div class="footnote">
            자료: 통계청, 지역소득(GRDP)
        </div>
    </div>
</body>
</html>
"""
    return html


def generate_statistics_report_html(excel_path, year, quarter, raw_excel_path=None):
    """통계표 보도자료 HTML 생성"""
    try:
        generator_path = TEMPLATES_DIR / 'statistics_table_generator.py'
        if not generator_path.exists():
            return None, f"통계표 Generator를 찾을 수 없습니다"
        
        spec = importlib.util.spec_from_file_location('통계표_generator', str(generator_path))
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        
        generator = module.통계표Generator(
            excel_path,
            raw_excel_path=raw_excel_path,
            current_year=year,
            current_quarter=quarter
        )
        template_path = TEMPLATES_DIR / 'statistics_table_template.html'
        
        html_content = generator.render_html(str(template_path), year=year, quarter=quarter)
        
        return _remove_page_numbers(html_content), None
        
    except Exception as e:
        import traceback
        error_msg = f"통계표 보도자료 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return None, error_msg


def generate_individual_statistics_html(excel_path, stat_config, year, quarter, raw_excel_path=None):
    """개별 통계표 HTML 생성"""
    try:
        stat_id = stat_config['id']
        template_name = stat_config['template']
        table_name = stat_config.get('table_name')
        
        # 통계표 Generator 모듈 로드
        generator_path = TEMPLATES_DIR / 'statistics_table_generator.py'
        if generator_path.exists():
            spec = importlib.util.spec_from_file_location('statistics_table_generator', str(generator_path))
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            generator = module.StatisticsTableGenerator(
                excel_path,
                raw_excel_path=raw_excel_path,
                current_year=year,
                current_quarter=quarter
            )
        else:
            generator = None
        
        PAGE1_REGIONS = ["전국", "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종"]
        PAGE2_REGIONS = ["경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"]
        
        # 통계표 목차
        if stat_id == 'stat_toc':
            toc_items = [
                {'number': 1, 'name': '광공업생산지수'},
                {'number': 2, 'name': '서비스업생산지수'},
                {'number': 3, 'name': '소매판매액지수'},
                {'number': 4, 'name': '건설수주액'},
                {'number': 5, 'name': '고용률'},
                {'number': 6, 'name': '실업률'},
                {'number': 7, 'name': '국내 인구이동'},
                {'number': 8, 'name': '수출액'},
                {'number': 9, 'name': '수입액'},
                {'number': 10, 'name': '소비자물가지수'},
            ]
            template_data = {
                'year': year,
                'quarter': quarter,
                'toc_items': toc_items,
                'page_number': 21
            }
        
        # 통계표 - 개별 지표
        elif table_name and table_name != 'GRDP' and generator:
            table_order = ['광공업생산지수', '서비스업생산지수', '소매판매액지수', '건설수주액',
                          '고용률', '실업률', '국내인구이동', '수출액', '수입액', '소비자물가지수']
            try:
                table_index = table_order.index(table_name) + 1
            except ValueError:
                table_index = 1
            
            config = generator.TABLE_CONFIG.get(table_name)
            if config:
                data = generator.extract_table_data(table_name)
                
                # 연도 키: JSON 데이터에서 가져오거나 기본값 사용
                yearly_years = data.get('yearly_years', ["2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"])
                
                # 분기 키: 실제 데이터에 있는 분기만 사용 (데이터 없는 분기 제외)
                quarterly_keys = data.get('quarterly_keys', [])
                if not quarterly_keys and data.get('quarterly'):
                    # quarterly_keys가 없으면 quarterly 딕셔너리에서 키 추출 후 정렬
                    quarterly_keys = sorted(data['quarterly'].keys(), key=lambda x: (
                        int(x[:4]), int(x[5]) if len(x) > 5 else 0
                    ))
                
                page_base = 22 + (table_index - 1) * 2
                
                template_data = {
                    'year': year,
                    'quarter': quarter,
                    'index': table_index,
                    'title': table_name,
                    'unit': config['단위'],
                    'data': data if data else {'yearly': {}, 'quarterly': {}},
                    'page1_regions': PAGE1_REGIONS,
                    'page2_regions': PAGE2_REGIONS,
                    'yearly_years': yearly_years,
                    'quarterly_keys': quarterly_keys,
                    'page_number_1': page_base,
                    'page_number_2': page_base + 1
                }
            else:
                return None, f"통계표 설정을 찾을 수 없습니다: {table_name}"
        
        # 통계표 - GRDP
        elif stat_id == 'stat_grdp':
            if generator:
                grdp_data = generator._create_grdp_placeholder()
            else:
                grdp_data = {
                    'data': {
                        'yearly': {},
                        'quarterly': {},
                        'yearly_years': [],
                        'quarterly_keys': []
                    }
                }
            
            # grdp_data에서 yearly_years와 quarterly_keys 가져오기
            data_dict = grdp_data.get('data', {'yearly': {}, 'quarterly': {}})
            yearly_years = data_dict.get('yearly_years', ["2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"])
            quarterly_keys = data_dict.get('quarterly_keys', [])
            
            template_data = {
                'year': year,
                'quarter': quarter,
                'data': data_dict,
                'page1_regions': PAGE1_REGIONS,
                'page2_regions': PAGE2_REGIONS,
                'yearly_years': yearly_years,
                'quarterly_keys': quarterly_keys,
                'page_number_1': 42,
                'page_number_2': 43
            }
        
        # 부록 - 주요 용어 정의
        elif stat_id == 'stat_appendix':
            terms_page1 = [
                {"term": "불변지수", "definition": "불변지수는 가격 변동분이 제외된 수량 변동분만 포함되어 있음을 의미하며, 성장 수준 분석(전년동분기비)에 활용됨"},
                {"term": "광공업생산지수", "definition": "한국표준산업분류 상의 3개 대분류(B, C, D)를 대상으로 광업제조업동향조사의 월별 품목별 생산·출하(내수 및 수출)·재고 및 생산능력·가동률지수를 기초로 작성됨"},
                {"term": "서비스업생산지수", "definition": "한국표준산업분류 상의 13개 대분류(E, G, H, I, J, K, L, M, N, P, Q, R, S)를 대상으로 서비스업동향조사의 월별 매출액을 기초로 작성됨"},
                {"term": "소매판매액지수", "definition": "한국표준산업분류 상의 '자동차 판매업 중 승용차'와 '소매업'을 대상으로 서비스업동향조사의 월별 상품판매액을 기초로 작성됨"},
                {"term": "건설수주", "definition": "종합건설업 등록업체 중 전전년 「건설업조사」 결과를 기준으로 기성액 순위 상위 기업체(대표도: 54%)의 국내공사에 대한 건설수주액임"},
                {"term": "소비자물가지수", "definition": "가구에서 일상생활을 영위하기 위해 구입하는 상품과 서비스의 평균적인 가격변동을 측정한 지수임"},
                {"term": "지역내총생산", "definition": "일정 기간 동안에 일정 지역 내에서 새로이 창출된 최종생산물을 시장가격으로 평가한 가치의 합임"},
            ]
            terms_page2 = [
                {"term": "고용률", "definition": "만 15세 이상 인구 중 취업자가 차지하는 비율로, 노동시장의 고용흡수력을 나타내는 지표"},
                {"term": "실업률", "definition": "경제활동인구 중 실업자가 차지하는 비율로, 노동시장의 수급상황을 파악하는 대표적 지표"},
                {"term": "국내인구이동", "definition": "주민등록법에 의한 전입신고를 집계한 것으로, 시·도 간 순이동을 의미함"},
                {"term": "수출액", "definition": "관세선을 통과하여 외국으로 반출하는 물품의 가액으로, FOB(본선인도가격) 기준으로 집계"},
                {"term": "수입액", "definition": "관세선을 통과하여 국내로 반입하는 물품의 가액으로, CIF(운임·보험료포함가격) 기준으로 집계"},
            ]
            
            template_data = {
                'year': year,
                'quarter': quarter,
                'terms_page1': terms_page1,
                'terms_page2': terms_page2,
                'page_number_1': 44,
                'page_number_2': 45
            }
        
        else:
            return None, f"알 수 없는 통계표 ID: {stat_id}"
        
        # 템플릿 렌더링
        template_path = TEMPLATES_DIR / template_name
        if not template_path.exists():
            return None, f"템플릿을 찾을 수 없습니다: {template_name}"
        
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        html_content = template.render(**template_data)
        return _remove_page_numbers(html_content), None
        
    except Exception as e:
        import traceback
        error_msg = f"개별 통계표 생성 오류 ({stat_config.get('name', 'unknown')}): {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return None, error_msg

