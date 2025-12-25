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

# 데이터 변환 모듈 임포트
from data_converter import DataConverter, convert_raw_to_analysis

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

# ===== 결측치 처리를 위한 Jinja2 커스텀 필터 =====
def is_missing(value):
    """결측치 여부 확인"""
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() in ['', '-', 'N/A', '없음', '미입력', '비공개']
    if isinstance(value, float) and pd.isna(value):
        return True
    return False

def format_value(value, format_str="%.1f", placeholder="[  ]"):
    """값 포맷팅 (결측치는 플레이스홀더로)"""
    if is_missing(value):
        return f'<span class="editable-placeholder" contenteditable="true">{placeholder}</span>'
    try:
        if format_str:
            return format_str % float(value)
        return str(value)
    except (ValueError, TypeError):
        return f'<span class="editable-placeholder" contenteditable="true">{placeholder}</span>'

def editable(value, format_str="%.1f"):
    """편집 가능한 값 표시 (결측치는 노란색 하이라이트)"""
    if is_missing(value):
        return f'<span class="editable-placeholder" contenteditable="true">[  ]</span>'
    try:
        formatted = format_str % float(value) if format_str else str(value)
        return formatted
    except (ValueError, TypeError):
        return f'<span class="editable-placeholder" contenteditable="true">[  ]</span>'

# Jinja2 필터 등록
app.jinja_env.filters['is_missing'] = is_missing
app.jinja_env.filters['format_value'] = format_value
app.jinja_env.filters['editable'] = editable
app.jinja_env.globals['is_missing'] = is_missing

# ===== 요약 보고서 목록 (표지-일러두기-목차-인포그래픽-요약 순서) =====
SUMMARY_REPORTS = [
    {
        'id': 'cover',
        'name': '표지',
        'sheet': None,
        'generator': None,
        'template': '표지_template.html',
        'icon': '📑',
        'category': 'summary'
    },
    {
        'id': 'guide',
        'name': '일러두기',
        'sheet': None,
        'generator': None,
        'template': '일러두기_template.html',
        'icon': '📖',
        'category': 'summary'
    },
    {
        'id': 'toc',
        'name': '목차',
        'sheet': None,
        'generator': None,
        'template': '목차_template.html',
        'icon': '📋',
        'category': 'summary'
    },
    {
        'id': 'infographic',
        'name': '인포그래픽',
        'sheet': 'multiple',
        'generator': '인포그래픽_generator.py',
        'template': '인포그래픽_js_template.html',
        'icon': '📊',
        'category': 'summary'
    },
    {
        'id': 'summary_overview',
        'name': '요약-지역경제동향',
        'sheet': 'multiple',
        'generator': '요약_지역경제동향_generator.py',
        'template': '요약_지역경제동향_template.html',
        'icon': '📈',
        'category': 'summary'
    },
    {
        'id': 'summary_production',
        'name': '요약-생산',
        'sheet': 'multiple',
        'generator': '요약_생산_generator.py',
        'template': '요약_생산_template.html',
        'icon': '🏭',
        'category': 'summary'
    },
    {
        'id': 'summary_consumption',
        'name': '요약-소비건설',
        'sheet': 'multiple',
        'generator': '요약_소비건설_generator.py',
        'template': '요약_소비건설_template.html',
        'icon': '🛒',
        'category': 'summary'
    },
    {
        'id': 'summary_trade_price',
        'name': '요약-수출물가',
        'sheet': 'multiple',
        'generator': '요약_수출물가_generator.py',
        'template': '요약_수출물가_template.html',
        'icon': '📦',
        'category': 'summary'
    },
    {
        'id': 'summary_employment',
        'name': '요약-고용인구',
        'sheet': 'multiple',
        'generator': '요약_고용인구_generator.py',
        'template': '요약_고용인구_template.html',
        'icon': '👔',
        'category': 'summary'
    },
]

# ===== 부문별 보고서 순서 설정 (광공업생산-서비스업생산-소비동향-건설동향-수출-수입-물가동향-고용률-실업률-국내인구이동) =====
SECTOR_REPORTS = [
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
        'id': 'construction',
        'name': '건설동향',
        'sheet': "F'분석",
        'generator': '건설동향_generator.py',
        'template': '건설동향_template.html',
        'icon': '🏗️',
        'category': 'construction'
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
        'id': 'price',
        'name': '물가동향',
        'sheet': 'E(품목성질물가)분석',
        'generator': '물가동향_generator.py',
        'template': '물가동향_template.html',
        'icon': '💰',
        'category': 'price'
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
        'id': 'population',
        'name': '국내인구이동',
        'sheet': 'I(순인구이동)집계',
        'generator': '국내인구이동_generator.py',
        'template': '국내인구이동_template.html',
        'icon': '👥',
        'category': 'population'
    }
]

# 전체 보고서 순서 (요약 → 부문별)
REPORT_ORDER = SUMMARY_REPORTS + SECTOR_REPORTS

# ===== 통계표 보고서 목록 (통계표-목차 → 각 지표 → GRDP → 부록) =====
STATISTICS_REPORTS = [
    {
        'id': 'stat_toc',
        'name': '통계표-목차',
        'table_name': None,
        'template': '통계표_목차_template.html',
        'icon': '📋',
        'category': 'statistics'
    },
    {
        'id': 'stat_mining',
        'name': '통계표-광공업생산지수',
        'table_name': '광공업생산지수',
        'template': '통계표_지표_template.html',
        'icon': '🏭',
        'category': 'statistics'
    },
    {
        'id': 'stat_service',
        'name': '통계표-서비스업생산지수',
        'table_name': '서비스업생산지수',
        'template': '통계표_지표_template.html',
        'icon': '🏢',
        'category': 'statistics'
    },
    {
        'id': 'stat_retail',
        'name': '통계표-소매판매액지수',
        'table_name': '소매판매액지수',
        'template': '통계표_지표_template.html',
        'icon': '🛒',
        'category': 'statistics'
    },
    {
        'id': 'stat_construction',
        'name': '통계표-건설수주액',
        'table_name': '건설수주액',
        'template': '통계표_지표_template.html',
        'icon': '🏗️',
        'category': 'statistics'
    },
    {
        'id': 'stat_employment',
        'name': '통계표-고용률',
        'table_name': '고용률',
        'template': '통계표_지표_template.html',
        'icon': '👔',
        'category': 'statistics'
    },
    {
        'id': 'stat_unemployment',
        'name': '통계표-실업률',
        'table_name': '실업률',
        'template': '통계표_지표_template.html',
        'icon': '📉',
        'category': 'statistics'
    },
    {
        'id': 'stat_population',
        'name': '통계표-국내인구이동',
        'table_name': '국내인구이동',
        'template': '통계표_지표_template.html',
        'icon': '👥',
        'category': 'statistics'
    },
    {
        'id': 'stat_export',
        'name': '통계표-수출액',
        'table_name': '수출액',
        'template': '통계표_지표_template.html',
        'icon': '📦',
        'category': 'statistics'
    },
    {
        'id': 'stat_import',
        'name': '통계표-수입액',
        'table_name': '수입액',
        'template': '통계표_지표_template.html',
        'icon': '🚢',
        'category': 'statistics'
    },
    {
        'id': 'stat_price',
        'name': '통계표-소비자물가지수',
        'table_name': '소비자물가지수',
        'template': '통계표_지표_template.html',
        'icon': '💰',
        'category': 'statistics'
    },
    {
        'id': 'stat_grdp',
        'name': '통계표-참고-GRDP',
        'table_name': 'GRDP',
        'template': '통계표_GRDP_template.html',
        'icon': '📊',
        'category': 'statistics'
    },
    {
        'id': 'stat_appendix',
        'name': '부록-주요용어정의',
        'table_name': None,
        'template': '통계표_부록_template.html',
        'icon': '📖',
        'category': 'statistics'
    },
]

# 시도별 보고서 목록 (17개 시도 + 참고_GRDP)
REGIONAL_REPORTS = [
    {'id': 'region_seoul', 'name': '서울', 'full_name': '서울특별시', 'index': 1, 'icon': '🏙️'},
    {'id': 'region_busan', 'name': '부산', 'full_name': '부산광역시', 'index': 2, 'icon': '🌊'},
    {'id': 'region_daegu', 'name': '대구', 'full_name': '대구광역시', 'index': 3, 'icon': '🏛️'},
    {'id': 'region_incheon', 'name': '인천', 'full_name': '인천광역시', 'index': 4, 'icon': '✈️'},
    {'id': 'region_gwangju', 'name': '광주', 'full_name': '광주광역시', 'index': 5, 'icon': '🎨'},
    {'id': 'region_daejeon', 'name': '대전', 'full_name': '대전광역시', 'index': 6, 'icon': '🔬'},
    {'id': 'region_ulsan', 'name': '울산', 'full_name': '울산광역시', 'index': 7, 'icon': '🚗'},
    {'id': 'region_sejong', 'name': '세종', 'full_name': '세종특별자치시', 'index': 8, 'icon': '🏛️'},
    {'id': 'region_gyeonggi', 'name': '경기', 'full_name': '경기도', 'index': 9, 'icon': '🏘️'},
    {'id': 'region_gangwon', 'name': '강원', 'full_name': '강원특별자치도', 'index': 10, 'icon': '⛰️'},
    {'id': 'region_chungbuk', 'name': '충북', 'full_name': '충청북도', 'index': 11, 'icon': '🌾'},
    {'id': 'region_chungnam', 'name': '충남', 'full_name': '충청남도', 'index': 12, 'icon': '🌅'},
    {'id': 'region_jeonbuk', 'name': '전북', 'full_name': '전북특별자치도', 'index': 13, 'icon': '🌿'},
    {'id': 'region_jeonnam', 'name': '전남', 'full_name': '전라남도', 'index': 14, 'icon': '🍃'},
    {'id': 'region_gyeongbuk', 'name': '경북', 'full_name': '경상북도', 'index': 15, 'icon': '🏔️'},
    {'id': 'region_gyeongnam', 'name': '경남', 'full_name': '경상남도', 'index': 16, 'icon': '🌳'},
    {'id': 'region_jeju', 'name': '제주', 'full_name': '제주특별자치도', 'index': 17, 'icon': '🏝️'},
    {'id': 'reference_grdp', 'name': '참고_GRDP', 'full_name': '분기 지역내총생산(GRDP)', 'index': 18, 'icon': '📊', 'is_reference': True},
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
    # - 전국 데이터는 generator에서 분류단계 0(총지수) 또는 시도별 합계로 이미 계산됨
    # - 결측치 체크는 최소한으로 유지 (실제 렌더링에 필수적인 것만)
    REQUIRED_FIELDS = {
        'manufacturing': [],  # generator가 전국 데이터를 분류단계 0에서 추출
        'service': [],        # generator가 전국 데이터를 분류단계 0에서 추출
        'consumption': [],    # generator가 전국 데이터를 분류단계 0에서 추출
        'employment': [],     # generator가 전국 데이터를 추출
        'unemployment': [],   # generator가 전국 데이터를 추출
        'price': [],          # generator가 전국 데이터를 추출
        'export': [],         # generator가 전국 데이터를 추출
        'import': [],         # generator가 전국 데이터를 추출
        'population': [],     # generator가 전국 데이터를 추출
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


def generate_report_html(excel_path, report_config, year, quarter, custom_data=None, raw_excel_path=None):
    """보고서 HTML 생성"""
    try:
        generator_name = report_config['generator']
        template_name = report_config['template']
        report_name = report_config['name']
        report_id = report_config['id']
        
        print(f"\n[DEBUG] ========== {report_name} 보고서 생성 시작 ==========")
        print(f"[DEBUG] Generator: {generator_name}")
        print(f"[DEBUG] Template: {template_name}")
        if raw_excel_path:
            print(f"[DEBUG] 기초자료 사용: {raw_excel_path}")
        
        # Generator 모듈 로드
        module = load_generator_module(generator_name)
        if not module:
            print(f"[ERROR] Generator 모듈을 찾을 수 없습니다: {generator_name}")
            return None, f"Generator 모듈을 찾을 수 없습니다: {generator_name}", []
        
        # 사용 가능한 함수 확인
        available_funcs = [name for name in dir(module) if not name.startswith('_')]
        print(f"[DEBUG] 모듈 내 함수/클래스: {[f for f in available_funcs if 'generate' in f.lower() or 'Generator' in f or f == 'load_data']}")
        
        # Generator 클래스 찾기
        generator_class = None
        for name in dir(module):
            obj = getattr(module, name)
            if isinstance(obj, type) and name.endswith('Generator'):
                generator_class = obj
                print(f"[DEBUG] Generator 클래스 발견: {name}")
                break
        
        data = None
        
        # ========== 데이터 추출 방식 결정 ==========
        
        # 방법 1: generate_report_data 함수 사용 (물가동향, 실업률, 수출, 수입, 국내인구이동, 건설동향)
        if hasattr(module, 'generate_report_data'):
            print(f"[DEBUG] generate_report_data 함수 사용")
            # 기초자료 지원 여부 확인 (함수 시그니처 체크는 어려우므로 try-except로 처리)
            try:
                # 기초자료 경로가 있고 함수가 raw_excel_path 파라미터를 받을 수 있는지 확인
                if raw_excel_path:
                    # 함수 시그니처 확인 불가능하므로 일단 전달 시도
                    import inspect
                    sig = inspect.signature(module.generate_report_data)
                    params = sig.parameters
                    if 'raw_excel_path' in params or 'use_raw_data' in params:
                        data = module.generate_report_data(excel_path, raw_excel_path=raw_excel_path, 
                                                         year=year, quarter=quarter)
                    else:
                        data = module.generate_report_data(excel_path)
                else:
                    data = module.generate_report_data(excel_path)
            except TypeError:
                # 파라미터가 맞지 않으면 기본 방식 사용
                data = module.generate_report_data(excel_path)
            except Exception as e:
                print(f"[WARNING] 기초자료 추출 실패, 분석표 사용: {e}")
                data = module.generate_report_data(excel_path)
            print(f"[DEBUG] 데이터 키: {list(data.keys()) if data else 'None'}")
        
        # 방법 2: generate_report 함수 직접 호출 (서비스업생산, 소비동향, 고용률)
        # - generate_report 함수가 완전한 데이터를 반환함
        elif hasattr(module, 'generate_report'):
            print(f"[DEBUG] generate_report 함수 직접 호출")
            template_path = TEMPLATES_DIR / template_name
            output_path = TEMPLATES_DIR / f"{report_name}_preview.html"
            # 기초자료 경로가 있으면 전달 시도 (함수 시그니처 확인)
            try:
                import inspect
                sig = inspect.signature(module.generate_report)
                params = sig.parameters
                if 'raw_excel_path' in params:
                    data = module.generate_report(excel_path, template_path, output_path, 
                                                 raw_excel_path=raw_excel_path, year=year, quarter=quarter)
                else:
                    data = module.generate_report(excel_path, template_path, output_path)
            except (TypeError, AttributeError):
                # 파라미터가 맞지 않으면 기본 방식 사용
                data = module.generate_report(excel_path, template_path, output_path)
            print(f"[DEBUG] 추출된 데이터 키: {list(data.keys()) if data else 'None'}")
        
        # 방법 3: Generator 클래스 사용 (광공업생산)
        elif generator_class:
            print(f"[DEBUG] Generator 클래스 사용: {generator_class.__name__}")
            # Generator 클래스는 현재 기초자료 지원하지 않음 (향후 확장 가능)
            # TODO: Generator 클래스에 raw_excel_path 파라미터 지원 추가 시 아래 주석 해제
            # if raw_excel_path:
            #     generator = generator_class(excel_path, raw_excel_path=raw_excel_path, year=year, quarter=quarter)
            # else:
            generator = generator_class(excel_path)
            data = generator.extract_all_data()
            print(f"[DEBUG] 추출된 데이터 키: {list(data.keys()) if data else 'None'}")
        
        else:
            error_msg = f"유효한 Generator를 찾을 수 없습니다: {generator_name}"
            print(f"[ERROR] {error_msg}")
            print(f"[ERROR] 사용 가능한 함수: {available_funcs}")
            return None, error_msg, []
        
        # ========== Top3 regions 후처리 ==========
        # 양쪽 키 이름 모두 제공 (템플릿 호환성: change/growth_rate, age_groups/industries)
        if data and 'regional_data' in data:
            # 이미 top3가 있으면 호환성 키만 추가
            if 'top3_increase_regions' not in data:
                top3_increase = []
                for r in data['regional_data'].get('increase_regions', [])[:3]:
                    rate_value = r.get('change', r.get('growth_rate', 0))
                    items = r.get('top_age_groups', r.get('industries', r.get('top_industries', [])))
                    top3_increase.append({
                        'region': r.get('region', ''),
                        'change': rate_value,
                        'growth_rate': rate_value,
                        'age_groups': items,
                        'industries': items
                    })
                data['top3_increase_regions'] = top3_increase
            else:
                # 기존 데이터에 호환성 키 추가
                for r in data['top3_increase_regions']:
                    if 'growth_rate' not in r:
                        r['growth_rate'] = r.get('change', 0)
                    if 'change' not in r:
                        r['change'] = r.get('growth_rate', 0)
                    if 'industries' not in r:
                        r['industries'] = r.get('age_groups', r.get('top_industries', []))
                    if 'age_groups' not in r:
                        r['age_groups'] = r.get('industries', [])
            
            if 'top3_decrease_regions' not in data:
                top3_decrease = []
                for r in data['regional_data'].get('decrease_regions', [])[:3]:
                    rate_value = r.get('change', r.get('growth_rate', 0))
                    items = r.get('top_age_groups', r.get('industries', r.get('top_industries', [])))
                    top3_decrease.append({
                        'region': r.get('region', ''),
                        'change': rate_value,
                        'growth_rate': rate_value,
                        'age_groups': items,
                        'industries': items
                    })
                data['top3_decrease_regions'] = top3_decrease
            else:
                # 기존 데이터에 호환성 키 추가
                for r in data['top3_decrease_regions']:
                    if 'growth_rate' not in r:
                        r['growth_rate'] = r.get('change', 0)
                    if 'change' not in r:
                        r['change'] = r.get('growth_rate', 0)
                    if 'industries' not in r:
                        r['industries'] = r.get('age_groups', r.get('top_industries', []))
                    if 'age_groups' not in r:
                        r['age_groups'] = r.get('industries', [])
            
            print(f"[DEBUG] Top3 regions 후처리 완료")
        
        # ========== 커스텀 데이터 병합 (사용자가 입력한 결측치) ==========
        if custom_data:
            for key, value in custom_data.items():
                keys = key.split('.')
                obj = data
                for k in keys[:-1]:
                    if '[' in k:
                        name, idx = k.replace(']', '').split('[')
                        obj = obj[name][int(idx)]
                    else:
                        if k not in obj:
                            obj[k] = {}
                        obj = obj[k]
                final_key = keys[-1]
                if '[' in final_key:
                    name, idx = final_key.replace(']', '').split('[')
                    obj[name][int(idx)] = value
                else:
                    obj[final_key] = value
        
        # 결측치 확인
        missing = check_missing_data(data, report_id)
        
        # ========== 템플릿 렌더링 ==========
        template_path = TEMPLATES_DIR / template_name
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        # 커스텀 필터 등록
        template.environment.filters['format_value'] = format_value
        template.environment.filters['is_missing'] = is_missing
        
        # 모든 템플릿은 {{ xxx }} 형태로 직접 접근 (통일된 방식)
        html_content = template.render(**data)
        
        print(f"[DEBUG] 보고서 생성 성공!")
        return html_content, None, missing
        
    except Exception as e:
        import traceback
        error_msg = f"보고서 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return None, error_msg, []


def generate_regional_report_html(excel_path, region_name, is_reference=False):
    """시도별 보고서 HTML 생성"""
    try:
        # 참고_GRDP인 경우 별도 처리
        if region_name == '참고_GRDP' or is_reference:
            return generate_grdp_reference_html(excel_path)
        
        # 시도별 Generator 모듈 로드
        generator_path = TEMPLATES_DIR / '시도별_generator.py'
        if not generator_path.exists():
            return None, f"시도별 Generator를 찾을 수 없습니다"
        
        spec = importlib.util.spec_from_file_location('시도별_generator', str(generator_path))
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        
        # Generator 클래스 사용
        generator = module.시도별Generator(excel_path)
        template_path = TEMPLATES_DIR / '시도별_template.html'
        
        # HTML 생성
        html_content = generator.render_html(region_name, str(template_path))
        
        return html_content, None
        
    except Exception as e:
        import traceback
        error_msg = f"시도별 보고서 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return None, error_msg


def generate_grdp_reference_html(excel_path):
    """참고_GRDP 보고서 HTML 생성
    
    데이터 소스 우선순위:
    1. 세션에 저장된 GRDP 데이터 (기초자료에서 추출된 경우)
    2. grdp_extracted.json 파일
    3. 기초자료 수집표에서 직접 추출
    4. 참고_GRDP Generator (플레이스홀더)
    5. KOSIS API (비활성화 상태)
    """
    try:
        year = session.get('year', 2025)
        quarter = session.get('quarter', 2)
        
        grdp_data = None
        
        # 1. 세션에서 추출된 GRDP 데이터 확인 (기초자료에서 추출된 경우)
        if 'grdp_data' in session and session['grdp_data']:
            grdp_data = session['grdp_data']
            print(f"[GRDP] 세션에서 GRDP 데이터 로드 (전국 {grdp_data['national_summary']['growth_rate']}%)")
        
        # 2. 추출된 JSON 파일 확인
        if grdp_data is None:
            grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
            if grdp_json_path.exists():
                with open(grdp_json_path, 'r', encoding='utf-8') as f:
                    grdp_data = json.load(f)
                print(f"[GRDP] JSON 파일에서 GRDP 데이터 로드")
        
        # 3. 기초자료 수집표에서 직접 추출 시도
        if grdp_data is None and session.get('raw_excel_path'):
            raw_path = session.get('raw_excel_path')
            try:
                converter = DataConverter(raw_path)
                grdp_data = converter.extract_grdp_data()
                session['grdp_data'] = grdp_data
                print(f"[GRDP] 기초자료에서 GRDP 데이터 추출")
            except Exception as e:
                print(f"[GRDP] 기초자료 추출 실패: {e}")
        
        # 4. 참고_GRDP Generator 로드 시도 (기존 방식)
        if grdp_data is None:
            grdp_generator_path = TEMPLATES_DIR / '참고_GRDP_generator.py'
            if grdp_generator_path.exists():
                spec = importlib.util.spec_from_file_location('참고_GRDP_generator', str(grdp_generator_path))
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                
                if hasattr(module, 'generate_report_data'):
                    grdp_data = module.generate_report_data(excel_path, year, quarter, use_sample=True)
        
        # 5. KOSIS API에서 데이터 가져오기 시도 (현재 비활성화)
        # TODO: 향후 활성화 시 아래 주석을 해제하고 kosis_grdp_fetcher 모듈의 ENABLE_KOSIS_GRDP_FETCH를 True로 설정하세요
        # if grdp_data is None:
        #     try:
        #         from kosis_grdp_fetcher import fetch_grdp_from_kosis
        #         print(f"[GRDP] KOSIS API에서 GRDP 데이터 수집 시도")
        #         kosis_data = fetch_grdp_from_kosis(
        #             start_year=2016,
        #             start_quarter=1,
        #             end_year=year,
        #             end_quarter=quarter
        #         )
        #         if kosis_data:
        #             grdp_data = kosis_data
        #             # 세션에 저장 (다음 요청 시 재사용)
        #             session['grdp_data'] = grdp_data
        #             print(f"[GRDP] KOSIS API에서 GRDP 데이터 수집 성공")
        #     except ImportError:
        #         print(f"[GRDP] kosis_grdp_fetcher 모듈을 찾을 수 없습니다")
        #     except Exception as e:
        #         print(f"[GRDP] KOSIS API 데이터 수집 실패: {e}")
        
        # 6. Generator가 없거나 실패하면 기본 데이터 사용
        if grdp_data is None:
            grdp_data = _get_default_grdp_data(year, quarter)
        
        # 참고_GRDP 템플릿 렌더링
        template_path = TEMPLATES_DIR / '참고_GRDP_template.html'
        if template_path.exists():
            with open(template_path, 'r', encoding='utf-8') as f:
                template = Template(f.read())
            html_content = template.render(**grdp_data)
        else:
            # 기본 GRDP 참고자료 HTML 생성
            html_content = _generate_default_grdp_html(grdp_data)
        
        return html_content, None
        
    except Exception as e:
        import traceback
        error_msg = f"참고_GRDP 보고서 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return None, error_msg


def _get_default_grdp_data(year, quarter):
    """기본 GRDP 데이터"""
    regions = ['전국', '서울', '인천', '경기', '대전', '세종', '충북', '충남',
               '광주', '전북', '전남', '제주', '대구', '경북', '강원', '부산', '울산', '경남']
    
    regional_data = []
    region_groups = {
        '서울': '경인', '인천': '경인', '경기': '경인',
        '대전': '충청', '세종': '충청', '충북': '충청', '충남': '충청',
        '광주': '호남', '전북': '호남', '전남': '호남', '제주': '호남',
        '대구': '동북', '경북': '동북', '강원': '동북',
        '부산': '동남', '울산': '동남', '경남': '동남'
    }
    
    for region in regions:
        regional_data.append({
            'region': region,
            'region_group': region_groups.get(region, ''),
            'growth_rate': 0.0,
            'manufacturing': 0.0,
            'construction': 0.0,
            'service': 0.0,
            'other': 0.0,
            'placeholder': True
        })
    
    return {
        'report_info': {
            'year': year,
            'quarter': quarter,
            'page_number': ''
        },
        'national_summary': {
            'growth_rate': 0.0,
            'direction': '증가',
            'contributions': {
                'manufacturing': 0.0,
                'construction': 0.0,
                'service': 0.0,
                'other': 0.0
            },
            'placeholder': True
        },
        'top_region': {
            'name': '-',
            'growth_rate': 0.0,
            'contributions': {
                'manufacturing': 0.0,
                'construction': 0.0,
                'service': 0.0,
                'other': 0.0
            },
            'placeholder': True
        },
        'regional_data': regional_data,
        'chart_config': {
            'y_axis': {
                'min': -6,
                'max': 8,
                'step': 2
            }
        }
    }


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
            <p>· 본 보고서에서는 분기 GRDP의 전년동기비 증감률을 시도별로 제공합니다.</p>
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
    """통계표 보고서 HTML 생성"""
    try:
        # 통계표 Generator 모듈 로드
        generator_path = TEMPLATES_DIR / '통계표_generator.py'
        if not generator_path.exists():
            return None, f"통계표 Generator를 찾을 수 없습니다"
        
        spec = importlib.util.spec_from_file_location('통계표_generator', str(generator_path))
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        
        # Generator 클래스 사용 (기초자료 경로 전달)
        generator = module.통계표Generator(
            excel_path,
            raw_excel_path=raw_excel_path,
            current_year=year,
            current_quarter=quarter
        )
        template_path = TEMPLATES_DIR / '통계표_template.html'
        
        # HTML 생성
        html_content = generator.render_html(str(template_path), year=year, quarter=quarter)
        
        return html_content, None
        
    except Exception as e:
        import traceback
        error_msg = f"통계표 보고서 생성 오류: {str(e)}"
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
        generator_path = TEMPLATES_DIR / '통계표_generator.py'
        if generator_path.exists():
            spec = importlib.util.spec_from_file_location('통계표_generator', str(generator_path))
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            generator = module.통계표Generator(
                excel_path,
                raw_excel_path=raw_excel_path,
                current_year=year,
                current_quarter=quarter
            )
        else:
            generator = None
        
        # 페이지 1/2 지역 목록
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
            # 지표 인덱스 계산
            table_order = ['광공업생산지수', '서비스업생산지수', '소매판매액지수', '건설수주액',
                          '고용률', '실업률', '국내인구이동', '수출액', '수입액', '소비자물가지수']
            try:
                table_index = table_order.index(table_name) + 1
            except ValueError:
                table_index = 1
            
            # 데이터 추출
            config = generator.TABLE_CONFIG.get(table_name)
            if config:
                data = generator.extract_table_data(table_name)
                
                # 연도/분기 키 목록
                yearly_years = ["2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"]
                quarterly_keys = [
                    "2016.4/4",
                    "2017.1/4", "2017.2/4", "2017.3/4", "2017.4/4",
                    "2018.1/4", "2018.2/4", "2018.3/4", "2018.4/4",
                    "2019.1/4", "2019.2/4", "2019.3/4", "2019.4/4",
                    "2020.1/4", "2020.2/4", "2020.3/4", "2020.4/4",
                    "2021.1/4", "2021.2/4", "2021.3/4", "2021.4/4",
                    "2022.1/4", "2022.2/4", "2022.3/4", "2022.4/4",
                    "2023.1/4", "2023.2/4", "2023.3/4", "2023.4/4",
                    "2024.1/4", "2024.2/4", "2024.3/4", "2024.4/4",
                    "2025.1/4", f"2025.{quarter}/4p"
                ]
                
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
            
            yearly_years = ["2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"]
            quarterly_keys = [
                "2016.4/4",
                "2017.1/4", "2017.2/4", "2017.3/4", "2017.4/4",
                "2018.1/4", "2018.2/4", "2018.3/4", "2018.4/4",
                "2019.1/4", "2019.2/4", "2019.3/4", "2019.4/4",
                "2020.1/4", "2020.2/4", "2020.3/4", "2020.4/4",
                "2021.1/4", "2021.2/4", "2021.3/4", "2021.4/4",
                "2022.1/4", "2022.2/4", "2022.3/4", "2022.4/4",
                "2023.1/4", "2023.2/4", "2023.3/4", "2023.4/4",
                "2024.1/4", "2024.2/4", "2024.3/4", "2024.4/4",
                "2025.1/4"
            ]
            
            template_data = {
                'year': year,
                'quarter': quarter,
                'data': grdp_data.get('data', {'yearly': {}, 'quarterly': {}}),
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
        return html_content, None
        
    except Exception as e:
        import traceback
        error_msg = f"개별 통계표 생성 오류 ({stat_config.get('name', 'unknown')}): {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return None, error_msg


@app.route('/')
def index():
    """메인 대시보드 페이지"""
    return render_template('dashboard.html', reports=REPORT_ORDER, regional_reports=REGIONAL_REPORTS)


def detect_file_type(filepath: str) -> str:
    """엑셀 파일 유형 자동 감지 (기초자료 수집표 vs 분석표)"""
    try:
        xl = pd.ExcelFile(filepath)
        sheet_names = xl.sheet_names
        
        # 기초자료 수집표 특징: '광공업생산', '서비스업생산', '분기 GRDP' 등의 시트
        raw_indicators = ['광공업생산', '서비스업생산', '고용률', '분기 GRDP']
        raw_count = sum(1 for s in raw_indicators if s in sheet_names)
        
        # 분석표 특징: 'A 분석', 'B 분석' 등의 시트
        analysis_indicators = ['A 분석', 'B 분석', 'C 분석', 'D(고용률)분석']
        analysis_count = sum(1 for s in analysis_indicators if s in sheet_names)
        
        if raw_count >= 2:
            return 'raw'  # 기초자료 수집표
        elif analysis_count >= 2:
            return 'analysis'  # 분석표
        else:
            # 파일명으로 추정
            filename = Path(filepath).stem.lower()
            if '기초' in filename or '수집' in filename:
                return 'raw'
            elif '분석' in filename:
                return 'analysis'
            return 'analysis'  # 기본값
    except Exception as e:
        print(f"[경고] 파일 유형 감지 실패: {e}")
        return 'analysis'


@app.route('/api/upload', methods=['POST'])
def upload_excel():
    """엑셀 파일 업로드 (기초자료 수집표 또는 분석표)"""
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
    
    # 파일 유형 자동 감지
    file_type = detect_file_type(str(filepath))
    
    # 파일 유형에 따른 처리
    analysis_path = str(filepath)
    grdp_data = None
    conversion_info = None
    
    if file_type == 'raw':
        # 기초자료 수집표인 경우: 분석표로 변환 + GRDP 추출
        print(f"[업로드] 기초자료 수집표 감지 → 분석표 변환 시작")
        try:
            converter = DataConverter(str(filepath))
            
            # 분석표 생성 (수식 보존 - 템플릿 복사 후 집계 시트 데이터 교체)
            analysis_output = str(UPLOAD_FOLDER / f"분석표_{converter.year}년_{converter.quarter}분기_자동생성.xlsx")
            analysis_path = converter.convert_all(analysis_output)
            
            download_filename = Path(analysis_path).name
            
            # GRDP 데이터 추출
            grdp_data = converter.extract_grdp_data()
            
            # GRDP JSON 저장
            grdp_json_path = TEMPLATES_DIR / 'grdp_extracted.json'
            with open(grdp_json_path, 'w', encoding='utf-8') as f:
                json.dump(grdp_data, f, ensure_ascii=False, indent=2)
            
            # 다운로드용 파일 경로 세션에 저장 (수식 보존 버전)
            session['download_analysis_path'] = analysis_path
            
            conversion_info = {
                'original_file': filename,
                'analysis_file': download_filename,
                'grdp_extracted': True,
                'national_growth_rate': grdp_data['national_summary']['growth_rate'],
                'top_region': grdp_data['top_region']['name'],
                'top_region_growth': grdp_data['top_region']['growth_rate']
            }
            
            print(f"[업로드] 변환 완료:")
            print(f"  - 분석표 (수식 보존): {download_filename}")
            print(f"[업로드] GRDP 추출 - 전국: {grdp_data['national_summary']['growth_rate']}%, 1위: {grdp_data['top_region']['name']}")
            
        except Exception as e:
            import traceback
            print(f"[오류] 기초자료 변환 실패: {e}")
            traceback.print_exc()
            # 변환 실패 시 원본 파일 사용
            analysis_path = str(filepath)
    
    # 연도/분기 추출
    year, quarter = extract_year_quarter_from_excel(analysis_path)
    
    # 세션에 저장
    session['excel_path'] = analysis_path
    session['raw_excel_path'] = str(filepath) if file_type == 'raw' else None
    session['year'] = year
    session['quarter'] = quarter
    session['file_type'] = file_type
    
    # GRDP 데이터 세션 저장
    if grdp_data:
        session['grdp_data'] = grdp_data
    
    return jsonify({
        'success': True,
        'filename': filename,
        'file_type': file_type,
        'year': year,
        'quarter': quarter,
        'reports': REPORT_ORDER,
        'regional_reports': REGIONAL_REPORTS,
        'conversion_info': conversion_info
    })


@app.route('/api/download-analysis', methods=['GET'])
def download_analysis():
    """생성된 분석표 다운로드 (수식 유지 버전)"""
    from flask import send_file
    
    # 수식 유지 버전 다운로드 (다운로드 전용 경로 사용)
    excel_path = session.get('download_analysis_path') or session.get('excel_path')
    
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '분석표 파일을 찾을 수 없습니다.'}), 404
    
    # 파일명 추출
    filename = Path(excel_path).name
    
    return send_file(
        excel_path,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


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
    
    # 기초자료 경로 가져오기
    raw_excel_path = session.get('raw_excel_path')
    
    # HTML 생성
    html_content, error, missing_fields = generate_report_html(
        excel_path, report_config, year, quarter, custom_data, raw_excel_path
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


@app.route('/api/generate-summary-preview', methods=['POST'])
def generate_summary_preview():
    """요약 보고서 미리보기 생성 (표지, 목차, 인포그래픽 등)"""
    data = request.get_json()
    report_id = data.get('report_id')
    year = data.get('year', session.get('year', 2025))
    quarter = data.get('quarter', session.get('quarter', 2))
    custom_data = data.get('custom_data', {})
    contact_info_input = data.get('contact_info', {})
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    # 요약 보고서 설정 찾기
    report_config = next((r for r in SUMMARY_REPORTS if r['id'] == report_id), None)
    if not report_config:
        return jsonify({'success': False, 'error': f'요약 보고서를 찾을 수 없습니다: {report_id}'})
    
    try:
        template_name = report_config['template']
        generator_name = report_config.get('generator')
        
        # 기본 report_info
        report_data = {
            'report_info': {
                'year': year,
                'quarter': quarter,
                'organization': '통계청',
                'department': '경제통계심의관'
            }
        }
        
        # Generator가 있는 경우 (인포그래픽 등)
        if generator_name:
            module = load_generator_module(generator_name)
            if module and hasattr(module, 'generate_report_data'):
                generated_data = module.generate_report_data(excel_path)
                report_data.update(generated_data)
        
        # ===== 템플릿별 기본 데이터 제공 =====
        
        # 목차 (toc)
        if report_id == 'toc':
            report_data['sections'] = {
                'summary': {'page': 1},
                'sector': {
                    'page': 5,
                    'entries': [
                        {'number': 1, 'name': '광공업생산', 'page': 5},
                        {'number': 2, 'name': '서비스업생산', 'page': 7},
                        {'number': 3, 'name': '소비동향', 'page': 9},
                        {'number': 4, 'name': '건설동향', 'page': 11},
                        {'number': 5, 'name': '수출', 'page': 13},
                        {'number': 6, 'name': '수입', 'page': 15},
                        {'number': 7, 'name': '물가동향', 'page': 17},
                        {'number': 8, 'name': '고용률', 'page': 19},
                        {'number': 9, 'name': '실업률', 'page': 21},
                        {'number': 10, 'name': '국내인구이동', 'page': 23},
                    ]
                },
                'region': {
                    'page': 25,
                    'entries': [
                        {'number': 1, 'name': '서울특별시', 'page': 25},
                        {'number': 2, 'name': '부산광역시', 'page': 27},
                        {'number': 3, 'name': '대구광역시', 'page': 29},
                        {'number': 4, 'name': '인천광역시', 'page': 31},
                        {'number': 5, 'name': '광주광역시', 'page': 33},
                        {'number': 6, 'name': '대전광역시', 'page': 35},
                        {'number': 7, 'name': '울산광역시', 'page': 37},
                        {'number': 8, 'name': '세종특별자치시', 'page': 39},
                        {'number': 9, 'name': '경기도', 'page': 41},
                        {'number': 10, 'name': '강원특별자치도', 'page': 43},
                        {'number': 11, 'name': '충청북도', 'page': 45},
                        {'number': 12, 'name': '충청남도', 'page': 47},
                        {'number': 13, 'name': '전북특별자치도', 'page': 49},
                        {'number': 14, 'name': '전라남도', 'page': 51},
                        {'number': 15, 'name': '경상북도', 'page': 53},
                        {'number': 16, 'name': '경상남도', 'page': 55},
                        {'number': 17, 'name': '제주특별자치도', 'page': 57},
                    ]
                },
                'reference': {'name': '분기 지역내총생산(GRDP)', 'page': 59},
                'statistics': {'page': 61},
                'appendix': {'page': 75}
            }
        
        # 일러두기 (guide)
        elif report_id == 'guide':
            report_data['intro'] = {
                'background': '지역경제동향은 시·도별 경제 현황을 생산, 소비, 건설, 수출입, 물가, 고용, 인구 등의 주요 경제지표를 통하여 분석한 자료입니다.',
                'purpose': '지역경제의 동향 파악과 지역개발정책 수립 및 평가의 기초자료로 활용하고자 작성합니다.'
            }
            report_data['content'] = {
                'description': f'본 보도자료는 {year}년 {quarter}/4분기 시·도별 지역경제동향을 수록하였습니다.',
                'indicator_note': '수록 지표는 총 7개 부문으로 다음과 같습니다.',
                'indicators': [
                    {'type': '생산', 'stat_items': ['광공업생산지수', '서비스업생산지수']},
                    {'type': '소비', 'stat_items': ['소매판매액지수']},
                    {'type': '건설', 'stat_items': ['건설수주액']},
                    {'type': '수출입', 'stat_items': ['수출액', '수입액']},
                    {'type': '물가', 'stat_items': ['소비자물가지수']},
                    {'type': '고용', 'stat_items': ['고용률', '실업률']},
                    {'type': '인구', 'stat_items': ['국내인구이동']}
                ]
            }
            report_data['contacts'] = [
                {'category': '생산', 'statistics_name': '광공업생산지수', 'department': '광업제조업동향과', 'phone': '042-481-2183'},
                {'category': '생산', 'statistics_name': '서비스업생산지수', 'department': '서비스업동향과', 'phone': '042-481-2196'},
                {'category': '소비', 'statistics_name': '소매판매액지수', 'department': '서비스업동향과', 'phone': '042-481-2199'},
                {'category': '건설', 'statistics_name': '건설수주액', 'department': '건설동향과', 'phone': '042-481-2556'},
                {'category': '수출입', 'statistics_name': '수출입액', 'department': '관세청', 'phone': '-'},
                {'category': '물가', 'statistics_name': '소비자물가지수', 'department': '물가동향과', 'phone': '042-481-2532'},
                {'category': '고용', 'statistics_name': '고용률, 실업률', 'department': '고용통계과', 'phone': '042-481-2264'},
                {'category': '인구', 'statistics_name': '국내인구이동', 'department': '인구동향과', 'phone': '042-481-2252'}
            ]
            report_data['references'] = [
                {'content': '본 자료는 통계청 홈페이지(http://kostat.go.kr)에서 확인하실 수 있습니다.'},
                {'content': '관련 통계표는 KOSIS(국가통계포털, http://kosis.kr)에서 이용하실 수 있습니다.'}
            ]
            report_data['notes'] = [
                '자료에 수록된 값은 잠정치이므로 추후 수정될 수 있습니다.'
            ]
        
        # 요약-지역경제동향 (summary_overview)
        elif report_id == 'summary_overview':
            report_data['summary'] = _get_summary_overview_data(excel_path, year, quarter)
            report_data['table_data'] = _get_summary_table_data(excel_path)
            report_data['page_number'] = 1
        
        # 요약-생산 (summary_production)
        elif report_id == 'summary_production':
            report_data.update(_get_production_summary_data(excel_path, year, quarter))
            report_data['page_number'] = 2
        
        # 요약-소비건설 (summary_consumption)
        elif report_id == 'summary_consumption':
            report_data.update(_get_consumption_construction_data(excel_path, year, quarter))
            report_data['page_number'] = 3
        
        # 요약-수출물가 (summary_trade_price)
        elif report_id == 'summary_trade_price':
            report_data.update(_get_trade_price_data(excel_path, year, quarter))
            report_data['page_number'] = 4
        
        # 요약-고용인구 (summary_employment)
        elif report_id == 'summary_employment':
            report_data.update(_get_employment_population_data(excel_path, year, quarter))
            report_data['page_number'] = 5
        
        # 담당자 정보 추가
        report_data['release_info'] = {
            'release_datetime': contact_info_input.get('release_datetime', '2025. 8. 12.(화) 12:00'),
            'distribution_datetime': contact_info_input.get('distribution_datetime', '2025. 8. 12.(화) 08:30')
        }
        report_data['contact_info'] = {
            'department': contact_info_input.get('department', '통계청 경제통계국'),
            'division': contact_info_input.get('division', '소득통계과'),
            'manager_title': contact_info_input.get('manager_title', '과 장'),
            'manager_name': contact_info_input.get('manager_name', '정선경'),
            'manager_phone': contact_info_input.get('manager_phone', '042-481-2206'),
            'staff_title': contact_info_input.get('staff_title', '사무관'),
            'staff_name': contact_info_input.get('staff_name', '윤민희'),
            'staff_phone': contact_info_input.get('staff_phone', '042-481-2226')
        }
        
        # 커스텀 데이터 병합
        if custom_data:
            for key, value in custom_data.items():
                report_data[key] = value
        
        # 템플릿 렌더링
        template_path = TEMPLATES_DIR / template_name
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        html_content = template.render(**report_data)
        
        return jsonify({
            'success': True,
            'html': html_content,
            'missing_fields': [],
            'report_id': report_id,
            'report_name': report_config['name']
        })
        
    except Exception as e:
        import traceback
        error_msg = f"요약 보고서 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': error_msg})


def _get_summary_overview_data(excel_path, year, quarter):
    """요약-지역경제동향 데이터 추출"""
    try:
        xl = pd.ExcelFile(excel_path)
        
        # 광공업 데이터 (A 분석)
        mining_data = _extract_sector_summary(xl, 'A 분석')
        # 서비스업 데이터 (B 분석)
        service_data = _extract_sector_summary(xl, 'B 분석')
        # 소비 데이터 (C 분석)
        consumption_data = _extract_sector_summary(xl, 'C 분석')
        # 수출 데이터 (G 분석)
        export_data = _extract_sector_summary(xl, 'G 분석')
        # 물가 데이터 (E 분석)
        price_data = _extract_sector_summary(xl, 'E(품목성질물가)분석')
        # 고용 데이터 (D 분석)
        employment_data = _extract_sector_summary(xl, 'D(고용률)분석')
        
        return {
            'production': {
                'mining': mining_data,
                'service': service_data
            },
            'consumption': consumption_data,
            'exports': export_data,
            'price': price_data,
            'employment': employment_data
        }
    except Exception as e:
        print(f"요약 데이터 추출 오류: {e}")
        return _get_default_summary_data()


def _extract_sector_summary(xl, sheet_name):
    """시트에서 요약 데이터 추출"""
    try:
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        increase_regions = []
        decrease_regions = []
        nationwide = 0.0
        
        for i, row in df.iterrows():
            try:
                region = str(row[2]).strip()
                if str(row[3]) == '0':
                    value = float(row[19]) if not pd.isna(row[19]) else 0.0
                    if region == '전국':
                        nationwide = value
                    elif region in regions:
                        if value >= 0:
                            increase_regions.append({'name': region, 'value': value})
                        else:
                            decrease_regions.append({'name': region, 'value': value})
            except:
                continue
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': round(nationwide, 1),
            'increase_regions': increase_regions[:3],
            'decrease_regions': decrease_regions[:3],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3],
            'below_regions': decrease_regions[:3],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions)
        }
    except Exception as e:
        print(f"{sheet_name} 데이터 추출 오류: {e}")
        return _get_default_sector_summary()


def _get_default_summary_data():
    """기본 요약 데이터"""
    default_sector = _get_default_sector_summary()
    return {
        'production': {
            'mining': default_sector,
            'service': default_sector
        },
        'consumption': default_sector,
        'exports': default_sector,
        'price': default_sector,
        'employment': default_sector
    }


def _get_default_sector_summary():
    """기본 부문 요약 데이터"""
    return {
        'nationwide': 0.0,
        'increase_regions': [{'name': '-', 'value': 0.0}],
        'decrease_regions': [{'name': '-', 'value': 0.0}],
        'increase_count': 0,
        'decrease_count': 0,
        'above_regions': [{'name': '-', 'value': 0.0}],
        'below_regions': [{'name': '-', 'value': 0.0}],
        'above_count': 0,
        'below_count': 0
    }


def _get_summary_table_data(excel_path):
    """요약 테이블 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        region_groups = [
            {'name': '수도권', 'regions': ['서울', '인천', '경기']},
            {'name': '충청권', 'regions': ['대전', '세종', '충북', '충남']},
            {'name': '호남권', 'regions': ['광주', '전북', '전남']},
            {'name': '영남권', 'regions': ['부산', '대구', '울산', '경북', '경남']},
            {'name': '기타', 'regions': ['강원', '제주']}
        ]
        
        nationwide_data = {
            'mining_production': 0.0, 'service_production': 0.0, 'retail_sales': 0.0,
            'exports': 0.0, 'price': 0.0, 'employment': 0.0
        }
        
        # 각 시트에서 전국 데이터 추출
        sheet_mapping = {
            'A 분석': 'mining_production',
            'B 분석': 'service_production',
            'C 분석': 'retail_sales',
            'G 분석': 'exports',
            'E(품목성질물가)분석': 'price',
            'D(고용률)분석': 'employment'
        }
        
        for sheet_name, key in sheet_mapping.items():
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
                for i, row in df.iterrows():
                    if str(row[2]).strip() == '전국' and str(row[3]) == '0':
                        nationwide_data[key] = round(float(row[19]), 1) if not pd.isna(row[19]) else 0.0
                        break
            except:
                continue
        
        # 지역 그룹별 데이터 생성 (빈 데이터로 초기화)
        for group in region_groups:
            group['regions'] = [{'name': r, 'mining_production': 0.0, 'service_production': 0.0,
                                 'retail_sales': 0.0, 'exports': 0.0, 'price': 0.0, 'employment': 0.0}
                               for r in group['regions']]
        
        return {
            'nationwide': nationwide_data,
            'region_groups': region_groups
        }
    except Exception as e:
        print(f"요약 테이블 데이터 오류: {e}")
        return {'nationwide': {'mining_production': 0.0, 'service_production': 0.0, 'retail_sales': 0.0,
                              'exports': 0.0, 'price': 0.0, 'employment': 0.0}, 'region_groups': []}


def _get_production_summary_data(excel_path, year, quarter):
    """요약-생산 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        mining = _extract_chart_data(xl, 'A 분석')
        service = _extract_chart_data(xl, 'B 분석')
        
        return {
            'mining_production': mining,
            'service_production': service
        }
    except Exception as e:
        print(f"생산 요약 데이터 오류: {e}")
        return {
            'mining_production': _get_default_chart_data(),
            'service_production': _get_default_chart_data()
        }


def _get_consumption_construction_data(excel_path, year, quarter):
    """요약-소비건설 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        retail = _extract_chart_data(xl, 'C 분석')
        
        # 건설 데이터
        construction = {
            'nationwide': {'amount': '0', 'change': 0.0},
            'increase_regions': [],
            'decrease_regions': [],
            'increase_count': 0,
            'decrease_count': 0,
            'chart_data': []
        }
        try:
            df = pd.read_excel(xl, sheet_name="F'분석", header=None)
            for i, row in df.iterrows():
                if str(row[2]).strip() == '전국' and str(row[3]) == '0':
                    construction['nationwide']['change'] = round(float(row[19]), 1) if not pd.isna(row[19]) else 0.0
                    break
        except:
            pass
        
        return {
            'retail_sales': retail,
            'construction': construction
        }
    except Exception as e:
        print(f"소비건설 요약 데이터 오류: {e}")
        return {
            'retail_sales': _get_default_chart_data(),
            'construction': {'nationwide': {'amount': '0', 'change': 0.0}, 'increase_regions': [], 
                           'decrease_regions': [], 'increase_count': 0, 'decrease_count': 0, 'chart_data': []}
        }


def _get_trade_price_data(excel_path, year, quarter):
    """요약-수출물가 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        exports = _extract_chart_data(xl, 'G 분석', is_trade=True)
        price = _extract_chart_data(xl, 'E(품목성질물가)분석')
        
        return {
            'exports': exports,
            'price': price
        }
    except Exception as e:
        print(f"수출 데이터 추출 오류: {e}")
        return {
            'exports': _get_default_trade_data(),
            'price': _get_default_chart_data()
        }


def _get_employment_population_data(excel_path, year, quarter):
    """요약-고용인구 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        employment = _extract_chart_data(xl, 'D(고용률)분석', is_employment=True)
        
        # 인구이동 데이터
        population = {
            'inflow_regions': [],
            'outflow_regions': [],
            'inflow_count': 0,
            'outflow_count': 0,
            'chart_data': []
        }
        try:
            df = pd.read_excel(xl, sheet_name='I(순인구이동)집계', header=None)
            regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                       '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
            
            for i, row in df.iterrows():
                region = str(row[2]).strip() if not pd.isna(row[2]) else ''
                if region in regions:
                    try:
                        value = int(float(row[19])) if not pd.isna(row[19]) else 0
                        if value > 0:
                            population['inflow_regions'].append({'name': region, 'value': value})
                        else:
                            population['outflow_regions'].append({'name': region, 'value': abs(value)})
                    except:
                        continue
            
            population['inflow_regions'].sort(key=lambda x: x['value'], reverse=True)
            population['outflow_regions'].sort(key=lambda x: x['value'], reverse=True)
            population['inflow_count'] = len(population['inflow_regions'])
            population['outflow_count'] = len(population['outflow_regions'])
        except Exception as e:
            print(f"인구이동 데이터 오류: {e}")
        
        return {
            'employment': employment,
            'population': population
        }
    except Exception as e:
        print(f"고용인구 요약 데이터 오류: {e}")
        return {
            'employment': _get_default_employment_data(),
            'population': {'inflow_regions': [], 'outflow_regions': [], 'inflow_count': 0, 
                          'outflow_count': 0, 'chart_data': []}
        }


def _extract_chart_data(xl, sheet_name, is_trade=False, is_employment=False):
    """차트용 데이터 추출"""
    try:
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        nationwide = {'index': 100.0, 'change': 0.0, 'rate': 60.0, 'amount': 0}
        increase_regions = []
        decrease_regions = []
        chart_data = []
        
        for i, row in df.iterrows():
            try:
                region = str(row[2]).strip()
                if str(row[3]) == '0':
                    index_val = float(row[18]) if not pd.isna(row[18]) else 100.0
                    change_val = float(row[19]) if not pd.isna(row[19]) else 0.0
                    
                    if region == '전국':
                        nationwide['index'] = round(index_val, 1)
                        nationwide['change'] = round(change_val, 1)
                        nationwide['rate'] = round(index_val, 1)
                        if is_trade:
                            nationwide['amount'] = round(index_val, 0)
                    elif region in regions:
                        data = {
                            'name': region, 'value': round(change_val, 1),
                            'index': round(index_val, 1), 'change': round(change_val, 1),
                            'rate': round(index_val, 1)
                        }
                        if is_trade:
                            data['amount'] = round(index_val, 0)
                            data['amount_normalized'] = min(100, max(0, index_val / 6))
                        if change_val >= 0:
                            increase_regions.append(data)
                        else:
                            decrease_regions.append(data)
                        chart_data.append(data)
            except:
                continue
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': nationwide,
            'increase_regions': increase_regions[:3],
            'decrease_regions': decrease_regions[:3],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3],
            'below_regions': decrease_regions[:3],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions),
            'chart_data': chart_data[:18]
        }
    except Exception as e:
        print(f"{sheet_name} 차트 데이터 오류: {e}")
        if is_trade:
            return _get_default_trade_data()
        elif is_employment:
            return _get_default_employment_data()
        return _get_default_chart_data()


def _get_default_chart_data():
    """기본 차트 데이터"""
    return {
        'nationwide': {'index': 100.0, 'change': 0.0},
        'increase_regions': [{'name': '-', 'value': 0.0, 'index': 100.0, 'change': 0.0}],
        'decrease_regions': [{'name': '-', 'value': 0.0, 'index': 100.0, 'change': 0.0}],
        'increase_count': 0, 'decrease_count': 0,
        'above_regions': [{'name': '-', 'value': 0.0}],
        'below_regions': [{'name': '-', 'value': 0.0}],
        'above_count': 0, 'below_count': 0,
        'chart_data': []
    }


def _get_default_trade_data():
    """기본 수출입 데이터"""
    return {
        'nationwide': {'amount': 0, 'change': 0.0},
        'increase_regions': [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
        'decrease_regions': [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
        'increase_count': 0, 'decrease_count': 0,
        'chart_data': []
    }


def _get_default_employment_data():
    """기본 고용 데이터"""
    return {
        'nationwide': {'rate': 60.0, 'change': 0.0},
        'increase_regions': [{'name': '-', 'value': 0.0, 'rate': 60.0, 'change': 0.0}],
        'decrease_regions': [{'name': '-', 'value': 0.0, 'rate': 60.0, 'change': 0.0}],
        'increase_count': 0, 'decrease_count': 0,
        'chart_data': []
    }


@app.route('/api/generate-regional-preview', methods=['POST'])
def generate_regional_preview():
    """시도별 보고서 미리보기 생성"""
    data = request.get_json()
    region_id = data.get('region_id')
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    # 지역 정보 찾기
    region_config = next((r for r in REGIONAL_REPORTS if r['id'] == region_id), None)
    if not region_config:
        return jsonify({'success': False, 'error': f'지역을 찾을 수 없습니다: {region_id}'})
    
    # 참고_GRDP 여부 확인
    is_reference = region_config.get('is_reference', False)
    
    # HTML 생성
    html_content, error = generate_regional_report_html(excel_path, region_config['name'], is_reference)
    
    if error:
        return jsonify({'success': False, 'error': error})
    
    return jsonify({
        'success': True,
        'html': html_content,
        'region_id': region_id,
        'region_name': region_config['name'],
        'full_name': region_config['full_name']
    })


@app.route('/api/generate-all-regional', methods=['POST'])
def generate_all_regional_reports():
    """시도별 보고서 전체 생성"""
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    generated_reports = []
    errors = []
    
    # 출력 디렉토리 생성
    output_dir = TEMPLATES_DIR / '시도별_output'
    output_dir.mkdir(exist_ok=True)
    
    for region_config in REGIONAL_REPORTS:
        html_content, error = generate_regional_report_html(excel_path, region_config['name'])
        
        if error:
            errors.append({'region_id': region_config['id'], 'error': error})
        else:
            # 파일 저장
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


@app.route('/api/generate-statistics-preview', methods=['POST'])
def generate_statistics_preview():
    """개별 통계표 보고서 미리보기 생성"""
    data = request.get_json()
    stat_id = data.get('stat_id')
    year = data.get('year', session.get('year', 2025))
    quarter = data.get('quarter', session.get('quarter', 2))
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    # 통계표 설정 찾기
    stat_config = next((s for s in STATISTICS_REPORTS if s['id'] == stat_id), None)
    if not stat_config:
        return jsonify({'success': False, 'error': f'통계표를 찾을 수 없습니다: {stat_id}'})
    
    # HTML 생성
    # 기초자료 경로 가져오기
    raw_excel_path = session.get('raw_excel_path')
    html_content, error = generate_individual_statistics_html(excel_path, stat_config, year, quarter, raw_excel_path)
    
    if error:
        return jsonify({'success': False, 'error': error})
    
    return jsonify({
        'success': True,
        'html': html_content,
        'stat_id': stat_id,
        'report_name': stat_config['name']
    })


@app.route('/api/generate-statistics-full-preview', methods=['POST'])
def generate_statistics_full_preview():
    """통계표 전체 보고서 미리보기 생성 (기존 방식)"""
    data = request.get_json()
    year = data.get('year', session.get('year', 2025))
    quarter = data.get('quarter', session.get('quarter', 2))
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    # HTML 생성
    # 기초자료 경로 가져오기
    raw_excel_path = session.get('raw_excel_path')
    html_content, error = generate_statistics_report_html(excel_path, year, quarter, raw_excel_path)
    
    if error:
        return jsonify({'success': False, 'error': error})
    
    return jsonify({
        'success': True,
        'html': html_content,
        'report_name': '통계표 (전체)'
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
        # 기초자료 경로 가져오기
        raw_excel_path = session.get('raw_excel_path')
        
        html_content, error, _ = generate_report_html(
            excel_path, report_config, year, quarter, custom_data, raw_excel_path
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
    return jsonify({'reports': REPORT_ORDER, 'regional_reports': REGIONAL_REPORTS})


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


@app.route('/preview/infographic')
def preview_infographic():
    """인포그래픽 미리보기 (직접 접근용)"""
    from flask import send_file
    output_path = TEMPLATES_DIR / '인포그래픽_output.html'
    if output_path.exists():
        return send_file(output_path)
    return "인포그래픽이 아직 생성되지 않았습니다.", 404


@app.route('/api/export-final', methods=['POST'])
def export_final_document():
    """모든 보고서를 하나의 HTML 문서로 합치기"""
    try:
        data = request.get_json()
        pages = data.get('pages', [])  # [{html: '...', title: '...'}, ...]
        year = data.get('year', session.get('year', 2025))
        quarter = data.get('quarter', session.get('quarter', 2))
        
        if not pages:
            return jsonify({'success': False, 'error': '페이지 데이터가 없습니다.'})
        
        # 최종 HTML 문서 생성
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
        
        /* 페이지 컨테이너 */
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
        
        /* 페이지 내용 */
        .page-content {{
            width: 100%;
            min-height: calc(297mm - 40mm);
        }}
        
        .page-content > * {{
            max-width: 100%;
        }}
        
        /* 페이지 번호 */
        .page-number {{
            position: absolute;
            bottom: 10mm;
            left: 0;
            right: 0;
            text-align: center;
            font-size: 10pt;
            color: #666;
        }}
        
        /* iframe 내용 스타일 오버라이드 */
        .page-content iframe {{
            border: none;
            width: 100%;
            min-height: 250mm;
        }}
        
        /* 인쇄용 스타일 */
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
        
        # 각 페이지 추가
        for idx, page in enumerate(pages, 1):
            page_html = page.get('html', '')
            page_title = page.get('title', f'페이지 {idx}')
            
            # HTML에서 body 내용만 추출
            body_content = page_html
            if '<body' in page_html:
                start = page_html.find('<body')
                start = page_html.find('>', start) + 1
                end = page_html.find('</body>')
                if end > start:
                    body_content = page_html[start:end]
            
            # 스타일 추출 (head에서)
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
        
        # 파일로 저장
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


@app.route('/uploads/<filename>')
def download_file(filename):
    """업로드된 파일 다운로드"""
    from flask import send_from_directory
    return send_from_directory(str(UPLOAD_FOLDER), filename, as_attachment=True)


@app.route('/view/<filename>')
def view_file(filename):
    """파일 직접 보기 (다운로드 없이)"""
    from flask import send_from_directory
    return send_from_directory(str(UPLOAD_FOLDER), filename, as_attachment=False)


@app.route('/api/export-hwpx', methods=['POST'])
def export_hwpx_document():
    """모든 보고서를 HWPX (한글) 파일로 내보내기"""
    try:
        from hwpx_generator import create_hwpx_from_html
        
        data = request.get_json()
        pages = data.get('pages', [])
        year = data.get('year', session.get('year', 2025))
        quarter = data.get('quarter', session.get('quarter', 2))
        
        if not pages:
            return jsonify({'success': False, 'error': '페이지 데이터가 없습니다.'})
        
        # HWPX 파일명 생성
        output_filename = f'지역경제동향_{year}년_{quarter}분기.hwpx'
        output_path = str(UPLOAD_FOLDER / output_filename)
        
        # HWPX 생성
        result = create_hwpx_from_html(pages, year, quarter, output_path)
        
        if result['success']:
            return jsonify({
                'success': True,
                'filename': output_filename,
                'download_url': f'/uploads/{output_filename}',
                'view_url': f'/view/{output_filename}',
                'total_pages': result.get('pages_count', len(pages)),
                'images_count': result.get('images_count', 0)
            })
        else:
            return jsonify({'success': False, 'error': result.get('error', '알 수 없는 오류')})
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/render-chart-image', methods=['POST'])
def render_chart_image():
    """차트/인포그래픽을 이미지로 렌더링 (클라이언트측에서 canvas.toDataURL 호출용)"""
    try:
        data = request.get_json()
        image_data = data.get('image_data', '')  # base64 data URL
        filename = data.get('filename', 'chart.png')
        
        if not image_data:
            return jsonify({'success': False, 'error': '이미지 데이터가 없습니다.'})
        
        # data:image/png;base64,... 형식 처리
        import re
        import base64
        
        match = re.match(r'data:([^;]+);base64,(.+)', image_data)
        if match:
            mimetype = match.group(1)
            img_data = base64.b64decode(match.group(2))
            
            # 파일 저장
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


if __name__ == '__main__':
    print("=" * 50)
    print("지역경제동향 보고서 생성 시스템")
    print("=" * 50)
    print(f"서버 시작: http://localhost:5050")
    print("=" * 50)
    app.run(debug=True, host='0.0.0.0', port=5050)

