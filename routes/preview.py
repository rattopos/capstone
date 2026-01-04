# -*- coding: utf-8 -*-
"""
미리보기 API 라우트
"""

from pathlib import Path

from flask import Blueprint, request, jsonify, session
from jinja2 import Template

from config.settings import TEMPLATES_DIR
from config.reports import (
    REPORT_ORDER, SUMMARY_REPORTS, REGIONAL_REPORTS, STATISTICS_REPORTS,
    PAGE_CONFIG, TOC_SECTOR_ITEMS, TOC_REGION_ITEMS
)
from utils.excel_utils import load_generator_module
from services.report_generator import (
    generate_report_html,
    generate_regional_report_html,
    generate_statistics_report_html,
    generate_individual_statistics_html
)
from services.summary_data import (
    get_summary_overview_data,
    get_summary_table_data,
    get_production_summary_data,
    get_consumption_construction_data,
    get_trade_price_data,
    get_employment_population_data
)

preview_bp = Blueprint('preview', __name__, url_prefix='/api')


@preview_bp.route('/generate-preview', methods=['POST'])
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
    
    report_config = next((r for r in REPORT_ORDER if r['id'] == report_id), None)
    if not report_config:
        return jsonify({'success': False, 'error': f'보도자료를 찾을 수 없습니다: {report_id}'})
    
    raw_excel_path = session.get('raw_excel_path')
    
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


@preview_bp.route('/generate-summary-preview', methods=['POST'])
def generate_summary_preview():
    """요약 보도자료 미리보기 생성 (표지, 목차, 인포그래픽 등)"""
    data = request.get_json()
    report_id = data.get('report_id')
    year = data.get('year', session.get('year', 2025))
    quarter = data.get('quarter', session.get('quarter', 2))
    custom_data = data.get('custom_data', {})
    contact_info_input = data.get('contact_info', {})
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    report_config = next((r for r in SUMMARY_REPORTS if r['id'] == report_id), None)
    if not report_config:
        return jsonify({'success': False, 'error': f'요약 보도자료를 찾을 수 없습니다: {report_id}'})
    
    try:
        template_name = report_config['template']
        generator_name = report_config.get('generator')
        
        report_data = {
            'report_info': {
                'year': year,
                'quarter': quarter,
                'organization': '통계청',
                'department': '경제통계심의관'
            }
        }
        
        if generator_name:
            module = load_generator_module(generator_name)
            if module and hasattr(module, 'generate_report_data'):
                generated_data = module.generate_report_data(excel_path)
                report_data.update(generated_data)
        
        # 템플릿별 기본 데이터 제공
        if report_id == 'toc':
            report_data['sections'] = _get_toc_sections()
        
        elif report_id == 'guide':
            report_data.update(_get_guide_data(year, quarter, contact_info_input))
        
        elif report_id == 'summary_overview':
            report_data['summary'] = get_summary_overview_data(excel_path, year, quarter)
            report_data['table_data'] = get_summary_table_data(excel_path)
            report_data['page_number'] = 1
        
        elif report_id == 'summary_production':
            report_data.update(get_production_summary_data(excel_path, year, quarter))
            report_data['page_number'] = 2
        
        elif report_id == 'summary_consumption':
            report_data.update(get_consumption_construction_data(excel_path, year, quarter))
            report_data['page_number'] = 3
        
        elif report_id == 'summary_trade_price':
            report_data.update(get_trade_price_data(excel_path, year, quarter))
            report_data['page_number'] = 4
        
        elif report_id == 'summary_employment':
            report_data.update(get_employment_population_data(excel_path, year, quarter))
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
        
        if custom_data:
            for key, value in custom_data.items():
                report_data[key] = value
        
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
        error_msg = f"요약 보도자료 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': error_msg})


@preview_bp.route('/generate-regional-preview', methods=['POST'])
def generate_regional_preview():
    """시도별 보도자료 미리보기 생성"""
    data = request.get_json()
    region_id = data.get('region_id')
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    region_config = next((r for r in REGIONAL_REPORTS if r['id'] == region_id), None)
    if not region_config:
        return jsonify({'success': False, 'error': f'지역을 찾을 수 없습니다: {region_id}'})
    
    is_reference = region_config.get('is_reference', False)
    
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


@preview_bp.route('/generate-statistics-preview', methods=['POST'])
def generate_statistics_preview():
    """개별 통계표 보도자료 미리보기 생성"""
    data = request.get_json()
    stat_id = data.get('stat_id')
    year = data.get('year', session.get('year', 2025))
    quarter = data.get('quarter', session.get('quarter', 2))
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    stat_config = next((s for s in STATISTICS_REPORTS if s['id'] == stat_id), None)
    if not stat_config:
        return jsonify({'success': False, 'error': f'통계표를 찾을 수 없습니다: {stat_id}'})
    
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


@preview_bp.route('/generate-statistics-full-preview', methods=['POST'])
def generate_statistics_full_preview():
    """통계표 전체 보도자료 미리보기 생성"""
    data = request.get_json()
    year = data.get('year', session.get('year', 2025))
    quarter = data.get('quarter', session.get('quarter', 2))
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    raw_excel_path = session.get('raw_excel_path')
    html_content, error = generate_statistics_report_html(excel_path, year, quarter, raw_excel_path)
    
    if error:
        return jsonify({'success': False, 'error': error})
    
    return jsonify({
        'success': True,
        'html': html_content,
        'report_name': '통계표 (전체)'
    })


def _get_toc_sections():
    """목차 섹션 데이터 - 페이지 단위로 동적 계산
    
    같은 항목이 여러 페이지인 경우 (1), (2) 등으로 구분
    """
    
    # 현재 페이지 번호 (요약부터 1페이지 시작)
    current_page = 1
    
    # 요약 섹션 시작 페이지
    summary_page = current_page
    summary_pages = sum(PAGE_CONFIG['summary'].values())
    current_page += summary_pages
    
    # 부문별 섹션 시작 페이지
    sector_page = current_page
    
    # 부문별 각 항목의 시작 페이지 계산 (페이지 단위)
    sector_entries = []
    sector_config = PAGE_CONFIG['sector']
    sector_order = ['manufacturing', 'service', 'consumption', 'construction', 
                    'export', 'import', 'price', 'employment', 'unemployment', 'population']
    
    # TOC_SECTOR_ITEMS 기반으로 통합 항목 처리
    item_pages_map = {}  # start_from -> [페이지번호들]
    for sector_id in sector_order:
        pages_count = sector_config.get(sector_id, 1)
        if sector_id not in item_pages_map:
            item_pages_map[sector_id] = []
        for i in range(pages_count):
            item_pages_map[sector_id].append(current_page + i)
        current_page += pages_count
    
    # 부문별 목차 항목 생성 (페이지 단위)
    entry_number = 1
    for item in TOC_SECTOR_ITEMS:
        start_from = item.get('start_from')
        pages = item_pages_map.get(start_from, [])
        
        if len(pages) == 1:
            # 1페이지짜리 항목
            sector_entries.append({
                'number': entry_number,
                'name': item['name'],
                'page': pages[0]
            })
            entry_number += 1
        else:
            # 여러 페이지인 경우 (1), (2) 등으로 구분
            for idx, page in enumerate(pages, 1):
                sector_entries.append({
                    'number': entry_number,
                    'name': f"{item['name']} ({idx})",
                    'page': page
                })
                entry_number += 1
    
    # 시도별 섹션 시작 페이지
    region_page = current_page
    
    # 시도별 목차 항목 생성 (페이지 단위)
    region_entries = []
    regional_pages = PAGE_CONFIG['regional']  # 각 시도당 페이지 수 (2)
    entry_number = 1
    
    for item in TOC_REGION_ITEMS:
        if regional_pages == 1:
            # 1페이지짜리 항목
            region_entries.append({
                'number': entry_number,
                'name': item['name'],
                'page': current_page
            })
            entry_number += 1
            current_page += 1
        else:
            # 여러 페이지인 경우 (1), (2) 등으로 구분
            for idx in range(1, regional_pages + 1):
                region_entries.append({
                    'number': entry_number,
                    'name': f"{item['name']} ({idx})",
                    'page': current_page
                })
                entry_number += 1
                current_page += 1
    
    # 참고 GRDP 페이지 (2페이지인 경우 구분)
    reference_page = current_page
    grdp_pages = PAGE_CONFIG['reference_grdp']
    reference_entries = []
    if grdp_pages > 1:
        for idx in range(1, grdp_pages + 1):
            reference_entries.append({
                'name': f'분기GRDP ({idx})',
                'page': current_page
            })
            current_page += 1
    else:
        current_page += grdp_pages
    
    # 통계표 섹션 시작 페이지
    statistics_page = current_page
    stat_config = PAGE_CONFIG['statistics']
    current_page += stat_config['toc']
    
    # 통계표 목차 항목 생성 (페이지 단위, 각 통계표 2페이지)
    statistics_entries = []
    stat_names = ['광공업생산지수', '서비스업생산지수', '소매판매액지수', '건설수주액',
                  '고용률', '실업률', '국내인구이동', '수출액', '수입액', '소비자물가지수', 'GRDP']
    entry_number = 1
    for stat_name in stat_names:
        pages_per_table = 2 if stat_name != 'GRDP' else 2  # 모든 통계표 2페이지
        for idx in range(1, pages_per_table + 1):
            statistics_entries.append({
                'number': entry_number,
                'name': f'{stat_name} ({idx})',
                'page': current_page
            })
            entry_number += 1
            current_page += 1
    
    # 부록 페이지
    appendix_page = current_page
    
    return {
        'summary': {'page': summary_page},
        'sector': {
            'page': sector_page,
            'entries': sector_entries
        },
        'region': {
            'page': region_page,
            'entries': region_entries
        },
        'reference': {
            'name': '분기GRDP', 
            'page': reference_page,
            'entries': reference_entries if reference_entries else None
        },
        'statistics': {
            'page': statistics_page,
            'entries': statistics_entries
        },
        'appendix': {'page': appendix_page}
    }


def _get_guide_data(year, quarter, contact_info=None):
    """일러두기 데이터"""
    # 관세청 담당자 정보 (contact_info에서 가져오거나 기본값 사용)
    customs_dept = '관세청 정보데이터기획담당관'
    customs_phone = '042-481-7845'
    
    if contact_info:
        customs_dept = contact_info.get('customs_department', customs_dept)
        customs_phone = contact_info.get('customs_phone', customs_phone)
    
    return {
        'intro': {
            'background': '지역경제동향은 시·도별 경제 현황을 생산, 소비, 건설, 수출입, 물가, 고용, 인구 등의 주요 경제지표를 통하여 분석한 자료입니다.',
            'purpose': '지역경제의 동향 파악과 지역개발정책 수립 및 평가의 기초자료로 활용하고자 작성합니다.'
        },
        'content': {
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
        },
        'contacts': [
            {'category': '생산', 'statistics_name': '광공업생산지수', 'department': '광업제조업동향과', 'phone': '042-481-2183'},
            {'category': '생산', 'statistics_name': '서비스업생산지수', 'department': '서비스업동향과', 'phone': '042-481-2196'},
            {'category': '소비', 'statistics_name': '소매판매액지수', 'department': '서비스업동향과', 'phone': '042-481-2199'},
            {'category': '건설', 'statistics_name': '건설수주액', 'department': '건설동향과', 'phone': '042-481-2556'},
            {'category': '수출입', 'statistics_name': '수출입액', 'department': customs_dept, 'phone': customs_phone},
            {'category': '물가', 'statistics_name': '소비자물가지수', 'department': '물가동향과', 'phone': '042-481-2532'},
            {'category': '고용', 'statistics_name': '고용률, 실업률', 'department': '고용통계과', 'phone': '042-481-2264'},
            {'category': '인구', 'statistics_name': '국내인구이동', 'department': '인구동향과', 'phone': '042-481-2252'}
        ],
        'references': [
            {'content': '본 자료는 통계청 홈페이지(http://kostat.go.kr)에서 확인하실 수 있습니다.'},
            {'content': '관련 통계표는 KOSIS(국가통계포털, http://kosis.kr)에서 이용하실 수 있습니다.'}
        ],
        'notes': [
            '자료에 수록된 값은 잠정치이므로 추후 수정될 수 있습니다.'
        ]
    }

