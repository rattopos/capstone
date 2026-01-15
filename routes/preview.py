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
    from utils.excel_utils import extract_year_quarter_from_data
    
    try:
        data = request.get_json()
        if data is None:
            print(f"[미리보기] ❌ JSON 파싱 실패: Content-Type={request.content_type}")
            return jsonify({'success': False, 'error': 'JSON 형식이 올바르지 않습니다'}), 400
    except Exception as e:
        print(f"[미리보기] ❌ JSON 파싱 예외: {e}")
        return jsonify({'success': False, 'error': f'요청 데이터 파싱 오류: {str(e)}'}), 400
    
    report_id = data.get('report_id')
    
    if not report_id:
        print(f"[미리보기] ❌ report_id 없음: data={data}")
        return jsonify({'success': False, 'error': 'report_id가 필요합니다'}), 400
    
    print(f"[미리보기] 요청: report_id={report_id}, year={data.get('year')}, quarter={data.get('quarter')}")
    
    # year/quarter가 None이거나 0이면 세션에서 가져오거나 추출
    year = data.get('year')
    quarter = data.get('quarter')
    
    # None이거나 0이면 세션에서 가져오기
    if not year or year == 0:
        year = session.get('year')
    if not quarter or quarter == 0:
        quarter = session.get('quarter')
    
    custom_data = data.get('custom_data', {})
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        print(f"[미리보기] ❌ 엑셀 파일 없음: excel_path={excel_path}")
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'}), 400
    
    # 연도/분기가 없으면 데이터에서 추출
    if not year or not quarter:
        try:
            from utils.excel_utils import get_previous_quarter
            default_year, default_quarter = get_previous_quarter()
            year, quarter = extract_year_quarter_from_data(excel_path, default_year=default_year, default_quarter=default_quarter)
            # 세션에 저장
            session['year'] = year
            session['quarter'] = quarter
            print(f"[미리보기] 데이터에서 연도/분기 추출: {year}년 {quarter}분기")
        except ValueError as e:
            print(f"[미리보기] ❌ 연도/분기 추출 실패: {e}")
            return jsonify({'success': False, 'error': f'연도/분기 정보를 추출할 수 없습니다: {str(e)}'}), 400
    
    report_config = next((r for r in REPORT_ORDER if r['id'] == report_id), None)
    if not report_config:
        return jsonify({'success': False, 'error': f'보도자료를 찾을 수 없습니다: {report_id}'}), 400
    
    # 기초자료 수집표는 사용하지 않음 (분석표만 사용)
    html_content, error, missing_fields = generate_report_html(
        excel_path, report_config, year, quarter, custom_data
    )
    
    if error:
        return jsonify({'success': False, 'error': error}), 400
    
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
    from utils.excel_utils import extract_year_quarter_from_data
    
    try:
        data = request.get_json()
        if data is None:
            print(f"[요약 미리보기] ❌ JSON 파싱 실패: Content-Type={request.content_type}")
            return jsonify({'success': False, 'error': 'JSON 형식이 올바르지 않습니다'}), 400
    except Exception as e:
        print(f"[요약 미리보기] ❌ JSON 파싱 예외: {e}")
        return jsonify({'success': False, 'error': f'요청 데이터 파싱 오류: {str(e)}'}), 400
    
    report_id = data.get('report_id')
    
    if not report_id:
        print(f"[요약 미리보기] ❌ report_id 없음: data={data}")
        return jsonify({'success': False, 'error': 'report_id가 필요합니다'}), 400
    
    print(f"[요약 미리보기] 요청: report_id={report_id}, year={data.get('year')}, quarter={data.get('quarter')}")
    
    # year/quarter가 None이거나 0이면 세션에서 가져오거나 추출
    year = data.get('year')
    quarter = data.get('quarter')
    
    # None이거나 0이면 세션에서 가져오기
    if not year or year == 0:
        year = session.get('year')
    if not quarter or quarter == 0:
        quarter = session.get('quarter')
    
    custom_data = data.get('custom_data', {})
    contact_info_input = data.get('contact_info', {})
    
    # 표지, 일러두기, 목차는 엑셀 파일 없이도 생성 가능
    static_reports = ['cover', 'guide', 'toc']
    
    excel_path = session.get('excel_path')
    if report_id not in static_reports:
        if not excel_path or not Path(excel_path).exists():
            print(f"[요약 미리보기] ❌ 엑셀 파일 없음: excel_path={excel_path}")
            return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'}), 400
        
        # 연도/분기가 없으면 데이터에서 추출
        if not year or not quarter:
            try:
                from utils.excel_utils import get_previous_quarter
                default_year, default_quarter = get_previous_quarter()
                year, quarter = extract_year_quarter_from_data(excel_path, default_year=default_year, default_quarter=default_quarter)
                session['year'] = year
                session['quarter'] = quarter
                print(f"[요약 미리보기] 데이터에서 연도/분기 추출: {year}년 {quarter}분기")
            except ValueError as e:
                print(f"[요약 미리보기] ❌ 연도/분기 추출 실패: {e}")
                return jsonify({'success': False, 'error': f'연도/분기 정보를 추출할 수 없습니다: {str(e)}'}), 400
    
    report_config = next((r for r in SUMMARY_REPORTS if r['id'] == report_id), None)
    if not report_config:
        return jsonify({'success': False, 'error': f'요약 보도자료를 찾을 수 없습니다: {report_id}'}), 400
    
    try:
        template_name = report_config['template']
        generator_name = report_config.get('generator')
        
        report_data = {
            'report_info': {
                'year': year,
                'quarter': quarter,
                'organization': '국가데이터처',
                'department': '경제동향통계심의관 지역경제동향과',
                'contact_phone': '042-481-xxxx'
            }
        }
        
        if generator_name:
            try:
                module = load_generator_module(generator_name)
                if module is None:
                    error_msg = f"Generator 모듈을 로드할 수 없습니다: {generator_name}"
                    print(f"[PREVIEW] {error_msg}")
                    return jsonify({'success': False, 'error': error_msg}), 400
                
                if hasattr(module, 'generate_report_data'):
                    try:
                        generated_data = module.generate_report_data(excel_path)
                        if generated_data:
                            report_data.update(generated_data)
                            print(f"[PREVIEW] Generator 데이터 생성 성공: {generator_name}")
                        else:
                            print(f"[PREVIEW] Generator가 빈 데이터를 반환했습니다: {generator_name}")
                    except Exception as e:
                        import traceback
                        error_msg = f"Generator 데이터 생성 오류 ({generator_name}): {str(e)}"
                        print(f"[PREVIEW] {error_msg}")
                        traceback.print_exc()
                        return jsonify({'success': False, 'error': error_msg}), 400
                else:
                    error_msg = f"Generator에 generate_report_data 함수가 없습니다: {generator_name}"
                    print(f"[PREVIEW] {error_msg}")
                    return jsonify({'success': False, 'error': error_msg}), 400
            except Exception as e:
                import traceback
                error_msg = f"Generator 모듈 로드 오류 ({generator_name}): {str(e)}"
                print(f"[PREVIEW] {error_msg}")
                traceback.print_exc()
                return jsonify({'success': False, 'error': error_msg}), 400
        
        # 템플릿 파일 존재 확인
        template_path = TEMPLATES_DIR / template_name
        if not template_path.exists():
            error_msg = f"템플릿 파일을 찾을 수 없습니다: {template_name}"
            print(f"[PREVIEW] {error_msg}")
            return jsonify({'success': False, 'error': error_msg}), 400
        
        # 템플릿별 기본 데이터 제공
        if report_id == 'cover':
            # 표지는 스키마에서 기본값 로드 (엑셀 파일 필요 없음)
            print(f"[PREVIEW] 표지 템플릿 로드")
        
        elif report_id == 'toc':
            # 목차는 고정된 HTML 템플릿 사용 (동적 계산 없음)
            print(f"[PREVIEW] 목차 템플릿 로드 (고정 페이지 번호)")
        
        elif report_id == 'guide':
            try:
                report_data.update(_get_guide_data(year, quarter, contact_info_input))
                print(f"[PREVIEW] 일러두기 데이터 생성 완료")
            except Exception as e:
                import traceback
                error_msg = f"일러두기 데이터 생성 오류: {str(e)}"
                print(f"[PREVIEW] {error_msg}")
                traceback.print_exc()
                return jsonify({'success': False, 'error': error_msg}), 400
        
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
        
        # 담당자 정보 추가 (기본값 없음 - 실무자가 직접 입력해야 함)
        report_data['release_info'] = {
            'release_datetime': contact_info_input.get('release_datetime', ''),
            'distribution_datetime': contact_info_input.get('distribution_datetime', '')
        }
        report_data['contact_info'] = {
            'department': contact_info_input.get('department', ''),
            'division': contact_info_input.get('division', ''),
            'manager_title': contact_info_input.get('manager_title', ''),
            'manager_name': contact_info_input.get('manager_name', ''),
            'manager_phone': contact_info_input.get('manager_phone', ''),
            'staff_title': contact_info_input.get('staff_title', ''),
            'staff_name': contact_info_input.get('staff_name', ''),
            'staff_phone': contact_info_input.get('staff_phone', '')
        }
        
        if custom_data:
            for key, value in custom_data.items():
                report_data[key] = value
        
        # 템플릿 렌더링
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template = Template(f.read())
            
            html_content = template.render(**report_data)
            print(f"[PREVIEW] {report_id} 템플릿 렌더링 완료: {template_name}")
            
            return jsonify({
                'success': True,
                'html': html_content,
                'missing_fields': [],
                'report_id': report_id,
                'report_name': report_config['name']
            })
        except Exception as e:
            import traceback
            error_msg = f"템플릿 렌더링 오류 ({template_name}): {str(e)}"
            print(f"[PREVIEW] {error_msg}")
            traceback.print_exc()
            return jsonify({'success': False, 'error': error_msg}), 400
        
    except Exception as e:
        import traceback
        error_msg = f"요약 보도자료 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return jsonify({'success': False, 'error': error_msg}), 400


@preview_bp.route('/generate-regional-preview', methods=['POST'])
def generate_regional_preview():
    """시도별 보도자료 미리보기 생성"""
    data = request.get_json() or {}
    region_id = data.get('region_id')
    
    if not region_id:
        return jsonify({'success': False, 'error': 'region_id가 필요합니다'}), 400
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'}), 400
    
    region_config = next((r for r in REGIONAL_REPORTS if r['id'] == region_id), None)
    if not region_config:
        return jsonify({'success': False, 'error': f'지역을 찾을 수 없습니다: {region_id}'}), 400
    
    is_reference = region_config.get('is_reference', False)
    
    html_content, error = generate_regional_report_html(excel_path, region_config['name'], is_reference)
    
    if error:
        return jsonify({'success': False, 'error': error}), 400
    
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
    from utils.excel_utils import extract_year_quarter_from_data
    
    data = request.get_json() or {}
    stat_id = data.get('stat_id')
    
    if not stat_id:
        return jsonify({'success': False, 'error': 'stat_id가 필요합니다'}), 400
    
    year = data.get('year') or session.get('year')
    quarter = data.get('quarter') or session.get('quarter')
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'}), 400
    
    # 연도/분기가 없으면 데이터에서 추출
    if not year or not quarter:
        try:
            from utils.excel_utils import get_previous_quarter
            default_year, default_quarter = get_previous_quarter()
            year, quarter = extract_year_quarter_from_data(excel_path, default_year=default_year, default_quarter=default_quarter)
            session['year'] = year
            session['quarter'] = quarter
            print(f"[통계표 미리보기] 데이터에서 연도/분기 추출: {year}년 {quarter}분기")
        except ValueError as e:
            return jsonify({'success': False, 'error': f'연도/분기 정보를 추출할 수 없습니다: {str(e)}'}), 400
    
    stat_config = next((s for s in STATISTICS_REPORTS if s['id'] == stat_id), None)
    if not stat_config:
        return jsonify({'success': False, 'error': f'통계표를 찾을 수 없습니다: {stat_id}'}), 400
    
    # 기초자료 수집표는 사용하지 않음 (분석표만 사용)
    html_content, error = generate_individual_statistics_html(excel_path, stat_config, year, quarter)
    
    if error:
        return jsonify({'success': False, 'error': error}), 400
    
    return jsonify({
        'success': True,
        'html': html_content,
        'stat_id': stat_id,
        'report_name': stat_config['name']
    })


@preview_bp.route('/generate-statistics-full-preview', methods=['POST'])
def generate_statistics_full_preview():
    """통계표 전체 보도자료 미리보기 생성"""
    data = request.get_json() or {}
    year = data.get('year', session.get('year', 2025))
    quarter = data.get('quarter', session.get('quarter', 2))
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'}), 400
    
    # 기초자료 수집표는 사용하지 않음 (분석표만 사용)
    html_content, error = generate_statistics_report_html(excel_path, year, quarter)
    
    if error:
        return jsonify({'success': False, 'error': error}), 400
    
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
            'description': f'본 보도자료는 {year if year else "[연도]"}년 {quarter if quarter else "[분기]"}/4분기 시·도별 지역경제동향을 수록하였습니다.' if year and quarter else '본 보도자료는 [연도/분기 입력 필요] 시·도별 지역경제동향을 수록하였습니다.',
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
            {'category': '생산', 'statistics_name': '광공업생산지수', 'department': '산업동향과', 'phone': '042-481-2161'},
            {'category': '생산', 'statistics_name': '서비스업생산지수', 'department': '서비스업동향과', 'phone': '042-481-2190'},
            {'category': '소비', 'statistics_name': '소매판매액지수', 'department': '서비스업동향과', 'phone': '042-481-2197'},
            {'category': '건설', 'statistics_name': '건설수주액', 'department': '산업동향과', 'phone': '042-481-2158'},
            {'category': '수출입', 'statistics_name': '수출·수입', 'department': customs_dept, 'phone': customs_phone},
            {'category': '물가', 'statistics_name': '소비자물가지수', 'department': '물가동향과', 'phone': '042-481-2531'},
            {'category': '고용', 'statistics_name': '고용률, 실업률', 'department': '고용통계과', 'phone': '042-481-2265'},
            {'category': '인구', 'statistics_name': '국내인구이동', 'department': '인구추계팀', 'phone': '042-481-2514'}
        ],
        'references': [
            {'content': '본문에 수록된 자료는 국가데이터처 홈페이지(http://mods.go.kr) 및 국가통계포털(http://kosis.kr)을 통해 이용할 수 있습니다.'}
        ],
        'notes': [
            '자료에 수록된 값은 잠정치이므로 추후 수정될 수 있습니다.'
        ]
    }

