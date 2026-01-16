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

from config.settings import TEMPLATES_DIR, UPLOAD_FOLDER
from utils.filters import is_missing, format_value
from utils.text_utils import get_josa
from utils.excel_utils import load_generator_module
from utils.data_utils import check_missing_data
from .excel_cache import get_excel_file, clear_excel_cache


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
        data['report_info'] = {'year': year, 'quarter': quarter, 'page_number': ''}
        
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
        template.environment.filters['format_value'] = format_value
        template.environment.filters['is_missing'] = is_missing
        template.environment.filters['josa'] = get_josa
        html_content = template.render(**data)
        
        return html_content, None, []
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return None, f"스키마 기반 보도자료 생성 오류: {str(e)}", []


def generate_report_html(excel_path, report_config, year, quarter, custom_data=None, excel_file=None):
    """보도자료 HTML 생성 (최적화 버전 - 엑셀 파일 캐싱 지원)
    
    Args:
        excel_path: 분석표 엑셀 파일 경로
        report_config: 보도자료 설정 딕셔너리
        year: 연도
        quarter: 분기
        custom_data: 커스텀 데이터 (선택)
        excel_file: 캐시된 ExcelFile 객체 (선택사항, 있으면 재사용)
    
    주의: 기초자료 수집표는 사용하지 않으며, 분석표만 사용합니다.
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
        
        # 엑셀 파일 캐싱 (없으면 캐시에서 가져오기)
        if excel_file is None:
            excel_file = get_excel_file(excel_path, use_data_only=True)
        
        generator_name = report_config['generator']
        template_name = report_config['template']
        report_name = report_config['name']
        report_id = report_config['id']
        
        # 보도자료 생성 시작
        
        # Generator가 None인 경우 (일러두기 등) 스키마에서 기본값 로드
        if generator_name is None:
            return _generate_from_schema(template_name, report_id, year, quarter, custom_data)
        
        # Generator 모듈 로드
        module = load_generator_module(generator_name)
        if not module:
            print(f"[ERROR] Generator 모듈을 찾을 수 없습니다: {generator_name}")
            return None, f"Generator 모듈을 찾을 수 없습니다: {generator_name}", []
        
        # 사용 가능한 함수 확인
        available_funcs = [name for name in dir(module) if not name.startswith('_')]
        
        # Generator 클래스 찾기 (BaseGenerator 제외)
        generator_class = None
        for name in dir(module):
            obj = getattr(module, name)
            if isinstance(obj, type) and name.endswith('Generator') and name != 'BaseGenerator':
                generator_class = obj
                break
        
        data = None
        
        # 방법 1: generate_report_data 함수 사용
        # 주의: 기초자료 수집표는 사용하지 않으므로 분석표만 사용
        if hasattr(module, 'generate_report_data'):
            try:
                # 함수 시그니처 확인하여 year, quarter, excel_file 전달 시도
                import inspect
                sig = inspect.signature(module.generate_report_data)
                params = list(sig.parameters.keys())
                
                # 캐시된 excel_file 전달 시도
                call_kwargs = {}
                if 'excel_file' in params:
                    call_kwargs['excel_file'] = excel_file
                if 'year' in params:
                    call_kwargs['year'] = year
                if 'quarter' in params:
                    call_kwargs['quarter'] = quarter
                
                if call_kwargs:
                    data = module.generate_report_data(excel_path, **call_kwargs)
                elif 'year' in params and 'quarter' in params:
                    data = module.generate_report_data(excel_path, year=year, quarter=quarter)
                elif 'year' in params:
                    data = module.generate_report_data(excel_path, year=year)
                else:
                    # 분석표만 사용
                    data = module.generate_report_data(excel_path)
            except TypeError as e:
                # 파라미터가 맞지 않으면 기본 호출 시도
                try:
                    data = module.generate_report_data(excel_path, year=year, quarter=quarter)
                except TypeError:
                    data = module.generate_report_data(excel_path)
            except Exception as e:
                print(f"[WARNING] 데이터 생성 실패: {e}")
                try:
                    data = module.generate_report_data(excel_path, year=year, quarter=quarter)
                except:
                    data = module.generate_report_data(excel_path)
        
        # 방법 2: generate_report 함수 직접 호출
        # 주의: 기초자료 수집표는 사용하지 않으므로 분석표만 사용
        elif hasattr(module, 'generate_report'):
            template_path = TEMPLATES_DIR / template_name
            output_path = TEMPLATES_DIR / f"{report_name}_preview.html"
            try:
                # 분석표만 사용
                data = module.generate_report(excel_path, template_path, output_path)
            except (TypeError, AttributeError):
                data = module.generate_report(excel_path, template_path, output_path)
        
        # 방법 3: Generator 클래스 사용
        elif generator_class:
            try:
                # __init__ 시그니처 확인하여 year, quarter, excel_file 전달 시도
                import inspect
                sig = inspect.signature(generator_class.__init__)
                params = list(sig.parameters.keys())
                
                # year와 quarter는 반드시 포함 (명시적 전달)
                init_kwargs = {}
                if 'year' in params:
                    init_kwargs['year'] = year
                if 'quarter' in params:
                    init_kwargs['quarter'] = quarter
                if 'excel_file' in params:
                    init_kwargs['excel_file'] = excel_file
                
                # year와 quarter가 있으면 명시적으로 전달
                if 'year' in params and 'quarter' in params:
                    if 'excel_file' in params:
                        generator = generator_class(excel_path, year=year, quarter=quarter, excel_file=excel_file)
                    else:
                        generator = generator_class(excel_path, year=year, quarter=quarter)
                elif init_kwargs:
                    generator = generator_class(excel_path, **init_kwargs)
                else:
                    generator = generator_class(excel_path)
            except (TypeError, AttributeError):
                # 시그니처 확인 실패 시 year, quarter 포함하여 시도
                try:
                    generator = generator_class(excel_path, year=year, quarter=quarter)
                except TypeError:
                    # year, quarter 파라미터가 없으면 기본 초기화
                    generator = generator_class(excel_path)
            
            data = generator.extract_all_data()
        
        else:
            error_msg = f"유효한 Generator를 찾을 수 없습니다: {generator_name}"
            print(f"[ERROR] {error_msg}")
            print(f"[ERROR] 사용 가능한 함수: {available_funcs}")
            return None, error_msg, []
        
        # Top3 regions 후처리
        if data and 'regional_data' in data:
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
                for r in data['top3_decrease_regions']:
                    if 'growth_rate' not in r:
                        r['growth_rate'] = r.get('change', 0)
                    if 'change' not in r:
                        r['change'] = r.get('growth_rate', 0)
                    if 'industries' not in r:
                        r['industries'] = r.get('age_groups', r.get('top_industries', []))
                    if 'age_groups' not in r:
                        r['age_groups'] = r.get('industries', [])
            
        
        # 커스텀 데이터 병합
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
        
        # report_info 강제 추가/업데이트 (연도/분기 보장)
        if data is None:
            data = {}
        
        if 'report_info' not in data:
            data['report_info'] = {}
        
        # year, quarter가 None이 아니면 업데이트
        if year is not None:
            data['report_info']['year'] = year
        if quarter is not None:
            data['report_info']['quarter'] = quarter
        
        # report_info에 year나 quarter가 없으면 동적으로 추출 (하드코딩 제거)
        if 'year' not in data['report_info'] or data['report_info']['year'] is None:
            data['report_info']['year'] = year if year is not None else (data.get('year') or 2025)
        if 'quarter' not in data['report_info'] or data['report_info']['quarter'] is None:
            data['report_info']['quarter'] = quarter if quarter is not None else (data.get('quarter') or 2)
        
        # 페이지 번호는 더 이상 사용하지 않음 (목차 생성 중단)
        data['report_info']['page_number'] = ""
        
        
        # 결측치 확인
        missing = check_missing_data(data, report_id)
        
        # 템플릿 렌더링 전 데이터 키 로깅 (디버깅용)
        print(f"[DEBUG] {report_name} 템플릿 렌더링 전 데이터 키: {list(data.keys()) if data else 'None'}")
        if data:
            # 주요 키의 타입과 크기 정보도 출력
            for key, value in data.items():
                if isinstance(value, (dict, list)):
                    print(f"  - {key}: {type(value).__name__} (크기: {len(value) if hasattr(value, '__len__') else 'N/A'})")
                else:
                    print(f"  - {key}: {type(value).__name__}")
        
        # 템플릿 렌더링
        template_path = TEMPLATES_DIR / template_name
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        template.environment.filters['format_value'] = format_value
        template.environment.filters['is_missing'] = is_missing
        template.environment.filters['josa'] = get_josa
        
        html_content = template.render(**data)
        
        return html_content, None, missing
        
    except Exception as e:
        import traceback
        error_msg = f"보도자료 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return None, error_msg, []


def generate_regional_report_html(excel_path, region_name, is_reference=False):
    """시도별 보도자료 HTML 생성"""
    try:
        # 파일 존재 확인
        excel_path_obj = Path(excel_path)
        if not excel_path_obj.exists() or not excel_path_obj.is_file():
            error_msg = f"엑셀 파일을 찾을 수 없습니다: {excel_path}"
            print(f"[ERROR] {error_msg}")
            return None, error_msg
        
        
        generator_path = TEMPLATES_DIR / 'regional_generator.py'
        if not generator_path.exists():
            return None, f"시도별 Generator를 찾을 수 없습니다"
        
        spec = importlib.util.spec_from_file_location('regional_generator', str(generator_path))
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        
        generator = module.RegionalGenerator(excel_path)
        template_path = TEMPLATES_DIR / 'regional_template.html'
        
        html_content = generator.render_html(region_name, str(template_path))
        
        return html_content, None
        
    except Exception as e:
        import traceback
        error_msg = f"시도별 보도자료 생성 오류: {str(e)}"
        print(f"[ERROR] {error_msg}")
        traceback.print_exc()
        return None, error_msg


def generate_statistics_report_html(excel_path, year, quarter):
    """통계표 보도자료 HTML 생성
    
    Args:
        excel_path: 분석표 엑셀 파일 경로
        year: 연도
        quarter: 분기
    
    주의: 고객사 요청으로 통계표 섹션 전체를 생성하지 않기로 결정됨
    """
    # 통계표 생성 비활성화
    return None, "통계표 생성이 비활성화되었습니다."


def generate_individual_statistics_html(excel_path, stat_config, year, quarter):
    """개별 통계표 HTML 생성
    
    Args:
        excel_path: 분석표 엑셀 파일 경로
        stat_config: 통계표 설정 딕셔너리
        year: 연도
        quarter: 분기
    
    주의: 고객사 요청으로 통계표 섹션 전체를 생성하지 않기로 결정됨
    """
    # 통계표 생성 비활성화
    return None, "통계표 생성이 비활성화되었습니다."
    
    # 아래 코드는 통계표 생성이 비활성화되어 실행되지 않음
    # 필요시 주석을 해제하여 다시 활성화 가능
    """
        stat_id = stat_config['id']
        template_name = stat_config['template']
        table_name = stat_config.get('table_name')
        
        # 통계표 Generator 모듈 로드
        generator_path = TEMPLATES_DIR / 'statistics_table_generator.py'
        if generator_path.exists():
            spec = importlib.util.spec_from_file_location('statistics_table_generator', str(generator_path))
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            
            # Generator 초기화 시그니처 확인하여 raw_excel_path 파라미터 제거
            import inspect
            sig = inspect.signature(module.StatisticsTableGenerator.__init__)
            params = list(sig.parameters.keys())
            
            if 'raw_excel_path' in params:
                # raw_excel_path 파라미터가 있으면 None으로 전달 (하위 호환성)
                generator = module.StatisticsTableGenerator(
                    excel_path,
                    raw_excel_path=None,
                    current_year=year,
                    current_quarter=quarter
                )
            else:
                # raw_excel_path 파라미터가 없으면 제거된 버전
                generator = module.StatisticsTableGenerator(
                    excel_path,
                    current_year=year,
                    current_quarter=quarter
                )
        else:
            generator = None
        
        PAGE1_REGIONS = ["전국", "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종"]
        PAGE2_REGIONS = ["경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"]
        
        # 통계표 목차 (고객사 요구사항 변경으로 더 이상 생성하지 않음)
        # if stat_id == 'stat_toc':
        #     toc_items = [
        #         {'number': 1, 'name': '광공업생산지수'},
        #         {'number': 2, 'name': '서비스업생산지수'},
        #         {'number': 3, 'name': '소매판매액지수'},
        #         {'number': 4, 'name': '건설수주액'},
        #         {'number': 5, 'name': '고용률'},
        #         {'number': 6, 'name': '실업률'},
        #         {'number': 7, 'name': '국내 인구이동'},
        #         {'number': 8, 'name': '수출액'},
        #         {'number': 9, 'name': '수입액'},
        #         {'number': 10, 'name': '소비자물가지수'},
        #     ]
        #     template_data = {
        #         'year': year,
        #         'quarter': quarter,
        #         'toc_items': toc_items
        #     }
        
        # 통계표 - 개별 지표
        if table_name and table_name != 'GRDP' and generator:
            table_order = ['광공업생산지수', '서비스업생산지수', '소매판매액지수', '건설수주액',
                          '고용률', '실업률', '국내인구이동', '수출액', '수입액', '소비자물가지수']
            try:
                table_index = table_order.index(table_name) + 1
            except ValueError:
                table_index = 1
            
            try:
            config = generator.TABLE_CONFIG.get(table_name)
                if not config:
                    print(f"[통계표] 설정 없음: {table_name}, 빈 데이터 반환")
                    data = generator._create_empty_table_data()
                else:
                data = generator.extract_table_data(table_name)
                    # data가 None이면 빈 데이터로 대체
                    if data is None:
                        print(f"[통계표] 데이터 추출 실패: {table_name}, 빈 데이터 반환")
                        data = generator._create_empty_table_data()
            except Exception as e:
                import traceback
                print(f"[통계표] 데이터 추출 중 오류: {table_name} - {e}")
                traceback.print_exc()
                # 오류 발생 시 빈 데이터 반환
                try:
                    data = generator._create_empty_table_data()
                except:
                    data = {
                        'yearly': {},
                        'quarterly': {},
                        'yearly_years': [],
                        'quarterly_keys': []
                    }
                
                # 연도 키: JSON 데이터에서 가져오거나 기본값 사용
                yearly_years = data.get('yearly_years', ["2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"])
                
                # 분기 키: 실제 데이터에 있는 분기만 사용 (데이터 없는 분기 제외)
                quarterly_keys = data.get('quarterly_keys', [])
                if not quarterly_keys and data.get('quarterly'):
                    # quarterly_keys가 없으면 quarterly 딕셔너리에서 키 추출 후 정렬
                    quarterly_keys = sorted(data['quarterly'].keys(), key=lambda x: (
                        int(x[:4]), int(x[5]) if len(x) > 5 else 0
                    ))
                
            # page_base 계산 제거 (페이지 번호는 더 이상 사용하지 않음, 목차 생성 중단)
            # page_base = 22 + (table_index - 1) * 2
            
            # config가 없어도 기본값 사용
            unit = config.get('단위', '[자료 없음]') if config else '[자료 없음]'
                
                template_data = {
                    'year': year,
                    'quarter': quarter,
                    'index': table_index,
                    'title': table_name,
                'unit': unit,
                'data': data if data else {'yearly': {}, 'quarterly': {}, 'yearly_years': [], 'quarterly_keys': []},
                    'page1_regions': PAGE1_REGIONS,
                    'page2_regions': PAGE2_REGIONS,
                    'yearly_years': yearly_years,
                'quarterly_keys': quarterly_keys
                }
        
        # 통계표 - GRDP
        elif stat_id == 'stat_grdp':
            if generator:
                try:
                grdp_data = generator._create_grdp_placeholder()
                except Exception as e:
                    print(f"[통계표] GRDP 데이터 생성 실패: {e}")
                    grdp_data = {
                        'title': '분기 지역내총생산(GRDP)',
                        'unit': '[전년동기비, %]',
                        'data': {
                            'yearly': {},
                            'quarterly': {},
                            'yearly_years': [],
                            'quarterly_keys': []
                        }
                    }
            else:
                grdp_data = {
                    'title': '분기 지역내총생산(GRDP)',
                    'unit': '[전년동기비, %]',
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
                'quarterly_keys': quarterly_keys
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
                'terms_page2': terms_page2
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
    """
