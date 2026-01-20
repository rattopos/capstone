# -*- coding: utf-8 -*-
"""
요약 보도자료 데이터 추출 서비스
"""

import pandas as pd
from pathlib import Path
from utils.excel_utils import load_generator_module
from services.excel_processor import preprocess_excel
from config.reports import REGION_GROUPS
from services.excel_cache import get_sector_data


def safe_float(value, default=None):
    """안전한 float 변환 함수 (NaN, '-', 빈 문자열 체크 포함)"""
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
        if isinstance(value, str):
            value = value.strip()
            if value == '-' or value == '' or value.lower() in ['없음', 'nan', 'none']:
                return default
        result = float(value)
        if pd.isna(result):
            return default
        return result
    except (ValueError, TypeError):
        return default


# 지역명 정식 명칭 → 약칭 매핑
REGION_NAME_MAP = {
    '서울특별시': '서울', '부산광역시': '부산', '대구광역시': '대구',
    '인천광역시': '인천', '광주광역시': '광주', '대전광역시': '대전',
    '울산광역시': '울산', '세종특별자치시': '세종', '경기도': '경기',
    '강원도': '강원', '충청북도': '충북', '충청남도': '충남',
    '전라북도': '전북', '전라남도': '전남', '경상북도': '경북',
    '경상남도': '경남', '제주특별자치도': '제주',
    # 추가 변형 (강원특별자치도 등)
    '강원특별자치도': '강원', '전북특별자치도': '전북',
}

# 17개 시도 목록 (상수)
VALID_REGIONS = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                  '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']

SHEET_REPORT_ID_MAP = {
    'A 분석': 'manufacturing',
    'B 분석': 'service',
    'C 분석': 'consumption',
    'G 분석': 'export',
    'E(품목성질물가)분석': 'price',
    'D(고용률)분석': 'employment',
    "F'(건설)집계": 'construction'
}


def normalize_region_name(name):
    """지역명을 약칭으로 정규화"""
    if not name:
        return name
    name = str(name).strip()
    return REGION_NAME_MAP.get(name, name)


def _get_excel_path(xl_or_path):
    if isinstance(xl_or_path, pd.ExcelFile):
        return xl_or_path.io
    return xl_or_path


def _get_calculated_excel_path(excel_path: str) -> str:
    """수식 계산 로직으로 계산된 임시 파일 경로 반환 (전역 캐시 사용)."""
    from services.excel_cache import get_cached_calculated_path, set_cached_calculated_path
    from config.settings import TEMP_CALCULATED_DIR

    cached_path = get_cached_calculated_path(excel_path)
    if cached_path:
        return cached_path

    TEMP_CALCULATED_DIR.mkdir(parents=True, exist_ok=True)
    output_path = TEMP_CALCULATED_DIR / f"{Path(excel_path).stem}_calculated.xlsx"
    result_path, success, _ = preprocess_excel(
        excel_path,
        str(output_path),
        force_calculation=True
    )

    if success and result_path:
        set_cached_calculated_path(excel_path, result_path)

    return result_path


def _read_sheet_df(xl_or_path, sheet_name, data_only=None):
    """분석 시트는 수식 계산값(data_only)으로 읽는다."""
    excel_path = _get_excel_path(xl_or_path)
    if data_only is None:
        data_only = '분석' in sheet_name

    if data_only:
        calculated_path = _get_calculated_excel_path(excel_path)
        return pd.read_excel(calculated_path, sheet_name=sheet_name, header=None)

    if isinstance(xl_or_path, pd.ExcelFile):
        return pd.read_excel(xl_or_path, sheet_name=sheet_name, header=None)
    return pd.read_excel(excel_path, sheet_name=sheet_name, header=None)


def _build_chart_data_from_sector_cache(sector_payload: dict, is_trade: bool = False, is_employment: bool = False) -> dict:
    """부문별 캐시 데이터로 요약 차트 구조 생성"""
    data = sector_payload.get('data', {}) if isinstance(sector_payload, dict) else {}
    table_data = sector_payload.get('table_data') or data.get('table_data') or []
    if not table_data:
        table_df = sector_payload.get('table_df') if isinstance(sector_payload, dict) else None
        if isinstance(table_df, pd.DataFrame):
            try:
                table_data = table_df.to_dict(orient='records')
            except Exception:
                table_data = []
        elif isinstance(table_df, list):
            table_data = table_df
    regional_data = data.get('regional_data') or {}
    nationwide_data = data.get('nationwide_data') or {}

    def _pick_change(row: dict) -> float:
        for key in ('change_rate', 'growth_rate', 'change'):
            if row.get(key) is not None:
                return row.get(key)
        return 0.0

    def _pick_value(row: dict):
        for key in ('value', 'index', 'rate', 'employment_rate', 'amount'):
            if row.get(key) is not None:
                return row.get(key)
        return None

    region_changes = {}
    region_values = {}
    for row in table_data:
        if not isinstance(row, dict):
            continue
        region_name = row.get('region_name') or row.get('region') or row.get('name')
        if not region_name:
            continue
        region_changes[region_name] = _pick_change(row)
        region_values[region_name] = _pick_value(row)

    increase_regions = []
    decrease_regions = []
    chart_data = []

    for region in VALID_REGIONS:
        change_val = region_changes.get(region, 0.0)
        value_val = region_values.get(region, 0.0)

        data_row = {
            'name': region,
            'value': change_val,
            'index': value_val,
            'change': change_val,
            'rate': value_val
        }

        if is_trade:
            amount = value_val if value_val is not None else 0.0
            try:
                amount_normalized = min(100, max(0, float(amount) * 10))
            except (ValueError, TypeError):
                amount_normalized = 0.0
            data_row['amount'] = amount
            data_row['amount_normalized'] = amount_normalized

        if change_val >= 0:
            increase_regions.append(data_row)
        else:
            decrease_regions.append(data_row)
        chart_data.append(data_row)

    increase_regions.sort(key=lambda x: x['value'], reverse=True)
    decrease_regions.sort(key=lambda x: x['value'])

    nationwide_change = None
    nationwide_value = None
    if isinstance(nationwide_data, dict):
        for key in ('growth_rate', 'change_rate', 'change'):
            if nationwide_data.get(key) is not None:
                nationwide_change = nationwide_data.get(key)
                break
        for key in ('production_index', 'index_value', 'value', 'rate', 'amount', 'employment_rate'):
            if nationwide_data.get(key) is not None:
                nationwide_value = nationwide_data.get(key)
                break

    if is_employment and nationwide_change is None:
        nationwide_change = 0.0
    nationwide = {'change': nationwide_change}
    # change_rate 필드 보장
    nationwide['change_rate'] = nationwide_change if nationwide_change is not None else 0.0
    if is_trade:
        nationwide['amount'] = nationwide_value if nationwide_value is not None else 0.0
    else:
        nationwide['index'] = nationwide_value
        if is_employment:
            nationwide['rate'] = nationwide_value if nationwide_value is not None else 0.0

    return {
        'nationwide': nationwide,
        'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
        'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
        'increase_count': len(increase_regions),
        'decrease_count': len(decrease_regions),
        'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
        'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
        'above_count': len(increase_regions),
        'below_count': len(decrease_regions),
        'chart_data': chart_data[:18]
    }


def get_summary_overview_data(excel_path, year, quarter):
    """
    요약-지역경제동향 데이터 추출
    
    ★ 핵심 원칙: [행렬 데이터 구축 -> 열 단위 분석 -> 문장 생성] 순서
    - Step 1: 통합 매트릭스(comprehensive_table) 생성 (SSOT)
    - Step 2: 부문별(Column) 분석 - comprehensive_table에서 각 부문 데이터 추출
    - Step 3: 부문별 요약 문장 생성 - 추출된 데이터로 나레이션 생성
    """
    try:
        xl = pd.ExcelFile(excel_path)

        mining = _extract_chart_data(xl, 'A 분석', year=year, quarter=quarter)
        service = _extract_chart_data(xl, 'B 분석', year=year, quarter=quarter)
        consumption = _extract_chart_data(xl, 'C 분석', year=year, quarter=quarter)
        exports = _extract_chart_data(xl, 'G 분석', is_trade=True, year=year, quarter=quarter)
        price = _extract_chart_data(xl, 'E(품목성질물가)분석', year=year, quarter=quarter)
        employment = _extract_chart_data(xl, 'D(고용률)분석', is_employment=True, year=year, quarter=quarter)

        return {
            'production': {
                'mining': _summary_from_chart(mining),
                'service': _summary_from_chart(service)
            },
            'consumption': _summary_from_chart(consumption),
            'exports': _summary_from_chart(exports),
            'price': _summary_from_chart(price, include_above_below=True),
            'employment': _summary_from_chart(employment)
        }

    except Exception as e:
        print(f"🔍 [디버그] 요약 데이터 추출 오류:")
        print(f"  - 오류: {e}")
        print(f"  - excel_path: {excel_path}")
        print(f"  - year: {year}, quarter: {quarter}")
        import traceback
        traceback.print_exc()
        # 기본값/폴백 사용 금지: ValueError 발생
        raise ValueError(f"요약 데이터 추출 실패: {e}. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")


def _build_comprehensive_table(excel_path, year=None, quarter=None):
    """
    Step 1: 통합 매트릭스 생성
    17개 시도별로 [광공업, 서비스업, 소비, 수출, 물가, 고용] 데이터를 모두 담은 리스트 생성
    이 리스트가 요약 페이지 하단의 '주요 지역경제 지표' 테이블이 됩니다.
    
    반환 형식:
    [
        {'name': '서울', 'mining_production': 2.1, 'service_production': 1.5, ...},
        {'name': '부산', 'mining_production': -1.2, 'service_production': 0.8, ...},
        ...
    ]
    """
    # 기존 get_summary_table_data를 활용하되, comprehensive_table 형태로 변환
    table_data = get_summary_table_data(excel_path, year, quarter)
    
    nationwide = table_data.get('nationwide', {})
    region_groups = table_data.get('region_groups', [])
    
    # 모든 지역 데이터를 flat list로 변환 (comprehensive_table)
    comprehensive_table = []
    
    # 전국 데이터 추가 (참고용)
    comprehensive_table.append({
        'name': '전국',
        'mining_production': nationwide.get('mining_production'),
        'service_production': nationwide.get('service_production'),
        'retail_sales': nationwide.get('retail_sales'),
        'exports': nationwide.get('exports'),
        'price': nationwide.get('price'),
        'employment': nationwide.get('employment'),
    })

    for group in region_groups:
        for region in group.get('regions', []):
            comprehensive_table.append({
                'name': region.get('name'),
                'mining_production': region.get('mining_production'),
                'service_production': region.get('service_production'),
                'retail_sales': region.get('retail_sales'),
                'exports': region.get('exports'),
                'price': region.get('price'),
                'employment': region.get('employment'),
            })

    return comprehensive_table


def _compute_above_below_by_nationwide(chart_data):
    if chart_data is None:
        return None
    nationwide = chart_data.get('nationwide', {}).get('change')
    rows = chart_data.get('chart_data', [])
    if nationwide is None or not rows:
        return None

    above_regions = []
    below_regions = []

    for item in rows:
        name = item.get('name')
        if name not in VALID_REGIONS:
            continue
        value = item.get('value', item.get('change'))
        if value is None:
            continue
        entry = {'name': name, 'value': value}
        if value >= nationwide:
            above_regions.append(entry)
        else:
            below_regions.append(entry)

    above_regions.sort(key=lambda x: x['value'], reverse=True)
    below_regions.sort(key=lambda x: x['value'])

    return above_regions, below_regions


def _format_region_entries(regions, max_items=3):
    entries = []
    for region in (regions or [])[:max_items]:
        name = region.get('name') if isinstance(region, dict) else None
        if not name or name == '-':
            continue
        value = safe_float(region.get('value'), None)
        if value is None:
            entries.append(f"{name}")
        else:
            entries.append(f"{name}({value:.1f}%)")
    return entries


def _build_region_phrase(regions, count):
    entries = _format_region_entries(regions, max_items=3)
    if not entries:
        return "해당 시도는"

    list_text = ', '.join(entries)
    count_value = count if isinstance(count, int) else 0

    if count_value >= 4:
        return f"{list_text} 등 {count_value}개 시도는"

    last_name = None
    for region in reversed((regions or [])[:3]):
        if isinstance(region, dict) and region.get('name') and region.get('name') != '-':
            last_name = region.get('name')
            break

    if not last_name:
        return list_text

    try:
        from utils.text_utils import get_josa
        josa = get_josa(last_name, "은/는")
    except Exception:
        josa = "은"

    return f"{list_text}{josa}"


def _summary_from_chart(chart_data, include_above_below=False):
    if chart_data is None:
        chart_data = {}
    summary = {
        'increase_regions': chart_data.get('increase_regions', []),
        'decrease_regions': chart_data.get('decrease_regions', []),
        'increase_count': chart_data.get('increase_count', 0),
        'decrease_count': chart_data.get('decrease_count', 0),
        'nationwide': chart_data.get('nationwide', {}).get('change')
    }

    if include_above_below:
        comparison = _compute_above_below_by_nationwide(chart_data)
        if comparison:
            above_regions, below_regions = comparison
            summary['above_regions'] = above_regions[:3] if above_regions else [{'name': '-', 'value': 0.0}]
            summary['below_regions'] = below_regions[:3] if below_regions else [{'name': '-', 'value': 0.0}]
            summary['above_count'] = len(above_regions)
            summary['below_count'] = len(below_regions)
        else:
            summary['above_regions'] = chart_data.get('above_regions', summary['increase_regions'])
            summary['below_regions'] = chart_data.get('below_regions', summary['decrease_regions'])
            summary['above_count'] = chart_data.get('above_count', summary['increase_count'])
            summary['below_count'] = chart_data.get('below_count', summary['decrease_count'])

        summary['below_phrase'] = _build_region_phrase(summary.get('below_regions'), summary.get('below_count'))
        summary['above_phrase'] = _build_region_phrase(summary.get('above_regions'), summary.get('above_count'))

    return summary


def _build_region_value_map(chart_data):
    if chart_data is None:
        return {}
    return {
        item.get('name'): item.get('value', 0.0)
        for item in chart_data.get('chart_data', [])
    }


def get_summary_table_data(excel_path, year=None, quarter=None):
    """요약-지역경제동향 하단 표 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)

        def _chart_from_cache(report_id, is_trade=False, is_employment=False):
            if year is None or quarter is None:
                return None
            cached = get_sector_data(excel_path, year, quarter, report_id)
            if cached:
                return _build_chart_data_from_sector_cache(cached, is_trade=is_trade, is_employment=is_employment)
            return None

        mining = _chart_from_cache('manufacturing') or _extract_chart_data(xl, 'A 분석', year=year, quarter=quarter)
        service = _chart_from_cache('service') or _extract_chart_data(xl, 'B 분석', year=year, quarter=quarter)
        retail = _chart_from_cache('consumption') or _extract_chart_data(xl, 'C 분석', year=year, quarter=quarter)
        exports = _chart_from_cache('export', is_trade=True) or _extract_chart_data(xl, 'G 분석', is_trade=True, year=year, quarter=quarter)
        price = _chart_from_cache('price') or _extract_chart_data(xl, 'E(품목성질물가)분석', year=year, quarter=quarter)
        employment = _chart_from_cache('employment', is_employment=True) or _extract_chart_data(xl, 'D(고용률)분석', is_employment=True, year=year, quarter=quarter)

        mining = mining or {}
        service = service or {}
        retail = retail or {}
        exports = exports or {}
        price = price or {}
        employment = employment or {}

        mining_map = _build_region_value_map(mining)
        service_map = _build_region_value_map(service)
        retail_map = _build_region_value_map(retail)
        exports_map = _build_region_value_map(exports)
        price_map = _build_region_value_map(price)
        employment_map = _build_region_value_map(employment)

        nationwide = {
            'mining_production': mining.get('nationwide', {}).get('change'),
            'service_production': service.get('nationwide', {}).get('change'),
            'retail_sales': retail.get('nationwide', {}).get('change'),
            'exports': exports.get('nationwide', {}).get('change'),
            'price': price.get('nationwide', {}).get('change'),
            'employment': employment.get('nationwide', {}).get('change')
        }

        region_groups = []
        for group_name, regions in REGION_GROUPS.items():
            group_regions = []
            for region in regions:
                group_regions.append({
                    'name': region,
                    'mining_production': mining_map.get(region, 0.0),
                    'service_production': service_map.get(region, 0.0),
                    'retail_sales': retail_map.get(region, 0.0),
                    'exports': exports_map.get(region, 0.0),
                    'price': price_map.get(region, 0.0),
                    'employment': employment_map.get(region, 0.0)
                })
            region_groups.append({'name': group_name, 'regions': group_regions})

        return {
            'nationwide': nationwide,
            'region_groups': region_groups
        }
    except Exception as e:
        print(f"🔍 [디버그] 요약 표 데이터 추출 오류:")
        print(f"  - 오류: {e}")
        print(f"  - excel_path: {excel_path}")
        import traceback
        traceback.print_exc()
        raise ValueError(f"요약 표 데이터 추출 실패: {e}. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")


def get_production_summary_data(excel_path, year, quarter):
    """요약-생산 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        mining = _extract_chart_data(xl, 'A 분석', year=year, quarter=quarter)
        service = _extract_chart_data(xl, 'B 분석', year=year, quarter=quarter)

        return {
            'mining_production': mining,
            'service_production': service,
            'report_info': {'year': year, 'quarter': quarter, 'page_number': ''}
        }
    except Exception as e:
        print(f"🔍 [디버그] 생산 요약 데이터 추출 오류:")
        print(f"  - 오류: {e}")
        print(f"  - excel_path: {excel_path}")
        import traceback
        traceback.print_exc()
        raise ValueError(f"생산 요약 데이터 추출 실패: {e}. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")


def get_consumption_construction_data(excel_path, year, quarter):
    """요약-소비/건설 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        retail = _extract_chart_data(xl, 'C 분석', year=year, quarter=quarter)
        retail['qoq_change'] = None

        construction = _extract_construction_chart_data(xl, year=year, quarter=quarter)

        return {
            'retail_sales': retail,
            'construction': construction,
            'report_info': {'year': year, 'quarter': quarter, 'page_number': ''}
        }
    except Exception as e:
        print(f"소비건설 요약 데이터 오류: {e}")
        print(f"🔍 [디버그] 소비건설 데이터 추출 오류:")
        print(f"  - 오류: {e}")
        print(f"  - excel_path: {excel_path}")
        import traceback
        traceback.print_exc()
        raise ValueError(f"소비건설 데이터 추출 실패: {e}. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")


def _extract_construction_chart_data(xl, year=None, quarter=None):
    """건설수주액 차트 데이터 추출"""
    try:
        excel_path = _get_excel_path(xl)
        cached = None
        if year is not None and quarter is not None:
            cached = get_sector_data(excel_path, year, quarter, SHEET_REPORT_ID_MAP.get("F'(건설)집계"))
        if cached:
            cached_data = cached.get('data') if isinstance(cached, dict) else None
            cached_table = cached.get('table_data') or (cached_data.get('table_data') if isinstance(cached_data, dict) else None)
            has_amount = False
            if isinstance(cached_table, list):
                for row in cached_table:
                    if isinstance(row, dict) and row.get('amount') is not None:
                        has_amount = True
                        break
            if has_amount:
                return _build_chart_data_from_sector_cache(cached)

        regions = VALID_REGIONS.copy()
        
        nationwide = {'amount': 0, 'change': 0.0}
        increase_regions = []
        decrease_regions = []
        chart_data = []
        
        # F'(건설)집계 시트에서 데이터 추출
        if "F'(건설)집계" in xl.sheet_names:
            df = _read_sheet_df(xl, "F'(건설)집계", data_only=False)
            
            for i, row in df.iterrows():
                try:
                    region = str(row[1]).strip() if pd.notna(row[1]) else ''
                    code = str(row[2]).strip() if pd.notna(row[2]) else ''
                    
                    # 총계 행 (code == '0')
                    if code == '0':
                        # 현재 분기 값 (열 19)과 전년동분기 값 (열 15)
                        curr_val = safe_float(row[19])
                        prev_val = safe_float(row[15])
                        
                        # 증감률 계산
                        if prev_val is not None and prev_val != 0:
                            change = round((curr_val - prev_val) / prev_val * 100, 1)
                        else:
                            change = None
                        
                        # 금액 (조원 단위)
                        amount = round(curr_val / 10000, 1) if curr_val is not None else 0
                        amount_normalized = min(100, max(0, amount * 10))
                        
                        if region == '전국':
                            nationwide['amount'] = amount
                            nationwide['change'] = change
                        elif region in regions:
                            data = {
                                'name': region,
                                'value': change,
                                'amount': amount,
                                'amount_normalized': amount_normalized,
                                'change': change
                            }
                            
                            if change >= 0:
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
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'chart_data': chart_data[:18]
        }
    except Exception as e:
        print(f"🔍 [디버그] 건설 차트 데이터 추출 오류:")
        print(f"  - 오류: {e}")
        import traceback
        traceback.print_exc()
        # 기본값/폴백 사용 금지: ValueError 발생
        raise ValueError(f"건설 차트 데이터 추출 실패: {e}. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")


def _get_default_construction_data():
    """기본 건설 데이터"""
    return {
        'nationwide': {'amount': 0, 'change': 0.0},
        'increase_regions': [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
        'decrease_regions': [{'name': '-', 'value': 0.0, 'amount': 0, 'amount_normalized': 0}],
        'increase_count': 0, 'decrease_count': 0,
        'chart_data': []
    }


def get_trade_price_data(excel_path, year, quarter):
    """요약-수출물가 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        exports = _extract_chart_data(xl, 'G 분석', is_trade=True, year=year, quarter=quarter)
        price = _extract_chart_data(xl, 'E(품목성질물가)분석', year=year, quarter=quarter)

        comparison = _compute_above_below_by_nationwide(price)
        if comparison:
            above_regions, below_regions = comparison
            price['above_regions'] = above_regions[:3] if above_regions else [{'name': '-', 'value': 0.0}]
            price['below_regions'] = below_regions[:3] if below_regions else [{'name': '-', 'value': 0.0}]
            price['above_count'] = len(above_regions)
            price['below_count'] = len(below_regions)

        price['below_phrase'] = _build_region_phrase(price.get('below_regions'), price.get('below_count'))
        price['above_phrase'] = _build_region_phrase(price.get('above_regions'), price.get('above_count'))
        
        return {
            'exports': exports,
            'price': price
        }
    except Exception as e:
        print(f"🔍 [디버그] 수출물가 데이터 추출 오류:")
        print(f"  - 오류: {e}")
        print(f"  - excel_path: {excel_path}")
        print(f"  - year: {year}, quarter: {quarter}")
        import traceback
        traceback.print_exc()
        # 기본값/폴백 사용 금지: ValueError 발생
        raise ValueError(f"수출물가 데이터 추출 실패: {e}. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")


def get_employment_population_data(excel_path, year, quarter):
    """요약-고용인구 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        employment = _extract_chart_data(xl, 'D(고용률)분석', is_employment=True, year=year, quarter=quarter)
        
        population = {
            'inflow_regions': [],
            'outflow_regions': [],
            'inflow_count': 0,
            'outflow_count': 0,
            'chart_data': []
        }
        try:
            df = _read_sheet_df(xl, 'I(순인구이동)집계', data_only=False)
            regions = VALID_REGIONS.copy()
            
            # 시트 구조: col4=지역이름, col5=분류단계(0=합계), col25=2025 2/4분기
            # 합계(분류단계 0) 행만 추출
            processed_regions = set()
            region_data = {}  # 지역별 데이터 저장
            
            for i, row in df.iterrows():
                region = str(row[4]).strip() if pd.notna(row[4]) else ''
                division = str(row[5]).strip() if pd.notna(row[5]) else ''
                
                # 합계 행 (분류단계 0)만 처리, 중복 지역 방지
                if division == '0' and region in regions and region not in processed_regions:
                    try:
                        # 2025 2/4분기 데이터 (열 25)
                        curr_value = safe_float(row[25])
                        value = int(round(curr_value / 1000)) if curr_value is not None else 0
                        # 국내인구이동은 증감률 계산하지 않음 (raw data만 사용)
                        change = None
                        
                        processed_regions.add(region)
                        region_data[region] = {'value': value, 'change': change}
                        
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
            
            # chart_data 구성 - 지역 순서대로
            for region in regions:
                if region in region_data:
                    data = region_data[region]
                    population['chart_data'].append({
                        'name': region,
                        'value': data['value'],  # 순이동량 (천명)
                        'change': data['change']  # 전년동분기대비 증감률 (%)
                    })
                else:
                    population['chart_data'].append({
                        'name': region,
                        'value': 0,
                        'change': None
                    })
                    
        except Exception as e:
            print(f"인구이동 데이터 오류: {e}")
            import traceback
            traceback.print_exc()
        
        return {
            'employment': employment,
            'population': population
        }
    except Exception as e:
        print(f"🔍 [디버그] 고용인구 요약 데이터 추출 오류:")
        print(f"  - 오류: {e}")
        print(f"  - excel_path: {excel_path}")
        print(f"  - year: {year}, quarter: {quarter}")
        import traceback
        traceback.print_exc()
        # 기본값/폴백 사용 금지: ValueError 발생
        raise ValueError(f"고용인구 요약 데이터 추출 실패: {e}. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")


def _extract_employment_from_aggregate(xl, config, regions):
    """고용률 집계에서 단순 퍼센트포인트 차이 계산"""
    df_rate = _read_sheet_df(xl, config['rate_sheet'], data_only=False)
    rate_region_col = config['rate_region_col']
    rate_code_col = config.get('rate_code_col')
    rate_division_col = config.get('rate_division_col')
    rate_total_code = config['rate_total_code']
    rate_value_col = config['rate_value_col']
    prev_rate_col = config.get('prev_rate_col', rate_value_col - 4)

    nationwide = {'index': 0.0, 'change': 0.0, 'rate': 0.0, 'amount': 0}
    region_changes = {}
    region_indices = {}

    for _, row in df_rate.iterrows():
        try:
            region_raw = str(row[rate_region_col]).strip() if pd.notna(row[rate_region_col]) else ''
            region = normalize_region_name(region_raw)

            is_total = False
            if rate_code_col is not None:
                code = str(row[rate_code_col]).strip() if pd.notna(row[rate_code_col]) else ''
                is_total = (code == rate_total_code)
            elif rate_division_col is not None:
                division = str(row[rate_division_col]).strip() if pd.notna(row[rate_division_col]) else ''
                is_total = (division == rate_total_code)

            if not is_total:
                continue

            rate_val = safe_float(row[rate_value_col])
            prev_rate = safe_float(row[prev_rate_col])
            if rate_val is None or prev_rate is None:
                continue

            change_val = round(rate_val - prev_rate, 1)

            if region == '전국':
                nationwide['rate'] = round(rate_val, 1)
                nationwide['index'] = round(rate_val, 1)
                nationwide['change'] = change_val
            elif region in regions:
                if region not in region_indices:
                    region_indices[region] = round(rate_val, 1)
                    region_changes[region] = change_val
        except Exception:
            continue

    increase_regions = []
    decrease_regions = []
    chart_data = []

    for region in regions:
        change_val = region_changes.get(region, 0.0)
        index_val = region_indices.get(region, 0.0)

        data = {
            'name': region,
            'value': change_val,
            'index': index_val,
            'change': change_val,
            'rate': index_val
        }

        if change_val >= 0:
            increase_regions.append(data)
        else:
            decrease_regions.append(data)
        chart_data.append(data)

    increase_regions.sort(key=lambda x: x['value'], reverse=True)
    decrease_regions.sort(key=lambda x: x['value'])

    return {
        'nationwide': nationwide,
        'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
        'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
        'increase_count': len(increase_regions),
        'decrease_count': len(decrease_regions),
        'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
        'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
        'above_count': len(increase_regions),
        'below_count': len(decrease_regions),
        'chart_data': chart_data[:18]
    }


def _extract_chart_data(xl, sheet_name, is_trade=False, is_employment=False, year=None, quarter=None):
    """차트용 데이터 추출 (분석 시트 우선, 없거나 비어있으면 집계 시트 사용)"""

    # === table_locations/reports 기반 config 적용 ===
    from config.reports import SECTOR_REPORTS
    from config.table_locations import load_table_locations
    regions = VALID_REGIONS.copy()
    excel_path = _get_excel_path(xl)
    report_id = SHEET_REPORT_ID_MAP.get(sheet_name)
    cached = None
    if report_id and year is not None and quarter is not None:
        cached = get_sector_data(excel_path, year, quarter, report_id)
    if cached:
        return _build_chart_data_from_sector_cache(cached, is_trade=is_trade, is_employment=is_employment)

    # SECTOR_REPORTS에서 해당 sheet의 config 찾기
    sector_config = None
    for config in SECTOR_REPORTS:
        if config.get('sheet') == sheet_name or config.get('aggregation_structure', {}).get('sheet') == sheet_name:
            sector_config = config
            break
    if not sector_config:
        raise ValueError(f"[table_locations] 시트 설정을 찾을 수 없습니다: {sheet_name}. 반드시 데이터를 찾아야 합니다.")

    agg_struct = sector_config.get('aggregation_structure', {})
    agg_range = sector_config.get('aggregation_range', None)
    sheet_to_read = agg_struct.get('sheet', sheet_name)
    # pandas로 범위 추출
    import pandas as pd
    df = None
    if sheet_to_read:
        try:
            df_full = pd.read_excel(excel_path, sheet_name=sheet_to_read, header=None)
            if agg_range:
                from openpyxl.utils import column_index_from_string
                def _col_to_index(col_value):
                    if col_value is None:
                        return None
                    if isinstance(col_value, int):
                        return col_value
                    if isinstance(col_value, str) and col_value.strip():
                        return column_index_from_string(col_value.strip().upper()) - 1
                    return None
                row_start = max((agg_range.get('start_row', 1) - 1), 0)
                row_end = agg_range.get('end_row', len(df_full))
                col_start = _col_to_index(agg_range.get('start_col'))
                col_end = _col_to_index(agg_range.get('end_col'))
                if col_end is not None:
                    col_end += 1
                df = df_full.iloc[row_start:row_end, col_start:col_end].copy()
            else:
                df = df_full.copy()
        except Exception as e:
            raise ValueError(f"[table_locations] 표 추출 실패: {sheet_to_read}, {agg_range}, {e}")
    if df is None or df.empty:
        raise ValueError(f"[table_locations] 추출된 표 데이터가 없습니다: {sheet_to_read}, {agg_range}")

    # 헤더 포함 여부에 따라 첫 행을 컬럼명으로 지정
    if sector_config.get('header_included') and not df.empty:
        try:
            df.columns = df.iloc[0].tolist()
            df = df.iloc[1:].reset_index(drop=True)
        except Exception as e:
            print(f"[WARNING] 헤더 변환 실패: {e}")

    # 전국/지역별 change_rate 등 주요 값 추출 (예시: '전국' 행의 '증감률' 컬럼)
    # 실제 프로젝트별로 컬럼명/구조가 다를 수 있으므로, 여기서는 예시로 '전국' 행의 첫 번째 float 컬럼을 change_rate로 사용
    nationwide_row = None
    for i, row in df.iterrows():
        region_val = str(row[0]).strip() if not pd.isna(row[0]) else ''
        if region_val == '전국':
            nationwide_row = row
            break
    change_rate = None
    if nationwide_row is not None:
        for v in nationwide_row:
            try:
                fval = float(v)
                change_rate = fval
                break
            except Exception:
                continue
    if change_rate is None:
        raise ValueError(f"[table_locations] 전국 change_rate를 추출할 수 없습니다. 반드시 지정된 범위에서 추출해야 합니다.")

    # 결과 dict 예시 (실제 템플릿/데이터 구조에 맞게 확장 필요)

        
        nationwide_change_set = False  # 전국 증감률이 설정되었는지 추적
        
        for i, row in df.iterrows():
            try:
                region = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
                
                # 총지수 행인지 확인
                is_total_row = False
                if code_col is not None:
                    code = str(row[code_col]).strip() if pd.notna(row[code_col]) else ''
                    is_total_row = (code == total_code)
                elif division_col is not None:
                    division = str(row[division_col]).strip() if pd.notna(row[division_col]) else ''
                    is_total_row = (division == total_code)
                
                if is_total_row:
                    # 유효한 숫자 값인지 확인
                    change_val = None
                    if change_col < len(row):
                        change_val = safe_float(row[change_col], None)
                        if change_val is not None:
                            change_val = round(change_val, 1)
                    
                    if region == '전국':
                        # 첫 번째 유효한 전국 값만 사용
                        if not nationwide_change_set and change_val is not None:
                            nationwide['change'] = change_val
                            nationwide_change_set = True
                    elif region in regions and change_val is not None:
                        # 첫 번째 유효한 지역 값만 사용
                        if region not in region_changes:
                            region_changes[region] = change_val
            except:
                continue
        
        # 집계 시트에서 지수/고용률 값 추출
        region_indices = {}
        
        if is_employment and 'rate_sheet' in config:
            # 고용률 집계 시트에서 값 추출
            try:
                df_rate = _read_sheet_df(xl, config['rate_sheet'], data_only=False)
                rate_region_col = config['rate_region_col']
                rate_code_col = config.get('rate_code_col')
                rate_division_col = config.get('rate_division_col')
                rate_total_code = config['rate_total_code']
                rate_value_col = config['rate_value_col']
                prev_rate_col = config.get('prev_rate_col', rate_value_col - 4)
                
                for i, row in df_rate.iterrows():
                    try:
                        region_raw = str(row[rate_region_col]).strip() if pd.notna(row[rate_region_col]) else ''
                        region = normalize_region_name(region_raw)  # 지역명 정규화
                        
                        # 코드 컬럼 또는 division 컬럼으로 총계 행 확인
                        is_total = False
                        if rate_code_col is not None:
                            code = str(row[rate_code_col]).strip() if pd.notna(row[rate_code_col]) else ''
                            is_total = (code == rate_total_code)
                        elif rate_division_col is not None:
                            division = str(row[rate_division_col]).strip() if pd.notna(row[rate_division_col]) else ''
                            is_total = (division == rate_total_code)
                        
                        if is_total:
                            rate_val = safe_float(row[rate_value_col])
                            prev_rate = safe_float(row[prev_rate_col])
                            change_val = round(rate_val - prev_rate, 1) if (rate_val is not None and prev_rate is not None) else None
                            
                            if region == '전국':
                                nationwide['rate'] = round(rate_val, 1)
                                nationwide['index'] = round(rate_val, 1)
                                nationwide['change'] = change_val
                            elif region in regions:
                                region_indices[region] = round(rate_val, 1)
                                region_changes[region] = change_val
                    except:
                        continue
            except Exception as e:
                print(f"고용률 집계 시트 오류: {e}")
        
        elif 'index_sheet' in config:
            # 지수 집계 시트에서 값 추출
            try:
                df_index = _read_sheet_df(xl, config['index_sheet'], data_only=False)
                idx_region_col = config['index_region_col']
                idx_code_col = config.get('index_code_col')
                idx_division_col = config.get('index_division_col')
                idx_total_code = config['index_total_code']
                idx_value_col = config['index_value_col']
                
                nationwide_index_set = False  # 전국 지수가 설정되었는지 추적
                
                for i, row in df_index.iterrows():
                    try:
                        region_raw = str(row[idx_region_col]).strip() if pd.notna(row[idx_region_col]) else ''
                        region = normalize_region_name(region_raw)  # 지역명 정규화
                        
                        is_total = False
                        if idx_code_col is not None:
                            code = str(row[idx_code_col]).strip() if pd.notna(row[idx_code_col]) else ''
                            is_total = (code == str(idx_total_code))
                        elif idx_division_col is not None:
                            division = str(row[idx_division_col]).strip() if pd.notna(row[idx_division_col]) else ''
                            is_total = (division == str(idx_total_code))
                        
                        if is_total:
                            # 유효한 숫자 값인지 확인
                            index_val = safe_float(row[idx_value_col], None)
                            if index_val is not None:
                                index_val = round(index_val, 1)
                            
                            if region == '전국':
                                # 첫 번째 유효한 전국 값만 사용
                                if not nationwide_index_set and index_val is not None:
                                    nationwide['index'] = index_val
                                    if is_trade:
                                        nationwide['amount'] = round(index_val, 0)
                                    nationwide_index_set = True
                            elif region in regions and index_val is not None:
                                # 첫 번째 유효한 지역 값만 사용
                                if region not in region_indices:
                                    region_indices[region] = index_val
                    except:
                        continue
            except Exception as e:
                print(f"지수 집계 시트 오류: {e}")
        
        # 수출액 특별 처리 (G 분석) - 금액을 억달러 단위로 변환
        if is_trade and config.get('is_amount'):
            try:
                # G(수출)집계 시트에서 수출액 가져오기
                if 'G(수출)집계' in xl.sheet_names:
                    df_export = _read_sheet_df(xl, 'G(수출)집계', data_only=False)
                    for i, row in df_export.iterrows():
                        try:
                            region = str(row[3]).strip() if pd.notna(row[3]) else ''
                            division = str(row[4]).strip() if pd.notna(row[4]) else ''
                            if division == '0':
                                # 2025 2/4분기 수출액 (열 26, 백만달러 → 억달러 변환)
                                amount_val = safe_float(row[26])
                                amount_val = amount_val if amount_val is not None else 0
                                amount_in_billion = round(amount_val * 100, 0)  # 백만달러 → 억달러 (요청: 100배)
                                if region == '전국':
                                    nationwide['amount'] = amount_in_billion
                                    nationwide['index'] = amount_in_billion  # 차트용
                                elif region in regions:
                                    region_indices[region] = amount_in_billion
                        except:
                            continue
            except Exception as e:
                print(f"수출 집계 시트 오류: {e}")
        
        # 차트 데이터 구성
        for region in regions:
            change_val = region_changes.get(region, 0.0)
            index_val = region_indices.get(region, 100.0)
            
            data = {
                'name': region,
                'value': change_val,
                'index': index_val,
                'change': change_val,
                'rate': index_val
            }
            
            if is_trade:
                data['amount'] = index_val
                data['amount_normalized'] = min(100, max(0, index_val / 6))
            
            if change_val >= 0:
                increase_regions.append(data)
            else:
                decrease_regions.append(data)
            chart_data.append(data)
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])

        return {
            'nationwide': nationwide,
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions),
            'chart_data': chart_data[:18]
        }


def _extract_chart_data_from_raw(xl, config, regions, is_trade=False, is_employment=False):
    """기초자료 사용을 차단하기 위한 가드"""
    raise ValueError("기초자료 시트는 사용하지 않습니다. 분석표 기반 데이터만 허용됩니다.")


def _extract_chart_data_from_aggregate(xl, config, regions, is_trade=False):
    """집계 시트에서 차트 데이터 추출 및 전년동기비 계산"""
    try:
        try:
            df = _read_sheet_df(xl, config['aggregate_sheet'], data_only=False)
        except Exception as e:
            print(f"🔍 [디버그] {config['aggregate_sheet']} 시트 읽기 오류: {e}")
            return None

        region_col = config['region_col']
        code_col = config.get('code_col')
        total_code = config['total_code']
        curr_col = config['curr_col']
        prev_col = config['prev_col']
        is_amount = config.get('is_amount', False)
        
        nationwide = {'index': 100.0, 'change': 0.0, 'rate': 60.0, 'amount': 0}
        increase_regions = []
        decrease_regions = []
        chart_data = []
        
        for i, row in df.iterrows():
            region = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
            code = str(row[code_col]).strip() if code_col is not None and pd.notna(row[code_col]) else ''
            
            # 총계 코드 확인
            if code != str(total_code):
                continue
                
            curr_val = safe_float(row[curr_col], 0)
            prev_val = safe_float(row[prev_col], 0)
            
            if prev_val is not None and prev_val != 0:
                change = round((curr_val - prev_val) / prev_val * 100, 1)
            else:
                change = 0.0
                
            data = {
                'name': region,
                'value': change,
                'index': round(curr_val, 1),
                'change': change,
                'rate': round(curr_val, 1)
            }
            
            if is_trade or is_amount:
                amount = round(curr_val / 100, 1) if curr_val > 1000 else round(curr_val, 1)
                data['amount'] = amount
                data['amount_normalized'] = min(100, max(0, curr_val / 600))
                
            if region == '전국':
                nationwide['index'] = round(curr_val, 1)
                nationwide['change'] = change
                nationwide['rate'] = round(curr_val, 1)
                if is_trade or is_amount:
                    nationwide['amount'] = data.get('amount', 0)
            elif region in regions:
                if change >= 0:
                    increase_regions.append(data)
                else:
                    decrease_regions.append(data)
                chart_data.append(data)
                
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': nationwide,
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions),
            'chart_data': chart_data[:18]
        }
            
    except Exception as e:
        print(f"🔍 [디버그] 집계 시트 데이터 추출 중 오류: {e}")
        import traceback
        traceback.print_exc()
        return None



def _get_default_trade_data():
    """기본 수출입 데이터"""
    return {
        'nationwide': {'amount': None, 'change': None},
        'increase_regions': [],
        'decrease_regions': [],
        'increase_count': 0, 'decrease_count': 0,
        'chart_data': []
    }


def _get_default_employment_data():
    """기본 고용 데이터"""
    return {
        'nationwide': {'rate': None, 'change': None},
        'increase_regions': [],
        'decrease_regions': [],
        'increase_count': 0, 'decrease_count': 0,
        'chart_data': []
    }

