# -*- coding: utf-8 -*-
"""
요약 보고서 데이터 추출 서비스
"""

import pandas as pd


def get_summary_overview_data(excel_path, year, quarter):
    """요약-지역경제동향 데이터 추출"""
    try:
        xl = pd.ExcelFile(excel_path)
        
        mining_data = _extract_sector_summary(xl, 'A 분석')
        service_data = _extract_sector_summary(xl, 'B 분석')
        consumption_data = _extract_sector_summary(xl, 'C 분석')
        export_data = _extract_sector_summary(xl, 'G 분석')
        price_data = _extract_sector_summary(xl, 'E(품목성질물가)분석')
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
    """시트에서 요약 데이터 추출 (집계 시트에서 전년동기비 계산)"""
    try:
        regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 분석 시트 -> 집계 시트 매핑 (전년동기비 계산 필요)
        aggregate_config = {
            'A 분석': {
                'aggregate_sheet': 'A(광공업생산)집계',
                'region_col': 1, 'code_col': 4, 'total_code': 'BCD',
                'curr_col': 26, 'prev_col': 22,  # 2025 2/4 (col26), 2024 2/4 (col22)
            },
            'B 분석': {
                'aggregate_sheet': 'B(서비스업생산)집계',
                'region_col': 1, 'code_col': 4, 'total_code': 'E~S',
                'curr_col': 25, 'prev_col': 21,  # 2025 2/4 (col25), 2024 2/4 (col21)
            },
            'C 분석': {
                'aggregate_sheet': 'C(소비)집계',
                'region_col': 1, 'division_col': 2, 'total_code': '0',
                'curr_col': 24, 'prev_col': 20,  # 2025 2/4 (col24), 2024 2/4 (col20)
            },
            'G 분석': {
                'aggregate_sheet': 'G(수출)집계',
                'region_col': 1, 'division_col': 2, 'total_code': '0',
                'curr_col': 26, 'prev_col': 22,  # 2025 2/4 (col26), 2024 2/4 (col22)
            },
            'E(품목성질물가)분석': {
                'use_custom_extractor': True,
                'extractor': '_extract_price_summary_from_aggregate'
            },
            'D(고용률)분석': {
                'use_custom_extractor': True,
                'extractor': '_extract_employment_summary_from_aggregate'
            },
        }
        
        config = aggregate_config.get(sheet_name)
        if not config:
            return _get_default_sector_summary()
        
        # 커스텀 추출기 사용
        if config.get('use_custom_extractor'):
            if config['extractor'] == '_extract_price_summary_from_aggregate':
                return _extract_price_summary_from_aggregate(xl, regions)
            elif config['extractor'] == '_extract_employment_summary_from_aggregate':
                return _extract_employment_summary_from_aggregate(xl, regions)
        
        # 집계 시트에서 전년동기비 계산
        agg_sheet = config['aggregate_sheet']
        if agg_sheet not in xl.sheet_names:
            print(f"집계 시트 없음: {agg_sheet}")
            return _get_default_sector_summary()
        
        df = pd.read_excel(xl, sheet_name=agg_sheet, header=None)
        
        region_col = config['region_col']
        code_col = config.get('code_col')
        division_col = config.get('division_col')
        total_code = config['total_code']
        curr_col = config['curr_col']
        prev_col = config['prev_col']
        
        increase_regions = []
        decrease_regions = []
        nationwide = 0.0
        
        for i, row in df.iterrows():
            try:
                region = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
                
                # 총지수 행 찾기
                is_total_row = False
                if code_col is not None:
                    code = str(row[code_col]).strip() if pd.notna(row[code_col]) else ''
                    is_total_row = (code == total_code)
                elif division_col is not None:
                    division = str(row[division_col]).strip() if pd.notna(row[division_col]) else ''
                    is_total_row = (division == total_code)
                
                if is_total_row:
                    # 전년동기비 계산
                    curr_val = float(row[curr_col]) if pd.notna(row[curr_col]) else 0
                    prev_val = float(row[prev_col]) if pd.notna(row[prev_col]) else 0
                    
                    if prev_val != 0:
                        change = round((curr_val - prev_val) / prev_val * 100, 1)
                    else:
                        change = 0.0
                    
                    if region == '전국':
                        nationwide = change
                    elif region in regions:
                        if change >= 0:
                            increase_regions.append({'name': region, 'value': change})
                        else:
                            decrease_regions.append({'name': region, 'value': change})
            except Exception as e:
                continue
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': round(nationwide, 1),
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions)
        }
    except Exception as e:
        print(f"{sheet_name} 데이터 추출 오류: {e}")
        import traceback
        traceback.print_exc()
        return _get_default_sector_summary()


def _extract_price_summary_from_aggregate(xl, regions):
    """E(품목성질물가)집계 시트에서 소비자물가 증감률 추출"""
    try:
        df = pd.read_excel(xl, sheet_name='E(품목성질물가)집계', header=None)
        
        # 열 구조: 0=지역이름, 1=분류단계, 2=가중치, 3=분류이름
        # 열 20=2024 2/4분기, 열 24=2025 2/4분기
        
        increase_regions = []
        decrease_regions = []
        nationwide = 0.0
        
        for i, row in df.iterrows():
            try:
                region = str(row[0]).strip() if pd.notna(row[0]) else ''
                division = str(row[1]).strip() if pd.notna(row[1]) else ''
                
                # 총지수 행 (division == '0')
                if division == '0':
                    # 2025 2/4분기 지수 (열 24)와 2024 2/4분기 지수 (열 20)
                    curr_val = float(row[24]) if pd.notna(row[24]) else 0
                    prev_val = float(row[20]) if pd.notna(row[20]) else 0
                    
                    # 전년동분기 대비 증감률 계산
                    if prev_val != 0:
                        change = round((curr_val - prev_val) / prev_val * 100, 1)
                    else:
                        change = 0.0
                    
                    if region == '전국':
                        nationwide = change
                    elif region in regions:
                        if change >= 0:
                            increase_regions.append({'name': region, 'value': change})
                        else:
                            decrease_regions.append({'name': region, 'value': change})
            except Exception as e:
                continue
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': round(nationwide, 1),
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions)
        }
    except Exception as e:
        print(f"물가 집계 데이터 추출 오류: {e}")
        import traceback
        traceback.print_exc()
        return _get_default_sector_summary()


def _extract_employment_summary_from_aggregate(xl, regions):
    """D(고용률)집계 시트에서 고용률 증감 추출"""
    try:
        df = pd.read_excel(xl, sheet_name='D(고용률)집계', header=None)
        
        # 열 구조: 0=지역코드, 1=지역이름, 2=분류단계, 3=산업이름
        # 열 24=2025 2/4분기 고용률, 열 20=2024 2/4분기 고용률
        
        increase_regions = []
        decrease_regions = []
        nationwide = 0.0
        
        for i, row in df.iterrows():
            try:
                region = str(row[1]).strip() if pd.notna(row[1]) else ''
                division = str(row[2]).strip() if pd.notna(row[2]) else ''
                industry = str(row[3]).strip() if pd.notna(row[3]) else ''
                
                # 총계 행 (division == '0' 또는 industry == '계')
                if division == '0' or industry == '계':
                    # 2025 2/4분기 고용률 (열 24)와 2024 2/4분기 고용률 (열 20)
                    curr_val = float(row[24]) if pd.notna(row[24]) else 0
                    prev_val = float(row[20]) if pd.notna(row[20]) else 0
                    
                    # 전년동분기 대비 증감 (고용률은 %p 단위)
                    change = round(curr_val - prev_val, 1)
                    
                    if region == '전국':
                        nationwide = change
                    elif region in regions:
                        if change >= 0:
                            increase_regions.append({'name': region, 'value': change})
                        else:
                            decrease_regions.append({'name': region, 'value': change})
            except Exception as e:
                continue
        
        increase_regions.sort(key=lambda x: x['value'], reverse=True)
        decrease_regions.sort(key=lambda x: x['value'])
        
        return {
            'nationwide': round(nationwide, 1),
            'increase_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'decrease_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
            'above_regions': increase_regions[:3] if increase_regions else [{'name': '-', 'value': 0.0}],
            'below_regions': decrease_regions[:3] if decrease_regions else [{'name': '-', 'value': 0.0}],
            'above_count': len(increase_regions),
            'below_count': len(decrease_regions)
        }
    except Exception as e:
        print(f"고용률 집계 데이터 추출 오류: {e}")
        import traceback
        traceback.print_exc()
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


def get_summary_table_data(excel_path):
    """요약 테이블 데이터 (집계 시트에서 전년동기비 계산)"""
    try:
        xl = pd.ExcelFile(excel_path)
        all_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                       '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 집계 시트 설정 (전년동기비 계산)
        # curr_col: 2025 2/4, prev_col: 2024 2/4
        sheet_configs = {
            'mining_production': {
                'sheet': 'A(광공업생산)집계',
                'region_col': 1, 'code_col': 4, 'total_code': 'BCD',
                'curr_col': 26, 'prev_col': 22,  # 2025 2/4, 2024 2/4
                'calc_type': 'growth_rate'
            },
            'service_production': {
                'sheet': 'B(서비스업생산)집계',
                'region_col': 1, 'code_col': 4, 'total_code': 'E~S',
                'curr_col': 25, 'prev_col': 21,
                'calc_type': 'growth_rate'
            },
            'retail_sales': {
                'sheet': 'C(소비)집계',
                'region_col': 1, 'division_col': 2, 'total_code': '0',
                'curr_col': 24, 'prev_col': 20,
                'calc_type': 'growth_rate'
            },
            'exports': {
                'sheet': 'G(수출)집계',
                'region_col': 1, 'division_col': 2, 'total_code': '0',
                'curr_col': 26, 'prev_col': 22,
                'calc_type': 'growth_rate'
            },
            'price': {
                'sheet': 'E(품목성질물가)집계',
                'region_col': 0, 'division_col': 1, 'total_code': '0',
                'curr_col': 24, 'prev_col': 20,  # 2025 2/4, 2024 2/4
                'calc_type': 'growth_rate'
            },
            'employment': {
                'sheet': 'D(고용률)집계',
                'region_col': 1, 'division_col': 2, 'total_code': '0',
                'curr_col': 24, 'prev_col': 20,  # 2025 2/4, 2024 2/4
                'calc_type': 'difference'  # 고용률은 %p
            },
        }
        
        nationwide_data = {
            'mining_production': 0.0, 'service_production': 0.0, 'retail_sales': 0.0,
            'exports': 0.0, 'price': 0.0, 'employment': 0.0
        }
        
        region_data = {r: {'name': r, 'mining_production': 0.0, 'service_production': 0.0,
                          'retail_sales': 0.0, 'exports': 0.0, 'price': 0.0, 'employment': 0.0}
                      for r in all_regions}
        
        for key, config in sheet_configs.items():
            try:
                sheet_name = config['sheet']
                if sheet_name not in xl.sheet_names:
                    print(f"시트 없음: {sheet_name}")
                    continue
                    
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
                region_col = config['region_col']
                code_col = config.get('code_col')
                division_col = config.get('division_col')
                total_code = config['total_code']
                curr_col = config['curr_col']
                prev_col = config['prev_col']
                calc_type = config['calc_type']
                
                for i, row in df.iterrows():
                    try:
                        region = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
                        
                        # 총지수 행 찾기
                        is_total_row = False
                        if code_col is not None:
                            code = str(row[code_col]).strip() if pd.notna(row[code_col]) else ''
                            is_total_row = (code == total_code)
                        elif division_col is not None:
                            division = str(row[division_col]).strip() if pd.notna(row[division_col]) else ''
                            is_total_row = (division == total_code)
                        
                        if is_total_row:
                            curr_val = float(row[curr_col]) if pd.notna(row[curr_col]) else 0
                            prev_val = float(row[prev_col]) if pd.notna(row[prev_col]) else 0
                            
                            # 계산 방식에 따라 증감률 또는 차이 계산
                            if calc_type == 'difference':
                                value = round(curr_val - prev_val, 1)
                            else:  # growth_rate
                                if prev_val != 0:
                                    value = round((curr_val - prev_val) / prev_val * 100, 1)
                                else:
                                    value = 0.0
                            
                            if region == '전국':
                                nationwide_data[key] = value
                            elif region in all_regions:
                                region_data[region][key] = value
                    except:
                        continue
            except Exception as e:
                print(f"{config.get('sheet', key)} 테이블 데이터 추출 오류: {e}")
                continue
        
        region_groups = [
            {'name': '경인', 'regions': [region_data['서울'], region_data['인천'], region_data['경기']]},
            {'name': '충청', 'regions': [region_data['대전'], region_data['세종'], region_data['충북'], region_data['충남']]},
            {'name': '호남', 'regions': [region_data['광주'], region_data['전북'], region_data['전남'], region_data['제주']]},
            {'name': '동북', 'regions': [region_data['대구'], region_data['경북'], region_data['강원']]},
            {'name': '동남', 'regions': [region_data['부산'], region_data['울산'], region_data['경남']]},
        ]
        
        return {
            'nationwide': nationwide_data,
            'region_groups': region_groups
        }
    except Exception as e:
        print(f"요약 테이블 데이터 오류: {e}")
        import traceback
        traceback.print_exc()
        return {'nationwide': {'mining_production': 0.0, 'service_production': 0.0, 'retail_sales': 0.0,
                              'exports': 0.0, 'price': 0.0, 'employment': 0.0}, 'region_groups': []}


def get_production_summary_data(excel_path, year, quarter):
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


def get_consumption_construction_data(excel_path, year, quarter):
    """요약-소비건설 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        retail = _extract_chart_data(xl, 'C 분석')
        
        # 건설 데이터 추출
        construction = _extract_construction_chart_data(xl)
        
        return {
            'retail_sales': retail,
            'construction': construction
        }
    except Exception as e:
        print(f"소비건설 요약 데이터 오류: {e}")
        return {
            'retail_sales': _get_default_chart_data(),
            'construction': _get_default_construction_data()
        }


def _extract_construction_chart_data(xl):
    """건설수주액 차트 데이터 추출"""
    try:
        regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        nationwide = {'amount': 0, 'change': 0.0}
        increase_regions = []
        decrease_regions = []
        chart_data = []
        
        # F'(건설)집계 시트에서 데이터 추출
        if "F'(건설)집계" in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name="F'(건설)집계", header=None)
            
            for i, row in df.iterrows():
                try:
                    region = str(row[1]).strip() if pd.notna(row[1]) else ''
                    code = str(row[2]).strip() if pd.notna(row[2]) else ''
                    
                    # 총계 행 (code == '0')
                    if code == '0':
                        # 현재 분기 값 (열 19)과 전년동분기 값 (열 15)
                        curr_val = float(row[19]) if pd.notna(row[19]) else 0
                        prev_val = float(row[15]) if pd.notna(row[15]) else 0
                        
                        # 증감률 계산
                        if prev_val != 0:
                            change = round((curr_val - prev_val) / prev_val * 100, 1)
                        else:
                            change = 0.0
                        
                        # 금액 (백억원 단위)
                        amount = int(round(curr_val / 10, 0))
                        amount_normalized = min(100, max(0, curr_val / 30))  # 최대 3000억원 기준
                        
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
        print(f"건설 차트 데이터 추출 오류: {e}")
        return _get_default_construction_data()


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


def get_employment_population_data(excel_path, year, quarter):
    """요약-고용인구 데이터"""
    try:
        xl = pd.ExcelFile(excel_path)
        employment = _extract_chart_data(xl, 'D(고용률)분석', is_employment=True)
        
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
            
            # 시트 구조: col1=지역이름, col2=분류(유입/유출/순인구이동), col24=2025 2/4분기
            # 순인구이동 수 행만 추출
            processed_regions = set()
            
            for i, row in df.iterrows():
                region = str(row[1]).strip() if pd.notna(row[1]) else ''
                category = str(row[2]).strip() if pd.notna(row[2]) else ''
                
                # 순인구이동 수 행만 처리, 중복 지역 방지
                if '순인구이동' in category and region in regions and region not in processed_regions:
                    try:
                        # 2025 2/4분기 데이터 (열 24)
                        value = int(float(row[24])) if pd.notna(row[24]) else 0
                        processed_regions.add(region)
                        
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
            import traceback
            traceback.print_exc()
        
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
    """차트용 데이터 추출 (시트별 열 설정 적용)"""
    try:
        regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 시트별 설정 (분석 시트와 집계 시트 매핑)
        sheet_config = {
            'A 분석': {
                'region_col': 3, 'code_col': 6, 'total_code': 'BCD',
                'change_col': 21,  # 증감률
                'index_sheet': 'A(광공업생산)집계',
                'index_region_col': 4, 'index_code_col': 7, 'index_total_code': 'BCD',
                'index_value_col': 26  # 2025 2/4분기 지수
            },
            'B 분석': {
                'region_col': 3, 'code_col': 6, 'total_code': 'E~S',
                'change_col': 20,  # 증감률
                'index_sheet': 'B(서비스업생산)집계',
                'index_region_col': 3, 'index_code_col': 6, 'index_total_code': 'E~S',
                'index_value_col': 25  # 2025 2/4분기 지수
            },
            'C 분석': {
                'region_col': 3, 'division_col': 4, 'total_code': '0',
                'change_col': 20,  # 증감률
                'index_sheet': 'C(소비)집계',
                'index_region_col': 2, 'index_division_col': 3, 'index_total_code': 0,
                'index_value_col': 24  # 2025 2/4분기 지수
            },
            'G 분석': {
                'region_col': 3, 'division_col': 4, 'total_code': '0',
                'change_col': 22,  # 증감률
                'index_sheet': 'G(수출)집계',
                'index_region_col': 3, 'index_division_col': 4, 'index_total_code': '0',
                'index_value_col': 26,  # 2025 2/4분기 수출액
                'is_amount': True  # 금액 단위 (억달러 변환)
            },
            'E(품목성질물가)분석': {
                'region_col': 0, 'division_col': 1, 'total_code': '0',
                'change_col': 16,  # 증감률
                'index_sheet': 'E(품목성질물가)집계',
                'index_region_col': 0, 'index_division_col': 1, 'index_total_code': 0,
                'index_value_col': 21  # 2025 2/4분기 지수
            },
            'D(고용률)분석': {
                'region_col': 2, 'division_col': 3, 'total_code': '0',
                'rate_sheet': 'D(고용률)집계',
                'rate_region_col': 1, 'rate_division_col': 2, 'rate_total_code': '0',
                'rate_value_col': 21,  # 2025 2/4분기 고용률
                'prev_rate_col': 17  # 2024 2/4분기 고용률 (증감 계산용)
            },
        }
        
        config = sheet_config.get(sheet_name, {})
        if not config:
            return _get_default_chart_data()
        
        # 분석 시트에서 증감률 추출
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        
        nationwide = {'index': 100.0, 'change': 0.0, 'rate': 60.0, 'amount': 0}
        increase_regions = []
        decrease_regions = []
        chart_data = []
        region_changes = {}  # 지역별 증감률 저장
        
        region_col = config['region_col']
        code_col = config.get('code_col')
        division_col = config.get('division_col')
        total_code = config['total_code']
        change_col = config.get('change_col', 20)
        
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
                    if change_col < len(row) and pd.notna(row[change_col]):
                        try:
                            change_val = round(float(row[change_col]), 1)
                        except (ValueError, TypeError):
                            change_val = None
                    
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
                df_rate = pd.read_excel(xl, sheet_name=config['rate_sheet'], header=None)
                rate_region_col = config['rate_region_col']
                rate_division_col = config['rate_division_col']
                rate_total_code = config['rate_total_code']
                rate_value_col = config['rate_value_col']
                prev_rate_col = config.get('prev_rate_col', rate_value_col - 4)
                
                for i, row in df_rate.iterrows():
                    try:
                        region = str(row[rate_region_col]).strip() if pd.notna(row[rate_region_col]) else ''
                        division = str(row[rate_division_col]).strip() if pd.notna(row[rate_division_col]) else ''
                        
                        if division == rate_total_code:
                            rate_val = float(row[rate_value_col]) if pd.notna(row[rate_value_col]) else 60.0
                            prev_rate = float(row[prev_rate_col]) if pd.notna(row[prev_rate_col]) else rate_val
                            change_val = round(rate_val - prev_rate, 1)
                            
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
                df_index = pd.read_excel(xl, sheet_name=config['index_sheet'], header=None)
                idx_region_col = config['index_region_col']
                idx_code_col = config.get('index_code_col')
                idx_division_col = config.get('index_division_col')
                idx_total_code = config['index_total_code']
                idx_value_col = config['index_value_col']
                
                nationwide_index_set = False  # 전국 지수가 설정되었는지 추적
                
                for i, row in df_index.iterrows():
                    try:
                        region = str(row[idx_region_col]).strip() if pd.notna(row[idx_region_col]) else ''
                        
                        is_total = False
                        if idx_code_col is not None:
                            code = str(row[idx_code_col]).strip() if pd.notna(row[idx_code_col]) else ''
                            is_total = (code == str(idx_total_code))
                        elif idx_division_col is not None:
                            division = str(row[idx_division_col]).strip() if pd.notna(row[idx_division_col]) else ''
                            is_total = (division == str(idx_total_code))
                        
                        if is_total:
                            # 유효한 숫자 값인지 확인
                            index_val = None
                            if pd.notna(row[idx_value_col]):
                                try:
                                    index_val = round(float(row[idx_value_col]), 1)
                                except (ValueError, TypeError):
                                    index_val = None
                            
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
                    df_export = pd.read_excel(xl, sheet_name='G(수출)집계', header=None)
                    for i, row in df_export.iterrows():
                        try:
                            region = str(row[3]).strip() if pd.notna(row[3]) else ''
                            division = str(row[4]).strip() if pd.notna(row[4]) else ''
                            if division == '0':
                                # 2025 2/4분기 수출액 (열 26, 백만달러 → 억달러 변환)
                                amount_val = float(row[26]) if pd.notna(row[26]) else 0
                                amount_in_billion = round(amount_val / 100, 0)  # 백만달러 → 억달러
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
    except Exception as e:
        print(f"{sheet_name} 차트 데이터 오류: {e}")
        import traceback
        traceback.print_exc()
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

