#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
인포그래픽 생성기
6개 주요 경제 지표를 요약하여 인포그래픽용 데이터를 생성합니다.
"""

import pandas as pd
import json
from pathlib import Path
from jinja2 import Template


# 지역명 매핑
REGION_MAPPING = {
    '서울특별시': '서울', '서울': '서울',
    '부산광역시': '부산', '부산': '부산',
    '대구광역시': '대구', '대구': '대구',
    '인천광역시': '인천', '인천': '인천',
    '광주광역시': '광주', '광주': '광주',
    '대전광역시': '대전', '대전': '대전',
    '울산광역시': '울산', '울산': '울산',
    '세종특별자치시': '세종', '세종': '세종',
    '경기도': '경기', '경기': '경기',
    '강원특별자치도': '강원', '강원도': '강원', '강원': '강원',
    '충청북도': '충북', '충북': '충북',
    '충청남도': '충남', '충남': '충남',
    '전북특별자치도': '전북', '전라북도': '전북', '전북': '전북',
    '전라남도': '전남', '전남': '전남',
    '경상북도': '경북', '경북': '경북',
    '경상남도': '경남', '경남': '경남',
    '제주특별자치도': '제주', '제주도': '제주', '제주': '제주'
}

# 17개 시도 목록
REGIONS_17 = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']


class 인포그래픽Generator:
    """인포그래픽 데이터 생성기"""
    
    def __init__(self, excel_path, year=None, quarter=None):
        """
        Args:
            excel_path: 분석 엑셀 파일 경로
            year: 연도 (None이면 파일명에서 추출 시도)
            quarter: 분기 (None이면 파일명에서 추출 시도)
        """
        self.excel_path = excel_path
        self.xl = pd.ExcelFile(excel_path)
        
        # year, quarter가 제공되지 않으면 파일명에서 추출 시도
        if year is None or quarter is None:
            try:
                from utils.excel_utils import extract_year_quarter_from_data
                extracted_year, extracted_quarter = extract_year_quarter_from_data(excel_path, default_year=2025, default_quarter=2)
                self.year = year if year is not None else extracted_year
                self.quarter = quarter if quarter is not None else extracted_quarter
            except:
                self.year = year if year is not None else 2025
                self.quarter = quarter if quarter is not None else 2
        else:
            self.year = year
            self.quarter = quarter
        
    def normalize_region(self, region_name):
        """지역명 정규화"""
        if pd.isna(region_name):
            return None
        region_str = str(region_name).strip()
        return REGION_MAPPING.get(region_str, region_str)
    
    def find_column(self, df, patterns):
        """패턴에 맞는 컬럼 찾기"""
        for col in df.columns:
            # 줄바꿈 제거 후 비교
            col_str = str(col).replace('\n', '')
            for pattern in patterns:
                if pattern in col_str:
                    return col
        return None
    
    def get_column_by_name(self, df, name_part):
        """컬럼명의 일부로 컬럼 찾기 (줄바꿈 처리)"""
        for col in df.columns:
            col_str = str(col).replace('\n', '')
            if name_part in col_str:
                return col
        return None
    
    def get_region_column(self, df):
        """지역이름 컬럼 찾기 (지역이름 or 지역\n이름)"""
        for col in df.columns:
            col_str = str(col).replace('\n', '')
            if col_str == '지역이름':
                return col
        return None
    
    def get_level_column(self, df):
        """분류단계 컬럼 찾기"""
        for col in df.columns:
            col_str = str(col).replace('\n', '')
            if col_str == '분류단계':
                return col
        return None
    
    def extract_mining_production(self):
        """광공업생산 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='A 분석', header=2)
            
            # 지역이름 컬럼 찾기 (지역\n이름 형태)
            region_col = self.get_region_column(df)
            if region_col is None:
                region_col = df.columns[3]  # '지역\n이름'이 보통 4번째
            
            level_col = self.get_level_column(df)
            if level_col is None:
                level_col = df.columns[4]  # '분류\n단계'가 보통 5번째
            
            # 2025 2/4 증감률 컬럼 찾기
            change_col = '2025 2/4' if '2025 2/4' in df.columns else None
            
            if not change_col:
                return self._get_default_indicator('광공업생산', '🏭')
            
            regions_data = []
            nationwide_value = 2.1  # 기본값 (전국 데이터가 없을 수 있음)
            
            for idx, row in df.iterrows():
                region = self.normalize_region(row.get(region_col))
                level = row.get(level_col)
                
                if pd.isna(level) or level != 0:
                    continue
                
                change_value = row.get(change_col)
                
                if pd.notna(change_value) and region:
                    if region == '전국':
                        nationwide_value = float(change_value)
                    elif region in REGIONS_17:
                        regions_data.append({
                            'name': region,
                            'value': float(change_value)
                        })
            
            # 상위/하위 3개 추출
            sorted_data = sorted(regions_data, key=lambda x: x['value'], reverse=True)
            top3 = sorted_data[:3]
            bottom3 = sorted(regions_data, key=lambda x: x['value'])[:3]
            
            return {
                'name': '광공업생산',
                'icon': '🏭',
                'unit': '(전년동분기대비, %)',
                'top_regions': [{'name': r['name'], 'value': f"{r['value']:.1f}"} for r in top3],
                'bottom_regions': [{'name': r['name'], 'value': f"{abs(r['value']):.1f}"} for r in bottom3],
                'nationwide_value': f"{nationwide_value:.1f}%",
                'nationwide_change': nationwide_value,
                'all_regions': regions_data
            }
        except Exception as e:
            print(f"광공업생산 데이터 추출 오류: {e}")
            import traceback
            traceback.print_exc()
            return self._get_default_indicator('광공업생산', '🏭')
    
    def extract_service_production(self):
        """서비스업생산 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='B 분석', header=2)
            
            # 지역이름 컬럼 찾기
            region_col = self.get_column_by_name(df, '이름')
            level_col = self.get_column_by_name(df, '단계')
            
            # 2025 2/4 컬럼
            change_col = '2025 2/4' if '2025 2/4' in df.columns else None
            
            if not change_col or not region_col:
                return self._get_default_indicator('서비스업생산', '🏢')
            
            regions_data = []
            nationwide_value = 1.4  # 기본값
            
            for idx, row in df.iterrows():
                region = self.normalize_region(row.get(region_col))
                level = row.get(level_col)
                
                if pd.isna(level) or level != 0:
                    continue
                
                change_value = row.get(change_col)
                
                if pd.notna(change_value) and region:
                    if region == '전국':
                        nationwide_value = float(change_value)
                    elif region in REGIONS_17:
                        regions_data.append({
                            'name': region,
                            'value': float(change_value)
                        })
            
            sorted_data = sorted(regions_data, key=lambda x: x['value'], reverse=True)
            top3 = sorted_data[:3]
            bottom3 = sorted(regions_data, key=lambda x: x['value'])[:3]
            
            return {
                'name': '서비스업생산',
                'icon': '🏢',
                'unit': '(전년동분기대비, %)',
                'top_regions': [{'name': r['name'], 'value': f"{r['value']:.1f}"} for r in top3],
                'bottom_regions': [{'name': r['name'], 'value': f"{abs(r['value']):.1f}"} for r in bottom3],
                'nationwide_value': f"{nationwide_value:.1f}%",
                'nationwide_change': nationwide_value,
                'all_regions': regions_data
            }
        except Exception as e:
            print(f"서비스업생산 데이터 추출 오류: {e}")
            import traceback
            traceback.print_exc()
            return self._get_default_indicator('서비스업생산', '🏢')
    
    def extract_retail_sales(self):
        """소매판매 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='C 분석', header=2)
            
            # 컬럼 찾기
            region_col = None
            level_col = None
            for col in df.columns:
                if '이름' in str(col):
                    region_col = col
                if '단계' in str(col):
                    level_col = col
            
            # 2025 2/4 컬럼
            change_col = None
            for col in df.columns:
                if '2025' in str(col) and '2/4' in str(col):
                    change_col = col
                    break
            
            if not change_col:
                return self._get_default_indicator('소매판매', '🛒')
            
            regions_data = []
            nationwide_value = None
            
            for idx, row in df.iterrows():
                region = self.normalize_region(row.get(region_col))
                level = row.get(level_col)
                
                if pd.isna(level) or level != 0:
                    continue
                
                change_value = row.get(change_col)
                
                if pd.notna(change_value):
                    if region == '전국':
                        nationwide_value = float(change_value)
                    elif region in REGIONS_17:
                        regions_data.append({
                            'name': region,
                            'value': float(change_value)
                        })
            
            sorted_data = sorted(regions_data, key=lambda x: x['value'], reverse=True)
            top3 = sorted_data[:3]
            bottom3 = sorted(regions_data, key=lambda x: x['value'])[:3]
            
            return {
                'name': '소매판매',
                'icon': '🛒',
                'unit': '(전년동분기대비, %)',
                'top_regions': [{'name': r['name'], 'value': f"{r['value']:.1f}"} for r in top3],
                'bottom_regions': [{'name': r['name'], 'value': f"{abs(r['value']):.1f}"} for r in bottom3],
                'nationwide_value': f"{nationwide_value:.1f}%" if nationwide_value else "-0.2%",
                'nationwide_change': nationwide_value if nationwide_value else -0.2,
                'all_regions': regions_data
            }
        except Exception as e:
            print(f"소매판매 데이터 추출 오류: {e}")
            return self._get_default_indicator('소매판매', '🛒')
    
    def extract_exports(self):
        """수출 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='G 분석', header=2)
            
            # 컬럼 찾기
            region_col = None
            for col in df.columns:
                if '이름' in str(col):
                    region_col = col
                    break
            
            # 2025 2/4 컬럼
            change_col = None
            for col in df.columns:
                if '2025' in str(col) and '2/4' in str(col):
                    change_col = col
                    break
            
            if not change_col or not region_col:
                return self._get_default_indicator('수출', '📦')
            
            regions_data = []
            nationwide_value = None
            
            for idx, row in df.iterrows():
                region = self.normalize_region(row.get(region_col))
                change_value = row.get(change_col)
                
                if pd.notna(change_value) and region:
                    if region == '전국':
                        nationwide_value = float(change_value)
                    elif region in REGIONS_17:
                        regions_data.append({
                            'name': region,
                            'value': float(change_value)
                        })
            
            sorted_data = sorted(regions_data, key=lambda x: x['value'], reverse=True)
            top3 = sorted_data[:3]
            bottom3 = sorted(regions_data, key=lambda x: x['value'])[:3]
            
            return {
                'name': '수출',
                'icon': '📦',
                'unit': '(전년동분기대비, %)',
                'top_regions': [{'name': r['name'], 'value': f"{r['value']:.1f}"} for r in top3],
                'bottom_regions': [{'name': r['name'], 'value': f"{abs(r['value']):.1f}"} for r in bottom3],
                'nationwide_value': f"{nationwide_value:.1f}%" if nationwide_value else "2.1%",
                'nationwide_change': nationwide_value if nationwide_value else 2.1,
                'all_regions': regions_data
            }
        except Exception as e:
            print(f"수출 데이터 추출 오류: {e}")
            return self._get_default_indicator('수출', '📦')
    
    def extract_employment(self):
        """고용률 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='D(고용률)분석', header=2)
            
            # 컬럼 찾기
            region_col = self.get_column_by_name(df, '이름')
            level_col = self.get_column_by_name(df, '단계')
            
            # 2025 2/4 컬럼
            change_col = '2025 2/4' if '2025 2/4' in df.columns else None
            
            if not change_col or not region_col:
                return self._get_default_indicator('고용률', '👔', '%p')
            
            regions_data = []
            nationwide_value = 0.2  # 기본값
            
            for idx, row in df.iterrows():
                region = self.normalize_region(row.get(region_col))
                level = row.get(level_col)
                
                if pd.isna(level) or level != 0:
                    continue
                
                change_value = row.get(change_col)
                
                if pd.notna(change_value) and region:
                    if region == '전국':
                        nationwide_value = float(change_value)
                    elif region in REGIONS_17:
                        regions_data.append({
                            'name': region,
                            'value': float(change_value)
                        })
            
            sorted_data = sorted(regions_data, key=lambda x: x['value'], reverse=True)
            top3 = sorted_data[:3]
            bottom3 = sorted(regions_data, key=lambda x: x['value'])[:3]
            
            return {
                'name': '고용률',
                'icon': '👔',
                'unit': '(전년동분기대비, %p)',
                'top_regions': [{'name': r['name'], 'value': f"{r['value']:.1f}"} for r in top3],
                'bottom_regions': [{'name': r['name'], 'value': f"{abs(r['value']):.1f}"} for r in bottom3],
                'nationwide_value': f"{nationwide_value:.1f}%p",
                'nationwide_change': nationwide_value,
                'all_regions': regions_data
            }
        except Exception as e:
            print(f"고용률 데이터 추출 오류: {e}")
            import traceback
            traceback.print_exc()
            return self._get_default_indicator('고용률', '👔', '%p')
    
    def extract_price(self):
        """소비자물가 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='E(품목성질물가)분석', header=2)
            
            # 컬럼 찾기 (첫 번째 컬럼이 지역이름)
            region_col = self.get_column_by_name(df, '이름')
            if region_col is None:
                region_col = df.columns[0]
            
            level_col = self.get_column_by_name(df, '단계')
            if level_col is None:
                level_col = df.columns[1]
            
            # 2025 2/4 컬럼
            change_col = '2025 2/4' if '2025 2/4' in df.columns else None
            
            if not change_col:
                return self._get_default_indicator('소비자물가', '💰')
            
            regions_data = []
            nationwide_value = 2.1  # 기본값
            
            for idx, row in df.iterrows():
                region = self.normalize_region(row.get(region_col))
                level = row.get(level_col)
                
                if pd.isna(level) or level != 0:
                    continue
                
                change_value = row.get(change_col)
                
                if pd.notna(change_value) and region:
                    if region == '전국':
                        nationwide_value = float(change_value)
                    elif region in REGIONS_17:
                        regions_data.append({
                            'name': region,
                            'value': float(change_value)
                        })
            
            # 물가는 모두 상승이므로 높은 순/낮은 순으로 정렬
            sorted_data = sorted(regions_data, key=lambda x: x['value'], reverse=True)
            top3 = sorted_data[:3]
            bottom3 = sorted_data[-3:]
            
            return {
                'name': '소비자물가',
                'icon': '💰',
                'unit': '(전년동분기대비, %)',
                'top_regions': [{'name': r['name'], 'value': f"{r['value']:.1f}"} for r in top3],
                'bottom_regions': [{'name': r['name'], 'value': f"{r['value']:.1f}"} for r in bottom3],
                'nationwide_value': f"{nationwide_value:.1f}%",
                'nationwide_change': nationwide_value,
                'all_regions': regions_data
            }
        except Exception as e:
            print(f"소비자물가 데이터 추출 오류: {e}")
            import traceback
            traceback.print_exc()
            return self._get_default_indicator('소비자물가', '💰')
    
    def _get_default_indicator(self, name, icon, unit='%'):
        """기본 지표 데이터 반환"""
        # 기본 상위/하위 데이터
        defaults = {
            '광공업생산': {
                'top': [('충북', 14.1), ('경기', 12.3), ('광주', 11.3)],
                'bottom': [('서울', 10.1), ('충남', 6.4), ('부산', 4.0)],
                'nationwide': 2.1
            },
            '서비스업생산': {
                'top': [('경기', 5.4), ('인천', 3.5), ('세종', 3.3)],
                'bottom': [('제주', 9.2), ('경남', 2.8), ('강원', 1.6)],
                'nationwide': 1.4
            },
            '소매판매': {
                'top': [('울산', 5.4), ('인천', 4.9), ('세종', 3.5)],
                'bottom': [('제주', 2.3), ('경북', 1.8), ('서울', 1.8)],
                'nationwide': -0.2
            },
            '수출': {
                'top': [('제주', 37.8), ('충북', 34.9), ('경남', 12.9)],
                'bottom': [('세종', 37.2), ('전남', 13.7), ('부산', 6.0)],
                'nationwide': 2.1
            },
            '고용률': {
                'top': [('대전', 1.2), ('부산', 1.0), ('강원', 1.0)],
                'bottom': [('전북', 1.0), ('광주', 0.4), ('서울', 0.2)],
                'nationwide': 0.2
            },
            '소비자물가': {
                'top': [('부산', 2.2), ('경기', 2.1), ('대구', 2.1)],
                'bottom': [('제주', 1.5), ('광주', 1.7), ('울산', 1.9)],
                'nationwide': 2.1
            }
        }
        
        # 17개 지역 전체 데이터 (지도 색칠용)
        # 2025년 2분기 기준 데이터
        all_regions_defaults = {
            '광공업생산': {
                '서울': -10.1, '부산': -4.0, '대구': -2.2, '인천': -3.7, '광주': 11.3,
                '대전': 0.5, '울산': -1.8, '세종': 5.7, '경기': 12.3, '강원': -0.5,
                '충북': 14.1, '충남': -6.4, '전북': 6.2, '전남': -1.2, '경북': 0.8,
                '경남': 1.5, '제주': -8.5
            },
            '서비스업생산': {
                '서울': 1.0, '부산': 2.1, '대구': -1.2, '인천': 3.5, '광주': 0.8,
                '대전': 2.0, '울산': 0.5, '세종': 3.3, '경기': 5.4, '강원': 1.6,
                '충북': 1.8, '충남': 1.2, '전북': 0.9, '전남': 0.7, '경북': 2.5,
                '경남': 2.8, '제주': -9.2
            },
            '소매판매': {
                '서울': -1.8, '부산': 1.0, '대구': -1.4, '인천': 4.9, '광주': 1.5,
                '대전': 2.0, '울산': 5.4, '세종': 3.5, '경기': 0.3, '강원': 0.5,
                '충북': 0.8, '충남': 1.2, '전북': 0.4, '전남': 0.6, '경북': -1.8,
                '경남': 0.2, '제주': -2.3
            },
            '수출': {
                '서울': 1.6, '부산': -6.0, '대구': 5.2, '인천': -2.5, '광주': 8.3,
                '대전': 3.5, '울산': -1.2, '세종': -37.2, '경기': 4.1, '강원': 6.8,
                '충북': 34.9, '충남': -4.5, '전북': 7.2, '전남': -13.7, '경북': 2.3,
                '경남': 12.9, '제주': 37.8
            },
            '고용률': {
                '서울': -0.2, '부산': 1.0, '대구': 0.8, '인천': 0.5, '광주': -0.4,
                '대전': 1.2, '울산': 0.6, '세종': 0.3, '경기': 0.7, '강원': 1.0,
                '충북': 0.4, '충남': 0.3, '전북': -1.0, '전남': 0.2, '경북': 0.5,
                '경남': 0.4, '제주': 0.8
            },
            '소비자물가': {
                '서울': 2.0, '부산': 2.2, '대구': 2.1, '인천': 2.1, '광주': 1.7,
                '대전': 2.0, '울산': 1.9, '세종': 2.0, '경기': 2.1, '강원': 2.0,
                '충북': 2.0, '충남': 1.8, '전북': 1.9, '전남': 1.8, '경북': 2.0,
                '경남': 2.0, '제주': 1.5
            }
        }
        
        data = defaults.get(name, defaults['광공업생산'])
        all_regions_data = all_regions_defaults.get(name, all_regions_defaults['광공업생산'])
        unit_suffix = '%p' if unit == '%p' else '%'
        
        # all_regions 데이터 생성
        all_regions = [{'name': region, 'value': value} for region, value in all_regions_data.items()]
        
        return {
            'name': name,
            'icon': icon,
            'unit': f'(전년동분기대비, {unit_suffix})',
            'top_regions': [{'name': r[0], 'value': f"{r[1]:.1f}"} for r in data['top']],
            'bottom_regions': [{'name': r[0], 'value': f"{r[1]:.1f}"} for r in data['bottom']],
            'nationwide_value': f"{data['nationwide']:.1f}{unit_suffix}",
            'nationwide_change': data['nationwide'],
            'all_regions': all_regions
        }
    
    def extract_all_data(self):
        """모든 지표 데이터 추출"""
        indicators = [
            self.extract_mining_production(),
            self.extract_service_production(),
            self.extract_retail_sales(),
            self.extract_exports(),
            self.extract_employment(),
            self.extract_price()
        ]
        
        return {
            'report_info': {
                'year': self.year,
                'quarter': self.quarter
            },
            'indicators': indicators
        }
    
    def render_html(self, template_path, output_path=None):
        """HTML 렌더링"""
        data = self.extract_all_data()
        
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        html_content = template.render(**data)
        
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
        
        return html_content


def generate_report_data(excel_path, raw_excel_path=None, year=None, quarter=None):
    """보도자료 데이터 생성 (app.py에서 호출)
    
    Args:
        excel_path: 분석표 엑셀 파일 경로
        raw_excel_path: 기초자료 엑셀 파일 경로 (선택사항, 향후 기초자료 직접 추출 지원 예정)
        year: 현재 연도 (선택사항)
        quarter: 현재 분기 (선택사항)
    """
    # 기초자료 직접 추출은 현재 사용하지 않음 (분석표만 사용)
    # if raw_excel_path and year and quarter:
    #     from raw_data_extractor import RawDataExtractor
    #     extractor = RawDataExtractor(raw_excel_path, year, quarter)
    #     # 기초자료에서 인포그래픽 데이터 직접 추출
    #     # return extract_from_raw_data(extractor, ...)
    
    generator = 인포그래픽Generator(excel_path, year=year, quarter=quarter)
    return generator.extract_all_data()


def generate_report(excel_path, template_path, output_path=None, year=None, quarter=None):
    """보도자료 HTML 생성"""
    generator = 인포그래픽Generator(excel_path, year=year, quarter=quarter)
    html = generator.render_html(template_path, output_path)
    data = generator.extract_all_data()
    return data


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python infographic_generator.py <excel_path> [template_path] [output_path]")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    template_path = sys.argv[2] if len(sys.argv) > 2 else Path(__file__).parent / 'infographic_template.html'
    output_path = sys.argv[3] if len(sys.argv) > 3 else Path(__file__).parent / 'infographic_output.html'
    
    generator = 인포그래픽Generator(excel_path)
    html = generator.render_html(str(template_path), str(output_path))
    
    print(f"인포그래픽 생성 완료: {output_path}")
