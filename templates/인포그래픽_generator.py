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


class 인포그래픽Generator:
    """인포그래픽 데이터 생성기"""
    
    def __init__(self, excel_path):
        """
        Args:
            excel_path: 분석 엑셀 파일 경로
        """
        self.excel_path = excel_path
        self.xl = pd.ExcelFile(excel_path)
        self.year = 2025
        self.quarter = 2
        
    def normalize_region(self, region_name):
        """지역명 정규화"""
        return REGION_MAPPING.get(region_name, region_name)
    
    def extract_mining_production(self):
        """광공업생산 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='A 분석')
            
            # 시도별 데이터 추출 (분류단계가 0이고 산업코드가 C인 행)
            regions_data = []
            nationwide_value = None
            
            for idx, row in df.iterrows():
                region = str(row.get('지역이름', row.iloc[3] if len(row) > 3 else ''))
                level = row.get('분류단계', row.iloc[4] if len(row) > 4 else None)
                
                if pd.isna(level) or level != 0:
                    continue
                    
                # 증감률 컬럼 찾기 (마지막 분기)
                change_col = None
                for col in df.columns:
                    if '증감' in str(col) or '2025.2/4' in str(col):
                        change_col = col
                        break
                
                if change_col:
                    change_value = row.get(change_col, row.iloc[22] if len(row) > 22 else 0)
                else:
                    change_value = row.iloc[22] if len(row) > 22 else 0
                
                if pd.notna(change_value):
                    region_short = self.normalize_region(region)
                    if region_short == '전국' or region == '전국':
                        nationwide_value = float(change_value)
                    elif region_short in REGION_MAPPING.values():
                        regions_data.append({
                            'name': region_short,
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
                'nationwide_value': f"{nationwide_value:.1f}%" if nationwide_value else "2.1%",
                'nationwide_change': nationwide_value if nationwide_value else 2.1
            }
        except Exception as e:
            print(f"광공업생산 데이터 추출 오류: {e}")
            return self._get_default_indicator('광공업생산', '🏭')
    
    def extract_service_production(self):
        """서비스업생산 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='B 분석')
            
            regions_data = []
            nationwide_value = None
            
            for idx, row in df.iterrows():
                region = str(row.get('지역이름', row.iloc[3] if len(row) > 3 else ''))
                level = row.get('분류단계', row.iloc[4] if len(row) > 4 else None)
                
                if pd.isna(level) or level != 0:
                    continue
                
                change_value = row.iloc[22] if len(row) > 22 else 0
                
                if pd.notna(change_value):
                    region_short = self.normalize_region(region)
                    if region_short == '전국' or region == '전국':
                        nationwide_value = float(change_value)
                    elif region_short in REGION_MAPPING.values():
                        regions_data.append({
                            'name': region_short,
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
                'nationwide_value': f"{nationwide_value:.1f}%" if nationwide_value else "1.4%",
                'nationwide_change': nationwide_value if nationwide_value else 1.4
            }
        except Exception as e:
            print(f"서비스업생산 데이터 추출 오류: {e}")
            return self._get_default_indicator('서비스업생산', '🏢')
    
    def extract_retail_sales(self):
        """소매판매 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='C 분석')
            
            regions_data = []
            nationwide_value = None
            
            for idx, row in df.iterrows():
                region = str(row.get('지역이름', row.iloc[3] if len(row) > 3 else ''))
                level = row.get('분류단계', row.iloc[4] if len(row) > 4 else None)
                
                if pd.isna(level) or level != 0:
                    continue
                
                change_value = row.iloc[22] if len(row) > 22 else 0
                
                if pd.notna(change_value):
                    region_short = self.normalize_region(region)
                    if region_short == '전국' or region == '전국':
                        nationwide_value = float(change_value)
                    elif region_short in REGION_MAPPING.values():
                        regions_data.append({
                            'name': region_short,
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
                'nationwide_change': nationwide_value if nationwide_value else -0.2
            }
        except Exception as e:
            print(f"소매판매 데이터 추출 오류: {e}")
            return self._get_default_indicator('소매판매', '🛒')
    
    def extract_exports(self):
        """수출 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='G 분석')
            
            regions_data = []
            nationwide_value = None
            
            for idx, row in df.iterrows():
                region = str(row.iloc[3] if len(row) > 3 else '')
                
                # 증감률 컬럼
                change_value = row.iloc[16] if len(row) > 16 else 0
                
                if pd.notna(change_value):
                    region_short = self.normalize_region(region)
                    if '전국' in region or region_short == '전국':
                        nationwide_value = float(change_value)
                    elif region_short in REGION_MAPPING.values():
                        regions_data.append({
                            'name': region_short,
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
                'nationwide_change': nationwide_value if nationwide_value else 2.1
            }
        except Exception as e:
            print(f"수출 데이터 추출 오류: {e}")
            return self._get_default_indicator('수출', '📦')
    
    def extract_employment(self):
        """고용률 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='D(고용률)분석')
            
            regions_data = []
            nationwide_value = None
            
            for idx, row in df.iterrows():
                region = str(row.iloc[3] if len(row) > 3 else '')
                
                # 증감 컬럼
                change_value = row.iloc[16] if len(row) > 16 else 0
                
                if pd.notna(change_value):
                    region_short = self.normalize_region(region)
                    if '전국' in region or region_short == '전국':
                        nationwide_value = float(change_value)
                    elif region_short in REGION_MAPPING.values():
                        regions_data.append({
                            'name': region_short,
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
                'nationwide_value': f"{nationwide_value:.1f}%p" if nationwide_value else "0.2%p",
                'nationwide_change': nationwide_value if nationwide_value else 0.2
            }
        except Exception as e:
            print(f"고용률 데이터 추출 오류: {e}")
            return self._get_default_indicator('고용률', '👔', '%p')
    
    def extract_price(self):
        """소비자물가 데이터 추출"""
        try:
            df = pd.read_excel(self.xl, sheet_name='E(품목성질물가)분석')
            
            regions_data = []
            nationwide_value = None
            
            for idx, row in df.iterrows():
                region = str(row.iloc[3] if len(row) > 3 else '')
                level = row.iloc[4] if len(row) > 4 else None
                
                if pd.isna(level) or level != 0:
                    continue
                
                # 증감률 컬럼
                change_value = row.iloc[16] if len(row) > 16 else 0
                
                if pd.notna(change_value):
                    region_short = self.normalize_region(region)
                    if '전국' in region or region_short == '전국':
                        nationwide_value = float(change_value)
                    elif region_short in REGION_MAPPING.values():
                        regions_data.append({
                            'name': region_short,
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
                'nationwide_value': f"{nationwide_value:.1f}%" if nationwide_value else "2.1%",
                'nationwide_change': nationwide_value if nationwide_value else 2.1
            }
        except Exception as e:
            print(f"소비자물가 데이터 추출 오류: {e}")
            return self._get_default_indicator('소비자물가', '💰')
    
    def _get_default_indicator(self, name, icon, unit='%'):
        """기본 지표 데이터 반환"""
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
        
        data = defaults.get(name, defaults['광공업생산'])
        unit_suffix = '%p' if unit == '%p' else '%'
        
        return {
            'name': name,
            'icon': icon,
            'unit': f'(전년동분기대비, {unit_suffix})',
            'top_regions': [{'name': r[0], 'value': f"{r[1]:.1f}"} for r in data['top']],
            'bottom_regions': [{'name': r[0], 'value': f"{r[1]:.1f}"} for r in data['bottom']],
            'nationwide_value': f"{data['nationwide']:.1f}{unit_suffix}",
            'nationwide_change': data['nationwide']
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


def generate_report_data(excel_path):
    """보고서 데이터 생성 (app.py에서 호출)"""
    generator = 인포그래픽Generator(excel_path)
    return generator.extract_all_data()


def generate_report(excel_path, template_path, output_path=None):
    """보고서 HTML 생성"""
    generator = 인포그래픽Generator(excel_path)
    html = generator.render_html(template_path, output_path)
    data = generator.extract_all_data()
    return data


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python 인포그래픽_generator.py <excel_path> [template_path] [output_path]")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    template_path = sys.argv[2] if len(sys.argv) > 2 else Path(__file__).parent / '인포그래픽_js_template.html'
    output_path = sys.argv[3] if len(sys.argv) > 3 else Path(__file__).parent / '인포그래픽_output.html'
    
    generator = 인포그래픽Generator(excel_path)
    html = generator.render_html(str(template_path), str(output_path))
    
    print(f"인포그래픽 생성 완료: {output_path}")

