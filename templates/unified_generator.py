#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
통합 보고서 Generator V2
모든 부문의 보고서를 생성하는 범용 Generator
집계 시트 기반, 완전 동적 매핑
"""

import pandas as pd
from typing import Dict, Any, List, Optional
from pathlib import Path

try:
    from .base_generator import BaseGenerator
    from config.report_configs import (
        get_report_config, REPORT_CONFIGS,
        REGION_DISPLAY_MAPPING, REGION_GROUPS, VALID_REGIONS
    )
except ImportError:
    import sys
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from templates.base_generator import BaseGenerator
    from config.report_configs import (
        get_report_config, REPORT_CONFIGS,
        REGION_DISPLAY_MAPPING, REGION_GROUPS, VALID_REGIONS
    )


class UnifiedReportGenerator(BaseGenerator):
    """
    통합 보고서 Generator (집계 시트 기반)
    
    mining_manufacturing_generator의 검증된 로직을 기반으로 구현
    """
    
    # 집계 시트 구조
    DATA_START_ROW = 3
    
    def __init__(self, report_type: str, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__(excel_path, year, quarter, excel_file)
        
        # 설정 로드
        self.config = get_report_config(report_type)
        if not self.config:
            raise ValueError(f"Unknown report type: {report_type}")
        
        self.report_type = report_type
        self.report_id = self.config['report_id']
        self.name_mapping = self.config.get('name_mapping', {})
        
        # 집계 시트 구조 (설정에서 로드)
        agg_struct = self.config.get('aggregation_structure', {})
        self.region_name_col = agg_struct.get('region_name_col', 4)
        self.industry_code_col = agg_struct.get('industry_code_col', 7)
        self.total_code = agg_struct.get('total_code', 'BCD')
        
        # 집계 시트
        self.df_aggregation = None
        self.target_col = None
        self.prev_y_col = None
        
        print(f"[{self.config['name']}] Generator 초기화")
    
    def _get_region_display_name(self, region: str) -> str:
        """지역명 변환"""
        return REGION_DISPLAY_MAPPING.get(region, region)
    
    def load_data(self):
        """집계 시트 로드"""
        xl = self.load_excel()
        
        # 집계 시트 찾기
        agg_sheets = self.config['sheets']['aggregation']
        fallback_sheets = self.config['sheets']['fallback']
        
        agg_sheet, _ = self.find_sheet_with_fallback(agg_sheets, fallback_sheets)
        
        if not agg_sheet:
            raise ValueError(f"[{self.config['name']}] 집계 시트를 찾을 수 없습니다")
        
        self.df_aggregation = self.get_sheet(agg_sheet)
        print(f"[{self.config['name']}] ✅ 집계 시트: '{agg_sheet}' ({len(self.df_aggregation)}행 × {len(self.df_aggregation.columns)}열)")
        
        # 동적 컬럼 찾기
        self._find_data_columns()
    
    def _find_data_columns(self):
        """데이터 컬럼 동적 탐색"""
        df = self.df_aggregation
        
        # 헤더 행에서 기간 컬럼 찾기
        for header_idx in range(min(5, len(df))):
            header_row = df.iloc[header_idx]
            
            # target_col 찾기
            if self.target_col is None:
                try:
                    self.target_col = self.find_target_col_index(header_row, self.year, self.quarter)
                    if self.target_col is not None:
                        print(f"[{self.config['name']}] ✅ Target 컬럼: {self.target_col} ({self.year} {self.quarter}/4)")
                except ValueError:
                    pass  # 이 행에서는 못 찾음, 다음 행 시도
            
            # prev_y_col 찾기
            if self.prev_y_col is None:
                try:
                    self.prev_y_col = self.find_target_col_index(header_row, self.year - 1, self.quarter)
                    if self.prev_y_col is not None:
                        print(f"[{self.config['name']}] ✅ 전년 컬럼: {self.prev_y_col} ({self.year - 1} {self.quarter}/4)")
                except ValueError:
                    pass  # 이 행에서는 못 찾음, 다음 행 시도
        
        # Fallback
        if self.target_col is None:
            self.target_col = len(df.columns) - 1
            print(f"[{self.config['name']}] ⚠️ Target 컬럼 fallback: {self.target_col}")
        
        if self.prev_y_col is None:
            self.prev_y_col = self.target_col - 4
            print(f"[{self.config['name']}] ⚠️ 전년 컬럼 fallback: {self.prev_y_col}")
    
    def _extract_table_data_ssot(self) -> List[Dict[str, Any]]:
        """
        집계 시트에서 전국 + 17개 시도 데이터 추출 (SSOT)
        mining_manufacturing_generator의 로직 그대로 사용
        """
        df = self.df_aggregation
        
        # 데이터 행만 (헤더 제외)
        data_df = df.iloc[self.DATA_START_ROW:].copy() if self.DATA_START_ROW < len(df) else df
        
        # 지역 목록
        regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        table_data = []
        
        for region in regions:
            # 지역명으로 필터링 (설정에서 가져온 컬럼 사용)
            region_filter = data_df[
                data_df.iloc[:, self.region_name_col].astype(str).str.strip() == region
            ]
            
            if region_filter.empty:
                continue
            
            # 총지수 행 찾기 (설정에서 가져온 컬럼 및 코드 사용)
            # 디버깅: 실제 코드 값 확인
            if region == '전국':
                industry_codes = region_filter.iloc[:, self.industry_code_col].astype(str).head(5).tolist()
                print(f"[{self.config['name']}] 디버그: {region} 산업코드 (처음 5개): {industry_codes}")
                print(f"[{self.config['name']}] 디버그: 찾으려는 코드: '{self.total_code}'")
            
            region_total = region_filter[
                region_filter.iloc[:, self.industry_code_col].astype(str).str.contains(self.total_code, na=False, regex=False)
            ]
            
            if region_total.empty:
                # Fallback: 첫 번째 행
                print(f"[{self.config['name']}] ⚠️ {region}: 코드 '{self.total_code}' 찾기 실패, 첫 번째 행 사용")
                region_total = region_filter.head(1)
            
            if region_total.empty:
                continue
            
            row = region_total.iloc[0]
            
            # 지수 추출
            idx_current = self.safe_float(row.iloc[self.target_col] if self.target_col < len(row) else None, None)
            idx_prev_year = self.safe_float(row.iloc[self.prev_y_col] if self.prev_y_col < len(row) else None, None)
            
            if idx_current is None:
                continue
            
            # 증감률 계산
            if idx_prev_year and idx_prev_year != 0:
                change_rate = round(((idx_current - idx_prev_year) / idx_prev_year) * 100, 1)
            else:
                change_rate = None
            
            table_data.append({
                'region_name': region,
                'region_display': self._get_region_display_name(region),
                'value': round(idx_current, 1),
                'prev_value': round(idx_prev_year, 1) if idx_prev_year else None,
                'change_rate': change_rate
            })
            
            print(f"[{self.config['name']}] ✅ {region}: 지수={idx_current:.1f}, 증감률={change_rate}%")
        
        return table_data
    
    def extract_nationwide_data(self, table_data: List[Dict] = None) -> Dict[str, Any]:
        """전국 데이터 추출 - 템플릿 호환 필드명"""
        if table_data is None:
            table_data = self._extract_table_data_ssot()
        
        nationwide = next((d for d in table_data if d['region_name'] == '전국'), None)
        
        if not nationwide:
            return {
                'production_index': 100.0,
                'sales_index': 100.0,
                'service_index': 100.0,
                'growth_rate': 0.0,
                'main_industries': [],
                'main_businesses': [],
                'main_items': []
            }
        
        # 보고서 타입별 지수명 결정
        index_value = nationwide['value']
        growth_rate = nationwide['change_rate'] if nationwide['change_rate'] else 0.0
        
        # 모든 필드명 포함 (템플릿 호환)
        empty_items = []  # TODO: 업종별 데이터 추가
        
        return {
            'production_index': index_value,
            'sales_index': index_value,
            'service_index': index_value,
            'growth_rate': growth_rate,
            'main_industries': empty_items,
            'main_increase_industries': empty_items,
            'main_decrease_industries': empty_items,
            'main_businesses': empty_items,
            'main_items': empty_items
        }
    
    def extract_regional_data(self, table_data: List[Dict] = None) -> Dict[str, Any]:
        """시도별 데이터 추출"""
        if table_data is None:
            table_data = self._extract_table_data_ssot()
        
        # 전국 제외
        regional = [d for d in table_data if d['region_name'] != '전국']
        
        # 증가/감소 분류
        increase = [r for r in regional if r['change_rate'] and r['change_rate'] > 0]
        decrease = [r for r in regional if r['change_rate'] and r['change_rate'] < 0]
        
        increase.sort(key=lambda x: x['change_rate'], reverse=True)
        decrease.sort(key=lambda x: x['change_rate'])
        
        return {
            'increase_regions': increase,
            'decrease_regions': decrease,
            'all_regions': regional
        }
    
    def _generate_summary_table(self, table_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """summary_table 생성 (템플릿 호환 형식)"""
        # 지역 표시명 매핑 (띄어쓰기 포함)
        region_display = {
            '전국': '전 국', '서울': '서 울', '부산': '부 산', '대구': '대 구', '인천': '인 천',
            '광주': '광 주', '대전': '대 전', '울산': '울 산', '세종': '세 종', '경기': '경 기',
            '강원': '강 원', '충북': '충 북', '충남': '충 남', '전북': '전 북', '전남': '전 남',
            '경북': '경 북', '경남': '경 남', '제주': '제 주'
        }
        
        # 지역 그룹 정보
        REGION_GROUPS = {
            '경인': ['서울', '인천', '경기'],
            '충청': ['대전', '세종', '충북', '충남'],
            '호남': ['광주', '전북', '전남', '제주'],
            '동북': ['대구', '경북', '강원'],
            '동남': ['부산', '울산', '경남']
        }
        
        regions_list = []
        
        # 전국 먼저 추가
        nationwide = next((d for d in table_data if d['region_name'] == '전국'), None)
        if nationwide:
            regions_list.append({
                'region': region_display.get('전국', '전 국'),
                'growth_rates': [None, None, None, nationwide.get('change_rate', 0.0)],
                'indices': [nationwide.get('prev_value', 0.0) if nationwide.get('prev_value') else 0.0,
                           nationwide.get('value', 0.0)],
                'group': None,
                'rowspan': None
            })
        
        # 권역별 시도 추가
        region_group_order = ['경인', '충청', '호남', '동북', '동남']
        
        for group_name in region_group_order:
            sidos = REGION_GROUPS.get(group_name, [])
            for idx, sido in enumerate(sidos):
                region_data = next((d for d in table_data if d['region_name'] == sido), None)
                if not region_data:
                    continue
                
                row_data = {
                    'region': region_display.get(sido, sido),
                    'growth_rates': [None, None, None, region_data.get('change_rate', 0.0)],
                    'indices': [region_data.get('prev_value', 0.0) if region_data.get('prev_value') else 0.0,
                               region_data.get('value', 0.0)],
                    'group': group_name if idx == 0 else None,
                    'rowspan': len(sidos) if idx == 0 else None
                }
                regions_list.append(row_data)
        
        # base_year는 설정에서 가져오거나 기본값 사용
        base_year = self.config.get('base_year', 2020)
        
        # columns 정보 (연도/분기 기반)
        growth_rate_columns = [
            f"{self.year - 2}.2/4",
            f"{self.year - 1}.2/4",
            f"{self.year}.1/4",
            f"{self.year}.2/4p"
        ]
        index_columns = [
            f"{self.year - 1}.2/4",
            f"{self.year}.2/4p"
        ]
        
        return {
            'base_year': base_year,
            'columns': {
                'growth_rate_columns': growth_rate_columns,
                'index_columns': index_columns
            },
            'regions': regions_list
        }
    
    def extract_all_data(self) -> Dict[str, Any]:
        """전체 데이터 추출 - 템플릿 호환 구조"""
        # 데이터 로드
        self.load_data()
        
        # Table Data (SSOT)
        table_data = self._extract_table_data_ssot()
        
        # Text Data
        nationwide = self.extract_nationwide_data(table_data)
        regional = self.extract_regional_data(table_data)
        
        # Top3 regions (템플릿 호환 필드명으로 생성)
        top3_increase = []
        for r in regional['increase_regions'][:3]:
            top3_increase.append({
                'region': r['region_name'],
                'growth_rate': r['change_rate'] if r['change_rate'] else 0.0,
                'industries': []  # TODO: 업종별 데이터 추가
            })
        
        top3_decrease = []
        for r in regional['decrease_regions'][:3]:
            top3_decrease.append({
                'region': r['region_name'],
                'growth_rate': r['change_rate'] if r['change_rate'] else 0.0,
                'industries': [],  # TODO: 업종별 데이터 추가
                'main_business': ''  # TODO: 소비동향용 주요 업태
            })
        
        # Summary Box
        summary_box = {
            'main_regions': [{'region': r['region'], 'items': r['industries']} for r in top3_increase],
            'region_count': len(regional['increase_regions'])
        }
        
        # Regional data 필드명 변환 (템플릿 호환)
        regional_converted = {
            'increase_regions': [
                {
                    'region': r['region_name'],
                    'growth_rate': r['change_rate'] if r['change_rate'] else 0.0,
                    'value': r['value'],
                    'top_industries': []  # TODO: 업종별 데이터 추가
                }
                for r in regional['increase_regions']
            ],
            'decrease_regions': [
                {
                    'region': r['region_name'],
                    'growth_rate': r['change_rate'] if r['change_rate'] else 0.0,
                    'value': r['value'],
                    'top_industries': []  # TODO: 업종별 데이터 추가
                }
                for r in regional['decrease_regions']
            ],
            'all_regions': regional['all_regions']
        }
        
        # Table data 필드명 변환 (템플릿 호환)
        table_data_converted = [
            {
                'region': r['region_name'],
                'region_name': r['region_name'],
                'value': r['value'],
                'growth_rate': r['change_rate'] if r['change_rate'] else 0.0,
                'change_rate': r['change_rate'] if r['change_rate'] else 0.0
            }
            for r in table_data
        ]
        
        # summary_table 생성 (템플릿에서 필요)
        summary_table = self._generate_summary_table(table_data)
        
        # 템플릿별 추가 필드 (TODO: 실제 데이터로 대체)
        extra_fields = {}
        
        # 소비동향용 필드
        if self.report_type == 'consumption':
            extra_fields['increase_businesses_text'] = ''  # TODO: 업태별 데이터 추가
            extra_fields['decrease_businesses_text'] = ''  # TODO: 업태별 데이터 추가
        
        # 서비스업용 필드
        elif self.report_type == 'service':
            extra_fields['increase_industries_text'] = ''  # TODO: 업종별 데이터 추가
            extra_fields['decrease_industries_text'] = ''  # TODO: 업종별 데이터 추가
        
        # footer_info 생성 (템플릿에서 필요)
        footer_source_map = {
            'mining': '자료: 국가데이터처 국가통계포털(KOSIS), 광업제조업동향조사',
            'service': '자료: 국가데이터처 국가통계포털(KOSIS), 서비스업동향조사',
            'consumption': '자료: 국가데이터처 국가통계포털(KOSIS), 서비스업동향조사'
        }
        footer_info = {
            'source': footer_source_map.get(self.report_type, '자료: 국가데이터처'),
            'page_num': '- 1 -'  # 기본값, 필요시 설정에서 가져올 수 있음
        }
        
        return {
            'report_info': {
                'year': self.year,
                'quarter': self.quarter,
                'report_type': self.report_type,
                'report_name': self.config['name'],
                'index_name': self.config.get('index_name', '지수'),
                'item_name': self.config.get('item_name', '항목'),
                'period': f"{self.year}년 {self.quarter}분기"
            },
            'footer_info': footer_info,  # 템플릿에서 필요
            'summary_box': summary_box,
            'nationwide_data': nationwide,
            'regional_data': regional_converted,
            'table_data': table_data_converted,
            'summary_table': summary_table,  # 템플릿에서 필요
            'top3_increase_regions': top3_increase,
            'top3_decrease_regions': top3_decrease,
            **extra_fields  # 추가 필드 병합
        }


# 하위 호환성 Wrapper
class MiningManufacturingGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('mining', excel_path, year, quarter, excel_file)


class ServiceIndustryGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('service', excel_path, year, quarter, excel_file)


class ConsumptionGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('consumption', excel_path, year, quarter, excel_file)


if __name__ == '__main__':
    # 테스트
    base_path = Path(__file__).parent.parent
    excel_path = base_path / '분석표_25년 3분기_캡스톤(업데이트).xlsx'
    
    print("=" * 70)
    print("통합 Generator V2 테스트 (집계 시트 기반)")
    print("=" * 70)
    
    for report_type in ['mining', 'service', 'consumption']:
        print(f"\n{'='*70}")
        print(f"[TEST] {REPORT_CONFIGS[report_type]['name']}")
        print(f"{'='*70}\n")
        
        try:
            generator = UnifiedReportGenerator(report_type, str(excel_path), 2025, 3)
            data = generator.extract_all_data()
            
            # 결과 출력
            print(f"\n[결과] ✅ 데이터 추출 완료")
            nationwide = data['nationwide_data']
            print(f"  전국: 지수={nationwide['production_index']:.1f}, 증감률={nationwide['growth_rate']}%")
            
            regional = data['regional_data']
            print(f"  지역: 증가={len(regional['increase_regions'])}개, 감소={len(regional['decrease_regions'])}개")
            
            if regional['increase_regions']:
                top = regional['increase_regions'][0]
                print(f"  최고: {top['region_name']} ({top['change_rate']}%)")
            
        except Exception as e:
            print(f"\n[오류] ❌ {e}")
            import traceback
            traceback.print_exc()
