#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
생산 관련 데이터 추출기
광공업생산, 서비스업생산 보도자료 데이터 추출
"""

from typing import Dict, List, Any, Optional
import pandas as pd

from .base import BaseExtractor
from .config import ALL_REGIONS, RAW_SHEET_QUARTER_COLS


class ProductionExtractor(BaseExtractor):
    """생산 관련 데이터 추출기 (광공업생산, 서비스업생산)"""
    
    # 권역별 시도 그룹
    REGION_GROUPS: Dict[str, List[str]] = {
        "수도권": ["서울", "인천", "경기"],
        "동남권": ["부산", "울산", "경남"],
        "대경권": ["대구", "경북"],
        "호남권": ["광주", "전북", "전남"],
        "충청권": ["대전", "세종", "충북", "충남"],
        "강원제주": ["강원", "제주"]
    }
    
    # 지역 표시명
    REGION_DISPLAY: Dict[str, str] = {
        '전국': '전 국', '서울': '서 울', '부산': '부 산', '대구': '대 구', '인천': '인 천',
        '광주': '광 주', '대전': '대 전', '울산': '울 산', '세종': '세 종', '경기': '경 기',
        '강원': '강 원', '충북': '충 북', '충남': '충 남', '전북': '전 북', '전남': '전 남',
        '경북': '경 북', '경남': '경 남', '제주': '제 주'
    }
    
    # =========================================================================
    # 광공업생산
    # =========================================================================
    
    def extract_manufacturing_data(self) -> Dict[str, Any]:
        """광공업생산 보도자료 데이터 추출"""
        return self._extract_production_data('광공업생산', '광공업생산지수')
    
    # =========================================================================
    # 서비스업생산
    # =========================================================================
    
    def extract_service_data(self) -> Dict[str, Any]:
        """서비스업생산 보도자료 데이터 추출"""
        return self._extract_production_data('서비스업생산', '서비스업생산지수')
    
    # =========================================================================
    # 공통 생산 데이터 추출
    # =========================================================================
    
    def _extract_production_data(self, sheet_name: str, title: str) -> Dict[str, Any]:
        """생산 관련 보도자료 공통 추출 로직"""
        report_data = self._create_base_report_data(title)
        
        # 증감률 추출
        quarterly_growth = self._extract_quarterly_growth(sheet_name)
        yearly_growth = self._extract_yearly_growth(sheet_name)
        
        # 현재 분기 데이터
        current_data = self._get_current_quarter_data(quarterly_growth)
        
        # 전국 데이터
        national_rate = current_data.get('전국')
        report_data['national_summary'] = self._create_national_summary(national_rate)
        report_data['nationwide_data'] = self._create_nationwide_data(national_rate)
        
        # 지역별 데이터
        regional_list = self._process_regional_data(current_data)
        increase_regions, decrease_regions = self._classify_regions(regional_list)
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        
        report_data['summary_box'] = {
            'main_increase_regions': increase_regions[:3],
            'main_decrease_regions': decrease_regions[:3],
            'region_count': len(increase_regions),
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
        }
        
        report_data['yearly_data'] = yearly_growth
        report_data['quarterly_data'] = quarterly_growth
        
        # 요약 테이블 생성
        config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        report_data['summary_table'] = self._generate_summary_table(
            sheet_name, config, quarterly_growth
        )
        
        return report_data
    
    def _create_base_report_data(self, title: str) -> Dict[str, Any]:
        """기본 보고서 데이터 구조 생성"""
        return {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': title,
            },
            'national_summary': {},
            'regional_data': {},
            'table_data': [],
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
            'summary_box': {},
            'nationwide_data': {},
        }
    
    def _extract_quarterly_growth(self, sheet_name: str) -> Dict[str, Dict[str, float]]:
        """분기별 전년동기비 증감률 추출"""
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        
        result = {}
        
        # 최근 4개 분기에 대해 증감률 계산 (테이블 열 순서와 일치)
        # 1. 전년동분기
        # 2. 2분기 전 (전년동분기 다음 분기)
        # 3. 직전 분기
        # 4. 현재 분기
        q2_q = self.current_quarter + 1 if self.current_quarter < 4 else 1
        q2_y = self.current_year - 1 if self.current_quarter < 4 else self.current_year
        q3_q = self.current_quarter - 1 if self.current_quarter > 1 else 4
        q3_y = self.current_year if self.current_quarter > 1 else self.current_year - 1
        
        quarters_to_extract = [
            (self.current_year - 1, self.current_quarter),  # 전년동분기
            (q2_y, q2_q),  # 2분기 전
            (q3_y, q3_q),  # 직전 분기
            (self.current_year, self.current_quarter),  # 현재 분기
        ]
        
        for year, quarter in quarters_to_extract:
            quarter_key = f"{year}.{quarter}/4"
            if quarter_key.endswith(f"{self.current_year}.{self.current_quarter}/4"):
                quarter_key += "p"  # 현재 분기에 p 추가
            
            current_col_key = f"{year}_{quarter}Q"
            prev_col_key = f"{year - 1}_{quarter}Q"
            
            current_col = config.get(current_col_key)
            prev_col = config.get(prev_col_key)
            
            if current_col is None or prev_col is None:
                continue
            
            quarter_data = {}
            for row_idx in range(len(df)):
                try:
                    region = str(df.iloc[row_idx, region_col]).strip()
                    region = self.normalize_region(region)
                    if region not in ALL_REGIONS:
                        continue
                    
                    level = df.iloc[row_idx, level_col]
                    if pd.isna(level) or str(level).strip() != '0':
                        continue
                    
                    if region in quarter_data:
                        continue
                    
                    current_val = self.safe_float(df.iloc[row_idx, current_col])
                    prev_val = self.safe_float(df.iloc[row_idx, prev_col])
                    
                    growth_rate = self.calculate_growth_rate(current_val, prev_val)
                    if growth_rate is not None:
                        quarter_data[region] = growth_rate
                except (IndexError, ValueError):
                    continue
            
            if quarter_data:
                result[quarter_key] = quarter_data
        
        return result
    
    def _extract_yearly_growth(self, sheet_name: str) -> Dict[str, Dict[str, float]]:
        """연도별 전년대비 증감률 추출"""
        # 간소화된 구현 - 필요시 확장
        return {}
    
    def _get_current_quarter_data(self, quarterly_growth: Dict) -> Dict[str, float]:
        """현재 분기 데이터 추출"""
        key_with_p = f"{self.current_year}.{self.current_quarter}/4p"
        key_without_p = f"{self.current_year}.{self.current_quarter}/4"
        return quarterly_growth.get(key_with_p, quarterly_growth.get(key_without_p, {}))
    
    def _create_national_summary(self, rate: Optional[float]) -> Dict[str, Any]:
        """전국 요약 데이터 생성"""
        if rate is None:
            return {'growth_rate': None, 'direction': 'N/A', 'trend': 'N/A'}
        
        direction = '증가' if rate > 0 else ('감소' if rate < 0 else '보합')
        trend = '확대' if rate > 0 else ('축소' if rate < 0 else '유지')
        return {'growth_rate': rate, 'direction': direction, 'trend': trend}
    
    def _create_nationwide_data(self, rate: Optional[float]) -> Dict[str, Any]:
        """전국 데이터 (템플릿 호환용) 생성"""
        return {
            'production_index': None,
            'growth_rate': rate if rate is not None else None,
            'main_industries': [],
        }
    
    def _process_regional_data(self, current_data: Dict) -> List[Dict]:
        """지역별 데이터 처리"""
        regional_list = []
        for region in ALL_REGIONS:
            if region == '전국':
                continue
            rate = current_data.get(region)
            if rate is None:
                continue
            
            direction = '증가' if rate > 0 else ('감소' if rate < 0 else '보합')
            regional_list.append({
                'region': region,
                'growth_rate': rate,
                'direction': direction,
                'industries': [],
            })
        return regional_list
    
    def _classify_regions(self, regional_list: List[Dict]) -> tuple:
        """지역을 증가/감소로 분류"""
        increase = sorted(
            [r for r in regional_list if r.get('growth_rate') and r['growth_rate'] > 0],
            key=lambda x: x['growth_rate'], reverse=True
        )
        decrease = sorted(
            [r for r in regional_list if r.get('growth_rate') and r['growth_rate'] < 0],
            key=lambda x: x['growth_rate']
        )
        return increase, decrease
    
    # =========================================================================
    # 요약 테이블 생성
    # =========================================================================
    
    def _generate_summary_table(
        self, 
        sheet_name: str, 
        config: Dict, 
        quarterly_growth: Dict
    ) -> Dict[str, Any]:
        """요약 테이블 데이터 생성"""
        # 테이블 분기 쌍 계산
        table_q_pairs = self._calculate_table_quarters()
        
        # 지역별 데이터 수집
        region_data = self._collect_region_data(quarterly_growth, table_q_pairs)
        
        # 원지수 데이터 추출
        raw_indices = self._extract_raw_indices(sheet_name, config)
        for region in ALL_REGIONS:
            if region in region_data:
                region_data[region]['indices'] = raw_indices.get(region, [None, None])
        
        # 테이블 행 생성
        rows = self._create_table_rows(region_data)
        
        return {
            'base_year': 2020,
            'columns': {
                'growth_rate_columns': [label for label in table_q_pairs],
                'index_columns': [
                    f"{self.current_year - 1}.{self.current_quarter}/4",
                    f"{self.current_year}.{self.current_quarter}/4p"
                ],
            },
            'regions': rows,
            'rows': rows,
        }
    
    def _calculate_table_quarters(self) -> List[str]:
        """테이블에 표시할 분기 라벨 계산"""
        labels = []
        
        # 1. 전년동분기
        labels.append(f"{self.current_year - 1}.{self.current_quarter}/4")
        
        # 2. 2분기 전
        q2 = self.current_quarter + 1 if self.current_quarter < 4 else 1
        y2 = self.current_year - 1 if self.current_quarter < 4 else self.current_year
        labels.append(f"{y2}.{q2}/4")
        
        # 3. 직전 분기
        q3 = self.current_quarter - 1 if self.current_quarter > 1 else 4
        y3 = self.current_year if self.current_quarter > 1 else self.current_year - 1
        labels.append(f"{y3}.{q3}/4")
        
        # 4. 현재 분기
        labels.append(f"{self.current_year}.{self.current_quarter}/4p")
        
        return labels
    
    def _collect_region_data(
        self, 
        quarterly_growth: Dict, 
        quarter_labels: List[str]
    ) -> Dict[str, Dict]:
        """지역별 데이터 수집"""
        region_data = {}
        
        for label in quarter_labels:
            # p 없는 버전도 시도
            quarter_data = quarterly_growth.get(label, quarterly_growth.get(label.rstrip('p'), {}))
            
            for region in ALL_REGIONS:
                if region not in region_data:
                    region_data[region] = {'growth_rates': [], 'indices': [None, None]}
                rate = quarter_data.get(region)
                region_data[region]['growth_rates'].append(rate)
        
        return region_data
    
    def _extract_raw_indices(self, sheet_name: str, config: Dict) -> Dict[str, List[float]]:
        """원지수 데이터 추출"""
        result = {}
        df = self._load_sheet(sheet_name)
        if df is None:
            return result
        
        prev_key = f"{self.current_year - 1}_{self.current_quarter}Q"
        curr_key = f"{self.current_year}_{self.current_quarter}Q"
        
        prev_col = config.get(prev_key)
        curr_col = config.get(curr_key)
        
        if prev_col is None or curr_col is None:
            return result
        
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        
        for row_idx in range(len(df)):
            try:
                region = str(df.iloc[row_idx, region_col]).strip()
                region = self.normalize_region(region)
                if region not in ALL_REGIONS or region in result:
                    continue
                
                level = df.iloc[row_idx, level_col]
                if pd.isna(level) or str(level).strip() not in ['0', '0.0', '총지수', '계']:
                    continue
                
                prev_val = self.safe_float(df.iloc[row_idx, prev_col])
                curr_val = self.safe_float(df.iloc[row_idx, curr_col])
                
                prev_idx = round(prev_val, 1) if prev_val is not None else None
                curr_idx = round(curr_val, 1) if curr_val is not None else None
                
                result[region] = [prev_idx, curr_idx]
            except (IndexError, ValueError):
                continue
        
        return result
    
    def _create_table_rows(self, region_data: Dict) -> List[Dict]:
        """테이블 행 생성"""
        rows = []
        
        # 전국 행
        national = region_data.get('전국', {'growth_rates': [None]*4, 'indices': [None, None]})
        rows.append({
            'region': '전 국',
            'group': None,
            'growth_rates': national['growth_rates'][:4],
            'indices': national['indices'][:2],
        })
        
        # 권역별 시도
        for group_name in ['수도권', '동남권', '대경권', '호남권', '충청권', '강원제주']:
            sidos = self.REGION_GROUPS[group_name]
            for idx, sido in enumerate(sidos):
                sido_data = region_data.get(sido, {'growth_rates': [None]*4, 'indices': [None, None]})
                
                row = {
                    'region': self.REGION_DISPLAY.get(sido, sido),
                    'growth_rates': sido_data['growth_rates'][:4],
                    'indices': sido_data['indices'][:2],
                }
                
                if idx == 0:
                    row['group'] = group_name
                    row['rowspan'] = len(sidos)
                else:
                    row['group'] = None
                
                rows.append(row)
        
        return rows
