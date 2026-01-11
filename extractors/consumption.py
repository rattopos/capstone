#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
소비/건설 관련 데이터 추출기
소비동향, 건설동향 보도자료 데이터 추출
"""

from typing import Dict, List, Any, Optional
import pandas as pd

from .base import BaseExtractor
from .config import ALL_REGIONS, RAW_SHEET_QUARTER_COLS


class ConsumptionConstructionExtractor(BaseExtractor):
    """소비/건설 관련 데이터 추출기"""
    
    REGION_GROUPS = {
        "수도권": ["서울", "인천", "경기"],
        "동남권": ["부산", "울산", "경남"],
        "대경권": ["대구", "경북"],
        "호남권": ["광주", "전북", "전남"],
        "충청권": ["대전", "세종", "충북", "충남"],
        "강원제주": ["강원", "제주"]
    }
    
    REGION_DISPLAY = {
        '전국': '전 국', '서울': '서 울', '부산': '부 산', '대구': '대 구', '인천': '인 천',
        '광주': '광 주', '대전': '대 전', '울산': '울 산', '세종': '세 종', '경기': '경 기',
        '강원': '강 원', '충북': '충 북', '충남': '충 남', '전북': '전 북', '전남': '전 남',
        '경북': '경 북', '경남': '경 남', '제주': '제 주'
    }
    
    def extract_consumption_data(self) -> Dict[str, Any]:
        """소비동향 보도자료 데이터 추출"""
        sheet_name = '소비(소매, 추가)'
        config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '소비동향',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
            'summary_box': {},
            'nationwide_data': {},
        }
        
        quarterly_growth = self._extract_quarterly_growth(sheet_name, config)
        current_data = self._get_current_quarter_data(quarterly_growth)
        
        national_rate = current_data.get('전국')
        report_data['national_summary'] = {
            'growth_rate': national_rate,
            'direction': self._get_direction(national_rate),
        }
        report_data['nationwide_data'] = {
            'sales_index': None,
            'growth_rate': national_rate,
            'main_categories': [],
        }
        
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
        }
        
        report_data['quarterly_data'] = quarterly_growth
        report_data['summary_table'] = self._generate_summary_table(sheet_name, config, quarterly_growth)
        
        return report_data
    
    def extract_construction_data(self) -> Dict[str, Any]:
        """건설동향 보도자료 데이터 추출"""
        sheet_name = '건설 (공표자료)'
        config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '건설동향',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
            'summary_box': {},
            'nationwide_data': {},
        }
        
        quarterly_growth = self._extract_quarterly_growth(sheet_name, config)
        current_data = self._get_current_quarter_data(quarterly_growth)
        
        # 토목/건축 데이터 추출
        civil_building_data = self._extract_civil_building_data(sheet_name, config)
        
        national_rate = current_data.get('전국')
        report_data['national_summary'] = {
            'growth_rate': national_rate,
            'direction': self._get_direction(national_rate),
        }
        
        # nationwide_data에 토목/건축 증감률 추가
        national_civil_building = civil_building_data.get('전국', {})
        report_data['nationwide_data'] = {
            'order_amount': None,
            'growth_rate': national_rate,
            'main_sectors': [],
            'civil_growth': national_civil_building.get('civil_growth'),
            'building_growth': national_civil_building.get('building_growth'),
            'civil_subtypes': '철도·궤도, 기계설치' if (national_civil_building.get('civil_growth') or 0) > 0 else '토지조성, 치산·치수',
            'building_subtypes': '주택, 관공서 등',
        }
        
        regional_list = self._process_regional_data(current_data)
        
        # 지역별 토목/건축 데이터 추가
        for region_data in regional_list:
            region = region_data['region']
            cb_data = civil_building_data.get(region, {})
            region_data['civil_growth'] = cb_data.get('civil_growth')
            region_data['building_growth'] = cb_data.get('building_growth')
            region_data['civil_subtypes'] = '철도·궤도, 도로·교량' if (cb_data.get('civil_growth') or 0) > 0 else '토지조성, 치산·치수'
            region_data['building_subtypes'] = '주택, 관공서 등' if (cb_data.get('building_growth') or 0) > 0 else '사무실·점포, 관공서 등'
        
        increase_regions, decrease_regions = self._classify_regions(regional_list)
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        
        report_data['quarterly_data'] = quarterly_growth
        report_data['summary_table'] = self._generate_summary_table(sheet_name, config, quarterly_growth)
        
        return report_data
    
    def _extract_civil_building_data(self, sheet_name: str, config: Dict) -> Dict[str, Dict]:
        """토목/건축 증감률 추출"""
        result = {}
        df = self._load_sheet(sheet_name)
        if df is None:
            return result
        
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        code_col = 3  # 분류 코드 열
        
        curr_key = f"{self.current_year}_{self.current_quarter}Q"
        prev_key = f"{self.current_year - 1}_{self.current_quarter}Q"
        curr_col = config.get(curr_key)
        prev_col = config.get(prev_key)
        
        if curr_col is None or prev_col is None:
            print(f"[건설 토목/건축] 열 인덱스를 찾을 수 없음: curr={curr_key}, prev={prev_key}")
            return result
        
        current_region = None
        region_data = {}
        
        for row_idx in range(len(df)):
            try:
                region = str(df.iloc[row_idx, region_col]).strip()
                region = self.normalize_region(region)
                level = str(df.iloc[row_idx, level_col]).strip()
                code = str(df.iloc[row_idx, code_col]).strip()
                
                # 총계 행 (level='0') - 새 지역 시작
                if level == '0' and code == '0':
                    if region in ALL_REGIONS:
                        current_region = region
                        if current_region not in result:
                            result[current_region] = {}
                
                # 건축(code='1') 또는 토목(code='2') 행
                if current_region and level == '1' and code in ['1', '2']:
                    curr_val = self.safe_float(df.iloc[row_idx, curr_col])
                    prev_val = self.safe_float(df.iloc[row_idx, prev_col])
                    rate = self.calculate_growth_rate(curr_val, prev_val)
                    
                    if code == '1':  # 건축
                        result[current_region]['building_growth'] = rate
                    elif code == '2':  # 토목
                        result[current_region]['civil_growth'] = rate
                        
            except (IndexError, ValueError) as e:
                continue
        
        return result
    
    def _extract_quarterly_growth(self, sheet_name: str, config: Dict) -> Dict:
        """분기별 증감률 추출"""
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        result = {}
        
        # 4개 분기 추출 (테이블 열 순서와 일치)
        q2_q = self.current_quarter + 1 if self.current_quarter < 4 else 1
        q2_y = self.current_year - 1 if self.current_quarter < 4 else self.current_year
        q3_q = self.current_quarter - 1 if self.current_quarter > 1 else 4
        q3_y = self.current_year if self.current_quarter > 1 else self.current_year - 1
        
        quarters = [
            (self.current_year - 1, self.current_quarter),  # 전년동분기
            (q2_y, q2_q),  # 2분기 전
            (q3_y, q3_q),  # 직전 분기
            (self.current_year, self.current_quarter),  # 현재 분기
        ]
        
        for year, quarter in quarters:
            label = f"{year}.{quarter}/4"
            if year == self.current_year and quarter == self.current_quarter:
                label += "p"
            
            current_key = f"{year}_{quarter}Q"
            prev_key = f"{year - 1}_{quarter}Q"
            current_col = config.get(current_key)
            prev_col = config.get(prev_key)
            
            if current_col is None or prev_col is None:
                continue
            
            quarter_data = {}
            for row_idx in range(len(df)):
                try:
                    region = str(df.iloc[row_idx, region_col]).strip()
                    region = self.normalize_region(region)
                    if region not in ALL_REGIONS or region in quarter_data:
                        continue
                    
                    level = df.iloc[row_idx, level_col]
                    if pd.isna(level) or str(level).strip() != '0':
                        continue
                    
                    curr_val = self.safe_float(df.iloc[row_idx, current_col])
                    prev_val = self.safe_float(df.iloc[row_idx, prev_col])
                    rate = self.calculate_growth_rate(curr_val, prev_val)
                    if rate is not None:
                        quarter_data[region] = rate
                except (IndexError, ValueError):
                    continue
            
            if quarter_data:
                result[label] = quarter_data
        
        return result
    
    def _get_current_quarter_data(self, quarterly_growth: Dict) -> Dict:
        key_p = f"{self.current_year}.{self.current_quarter}/4p"
        key = f"{self.current_year}.{self.current_quarter}/4"
        return quarterly_growth.get(key_p, quarterly_growth.get(key, {}))
    
    def _get_direction(self, rate: Optional[float]) -> str:
        if rate is None:
            return 'N/A'
        return '증가' if rate > 0 else ('감소' if rate < 0 else '보합')
    
    def _process_regional_data(self, current_data: Dict) -> List[Dict]:
        regional_list = []
        for region in ALL_REGIONS:
            if region == '전국':
                continue
            rate = current_data.get(region)
            if rate is None:
                continue
            regional_list.append({
                'region': region,
                'growth_rate': rate,
                'direction': self._get_direction(rate),
            })
        return regional_list
    
    def _classify_regions(self, regional_list: List[Dict]) -> tuple:
        increase = sorted(
            [r for r in regional_list if r.get('growth_rate') and r['growth_rate'] > 0],
            key=lambda x: x['growth_rate'], reverse=True
        )
        decrease = sorted(
            [r for r in regional_list if r.get('growth_rate') and r['growth_rate'] < 0],
            key=lambda x: x['growth_rate']
        )
        return increase, decrease
    
    def _generate_summary_table(self, sheet_name: str, config: Dict, quarterly_growth: Dict) -> Dict:
        quarter_labels = self._calculate_quarters()
        region_data = {}
        
        for label in quarter_labels:
            q_data = quarterly_growth.get(label, quarterly_growth.get(label.rstrip('p'), {}))
            for region in ALL_REGIONS:
                if region not in region_data:
                    region_data[region] = {'growth_rates': [], 'indices': [None, None]}
                region_data[region]['growth_rates'].append(q_data.get(region))
        
        raw_indices = self._extract_raw_indices(sheet_name, config)
        for region in ALL_REGIONS:
            if region in region_data:
                region_data[region]['indices'] = raw_indices.get(region, [None, None])
        
        rows = self._create_rows(region_data)
        
        return {
            'columns': {
                'growth_rate_columns': quarter_labels,
                'index_columns': [
                    f"{self.current_year - 1}.{self.current_quarter}/4",
                    f"{self.current_year}.{self.current_quarter}/4p"
                ],
            },
            'rows': rows,
            'regions': rows,
        }
    
    def _calculate_quarters(self) -> List[str]:
        labels = []
        labels.append(f"{self.current_year - 1}.{self.current_quarter}/4")
        q2 = self.current_quarter + 1 if self.current_quarter < 4 else 1
        y2 = self.current_year - 1 if self.current_quarter < 4 else self.current_year
        labels.append(f"{y2}.{q2}/4")
        q3 = self.current_quarter - 1 if self.current_quarter > 1 else 4
        y3 = self.current_year if self.current_quarter > 1 else self.current_year - 1
        labels.append(f"{y3}.{q3}/4")
        labels.append(f"{self.current_year}.{self.current_quarter}/4p")
        return labels
    
    def _extract_raw_indices(self, sheet_name: str, config: Dict) -> Dict:
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
                if pd.isna(level) or str(level).strip() not in ['0', '0.0']:
                    continue
                
                prev_val = self.safe_float(df.iloc[row_idx, prev_col])
                curr_val = self.safe_float(df.iloc[row_idx, curr_col])
                
                result[region] = [
                    round(prev_val, 1) if prev_val else None,
                    round(curr_val, 1) if curr_val else None
                ]
            except (IndexError, ValueError):
                continue
        
        return result
    
    def _create_rows(self, region_data: Dict) -> List[Dict]:
        rows = []
        national = region_data.get('전국', {'growth_rates': [None]*4, 'indices': [None, None]})
        rows.append({
            'region': '전 국',
            'group': None,
            'growth_rates': national['growth_rates'][:4],
            'indices': national['indices'][:2],
        })
        
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
