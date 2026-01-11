#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
고용/인구 관련 데이터 추출기
고용률, 실업률, 국내인구이동 보도자료 데이터 추출
"""

from typing import Dict, List, Any, Optional
import pandas as pd

from .base import BaseExtractor
from .config import ALL_REGIONS, RAW_SHEET_QUARTER_COLS


class EmploymentPopulationExtractor(BaseExtractor):
    """고용/인구 관련 데이터 추출기"""
    
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
    
    def extract_employment_rate_data(self) -> Dict[str, Any]:
        """고용률 보도자료 데이터 추출"""
        sheet_name = '연령별고용률'
        config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '고용률',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
            'summary_box': {},
            'nationwide_data': {},
        }
        
        # 고용률은 %p 차이로 계산
        quarterly_diff = self._extract_quarterly_diff(sheet_name, config)
        current_data = self._get_current_quarter_data(quarterly_diff)
        
        national_diff = current_data.get('전국')
        report_data['national_summary'] = {
            'change': national_diff,
            'direction': self._get_direction(national_diff),
        }
        report_data['nationwide_data'] = {
            'rate': None,
            'change': national_diff,
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
        
        report_data['quarterly_data'] = quarterly_diff
        report_data['summary_table'] = self._generate_employment_summary_table(sheet_name, config, quarterly_diff)
        
        return report_data
    
    def extract_unemployment_data(self) -> Dict[str, Any]:
        """실업률 보도자료 데이터 추출"""
        sheet_name = '실업자 수'
        config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '실업률',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
            'summary_box': {},
            'nationwide_data': {},
        }
        
        quarterly_diff = self._extract_quarterly_diff(sheet_name, config)
        current_data = self._get_current_quarter_data(quarterly_diff)
        
        national_diff = current_data.get('전국')
        report_data['national_summary'] = {
            'change': national_diff,
            'direction': self._get_direction(national_diff, reverse=True),
        }
        report_data['nationwide_data'] = {
            'rate': None,
            'change': national_diff,
        }
        
        regional_list = self._process_regional_data(current_data, reverse=True)
        increase_regions, decrease_regions = self._classify_regions(regional_list)
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        
        report_data['quarterly_data'] = quarterly_diff
        report_data['summary_table'] = self._generate_unemployment_summary_table(sheet_name, config, quarterly_diff)
        
        return report_data
    
    def extract_population_migration_data(self) -> Dict[str, Any]:
        """국내인구이동 보도자료 데이터 추출"""
        sheet_name = '시도 간 이동'
        config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '국내인구이동',
            },
            'regional_data': {},
            'top3_inflow_regions': [],
            'top3_outflow_regions': [],
            'summary_box': {},
        }
        
        # 순이동 데이터 추출 (절대값)
        quarterly_migration = self._extract_migration_data(sheet_name, config)
        current_data = self._get_current_quarter_data(quarterly_migration)
        
        regional_list = []
        for region in ALL_REGIONS:
            if region == '전국':
                continue
            migration = current_data.get(region)
            if migration is None:
                continue
            regional_list.append({
                'region': region,
                'net_migration': migration,
                'direction': '순유입' if migration > 0 else ('순유출' if migration < 0 else '없음'),
            })
        
        inflow = sorted([r for r in regional_list if r.get('net_migration') and r['net_migration'] > 0],
                       key=lambda x: x['net_migration'], reverse=True)
        outflow = sorted([r for r in regional_list if r.get('net_migration') and r['net_migration'] < 0],
                        key=lambda x: x['net_migration'])
        
        report_data['top3_inflow_regions'] = inflow[:3]
        report_data['top3_outflow_regions'] = outflow[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'inflow_regions': inflow,
            'outflow_regions': outflow,
        }
        
        report_data['quarterly_data'] = quarterly_migration
        report_data['summary_table'] = self._generate_migration_summary_table(sheet_name, config, quarterly_migration)
        
        return report_data
    
    def _extract_quarterly_diff(self, sheet_name: str, config: Dict) -> Dict:
        """분기별 전년동기 대비 차이 (%p) 추출"""
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        total_code = config.get('total_code', '계')
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
                    
                    # 분류 확인 (고용률은 '계', 실업자수는 컬럼 구조 다름)
                    if level_col < len(df.columns):
                        level = df.iloc[row_idx, level_col]
                        if pd.notna(level) and str(level).strip() != total_code and total_code != '계':
                            continue
                    
                    curr_val = self.safe_float(df.iloc[row_idx, current_col])
                    prev_val = self.safe_float(df.iloc[row_idx, prev_col])
                    diff = self.calculate_difference(curr_val, prev_val)
                    if diff is not None:
                        quarter_data[region] = diff
                except (IndexError, ValueError):
                    continue
            
            if quarter_data:
                result[label] = quarter_data
        
        return result
    
    def _extract_migration_data(self, sheet_name: str, config: Dict) -> Dict:
        """분기별 순이동 데이터 추출 (절대값)"""
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        region_col = config.get('region_col', 1)
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
            
            col_key = f"{year}_{quarter}Q"
            col_idx = config.get(col_key)
            
            if col_idx is None:
                continue
            
            quarter_data = {}
            for row_idx in range(len(df)):
                try:
                    region = str(df.iloc[row_idx, region_col]).strip()
                    region = self.normalize_region(region)
                    if region not in ALL_REGIONS or region == '전국' or region in quarter_data:
                        continue
                    
                    value = self.safe_float(df.iloc[row_idx, col_idx])
                    if value is not None:
                        # 천 명 단위로 변환
                        quarter_data[region] = round(value / 1000, 1)
                except (IndexError, ValueError):
                    continue
            
            if quarter_data:
                result[label] = quarter_data
        
        return result
    
    def _get_current_quarter_data(self, quarterly_data: Dict) -> Dict:
        key_p = f"{self.current_year}.{self.current_quarter}/4p"
        key = f"{self.current_year}.{self.current_quarter}/4"
        return quarterly_data.get(key_p, quarterly_data.get(key, {}))
    
    def _get_direction(self, value: Optional[float], reverse: bool = False) -> str:
        if value is None:
            return 'N/A'
        if reverse:
            return '개선' if value < 0 else ('악화' if value > 0 else '유지')
        return '상승' if value > 0 else ('하락' if value < 0 else '유지')
    
    def _process_regional_data(self, current_data: Dict, reverse: bool = False) -> List[Dict]:
        regional_list = []
        for region in ALL_REGIONS:
            if region == '전국':
                continue
            value = current_data.get(region)
            if value is None:
                continue
            regional_list.append({
                'region': region,
                'change': value,
                'direction': self._get_direction(value, reverse),
            })
        return regional_list
    
    def _classify_regions(self, regional_list: List[Dict]) -> tuple:
        increase = sorted(
            [r for r in regional_list if r.get('change') and r['change'] > 0],
            key=lambda x: x['change'], reverse=True
        )
        decrease = sorted(
            [r for r in regional_list if r.get('change') and r['change'] < 0],
            key=lambda x: x['change']
        )
        return increase, decrease
    
    def _generate_employment_summary_table(self, sheet_name: str, config: Dict, quarterly_diff: Dict) -> Dict:
        """고용률 요약 테이블 생성"""
        quarter_labels = self._calculate_quarters()
        region_data = {}
        
        for label in quarter_labels:
            q_data = quarterly_diff.get(label, quarterly_diff.get(label.rstrip('p'), {}))
            for region in ALL_REGIONS:
                if region not in region_data:
                    region_data[region] = {'changes': [], 'rates': [None, None]}
                region_data[region]['changes'].append(q_data.get(region))
        
        raw_rates = self._extract_raw_values(sheet_name, config)
        for region in ALL_REGIONS:
            if region in region_data:
                region_data[region]['rates'] = raw_rates.get(region, [None, None])
        
        rows = self._create_employment_rows(region_data)
        
        return {
            'columns': {
                'change_columns': quarter_labels,
                'rate_columns': [
                    f"{self.current_year - 1}.{self.current_quarter}/4",
                    f"{self.current_year}.{self.current_quarter}/4p"
                ],
            },
            'rows': rows,
            'regions': rows,
        }
    
    def _generate_unemployment_summary_table(self, sheet_name: str, config: Dict, quarterly_diff: Dict) -> Dict:
        """실업률 요약 테이블 생성"""
        return self._generate_employment_summary_table(sheet_name, config, quarterly_diff)
    
    def _generate_migration_summary_table(self, sheet_name: str, config: Dict, quarterly_migration: Dict) -> Dict:
        """인구이동 요약 테이블 생성"""
        quarter_labels = self._calculate_quarters()
        region_data = {}
        
        for label in quarter_labels:
            q_data = quarterly_migration.get(label, quarterly_migration.get(label.rstrip('p'), {}))
            for region in ALL_REGIONS:
                if region == '전국':
                    continue
                if region not in region_data:
                    region_data[region] = {'migrations': [], 'amounts': [None, None]}
                region_data[region]['migrations'].append(q_data.get(region))
        
        raw_amounts = self._extract_migration_amounts(sheet_name, config)
        for region in region_data:
            region_data[region]['amounts'] = raw_amounts.get(region, [None, None])
        
        rows = self._create_migration_rows(region_data)
        
        return {
            'columns': {
                'migration_columns': quarter_labels,
                'amount_columns': [
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
    
    def _extract_raw_values(self, sheet_name: str, config: Dict) -> Dict:
        """원시값 추출 (고용률, 실업률)"""
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
        
        for row_idx in range(len(df)):
            try:
                region = str(df.iloc[row_idx, region_col]).strip()
                region = self.normalize_region(region)
                if region not in ALL_REGIONS or region in result:
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
    
    def _extract_migration_amounts(self, sheet_name: str, config: Dict) -> Dict:
        """인구이동 절대값 추출"""
        return self._extract_raw_values(sheet_name, config)
    
    def _create_employment_rows(self, region_data: Dict) -> List[Dict]:
        rows = []
        national = region_data.get('전국', {'changes': [None]*4, 'rates': [None, None]})
        rows.append({
            'region': '전 국',
            'group': None,
            'changes': national['changes'][:4],
            'rates': national['rates'][:2],
        })
        
        for group_name in ['수도권', '동남권', '대경권', '호남권', '충청권', '강원제주']:
            sidos = self.REGION_GROUPS[group_name]
            for idx, sido in enumerate(sidos):
                sido_data = region_data.get(sido, {'changes': [None]*4, 'rates': [None, None]})
                row = {
                    'region': self.REGION_DISPLAY.get(sido, sido),
                    'changes': sido_data['changes'][:4],
                    'rates': sido_data['rates'][:2],
                }
                if idx == 0:
                    row['group'] = group_name
                    row['rowspan'] = len(sidos)
                else:
                    row['group'] = None
                rows.append(row)
        
        return rows
    
    def _create_migration_rows(self, region_data: Dict) -> List[Dict]:
        rows = []
        
        for group_name in ['수도권', '동남권', '대경권', '호남권', '충청권', '강원제주']:
            sidos = self.REGION_GROUPS[group_name]
            for idx, sido in enumerate(sidos):
                sido_data = region_data.get(sido, {'migrations': [None]*4, 'amounts': [None, None]})
                row = {
                    'region': self.REGION_DISPLAY.get(sido, sido),
                    'migrations': sido_data['migrations'][:4],
                    'amounts': sido_data['amounts'][:2],
                }
                if idx == 0:
                    row['group'] = group_name
                    row['rowspan'] = len(sidos)
                else:
                    row['group'] = None
                rows.append(row)
        
        return rows
