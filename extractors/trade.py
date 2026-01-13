#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
무역 관련 데이터 추출기
수출, 수입 보도자료 데이터 추출
"""

from typing import Dict, List, Any, Optional
import pandas as pd

from .base import BaseExtractor
from .config import ALL_REGIONS, RAW_SHEET_QUARTER_COLS


class TradeExtractor(BaseExtractor):
    """무역 관련 데이터 추출기 (수출, 수입)"""
    
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
    
    def extract_export_data(self) -> Dict[str, Any]:
        """수출 보도자료 데이터 추출"""
        return self._extract_trade_data('수출', '수출')
    
    def extract_import_data(self) -> Dict[str, Any]:
        """수입 보도자료 데이터 추출"""
        return self._extract_trade_data('수입', '수입')
    
    def _extract_trade_data(self, sheet_name: str, title: str) -> Dict[str, Any]:
        """무역 데이터 공통 추출 로직"""
        # 키워드 기반으로 시트 찾기
        actual_sheet_name = self._find_trade_sheet(sheet_name)
        if actual_sheet_name:
            sheet_name = actual_sheet_name
        
        config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        if not config:
            # 기본 무역 시트 config 사용
            base_config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
            config = base_config.copy()
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': title,
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
            'change': national_rate,
            'direction': self._get_direction(national_rate),
        }
        
        # 카테고리 결정 (수출/수입)
        category = '수출' if '수출' in sheet_name else '수입'
        
        # 전국 level=1 품목 추출
        national_items = self._extract_level_items(sheet_name, '전국', category, level=None)
        report_data['nationwide_data'] = {
            'amount': None,
            'change': national_rate,
            'main_items': national_items,
        }
        
        regional_list = self._process_regional_data(current_data, sheet_name, category)
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
    
    def _extract_quarterly_growth(self, sheet_name: str, config: Dict) -> Dict:
        """분기별 증감률 추출 (수출/수입용 - name_col에서 '계' 확인)"""
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        region_col = config.get('region_col', 1)
        name_col = config.get('name_col', 5)
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
                    
                    # 상품분류 확인 ('계'인 행만)
                    name_val = df.iloc[row_idx, name_col]
                    if pd.isna(name_val) or str(name_val).strip() != total_code:
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
    
    def _process_regional_data(self, current_data: Dict, sheet_name: str = None, category: str = None) -> List[Dict]:
        regional_list = []
        for region in ALL_REGIONS:
            if region == '전국':
                continue
            rate = current_data.get(region)
            if rate is None:
                continue
            
            # 지역별 level=2 품목 추출
            top_items = []
            if sheet_name and category:
                top_items = self._extract_level_items(sheet_name, region, category, level=None)
            
            regional_list.append({
                'region': region,
                'change': rate,
                'direction': self._get_direction(rate),
                'top_items': top_items,
                'items': top_items,  # 호환성
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
    
    def _generate_summary_table(self, sheet_name: str, config: Dict, quarterly_growth: Dict) -> Dict:
        quarter_labels = self._calculate_quarters()
        region_data = {}
        
        for label in quarter_labels:
            q_data = quarterly_growth.get(label, quarterly_growth.get(label.rstrip('p'), {}))
            for region in ALL_REGIONS:
                if region not in region_data:
                    region_data[region] = {'changes': [], 'amounts': [None, None]}
                region_data[region]['changes'].append(q_data.get(region))
        
        raw_amounts = self._extract_raw_amounts(sheet_name, config)
        for region in ALL_REGIONS:
            if region in region_data:
                region_data[region]['amounts'] = raw_amounts.get(region, [None, None])
        
        rows = self._create_rows(region_data)
        
        return {
            'columns': {
                'change_columns': quarter_labels,
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
    
    def _extract_raw_amounts(self, sheet_name: str, config: Dict) -> Dict:
        """원금액 추출 (백만달러 → 억달러 변환)"""
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
        name_col = config.get('name_col', 5)
        total_code = config.get('total_code', '계')
        
        for row_idx in range(len(df)):
            try:
                region = str(df.iloc[row_idx, region_col]).strip()
                region = self.normalize_region(region)
                if region not in ALL_REGIONS or region in result:
                    continue
                
                name_val = df.iloc[row_idx, name_col]
                if pd.isna(name_val) or str(name_val).strip() != total_code:
                    continue
                
                prev_val = self.safe_float(df.iloc[row_idx, prev_col])
                curr_val = self.safe_float(df.iloc[row_idx, curr_col])
                
                # 백만달러 → 억달러 변환 (/100)
                result[region] = [
                    round(prev_val / 100, 0) if prev_val else None,
                    round(curr_val / 100, 0) if curr_val else None
                ]
            except (IndexError, ValueError, ZeroDivisionError):
                continue
        
        return result
    
    def _create_rows(self, region_data: Dict) -> List[Dict]:
        rows = []
        national = region_data.get('전국', {'changes': [None]*4, 'amounts': [None, None]})
        rows.append({
            'region': '전 국',
            'group': None,
            'changes': national['changes'][:4],
            'amounts': national['amounts'][:2],
        })
        
        for group_name in ['수도권', '동남권', '대경권', '호남권', '충청권', '강원제주']:
            sidos = self.REGION_GROUPS[group_name]
            for idx, sido in enumerate(sidos):
                sido_data = region_data.get(sido, {'changes': [None]*4, 'amounts': [None, None]})
                row = {
                    'region': self.REGION_DISPLAY.get(sido, sido),
                    'changes': sido_data['changes'][:4],
                    'amounts': sido_data['amounts'][:2],
                }
                if idx == 0:
                    row['group'] = group_name
                    row['rowspan'] = len(sidos)
                else:
                    row['group'] = None
                rows.append(row)
        
        return rows
    
    def _extract_level_items(self, sheet_name: str, region: str, category: str, level: Optional[int] = None) -> List[Dict]:
        """level별 품목 추출 (기여도 기준 상위 3개)
        
        Args:
            sheet_name: 시트 이름
            region: 지역명
            category: 카테고리 ('수출' 또는 '수입')
            level: 분류단계 (None이면 데이터가 있는 모든 level 시도, 0 제외)
            
        Returns:
            [{'name': '축약이름', 'growth_rate': 증감률, 'contribution': 기여도}, ...]
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return []
        
        config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        name_col = config.get('name_col', 5)  # 품목이름 컬럼
        weight_col = config.get('weight_col', 3)  # 가중치 컬럼
        
        # 현재 분기 컬럼
        current_key = f"{self.current_year}_{self.current_quarter}Q"
        prev_key = f"{self.current_year - 1}_{self.current_quarter}Q"
        current_col = config.get(current_key)
        prev_col = config.get(prev_key)
        
        if current_col is None or prev_col is None:
            return []
        
        items = []
        
        for row_idx in range(len(df)):
            try:
                row_region = str(df.iloc[row_idx, region_col]).strip()
                row_region = self.normalize_region(row_region)
                
                if row_region != region:
                    continue
                
                row_level = df.iloc[row_idx, level_col]
                if pd.isna(row_level):
                    continue
                
                # level 비교: 숫자로 변환하여 비교 (2.0 == 2, "2" == 2 등 처리)
                row_level_num = None
                try:
                    row_level_num = float(str(row_level).strip())
                except (ValueError, TypeError):
                    pass
                
                # level 필터링: 지정된 level이 있으면 그것만, 없으면 0이 아닌 모든 level
                level_match = False
                if level is not None:
                    if row_level_num is not None:
                        level_num = float(level)
                        level_match = abs(row_level_num - level_num) <= 0.01
                    else:
                        level_match = str(row_level).strip() == str(level)
                else:
                    # level이 None이면 0이 아닌 모든 level 허용
                    if row_level_num is not None:
                        level_match = abs(row_level_num) > 0.01
                    else:
                        level_match = str(row_level).strip() not in ['0', '0.0', '총지수', '계']
                
                if not level_match:
                    continue
                
                # 품목 이름 추출
                item_name = str(df.iloc[row_idx, name_col]).strip() if name_col < len(df.columns) and pd.notna(df.iloc[row_idx, name_col]) else ''
                if not item_name:
                    continue
                
                # 증감률 계산
                current_val = self.safe_float(df.iloc[row_idx, current_col])
                prev_val = self.safe_float(df.iloc[row_idx, prev_col])
                growth_rate = self.calculate_growth_rate(current_val, prev_val)
                
                # 기여도 계산
                weight = self.safe_float(df.iloc[row_idx, weight_col]) if weight_col < len(df.columns) else None
                contribution = None
                if growth_rate is not None:
                    if weight is not None:
                        contribution = growth_rate * weight / 100.0
                    else:
                        contribution = growth_rate
                
                # 축약 이름 매칭
                shortened_name = self.get_shortened_name(item_name, category)
                
                items.append({
                    'name': shortened_name,
                    'growth_rate': growth_rate,
                    'contribution': contribution,
                })
            except (IndexError, ValueError):
                continue
        
        # 기여도 절대값 기준으로 정렬 (상위 3개)
        items.sort(key=lambda x: abs(x['contribution']) if x['contribution'] is not None else 0, reverse=True)
        
        return items[:3]
    
    def _find_trade_sheet(self, default_sheet_name: str) -> Optional[str]:
        """무역 시트 찾기 ('상품 이름' 키워드 사용)
        
        Args:
            default_sheet_name: 기본 시트명 ('수출' 또는 '수입')
            
        Returns:
            찾은 시트명 또는 None
        """
        found_sheet = self.find_sheet_by_content_keyword(
            content_keywords=['상품 이름'],
            sheet_name_keywords=[default_sheet_name]
        )
        
        return found_sheet