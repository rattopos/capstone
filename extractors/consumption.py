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
        # 키워드 기반으로 시트 찾기
        sheet_name = self._find_consumption_sheet()
        if not sheet_name:
            sheet_name = '소비(소매, 추가)'
        
        config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        if not config:
            # 기본 소비 시트 config 사용
            base_config = RAW_SHEET_QUARTER_COLS.get('소비(소매, 추가)', {})
            config = base_config.copy()
        
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
        
        # 전국 업태 추출 (level 자동 감지)
        national_businesses = self._extract_level_businesses(sheet_name, '전국', '소비', level=None)
        report_data['nationwide_data'] = {
            'sales_index': None,
            'growth_rate': national_rate,
            'main_businesses': national_businesses,
            'main_categories': national_businesses,  # 호환성
        }
        
        regional_list = self._process_regional_data(current_data, sheet_name, '소비')
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
        """건설동향 보도자료 데이터 추출
        
        '공정' 키워드로 시트를 찾아 데이터 추출
        """
        # '공정' 키워드로 시트 찾기
        sheet_name = self._find_construction_sheet()
        if not sheet_name:
            print("[건설동향] '공정' 키워드가 있는 시트를 찾을 수 없습니다.")
            # 기본값으로 시도
            sheet_name = '건설 (공표자료)'
        
        # 찾은 시트의 config 가져오기 (없으면 기본값 사용)
        config = RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        if not config:
            # 기본 건설 시트 config 사용
            base_config = RAW_SHEET_QUARTER_COLS.get('건설 (공표자료)', {})
            config = base_config.copy()
        
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
        
        # 전국 공종 추출 (level 자동 감지)
        national_sectors = self._extract_level_businesses(sheet_name, '전국', '건설', level=None)
        report_data['nationwide_data'] = {
            'order_amount': None,
            'growth_rate': national_rate,
            'main_sectors': national_sectors,
            'civil_growth': national_civil_building.get('civil_growth'),
            'building_growth': national_civil_building.get('building_growth'),
            'civil_subtypes': '철도·궤도, 기계설치' if (national_civil_building.get('civil_growth') or 0) > 0 else '토지조성, 치산·치수',
            'building_subtypes': '주택, 관공서 등',
        }
        
        regional_list = self._process_regional_data(current_data, sheet_name, '건설')
        
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
    
    def _find_construction_sheet(self) -> Optional[str]:
        """'공정' 키워드로 건설 시트 찾기
        
        '건설'과 '공정' 키워드를 모두 포함하는 시트를 우선적으로 찾고,
        없으면 '공정' 키워드만 있는 시트를 찾습니다.
        
        Returns:
            찾은 시트명 또는 None
        """
        xl = self._get_excel_file()
        
        # 1단계: '건설'과 '공정'을 모두 포함하는 시트 찾기
        for sheet_name in xl.sheet_names:
            if '건설' in sheet_name:
                try:
                    df_header = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=10)
                    
                    for i in range(min(10, len(df_header))):
                        for j in range(min(20, len(df_header.columns))):
                            val = str(df_header.iloc[i, j]) if pd.notna(df_header.iloc[i, j]) else ''
                            if '공정' in val:
                                print(f"[건설동향] '건설'과 '공정' 키워드로 시트 발견: {sheet_name}")
                                return sheet_name
                except Exception as e:
                    continue
        
        # 2단계: '공정' 키워드만 있는 시트 찾기 (건설 관련 시트 우선)
        construction_candidates = []
        other_candidates = []
        
        for sheet_name in xl.sheet_names:
            try:
                df_header = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=10)
                
                has_process_keyword = False
                for i in range(min(10, len(df_header))):
                    for j in range(min(20, len(df_header.columns))):
                        val = str(df_header.iloc[i, j]) if pd.notna(df_header.iloc[i, j]) else ''
                        if '공정' in val:
                            has_process_keyword = True
                            break
                    if has_process_keyword:
                        break
                
                if has_process_keyword:
                    if '건설' in sheet_name:
                        construction_candidates.append(sheet_name)
                    else:
                        other_candidates.append(sheet_name)
            except Exception as e:
                continue
        
        # 건설 관련 시트 우선 반환
        if construction_candidates:
            print(f"[건설동향] '공정' 키워드로 건설 시트 발견: {construction_candidates[0]}")
            return construction_candidates[0]
        
        # 다른 시트 중에서 선택 (일반적으로는 발생하지 않아야 함)
        if other_candidates:
            print(f"[건설동향] 경고: '공정' 키워드가 있지만 '건설'이 없는 시트 발견: {other_candidates[0]}")
            return other_candidates[0]
        
        return None
    
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
    
    def _process_regional_data(self, current_data: Dict, sheet_name: str = None, category: str = None) -> List[Dict]:
        regional_list = []
        for region in ALL_REGIONS:
            if region == '전국':
                continue
            rate = current_data.get(region)
            if rate is None:
                continue
            
            # 지역별 업태 추출 (level 자동 감지)
            top_businesses = []
            if sheet_name and category:
                top_businesses = self._extract_level_businesses(sheet_name, region, category, level=None)
            
            regional_list.append({
                'region': region,
                'growth_rate': rate,
                'direction': self._get_direction(rate),
                'top_businesses': top_businesses,
                'businesses': top_businesses,  # 호환성
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
    
    def _extract_level_businesses(self, sheet_name: str, region: str, category: str, level: Optional[int] = None) -> List[Dict]:
        """level별 업태 추출 (기여도 기준 상위 3개)
        
        Args:
            sheet_name: 시트 이름
            region: 지역명
            category: 카테고리 ('소비' 또는 '건설')
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
        name_col = config.get('name_col', 4)  # 업태이름 컬럼
        weight_col = config.get('weight_col', 3)  # 가중치 컬럼
        
        # 현재 분기 컬럼
        current_key = f"{self.current_year}_{self.current_quarter}Q"
        prev_key = f"{self.current_year - 1}_{self.current_quarter}Q"
        current_col = config.get(current_key)
        prev_col = config.get(prev_key)
        
        if current_col is None or prev_col is None:
            return []
        
        businesses = []
        
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
                
                # 업태 이름 추출
                business_name = str(df.iloc[row_idx, name_col]).strip() if name_col < len(df.columns) and pd.notna(df.iloc[row_idx, name_col]) else ''
                if not business_name:
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
                shortened_name = self.get_shortened_name(business_name, category)
                
                businesses.append({
                    'name': shortened_name,
                    'growth_rate': growth_rate,
                    'contribution': contribution,
                })
            except (IndexError, ValueError):
                continue
        
        # 기여도 절대값 기준으로 정렬 (상위 3개)
        businesses.sort(key=lambda x: abs(x['contribution']) if x['contribution'] is not None else 0, reverse=True)
        
        return businesses[:3]
    
    def _find_consumption_sheet(self) -> Optional[str]:
        """소비 시트 찾기 ('업태' 키워드 사용)
        
        Returns:
            찾은 시트명 또는 None
        """
        found_sheet = self.find_sheet_by_content_keyword(
            content_keywords=['업태'],
            sheet_name_keywords=['소비', '소매']
        )
        
        return found_sheet