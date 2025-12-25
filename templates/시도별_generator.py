#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
시도별 경제동향 보고서 생성기

17개 시도별 경제동향 데이터를 엑셀에서 추출하고,
Jinja2 템플릿을 사용하여 HTML 보고서를 생성합니다.
효율성을 위해 데이터를 캐싱합니다.
"""

import pandas as pd
import json
from jinja2 import Environment, FileSystemLoader
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, field
from functools import lru_cache


@dataclass
class RegionInfo:
    """지역 정보"""
    code: int
    name: str
    full_name: str
    index: int  # 1-17


# 17개 시도 정보
REGIONS = [
    RegionInfo(11, "서울", "서울특별시", 1),
    RegionInfo(21, "부산", "부산광역시", 2),
    RegionInfo(22, "대구", "대구광역시", 3),
    RegionInfo(23, "인천", "인천광역시", 4),
    RegionInfo(24, "광주", "광주광역시", 5),
    RegionInfo(25, "대전", "대전광역시", 6),
    RegionInfo(26, "울산", "울산광역시", 7),
    RegionInfo(29, "세종", "세종특별자치시", 8),
    RegionInfo(31, "경기", "경기도", 9),
    RegionInfo(32, "강원", "강원특별자치도", 10),
    RegionInfo(33, "충북", "충청북도", 11),
    RegionInfo(34, "충남", "충청남도", 12),
    RegionInfo(35, "전북", "전북특별자치도", 13),
    RegionInfo(36, "전남", "전라남도", 14),
    RegionInfo(37, "경북", "경상북도", 15),
    RegionInfo(38, "경남", "경상남도", 16),
    RegionInfo(39, "제주", "제주특별자치도", 17),
]


class DataCache:
    """엑셀 데이터 캐시 클래스"""
    
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self._sheets: Dict[str, pd.DataFrame] = {}
        self._nationwide_cache: Dict[str, Any] = {}
        
    def get_sheet(self, sheet_name: str) -> pd.DataFrame:
        """시트 데이터를 캐시에서 가져오거나 로드"""
        if sheet_name not in self._sheets:
            self._sheets[sheet_name] = pd.read_excel(
                self.excel_path, 
                sheet_name=sheet_name, 
                header=None
            )
        return self._sheets[sheet_name]
    
    def preload_all_sheets(self):
        """모든 필요한 시트를 미리 로드"""
        required_sheets = [
            'A 분석', 'A(광공업생산)집계',
            'B 분석', 'B(서비스업생산)집계',
            'C 분석', 'C(소비)집계',
            "F'분석", "F'(건설)집계",
            'G 분석', 'G(수출)집계',
            'H 분석', 'H(수입)집계',
            'E(지출목적물가) 분석', 'E(지출목적물가)집계',
            'D(고용률)분석', 'D(고용률)집계',
            'I(순인구이동)집계',
        ]
        for sheet in required_sheets:
            try:
                self.get_sheet(sheet)
            except Exception as e:
                print(f"Warning: Could not load sheet '{sheet}': {e}")


class 시도별Generator:
    """시도별 경제동향 보고서 생성 클래스"""
    
    # 업종명 매핑 사전 (엑셀 데이터 → 보고서 표기명)
    INDUSTRY_NAME_MAP = {
        # 광공업
        "전자 부품, 컴퓨터, 영상, 음향 및 통신장비 제조업": "반도체·전자부품",
        "의료, 정밀, 광학 기기 및 시계 제조업": "의료·정밀",
        "의료용 물질 및 의약품 제조업": "의약품",
        "기타 운송장비 제조업": "기타 운송장비",
        "기타 기계 및 장비 제조업": "기타기계장비",
        "전기장비 제조업": "전기장비",
        "자동차 및 트레일러 제조업": "자동차·트레일러",
        "전기, 가스, 증기 및 공기 조절 공급업": "전기·가스업",
        "전기업 및 가스업": "전기·가스업",
        "식료품 제조업": "식료품",
        "금속 가공제품 제조업; 기계 및 가구 제외": "금속가공제품",
        "1차 금속 제조업": "1차금속",
        "화학 물질 및 화학제품 제조업; 의약품 제외": "화학물질",
        "담배 제조업": "담배",
        "고무 및 플라스틱제품 제조업": "고무·플라스틱",
        "비금속 광물제품 제조업": "비금속광물",
        "섬유제품 제조업; 의복 제외": "섬유제품",
        "금속 광업": "금속광업",
        "산업용 기계 및 장비 수리업": "산업용기계",
        "펄프, 종이 및 종이제품 제조업": "펄프·종이",
        "인쇄 및 기록매체 복제업": "인쇄",
        "음료 제조업": "음료",
        "가구 제조업": "가구",
        "기타 제품 제조업": "기타제품",
        "가죽, 가방 및 신발 제조업": "가죽·신발",
        "의복, 의복액세서리 및 모피제품 제조업": "의복·모피",
        "코크스, 연탄 및 석유정제품 제조업": "석유정제품",
        "목재 및 나무제품 제조업; 가구 제외": "목재제품",
        "비금속광물 광업; 연료용 제외": "비금속광물광업",
        # 서비스업
        "운수 및 창고업": "운수·창고업",
        "도매 및 소매업": "도소매업",
        "숙박 및 음식점업": "숙박·음식점업",
        "정보통신업": "정보통신업",
        "금융 및 보험업": "금융·보험",
        "부동산업": "부동산",
        "전문, 과학 및 기술 서비스업": "전문·과학·기술",
        "사업시설 관리, 사업 지원 및 임대 서비스업": "사업시설관리·사업지원·임대",
        "교육 서비스업": "교육서비스",
        "보건업 및 사회복지 서비스업": "보건·사회복지",
        "예술, 스포츠 및 여가관련 서비스업": "예술·스포츠·여가",
        "협회 및 단체, 수리 및 기타 개인 서비스업": "협회·수리·개인서비스",
    }
    
    # 소매판매 업태 매핑
    RETAIL_NAME_MAP = {
        "승용차 및 연료 소매점": "승용차·연료소매점",
        "전문 소매점": "전문소매점",
        "슈퍼마켓· 잡화점 및 편의점": "슈퍼마켓·잡화점·편의점",
        "대형 마트": "대형마트",
        "백화점": "백화점",
        "면세점": "면세점",
        "무점포 소매": "무점포소매",
    }
    
    # 분기별 컬럼 매핑 (각 시트별로 다름)
    QUARTER_COLS = {
        'A 분석': {'2023_2Q': 13, '2024_2Q': 17, '2025_1Q': 20, '2025_2Q': 21},
        'B 분석': {'2023_2Q': 12, '2024_2Q': 16, '2025_1Q': 19, '2025_2Q': 20},
        'C 분석': {'2023_2Q': 12, '2024_2Q': 16, '2025_1Q': 19, '2025_2Q': 20},
        "F'분석": {'2023_2Q': 11, '2024_2Q': 15, '2025_1Q': 18, '2025_2Q': 19},
        'G 분석': {'2023_2Q': 14, '2024_2Q': 18, '2025_1Q': 21, '2025_2Q': 22},
        'H 분석': {'2023_2Q': 14, '2024_2Q': 18, '2025_1Q': 21, '2025_2Q': 22},
        'E(지출목적물가) 분석': {'2023_2Q': 13, '2024_2Q': 17, '2025_1Q': 20, '2025_2Q': 21},
        'D(고용률)분석': {'2023_2Q': 10, '2024_2Q': 14, '2025_1Q': 17, '2025_2Q': 18},
        'I(순인구이동)집계': {'2023_2Q': 17, '2024_2Q': 21, '2025_1Q': 24, '2025_2Q': 25},
    }
    
    def __init__(self, excel_path: str):
        """초기화"""
        self.excel_path = excel_path
        self.cache = DataCache(excel_path)
        self.cache.preload_all_sheets()
        self._nationwide_data: Dict[str, Any] = {}
        
    def _clean_name(self, name: str) -> str:
        """이름 정제 (공백 및 특수문자 처리)"""
        if pd.isna(name):
            return ""
        return str(name).strip().replace("\u3000", "").replace("　", "").strip()
    
    def _get_display_name(self, raw_name: str, name_map: Dict[str, str]) -> str:
        """보고서 표기명으로 변환"""
        cleaned = self._clean_name(raw_name)
        for key, value in name_map.items():
            if key in cleaned or cleaned in key:
                return value
        return cleaned
    
    def _get_quarter_value(self, row: pd.Series, sheet_name: str, quarter: str) -> float:
        """분기별 값 추출"""
        col = self.QUARTER_COLS.get(sheet_name, {}).get(quarter)
        if col is None:
            return 0.0
        val = row.iloc[col] if col < len(row) else None
        # '없음' 등의 문자열 처리
        if pd.isna(val) or val == '없음' or val == '-':
            return 0.0
        try:
            return round(float(val), 1)
        except (ValueError, TypeError):
            return 0.0
    
    def _get_chart_time_series(self, df: pd.DataFrame, region: str, 
                               region_col: int, code_col: int, total_code: str,
                               sheet_name: str) -> List[Dict[str, Any]]:
        """차트용 시계열 데이터 추출"""
        row = df[(df[region_col] == region) & (df[code_col].astype(str) == total_code)]
        if len(row) == 0:
            return []
        row = row.iloc[0]
        
        quarters = ["'23.2/4", "'24.2/4", "'25.2/4"]
        keys = ['2023_2Q', '2024_2Q', '2025_2Q']
        
        return [
            {"period": q, "value": self._get_quarter_value(row, sheet_name, k)}
            for q, k in zip(quarters, keys)
        ]
    
    def _extract_top_industries(self, df: pd.DataFrame, region: str,
                                region_col: int, name_col: int, 
                                contribution_col: int, quarter_col: int,
                                name_map: Dict[str, str],
                                direction: str = 'increase',
                                top_n: int = 3) -> List[Dict[str, Any]]:
        """상위/하위 기여 산업 추출"""
        region_data = df[df[region_col] == region]
        
        # 기여도가 있는 행만 필터링
        if contribution_col < len(df.columns):
            region_data = region_data[pd.notna(region_data[contribution_col])]
        
        if len(region_data) == 0:
            return []
        
        # 정렬
        ascending = direction == 'decrease'
        sorted_data = region_data.sort_values(contribution_col, ascending=ascending)
        
        # 기여도 기준 필터링
        if direction == 'increase':
            sorted_data = sorted_data[sorted_data[contribution_col] > 0]
        else:
            sorted_data = sorted_data[sorted_data[contribution_col] < 0]
        
        result = []
        for _, row in sorted_data.head(top_n).iterrows():
            name = self._get_display_name(row[name_col], name_map)
            growth_rate = float(row[quarter_col]) if pd.notna(row[quarter_col]) else 0.0
            contribution = float(row[contribution_col]) if pd.notna(row[contribution_col]) else 0.0
            
            result.append({
                "name": name,
                "growth_rate": round(growth_rate, 1),
                "contribution": round(contribution, 4)
            })
        
        return result
    
    def extract_manufacturing_data(self, region: str) -> Dict[str, Any]:
        """광공업생산 데이터 추출"""
        df = self.cache.get_sheet('A 분석')
        
        # 총지수 행
        total_row = df[(df[3] == region) & (df[6] == 'BCD')]
        if len(total_row) == 0:
            return {"total_growth_rate": 0, "direction": "감소", 
                    "increase_industries": [], "decrease_industries": []}
        
        total_row = total_row.iloc[0]
        growth_rate = self._get_quarter_value(total_row, 'A 분석', '2025_2Q')
        
        # 중분류 데이터 (분류단계 2)
        industries = df[(df[3] == region) & (df[4].astype(str) == '2') & (pd.notna(df[28]))]
        
        # 기여도 순 정렬하여 상위/하위 추출
        increase_industries = []
        decrease_industries = []
        
        if len(industries) > 0:
            sorted_pos = industries[industries[28] > 0].sort_values(28, ascending=False)
            sorted_neg = industries[industries[28] < 0].sort_values(28, ascending=True)
            
            for _, row in sorted_pos.head(2).iterrows():
                increase_industries.append({
                    "name": self._get_display_name(row[7], self.INDUSTRY_NAME_MAP),
                    "growth_rate": round(float(row[21]) if pd.notna(row[21]) else 0, 1),
                    "contribution": round(float(row[28]), 4)
                })
            
            for _, row in sorted_neg.head(2).iterrows():
                decrease_industries.append({
                    "name": self._get_display_name(row[7], self.INDUSTRY_NAME_MAP),
                    "growth_rate": round(float(row[21]) if pd.notna(row[21]) else 0, 1),
                    "contribution": round(float(row[28]), 4)
                })
        
        return {
            "total_growth_rate": growth_rate,
            "direction": "증가" if growth_rate > 0 else "감소",
            "increase_industries": increase_industries,
            "decrease_industries": decrease_industries
        }
    
    def extract_service_data(self, region: str) -> Dict[str, Any]:
        """서비스업생산 데이터 추출"""
        df = self.cache.get_sheet('B 분석')
        
        # 총지수 행
        total_row = df[(df[3] == region) & (df[6] == 'E~S')]
        if len(total_row) == 0:
            return {"total_growth_rate": 0, "direction": "증가",
                    "increase_industries": [], "decrease_industries": []}
        
        total_row = total_row.iloc[0]
        growth_rate = self._get_quarter_value(total_row, 'B 분석', '2025_2Q')
        
        # 중분류 데이터
        industries = df[(df[3] == region) & (df[4].astype(str) == '1')]
        
        increase_industries = []
        decrease_industries = []
        
        if len(industries) > 0:
            # 증감률 기준으로 정렬
            quarter_col = self.QUARTER_COLS['B 분석']['2025_2Q']
            sorted_data = industries.sort_values(quarter_col, ascending=False)
            
            # 양수 증감률 (증가)
            pos_data = sorted_data[sorted_data[quarter_col] > 0]
            for _, row in pos_data.head(2).iterrows():
                increase_industries.append({
                    "name": self._get_display_name(row[7], self.INDUSTRY_NAME_MAP),
                    "growth_rate": round(float(row[quarter_col]) if pd.notna(row[quarter_col]) else 0, 1)
                })
            
            # 음수 증감률 (감소)
            neg_data = sorted_data[sorted_data[quarter_col] < 0].sort_values(quarter_col)
            for _, row in neg_data.head(2).iterrows():
                decrease_industries.append({
                    "name": self._get_display_name(row[7], self.INDUSTRY_NAME_MAP),
                    "growth_rate": round(float(row[quarter_col]) if pd.notna(row[quarter_col]) else 0, 1)
                })
        
        return {
            "total_growth_rate": growth_rate,
            "direction": "증가" if growth_rate > 0 else "감소",
            "increase_industries": increase_industries,
            "decrease_industries": decrease_industries
        }
    
    def extract_retail_data(self, region: str) -> Dict[str, Any]:
        """소매판매 데이터 추출"""
        df = self.cache.get_sheet('C 분석')
        
        # 총지수 행
        total_row = df[(df[3] == region) & (df[4].astype(str) == '0')]
        if len(total_row) == 0:
            return {"total_growth_rate": 0, "direction": "감소",
                    "increase_categories": [], "decrease_categories": []}
        
        total_row = total_row.iloc[0]
        growth_rate = self._get_quarter_value(total_row, 'C 분석', '2025_2Q')
        
        # 업태별 데이터
        categories = df[(df[3] == region) & (df[4].astype(str) == '1')]
        
        increase_categories = []
        decrease_categories = []
        
        if len(categories) > 0:
            quarter_col = self.QUARTER_COLS['C 분석']['2025_2Q']
            sorted_data = categories.sort_values(quarter_col, ascending=False)
            
            pos_data = sorted_data[sorted_data[quarter_col] > 0]
            for _, row in pos_data.head(2).iterrows():
                increase_categories.append({
                    "name": self._get_display_name(row[7], self.RETAIL_NAME_MAP),
                    "growth_rate": round(float(row[quarter_col]) if pd.notna(row[quarter_col]) else 0, 1)
                })
            
            neg_data = sorted_data[sorted_data[quarter_col] < 0].sort_values(quarter_col)
            for _, row in neg_data.head(2).iterrows():
                decrease_categories.append({
                    "name": self._get_display_name(row[7], self.RETAIL_NAME_MAP),
                    "growth_rate": round(float(row[quarter_col]) if pd.notna(row[quarter_col]) else 0, 1)
                })
        
        return {
            "total_growth_rate": growth_rate,
            "direction": "증가" if growth_rate > 0 else "감소",
            "increase_categories": increase_categories,
            "decrease_categories": decrease_categories
        }
    
    def extract_construction_data(self, region: str) -> Dict[str, Any]:
        """건설수주 데이터 추출"""
        df = self.cache.get_sheet("F'분석")
        
        # 총지수 행 (합계)
        total_row = df[(df[2] == region) & (df[3].astype(str) == '0')]
        if len(total_row) == 0:
            return {"total_growth_rate": 0, "direction": "증가",
                    "increase_categories": [{"name": "[품목명]", "growth_rate": 0, "placeholder": True}], 
                    "decrease_categories": [{"name": "[품목명]", "growth_rate": 0, "placeholder": True}]}
        
        total_row = total_row.iloc[0]
        quarter_col = self.QUARTER_COLS["F'분석"]['2025_2Q']
        growth_rate = round(float(total_row[quarter_col]) if pd.notna(total_row[quarter_col]) else 0, 1)
        
        # 건축/토목 분류
        categories = df[(df[2] == region) & (df[3].astype(str).isin(['1', '2']))]
        
        increase_categories = []
        decrease_categories = []
        
        if len(categories) > 0:
            for _, row in categories.iterrows():
                name = self._clean_name(row[6])
                val = row[quarter_col]
                # '없음' 등의 문자열 처리 - 건축/토목 외 품목은 비공개
                if pd.isna(val) or val == '없음' or val == '-':
                    continue
                try:
                    rate = round(float(val), 1)
                except (ValueError, TypeError):
                    continue
                category_info = {"name": name, "growth_rate": rate, "placeholder": False}
                
                if rate > 0:
                    increase_categories.append(category_info)
                elif rate < 0:
                    decrease_categories.append(category_info)
        
        # 플레이스홀더 추가 (데이터가 부족한 경우)
        if len(increase_categories) == 0:
            increase_categories.append({"name": "[품목명]", "growth_rate": 0, "placeholder": True})
        if len(decrease_categories) == 0:
            decrease_categories.append({"name": "[품목명]", "growth_rate": 0, "placeholder": True})
        
        return {
            "total_growth_rate": growth_rate,
            "direction": "증가" if growth_rate > 0 else "감소",
            "increase_categories": sorted([c for c in increase_categories if not c.get('placeholder')], key=lambda x: x['growth_rate'], reverse=True)[:2] or [{"name": "[품목명]", "growth_rate": 0, "placeholder": True}],
            "decrease_categories": sorted([c for c in decrease_categories if not c.get('placeholder')], key=lambda x: x['growth_rate'])[:2] or [{"name": "[품목명]", "growth_rate": 0, "placeholder": True}]
        }
    
    def extract_export_data(self, region: str) -> Dict[str, Any]:
        """수출 데이터 추출"""
        df = self.cache.get_sheet('G 분석')
        
        # 총계 행
        total_row = df[(df[3] == region) & (df[4].astype(str) == '0')]
        if len(total_row) == 0:
            return {"total_growth_rate": 0, "direction": "증가",
                    "increase_items": [{"name": "[품목명]", "growth_rate": 0, "placeholder": True}], 
                    "decrease_items": [{"name": "[품목명]", "growth_rate": 0, "placeholder": True}]}
        
        total_row = total_row.iloc[0]
        growth_rate = self._get_quarter_value(total_row, 'G 분석', '2025_2Q')
        
        # 품목별 데이터 (기여도 기준) - 분류단계 2 (소분류)
        items = df[(df[3] == region) & (df[4].astype(str) == '2') & (pd.notna(df[26]))]
        
        increase_items = []
        decrease_items = []
        
        if len(items) > 0:
            quarter_col = self.QUARTER_COLS['G 분석']['2025_2Q']
            
            # 기여도 양수 (증가 기여)
            pos_items = items[items[26] > 0].sort_values(26, ascending=False)
            for _, row in pos_items.head(2).iterrows():
                val = row[quarter_col]
                if pd.isna(val) or val == '없음' or val == '-':
                    rate = 0.0
                    is_placeholder = True
                else:
                    try:
                        rate = round(float(val), 1)
                        is_placeholder = False
                    except (ValueError, TypeError):
                        rate = 0.0
                        is_placeholder = True
                increase_items.append({
                    "name": self._clean_name(row[8]),
                    "growth_rate": rate,
                    "placeholder": is_placeholder
                })
            
            # 기여도 음수 (감소 기여)
            neg_items = items[items[26] < 0].sort_values(26)
            for _, row in neg_items.head(2).iterrows():
                val = row[quarter_col]
                if pd.isna(val) or val == '없음' or val == '-':
                    rate = 0.0
                    is_placeholder = True
                else:
                    try:
                        rate = round(float(val), 1)
                        is_placeholder = False
                    except (ValueError, TypeError):
                        rate = 0.0
                        is_placeholder = True
                decrease_items.append({
                    "name": self._clean_name(row[8]),
                    "growth_rate": rate,
                    "placeholder": is_placeholder
                })
        
        # 플레이스홀더 추가 (데이터가 부족한 경우)
        if len(increase_items) == 0:
            increase_items.append({"name": "[품목명]", "growth_rate": 0, "placeholder": True})
        if len(decrease_items) == 0:
            decrease_items.append({"name": "[품목명]", "growth_rate": 0, "placeholder": True})
        
        return {
            "total_growth_rate": growth_rate,
            "direction": "증가" if growth_rate > 0 else "감소",
            "increase_items": increase_items,
            "decrease_items": decrease_items
        }
    
    def extract_import_data(self, region: str) -> Dict[str, Any]:
        """수입 데이터 추출"""
        df = self.cache.get_sheet('H 분석')
        
        # 총계 행
        total_row = df[(df[3] == region) & (df[4].astype(str) == '0')]
        if len(total_row) == 0:
            return {"total_growth_rate": 0, "direction": "감소",
                    "increase_items": [{"name": "[품목명]", "growth_rate": 0, "placeholder": True}],
                    "decrease_items": [{"name": "[품목명]", "growth_rate": 0, "placeholder": True}]}
        
        total_row = total_row.iloc[0]
        growth_rate = self._get_quarter_value(total_row, 'H 분석', '2025_2Q')
        
        # 품목별 데이터 - 분류단계 2 (소분류)
        items = df[(df[3] == region) & (df[4].astype(str) == '2') & (pd.notna(df[26]))]
        
        increase_items = []
        decrease_items = []
        
        if len(items) > 0:
            quarter_col = self.QUARTER_COLS['H 분석']['2025_2Q']
            
            pos_items = items[items[26] > 0].sort_values(26, ascending=False)
            for _, row in pos_items.head(2).iterrows():
                val = row[quarter_col]
                if pd.isna(val) or val == '없음' or val == '-':
                    rate = 0.0
                    is_placeholder = True
                else:
                    try:
                        rate = round(float(val), 1)
                        is_placeholder = False
                    except (ValueError, TypeError):
                        rate = 0.0
                        is_placeholder = True
                increase_items.append({
                    "name": self._clean_name(row[8]),
                    "growth_rate": rate,
                    "placeholder": is_placeholder
                })
            
            neg_items = items[items[26] < 0].sort_values(26)
            for _, row in neg_items.head(2).iterrows():
                val = row[quarter_col]
                if pd.isna(val) or val == '없음' or val == '-':
                    rate = 0.0
                    is_placeholder = True
                else:
                    try:
                        rate = round(float(val), 1)
                        is_placeholder = False
                    except (ValueError, TypeError):
                        rate = 0.0
                        is_placeholder = True
                decrease_items.append({
                    "name": self._clean_name(row[8]),
                    "growth_rate": rate,
                    "placeholder": is_placeholder
                })
        
        # 플레이스홀더 추가 (데이터가 부족한 경우)
        if len(increase_items) == 0:
            increase_items.append({"name": "[품목명]", "growth_rate": 0, "placeholder": True})
        if len(decrease_items) == 0:
            decrease_items.append({"name": "[품목명]", "growth_rate": 0, "placeholder": True})
        
        return {
            "total_growth_rate": growth_rate,
            "direction": "증가" if growth_rate > 0 else "감소",
            "increase_items": increase_items,
            "decrease_items": decrease_items
        }
    
    def extract_consumer_price_data(self, region: str) -> Dict[str, Any]:
        """소비자물가 데이터 추출"""
        df = self.cache.get_sheet('E(지출목적물가) 분석')
        
        # 총지수 행
        total_row = df[(df[3] == region) & (df[4].astype(str) == '0') & (df[8] == '총지수')]
        if len(total_row) == 0:
            return {"total_growth_rate": 0, "direction": "상승",
                    "increase_items": [], "decrease_items": []}
        
        total_row = total_row.iloc[0]
        growth_rate = self._get_quarter_value(total_row, 'E(지출목적물가) 분석', '2025_2Q')
        
        # 품목별 데이터
        items = df[(df[3] == region) & (df[4].astype(str) == '1')]
        
        increase_items = []
        decrease_items = []
        
        if len(items) > 0:
            quarter_col = self.QUARTER_COLS['E(지출목적물가) 분석']['2025_2Q']
            sorted_data = items.sort_values(quarter_col, ascending=False)
            
            pos_data = sorted_data[sorted_data[quarter_col] > 0]
            for _, row in pos_data.head(2).iterrows():
                increase_items.append({
                    "name": self._clean_name(row[8]),
                    "growth_rate": round(float(row[quarter_col]) if pd.notna(row[quarter_col]) else 0, 1)
                })
            
            neg_data = sorted_data[sorted_data[quarter_col] < 0].sort_values(quarter_col)
            for _, row in neg_data.head(2).iterrows():
                decrease_items.append({
                    "name": self._clean_name(row[8]),
                    "growth_rate": round(float(row[quarter_col]) if pd.notna(row[quarter_col]) else 0, 1)
                })
        
        return {
            "total_growth_rate": growth_rate,
            "direction": "상승" if growth_rate > 0 else "하락",
            "increase_items": increase_items,
            "decrease_items": decrease_items
        }
    
    def extract_employment_data(self, region: str) -> Dict[str, Any]:
        """고용률 데이터 추출"""
        df = self.cache.get_sheet('D(고용률)분석')
        
        # 총계 행
        total_row = df[(df[2] == region) & (df[3].astype(str) == '0')]
        if len(total_row) == 0:
            return {"total_change": 0, "direction": "하락",
                    "increase_age_groups": [], "decrease_age_groups": []}
        
        total_row = total_row.iloc[0]
        quarter_col = self.QUARTER_COLS['D(고용률)분석']['2025_2Q']
        total_change = round(float(total_row[quarter_col]) if pd.notna(total_row[quarter_col]) else 0, 1)
        
        # 연령대별 데이터
        age_groups = df[(df[2] == region) & (df[3].astype(str) == '1')]
        
        increase_groups = []
        decrease_groups = []
        
        age_name_map = {
            "15 - 29세": "20대",
            "30 - 39세": "30대",
            "40 - 49세": "40대",
            "50 - 59세": "50대",
            "60세이상": "60대"
        }
        
        if len(age_groups) > 0:
            for _, row in age_groups.iterrows():
                age_name = self._clean_name(row[5])
                change = round(float(row[quarter_col]) if pd.notna(row[quarter_col]) else 0, 1)
                display_name = age_name_map.get(age_name, age_name)
                
                if change > 0:
                    increase_groups.append({"name": display_name, "change": change})
                elif change < 0:
                    decrease_groups.append({"name": display_name, "change": change})
        
        # 정렬
        increase_groups = sorted(increase_groups, key=lambda x: x['change'], reverse=True)[:2]
        decrease_groups = sorted(decrease_groups, key=lambda x: x['change'])[:2]
        
        return {
            "total_change": total_change,
            "direction": "상승" if total_change > 0 else "하락",
            "increase_age_groups": increase_groups,
            "decrease_age_groups": decrease_groups
        }
    
    def extract_migration_data(self, region: str) -> Dict[str, Any]:
        """인구이동 데이터 추출"""
        df = self.cache.get_sheet('I(순인구이동)집계')
        
        # 총계 행
        total_row = df[(df[4] == region) & (df[5].astype(str) == '0')]
        if len(total_row) == 0:
            return {"net_migration": 0, "direction": "순유출",
                    "inflow_age_groups": [], "outflow_age_groups": []}
        
        total_row = total_row.iloc[0]
        quarter_col = self.QUARTER_COLS['I(순인구이동)집계']['2025_2Q']
        net_migration = int(total_row[quarter_col]) if pd.notna(total_row[quarter_col]) else 0
        
        # 연령대별 데이터
        age_groups = df[(df[4] == region) & (df[5].astype(str) == '1')]
        
        inflow_groups = []
        outflow_groups = []
        
        if len(age_groups) > 0:
            for _, row in age_groups.iterrows():
                age_name = self._clean_name(row[7])
                count = int(row[quarter_col]) if pd.notna(row[quarter_col]) else 0
                
                if count > 0:
                    inflow_groups.append({"name": age_name, "count": count})
                elif count < 0:
                    outflow_groups.append({"name": age_name, "count": abs(count)})
        
        # 정렬 (절대값 기준)
        inflow_groups = sorted(inflow_groups, key=lambda x: x['count'], reverse=True)[:2]
        outflow_groups = sorted(outflow_groups, key=lambda x: x['count'], reverse=True)[:2]
        
        return {
            "net_migration": abs(net_migration),
            "direction": "순유입" if net_migration > 0 else "순유출",
            "inflow_age_groups": inflow_groups,
            "outflow_age_groups": outflow_groups
        }
    
    def _get_nationwide_chart_data(self, sheet_name: str, region_col: int, 
                                   code_col: int, total_code: str) -> List[Dict[str, Any]]:
        """전국 차트 데이터 추출 (캐싱)"""
        cache_key = f"{sheet_name}_{total_code}"
        if cache_key not in self._nationwide_data:
            df = self.cache.get_sheet(sheet_name)
            self._nationwide_data[cache_key] = self._get_chart_time_series(
                df, '전국', region_col, code_col, total_code, sheet_name
            )
        return self._nationwide_data[cache_key]
    
    def extract_chart_data(self, region: str) -> Dict[str, Any]:
        """차트 데이터 추출"""
        charts = {}
        
        # 1. 광공업생산 차트
        df_mfg = self.cache.get_sheet('A 분석')
        nationwide_mfg = self._get_nationwide_chart_data('A 분석', 3, 6, 'BCD')
        region_mfg = self._get_chart_time_series(df_mfg, region, 3, 6, 'BCD', 'A 분석')
        
        charts['manufacturing'] = {
            "title": "< 광공업생산 전년동분기대비 증감률(%) >",
            "yAxisMin": -12,
            "yAxisMax": 12,
            "yAxisStep": 4,
            "series": [
                {"name": "전국", "data": nationwide_mfg, "color": "#1f77b4"},
                {"name": region, "data": region_mfg, "color": "#ff7f0e"}
            ]
        }
        
        # 2. 서비스업생산 차트
        df_svc = self.cache.get_sheet('B 분석')
        nationwide_svc = self._get_nationwide_chart_data('B 분석', 3, 6, 'E~S')
        region_svc = self._get_chart_time_series(df_svc, region, 3, 6, 'E~S', 'B 분석')
        
        charts['service'] = {
            "title": "< 서비스업생산 전년동분기대비 증감률(%) >",
            "yAxisMin": 0,
            "yAxisMax": 8,
            "yAxisStep": 2,
            "series": [
                {"name": "전국", "data": nationwide_svc, "color": "#1f77b4"},
                {"name": region, "data": region_svc, "color": "#ff7f0e"}
            ]
        }
        
        # 3. 소매판매 차트
        df_retail = self.cache.get_sheet('C 분석')
        nationwide_retail = self._get_nationwide_chart_data('C 분석', 3, 4, '0')
        region_retail = self._get_chart_time_series(df_retail, region, 3, 4, '0', 'C 분석')
        
        charts['retail'] = {
            "title": "< 소매판매 전년동분기대비 증감률(%) >",
            "yAxisMin": -8,
            "yAxisMax": 0,
            "yAxisStep": 2,
            "series": [
                {"name": "전국", "data": nationwide_retail, "color": "#1f77b4"},
                {"name": region, "data": region_retail, "color": "#ff7f0e"}
            ]
        }
        
        # 4. 건설수주 차트
        df_const = self.cache.get_sheet("F'분석")
        nationwide_const = self._get_nationwide_chart_data("F'분석", 2, 3, '0')
        region_const = self._get_chart_time_series(df_const, region, 2, 3, '0', "F'분석")
        
        charts['construction'] = {
            "title": "< 건설수주 전년동분기대비 증감률(%) >",
            "yAxisMin": -60,
            "yAxisMax": 150,
            "yAxisStep": 30,
            "series": [
                {"name": "전국", "data": nationwide_const, "color": "#1f77b4"},
                {"name": region, "data": region_const, "color": "#ff7f0e"}
            ]
        }
        
        # 5. 수출입 차트 (4개 시리즈)
        df_export = self.cache.get_sheet('G 분석')
        df_import = self.cache.get_sheet('H 분석')
        
        nationwide_export = self._get_nationwide_chart_data('G 분석', 3, 4, '0')
        region_export = self._get_chart_time_series(df_export, region, 3, 4, '0', 'G 분석')
        nationwide_import = self._get_nationwide_chart_data('H 분석', 3, 4, '0')
        region_import = self._get_chart_time_series(df_import, region, 3, 4, '0', 'H 분석')
        
        charts['export_import'] = {
            "title": "< 수출입 전년동분기대비 증감률(%) >",
            "yAxisMin": -24,
            "yAxisMax": 24,
            "yAxisStep": 12,
            "series": [
                {"name": f"수출(전국)", "data": nationwide_export, "color": "#1f77b4"},
                {"name": f"수출({region})", "data": region_export, "color": "#ff7f0e"},
                {"name": f"수입(전국)", "data": nationwide_import, "color": "#1f77b4", "dashStyle": "dash"},
                {"name": f"수입({region})", "data": region_import, "color": "#ff7f0e", "dashStyle": "dash"}
            ]
        }
        
        # 6. 소비자물가 차트
        df_price = self.cache.get_sheet('E(지출목적물가) 분석')
        
        # 총지수 데이터 추출 (별도 처리 필요)
        def get_price_series(df, region_name):
            rows = df[(df[3] == region_name) & (df[4].astype(str) == '0') & (df[8] == '총지수')]
            if len(rows) == 0:
                return []
            row = rows.iloc[0]
            quarters = ["'23.2/4", "'24.2/4", "'25.2/4"]
            keys = ['2023_2Q', '2024_2Q', '2025_2Q']
            return [{"period": q, "value": self._get_quarter_value(row, 'E(지출목적물가) 분석', k)} for q, k in zip(quarters, keys)]
        
        charts['consumer_price'] = {
            "title": "< 소비자물가 전년동분기대비 등락률(%) >",
            "yAxisMin": 0,
            "yAxisMax": 5,
            "yAxisStep": 1,
            "series": [
                {"name": "전국", "data": get_price_series(df_price, '전국'), "color": "#1f77b4"},
                {"name": region, "data": get_price_series(df_price, region), "color": "#ff7f0e"}
            ]
        }
        
        # 7. 고용률 차트
        df_emp = self.cache.get_sheet('D(고용률)분석')
        
        def get_emp_series(df, region_name, age_filter=None):
            if age_filter:
                rows = df[(df[2] == region_name) & (df[5].astype(str).str.contains(age_filter, na=False))]
            else:
                rows = df[(df[2] == region_name) & (df[3].astype(str) == '0')]
            if len(rows) == 0:
                return []
            row = rows.iloc[0]
            quarters = ["'23.2/4", "'24.2/4", "'25.2/4"]
            keys = ['2023_2Q', '2024_2Q', '2025_2Q']
            return [{"period": q, "value": self._get_quarter_value(row, 'D(고용률)분석', k)} for q, k in zip(quarters, keys)]
        
        charts['employment_rate'] = {
            "title": "< 고용률 전년동분기대비 증감(%p) >",
            "yAxisMin": -8,
            "yAxisMax": 8,
            "yAxisStep": 3,
            "series": [
                {"name": "전국", "data": get_emp_series(df_emp, '전국'), "color": "#1f77b4"},
                {"name": region, "data": get_emp_series(df_emp, region), "color": "#ff7f0e"},
                {"name": f"{region} 20-29세", "data": get_emp_series(df_emp, region, '15 - 29'), "color": "#2ca02c"}
            ]
        }
        
        # 8. 인구순이동 차트
        df_mig = self.cache.get_sheet('I(순인구이동)집계')
        
        def get_mig_series(df, region_name, age_filter=None, convert_to_thousands=True):
            if age_filter:
                # 20-29세 = 20~24세 + 25~29세
                rows_20_24 = df[(df[4] == region_name) & (df[7].astype(str).str.contains('20~24', na=False))]
                rows_25_29 = df[(df[4] == region_name) & (df[7].astype(str).str.contains('25~29', na=False))]
                if len(rows_20_24) == 0 or len(rows_25_29) == 0:
                    return []
                
                quarters = ["'23.2/4", "'24.2/4", "'25.2/4"]
                keys = ['2023_2Q', '2024_2Q', '2025_2Q']
                result = []
                for q, k in zip(quarters, keys):
                    col = self.QUARTER_COLS['I(순인구이동)집계'][k]
                    val_20_24 = float(rows_20_24.iloc[0][col]) if pd.notna(rows_20_24.iloc[0][col]) else 0
                    val_25_29 = float(rows_25_29.iloc[0][col]) if pd.notna(rows_25_29.iloc[0][col]) else 0
                    total = val_20_24 + val_25_29
                    if convert_to_thousands:
                        total = round(total / 1000, 1)
                    result.append({"period": q, "value": total})
                return result
            else:
                rows = df[(df[4] == region_name) & (df[5].astype(str) == '0')]
                if len(rows) == 0:
                    return []
                row = rows.iloc[0]
                quarters = ["'23.2/4", "'24.2/4", "'25.2/4"]
                keys = ['2023_2Q', '2024_2Q', '2025_2Q']
                result = []
                for q, k in zip(quarters, keys):
                    col = self.QUARTER_COLS['I(순인구이동)집계'][k]
                    val = float(row[col]) if pd.notna(row[col]) else 0
                    if convert_to_thousands:
                        val = round(val / 1000, 1)
                    result.append({"period": q, "value": val})
                return result
        
        charts['population_migration'] = {
            "title": "< 인구순이동(천명) >",
            "yAxisMin": -20,
            "yAxisMax": 20,
            "yAxisStep": 10,
            "series": [
                {"name": region, "data": get_mig_series(df_mig, region), "color": "#1f77b4"},
                {"name": f"{region}(20-29세)", "data": get_mig_series(df_mig, region, '20-29'), "color": "#ff7f0e"}
            ]
        }
        
        return charts
    
    def extract_summary_table(self, region: str) -> Dict[str, Any]:
        """주요지표 요약 테이블 데이터 추출"""
        region_info = next((r for r in REGIONS if r.name == region), None)
        if not region_info:
            return {}
        
        periods = ["'23.2/4", "'24.2/4", "'25.1/4", "'25.2/4"]
        quarter_keys = ['2023_2Q', '2024_2Q', '2025_1Q', '2025_2Q']
        
        rows = []
        
        for period, q_key in zip(periods, quarter_keys):
            row_data = {"period": period}
            
            # 광공업생산
            df = self.cache.get_sheet('A 분석')
            mfg_row = df[(df[3] == region) & (df[6] == 'BCD')]
            if len(mfg_row) > 0:
                row_data['manufacturing'] = self._get_quarter_value(mfg_row.iloc[0], 'A 분석', q_key)
            else:
                row_data['manufacturing'] = 0
            
            # 서비스업생산
            df = self.cache.get_sheet('B 분석')
            svc_row = df[(df[3] == region) & (df[6] == 'E~S')]
            if len(svc_row) > 0:
                row_data['service'] = self._get_quarter_value(svc_row.iloc[0], 'B 분석', q_key)
            else:
                row_data['service'] = 0
            
            # 소매판매
            df = self.cache.get_sheet('C 분석')
            retail_row = df[(df[3] == region) & (df[4].astype(str) == '0')]
            if len(retail_row) > 0:
                row_data['retail'] = self._get_quarter_value(retail_row.iloc[0], 'C 분석', q_key)
            else:
                row_data['retail'] = 0
            
            # 건설수주
            df = self.cache.get_sheet("F'분석")
            const_row = df[(df[2] == region) & (df[3].astype(str) == '0')]
            if len(const_row) > 0:
                val = const_row.iloc[0][self.QUARTER_COLS["F'분석"][q_key]]
                if pd.isna(val) or val == '없음' or val == '-':
                    row_data['construction'] = 0
                else:
                    try:
                        row_data['construction'] = round(float(val), 1)
                    except (ValueError, TypeError):
                        row_data['construction'] = 0
            else:
                row_data['construction'] = 0
            
            # 수출
            df = self.cache.get_sheet('G 분석')
            exp_row = df[(df[3] == region) & (df[4].astype(str) == '0')]
            if len(exp_row) > 0:
                row_data['export'] = self._get_quarter_value(exp_row.iloc[0], 'G 분석', q_key)
            else:
                row_data['export'] = 0
            
            # 수입
            df = self.cache.get_sheet('H 분석')
            imp_row = df[(df[3] == region) & (df[4].astype(str) == '0')]
            if len(imp_row) > 0:
                row_data['import'] = self._get_quarter_value(imp_row.iloc[0], 'H 분석', q_key)
            else:
                row_data['import'] = 0
            
            # 소비자물가
            df = self.cache.get_sheet('E(지출목적물가) 분석')
            price_row = df[(df[3] == region) & (df[4].astype(str) == '0') & (df[8] == '총지수')]
            if len(price_row) > 0:
                row_data['consumer_price'] = self._get_quarter_value(price_row.iloc[0], 'E(지출목적물가) 분석', q_key)
            else:
                row_data['consumer_price'] = 0
            
            # 고용률
            df = self.cache.get_sheet('D(고용률)분석')
            emp_row = df[(df[2] == region) & (df[3].astype(str) == '0')]
            if len(emp_row) > 0:
                row_data['employment_rate'] = self._get_quarter_value(emp_row.iloc[0], 'D(고용률)분석', q_key)
            else:
                row_data['employment_rate'] = 0
            
            # 인구순이동 (천명 단위)
            df = self.cache.get_sheet('I(순인구이동)집계')
            mig_row = df[(df[4] == region) & (df[5].astype(str) == '0')]
            if len(mig_row) > 0:
                col = self.QUARTER_COLS['I(순인구이동)집계'][q_key]
                val = float(mig_row.iloc[0][col]) if pd.notna(mig_row.iloc[0][col]) else 0
                row_data['migration_total'] = round(val / 1000, 1)
            else:
                row_data['migration_total'] = 0
            
            # 인구순이동 20-29세
            mig_20_24 = df[(df[4] == region) & (df[7].astype(str).str.contains('20~24', na=False))]
            mig_25_29 = df[(df[4] == region) & (df[7].astype(str).str.contains('25~29', na=False))]
            if len(mig_20_24) > 0 and len(mig_25_29) > 0:
                col = self.QUARTER_COLS['I(순인구이동)집계'][q_key]
                val_20_24 = float(mig_20_24.iloc[0][col]) if pd.notna(mig_20_24.iloc[0][col]) else 0
                val_25_29 = float(mig_25_29.iloc[0][col]) if pd.notna(mig_25_29.iloc[0][col]) else 0
                row_data['migration_20_29'] = round((val_20_24 + val_25_29) / 1000, 1)
            else:
                row_data['migration_20_29'] = 0
            
            rows.append(row_data)
        
        return {
            "title": f"《 {region_info.full_name} 주요지표 》",
            "rows": rows
        }
    
    def extract_all_data(self, region: str) -> Dict[str, Any]:
        """지역별 모든 데이터 추출"""
        region_info = next((r for r in REGIONS if r.name == region), None)
        if not region_info:
            raise ValueError(f"Unknown region: {region}")
        
        return {
            "report_info": {
                "year": 2025,
                "quarter": 2,
                "region": region,
                "region_full_name": region_info.full_name,
                "region_code": region_info.code,
                "region_index": region_info.index,
                "page_number": 15 + region_info.index  # 페이지 번호 계산
            },
            "production": {
                "manufacturing": self.extract_manufacturing_data(region),
                "service": self.extract_service_data(region)
            },
            "consumption_construction": {
                "retail": self.extract_retail_data(region),
                "construction": self.extract_construction_data(region)
            },
            "export_import_price": {
                "export": self.extract_export_data(region),
                "import": self.extract_import_data(region),
                "consumer_price": self.extract_consumer_price_data(region)
            },
            "employment_migration": {
                "employment_rate": self.extract_employment_data(region),
                "population_migration": self.extract_migration_data(region)
            },
            "summary_table": self.extract_summary_table(region),
            "charts": self.extract_chart_data(region)
        }
    
    def render_html(self, region: str, template_path: str, output_path: str = None) -> str:
        """HTML 보고서 렌더링"""
        data = self.extract_all_data(region)
        
        # Jinja2 환경 설정
        template_dir = Path(template_path).parent
        env = Environment(loader=FileSystemLoader(str(template_dir)))
        template = env.get_template(Path(template_path).name)
        
        # 렌더링
        html_content = template.render(**data)
        
        # 파일 저장
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            print(f"보고서가 생성되었습니다: {output_path}")
        
        return html_content
    
    def render_all_regions(self, template_path: str, output_dir: str):
        """17개 시도 모두 렌더링"""
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        for region_info in REGIONS:
            output_file = output_path / f"{region_info.name}_output.html"
            self.render_html(region_info.name, template_path, str(output_file))
            print(f"Generated: {output_file}")
    
    def export_data_json(self, region: str, output_path: str):
        """추출된 데이터를 JSON으로 내보내기"""
        data = self.extract_all_data(region)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print(f"데이터가 저장되었습니다: {output_path}")


def main():
    """메인 실행 함수"""
    import argparse
    
    parser = argparse.ArgumentParser(description='시도별 경제동향 보고서 생성기')
    parser.add_argument('--excel', '-e', required=True, help='엑셀 파일 경로')
    parser.add_argument('--template', '-t', required=True, help='템플릿 파일 경로')
    parser.add_argument('--region', '-r', help='지역명 (예: 서울). 미지정시 전체 생성')
    parser.add_argument('--output', '-o', help='출력 HTML 파일 경로 또는 디렉토리')
    parser.add_argument('--json', '-j', help='데이터 JSON 출력 경로')
    parser.add_argument('--all', '-a', action='store_true', help='17개 시도 전체 생성')
    
    args = parser.parse_args()
    
    generator = 시도별Generator(args.excel)
    
    if args.all:
        output_dir = args.output or 'output'
        generator.render_all_regions(args.template, output_dir)
    elif args.region:
        if args.json:
            generator.export_data_json(args.region, args.json)
        
        if args.output:
            generator.render_html(args.region, args.template, args.output)
        elif not args.json:
            html = generator.render_html(args.region, args.template)
            print(html)
    else:
        print("지역을 지정하거나 --all 옵션을 사용하세요.")
        print("사용 가능한 지역: ", [r.name for r in REGIONS])


if __name__ == '__main__':
    main()

