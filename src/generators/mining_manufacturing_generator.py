#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
광공업생산 보도자료 생성기

엑셀 데이터를 읽어 스키마에 맞게 데이터를 추출하고,
Jinja2 템플릿을 사용하여 HTML 보도자료를 생성합니다.
"""

import pandas as pd
import json
from pathlib import Path
from typing import Dict, Any, Optional, Tuple, List

from .base import BaseGenerator
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from src.utils.excel_heuristic_parser import ExcelHeuristicParser
from utils.column_detector import get_column_mapping


class 광공업생산Generator(BaseGenerator):
    """광공업생산 보도자료 생성 클래스"""
    
    # 업종명 매핑 사전 (엑셀 데이터 → 보도자료 표기명)
    INDUSTRY_NAME_MAP = {
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
        "의복, 의복액세서리 및 모피제품 제조업": "의복",
        "코크스, 연탄 및 석유정제품 제조업": "석유정제품",
        "목재 및 나무제품 제조업; 가구 제외": "목재제품",
        "비금속광물 광업; 연료용 제외": "비금속광물광업",
    }
    
    # 표에 포함되는 지역 그룹
    REGION_GROUPS = {
        "전 국": {"regions": ["전 국"], "group": None},
        "경인": {"regions": ["서 울", "인 천", "경 기"], "group": "경인"},
        "충청": {"regions": ["대 전", "세 종", "충 북", "충 남"], "group": "충청"},
        "호남": {"regions": ["광 주", "전 북", "전 남", "제 주"], "group": "호남"},
        "동북": {"regions": ["대 구", "경 북", "강 원"], "group": "동북"},
        "동남": {"regions": ["부 산", "울 산", "경 남"], "group": "동남"},
    }
    
    # 지역명 정규화 (띄어쓰기 포함 표기)
    REGION_DISPLAY_MAP = {
        "전국": "전 국",
        "서울": "서 울",
        "부산": "부 산",
        "대구": "대 구",
        "인천": "인 천",
        "광주": "광 주",
        "대전": "대 전",
        "울산": "울 산",
        "세종": "세 종",
        "경기": "경 기",
        "강원": "강 원",
        "충북": "충 북",
        "충남": "충 남",
        "전북": "전 북",
        "전남": "전 남",
        "경북": "경 북",
        "경남": "경 남",
        "제주": "제 주",
    }
    
    def __init__(
        self,
        excel_path: str,
        raw_excel_path: Optional[str] = None,
        year: Optional[int] = None,
        quarter: Optional[int] = None
    ):
        """
        초기화
        
        Args:
            excel_path: 엑셀 파일 경로
            raw_excel_path: 기초자료 엑셀 파일 경로 (선택적)
            year: 연도 (선택적)
            quarter: 분기 (선택적)
        """
        super().__init__(excel_path, raw_excel_path, year, quarter)
        self.df_analysis: Optional[pd.DataFrame] = None
        self.df_aggregation: Optional[pd.DataFrame] = None
        self.use_raw_data: bool = False  # 기초자료 시트 사용 여부
        
    def load_data(self) -> None:
        """엑셀 데이터 로드 (휴리스틱 파서 사용)"""
        parser = ExcelHeuristicParser(str(self.excel_path))
        
        try:
            # 집계 시트 찾기 (수집표만 사용)
            agg_result = parser.find_target_sheet(
                keywords=['광공업생산', '광공업생산지수'],
                required_columns=['지역', '분류', '산업'],
                required_row_labels=['전국', 'BCD', '총지수']
            )
            
            if agg_result:
                agg_sheet_name, self.df_aggregation = agg_result
                self.use_raw_data = True  # 수집표만 사용
                print(f"[시트] 수집표 시트 사용: '{agg_sheet_name}'")
            else:
                raise ValueError(f"광공업생산 시트를 찾을 수 없습니다. 시트: {parser.xl.sheet_names}")
            
            # 분석 시트 찾기 (수집표만 사용)
            analysis_result = parser.find_target_sheet(
                keywords=['광공업생산', '광공업생산지수'],
                required_columns=['지역', '분류', '산업'],
                required_row_labels=['전국']
            )
            
            if analysis_result:
                analysis_sheet_name, self.df_analysis = analysis_result
                self.use_raw_data = True  # 수집표만 사용
                print(f"[시트] 수집표 시트 사용: '{analysis_sheet_name}'")
                
                # 헤더 행 동적 찾기
                header_row = parser.locate_table_start(
                    self.df_analysis,
                    anchor_keywords=['지역', '분류', '산업', '2024', '2025']
                )
                
                # 분석 시트에 실제 데이터가 있는지 확인 (수식 미계산 체크)
                # 동적 컬럼 감지를 사용하여 '전국' 행 찾기
                test_row = None
                if header_row is not None:
                    # 헤더 행 이후에서 '전국' 찾기
                    for col_idx in [1, 3, 4]:  # 기초자료/분석표 형식 모두 고려
                        if col_idx < len(self.df_analysis.columns):
                            test_row = self.df_analysis[
                                (self.df_analysis[col_idx] == '전국') | 
                                (self.df_analysis[col_idx].astype(str).str.contains('전국', na=False))
                            ]
                            if not test_row.empty:
                                break
                
                if test_row is None or test_row.empty or test_row.iloc[0].isna().sum() > 20:
                    print(f"[광공업생산] 분석 시트가 비어있음 → 집계 시트에서 직접 계산")
                    self.use_aggregation_only = True
                else:
                    self.use_aggregation_only = False
            else:
                # 분석 시트를 찾지 못한 경우 집계 시트 사용
                self.df_analysis = self.df_aggregation.copy()
                self.use_aggregation_only = True
                print(f"[광공업생산] 분석 시트를 찾을 수 없음 → 집계 시트 사용")
        
        finally:
            parser.close()
        
    def _get_industry_display_name(self, raw_name: str) -> str:
        """업종명을 보도자료 표기명으로 변환
        
        우선순위:
        1. 분류 이름 축약.xlsx 파일의 매핑 (BaseExtractor 사용)
        2. 하드코딩된 INDUSTRY_NAME_MAP
        3. 원본 이름 그대로 반환
        """
        if not raw_name:
            return raw_name
        
        # 공백 제거
        cleaned = raw_name.strip().replace("\u3000", "").replace("　", "")
        
        # 1. 분류 이름 축약.xlsx 파일의 매핑 사용 (우선)
        try:
            from extractors.base import BaseExtractor
            mapping = BaseExtractor.load_classification_name_mapping()
            if '광공업' in mapping and cleaned in mapping['광공업']:
                shortened = mapping['광공업'][cleaned]
                if shortened:  # 축약 이름이 있으면 사용
                    return shortened
        except Exception as e:
            # 매핑 파일 로드 실패 시 하드코딩된 매핑 사용
            pass
        
        # 2. 하드코딩된 매핑 사용
        for key, value in self.INDUSTRY_NAME_MAP.items():
            if key in cleaned or cleaned in key:
                return value
        
        # 3. 원본 이름 반환
        return cleaned
    
    def _get_region_display_name(self, raw_name: str) -> str:
        """지역명을 표시용으로 변환"""
        return self.REGION_DISPLAY_MAP.get(raw_name, raw_name)
    
    def generate(self) -> Dict[str, Any]:
        """
        보도자료 데이터 생성 (BaseGenerator 추상 메서드 구현)
        
        Returns:
            생성된 데이터 딕셔너리
        """
        return self.extract_all_data()
    
    def extract_nationwide_data(self) -> Dict[str, Any]:
        """전국 데이터 추출"""
        # 집계 시트 기반으로 추출 (분석 시트가 비어있는 경우 포함)
        if hasattr(self, 'use_aggregation_only') and self.use_aggregation_only:
            return self._extract_nationwide_from_aggregation()
        
        if self.use_raw_data:
            return self._extract_nationwide_from_raw_data()
        
        df = self.df_analysis
        
        # 전국 총지수 행 찾기 (컬럼 인덱스 자동 감지)
        nationwide_total = None
        for col_offset in [0, 1]:  # df[3] 또는 df[4]에 지역명
            try:
                region_col = 3 + col_offset
                code_col = 6 + col_offset
                filtered = df[(df[region_col] == '전국') & (df[code_col] == 'BCD')]
                if not filtered.empty:
                    nationwide_total = filtered.iloc[0]
                    break
            except:
                continue
        
        if nationwide_total is None:
            return self._extract_nationwide_from_aggregation()
        
        # 전국 중분류 데이터 (분류단계 2)
        class_col = 4 + col_offset
        contrib_col = 28 + col_offset if 28 + col_offset < len(df.columns) else 28
        
        try:
            nationwide_industries = df[(df[region_col] == '전국') & (df[class_col].astype(str) == '2') & (pd.notna(df[contrib_col]))]
            sorted_industries = nationwide_industries.sort_values(contrib_col, ascending=False)
            increase_industries = sorted_industries[sorted_industries[contrib_col] > 0]
            decrease_industries = sorted_industries[sorted_industries[contrib_col] < 0].sort_values(contrib_col, ascending=True)
        except:
            increase_industries = pd.DataFrame()
            decrease_industries = pd.DataFrame()
        
        # 광공업생산지수 (집계 시트에서) - 동적 컬럼 감지 적용
        df_agg = self.df_aggregation
        column_mapping = get_column_mapping(df_agg, self.year, self.quarter)
        region_col = column_mapping['region_col']
        code_col = column_mapping['code_col']
        current_quarter_col = column_mapping.get('current_quarter_col')
        
        nationwide_agg_rows = df_agg[(df_agg[region_col] == '전국') & (df_agg[code_col] == 'BCD')]
        if nationwide_agg_rows.empty:
            production_index = None
        else:
            nationwide_agg = nationwide_agg_rows.iloc[0]
            production_index = self.safe_float(nationwide_agg[current_quarter_col], None) if current_quarter_col is not None else None
        
        growth_col = 21 + col_offset if 21 + col_offset < len(nationwide_total) else 21
        growth_rate = self.safe_float(nationwide_total[growth_col], None) if growth_col < len(nationwide_total) else None  # PM 요구사항: None으로 처리
        
        industry_name_col = 7 + col_offset
        
        return {
            "production_index": self.safe_float(production_index, None),  # PM 요구사항: None으로 처리
            "growth_rate": self.safe_round(growth_rate, 1, None),  # PM 요구사항: None으로 처리
            "growth_direction": "증가" if (growth_rate is not None and growth_rate > 0) else ("감소" if (growth_rate is not None and growth_rate < 0) else "N/A"),
            "main_increase_industries": [
                {
                    "name": self._get_industry_display_name(str(row[industry_name_col]) if pd.notna(row[industry_name_col]) else ''),
                    "growth_rate": self.safe_round(row[growth_col], 1, None) if growth_col < len(row) else None,  # PM 요구사항: None으로 처리
                    "contribution": self.safe_round(row[contrib_col], 6, None) if contrib_col < len(row) else None  # PM 요구사항: None으로 처리
                }
                for _, row in increase_industries.head(5).iterrows()
            ],
            "main_decrease_industries": [
                {
                    "name": self._get_industry_display_name(str(row[industry_name_col]) if pd.notna(row[industry_name_col]) else ''),
                    "growth_rate": self.safe_round(row[growth_col], 1, None) if growth_col < len(row) else None,  # PM 요구사항: None으로 처리
                    "contribution": self.safe_round(row[contrib_col], 6, None) if contrib_col < len(row) else None  # PM 요구사항: None으로 처리
                }
                for _, row in decrease_industries.head(5).iterrows()
            ]
        }
    
    def _extract_nationwide_from_aggregation(self) -> Dict[str, Any]:
        """집계 시트에서 전국 데이터 추출 (증감률 직접 계산)"""
        df = self.df_aggregation
        
        # 기초자료 형식: 1=지역이름, 2=분류단계, 3=가중치, 4=산업코드, 5=산업이름
        # 분석표 형식: 4=지역이름, 5=분류단계, 6=가중치, 7=산업코드, 8=산업이름
        # 형식 자동 감지
        is_raw_format = False
        if len(df.columns) > 1 and pd.notna(df.iloc[2, 1]) and '지역' in str(df.iloc[2, 1]):
            # 기초자료 형식 (헤더 행 2에 "지역"이 열 1에 있음)
            is_raw_format = True
            region_col = 1
            class_col = 2
            weight_col = 3
            code_col = 4
            name_col = 5
        else:
            # 분석표 형식
            region_col = 4
            class_col = 5
            weight_col = 6
            code_col = 7
            name_col = 8
        
        # 분기 데이터 컬럼 찾기 (헤더 행에서)
        header_row_idx = 2 if is_raw_format else 0
        if header_row_idx >= len(df):
            header_row_idx = 0
        
        header_row = df.iloc[header_row_idx]
        
        # 2024 2/4, 2025 2/4, 2025 3/4p 컬럼 찾기
        col_2024_2q = None
        col_2025_2q = None
        col_2025_3q = None
        
        for col_idx in range(len(header_row)):
            val = str(header_row[col_idx]) if pd.notna(header_row[col_idx]) else ''
            val_clean = val.strip().replace('.', ' ').replace('p', '').replace('P', '')
            
            if '2024' in val_clean and '2/4' in val_clean:
                col_2024_2q = col_idx
            if '2025' in val_clean and '2/4' in val_clean:
                col_2025_2q = col_idx
            if '2025' in val_clean and '3/4' in val_clean:
                col_2025_3q = col_idx
        
        # 분기 선택 (3분기면 3분기, 아니면 2분기)
        if self.quarter == 3 and col_2025_3q is not None:
            current_quarter_col = col_2025_3q
        elif col_2025_2q is not None:
            current_quarter_col = col_2025_2q
        else:
            # 기본값: 마지막 분기 컬럼
            current_quarter_col = len(header_row) - 1
        
        if col_2024_2q is None:
            # 2024 2/4를 찾지 못한 경우, 현재 분기의 전년동분기 찾기
            target_quarter = self.quarter if self.quarter else 2
            for col_idx in range(len(header_row)):
                val = str(header_row[col_idx]) if pd.notna(header_row[col_idx]) else ''
                val_clean = val.strip().replace('.', ' ').replace('p', '').replace('P', '')
                if '2024' in val_clean and f'{target_quarter}/4' in val_clean:
                    col_2024_2q = col_idx
                    break
        
        # 전국 총지수 행 (BCD)
        nationwide_rows = df[(df[region_col] == '전국') & (df[code_col] == 'BCD')]
        if nationwide_rows.empty:
            return self._get_default_nationwide_data()
        
        nationwide_total = nationwide_rows.iloc[0]
        
        # 당분기와 전년동분기 지수로 증감률 계산
        current_index = self.safe_float(nationwide_total[current_quarter_col], None)
        prev_year_index = self.safe_float(nationwide_total[col_2024_2q], None) if col_2024_2q is not None else None
        
        if current_index is None or prev_year_index is None or prev_year_index == 0:
            growth_rate = 0.0
        else:
            growth_rate = ((current_index - prev_year_index) / prev_year_index) * 100
        
        # 전국 중분류 업종별 데이터 (분류단계 2)
        nationwide_industries = df[(df[region_col] == '전국') & (df[class_col].astype(str) == '2')]
        
        industries = []
        for _, row in nationwide_industries.iterrows():
            curr = self.safe_float(row[current_quarter_col], None)
            prev = self.safe_float(row[col_2024_2q], None) if col_2024_2q is not None else None
            weight = self.safe_float(row[weight_col], None)  # PM 요구사항: None으로 처리
            
            # PM 요구사항: 데이터가 없으면 None 반환
            if curr is None or prev is None or prev == 0:
                continue  # 데이터 없으면 건너뛰기
            if prev_year_index is None or prev_year_index == 0 or weight is None:
                continue  # 계산 불가능하면 건너뛰기
            
            if curr is None or prev is None or prev == 0:
                continue
            if prev_year_index is None or prev_year_index == 0 or weight is None:
                continue
            
            ind_growth = ((curr - prev) / prev) * 100
            # 기여도 = (당기 - 전기) / 전국전기 * 가중치/10000
            contribution = (curr - prev) / prev_year_index * weight / 10000 * 100
            industries.append({
                'name': self._get_industry_display_name(str(row[name_col]) if pd.notna(row[name_col]) else ''),
                'growth_rate': round(ind_growth, 1),
                'contribution': round(contribution, 6)
            })
        
        # 기여도 순 정렬
        increase_industries = sorted([i for i in industries if i['contribution'] > 0], 
                                    key=lambda x: x['contribution'], reverse=True)[:5]
        decrease_industries = sorted([i for i in industries if i['contribution'] < 0], 
                                    key=lambda x: x['contribution'])[:5]
        
        return {
            "production_index": current_index,
            "growth_rate": round(growth_rate, 1),
            "growth_direction": "증가" if growth_rate > 0 else "감소",
            "main_increase_industries": increase_industries,
            "main_decrease_industries": decrease_industries
        }
    
    def _extract_nationwide_from_raw_data(self) -> Dict[str, Any]:
        """기초자료 시트에서 전국 데이터 추출"""
        df = self.df_analysis
        
        # 기초자료 시트 구조: 1=지역이름, 2=분류단계, 4=품목명, 5~=연도/분기 데이터
        # 헤더 행 찾기
        header_row = None
        for i in range(min(10, len(df))):
            row = df.iloc[i]
            row_str = ' '.join([str(v) for v in row.values[:10] if pd.notna(v)])
            if '지역' in row_str and ('2024' in row_str or '2025' in row_str):
                header_row = i
                break
        
        if header_row is None:
            header_row = 2
        
        # 당분기 컬럼 찾기 (마지막에서 두번째 또는 2025.2/4 등)
        current_quarter_col = None
        prev_year_col = None
        header = df.iloc[header_row] if header_row < len(df) else df.iloc[0]
        
        for col_idx in range(len(header) - 1, 4, -1):
            col_val = str(header[col_idx]) if pd.notna(header[col_idx]) else ''
            if '2025' in col_val and ('2/4' in col_val or '2' in col_val):
                current_quarter_col = col_idx
            if '2024' in col_val and ('2/4' in col_val or '2' in col_val):
                prev_year_col = col_idx
            if current_quarter_col and prev_year_col:
                break
        
        if current_quarter_col is None:
            current_quarter_col = len(header) - 2
        if prev_year_col is None:
            prev_year_col = current_quarter_col - 4
        
        # 전국 총지수 행 찾기
        nationwide_row = None
        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            region = str(row[1]).strip() if pd.notna(row[1]) else ''
            classification = str(row[2]).strip() if pd.notna(row[2]) else ''
            if region == '전국' and classification == '0':
                nationwide_row = row
                break
        
        if nationwide_row is None:
            return self._get_default_nationwide_data()
        
        # 증감률 계산
        current_val = self.safe_float(nationwide_row[current_quarter_col], 100)
        prev_val = self.safe_float(nationwide_row[prev_year_col], 100)
        if prev_val and prev_val != 0:
            growth_rate = ((current_val - prev_val) / prev_val) * 100
        else:
            growth_rate = 0.0
        
        # 업종별 데이터 추출 (분류단계 2)
        industries = []
        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            region = str(row[1]).strip() if pd.notna(row[1]) else ''
            classification = str(row[2]).strip() if pd.notna(row[2]) else ''
            if region == '전국' and classification == '2':
                current = self.safe_float(row[current_quarter_col], None)
                prev = self.safe_float(row[prev_year_col], None)
                if current is not None and prev is not None and prev != 0:
                    ind_growth = ((current - prev) / prev) * 100
                    industries.append({
                        'name': self._get_industry_display_name(str(row[4]) if pd.notna(row[4]) else ''),
                        'growth_rate': round(ind_growth, 1),
                        'contribution': round(ind_growth * 0.1, 6)  # 추정 기여도
                    })
        
        # 증가/감소 업종 분류
        increase_industries = sorted([i for i in industries if i['growth_rate'] > 0], 
                                    key=lambda x: x['growth_rate'], reverse=True)[:5]
        decrease_industries = sorted([i for i in industries if i['growth_rate'] < 0], 
                                    key=lambda x: x['growth_rate'])[:5]
        
        return {
            "production_index": current_val,
            "growth_rate": round(growth_rate, 1),
            "growth_direction": "증가" if growth_rate > 0 else "감소",
            "main_increase_industries": increase_industries,
            "main_decrease_industries": decrease_industries
        }
    
    def _get_default_nationwide_data(self) -> Dict[str, Any]:
        """기본 전국 데이터"""
        return {
            "production_index": 100.0,
            "growth_rate": 0.0,
            "growth_direction": "감소",
            "main_increase_industries": [],
            "main_decrease_industries": []
        }
    
    def extract_regional_data(self) -> Dict[str, Any]:
        """시도별 데이터 추출"""
        # 집계 시트 기반으로 추출
        if hasattr(self, 'use_aggregation_only') and self.use_aggregation_only:
            return self._extract_regional_from_aggregation()
        
        if self.use_raw_data:
            return self._extract_regional_from_raw_data()
        
        df = self.df_analysis
        
        # 개별 시도 목록 (수도, 충청 등 권역 제외)
        individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 컬럼 인덱스 자동 감지
        col_offset = 0
        for offset in [0, 1]:
            test = df[df[3 + offset] == '전국']
            if not test.empty:
                col_offset = offset
                break
        
        region_col = 3 + col_offset
        code_col = 6 + col_offset
        growth_col = 21 + col_offset
        contrib_col = 28 + col_offset if 28 + col_offset < len(df.columns) else 28
        industry_name_col = 7 + col_offset
        
        regions_data = []
        
        for region in individual_regions:
            # 해당 지역 총지수
            region_total = df[(df[region_col] == region) & (df[code_col] == 'BCD')]
            if region_total.empty:
                continue
            region_total = region_total.iloc[0]
            
            growth_rate = self.safe_float(region_total[growth_col], None) if growth_col < len(region_total) else None  # PM 요구사항: None으로 처리
            
            # 해당 지역 업종별 데이터
            try:
                region_industries = df[(df[region_col] == region) & (pd.notna(df[contrib_col]))]
                
                # 기여도 순 정렬 (증가는 높은 순, 감소는 낮은 순)
                if growth_rate >= 0:
                    sorted_ind = region_industries.sort_values(contrib_col, ascending=False)
                else:
                    sorted_ind = region_industries.sort_values(contrib_col, ascending=True)
                
                # 상위 3개 업종 (BCD 제외)
                top_industries = []
                industry_count = 0
                for _, row in sorted_ind.iterrows():
                    if industry_count >= 3:
                        break
                    if pd.notna(row[industry_name_col]) and str(row[code_col]) != 'BCD':
                        top_industries.append({
                            "name": self._get_industry_display_name(str(row[industry_name_col])),
                            "growth_rate": self.safe_round(row[growth_col], 1, None) if growth_col < len(row) else None,  # PM 요구사항: None으로 처리
                            "contribution": self.safe_round(row[contrib_col], 6, None) if contrib_col < len(row) else None  # PM 요구사항: None으로 처리
                        })
                        industry_count += 1
            except:
                top_industries = []
            
            regions_data.append({
                "region": region,
                "growth_rate": round(growth_rate, 1),
                "top_industries": top_industries
            })
        
        # 증가/감소 지역 분류
        increase_regions = sorted(
            [r for r in regions_data if r["growth_rate"] > 0],
            key=lambda x: x["growth_rate"],
            reverse=True
        )
        
        decrease_regions = sorted(
            [r for r in regions_data if r["growth_rate"] < 0],
            key=lambda x: x["growth_rate"]  # 가장 낮은 값(큰 감소)이 먼저
        )
        
        return {
            "increase_regions": increase_regions,
            "decrease_regions": decrease_regions,
            "region_count": len(increase_regions)
        }
    
    def _extract_regional_from_aggregation(self) -> Dict[str, Any]:
        """집계 시트에서 시도별 데이터 추출"""
        df = self.df_aggregation
        
        # 형식 자동 감지 (전국 데이터 추출과 동일한 로직)
        is_raw_format = False
        if len(df.columns) > 1 and pd.notna(df.iloc[2, 1]) and '지역' in str(df.iloc[2, 1]):
            is_raw_format = True
            region_col = 1
            class_col = 2
            weight_col = 3
            code_col = 4
            name_col = 5
        else:
            region_col = 4
            class_col = 5
            weight_col = 6
            code_col = 7
            name_col = 8
        
        # 분기 컬럼 찾기
        header_row_idx = 2 if is_raw_format else 0
        if header_row_idx >= len(df):
            header_row_idx = 0
        
        header_row = df.iloc[header_row_idx]
        
        col_2024_2q = None
        col_2025_2q = None
        col_2025_3q = None
        
        for col_idx in range(len(header_row)):
            val = str(header_row[col_idx]) if pd.notna(header_row[col_idx]) else ''
            val_clean = val.strip().replace('.', ' ').replace('p', '').replace('P', '')
            
            if '2024' in val_clean and '2/4' in val_clean:
                col_2024_2q = col_idx
            if '2025' in val_clean and '2/4' in val_clean:
                col_2025_2q = col_idx
            if '2025' in val_clean and '3/4' in val_clean:
                col_2025_3q = col_idx
        
        if self.quarter == 3 and col_2025_3q is not None:
            current_quarter_col = col_2025_3q
        elif col_2025_2q is not None:
            current_quarter_col = col_2025_2q
        else:
            current_quarter_col = len(header_row) - 1
        
        if col_2024_2q is None:
            target_quarter = self.quarter if self.quarter else 2
            for col_idx in range(len(header_row)):
                val = str(header_row[col_idx]) if pd.notna(header_row[col_idx]) else ''
                val_clean = val.strip().replace('.', ' ').replace('p', '').replace('P', '')
                if '2024' in val_clean and f'{target_quarter}/4' in val_clean:
                    col_2024_2q = col_idx
                    break
        
        individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 전국 전년동분기 지수 (기여도 계산용)
        nationwide_rows = df[(df[region_col] == '전국') & (df[code_col] == 'BCD')]
        nationwide_prev = self.safe_float(nationwide_rows.iloc[0][col_2024_2q], None) if (not nationwide_rows.empty and col_2024_2q is not None) else None
        
        regions_data = []
        
        for region in individual_regions:
            # 해당 지역 총지수 (BCD)
            region_total = df[(df[region_col] == region) & (df[code_col] == 'BCD')]
            if region_total.empty:
                continue
            region_total = region_total.iloc[0]
            
            # 증감률 계산
            current = self.safe_float(region_total[current_quarter_col], None)
            prev = self.safe_float(region_total[col_2024_2q], None) if col_2024_2q is not None else None
            
            if current is None or prev is None or prev == 0:
                growth_rate = 0.0
            else:
                growth_rate = ((current - prev) / prev) * 100
            
            # 해당 지역 업종별 데이터 (분류단계 2)
            region_industries = df[(df[region_col] == region) & (df[class_col].astype(str) == '2')]
            
            industries = []
            for _, row in region_industries.iterrows():
                curr = self.safe_float(row[current_quarter_col], None)
                prev_ind = self.safe_float(row[col_2024_2q], None) if col_2024_2q is not None else None
                weight = self.safe_float(row[weight_col], None)  # PM 요구사항: None으로 처리
                
                # PM 요구사항: 데이터가 없으면 None 반환
                if curr is None or prev_ind is None or prev_ind == 0:
                    continue  # 데이터 없으면 건너뛰기
                if nationwide_prev is None or nationwide_prev == 0 or weight is None:
                    continue  # 계산 불가능하면 건너뛰기
                
                ind_growth = ((curr - prev_ind) / prev_ind) * 100
                contribution = (curr - prev_ind) / nationwide_prev * weight / 10000 * 100
                industries.append({
                    'name': self._get_industry_display_name(str(row[name_col]) if pd.notna(row[name_col]) else ''),
                    'growth_rate': round(ind_growth, 1),
                    'contribution': round(contribution, 6)
                })
            
            # 기여도 순 정렬
            if growth_rate >= 0:
                sorted_ind = sorted(industries, key=lambda x: x['contribution'], reverse=True)
            else:
                sorted_ind = sorted(industries, key=lambda x: x['contribution'])
            
            regions_data.append({
                "region": region,
                "growth_rate": round(growth_rate, 1),
                "top_industries": sorted_ind[:3]
            })
        
        # 증가/감소 지역 분류
        increase_regions = sorted(
            [r for r in regions_data if r["growth_rate"] > 0],
            key=lambda x: x["growth_rate"],
            reverse=True
        )
        
        decrease_regions = sorted(
            [r for r in regions_data if r["growth_rate"] < 0],
            key=lambda x: x["growth_rate"]
        )
        
        # 전체 지역 데이터 (all 필드) - 증감률 순으로 정렬
        all_regions = sorted(regions_data, key=lambda x: x["growth_rate"], reverse=True)
        
        return {
            "increase_regions": increase_regions,
            "decrease_regions": decrease_regions,
            "all": all_regions,
            "region_count": len(increase_regions)
        }
    
    def _extract_regional_from_raw_data(self) -> Dict[str, Any]:
        """기초자료 시트에서 시도별 데이터 추출"""
        df = self.df_analysis
        
        individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 헤더 행 및 컬럼 인덱스 찾기
        header_row = 2
        current_quarter_col = None
        prev_year_col = None
        
        for i in range(min(10, len(df))):
            row = df.iloc[i]
            row_str = ' '.join([str(v) for v in row.values[:10] if pd.notna(v)])
            if '지역' in row_str and ('2024' in row_str or '2025' in row_str):
                header_row = i
                break
        
        header = df.iloc[header_row] if header_row < len(df) else df.iloc[0]
        
        for col_idx in range(len(header) - 1, 4, -1):
            col_val = str(header[col_idx]) if pd.notna(header[col_idx]) else ''
            if '2025' in col_val and ('2/4' in col_val or '2' in col_val):
                current_quarter_col = col_idx
            if '2024' in col_val and ('2/4' in col_val or '2' in col_val):
                prev_year_col = col_idx
            if current_quarter_col and prev_year_col:
                break
        
        if current_quarter_col is None:
            current_quarter_col = len(header) - 2
        if prev_year_col is None:
            prev_year_col = current_quarter_col - 4
        
        regions_data = []
        
        for region in individual_regions:
            # 해당 지역 총지수 (분류단계 0) 찾기
            region_row = None
            for i in range(header_row + 1, len(df)):
                row = df.iloc[i]
                r_name = str(row[1]).strip() if pd.notna(row[1]) else ''
                classification = str(row[2]).strip() if pd.notna(row[2]) else ''
                if r_name == region and classification == '0':
                    region_row = row
                    break
            
            if region_row is None:
                continue
            
            # 증감률 계산
            current_val = self.safe_float(region_row[current_quarter_col], None)
            prev_val = self.safe_float(region_row[prev_year_col], None)
            
            if current_val is not None and prev_val is not None and prev_val != 0:
                growth_rate = ((current_val - prev_val) / prev_val) * 100
            else:
                growth_rate = 0.0
            
            # 해당 지역의 업종별 데이터 (분류단계 2)
            industries = []
            for i in range(header_row + 1, len(df)):
                row = df.iloc[i]
                r_name = str(row[1]).strip() if pd.notna(row[1]) else ''
                classification = str(row[2]).strip() if pd.notna(row[2]) else ''
                if r_name == region and classification == '2':
                    current = self.safe_float(row[current_quarter_col], None)
                    prev = self.safe_float(row[prev_year_col], None)
                    if current is not None and prev is not None and prev != 0:
                        ind_growth = ((current - prev) / prev) * 100
                        industries.append({
                            'name': self._get_industry_display_name(str(row[4]) if pd.notna(row[4]) else ''),
                            'growth_rate': round(ind_growth, 1),
                            'contribution': round(ind_growth * 0.1, 6)
                        })
            
            # 기여도 순 정렬
            if growth_rate >= 0:
                sorted_ind = sorted(industries, key=lambda x: x['contribution'], reverse=True)
            else:
                sorted_ind = sorted(industries, key=lambda x: x['contribution'])
            
            regions_data.append({
                "region": region,
                "growth_rate": round(growth_rate, 1),
                "top_industries": sorted_ind[:3]
            })
        
        # 증가/감소 지역 분류
        increase_regions = sorted(
            [r for r in regions_data if r["growth_rate"] > 0],
            key=lambda x: x["growth_rate"],
            reverse=True
        )
        
        decrease_regions = sorted(
            [r for r in regions_data if r["growth_rate"] < 0],
            key=lambda x: x["growth_rate"]
        )
        
        return {
            "increase_regions": increase_regions,
            "decrease_regions": decrease_regions,
            "region_count": len(increase_regions)
        }
    
    def extract_summary_box(self) -> Dict[str, Any]:
        """회색 요약 박스 데이터 추출"""
        regional = self.extract_regional_data()
        
        # 증가 지역 중 상위 3개
        top_increase = regional["increase_regions"][:3]
        
        main_regions = []
        for r in top_increase:
            # 스키마 준수: industries는 항상 string[] 타입
            industries = []
            if "top_industries" in r and r["top_industries"]:
                for ind in r["top_industries"][:2]:
                    if isinstance(ind, dict) and "name" in ind:
                        industries.append(str(ind["name"]))
                    elif isinstance(ind, str):
                        industries.append(ind)
            main_regions.append({
                "region": r["region"],
                "industries": industries  # 항상 string[]
            })
        
        return {
            "main_increase_regions": main_regions,
            "region_count": regional["region_count"]
        }
    
    def extract_top3_regions(self) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        """상위 3개 증가/감소 지역 추출 (< 주요 증감 지역 및 업종 > 섹션용)"""
        regional = self.extract_regional_data()
        
        top3_increase = []
        for r in regional["increase_regions"][:3]:
            top3_increase.append({
                "region": r["region"],
                "growth_rate": r["growth_rate"],
                "industries": r["top_industries"][:3]
            })
        
        top3_decrease = []
        for r in regional["decrease_regions"][:3]:
            top3_decrease.append({
                "region": r["region"],
                "growth_rate": r["growth_rate"],
                "industries": r["top_industries"][:3]
            })
        
        return top3_increase, top3_decrease
    
    def extract_summary_table(self) -> Dict[str, Any]:
        """하단 표 데이터 추출"""
        df_agg = self.df_aggregation
        
        # 컬럼 정의
        columns = {
            "growth_rate_columns": ["2023.2/4", "2024.2/4", "2025.1/4", "2025.2/4p"],
            "index_columns": ["2024.2/4", "2025.2/4p"]
        }
        
        # 지역 순서 정의
        region_order = [
            {"region": "전국", "group": None},
            {"region": "서울", "group": "경인", "rowspan": 3},
            {"region": "인천", "group": None},
            {"region": "경기", "group": None},
            {"region": "대전", "group": "충청", "rowspan": 4},
            {"region": "세종", "group": None},
            {"region": "충북", "group": None},
            {"region": "충남", "group": None},
            {"region": "광주", "group": "호남", "rowspan": 4},
            {"region": "전북", "group": None},
            {"region": "전남", "group": None},
            {"region": "제주", "group": None},
            {"region": "대구", "group": "동북", "rowspan": 3},
            {"region": "경북", "group": None},
            {"region": "강원", "group": None},
            {"region": "부산", "group": "동남", "rowspan": 3},
            {"region": "울산", "group": None},
            {"region": "경남", "group": None},
        ]
        
        # 집계 시트 컬럼 매핑:
        # 18=2023.2/4, 22=2024.2/4, 25=2025.1/4, 26=2025.2/4p
        
        regions_data = []
        
        for r_info in region_order:
            region = r_info["region"]
            
            # 집계 데이터에서 해당 지역 찾기
            region_agg = df_agg[(df_agg[4] == region) & (df_agg[7] == 'BCD')]
            if region_agg.empty:
                continue
            region_agg = region_agg.iloc[0]
            
            # 지수 추출
            idx_2023_2q = self.safe_float(region_agg[18], 100)  # 2023.2/4
            idx_2024_2q = self.safe_float(region_agg[22], 100)  # 2024.2/4
            idx_2025_1q = self.safe_float(region_agg[25], 100)  # 2025.1/4
            idx_2025_2q = self.safe_float(region_agg[26], 100)  # 2025.2/4p
            
            # 전년동분기 지수로 증감률 계산
            idx_2022_2q = self.safe_float(region_agg[14], 100)  # 2022.2/4 (2023.2/4의 전년동분기)
            idx_2023_2q_for_2024 = self.safe_float(region_agg[18], 100)  # 2023.2/4 (2024.2/4의 전년동분기)
            idx_2024_1q = self.safe_float(region_agg[21], 100)  # 2024.1/4 (2025.1/4의 전년동분기)
            
            def calc_growth(curr, prev):
                # PM 요구사항: 데이터가 없으면 None 반환
                if curr is None or prev is None:
                    return None  # N/A 처리
                if prev != 0:
                    return round(((curr - prev) / prev) * 100, 1)
                return None  # 0으로 나누기 방지, N/A 처리
            
            growth_rates = [
                calc_growth(idx_2023_2q, idx_2022_2q),
                calc_growth(idx_2024_2q, idx_2023_2q_for_2024),
                calc_growth(idx_2025_1q, idx_2024_1q),
                calc_growth(idx_2025_2q, idx_2024_2q),
            ]
            
            indices = [
                round(idx_2024_2q, 1),
                round(idx_2025_2q, 1),
            ]
            
            row_data = {
                "region": self._get_region_display_name(region),
                "growth_rates": growth_rates,
                "indices": indices
            }
            
            if r_info.get("group"):
                row_data["group"] = r_info["group"]
                row_data["rowspan"] = r_info.get("rowspan", 1)
            
            regions_data.append(row_data)
        
        return {
            "title": "《 광공업생산지수 및 증감률》",
            "base_year": 2020,
            "columns": columns,
            "regions": regions_data
        }
    
    def extract_all_data(self) -> Dict[str, Any]:
        """모든 데이터 추출"""
        self.load_data()
        
        nationwide = self.extract_nationwide_data()
        regional = self.extract_regional_data()
        summary_box = self.extract_summary_box()
        top3_increase, top3_decrease = self.extract_top3_regions()
        summary_table = self.extract_summary_table()
        
        return {
            "report_info": {
                "year": 2025,
                "quarter": 2,
                "data_source": "국가데이터처 국가통계포털(KOSIS), 광업제조업동향조사"
            },
            "nationwide_data": nationwide,
            "regional_data": regional,
            "summary_box": summary_box,
            "top3_increase_regions": top3_increase,
            "top3_decrease_regions": top3_decrease,
            "summary_table": summary_table
        }
    


def main():
    """메인 실행 함수"""
    import argparse
    
    parser = argparse.ArgumentParser(description='광공업생산 보도자료 생성기')
    parser.add_argument('--excel', '-e', required=True, help='엑셀 파일 경로')
    parser.add_argument('--template', '-t', required=True, help='템플릿 파일 경로')
    parser.add_argument('--output', '-o', help='출력 HTML 파일 경로')
    parser.add_argument('--json', '-j', help='데이터 JSON 출력 경로')
    
    args = parser.parse_args()
    
    generator = 광공업생산Generator(args.excel)
    
    if args.json:
        generator.extract_all_data()  # 데이터 로드
        generator.export_data_json(args.json)
    
    if args.output:
        generator.render_html(args.template, args.output)
    elif not args.json:
        # 출력 경로가 지정되지 않으면 stdout으로 출력
        html = generator.render_html(args.template)
        print(html)


if __name__ == '__main__':
    main()

