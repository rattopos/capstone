#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
광공업생산 보도자료 생성기

엑셀 데이터를 읽어 스키마에 맞게 데이터를 추출하고,
Jinja2 템플릿을 사용하여 HTML 보도자료를 생성합니다.
"""

import pandas as pd
import json
from jinja2 import Environment, FileSystemLoader
from pathlib import Path
try:
    from .base_generator import BaseGenerator
except ImportError:
    # 직접 실행 시 상대 import 실패 방지
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).parent))
    from base_generator import BaseGenerator


class 광공업생산Generator(BaseGenerator):
    """광공업생산 보도자료 생성 클래스"""
    
    # ========================================
    # 분석 시트 컬럼 인덱스 (0-based, 문서 7.2 참조)
    # ========================================
    # A분석 시트 컬럼 구조:
    COL_REGION_NAME = 3        # 지역이름
    COL_CLASSIFICATION = 4     # 분류단계 (문서에 명시되지 않았으나 일반적으로 4)
    COL_INDUSTRY_CODE = 6      # 산업/업태코드
    COL_INDUSTRY_NAME = 7      # 산업/업태명
    COL_WEIGHT = 8             # 가중치
    COL_GROWTH_RATE = 21       # 2025_2Q 증감률 (문서: 2025_2Q 증감률)
    COL_CONTRIBUTION = 28      # 기여도
    
    # ========================================
    # 집계 시트 컬럼 인덱스 (1-based → 0-based 변환)
    # ========================================
    # A(광공업생산)집계 시트 구조 (문서 7.2):
    # 메타시작열: 4 (1-based) → 3 (0-based)
    # 연도시작열: 10 (1-based) → 9 (0-based)
    # 분기시작열: 15 (1-based) → 14 (0-based)
    # 가중치열: 7 (1-based) → 6 (0-based)
    AGG_COL_REGION_NAME = 3    # 지역이름 (메타시작열 4의 0-based: 3)
    AGG_COL_CLASSIFICATION = 4  # 분류단계 (메타시작열 + 1)
    AGG_COL_WEIGHT = 6         # 가중치 (가중치열 7의 0-based: 6)
    AGG_COL_INDUSTRY_CODE = 6  # 산업코드 (가중치열과 동일 위치 또는 다음 위치, 실제 확인 필요)
    AGG_COL_INDUSTRY_NAME = 7  # 산업이름 (산업코드 다음)
    AGG_COL_YEAR_START = 9     # 연도시작열 (2020년)
    AGG_COL_QUARTER_START = 14 # 분기시작열 (2022.1/4)
    # 분기 데이터: 14=2022.1/4, 15=2022.2/4, ..., 22=2024.2/4, 25=2025.1/4, 26=2025.2/4p
    AGG_COL_2024_2Q = 22       # 2024.2/4 (전년동분기)
    AGG_COL_2025_2Q = 26       # 2025.2/4p (당분기)
    AGG_COL_2023_2Q = 18       # 2023.2/4
    AGG_COL_2025_1Q = 25       # 2025.1/4
    AGG_COL_2024_1Q = 21       # 2024.1/4
    AGG_COL_2022_2Q = 14       # 2022.2/4
    
    # 데이터 시작 행 (보통 3~4행부터 시작, 0-based로는 2~3)
    DATA_START_ROW = 2         # 기본값: 3행 (0-based: 2)
    
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
    
    def __init__(self, excel_path: str, year=None, quarter=None):
        """
        초기화
        
        Args:
            excel_path: 엑셀 파일 경로
            year: 연도 (선택사항)
            quarter: 분기 (선택사항)
        """
        super().__init__(excel_path, year, quarter)
        self.df_analysis = None
        self.df_aggregation = None
        self.data = {}
        self.use_raw_data = False  # 기초자료 시트 사용 여부
        
    def load_data(self):
        """엑셀 데이터 로드"""
        xl = self.load_excel()
        
        # 집계 시트 찾기 (우선)
        agg_sheet, _ = self.find_sheet_with_fallback(
            ['A(광공업생산)집계', 'A 집계'],
            ['광공업생산', '광공업생산지수']
        )
        
        if agg_sheet:
            self.df_aggregation = self.get_sheet(agg_sheet)
        else:
            raise ValueError(f"광공업생산 집계 시트를 찾을 수 없습니다. 시트: {xl.sheet_names}")
        
        # 분석 시트 찾기
        analysis_sheet, self.use_raw_data = self.find_sheet_with_fallback(
            ['A 분석', 'A분석'],
            ['광공업생산', '광공업생산지수']
        )
        
        if analysis_sheet:
            self.df_analysis = self.get_sheet(analysis_sheet)
            # 분석 시트에 실제 데이터가 있는지 확인 (수식 미계산 체크)
            test_conditions = {self.COL_REGION_NAME: '전국'}
            has_data = self.check_sheet_has_data(
                self.df_analysis.iloc[self.DATA_START_ROW:],
                test_conditions,
                max_empty_cells=20
            )
            if not has_data:
                print(f"[광공업생산] 분석 시트가 비어있음 → 집계 시트에서 직접 계산")
                self.use_aggregation_only = True
            else:
                self.use_aggregation_only = False
        else:
            self.df_analysis = self.df_aggregation.copy()
            self.use_aggregation_only = True
        
    def _get_industry_display_name(self, raw_name: str) -> str:
        """업종명을 보도자료 표기명으로 변환"""
        # 공백 제거
        cleaned = raw_name.strip().replace("\u3000", "").replace("　", "")
        
        for key, value in self.INDUSTRY_NAME_MAP.items():
            if key in cleaned or cleaned in key:
                return value
        return cleaned
    
    def _get_region_display_name(self, raw_name: str) -> str:
        """지역명을 표시용으로 변환"""
        return self.REGION_DISPLAY_MAP.get(raw_name, raw_name)
    
    def extract_nationwide_data(self) -> dict:
        """전국 데이터 추출"""
        # 집계 시트 기반으로 추출 (분석 시트가 비어있는 경우 포함)
        if hasattr(self, 'use_aggregation_only') and self.use_aggregation_only:
            return self._extract_nationwide_from_aggregation()
        
        if self.use_raw_data:
            return self._extract_nationwide_from_raw_data()
        
        df = self.df_analysis
        
        # 헤더 행을 건너뛰고 데이터만 사용
        data_df = df.iloc[self.DATA_START_ROW:].copy()
        
        # 전국 총지수 행 찾기 (문서에 명시된 컬럼 인덱스 사용)
        nationwide_total = None
        try:
            filtered = data_df[(data_df[self.COL_REGION_NAME] == '전국') & 
                              (data_df[self.COL_INDUSTRY_CODE] == 'BCD')]
            if not filtered.empty:
                nationwide_total = filtered.iloc[0]
        except Exception as e:
            print(f"[광공업생산] 전국 데이터 찾기 실패: {e}")
        
        if nationwide_total is None:
            return self._extract_nationwide_from_aggregation()
        
        # 전국 중분류 데이터 (분류단계 2)
        try:
            nationwide_industries = data_df[(data_df[self.COL_REGION_NAME] == '전국') & 
                                            (data_df[self.COL_CLASSIFICATION].astype(str) == '2') & 
                                            (pd.notna(data_df[self.COL_CONTRIBUTION]))]
            sorted_industries = nationwide_industries.sort_values(self.COL_CONTRIBUTION, ascending=False)
            increase_industries = sorted_industries[sorted_industries[self.COL_CONTRIBUTION] > 0]
            decrease_industries = sorted_industries[sorted_industries[self.COL_CONTRIBUTION] < 0].sort_values(self.COL_CONTRIBUTION, ascending=True)
        except Exception as e:
            print(f"[광공업생산] 업종별 데이터 추출 실패: {e}")
            increase_industries = pd.DataFrame()
            decrease_industries = pd.DataFrame()
        
        # 광공업생산지수 (집계 시트에서)
        df_agg = self.df_aggregation
        nationwide_agg = df_agg[(df_agg[self.AGG_COL_REGION_NAME] == '전국') & 
                               (df_agg[self.AGG_COL_INDUSTRY_CODE] == 'BCD')].iloc[0]
        production_index = nationwide_agg[self.AGG_COL_2025_2Q]  # 2025.2/4p 컬럼
        
        # 증감률 추출 (문서에 명시된 컬럼 인덱스 사용)
        growth_rate = nationwide_total[self.COL_GROWTH_RATE] if pd.notna(nationwide_total[self.COL_GROWTH_RATE]) else 0
        
        return {
            "production_index": self.safe_float(production_index, 100.0),
            "growth_rate": self.safe_round(growth_rate, 1, 0.0),
            "growth_direction": "증가" if self.safe_float(growth_rate, 0) > 0 else "감소",
            "main_increase_industries": [
                {
                    "name": self._get_industry_display_name(str(row[self.COL_INDUSTRY_NAME]) if pd.notna(row[self.COL_INDUSTRY_NAME]) else ''),
                    "growth_rate": self.safe_round(row[self.COL_GROWTH_RATE], 1, 0.0) if self.COL_GROWTH_RATE < len(row) else 0.0,
                    "contribution": self.safe_round(row[self.COL_CONTRIBUTION], 6, 0.0) if self.COL_CONTRIBUTION < len(row) else 0.0
                }
                for _, row in increase_industries.head(5).iterrows()
            ],
            "main_decrease_industries": [
                {
                    "name": self._get_industry_display_name(str(row[self.COL_INDUSTRY_NAME]) if pd.notna(row[self.COL_INDUSTRY_NAME]) else ''),
                    "growth_rate": self.safe_round(row[self.COL_GROWTH_RATE], 1, 0.0) if self.COL_GROWTH_RATE < len(row) else 0.0,
                    "contribution": self.safe_round(row[self.COL_CONTRIBUTION], 6, 0.0) if self.COL_CONTRIBUTION < len(row) else 0.0
                }
                for _, row in decrease_industries.head(5).iterrows()
            ]
        }
    
    def _extract_nationwide_from_aggregation(self) -> dict:
        """집계 시트에서 전국 데이터 추출 (증감률 직접 계산)"""
        df = self.df_aggregation
        
        # 컬럼 구조: 4=지역이름, 5=분류단계, 6=가중치, 7=산업코드, 8=산업이름
        # 9~26: 연도/분기별 지수 데이터
        # 데이터 컬럼: 9=2020, 10=2021, 11=2022, 12=2023, 13=2024
        # 분기: 14=2022.1/4, ..., 22=2024.2/4, 23=2024.3/4, 24=2024.4/4, 25=2025.1/4, 26=2025.2/4p
        
        # 전국 총지수 행 (BCD) - 집계 시트는 1-based를 0-based로 변환
        nationwide_rows = df[(df[self.AGG_COL_REGION_NAME] == '전국') & 
                             (df[self.AGG_COL_INDUSTRY_CODE] == 'BCD')]
        if nationwide_rows.empty:
            return self._get_default_nationwide_data()
        
        nationwide_total = nationwide_rows.iloc[0]
        
        # 당분기(2025.2/4)와 전년동분기(2024.2/4) 지수로 증감률 계산
        current_index = self.safe_float(nationwide_total[self.AGG_COL_2025_2Q], 100)  # 2025.2/4p
        prev_year_index = self.safe_float(nationwide_total[self.AGG_COL_2024_2Q], 100)  # 2024.2/4
        
        if prev_year_index and prev_year_index != 0:
            growth_rate = ((current_index - prev_year_index) / prev_year_index) * 100
        else:
            growth_rate = 0.0
        
        # 전국 중분류 업종별 데이터 (분류단계 2)
        nationwide_industries = df[(df[self.AGG_COL_REGION_NAME] == '전국') & 
                                   (df[self.AGG_COL_CLASSIFICATION].astype(str) == '2')]
        
        industries = []
        for _, row in nationwide_industries.iterrows():
            curr = self.safe_float(row[self.AGG_COL_2025_2Q], None)
            prev = self.safe_float(row[self.AGG_COL_2024_2Q], None)
            weight = self.safe_float(row[self.AGG_COL_WEIGHT], 0)
            
            if curr is not None and prev is not None and prev != 0:
                ind_growth = ((curr - prev) / prev) * 100
                # 기여도 = (당기 - 전기) / 전국전기 * 가중치/10000
                contribution = (curr - prev) / prev_year_index * weight / 10000 * 100 if prev_year_index else 0
                industries.append({
                    'name': self._get_industry_display_name(str(row[self.AGG_COL_INDUSTRY_NAME]) if pd.notna(row[self.AGG_COL_INDUSTRY_NAME]) else ''),
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
    
    def _extract_nationwide_from_raw_data(self) -> dict:
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
    
    def _get_default_nationwide_data(self) -> dict:
        """기본 전국 데이터"""
        return {
            "production_index": 100.0,
            "growth_rate": 0.0,
            "growth_direction": "감소",
            "main_increase_industries": [],
            "main_decrease_industries": []
        }
    
    def extract_regional_data(self) -> dict:
        """시도별 데이터 추출"""
        # 집계 시트 기반으로 추출
        if hasattr(self, 'use_aggregation_only') and self.use_aggregation_only:
            return self._extract_regional_from_aggregation()
        
        if self.use_raw_data:
            return self._extract_regional_from_raw_data()
        
        df = self.df_analysis
        
        # 헤더 행을 건너뛰고 데이터만 사용
        data_df = df.iloc[self.DATA_START_ROW:].copy()
        
        # 개별 시도 목록 (수도, 충청 등 권역 제외)
        individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        regions_data = []
        
        for region in individual_regions:
            # 해당 지역 총지수 (문서에 명시된 컬럼 인덱스 사용)
            region_total = data_df[(data_df[self.COL_REGION_NAME] == region) & 
                                   (data_df[self.COL_INDUSTRY_CODE] == 'BCD')]
            if region_total.empty:
                continue
            region_total = region_total.iloc[0]
            
            growth_rate = self.safe_float(region_total[self.COL_GROWTH_RATE], 0)
            
            # 해당 지역 업종별 데이터
            try:
                region_industries = data_df[(data_df[self.COL_REGION_NAME] == region) & 
                                           (pd.notna(data_df[self.COL_CONTRIBUTION]))]
                
                # 기여도 순 정렬 (증가는 높은 순, 감소는 낮은 순)
                if growth_rate >= 0:
                    sorted_ind = region_industries.sort_values(self.COL_CONTRIBUTION, ascending=False)
                else:
                    sorted_ind = region_industries.sort_values(self.COL_CONTRIBUTION, ascending=True)
                
                # 상위 3개 업종 (BCD 제외)
                top_industries = []
                industry_count = 0
                for _, row in sorted_ind.iterrows():
                    if industry_count >= 3:
                        break
                    if pd.notna(row[self.COL_INDUSTRY_NAME]) and str(row[self.COL_INDUSTRY_CODE]) != 'BCD':
                        top_industries.append({
                            "name": self._get_industry_display_name(str(row[self.COL_INDUSTRY_NAME])),
                            "growth_rate": self.safe_round(row[self.COL_GROWTH_RATE], 1, 0.0) if self.COL_GROWTH_RATE < len(row) else 0.0,
                            "contribution": self.safe_round(row[self.COL_CONTRIBUTION], 6, 0.0) if self.COL_CONTRIBUTION < len(row) else 0.0
                        })
                        industry_count += 1
            except Exception as e:
                print(f"[광공업생산] {region} 업종별 데이터 추출 실패: {e}")
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
    
    def _extract_regional_from_aggregation(self) -> dict:
        """집계 시트에서 시도별 데이터 추출"""
        df = self.df_aggregation
        
        individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 전국 전년동분기 지수 (기여도 계산용)
        nationwide_rows = df[(df[self.AGG_COL_REGION_NAME] == '전국') & 
                            (df[self.AGG_COL_INDUSTRY_CODE] == 'BCD')]
        nationwide_prev = self.safe_float(nationwide_rows.iloc[0][self.AGG_COL_2024_2Q], 100) if not nationwide_rows.empty else 100
        
        regions_data = []
        
        for region in individual_regions:
            # 해당 지역 총지수 (BCD) - 집계 시트는 1-based를 0-based로 변환
            region_total = df[(df[self.AGG_COL_REGION_NAME] == region) & 
                             (df[self.AGG_COL_INDUSTRY_CODE] == 'BCD')]
            if region_total.empty:
                continue
            region_total = region_total.iloc[0]
            
            # 증감률 계산
            current = self.safe_float(region_total[self.AGG_COL_2025_2Q], 100)  # 2025.2/4p
            prev = self.safe_float(region_total[self.AGG_COL_2024_2Q], 100)  # 2024.2/4
            
            if prev and prev != 0:
                growth_rate = ((current - prev) / prev) * 100
            else:
                growth_rate = 0.0
            
            # 해당 지역 업종별 데이터 (분류단계 2)
            region_industries = df[(df[self.AGG_COL_REGION_NAME] == region) & 
                                  (df[self.AGG_COL_CLASSIFICATION].astype(str) == '2')]
            
            industries = []
            for _, row in region_industries.iterrows():
                curr = self.safe_float(row[self.AGG_COL_2025_2Q], None)
                prev_ind = self.safe_float(row[self.AGG_COL_2024_2Q], None)
                weight = self.safe_float(row[self.AGG_COL_WEIGHT], 0)
                
                if curr is not None and prev_ind is not None and prev_ind != 0:
                    ind_growth = ((curr - prev_ind) / prev_ind) * 100
                    contribution = (curr - prev_ind) / nationwide_prev * weight / 10000 * 100 if nationwide_prev else 0
                    industries.append({
                        'name': self._get_industry_display_name(str(row[self.AGG_COL_INDUSTRY_NAME]) if pd.notna(row[self.AGG_COL_INDUSTRY_NAME]) else ''),
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
        
        return {
            "increase_regions": increase_regions,
            "decrease_regions": decrease_regions,
            "region_count": len(increase_regions)
        }
    
    def _extract_regional_from_raw_data(self) -> dict:
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
    
    def extract_summary_box(self) -> dict:
        """회색 요약 박스 데이터 추출"""
        regional = self.extract_regional_data()
        
        # 증가 지역 중 상위 3개
        top_increase = regional["increase_regions"][:3]
        
        main_regions = []
        for r in top_increase:
            industries = [ind["name"] for ind in r["top_industries"][:2]]
            main_regions.append({
                "region": r["region"],
                "industries": industries
            })
        
        return {
            "main_increase_regions": main_regions,
            "region_count": regional["region_count"]
        }
    
    def extract_top3_regions(self) -> tuple:
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
    
    def extract_summary_table(self) -> dict:
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
        
        # 집계 시트 컬럼 매핑 (상수 사용):
        # AGG_COL_2023_2Q=18, AGG_COL_2024_2Q=22, AGG_COL_2025_1Q=25, AGG_COL_2025_2Q=26
        
        regions_data = []
        
        for r_info in region_order:
            region = r_info["region"]
            
            # 집계 데이터에서 해당 지역 찾기 (1-based를 0-based로 변환)
            region_agg = df_agg[(df_agg[self.AGG_COL_REGION_NAME] == region) & 
                               (df_agg[self.AGG_COL_INDUSTRY_CODE] == 'BCD')]
            if region_agg.empty:
                continue
            region_agg = region_agg.iloc[0]
            
            # 지수 추출 (상수 사용)
            idx_2023_2q = self.safe_float(region_agg[self.AGG_COL_2023_2Q], 100)  # 2023.2/4
            idx_2024_2q = self.safe_float(region_agg[self.AGG_COL_2024_2Q], 100)  # 2024.2/4
            idx_2025_1q = self.safe_float(region_agg[self.AGG_COL_2025_1Q], 100)  # 2025.1/4
            idx_2025_2q = self.safe_float(region_agg[self.AGG_COL_2025_2Q], 100)  # 2025.2/4p
            
            # 전년동분기 지수로 증감률 계산
            idx_2022_2q = self.safe_float(region_agg[self.AGG_COL_2022_2Q], 100)  # 2022.2/4 (2023.2/4의 전년동분기)
            idx_2023_2q_for_2024 = self.safe_float(region_agg[self.AGG_COL_2023_2Q], 100)  # 2023.2/4 (2024.2/4의 전년동분기)
            idx_2024_1q = self.safe_float(region_agg[self.AGG_COL_2024_1Q], 100)  # 2024.1/4 (2025.1/4의 전년동분기)
            
            def calc_growth(curr, prev):
                if prev and prev != 0:
                    return round(((curr - prev) / prev) * 100, 1)
                return 0.0
            
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
    
    def extract_all_data(self) -> dict:
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
    
    def render_html(self, template_path: str, output_path: str = None) -> str:
        """HTML 보도자료 렌더링"""
        data = self.extract_all_data()
        
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
            print(f"보도자료가 생성되었습니다: {output_path}")
        
        return html_content
    
    def export_data_json(self, output_path: str):
        """추출된 데이터를 JSON으로 내보내기"""
        data = self.extract_all_data()
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print(f"데이터가 저장되었습니다: {output_path}")


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

