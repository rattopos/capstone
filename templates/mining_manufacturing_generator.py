#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
광공업생산 보도자료 생성기

엑셀 데이터를 읽어 스키마에 맞게 데이터를 추출하고,
Jinja2 템플릿을 사용하여 HTML 보도자료를 생성합니다.
"""

import pandas as pd
import json
import re
from typing import Optional, Dict, Any
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


class MiningManufacturingGenerator(BaseGenerator):
    """광공업생산 보도자료 생성 클래스"""
    
    # 하위 호환성을 위한 별칭 (클래스 정의 후 파일 끝에서 할당)
    
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
    AGG_COL_REGION_NAME = 4    # 지역이름 (로그 기반 근본 수정: Index 4 확정)
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
        # 부모 생성자 호출 필수
        super().__init__(excel_path, year, quarter)
        self.report_id = 'mining'  # 보고서 ID (푸터 출처 매핑용)
        self.df_analysis = None
        self.df_aggregation = None
        self.data = {}
        self.use_raw_data = False  # 기초자료 시트 사용 여부
        
        # Step 2: 동적 할당 적용 - 하드코딩 절대 금지
        if self.year and self.quarter:
            from utils.excel_utils import get_period_context
            self.period_context = get_period_context(self.year, self.quarter)
            # 변수 선언 (코드 내의 모든 날짜 로직은 이 변수들만 사용)
            self.target_year = self.period_context['target_year']
            self.target_quarter = self.period_context['target_quarter']
            self.prev_q_year = self.period_context['prev_q_year']
            self.prev_q = self.period_context['prev_q']
            self.prev_y_year = self.period_context['prev_y_year']
            self.prev_y_quarter = self.period_context['prev_y_quarter']
            self.target_period = self.period_context['target_period']
            self.prev_q_period = self.period_context['prev_q_period']
            self.prev_y_period = self.period_context['prev_y_period']
            self.target_key = self.period_context['target_key']
            self.prev_quarter_key = self.period_context['prev_quarter_key']
            self.prev_year_key = self.period_context['prev_year_key']
        else:
            self.period_context = None
        
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
    
    def extract_nationwide_data(self, table_data: list = None) -> dict:
        """
        Step 2: Text Variables 추출 (Table Data 재사용)
        table_data에서 전국 데이터를 추출하여 사용
        """
        # table_data가 제공되지 않으면 SSOT에서 추출
        if table_data is None:
            table_data = self._extract_table_data_ssot()
        
        # table_data에서 전국 데이터 찾기 (SSOT)
        nationwide_table = next((d for d in table_data if d["region_name"] == "전국"), None)
        if nationwide_table is None:
            # fallback: 기존 로직 사용
            if hasattr(self, 'use_aggregation_only') and self.use_aggregation_only:
                return self._extract_nationwide_from_aggregation()
            if self.use_raw_data:
                return self._extract_nationwide_from_raw_data()
            return self._extract_nationwide_from_analysis()
        
        # SSOT에서 가져온 전국 데이터 사용
        production_index = nationwide_table["value"]  # 2025.2/4p 지수
        growth_rate = nationwide_table["change_rate"]  # 증감률 (반올림 완료)
        
        # 업종별 데이터는 기존 로직 유지 (엑셀에서 추출)
        if hasattr(self, 'use_aggregation_only') and self.use_aggregation_only:
            industries_data = self._extract_nationwide_industries_from_aggregation()
        elif self.use_raw_data:
            industries_data = self._extract_nationwide_industries_from_raw_data()
        else:
            industries_data = self._extract_nationwide_industries_from_analysis()
        
        return {
            "production_index": production_index,
            "growth_rate": growth_rate,  # SSOT에서 가져온 값
            "growth_direction": "증가" if growth_rate > 0 else "감소",
            "main_increase_industries": industries_data.get("increase", []),
            "main_decrease_industries": industries_data.get("decrease", [])
        }
    
    def _extract_nationwide_industries_from_analysis(self) -> dict:
        """분석 시트에서 전국 업종별 데이터 추출"""
        df = self.df_analysis
        data_df = df.iloc[self.DATA_START_ROW:].copy()
        
        # Step 2: 동적 컬럼 매핑 - 헤더에서 컬럼 찾기
        print(f"\n[광공업생산 분석시트] ===== 헤더 정보 분석 =====")
        for header_row_idx in range(min(5, len(df))):
            header_row = df.iloc[header_row_idx]
            header_values = [str(v) if pd.notna(v) else 'NaN' for v in header_row.values[:35]]
            print(f"[광공업생산 분석시트] 헤더 행 {header_row_idx}: {header_values}")
        
        # 동적으로 컬럼 인덱스 찾기
        growth_rate_col = self._find_column_by_header(df, ['증감률', 'growth', 'rate'], search_rows=5)
        contribution_col = self._find_column_by_header(df, ['기여도', 'contribution', '기여'], search_rows=5)
        weight_col = self._find_column_by_header(df, ['가중치', 'weight', '비중'], search_rows=5)
        
        # fallback: 기존 하드코딩된 인덱스 사용
        if growth_rate_col is None:
            growth_rate_col = self.COL_GROWTH_RATE
            print(f"[광공업생산 분석시트] ⚠️ 증감률 컬럼 fallback: {growth_rate_col}")
        else:
            print(f"[광공업생산 분석시트] ✅ 증감률 컬럼 발견: {growth_rate_col} (기존: {self.COL_GROWTH_RATE})")
        
        if contribution_col is None:
            contribution_col = self.COL_CONTRIBUTION
            print(f"[광공업생산 분석시트] ⚠️ 기여도 컬럼 fallback: {contribution_col}")
        else:
            print(f"[광공업생산 분석시트] ✅ 기여도 컬럼 발견: {contribution_col} (기존: {self.COL_CONTRIBUTION})")
        
        if weight_col is None:
            weight_col = self.COL_WEIGHT
            print(f"[광공업생산 분석시트] ⚠️ 가중치 컬럼 fallback: {weight_col}")
        else:
            print(f"[광공업생산 분석시트] ✅ 가중치 컬럼 발견: {weight_col} (기존: {self.COL_WEIGHT})")
        
        # 소비재 필터링 목록
        consumer_goods = ['식료품', '음료']
        
        try:
            nationwide_industries = data_df[(data_df[self.COL_REGION_NAME] == '전국') & 
                                            (data_df[self.COL_CLASSIFICATION].astype(str) == '2')]
            
            industries = []
            for _, row in nationwide_industries.iterrows():
                raw_growth = row[growth_rate_col] if growth_rate_col < len(row) else None
                raw_weight = row[weight_col] if weight_col < len(row) else None
                
                growth_rate = self.safe_float(raw_growth, None)
                weight = self.safe_float(raw_weight, 0)
                
                if growth_rate is None:
                    continue
                
                industry_name = self._get_industry_display_name(str(row[self.COL_INDUSTRY_NAME]) if pd.notna(row[self.COL_INDUSTRY_NAME]) else '')
                industry_code = str(row[self.COL_INDUSTRY_CODE]) if self.COL_INDUSTRY_CODE < len(row) and pd.notna(row[self.COL_INDUSTRY_CODE]) else ''
                
                # 소비재 필터링
                if any(consumer in industry_name for consumer in consumer_goods):
                    continue
                
                # 가중치가 없거나 0인 경우 하드코딩 가중치 사용
                if weight == 0 or weight is None:
                    weight = self._get_industry_weight_fallback(industry_name, industry_code)
                    print(f"[광공업생산] 가중치 없음 - {industry_name}에 하드코딩 가중치 {weight} 적용")
                
                # 기여도 절대값 = (증감률 × 가중치)의 절대값
                # 가중치가 이미 퍼센트 단위이므로 100으로 나누지 않고 직접 곱셈
                contribution_abs = abs(growth_rate * weight / 100)
                
                industries.append({
                    "name": industry_name,
                    "growth_rate": self.safe_round(growth_rate, 1, 0.0),
                    "contribution": self.safe_round(raw_weight, 6, 0.0) if raw_weight is not None else 0.0,  # 기존 호환성
                    "contribution_abs": contribution_abs  # 정렬용 절대값
                })
            
            # 기여도 절대값 순 정렬
            increase_industries = sorted([i for i in industries if i['growth_rate'] > 0], 
                                        key=lambda x: x['contribution_abs'], reverse=True)[:2]
            decrease_industries = sorted([i for i in industries if i['growth_rate'] < 0], 
                                        key=lambda x: x['contribution_abs'], reverse=True)[:2]
            
        except Exception as e:
            print(f"[광공업생산] 업종별 데이터 추출 실패: {e}")
            import traceback
            traceback.print_exc()
            return {"increase": [], "decrease": []}
        
        return {"increase": increase_industries, "decrease": decrease_industries}
    
    def _find_column_by_header(self, df: pd.DataFrame, patterns: list, search_rows: int = 5) -> Optional[int]:
        """헤더에서 패턴으로 컬럼 인덱스 찾기"""
        for row_idx in range(min(search_rows, len(df))):
            row = df.iloc[row_idx]
            for col_idx in range(len(row)):
                cell_value = row.iloc[col_idx] if hasattr(row, 'iloc') else row[col_idx]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()
                    for pattern in patterns:
                        if pattern in cell_str:
                            return col_idx
        return None
    
    def _extract_nationwide_industries_from_aggregation(self) -> dict:
        """집계 시트에서 전국 업종별 데이터 추출"""
        df = self.df_aggregation
        
        # 헤더 행 찾기
        header_row_idx = None
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            row_str = ' '.join([str(v) for v in row.values[:20] if pd.notna(v)])
            if any(year in row_str for year in ['2023', '2024', '2025', '2026']):
                header_row_idx = i
                break
        
        if header_row_idx is None:
            header_row_idx = 2
        
        header_row = df.iloc[header_row_idx]
        
        # 동적으로 타겟 컬럼 인덱스 찾기
        target_col_idx = self.find_target_col_index(header_row, self.year or 2025, self.quarter or 2)
        if target_col_idx == -1:
            print(f"[광공업생산] {self.year or 2025}년 {self.quarter or 2}분기 컬럼을 찾을 수 없어 기본 인덱스 사용")
            target_col_idx = self.AGG_COL_2025_2Q
        
        # 전년동분기 인덱스 계산 (상대적 위치 -4)
        prev_y_idx = target_col_idx - 4 if target_col_idx >= 4 else (self.AGG_COL_2024_2Q if hasattr(self, 'AGG_COL_2024_2Q') else target_col_idx - 1)
        
        # 전국 중분류 업종별 데이터 (분류단계 2)
        nationwide_industries = df[(df[self.AGG_COL_REGION_NAME] == '전국') & 
                                   (df[self.AGG_COL_CLASSIFICATION].astype(str) == '2')]
        
        # 전국 전년동분기 지수 (기여도 계산용)
        nationwide_rows = df[(df[self.AGG_COL_REGION_NAME] == '전국') & 
                            (df[self.AGG_COL_INDUSTRY_CODE] == 'BCD')]
        nationwide_prev = self.safe_float(nationwide_rows.iloc[0][prev_y_idx], 100) if not nationwide_rows.empty and prev_y_idx < len(nationwide_rows.iloc[0]) else 100
        
        # 소비재 필터링 목록
        consumer_goods = ['식료품', '음료']
        
        industries = []
        for _, row in nationwide_industries.iterrows():
            curr = self.safe_float(row[target_col_idx], None) if target_col_idx < len(row) else None
            prev = self.safe_float(row[prev_y_idx], None) if prev_y_idx < len(row) else None
            weight = self.safe_float(row[self.AGG_COL_WEIGHT], 0)
            
            if curr is not None and prev is not None and prev != 0:
                ind_growth = ((curr - prev) / prev) * 100
                contribution = (curr - prev) / nationwide_prev * weight / 10000 * 100 if nationwide_prev else 0
                
                industry_name = self._get_industry_display_name(str(row[self.AGG_COL_INDUSTRY_NAME]) if pd.notna(row[self.AGG_COL_INDUSTRY_NAME]) else '')
                
                # 소비재 필터링
                if any(consumer in industry_name for consumer in consumer_goods):
                    continue
                
                # 기여도 = (증감률 × 가중치)의 절대값
                contribution_abs = abs(ind_growth * weight / 10000)
                
                industries.append({
                    'name': industry_name,
                    'growth_rate': round(ind_growth, 1),
                    'contribution': round(contribution, 6),
                    'contribution_abs': contribution_abs  # 정렬용 절대값
                })
        
        # 기여도 절대값 순 정렬
        increase_industries = sorted([i for i in industries if i['growth_rate'] > 0], 
                                    key=lambda x: x['contribution_abs'], reverse=True)[:2]
        decrease_industries = sorted([i for i in industries if i['growth_rate'] < 0], 
                                    key=lambda x: x['contribution_abs'], reverse=True)[:2]
        
        return {"increase": increase_industries, "decrease": decrease_industries}
    
    def _extract_nationwide_industries_from_raw_data(self) -> dict:
        """기초자료 시트에서 전국 업종별 데이터 추출"""
        df = self.df_analysis
        
        # 헤더 행 찾기
        header_row_idx = None
        for i in range(min(10, len(df))):
            row = df.iloc[i]
            row_str = ' '.join([str(v) for v in row.values[:10] if pd.notna(v)])
            if '지역' in row_str and any(year in row_str for year in ['2023', '2024', '2025', '2026']):
                header_row_idx = i
                break
        
        if header_row_idx is None:
            header_row_idx = 2
        
        header_row = df.iloc[header_row_idx]
        
        # 동적으로 타겟 컬럼 인덱스 찾기
        target_col = self.find_target_col_index(header_row, self.year or 2025, self.quarter or 2)
        if target_col == -1:
            # fallback: 헤더에서 직접 찾기
            header = df.iloc[header_row_idx] if header_row_idx < len(df) else df.iloc[0]
            for col_idx in range(len(header) - 1, 4, -1):
                col_val = str(header[col_idx]) if pd.notna(header[col_idx]) else ''
                if f"{self.year or 2025}" in col_val and (f"{self.quarter or 2}/4" in col_val or f"{self.quarter or 2}분기" in col_val):
                    target_col = col_idx
                    break
            
            if target_col == -1:
                target_col = len(header) - 2
        
        # 전년동분기 인덱스 계산
        prev_y_col = target_col - 4 if target_col >= 4 else target_col - 1
        
        # 소비재 필터링 목록
        consumer_goods = ['식료품', '음료']
        
        # 업종별 데이터 추출 (분류단계 2)
        industries = []
        for i in range(header_row_idx + 1, len(df)):
            row = df.iloc[i]
            region = str(row[1]).strip() if pd.notna(row[1]) else ''
            classification = str(row[2]).strip() if pd.notna(row[2]) else ''
            if region == '전국' and classification == '2':
                current = self.safe_float(row[target_col], None) if target_col < len(row) else None
                prev = self.safe_float(row[prev_y_col], None) if prev_y_col < len(row) else None
                if current is not None and prev is not None and prev != 0:
                    ind_growth = ((current - prev) / prev) * 100
                    
                    industry_name = self._get_industry_display_name(str(row[4]) if pd.notna(row[4]) else '')
                    industry_code = str(row[3]) if 3 < len(row) and pd.notna(row[3]) else ''
                    
                    # 소비재 필터링
                    if any(consumer in industry_name for consumer in consumer_goods):
                        continue
                    
                    # 가중치가 없으므로 하드코딩 가중치 사용
                    weight = self._get_industry_weight_fallback(industry_name, industry_code)
                    
                    # 기여도 절대값 = (증감률 × 가중치)의 절대값
                    contribution_abs = abs(ind_growth * weight / 100)
                    
                    industries.append({
                        'name': industry_name,
                        'growth_rate': round(ind_growth, 1),
                        'contribution': round(ind_growth * 0.1, 6),  # 기존 호환성
                        'contribution_abs': contribution_abs  # 정렬용 절대값
                    })
        
        # 기여도 절대값 순 정렬
        increase_industries = sorted([i for i in industries if i['growth_rate'] > 0], 
                                    key=lambda x: x['contribution_abs'], reverse=True)[:2]
        decrease_industries = sorted([i for i in industries if i['growth_rate'] < 0], 
                                    key=lambda x: x['contribution_abs'], reverse=True)[:2]
        
        return {"increase": increase_industries, "decrease": decrease_industries}
    
    def _get_industry_weight_fallback(self, industry_name: str, industry_code: str = None) -> float:
        """
        가중치 컬럼을 찾지 못할 경우 사용하는 하드코딩 가중치 로직
        제조업 10대 주력 업종에 우선순위 부여
        """
        # 주력 업종 키워드 및 가중치 매핑 (상대적 가중치)
        major_industries = {
            # 전자/반도체 관련
            '전자': 100.0,
            '반도체': 100.0,
            '전자부품': 100.0,
            '컴퓨터': 80.0,
            '통신장비': 80.0,
            # 자동차 관련
            '자동차': 90.0,
            '트레일러': 90.0,
            # 기계장비 관련
            '기계': 70.0,
            '기계장비': 70.0,
            '산업용기계': 70.0,
            # 화학 관련
            '화학': 60.0,
            '화학물질': 60.0,
            '화학제품': 60.0,
            # 철강/금속 관련
            '1차금속': 50.0,
            '금속': 50.0,
            '철강': 50.0,
            # 기타 주요 업종
            '전기장비': 40.0,
            '의료': 30.0,
            '의약품': 30.0,
        }
        
        # 업종명에서 키워드 매칭
        industry_name_lower = industry_name.lower()
        for keyword, weight in major_industries.items():
            if keyword in industry_name_lower:
                return weight
        
        # 매칭되지 않으면 기본 가중치
        return 10.0
    
    def _extract_nationwide_from_analysis(self) -> dict:
        """분석 시트에서 전국 데이터 추출 (fallback)"""
        df = self.df_analysis
        data_df = df.iloc[self.DATA_START_ROW:].copy()
        
        # 헤더 행 찾기
        header_row_idx = None
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            row_str = ' '.join([str(v) for v in row.values[:30] if pd.notna(v)])
            if any(year in row_str for year in ['2023', '2024', '2025', '2026']):
                header_row_idx = i
                break
        
        if header_row_idx is None:
            header_row_idx = 2
        
        header_row = df.iloc[header_row_idx]
        
        # 동적으로 타겟 컬럼 인덱스 찾기 (증감률 컬럼)
        growth_rate_col = self.find_target_col_index(header_row, self.year or 2025, self.quarter or 2)
        if growth_rate_col == -1:
            # fallback: 헤더에서 직접 찾기
            growth_rate_col = self._find_column_by_header(df, ['증감률', 'growth', 'rate'], search_rows=5)
            if growth_rate_col is None:
                growth_rate_col = self.COL_GROWTH_RATE
                print(f"[광공업생산] ⚠️ 증감률 컬럼 fallback: {growth_rate_col}")
        
        nationwide_total = None
        try:
            filtered = data_df[(data_df[self.COL_REGION_NAME] == '전국') & 
                              (data_df[self.COL_INDUSTRY_CODE] == 'BCD')]
            if not filtered.empty:
                nationwide_total = filtered.iloc[0]
        except Exception as e:
            print(f"[광공업생산] 전국 데이터 찾기 실패: {e}")
        
        if nationwide_total is None:
            return self._get_default_nationwide_data()
        
        # 광공업생산지수 (집계 시트에서) - 동적 컬럼 사용
        df_agg = self.df_aggregation
        try:
            # 집계 시트 헤더 찾기
            agg_header_row_idx = None
            for i in range(min(5, len(df_agg))):
                row = df_agg.iloc[i]
                row_str = ' '.join([str(v) for v in row.values[:20] if pd.notna(v)])
                if any(year in row_str for year in ['2023', '2024', '2025', '2026']):
                    agg_header_row_idx = i
                    break
            
            if agg_header_row_idx is None:
                agg_header_row_idx = 2
            
            agg_header_row = df_agg.iloc[agg_header_row_idx]
            agg_target_col = self.find_target_col_index(agg_header_row, self.year or 2025, self.quarter or 2)
            if agg_target_col == -1:
                agg_target_col = self.AGG_COL_2025_2Q if hasattr(self, 'AGG_COL_2025_2Q') else (len(df_agg.columns) - 1)
            
            nationwide_agg_rows = df_agg[(df_agg[self.AGG_COL_REGION_NAME] == '전국') & 
                                         (df_agg[self.AGG_COL_INDUSTRY_CODE] == 'BCD')]
            if nationwide_agg_rows.empty:
                production_index = 100.0
            else:
                nationwide_agg = nationwide_agg_rows.iloc[0]
                production_index = self.safe_float(nationwide_agg[agg_target_col] if agg_target_col < len(nationwide_agg) else 100.0, 100.0)
        except Exception as e:
            production_index = 100.0
        
        growth_rate = nationwide_total[growth_rate_col] if growth_rate_col < len(nationwide_total) and pd.notna(nationwide_total[growth_rate_col]) else 0
        industries_data = self._extract_nationwide_industries_from_analysis()
        
        return {
            "production_index": production_index,
            "growth_rate": self.safe_round(growth_rate, 1, 0.0),
            "growth_direction": "증가" if self.safe_float(growth_rate, 0) > 0 else "감소",
            "main_increase_industries": industries_data.get("increase", []),
            "main_decrease_industries": industries_data.get("decrease", [])
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
        
        # 헤더 행 찾기
        header_row_idx = None
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            row_str = ' '.join([str(v) for v in row.values[:20] if pd.notna(v)])
            if any(year in row_str for year in ['2023', '2024', '2025', '2026']):
                header_row_idx = i
                break
        
        if header_row_idx is None:
            header_row_idx = 2
        
        header_row = df.iloc[header_row_idx]
        
        # 동적으로 타겟 컬럼 인덱스 찾기
        target_col = self.find_target_col_index(header_row, self.year or 2025, self.quarter or 2)
        if target_col == -1:
            print(f"[광공업생산] {self.year or 2025}년 {self.quarter or 2}분기 컬럼을 찾을 수 없어 기본 인덱스 사용")
            target_col = self.AGG_COL_2025_2Q if hasattr(self, 'AGG_COL_2025_2Q') else (len(df.columns) - 1)
        
        # 전년동분기 인덱스 계산 (상대적 위치 -4)
        prev_y_col = target_col - 4 if target_col >= 4 else (self.AGG_COL_2024_2Q if hasattr(self, 'AGG_COL_2024_2Q') else target_col - 1)
        
        # 당분기와 전년동분기 지수로 증감률 계산 (동적 컬럼 사용)
        current_index = self.safe_float(nationwide_total.iloc[target_col] if hasattr(nationwide_total, 'iloc') and target_col < len(nationwide_total) else (nationwide_total[target_col] if target_col < len(nationwide_total) else 100), 100)
        prev_year_index = self.safe_float(nationwide_total.iloc[prev_y_col] if hasattr(nationwide_total, 'iloc') and prev_y_col < len(nationwide_total) else (nationwide_total[prev_y_col] if prev_y_col < len(nationwide_total) else 100), 100)
        
        if prev_year_index and prev_year_index != 0:
            growth_rate = ((current_index - prev_year_index) / prev_year_index) * 100
        else:
            growth_rate = 0.0
        
        # 전국 중분류 업종별 데이터 (분류단계 2)
        nationwide_industries = df[(df[self.AGG_COL_REGION_NAME] == '전국') & 
                                   (df[self.AGG_COL_CLASSIFICATION].astype(str) == '2')]
        
        # 소비재 필터링 목록
        consumer_goods = ['식료품', '음료']
        
        industries = []
        for _, row in nationwide_industries.iterrows():
            curr = self.safe_float(row[target_col], None) if target_col < len(row) else None
            prev = self.safe_float(row[prev_y_col], None) if prev_y_col < len(row) else None
            weight = self.safe_float(row[self.AGG_COL_WEIGHT], 0)
            
            if curr is not None and prev is not None and prev != 0:
                ind_growth = ((curr - prev) / prev) * 100
                # 기여도 = (당기 - 전기) / 전국전기 * 가중치/10000
                contribution = (curr - prev) / prev_year_index * weight / 10000 * 100 if prev_year_index else 0
                
                industry_name = self._get_industry_display_name(str(row[self.AGG_COL_INDUSTRY_NAME]) if pd.notna(row[self.AGG_COL_INDUSTRY_NAME]) else '')
                industry_code = str(row[self.AGG_COL_INDUSTRY_CODE]) if self.AGG_COL_INDUSTRY_CODE < len(row) and pd.notna(row[self.AGG_COL_INDUSTRY_CODE]) else ''
                
                # 소비재 필터링
                if any(consumer in industry_name for consumer in consumer_goods):
                    continue
                
                # 가중치가 없거나 0인 경우 하드코딩 가중치 사용
                if weight == 0 or weight is None:
                    weight = self._get_industry_weight_fallback(industry_name, industry_code)
                    print(f"[광공업생산] 가중치 없음 - {industry_name}에 하드코딩 가중치 {weight} 적용")
                
                # 기여도 절대값 = (증감률 × 가중치)의 절대값
                # 가중치가 10000 단위이므로 10000으로 나눔
                contribution_abs = abs(ind_growth * weight / 10000)
                
                industries.append({
                    'name': industry_name,
                    'growth_rate': round(ind_growth, 1),
                    'contribution': round(contribution, 6),
                    'contribution_abs': contribution_abs  # 정렬용 절대값
                })
        
        # 기여도 절대값 순 정렬
        increase_industries = sorted([i for i in industries if i['growth_rate'] > 0], 
                                    key=lambda x: x['contribution_abs'], reverse=True)[:2]
        decrease_industries = sorted([i for i in industries if i['growth_rate'] < 0], 
                                    key=lambda x: x['contribution_abs'], reverse=True)[:2]
        
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
    
    def extract_regional_data(self, table_data: list = None) -> dict:
        """
        Step 2: Text Variables 추출 (Table Data 재사용)
        table_data를 증감률 기준으로 정렬하여 증가/감소 지역 추출
        """
        # table_data가 제공되지 않으면 SSOT에서 추출
        if table_data is None:
            table_data = self._extract_table_data_ssot()
        
        # 전국 제외하고 시도만 필터링
        regional_table_data = [d for d in table_data if d["region_name"] != "전국"]
        
        # 증감률 기준으로 내림차순 정렬
        sorted_regions = sorted(regional_table_data, key=lambda x: x["change_rate"], reverse=True)
        
        # 증가 지역: 양수(+)인 지역 중 상위 2~3개 (0.0은 완전히 제외)
        increase_regions_data = [r for r in sorted_regions if r.get("change_rate") is not None and r["change_rate"] > 0][:3]
        
        # 감소 지역: 음수(-)인 지역 중 하위 2~3개 (0.0은 완전히 제외)
        decrease_regions_data = sorted(
            [r for r in sorted_regions if r.get("change_rate") is not None and r["change_rate"] < 0],
            key=lambda x: x["change_rate"]
        )[:3]
        
        # 업종별 데이터는 기존 로직 유지 (엑셀에서 추출)
        individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 업종별 데이터 추출 (기존 로직)
        if hasattr(self, 'use_aggregation_only') and self.use_aggregation_only:
            industry_data = self._extract_regional_industries_from_aggregation(individual_regions)
        elif self.use_raw_data:
            industry_data = self._extract_regional_industries_from_raw_data(individual_regions)
        else:
            industry_data = self._extract_regional_industries_from_analysis(individual_regions)
        
        # table_data의 증감률과 업종 데이터 결합
        regions_data = []
        for region in individual_regions:
            # table_data에서 증감률 가져오기 (SSOT)
            region_table = next((d for d in regional_table_data if d["region_name"] == region), None)
            if region_table is None:
                continue
            
            # 업종 데이터 가져오기
            industries = industry_data.get(region, [])
            
            regions_data.append({
                "region": region,
                "growth_rate": region_table["change_rate"],  # SSOT에서 가져온 값
                "top_industries": industries[:3]  # 상위 3개만
            })
        
        # 증가/감소 지역 분류 (table_data 기준)
        # 0.0인 지역은 완전히 제외 (None도 제외)
        increase_regions = sorted(
            [r for r in regions_data if r.get("growth_rate") is not None and r["growth_rate"] > 0],
            key=lambda x: x["growth_rate"],
            reverse=True
        )
        
        decrease_regions = sorted(
            [r for r in regions_data if r.get("growth_rate") is not None and r["growth_rate"] < 0],
            key=lambda x: x["growth_rate"]
        )
        
        return {
            "increase_regions": increase_regions,
            "decrease_regions": decrease_regions,
            "region_count": len(increase_regions)
        }
    
    def _extract_regional_industries_from_analysis(self, individual_regions: list) -> dict:
        """분석 시트에서 지역별 업종 데이터 추출"""
        df = self.df_analysis
        data_df = df.iloc[self.DATA_START_ROW:].copy()
        industry_data = {}
        
        for region in individual_regions:
            try:
                region_industries = data_df[(data_df[self.COL_REGION_NAME] == region) & 
                                           (pd.notna(data_df[self.COL_CONTRIBUTION]))]
                
                top_industries = []
                industry_count = 0
                for _, row in region_industries.iterrows():
                    if industry_count >= 3:
                        break
                    if pd.notna(row[self.COL_INDUSTRY_NAME]) and str(row[self.COL_INDUSTRY_CODE]) != 'BCD':
                        top_industries.append({
                            "name": self._get_industry_display_name(str(row[self.COL_INDUSTRY_NAME])),
                            "growth_rate": self.safe_round(row[self.COL_GROWTH_RATE], 1, 0.0) if self.COL_GROWTH_RATE < len(row) else 0.0,
                            "contribution": self.safe_round(row[self.COL_CONTRIBUTION], 6, 0.0) if self.COL_CONTRIBUTION < len(row) else 0.0
                        })
                        industry_count += 1
                industry_data[region] = top_industries
            except Exception as e:
                print(f"[광공업생산] {region} 업종별 데이터 추출 실패: {e}")
                industry_data[region] = []
        
        return industry_data
    
    def _extract_regional_industries_from_aggregation(self, individual_regions: list) -> dict:
        """집계 시트에서 지역별 업종 데이터 추출"""
        df = self.df_aggregation
        industry_data = {}
        
        # 헤더 행 찾기
        header_row_idx = None
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            row_str = ' '.join([str(v) for v in row.values[:20] if pd.notna(v)])
            if any(year in row_str for year in ['2023', '2024', '2025', '2026']):
                header_row_idx = i
                break
        
        if header_row_idx is None:
            header_row_idx = 2
        
        header_row = df.iloc[header_row_idx]
        
        # 동적으로 타겟 컬럼 인덱스 찾기
        target_col = self.find_target_col_index(header_row, self.year or 2025, self.quarter or 2)
        if target_col == -1:
            print(f"[광공업생산] {self.year or 2025}년 {self.quarter or 2}분기 컬럼을 찾을 수 없어 기본 인덱스 사용")
            target_col = self.AGG_COL_2025_2Q if hasattr(self, 'AGG_COL_2025_2Q') else (len(df.columns) - 1)
        
        # 전년동분기 인덱스 계산 (상대적 위치 -4)
        prev_y_col = target_col - 4 if target_col >= 4 else (self.AGG_COL_2024_2Q if hasattr(self, 'AGG_COL_2024_2Q') else target_col - 1)
        
        # 전국 전년동분기 지수 (기여도 계산용)
        nationwide_rows = df[(df[self.AGG_COL_REGION_NAME] == '전국') & 
                            (df[self.AGG_COL_INDUSTRY_CODE] == 'BCD')]
        nationwide_prev = self.safe_float(nationwide_rows.iloc[0][prev_y_col], 100) if not nationwide_rows.empty and prev_y_col < len(nationwide_rows.iloc[0]) else 100
        
        for region in individual_regions:
            # 해당 지역 업종별 데이터 (분류단계 2)
            region_industries = df[(df[self.AGG_COL_REGION_NAME] == region) & 
                                  (df[self.AGG_COL_CLASSIFICATION].astype(str) == '2')]
            
            industries = []
            for _, row in region_industries.iterrows():
                curr = self.safe_float(row[target_col], None) if target_col < len(row) else None
                prev_ind = self.safe_float(row[prev_y_col], None) if prev_y_col < len(row) else None
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
            sorted_ind = sorted(industries, key=lambda x: x['contribution'], reverse=True)
            industry_data[region] = sorted_ind[:3]
        
        return industry_data
    
    def _extract_regional_industries_from_raw_data(self, individual_regions: list) -> dict:
        """기초자료 시트에서 지역별 업종 데이터 추출"""
        df = self.df_analysis
        industry_data = {}
        
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
        
        for region in individual_regions:
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
            
            sorted_ind = sorted(industries, key=lambda x: x['contribution'], reverse=True)
            industry_data[region] = sorted_ind[:3]
        
        return industry_data
    
    def _extract_regional_from_aggregation(self) -> dict:
        """집계 시트에서 시도별 데이터 추출"""
        df = self.df_aggregation
        
        individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 헤더 행 찾기
        header_row_idx = None
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            row_str = ' '.join([str(v) for v in row.values[:20] if pd.notna(v)])
            if any(year in row_str for year in ['2023', '2024', '2025', '2026']):
                header_row_idx = i
                break
        
        if header_row_idx is None:
            header_row_idx = 2
        
        header_row = df.iloc[header_row_idx]
        
        # 동적으로 타겟 컬럼 인덱스 찾기
        target_col = self.find_target_col_index(header_row, self.year or 2025, self.quarter or 2)
        if target_col == -1:
            print(f"[광공업생산] {self.year or 2025}년 {self.quarter or 2}분기 컬럼을 찾을 수 없어 기본 인덱스 사용")
            target_col = self.AGG_COL_2025_2Q if hasattr(self, 'AGG_COL_2025_2Q') else (len(df.columns) - 1)
        
        # 전년동분기 인덱스 계산 (상대적 위치 -4)
        prev_y_col = target_col - 4 if target_col >= 4 else (self.AGG_COL_2024_2Q if hasattr(self, 'AGG_COL_2024_2Q') else target_col - 1)
        
        # 전국 전년동분기 지수 (기여도 계산용)
        nationwide_rows = df[(df[self.AGG_COL_REGION_NAME] == '전국') & 
                            (df[self.AGG_COL_INDUSTRY_CODE] == 'BCD')]
        nationwide_prev = self.safe_float(nationwide_rows.iloc[0][prev_y_col], 100) if not nationwide_rows.empty and prev_y_col < len(nationwide_rows.iloc[0]) else 100
        
        regions_data = []
        
        for region in individual_regions:
            # 해당 지역 총지수 (BCD) - 집계 시트는 1-based를 0-based로 변환
            region_total = df[(df[self.AGG_COL_REGION_NAME] == region) & 
                             (df[self.AGG_COL_INDUSTRY_CODE] == 'BCD')]
            if region_total.empty:
                continue
            region_total = region_total.iloc[0]
            
            # 증감률 계산 (동적 컬럼 사용)
            current = self.safe_float(region_total.iloc[target_col] if hasattr(region_total, 'iloc') and target_col < len(region_total) else (region_total[target_col] if target_col < len(region_total) else 100), 100)
            prev = self.safe_float(region_total.iloc[prev_y_col] if hasattr(region_total, 'iloc') and prev_y_col < len(region_total) else (region_total[prev_y_col] if prev_y_col < len(region_total) else 100), 100)
            
            if prev and prev != 0:
                growth_rate = ((current - prev) / prev) * 100
            else:
                growth_rate = 0.0
            
            # 해당 지역 업종별 데이터 (분류단계 2)
            region_industries = df[(df[self.AGG_COL_REGION_NAME] == region) & 
                                  (df[self.AGG_COL_CLASSIFICATION].astype(str) == '2')]
            
            industries = []
            for _, row in region_industries.iterrows():
                curr = self.safe_float(row[target_col], None) if target_col < len(row) else None
                prev_ind = self.safe_float(row[prev_y_col], None) if prev_y_col < len(row) else None
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
        # 0.0인 지역은 완전히 제외 (None도 제외)
        increase_regions = sorted(
            [r for r in regions_data if r.get("growth_rate") is not None and r["growth_rate"] > 0],
            key=lambda x: x["growth_rate"],
            reverse=True
        )
        
        decrease_regions = sorted(
            [r for r in regions_data if r.get("growth_rate") is not None and r["growth_rate"] < 0],
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
        # 0.0인 지역은 완전히 제외 (None도 제외)
        increase_regions = sorted(
            [r for r in regions_data if r.get("growth_rate") is not None and r["growth_rate"] > 0],
            key=lambda x: x["growth_rate"],
            reverse=True
        )
        
        decrease_regions = sorted(
            [r for r in regions_data if r.get("growth_rate") is not None and r["growth_rate"] < 0],
            key=lambda x: x["growth_rate"]
        )
        
        return {
            "increase_regions": increase_regions,
            "decrease_regions": decrease_regions,
            "region_count": len(increase_regions)
        }
    
    def extract_summary_box(self, regional_data: dict = None, nationwide_data: dict = None) -> dict:
        """
        Step 3: 나레이션 생성 (Text Variables만 사용)
        회색 요약 박스 데이터 추출
        지그재그 화법을 사용한 대조 나레이션 생성
        
        논리 구조:
        - 전국 문장 (First Sentence): 전국 추세 설명
        - 시도 대조 문장 (Second Sentence - Zigzag): 전국 방향에 맞춰 '반대 지역 -> 같은 지역' 순서로 자동 배치
        """
        if regional_data is None:
            regional_data = self.extract_regional_data()
        if nationwide_data is None:
            nationwide_data = self.extract_nationwide_data()
        
        # 증가 지역 중 상위 3개
        top_increase = regional_data["increase_regions"][:3]
        
        main_regions = []
        for r in top_increase:
            industries = [ind["name"] for ind in r["top_industries"][:2]]
            main_regions.append({
                "region": r["region"],
                "industries": industries
            })
        
        # 지그재그 화법을 사용한 나레이션 생성
        from utils.text_utils import get_contrast_narrative
        
        # 1. 전국 데이터 및 상/하위 지역 추출
        nationwide_val = nationwide_data.get("growth_rate", 0.0)
        production_index = nationwide_data.get("production_index", 100.0)
        
        # 증가 지역 1위 (top_k=1, reverse=True)
        top_increase_regions = regional_data["increase_regions"][:1] if regional_data["increase_regions"] else []
        
        # 감소 지역 1위 (top_k=1, reverse=False)
        top_decrease_regions = regional_data["decrease_regions"][:1] if regional_data["decrease_regions"] else []
        
        # 증가 지역 데이터 포맷팅 (name, value 포함)
        inc_regions = [
            {"name": r["region"], "value": r["growth_rate"]}
            for r in top_increase_regions
        ]
        
        # 감소 지역 데이터 포맷팅 (name, value 포함)
        dec_regions = [
            {"name": r["region"], "value": r["growth_rate"]}
            for r in top_decrease_regions
        ]
        
        # 2. 전국 문장 (First Sentence) 생성
        from utils.text_utils import get_terms
        
        # 전국 증감률에 따라 적절한 업종 선택
        if nationwide_val > 0:
            # 전국이 증가하면 증가 업종 사용
            main_industries = nationwide_data.get("main_increase_industries", [])
        else:
            # 전국이 감소하면 감소 업종 사용
            main_industries = nationwide_data.get("main_decrease_industries", [])
        
        # 업종명 나열 (데이터 기반)
        # 데이터가 없으면 '주요 업종'으로 fallback 하되, 최대한 추출 시도
        top_industries = [ind["name"] for ind in main_industries[:3]] if main_industries else []
        industry_text = ", ".join(top_industries) if top_industries else "주요 업종"
        
        # 1. 서술어 결정 (수치 기반 강제, 어휘 통제 센터 사용)
        cause_verb, result_noun = get_terms('mining', nationwide_val)
        # 결과: cause_verb='늘어' 또는 '줄어', result_noun='증가' 또는 '감소'
        
        # 3. 문장 완성
        # 예: "전국 광공업생산(98.5)은 반도체, 자동차 등의 생산이 줄어 전년동분기대비 -1.5% 감소"
        from utils.text_utils import get_josa
        전국_josa = get_josa("전국", "Topic")
        sentence1 = f"전국{전국_josa} 광공업생산({production_index:.1f})은 {industry_text} 등의 생산이 {cause_verb} 전년동분기대비 {nationwide_val:.1f}% {result_noun}"
        
        # 3. 시도 대조 문장 (Second Sentence - Zigzag)
        # 전국 방향에 맞춰 '반대 지역 -> 같은 지역' 순서로 자동 배치
        # 어휘 통제 센터 사용: report_id='mining' 전달
        sentence2_base = get_contrast_narrative(nationwide_val, inc_regions, dec_regions, report_id='mining')
        
        # 업종 정보 추가 (감소 지역과 증가 지역의 주요 업종)
        # 어휘 통제 센터에서 가져온 어휘 사용
        dec_cause, dec_result = get_terms('mining', -1.0)  # 감소 지역용
        inc_cause, inc_result = get_terms('mining', 1.0)    # 증가 지역용
        
        sentence2 = sentence2_base
        if top_decrease_regions and top_decrease_regions[0].get("top_industries"):
            dec_industries = ", ".join([ind["name"] for ind in top_decrease_regions[0]["top_industries"][:2]])
            # 동적으로 생성된 어휘 사용
            dec_pattern = f"{dec_result}하였으나"
            dec_pattern2 = f"{dec_cause} {dec_result}"
            if dec_pattern in sentence2:
                sentence2 = sentence2.replace(dec_pattern, f"{dec_industries} 등의 생산이 {dec_cause} {dec_result}하였으나")
            elif dec_pattern2 in sentence2:
                sentence2 = sentence2.replace(dec_pattern2, f"{dec_industries} 등의 생산이 {dec_cause} {dec_result}")
        
        if top_increase_regions and top_increase_regions[0].get("top_industries"):
            inc_industries = ", ".join([ind["name"] for ind in top_increase_regions[0]["top_industries"][:2]])
            # 동적으로 생성된 어휘 사용
            inc_pattern = f"{inc_result}하였으나"
            inc_pattern2 = f"{inc_cause} {inc_result}"
            if inc_pattern2 in sentence2:
                sentence2 = sentence2.replace(inc_pattern2, f"{inc_industries} 등의 생산이 {inc_cause} {inc_result}")
            elif inc_pattern in sentence2:
                sentence2 = sentence2.replace(inc_pattern, f"{inc_industries} 등의 생산이 {inc_cause} {inc_result}하였으나")
        
        # 최종 본문 (전국 문장 + 시도 대조 문장)
        full_text = f"□ {sentence1}\n○ {sentence2}"
        
        return {
            "main_increase_regions": main_regions,
            "region_count": regional_data["region_count"],
            "contrast_narrative": sentence2,  # 지그재그 화법 나레이션 (시도 대조 문장만)
            "nationwide_sentence": sentence1,  # 전국 문장
            "full_narrative": full_text,  # 완전한 나레이션 (전국 + 시도)
            # 하위 호환성을 위한 기존 필드들
            "nationwide_val": nationwide_val,
            "top_increase": top_increase_regions,
            "top_decrease": top_decrease_regions
        }
    
    def extract_top3_regions(self, regional_data: dict = None) -> tuple:
        """
        Step 3: 나레이션 생성 (Text Variables만 사용)
        상위 3개 증가/감소 지역 추출 (< 주요 증감 지역 및 업종 > 섹션용)
        """
        if regional_data is None:
            regional_data = self.extract_regional_data()
        
        top3_increase = []
        for r in regional_data["increase_regions"][:3]:
            top3_increase.append({
                "region": r["region"],
                "growth_rate": r["growth_rate"],  # SSOT에서 가져온 값
                "industries": r["top_industries"][:3]
            })
        
        top3_decrease = []
        for r in regional_data["decrease_regions"][:3]:
            top3_decrease.append({
                "region": r["region"],
                "growth_rate": r["growth_rate"],  # SSOT에서 가져온 값
                "industries": r["top_industries"][:3]
            })
        
        return top3_increase, top3_decrease
    
    def _extract_table_data_ssot(self) -> list:
        """
        Step 1: Table Data 확정 (SSOT 생성)
        엑셀에서 전국 및 17개 시도의 데이터(지수, 증감률)를 모두 추출하여 table_data 리스트 생성
        반환 형식: [{"rank": int, "region_name": str, "value": float, "change_rate": float, ...}, ...]
        """
        df_agg = self.df_aggregation
        
        # Step 1: 디버깅 트랩 설치 - 헤더 정보 출력
        print(f"\n[광공업생산 SSOT] ===== 헤더 정보 분석 시작 =====")
        print(f"[광공업생산 SSOT] DataFrame shape: {df_agg.shape}")
        
        # 헤더 영역 (3~5번째 행) 출력
        for header_row_idx in range(min(5, len(df_agg))):
            header_row = df_agg.iloc[header_row_idx]
            header_values = [str(v) if pd.notna(v) else 'NaN' for v in header_row.values[:30]]
            print(f"[광공업생산 SSOT] 헤더 행 {header_row_idx}: {header_values[:15]}...")
        
        # Step 1-1: 지역명 컬럼 찾기 (로그 기반: Index 4 확정)
        # 헤더 행 2: ['조회용', '조회준비2', '조회준비1', '지역\n코드', '지역\n이름', ...]
        # 헤더 행 3: ['00BCD', 'BCD', 'BCD', '00', '전국', ...]
        # 따라서 지역명은 Index 4에 위치
        
        region_name_col = None
        region_code_col = None
        
        # 먼저 헤더에서 '지역 이름' 찾기
        for header_row_idx in range(min(5, len(df_agg))):
            header_row = df_agg.iloc[header_row_idx]
            for col_idx in range(min(10, len(header_row))):
                cell_value = str(header_row.iloc[col_idx]) if pd.notna(header_row.iloc[col_idx]) else ''
                if '지역' in cell_value and ('이름' in cell_value or 'name' in cell_value.lower()):
                    region_name_col = col_idx
                    print(f"[광공업생산 SSOT] ✅ 지역명 컬럼 발견: Index {region_name_col} (헤더 행 {header_row_idx}: '{cell_value}')")
                if '지역' in cell_value and ('코드' in cell_value or 'code' in cell_value.lower()):
                    region_code_col = col_idx
                    print(f"[광공업생산 SSOT] ✅ 지역코드 컬럼 발견: Index {region_code_col} (헤더 행 {header_row_idx}: '{cell_value}')")
        
        # 실제 데이터 행(헤더 행 3)에서 '전국'이 있는 컬럼 확인 (검증)
        if len(df_agg) > 3:
            data_row = df_agg.iloc[3]  # 헤더 행 3 (실제 데이터 시작)
            for col_idx in range(min(10, len(data_row))):
                cell_value = str(data_row.iloc[col_idx]).strip() if pd.notna(data_row.iloc[col_idx]) else ''
                if cell_value == '전국':
                    if region_name_col is None or region_name_col != col_idx:
                        print(f"[광공업생산 SSOT] ✅ 실제 데이터 행에서 '전국' 발견: Index {col_idx} (기존 추정: {region_name_col})")
                        region_name_col = col_idx
                    break
        
        # 로그 기반 확정: Index 4 (근본 수정)
        if region_name_col is None:
            region_name_col = 4  # 로그에서 확인된 정확한 인덱스
            print(f"[광공업생산 SSOT] ✅ 지역명 컬럼 Index 4로 확정 (로그 기반)")
        
        if region_code_col is None:
            region_code_col = 3  # 로그에서 확인된 정확한 인덱스
            print(f"[광공업생산 SSOT] ✅ 지역코드 컬럼 Index 3으로 확정 (로그 기반)")
        
        # Step 2: 동적 컬럼 매핑 - 헤더 검색
        if self.period_context:
            from utils.excel_utils import find_columns_by_period
            col_mapping = find_columns_by_period(df_agg, self.period_context, search_rows=5)
            target_col = col_mapping.get('target_col')
            prev_y_col = col_mapping.get('prev_y_col')
            print(f"[광공업생산 SSOT] period_context 사용: target_col={target_col}, prev_y_col={prev_y_col}")
        else:
            target_col = None
            prev_y_col = None
        
        # 헤더 파싱 실패 시 fallback: 헤더에서 직접 찾기
        if target_col is None:
            target_period = f"{self.year if self.year else 2025}.{self.quarter if self.quarter else 2}/4"
            prev_y_period = f"{self.year - 1 if self.year else 2024}.{self.quarter if self.quarter else 2}/4"
            from utils.excel_utils import find_column_by_header
            target_col = find_column_by_header(df_agg, target_period, search_rows=5)
            prev_y_col = find_column_by_header(df_agg, prev_y_period, search_rows=5)
            print(f"[광공업생산 SSOT] 헤더 직접 검색: '{target_period}' -> col={target_col}, '{prev_y_period}' -> col={prev_y_col}")
        
        # 최종 fallback: 마지막 컬럼과 그 이전 4개 컬럼 전 (하위 호환성)
        if target_col is None:
            target_col = len(df_agg.columns) - 1
            print(f"[광공업생산 SSOT] ⚠️ target_col fallback: {target_col}")
        if prev_y_col is None:
            prev_y_col = target_col - 4 if target_col >= 4 else target_col - 1
            print(f"[광공업생산 SSOT] ⚠️ prev_y_col fallback: {prev_y_col}")
        
        print(f"[광공업생산 SSOT] 최종 컬럼 매핑: region_name_col={region_name_col}, target_col={target_col}, prev_y_col={prev_y_col}")
        
        # Step 1-2: 산업코드 컬럼 동적 찾기 (헤더에서 '산업코드' 또는 '산업 코드' 찾기)
        industry_code_col = None
        for header_row_idx in range(min(5, len(df_agg))):
            header_row = df_agg.iloc[header_row_idx]
            for col_idx in range(min(10, len(header_row))):
                cell_value = str(header_row.iloc[col_idx]) if pd.notna(header_row.iloc[col_idx]) else ''
                if '산업' in cell_value and ('코드' in cell_value or 'code' in cell_value.lower()):
                    industry_code_col = col_idx
                    print(f"[광공업생산 SSOT] ✅ 산업코드 컬럼 발견: Index {industry_code_col} (헤더 행 {header_row_idx}: '{cell_value}')")
                    break
            if industry_code_col is not None:
                break
        
        # fallback: 기존 하드코딩된 인덱스 사용 (AGG_COL_INDUSTRY_CODE는 정의되지 않았으므로 region_name_col + 3 추정)
        if industry_code_col is None:
            # 일반적으로 산업코드는 지역이름 다음에 위치
            industry_code_col = region_name_col + 3 if region_name_col + 3 < len(df_agg.columns) else region_name_col + 1
            print(f"[광공업생산 SSOT] ⚠️ 산업코드 컬럼 fallback: {industry_code_col}")
        
        # 개별 시도 목록 (전국 포함)
        individual_regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        table_data = []
        
        for region in individual_regions:
            # 집계 데이터에서 해당 지역 찾기 (Index 4 사용 - 근본 수정)
            # 데이터 시작 행 이후부터 확인 (헤더 제외)
            data_df = df_agg.iloc[self.DATA_START_ROW:].copy() if self.DATA_START_ROW < len(df_agg) else df_agg
            
            # region_name_col(Index 4)로 필터링 (데이터 행만)
            if region_name_col >= len(data_df.columns):
                print(f"[광공업생산 SSOT] ⚠️ {region} - region_name_col[{region_name_col}] 범위 초과 (컬럼 수: {len(data_df.columns)})")
                continue
            
            # Index 4에서 지역명 확인 (안전한 문자열 변환)
            region_filter = data_df[
                data_df.iloc[:, region_name_col].astype(str).str.strip().str.replace('\n', '').str.replace(' ', '') == region
            ]
            
            if region_filter.empty:
                # 대체 검색: 공백 제거 없이 정확히 일치
                region_filter = data_df[
                    data_df.iloc[:, region_name_col].astype(str).str.strip() == region
                ]
            
            # BCD 코드로 추가 필터링 (산업코드 컬럼 사용)
            if industry_code_col is not None and industry_code_col < len(data_df.columns):
                # 산업코드 컬럼에서 'BCD' 찾기 (정확히 일치하거나 포함)
                region_agg = region_filter[
                    region_filter.iloc[:, industry_code_col].astype(str).str.strip().str.contains('BCD', na=False, regex=False)
                ]
            else:
                # 산업코드 컬럼이 없으면 region_name_col만으로 필터링
                # 하지만 BCD가 필요하므로 첫 번째 행만 사용 (총지수 행)
                region_agg = region_filter.head(1)
            
            if region_agg.empty:
                print(f"[광공업생산 SSOT] ⚠️ {region} 데이터 없음 (BCD 행 없음, region_name_col={region_name_col})")
                # 디버깅: 실제 데이터 행에서 Index 4 값 확인
                if len(data_df) > 0:
                    sample_values = data_df.iloc[:5, region_name_col].astype(str).tolist()
                    print(f"[광공업생산 SSOT] 디버깅: Index {region_name_col}의 처음 5개 값: {sample_values}")
                continue
            region_agg = region_agg.iloc[0]
            
            # Step 1: 디버깅 트랩 - 행 전체 데이터 출력
            row_data_list = [str(v) if pd.notna(v) else 'NaN' for v in region_agg.values[:30]]
            print(f"\n[광공업생산 SSOT] {region} 행 데이터 (처음 30개 컬럼): {row_data_list}")
            
            # 지수 추출 (동적 컬럼 사용)
            raw_current = None
            raw_prev_year = None
            
            # 자료형 안전 변환 (근본 수정): 모든 값 읽기 시 safe_float 사용
            if target_col < len(region_agg):
                raw_current = region_agg.iloc[target_col] if hasattr(region_agg, 'iloc') else region_agg[target_col]
                print(f"[광공업생산 SSOT] {region} - target_col[{target_col}] 원본값: {raw_current} (타입: {type(raw_current).__name__})")
                # parse_excel_value와 동일한 로직: safe_float 사용 (콤마, '-', None 처리)
                idx_current = self.safe_float(raw_current, None)
                print(f"[광공업생산 SSOT] {region} - target_col[{target_col}] 변환값: {idx_current}")
            else:
                print(f"[광공업생산 SSOT] ⚠️ {region} - target_col[{target_col}] 범위 초과 (행 길이: {len(region_agg)})")
                idx_current = None
            
            if prev_y_col < len(region_agg):
                raw_prev_year = region_agg.iloc[prev_y_col] if hasattr(region_agg, 'iloc') else region_agg[prev_y_col]
                print(f"[광공업생산 SSOT] {region} - prev_y_col[{prev_y_col}] 원본값: {raw_prev_year} (타입: {type(raw_prev_year).__name__})")
                # parse_excel_value와 동일한 로직: safe_float 사용
                idx_prev_year = self.safe_float(raw_prev_year, None)
                print(f"[광공업생산 SSOT] {region} - prev_y_col[{prev_y_col}] 변환값: {idx_prev_year}")
            else:
                print(f"[광공업생산 SSOT] ⚠️ {region} - prev_y_col[{prev_y_col}] 범위 초과 (행 길이: {len(region_agg)})")
                idx_prev_year = None
            
            # 결측치 체크 및 경고
            if idx_current is None:
                print(f"[광공업생산 SSOT] ❌ {region} - 현재 지수 추출 실패! 원본값: {raw_current}")
            if idx_prev_year is None:
                print(f"[광공업생산 SSOT] ❌ {region} - 전년 지수 추출 실패! 원본값: {raw_prev_year}")
            
            # 기본값 사용 금지 - 실제 데이터가 없으면 None 유지
            if idx_current is None:
                print(f"[광공업생산 SSOT] ⚠️ {region} - 현재 지수가 None이므로 이 지역을 건너뜁니다")
                continue
            if idx_prev_year is None or idx_prev_year == 0:
                print(f"[광공업생산 SSOT] ⚠️ {region} - 전년 지수가 None 또는 0이므로 증감률 계산 불가")
                change_rate = None
            else:
                # 전년동분기대비 증감률 계산
                change_rate = round(((idx_current - idx_prev_year) / idx_prev_year) * 100, 1)
            
            # 반올림 완료된 값으로 저장 (SSOT)
            table_data.append({
                "region_name": region,
                "region_display": self._get_region_display_name(region),
                "value": round(idx_current, 1),  # 현재 지수
                "prev_value": round(idx_prev_year, 1) if idx_prev_year else None,  # 전년동분기 지수
                "change_rate": change_rate,  # 증감률 (반올림 완료)
            })
            print(f"[광공업생산 SSOT] ✅ {region} - 지수: {idx_current}, 전년: {idx_prev_year}, 증감률: {change_rate}")
        
        print(f"[광공업생산 SSOT] ===== SSOT 추출 완료: {len(table_data)}개 지역 =====\n")
        
        return table_data
    
    def extract_summary_table(self, table_data: list = None) -> dict:
        """하단 표 데이터 추출 (table_data를 사용하여 일관성 유지)"""
        # table_data가 제공되지 않으면 SSOT에서 추출
        if table_data is None:
            table_data = self._extract_table_data_ssot()
        
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
        
        regions_data = []
        
        for r_info in region_order:
            region = r_info["region"]
            
            # table_data에서 해당 지역 찾기
            region_data = next((d for d in table_data if d["region_name"] == region), None)
            if region_data is None:
                # table_data에 없으면 집계 시트에서 직접 추출 (fallback)
                region_agg = df_agg[(df_agg[self.AGG_COL_REGION_NAME] == region) & 
                                   (df_agg[self.AGG_COL_INDUSTRY_CODE] == 'BCD')]
                if region_agg.empty:
                    continue
                region_agg = region_agg.iloc[0]
                
                idx_2023_2q = self.safe_float(region_agg[self.AGG_COL_2023_2Q], 100)
                idx_2024_2q = self.safe_float(region_agg[self.AGG_COL_2024_2Q], 100)
                idx_2025_1q = self.safe_float(region_agg[self.AGG_COL_2025_1Q], 100)
                idx_2025_2q = self.safe_float(region_agg[self.AGG_COL_2025_2Q], 100)
                idx_2022_2q = self.safe_float(region_agg[self.AGG_COL_2022_2Q], 100)
                idx_2024_1q = self.safe_float(region_agg[self.AGG_COL_2024_1Q], 100)
                
                def calc_growth(curr, prev):
                    if prev and prev != 0:
                        return round(((curr - prev) / prev) * 100, 1)
                    return 0.0
                
                growth_rates = [
                    calc_growth(idx_2023_2q, idx_2022_2q),
                    calc_growth(idx_2024_2q, idx_2023_2q),
                    calc_growth(idx_2025_1q, idx_2024_1q),
                    calc_growth(idx_2025_2q, idx_2024_2q),
                ]
                
                indices = [round(idx_2024_2q, 1), round(idx_2025_2q, 1)]
            else:
                # table_data 사용 (SSOT)
                # 추가 분기 데이터는 집계 시트에서 추출
                region_agg = df_agg[(df_agg[self.AGG_COL_REGION_NAME] == region) & 
                                   (df_agg[self.AGG_COL_INDUSTRY_CODE] == 'BCD')]
                if not region_agg.empty:
                    region_agg = region_agg.iloc[0]
                    idx_2023_2q = self.safe_float(region_agg[self.AGG_COL_2023_2Q], 100)
                    idx_2025_1q = self.safe_float(region_agg[self.AGG_COL_2025_1Q], 100)
                    idx_2022_2q = self.safe_float(region_agg[self.AGG_COL_2022_2Q], 100)
                    idx_2024_1q = self.safe_float(region_agg[self.AGG_COL_2024_1Q], 100)
                    
                    def calc_growth(curr, prev):
                        if prev and prev != 0:
                            return round(((curr - prev) / prev) * 100, 1)
                        return 0.0
                    
                    growth_rates = [
                        calc_growth(idx_2023_2q, idx_2022_2q),
                        calc_growth(region_data["prev_value"], idx_2023_2q),
                        calc_growth(idx_2025_1q, idx_2024_1q),
                        region_data["change_rate"],  # SSOT에서 가져온 값 사용
                    ]
                    
                    indices = [region_data["prev_value"], region_data["value"]]
                else:
                    continue
            
            row_data = {
                "region": region_data["region_display"] if region_data else self._get_region_display_name(region),
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
        """
        BaseGenerator의 추상 메서드 구현 필수
        반환값: 템플릿 렌더링에 필요한 모든 데이터를 담은 딕셔너리
        
        모든 데이터 추출 (Table First, Text Second 방식)
        
        Step 1: Table Data 확정 (SSOT 생성)
        Step 2: Text Variables 추출 (Table Data 재사용)
        Step 3: 나레이션 생성 (Text Variables만 사용)
        """
        print(f"[DEBUG] 광공업 데이터 추출 시작: {self.excel_path}")
        
        # 1. 엑셀 데이터 로드 (기존 로직 활용)
        self.load_data()
        
        # 2. 데이터 가공
        # Step 1: Table Data 확정 (SSOT 생성)
        table_data = self._extract_table_data_ssot()
        
        # Step 2: Text Variables 추출 (Table Data 재사용)
        nationwide = self.extract_nationwide_data(table_data)
        regional = self.extract_regional_data(table_data)
        
        # Step 3: 나레이션 생성 (Text Variables만 사용)
        summary_box = self.extract_summary_box(regional, nationwide)
        top3_increase, top3_decrease = self.extract_top3_regions(regional)
        summary_table = self.extract_summary_table(table_data)
        
        # 3. 필수 반환값 구성 (템플릿 변수명과 일치해야 함)
        # 하드코딩 제거: year와 quarter를 동적으로 사용
        from utils.text_utils import get_footer_source
        
        report_year = self.year if self.year else 2025
        report_quarter = self.quarter if self.quarter else 2
        
        # 푸터 정보 생성
        footer_source = get_footer_source(self.report_id)
        page_number = getattr(self, 'page_number', None) or "- 1 -"
        
        return {
            "report_info": {
                "year": report_year,
                "quarter": report_quarter,
                "data_source": "국가데이터처 국가통계포털(KOSIS), 광업제조업동향조사"
            },
            "footer_info": {
                "source": footer_source,
                "page_num": page_number
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
    
    generator = MiningManufacturingGenerator(args.excel)
    
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
    # 추상 메서드 구현 검증 테스트
    try:
        # 가짜 경로로 인스턴스화 시도
        gen = MiningManufacturingGenerator("dummy_path.xlsx")
        print("✅ 인스턴스화 성공: 추상 메서드 구현 완료")
        print(f"   - extract_all_data 메서드 타입: {type(gen.extract_all_data)}")
        print(f"   - 반환 타입 어노테이션: {gen.extract_all_data.__annotations__.get('return', 'N/A')}")
    except TypeError as e:
        print(f"❌ 인스턴스화 실패: {e}")
        import sys
        sys.exit(1)
    
    # main() 함수는 실제 사용 시에만 실행
    import sys
    if len(sys.argv) > 1:
        main()

# 하위 호환성을 위한 별칭
광공업생산Generator = MiningManufacturingGenerator

