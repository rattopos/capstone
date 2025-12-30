#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
광공업생산 보고서 생성기

엑셀 데이터를 읽어 스키마에 맞게 데이터를 추출하고,
Jinja2 템플릿을 사용하여 HTML 보고서를 생성합니다.
"""

import pandas as pd
import json
from jinja2 import Environment, FileSystemLoader
from pathlib import Path


def safe_float(value, default=None):
    """안전한 float 변환 함수 (NaN, '-' 체크 포함)"""
    if value is None:
        return default
    try:
        if pd.isna(value):
            return default
        if isinstance(value, str):
            value = value.strip()
            if value == '-' or value == '':
                return default
        result = float(value)
        if pd.isna(result):
            return default
        return result
    except (ValueError, TypeError):
        return default


def safe_round(value, decimals=1, default=None):
    """안전한 반올림 함수"""
    result = safe_float(value, default)
    if result is None:
        return default
    return round(result, decimals)


def find_sheet_with_fallback(xl, primary_sheets, fallback_sheets):
    """시트 찾기 - 기본 시트가 없으면 대체 시트 사용"""
    for sheet in primary_sheets:
        if sheet in xl.sheet_names:
            return sheet, False
    for sheet in fallback_sheets:
        if sheet in xl.sheet_names:
            print(f"[시트 대체] '{primary_sheets[0]}' → '{sheet}' (기초자료)")
            return sheet, True
    return None, False


class 광공업생산Generator:
    """광공업생산 보고서 생성 클래스"""
    
    # 업종명 매핑 사전 (엑셀 데이터 → 보고서 표기명)
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
    
    def __init__(self, excel_path: str):
        """
        초기화
        
        Args:
            excel_path: 엑셀 파일 경로
        """
        self.excel_path = excel_path
        self.df_analysis = None
        self.df_aggregation = None
        self.data = {}
        self.use_raw_data = False  # 기초자료 시트 사용 여부
        
    def load_data(self):
        """엑셀 데이터 로드"""
        xl = pd.ExcelFile(self.excel_path)
        
        # 분석 시트 찾기
        analysis_sheet, self.use_raw_data = find_sheet_with_fallback(
            xl, 
            ['A 분석', 'A분석'],
            ['광공업생산', '광공업생산지수']
        )
        
        if analysis_sheet:
            self.df_analysis = pd.read_excel(xl, sheet_name=analysis_sheet, header=None)
        else:
            raise ValueError(f"광공업생산 분석 시트를 찾을 수 없습니다. 시트: {xl.sheet_names}")
        
        # 집계 시트 찾기 (선택사항)
        agg_sheet, _ = find_sheet_with_fallback(
            xl,
            ['A(광공업생산)집계', 'A 집계'],
            ['광공업생산', '광공업생산지수']
        )
        
        if agg_sheet and agg_sheet != analysis_sheet:
            self.df_aggregation = pd.read_excel(xl, sheet_name=agg_sheet, header=None)
        else:
            self.df_aggregation = self.df_analysis.copy()
        
    def _get_industry_display_name(self, raw_name: str) -> str:
        """업종명을 보고서 표기명으로 변환"""
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
        df = self.df_analysis
        
        if self.use_raw_data:
            return self._extract_nationwide_from_raw_data()
        
        # 전국 총지수 행
        nationwide_total = df[(df[3] == '전국') & (df[6] == 'BCD')].iloc[0]
        
        # 전국 중분류 데이터 (분류단계 2)
        nationwide_industries = df[(df[3] == '전국') & (df[4].astype(str) == '2') & (pd.notna(df[28]))]
        
        # 기여도 순 정렬
        sorted_industries = nationwide_industries.sort_values(28, ascending=False)
        
        # 증가 업종 (기여도 양수)
        increase_industries = sorted_industries[sorted_industries[28] > 0]
        
        # 감소 업종 (기여도 음수)
        decrease_industries = sorted_industries[sorted_industries[28] < 0].sort_values(28, ascending=True)
        
        # 광공업생산지수 (집계 시트에서)
        df_agg = self.df_aggregation
        nationwide_agg = df_agg[(df_agg[4] == '전국') & (df_agg[7] == 'BCD')].iloc[0]
        production_index = nationwide_agg[26]  # 2025.2/4p 컬럼
        
        growth_rate = nationwide_total[21]  # 2025 2/4 증감률
        
        return {
            "production_index": safe_float(production_index, 100.0),
            "growth_rate": safe_round(growth_rate, 1, 0.0),
            "growth_direction": "증가" if safe_float(growth_rate, 0) > 0 else "감소",
            "main_increase_industries": [
                {
                    "name": self._get_industry_display_name(str(row[7])),
                    "growth_rate": safe_round(row[21], 1, 0.0),
                    "contribution": safe_round(row[28], 6, 0.0)
                }
                for _, row in increase_industries.head(5).iterrows()
            ],
            "main_decrease_industries": [
                {
                    "name": self._get_industry_display_name(str(row[7])),
                    "growth_rate": safe_round(row[21], 1, 0.0),
                    "contribution": safe_round(row[28], 6, 0.0)
                }
                for _, row in decrease_industries.head(5).iterrows()
            ]
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
        current_val = safe_float(nationwide_row[current_quarter_col], 100)
        prev_val = safe_float(nationwide_row[prev_year_col], 100)
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
                current = safe_float(row[current_quarter_col], None)
                prev = safe_float(row[prev_year_col], None)
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
        df = self.df_analysis
        
        if self.use_raw_data:
            return self._extract_regional_from_raw_data()
        
        # 개별 시도 목록 (수도, 충청 등 권역 제외)
        individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        regions_data = []
        
        for region in individual_regions:
            # 해당 지역 총지수
            region_total = df[(df[3] == region) & (df[6] == 'BCD')]
            if region_total.empty:
                continue
            region_total = region_total.iloc[0]
            
            growth_rate = safe_float(region_total[21], 0)
            
            # 해당 지역 업종별 데이터
            region_industries = df[(df[3] == region) & (pd.notna(df[28]))]
            
            # 기여도 순 정렬 (증가는 높은 순, 감소는 낮은 순)
            if growth_rate >= 0:
                sorted_ind = region_industries.sort_values(28, ascending=False)
            else:
                sorted_ind = region_industries.sort_values(28, ascending=True)
            
            # 상위 3개 업종 (BCD 제외)
            top_industries = []
            industry_count = 0
            for _, row in sorted_ind.iterrows():
                if industry_count >= 3:
                    break
                if pd.notna(row[7]) and str(row[6]) != 'BCD':
                    top_industries.append({
                        "name": self._get_industry_display_name(str(row[7])),
                        "growth_rate": safe_round(row[21], 1, 0.0),
                        "contribution": safe_round(row[28], 6, 0.0)
                    })
                    industry_count += 1
            
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
            current_val = safe_float(region_row[current_quarter_col], None)
            prev_val = safe_float(region_row[prev_year_col], None)
            
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
                    current = safe_float(row[current_quarter_col], None)
                    prev = safe_float(row[prev_year_col], None)
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
        df_analysis = self.df_analysis
        
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
            
            # 집계 데이터에서 해당 지역 찾기
            region_agg = df_agg[(df_agg[4] == region) & (df_agg[7] == 'BCD')]
            if region_agg.empty:
                continue
            region_agg = region_agg.iloc[0]
            
            # 분석 데이터에서 증감률 찾기
            region_analysis = df_analysis[(df_analysis[3] == region) & (df_analysis[6] == 'BCD')]
            if region_analysis.empty:
                continue
            region_analysis = region_analysis.iloc[0]
            
            # 증감률 (A 분석 시트 컬럼)
            growth_rates = [
                round(float(region_analysis[13]) if pd.notna(region_analysis[13]) else 0, 1),  # 2023 2/4
                round(float(region_analysis[17]) if pd.notna(region_analysis[17]) else 0, 1),  # 2024 2/4
                round(float(region_analysis[20]) if pd.notna(region_analysis[20]) else 0, 1),  # 2025 1/4
                round(float(region_analysis[21]) if pd.notna(region_analysis[21]) else 0, 1),  # 2025 2/4
            ]
            
            # 지수 (집계 시트 컬럼)
            indices = [
                round(float(region_agg[22]) if pd.notna(region_agg[22]) else 0, 1),  # 2024 2/4
                round(float(region_agg[26]) if pd.notna(region_agg[26]) else 0, 1),  # 2025 2/4
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
        """HTML 보고서 렌더링"""
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
            print(f"보고서가 생성되었습니다: {output_path}")
        
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
    
    parser = argparse.ArgumentParser(description='광공업생산 보고서 생성기')
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

