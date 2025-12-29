#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
통계표 생성기

분석표 엑셀에서 데이터를 추출하여 통계표 HTML을 생성합니다.
스키마를 통해 분석표의 구조 변경에 유연하게 대응할 수 있습니다.
"""

import pandas as pd
import json
import re
from pathlib import Path
from jinja2 import Environment, FileSystemLoader
from typing import Dict, List, Optional, Any

# RawDataExtractor 임포트
try:
    from raw_data_extractor import RawDataExtractor
except ImportError:
    RawDataExtractor = None


class StatisticsTableGenerator:
    """통계표 생성 클래스"""
    
    # 통계표 항목 정의 (집계 시트 기준, 전년동기비 계산 필요)
    TABLE_CONFIG = {
        "광공업생산지수": {
            "집계_시트": "A(광공업생산)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 4, "값": "BCD"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 9, "2018": 9, "2019": 9, "2020": 9, "2021": 10, "2022": 11, "2023": 12, "2024": 13},
            "분기_시작_컬럼": 14,  # 2022 2/4부터
            "계산방식": "growth_rate"
        },
        "서비스업생산지수": {
            "집계_시트": "B(서비스업생산)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 4, "값": "E~S"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 8, "2018": 8, "2019": 8, "2020": 8, "2021": 9, "2022": 10, "2023": 11, "2024": 12},
            "분기_시작_컬럼": 13,
            "계산방식": "growth_rate"
        },
        "소매판매액지수": {
            "집계_시트": "C(소비)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 5, "값": "A0"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 7, "2018": 7, "2019": 7, "2020": 7, "2021": 8, "2022": 9, "2023": 10, "2024": 11},
            "분기_시작_컬럼": 12,
            "계산방식": "growth_rate"
        },
        "건설수주액": {
            "집계_시트": "F'(건설)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 2, "값": "0"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 5, "2018": 5, "2019": 5, "2020": 5, "2021": 6, "2022": 7, "2023": 8, "2024": 9},
            "분기_시작_컬럼": 10,
            "계산방식": "growth_rate"
        },
        "고용률": {
            "집계_시트": "D(고용률)집계",
            "단위": "[전년동기비, %p]",
            "총지수_식별": {"컬럼": 3, "값": "계"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 7, "2018": 7, "2019": 7, "2020": 7, "2021": 8, "2022": 9, "2023": 10, "2024": 11},
            "분기_시작_컬럼": 12,
            "계산방식": "difference"  # %p 단위
        },
        "실업률": {
            "집계_시트": "D(실업)집계",
            "단위": "[전년동기비, %p]",
            "총지수_식별": {"컬럼": 1, "값": "계"},
            "지역_컬럼": 0,
            "분류단계_컬럼": 1,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 7, "2018": 7, "2019": 7, "2020": 7, "2021": 8, "2022": 9, "2023": 10, "2024": 11},
            "분기_시작_컬럼": 12,
            "계산방식": "difference_rate",
            "지역_매핑": True  # 서울특별시->서울 등 매핑 필요
        },
        "국내인구이동": {
            "집계_시트": "I(순인구이동)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 2, "값": "순인구이동 수"},  # col2가 분류
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 7, "2018": 7, "2019": 7, "2020": 7, "2021": 8, "2022": 9, "2023": 10, "2024": 11},
            "분기_시작_컬럼": 12,
            "계산방식": "growth_rate"
        },
        "수출액": {
            "집계_시트": "G(수출)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 2, "값": "0"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 9, "2018": 9, "2019": 9, "2020": 9, "2021": 10, "2022": 11, "2023": 12, "2024": 13},
            "분기_시작_컬럼": 14,
            "계산방식": "growth_rate"
        },
        "수입액": {
            "집계_시트": "H(수입)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 2, "값": "0"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 9, "2018": 9, "2019": 9, "2020": 9, "2021": 10, "2022": 11, "2023": 12, "2024": 13},
            "분기_시작_컬럼": 14,
            "계산방식": "growth_rate"
        },
        "소비자물가지수": {
            "집계_시트": "E(품목성질물가)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 1, "값": 0},
            "지역_컬럼": 0,
            "분류단계_컬럼": 1,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 7, "2018": 7, "2019": 7, "2020": 7, "2021": 8, "2022": 9, "2023": 10, "2024": 11},
            "분기_시작_컬럼": 12,
            "계산방식": "growth_rate"
        }
    }
    
    # 기초자료 시트 매핑
    RAW_SHEET_MAPPING = {
        "광공업생산지수": "광공업생산",
        "서비스업생산지수": "서비스업생산",
        "소매판매액지수": "소비(소매, 추가)",
        "건설수주액": "건설 (공표자료)",
        "고용률": "고용률",
        "실업률": "실업자 수",
        "수출액": "수출",
        "수입액": "수입",
        "국내인구이동": "시도 간 이동"
    }
    
    # 기초자료 시트별 컬럼 매핑 (분석표와 다른 구조 대응)
    # 각 시트에서 지역이름과 분류단계의 컬럼 인덱스
    RAW_COLUMN_MAPPING = {
        "광공업생산": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "서비스업생산": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "소비(소매, 추가)": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "건설 (공표자료)": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "고용률": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "difference"},  # %p 단위
        "실업자 수": {"지역_컬럼": 0, "분류단계_컬럼": 1, "분류값": "계", "계산방식": "growth_rate"},
        "수출": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "수입": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "시도 간 이동": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "순인구이동 수", "계산방식": "growth_rate"},
    }
    
    # 지역 목록 (페이지별)
    PAGE1_REGIONS = ["전국", "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종"]
    PAGE2_REGIONS = ["경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"]
    ALL_REGIONS = PAGE1_REGIONS + PAGE2_REGIONS
    
    def __init__(self, excel_path: str, historical_data_path: Optional[str] = None, 
                 raw_excel_path: Optional[str] = None, current_year: int = 2025, current_quarter: int = 2):
        """
        초기화
        
        Args:
            excel_path: 분석표 엑셀 파일 경로
            historical_data_path: 과거 데이터 JSON 파일 경로 (사용하지 않음, 호환성 유지)
            raw_excel_path: 기초자료 엑셀 파일 경로 (우선 사용)
            current_year: 현재 연도
            current_quarter: 현재 분기
        """
        self.excel_path = excel_path
        self.raw_excel_path = raw_excel_path
        self.current_year = current_year
        self.current_quarter = current_quarter
        self.historical_data_path = historical_data_path
        self.historical_data = {}
        self.cached_sheets = {}
        
        # RawDataExtractor 초기화
        self.raw_extractor = None
        if raw_excel_path and RawDataExtractor and Path(raw_excel_path).exists():
            try:
                self.raw_extractor = RawDataExtractor(raw_excel_path, current_year, current_quarter)
                print(f"[통계표] 기초자료 추출기 초기화 완료: {raw_excel_path}")
            except Exception as e:
                print(f"[통계표] 기초자료 추출기 초기화 실패: {e}")
        
        # 과거 데이터 로드 (하위 호환성, 사용하지 않음)
        if historical_data_path and Path(historical_data_path).exists():
            with open(historical_data_path, 'r', encoding='utf-8') as f:
                self.historical_data = json.load(f)
    
    def _load_sheet(self, sheet_name: str) -> pd.DataFrame:
        """엑셀 시트 로드 (캐싱)"""
        if sheet_name not in self.cached_sheets:
            try:
                self.cached_sheets[sheet_name] = pd.read_excel(
                    self.excel_path,
                    sheet_name=sheet_name,
                    header=None
                )
            except Exception as e:
                print(f"시트 로드 실패: {sheet_name} - {e}")
                return None
        return self.cached_sheets[sheet_name]
    
    def _extract_from_raw_data(self, raw_sheet_name: str, config: dict) -> Optional[Dict[str, Any]]:
        """기초자료에서 데이터 추출
        
        Args:
            raw_sheet_name: 기초자료 시트 이름
            config: 통계표 설정
            
        Returns:
            {
                'yearly': {'2016': {'전국': value, ...}, ...},
                'quarterly': {'2016.1/4': {'전국': value, ...}, ...}
            }
        """
        if not self.raw_extractor:
            return None
        
        # 기초자료 시트별 컬럼 매핑 사용 (분석표 설정이 아닌 기초자료 구조에 맞춤)
        raw_col_config = self.RAW_COLUMN_MAPPING.get(raw_sheet_name, {})
        
        region_column = raw_col_config.get("지역_컬럼", 1)
        classification_column = raw_col_config.get("분류단계_컬럼", 2)
        classification_value = raw_col_config.get("분류값", "0")
        
        calculation_method = raw_col_config.get("계산방식", "growth_rate")
        
        print(f"[통계표] 기초자료 추출 - 시트: {raw_sheet_name}, 지역컬럼: {region_column}, 분류컬럼: {classification_column}, 분류값: {classification_value}, 계산방식: {calculation_method}")
        
        # 계산방식에 따라 다른 함수 호출
        if calculation_method == "difference":
            # 차이 계산 (%p 단위)
            yearly_data = self.raw_extractor.extract_yearly_difference(
                raw_sheet_name,
                start_year=2016,
                region_column=region_column,
                classification_column=classification_column,
                classification_value=classification_value if classification_value else None
            )
            quarterly_data = self.raw_extractor.extract_quarterly_difference(
                raw_sheet_name,
                start_year=2016,
                start_quarter=1,
                region_column=region_column,
                classification_column=classification_column,
                classification_value=classification_value if classification_value else None
            )
        else:
            # 전년동기비 계산 (% 단위)
            yearly_data = self.raw_extractor.extract_yearly_growth_rate(
                raw_sheet_name,
                start_year=2016,
                region_column=region_column,
                classification_column=classification_column,
                classification_value=classification_value if classification_value else None
            )
            quarterly_data = self.raw_extractor.extract_quarterly_growth_rate(
                raw_sheet_name,
                start_year=2016,
                start_quarter=1,
                region_column=region_column,
                classification_column=classification_column,
                classification_value=classification_value if classification_value else None
            )
        
        print(f"[통계표] 기초자료 추출 결과 - 연도: {len(yearly_data)}개, 분기: {len(quarterly_data)}개")
        
        # 데이터 형식 변환 (분기 키 형식 통일: "2016 1/4" -> "2016.1/4")
        quarterly_formatted = {}
        for quarter_key, data in quarterly_data.items():
            # "2016 1/4" -> "2016.1/4" 형식으로 변환
            formatted_key = quarter_key.replace(" ", ".")
            quarterly_formatted[formatted_key] = data
        
        return {
            'yearly': yearly_data,
            'quarterly': quarterly_formatted
        }
    
    # 지역명 매핑 (실업률 시트용)
    REGION_NAME_MAPPING = {
        "전국": ["전국"],
        "서울": ["서울", "서울특별시"],
        "부산": ["부산", "부산광역시"],
        "대구": ["대구", "대구광역시"],
        "인천": ["인천", "인천광역시"],
        "광주": ["광주", "광주광역시"],
        "대전": ["대전", "대전광역시"],
        "울산": ["울산", "울산광역시"],
        "세종": ["세종", "세종특별자치시"],
        "경기": ["경기", "경기도"],
        "강원": ["강원", "강원도", "강원특별자치도"],
        "충북": ["충북", "충청북도"],
        "충남": ["충남", "충청남도"],
        "전북": ["전북", "전라북도", "전북특별자치도"],
        "전남": ["전남", "전라남도"],
        "경북": ["경북", "경상북도"],
        "경남": ["경남", "경상남도"],
        "제주": ["제주", "제주도", "제주특별자치도"]
    }
    
    def _get_total_row(self, df: pd.DataFrame, region: str, config: dict) -> Optional[pd.Series]:
        """특정 지역의 총지수 행 가져오기 (시트별 설정 사용)"""
        region_col = config["지역_컬럼"]
        id_col = config["총지수_식별"]["컬럼"]
        id_val = config["총지수_식별"]["값"]
        use_region_mapping = config.get("지역_매핑", False)
        
        # 값 타입에 따라 비교
        try:
            df_region = df[region_col].astype(str).str.strip()
            region_clean = str(region).strip()
            
            # 지역명 매핑 사용 여부
            if use_region_mapping:
                region_variants = self.REGION_NAME_MAPPING.get(region_clean, [region_clean])
                mask = df_region.isin(region_variants) & (df[id_col].astype(str).str.strip() == str(id_val).strip())
            else:
                mask = (df_region == region_clean) & (df[id_col].astype(str).str.strip() == str(id_val).strip())
            
            result = df[mask]
            
            if not result.empty:
                return result.iloc[0]
        except Exception as e:
            print(f"행 검색 오류: {region} - {e}")
        
        return None
    
    def _calculate_yoy_growth(self, current: float, previous: float) -> Optional[float]:
        """전년동기비 증감률 계산"""
        if pd.isna(current) or pd.isna(previous) or previous == 0:
            return None
        try:
            return ((float(current) - float(previous)) / float(previous)) * 100
        except (ValueError, TypeError):
            return None
    
    def _calculate_difference(self, current: float, previous: float) -> Optional[float]:
        """전년동기비 차이 계산 (%p)"""
        if pd.isna(current) or pd.isna(previous):
            return None
        try:
            return float(current) - float(previous)
        except (ValueError, TypeError):
            return None
    
    def _format_value(self, value, decimals: int = 1) -> str:
        """값 포맷팅"""
        if pd.isna(value):
            return "-"
        try:
            val = float(value)
            return f"{val:.{decimals}f}"
        except (ValueError, TypeError):
            return str(value)
    
    def _create_empty_table_data(self) -> Dict[str, Any]:
        """모든 연도/분기/지역에 기본값 '-'가 채워진 데이터 구조 생성 (2016년부터)"""
        yearly_years = ["2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"]
        
        # 2016년 1분기부터 현재 분기까지 생성
        quarterly_keys = []
        for year in range(2016, self.current_year + 1):
            start_q = 1 if year > 2016 else 1
            end_q = self.current_quarter if year == self.current_year else 4
            for quarter in range(start_q, end_q + 1):
                if year == self.current_year and quarter == self.current_quarter:
                    quarterly_keys.append(f"{year}.{quarter}/4p")
                else:
                    quarterly_keys.append(f"{year}.{quarter}/4")
        
        # 모든 연도에 대해 모든 지역의 기본값 '-' 설정
        yearly = {}
        for year in yearly_years:
            yearly[year] = {region: "-" for region in self.ALL_REGIONS}
        
        # 모든 분기에 대해 모든 지역의 기본값 '-' 설정
        quarterly = {}
        for quarter in quarterly_keys:
            quarterly[quarter] = {region: "-" for region in self.ALL_REGIONS}
        
        return {
            "yearly": yearly,
            "quarterly": quarterly,
            "yearly_years": yearly_years,
            "quarterly_keys": quarterly_keys
        }
    
    def extract_table_data(self, table_name: str) -> Dict[str, Any]:
        """특정 통계표의 데이터 추출 (집계 시트에서 전년동기비 계산)"""
        config = self.TABLE_CONFIG.get(table_name)
        if not config:
            raise ValueError(f"알 수 없는 통계표: {table_name}")
        
        # 데이터 구조 초기화 - 모든 연도/분기/지역에 기본값 '-' 설정
        data = self._create_empty_table_data()
        
        # 기초자료에서 직접 추출 (우선순위 1)
        if self.raw_extractor:
            raw_sheet_name = self.RAW_SHEET_MAPPING.get(table_name)
            if raw_sheet_name:
                try:
                    print(f"[통계표] 기초자료에서 추출: {table_name} (시트: {raw_sheet_name})")
                    raw_data = self._extract_from_raw_data(raw_sheet_name, config)
                    if raw_data:
                        # 기초자료에서 추출한 데이터로 업데이트
                        for year in raw_data.get('yearly', {}):
                            if year in data['yearly']:
                                data['yearly'][year].update(raw_data['yearly'][year])
                        
                        for quarter in raw_data.get('quarterly', {}):
                            if quarter in data['quarterly']:
                                data['quarterly'][quarter].update(raw_data['quarterly'][quarter])
                        
                        print(f"[통계표] 기초자료 추출 완료: {table_name}")
                        return data
                except Exception as e:
                    print(f"[통계표] 기초자료 추출 실패: {table_name} - {e}")
                    import traceback
                    traceback.print_exc()
        
        # 집계 시트에서 추출 (fallback)
        sheet_name = config.get("집계_시트")
        if not sheet_name:
            print(f"[통계표] 집계_시트 설정 없음: {table_name}")
            return data
            
        df = self._load_sheet(sheet_name)
        if df is None:
            print(f"[통계표] 시트를 찾을 수 없음: {sheet_name}")
            return data
        
        print(f"[통계표] 집계 시트에서 추출: {table_name} (시트: {sheet_name})")
        
        # 헤더 행에서 분기 컬럼 위치 파악
        header_row = df.iloc[2]
        quarter_col_map = {}
        for col_idx, header in enumerate(header_row):
            if pd.notna(header):
                header_str = str(header).strip()
                # "2022  2/4" 형식 파싱
                match = re.match(r'(\d{4})\s+(\d)/4', header_str)
                if match:
                    year = match.group(1)
                    q = match.group(2)
                    quarter_key = f"{year}.{q}/4"
                    quarter_col_map[quarter_key] = col_idx
        
        calculation_method = config.get("계산방식", "growth_rate")
        
        # 각 지역에 대해 데이터 추출
        for region in self.ALL_REGIONS:
            row = self._get_total_row(df, region, config)
            if row is None:
                continue
            
            # 연도별 데이터 (전년동기비 계산)
            for year in data["yearly_years"]:
                year_int = int(year)
                prev_year = str(year_int - 1)
                
                # 현재 연도와 전년도 컬럼 인덱스 찾기
                curr_col = config["연도_컬럼"].get(year)
                prev_col = config["연도_컬럼"].get(prev_year)
                
                if curr_col is not None and curr_col < len(row):
                    curr_val = row.iloc[curr_col]
                    
                    if prev_col is not None and prev_col < len(row):
                        prev_val = row.iloc[prev_col]
                        
                        if calculation_method == "difference":
                            result = self._calculate_difference(curr_val, prev_val)
                        else:
                            result = self._calculate_yoy_growth(curr_val, prev_val)
                        
                        if result is not None:
                            data["yearly"][year][region] = round(result, 1)
            
            # 분기별 데이터 (전년동기비 계산)
            for quarter_key in data["quarterly_keys"]:
                # "2024.2/4" -> 2024, 2 파싱
                match = re.match(r'(\d{4})\.(\d)/4(p?)', quarter_key)
                if not match:
                    continue
                
                year = int(match.group(1))
                q = int(match.group(2))
                prev_year_quarter_key = f"{year - 1}.{q}/4"
                
                # 현재 분기와 전년동분기 컬럼 찾기
                curr_col = quarter_col_map.get(quarter_key.rstrip('p'))
                prev_col = quarter_col_map.get(prev_year_quarter_key)
                
                if curr_col is not None and curr_col < len(row):
                    curr_val = row.iloc[curr_col]
                    
                    if prev_col is not None and prev_col < len(row):
                        prev_val = row.iloc[prev_col]
                        
                        if calculation_method == "difference":
                            result = self._calculate_difference(curr_val, prev_val)
                        else:
                            result = self._calculate_yoy_growth(curr_val, prev_val)
                        
                        if result is not None:
                            data["quarterly"][quarter_key][region] = round(result, 1)
        
        # quarterly_keys 정리 (데이터가 있는 분기만 유지)
        data["quarterly_keys"] = [
            q for q in data["quarterly_keys"] 
            if any(v != "-" for v in data["quarterly"].get(q, {}).values())
        ]
        
        return data
    
    def extract_all_tables(self, year: Optional[int] = None, quarter: Optional[int] = None) -> Dict[str, Any]:
        """모든 통계표 데이터 추출"""
        if year is None:
            year = self.current_year
        if quarter is None:
            quarter = self.current_quarter
        tables = []
        # 통계표 순서: 광공업생산지수-서비스업생산지수-소매판매액지수-건설수주액-고용률-실업률-국내인구이동-수출액-수입액-소비자물가지수
        table_order = [
            "광공업생산지수",
            "서비스업생산지수",
            "소매판매액지수",
            "건설수주액",
            "고용률",
            "실업률",
            "국내인구이동",
            "수출액",
            "수입액",
            "소비자물가지수"
        ]
        
        page_num = 22  # 시작 페이지 번호
        
        for idx, table_name in enumerate(table_order, 1):
            config = self.TABLE_CONFIG[table_name]
            data = self.extract_table_data(table_name)
            
            if data:
                tables.append({
                    "id": idx,
                    "title": table_name,
                    "unit": config["단위"],
                    "data": data,
                    "page_number_1": page_num,
                    "page_number_2": page_num + 1
                })
                page_num += 2
        
        # 부록 용어 정의
        appendix = {
            "terms": [
                {
                    "term": "불변지수",
                    "definition": "불변지수는 가격 변동분이 제외된 수량 변동분만 포함되어 있음을 의미하며, 성장 수준 분석(전년동분기비)에 활용됨"
                },
                {
                    "term": "광공업생산지수",
                    "definition": "한국표준산업분류 상의 3개 대분류(B, C, D)를 대상으로 광업제조업동향조사의 월별 품목별 생산·출하(내수 및 수출)·재고 및 생산능력·가동률지수를 기초로 작성됨"
                },
                {
                    "term": "서비스업생산지수",
                    "definition": "한국표준산업분류 상의 13개 대분류(E, G, H, I, J, K, L, M, N, P, Q, R, S)를 대상으로 서비스업동향조사의 월별 매출액을 기초로 작성됨"
                },
                {
                    "term": "소매판매액지수",
                    "definition": "한국표준산업분류 상의 '자동차 판매업 중 승용차'와 '소매업'을 대상으로 서비스업동향조사의 월별 상품판매액을 기초로 작성됨"
                },
                {
                    "term": "건설수주",
                    "definition": "종합건설업 등록업체 중 전전년 「건설업조사」 결과를 기준으로 기성액 순위 상위 기업체(대표도: 54%)의 국내공사에 대한 건설수주액임"
                },
                {
                    "term": "소비자물가지수",
                    "definition": "가구에서 일상생활을 영위하기 위해 구입하는 상품과 서비스의 평균적인 가격변동을 측정한 지수임"
                },
                {
                    "term": "지역내총생산",
                    "definition": "일정 기간 동안에 일정 지역 내에서 새로이 창출된 최종생산물을 시장가격으로 평가한 가치의 합임"
                }
            ]
        }
        
        # GRDP 데이터 (현재 N/A로 처리)
        grdp_data = self._create_grdp_placeholder()
        
        return {
            "report_info": {
                "year": year,
                "quarter": quarter
            },
            "tables": tables,
            "grdp": grdp_data,
            "appendix": appendix,
            "page_numbers": {
                "toc": 21,
                "appendix_1": page_num + 2,  # GRDP 2페이지 후
                "appendix_2": page_num + 3
            }
        }
    
    def _create_grdp_placeholder(self) -> Dict[str, Any]:
        """GRDP 데이터 생성 - grdp_extracted.json에서 데이터 로드 시도"""
        import json
        
        yearly_years = ["2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"]
        
        # 2017년 3/4분기부터 현재 분기까지 생성 (정답 이미지 기준)
        quarterly_keys = []
        for year in range(2017, self.current_year + 1):
            start_q = 3 if year == 2017 else 1
            end_q = self.current_quarter if year == self.current_year else 4
            for quarter in range(start_q, end_q + 1):
                if year == self.current_year and quarter == self.current_quarter:
                    quarterly_keys.append(f"{year}.{quarter}/4p")
                else:
                    quarterly_keys.append(f"{year}.{quarter}/4")
        
        # 기본값으로 플레이스홀더 (편집 가능) 생성
        yearly = {}
        for year in yearly_years:
            yearly[year] = {region: "-" for region in self.ALL_REGIONS}
        
        quarterly = {}
        for qk in quarterly_keys:
            quarterly[qk] = {region: "-" for region in self.ALL_REGIONS}
        
        # grdp_extracted.json에서 현재 분기 데이터 로드 시도
        try:
            grdp_json_path = Path(__file__).parent / 'grdp_extracted.json'
            if grdp_json_path.exists():
                with open(grdp_json_path, 'r', encoding='utf-8') as f:
                    grdp_data = json.load(f)
                
                # 현재 분기 키
                current_key = f"{self.current_year}.{self.current_quarter}/4p"
                
                if current_key in quarterly and 'regional_data' in grdp_data:
                    for item in grdp_data['regional_data']:
                        region = item.get('region', '')
                        growth_rate = item.get('growth_rate', 0)
                        if region in self.ALL_REGIONS and not item.get('placeholder', True):
                            quarterly[current_key][region] = round(growth_rate, 1)
                    
                    print(f"[통계표] GRDP JSON에서 {current_key} 데이터 로드 완료")
        except Exception as e:
            print(f"[통계표] GRDP JSON 로드 실패: {e}")
        
        return {
            "title": "분기 지역내총생산(GRDP)",
            "unit": "[전년동기비, %]",
            "page_number_1": 42,
            "page_number_2": 43,
            "data": {
                "yearly": yearly,
                "quarterly": quarterly,
                "yearly_years": yearly_years,
                "quarterly_keys": quarterly_keys
            }
        }
    
    def render_html(self, output_path: str, year: int = 2025, quarter: int = 2) -> str:
        """HTML 렌더링"""
        data = self.extract_all_tables(year, quarter)
        
        # Jinja2 환경 설정
        template_dir = Path(__file__).parent
        env = Environment(loader=FileSystemLoader(str(template_dir)))
        
        # 커스텀 필터 추가
        def round_value(value):
            if value is None or value == "-":
                return "-"
            if value == "N/A":
                return "N/A"
            try:
                return f"{float(value):.1f}"
            except (ValueError, TypeError):
                return str(value)
        
        env.filters['round_value'] = round_value
        
        template = env.get_template("statistics_table_template.html")
        html_content = template.render(**data)
        
        # 파일 저장
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"통계표가 생성되었습니다: {output_path}")
        return html_content
    
    def export_data_json(self, output_path: str, year: int = 2025, quarter: int = 2):
        """데이터를 JSON으로 내보내기"""
        data = self.extract_all_tables(year, quarter)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print(f"데이터가 저장되었습니다: {output_path}")
    
    def create_historical_template(self, output_path: str):
        """과거 데이터 템플릿 JSON 생성"""
        template = {}
        
        for table_name in self.TABLE_CONFIG.keys():
            template[table_name] = {
                "yearly": {},
                "quarterly": {}
            }
            
            # 과거 연도 (2017-2020)
            for year in ["2017", "2018", "2019", "2020"]:
                template[table_name]["yearly"][year] = {
                    region: None for region in self.ALL_REGIONS
                }
            
            # 과거 분기 (2016.4/4 ~ 2023.1/4)
            quarters = []
            for y in range(2016, 2024):
                for q in ["1/4", "2/4", "3/4", "4/4"]:
                    if y == 2016 and q != "4/4":
                        continue
                    if y == 2023 and q == "2/4":
                        break
                    quarters.append(f"{y}.{q}")
            
            for quarter in quarters:
                template[table_name]["quarterly"][quarter] = {
                    region: None for region in self.ALL_REGIONS
                }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(template, f, ensure_ascii=False, indent=2)
        
        print(f"과거 데이터 템플릿이 생성되었습니다: {output_path}")


def main():
    """메인 실행 함수"""
    import argparse
    
    parser = argparse.ArgumentParser(description='통계표 생성기')
    parser.add_argument('--excel', '-e', required=True, help='분석표 엑셀 파일 경로')
    parser.add_argument('--output', '-o', help='출력 HTML 파일 경로')
    parser.add_argument('--json', '-j', help='데이터 JSON 출력 경로')
    parser.add_argument('--historical', help='과거 데이터 JSON 파일 경로')
    parser.add_argument('--create-historical-template', help='과거 데이터 템플릿 생성')
    parser.add_argument('--year', '-y', type=int, default=2025, help='보고서 연도')
    parser.add_argument('--quarter', '-q', type=int, default=2, help='보고서 분기')
    
    args = parser.parse_args()
    
    generator = StatisticsTableGenerator(args.excel, args.historical)
    
    if args.create_historical_template:
        generator.create_historical_template(args.create_historical_template)
        return
    
    if args.json:
        generator.export_data_json(args.json, args.year, args.quarter)
    
    if args.output:
        generator.render_html(args.output, args.year, args.quarter)


if __name__ == '__main__':
    main()

