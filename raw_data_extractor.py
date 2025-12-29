#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
기초자료 추출기

기초자료 수집표 엑셀에서 통계표용 데이터를 추출합니다.
전년동기비 증감률 또는 차이(%p)를 계산합니다.
"""

import pandas as pd
import re
from pathlib import Path
from typing import Dict, List, Optional, Any


class RawDataExtractor:
    """기초자료 추출 클래스"""
    
    # 지역명 매핑 (다양한 형태 -> 표준 이름)
    REGION_MAPPING = {
        "전국": "전국",
        "서울": "서울", "서울특별시": "서울",
        "부산": "부산", "부산광역시": "부산",
        "대구": "대구", "대구광역시": "대구",
        "인천": "인천", "인천광역시": "인천",
        "광주": "광주", "광주광역시": "광주",
        "대전": "대전", "대전광역시": "대전",
        "울산": "울산", "울산광역시": "울산",
        "세종": "세종", "세종특별자치시": "세종",
        "경기": "경기", "경기도": "경기",
        "강원": "강원", "강원도": "강원", "강원특별자치도": "강원",
        "충북": "충북", "충청북도": "충북",
        "충남": "충남", "충청남도": "충남",
        "전북": "전북", "전라북도": "전북", "전북특별자치도": "전북",
        "전남": "전남", "전라남도": "전남",
        "경북": "경북", "경상북도": "경북",
        "경남": "경남", "경상남도": "경남",
        "제주": "제주", "제주도": "제주", "제주특별자치도": "제주"
    }
    
    ALL_REGIONS = ["전국", "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종",
                   "경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"]
    
    def __init__(self, excel_path: str, current_year: int = 2025, current_quarter: int = 2):
        """
        초기화
        
        Args:
            excel_path: 기초자료 수집표 엑셀 파일 경로
            current_year: 현재 연도
            current_quarter: 현재 분기
        """
        self.excel_path = excel_path
        self.current_year = current_year
        self.current_quarter = current_quarter
        self.cached_sheets = {}
        
        # 엑셀 파일 로드
        self.xl = pd.ExcelFile(excel_path)
        print(f"[기초자료 추출기] 초기화 완료: {excel_path}")
        print(f"[기초자료 추출기] 시트 목록: {self.xl.sheet_names}")
    
    def _load_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """시트 로드 (캐싱)"""
        if sheet_name not in self.cached_sheets:
            if sheet_name not in self.xl.sheet_names:
                print(f"[기초자료 추출기] 시트 없음: {sheet_name}")
                return None
            try:
                self.cached_sheets[sheet_name] = pd.read_excel(
                    self.xl, sheet_name=sheet_name, header=None
                )
            except Exception as e:
                print(f"[기초자료 추출기] 시트 로드 실패: {sheet_name} - {e}")
                return None
        return self.cached_sheets[sheet_name]
    
    def _normalize_region(self, region_name: str) -> Optional[str]:
        """지역명 정규화"""
        if pd.isna(region_name):
            return None
        region_clean = str(region_name).strip()
        return self.REGION_MAPPING.get(region_clean)
    
    def _parse_header_columns(self, df: pd.DataFrame, header_row_idx: int = 2) -> Dict[str, int]:
        """
        헤더 행에서 연도/분기 컬럼 매핑 추출
        
        Returns:
            {'2020': col_idx, '2021': col_idx, ..., '2024.1/4': col_idx, ...}
        """
        col_map = {}
        header_row = df.iloc[header_row_idx]
        
        for col_idx, header in enumerate(header_row):
            if pd.isna(header):
                continue
            header_str = str(header).strip()
            
            # 연도만 있는 경우 (예: "2020", "2020.0")
            year_match = re.match(r'^(\d{4})(\.0)?$', header_str)
            if year_match:
                year = year_match.group(1)
                col_map[year] = col_idx
                continue
            
            # 분기 형식 (예: "2022  2/4", "2025  2/4p")
            quarter_match = re.match(r'(\d{4})\s+(\d)/4(p)?', header_str)
            if quarter_match:
                year = quarter_match.group(1)
                q = quarter_match.group(2)
                is_preliminary = quarter_match.group(3) == 'p'
                quarter_key = f"{year}.{q}/4"
                if is_preliminary:
                    quarter_key += "p"
                col_map[quarter_key] = col_idx
        
        return col_map
    
    def _find_region_rows(self, df: pd.DataFrame, region_col: int, 
                          classification_col: Optional[int] = None,
                          classification_value: Optional[str] = None) -> Dict[str, int]:
        """
        각 지역의 데이터 행 찾기
        
        Returns:
            {'전국': row_idx, '서울': row_idx, ...}
        """
        region_rows = {}
        
        for row_idx, row in df.iterrows():
            region_raw = row.iloc[region_col] if region_col < len(row) else None
            region = self._normalize_region(region_raw)
            
            if region is None:
                continue
            
            # 분류값 필터링 (있으면)
            if classification_col is not None and classification_value is not None:
                class_val = str(row.iloc[classification_col]).strip() if classification_col < len(row) else ""
                if class_val != str(classification_value).strip():
                    continue
            
            # 첫 번째 일치하는 행 저장
            if region not in region_rows:
                region_rows[region] = row_idx
        
        return region_rows
    
    def extract_yearly_growth_rate(self, sheet_name: str, start_year: int = 2016,
                                   region_column: int = 1, 
                                   classification_column: Optional[int] = None,
                                   classification_value: Optional[str] = None) -> Dict[str, Dict[str, float]]:
        """
        연도별 전년비 증감률 추출
        
        Returns:
            {'2017': {'전국': 5.2, '서울': 3.1, ...}, '2018': {...}, ...}
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        # 헤더 파싱
        col_map = self._parse_header_columns(df)
        
        # 지역별 행 찾기
        region_rows = self._find_region_rows(df, region_column, classification_column, classification_value)
        
        result = {}
        
        for year in range(start_year, self.current_year + 1):
            year_str = str(year)
            prev_year_str = str(year - 1)
            
            curr_col = col_map.get(year_str)
            prev_col = col_map.get(prev_year_str)
            
            if curr_col is None:
                # 연도 컬럼이 없으면 스킵
                continue
            
            year_data = {}
            
            for region in self.ALL_REGIONS:
                row_idx = region_rows.get(region)
                if row_idx is None:
                    continue
                
                row = df.iloc[row_idx]
                
                try:
                    curr_val = float(row.iloc[curr_col]) if curr_col < len(row) and pd.notna(row.iloc[curr_col]) else None
                    prev_val = float(row.iloc[prev_col]) if prev_col and prev_col < len(row) and pd.notna(row.iloc[prev_col]) else None
                    
                    if curr_val is not None and prev_val is not None and prev_val != 0:
                        growth = ((curr_val - prev_val) / prev_val) * 100
                        year_data[region] = round(growth, 1)
                except (ValueError, TypeError, IndexError):
                    continue
            
            if year_data:
                result[year_str] = year_data
        
        return result
    
    def extract_quarterly_growth_rate(self, sheet_name: str, start_year: int = 2016, start_quarter: int = 1,
                                      region_column: int = 1,
                                      classification_column: Optional[int] = None,
                                      classification_value: Optional[str] = None) -> Dict[str, Dict[str, float]]:
        """
        분기별 전년동기비 증감률 추출
        
        Returns:
            {'2017 1/4': {'전국': 5.2, '서울': 3.1, ...}, ...}
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        col_map = self._parse_header_columns(df)
        region_rows = self._find_region_rows(df, region_column, classification_column, classification_value)
        
        result = {}
        
        # 분기 키들 생성
        for year in range(start_year, self.current_year + 1):
            sq = start_quarter if year == start_year else 1
            eq = self.current_quarter if year == self.current_year else 4
            
            for q in range(sq, eq + 1):
                # 현재 분기 키
                is_current = (year == self.current_year and q == self.current_quarter)
                quarter_key = f"{year}.{q}/4"
                if is_current:
                    quarter_key_with_p = quarter_key + "p"
                else:
                    quarter_key_with_p = quarter_key
                
                # 전년동분기 키
                prev_quarter_key = f"{year - 1}.{q}/4"
                
                # 컬럼 찾기
                curr_col = col_map.get(quarter_key_with_p) or col_map.get(quarter_key)
                prev_col = col_map.get(prev_quarter_key)
                
                if curr_col is None:
                    continue
                
                quarter_data = {}
                
                for region in self.ALL_REGIONS:
                    row_idx = region_rows.get(region)
                    if row_idx is None:
                        continue
                    
                    row = df.iloc[row_idx]
                    
                    try:
                        curr_val = float(row.iloc[curr_col]) if curr_col < len(row) and pd.notna(row.iloc[curr_col]) else None
                        prev_val = float(row.iloc[prev_col]) if prev_col and prev_col < len(row) and pd.notna(row.iloc[prev_col]) else None
                        
                        if curr_val is not None and prev_val is not None and prev_val != 0:
                            growth = ((curr_val - prev_val) / prev_val) * 100
                            quarter_data[region] = round(growth, 1)
                    except (ValueError, TypeError, IndexError):
                        continue
                
                if quarter_data:
                    # 출력 키는 공백 형식 사용 (통계표 생성기와 호환)
                    output_key = f"{year} {q}/4"
                    result[output_key] = quarter_data
        
        return result
    
    def extract_yearly_difference(self, sheet_name: str, start_year: int = 2016,
                                  region_column: int = 1,
                                  classification_column: Optional[int] = None,
                                  classification_value: Optional[str] = None) -> Dict[str, Dict[str, float]]:
        """
        연도별 전년비 차이 추출 (%p 단위)
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        col_map = self._parse_header_columns(df)
        region_rows = self._find_region_rows(df, region_column, classification_column, classification_value)
        
        result = {}
        
        for year in range(start_year, self.current_year + 1):
            year_str = str(year)
            prev_year_str = str(year - 1)
            
            curr_col = col_map.get(year_str)
            prev_col = col_map.get(prev_year_str)
            
            if curr_col is None:
                continue
            
            year_data = {}
            
            for region in self.ALL_REGIONS:
                row_idx = region_rows.get(region)
                if row_idx is None:
                    continue
                
                row = df.iloc[row_idx]
                
                try:
                    curr_val = float(row.iloc[curr_col]) if curr_col < len(row) and pd.notna(row.iloc[curr_col]) else None
                    prev_val = float(row.iloc[prev_col]) if prev_col and prev_col < len(row) and pd.notna(row.iloc[prev_col]) else None
                    
                    if curr_val is not None and prev_val is not None:
                        diff = curr_val - prev_val
                        year_data[region] = round(diff, 1)
                except (ValueError, TypeError, IndexError):
                    continue
            
            if year_data:
                result[year_str] = year_data
        
        return result
    
    def extract_quarterly_difference(self, sheet_name: str, start_year: int = 2016, start_quarter: int = 1,
                                     region_column: int = 1,
                                     classification_column: Optional[int] = None,
                                     classification_value: Optional[str] = None) -> Dict[str, Dict[str, float]]:
        """
        분기별 전년동기비 차이 추출 (%p 단위)
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        col_map = self._parse_header_columns(df)
        region_rows = self._find_region_rows(df, region_column, classification_column, classification_value)
        
        result = {}
        
        for year in range(start_year, self.current_year + 1):
            sq = start_quarter if year == start_year else 1
            eq = self.current_quarter if year == self.current_year else 4
            
            for q in range(sq, eq + 1):
                is_current = (year == self.current_year and q == self.current_quarter)
                quarter_key = f"{year}.{q}/4"
                if is_current:
                    quarter_key_with_p = quarter_key + "p"
                else:
                    quarter_key_with_p = quarter_key
                
                prev_quarter_key = f"{year - 1}.{q}/4"
                
                curr_col = col_map.get(quarter_key_with_p) or col_map.get(quarter_key)
                prev_col = col_map.get(prev_quarter_key)
                
                if curr_col is None:
                    continue
                
                quarter_data = {}
                
                for region in self.ALL_REGIONS:
                    row_idx = region_rows.get(region)
                    if row_idx is None:
                        continue
                    
                    row = df.iloc[row_idx]
                    
                    try:
                        curr_val = float(row.iloc[curr_col]) if curr_col < len(row) and pd.notna(row.iloc[curr_col]) else None
                        prev_val = float(row.iloc[prev_col]) if prev_col and prev_col < len(row) and pd.notna(row.iloc[prev_col]) else None
                        
                        if curr_val is not None and prev_val is not None:
                            diff = curr_val - prev_val
                            quarter_data[region] = round(diff, 1)
                    except (ValueError, TypeError, IndexError):
                        continue
                
                if quarter_data:
                    output_key = f"{year} {q}/4"
                    result[output_key] = quarter_data
        
        return result


def main():
    """테스트 실행"""
    import sys
    
    if len(sys.argv) < 2:
        print("사용법: python raw_data_extractor.py <엑셀파일경로>")
        return
    
    excel_path = sys.argv[1]
    extractor = RawDataExtractor(excel_path, 2025, 2)
    
    # 광공업생산 테스트
    print("\n=== 광공업생산 연도별 증감률 ===")
    yearly = extractor.extract_yearly_growth_rate("광공업생산", region_column=1, 
                                                   classification_column=2, classification_value="0")
    for year, data in sorted(yearly.items()):
        print(f"{year}: 전국={data.get('전국', '-')}, 서울={data.get('서울', '-')}")
    
    print("\n=== 광공업생산 분기별 증감률 ===")
    quarterly = extractor.extract_quarterly_growth_rate("광공업생산", region_column=1,
                                                         classification_column=2, classification_value="0")
    for qk, data in sorted(quarterly.items())[-5:]:
        print(f"{qk}: 전국={data.get('전국', '-')}, 서울={data.get('서울', '-')}")


if __name__ == '__main__':
    main()

