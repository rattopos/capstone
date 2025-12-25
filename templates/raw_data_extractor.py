#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
기초자료 직접 추출 클래스

기초자료 수집표에서 2016년 1분기부터 현재까지의 모든 데이터를 직접 추출합니다.
"""

import pandas as pd
import re
from typing import Dict, List, Optional, Any, Tuple
from pathlib import Path


class RawDataExtractor:
    """기초자료에서 직접 데이터 추출하는 클래스"""
    
    # 기초자료 시트 → 보고서 시트 매핑
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
    
    # 유효한 지역 목록
    ALL_REGIONS = [
        '전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
        '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주'
    ]
    
    def __init__(self, raw_excel_path: str, current_year: int, current_quarter: int):
        """
        Args:
            raw_excel_path: 기초자료 엑셀 파일 경로
            current_year: 현재 연도
            current_quarter: 현재 분기 (1-4)
        """
        self.raw_excel_path = Path(raw_excel_path)
        self.current_year = current_year
        self.current_quarter = current_quarter
        self._xl = None
        self._sheet_cache = {}
        self._header_cache = {}
    
    def _get_excel_file(self) -> pd.ExcelFile:
        """ExcelFile 객체 가져오기 (캐시)"""
        if self._xl is None:
            self._xl = pd.ExcelFile(self.raw_excel_path)
        return self._xl
    
    def _load_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """시트 로드 (캐시)"""
        if sheet_name not in self._sheet_cache:
            xl = self._get_excel_file()
            if sheet_name not in xl.sheet_names:
                return None
            try:
                self._sheet_cache[sheet_name] = pd.read_excel(
                    xl, sheet_name=sheet_name, header=None
                )
            except Exception as e:
                print(f"[RawDataExtractor] 시트 로드 실패: {sheet_name} - {e}")
                return None
        return self._sheet_cache[sheet_name]
    
    def parse_sheet_structure(self, sheet_name: str, header_row: int = 2) -> Dict:
        """기초자료 시트의 헤더에서 연도/분기별 컬럼 인덱스 매핑 생성
        
        Args:
            sheet_name: 시트 이름
            header_row: 헤더 행 인덱스 (0-based, 기본값 2 = 3행)
            
        Returns:
            {
                'years': {2020: col_idx, 2021: col_idx, ...},
                'quarters': {'2022 2/4': col_idx, '2022 3/4': col_idx, ...}
            }
        """
        cache_key = f"{sheet_name}_{header_row}"
        if cache_key in self._header_cache:
            return self._header_cache[cache_key]
        
        df = self._load_sheet(sheet_name)
        if df is None:
            return {'years': {}, 'quarters': {}}
        
        year_cols = {}
        quarter_cols = {}
        
        if header_row >= len(df):
            return {'years': {}, 'quarters': {}}
        
        for col_idx in range(len(df.columns)):
            val = df.iloc[header_row, col_idx]
            if pd.isna(val):
                continue
            
            val_str = str(val).strip()
            
            # 연도 패턴: 2020, 2021, ... (정수 또는 "2020.0")
            if isinstance(val, (int, float)) and 2000 <= int(val) <= 2100:
                year_cols[int(val)] = col_idx
            elif re.match(r'^(\d{4})\.?0?$', val_str):
                year = int(re.match(r'^(\d{4})\.?0?$', val_str).group(1))
                year_cols[year] = col_idx
            
            # 분기 패턴: "2022  2/4", "2022 2/4p", "2022.2/4" 등
            quarter_match = re.search(r'(\d{4})[.\s]*(\d)/4', val_str)
            if quarter_match:
                q_year = int(quarter_match.group(1))
                q_num = int(quarter_match.group(2))
                quarter_key = f"{q_year} {q_num}/4"
                quarter_cols[quarter_key] = col_idx
        
        result = {'years': year_cols, 'quarters': quarter_cols}
        self._header_cache[cache_key] = result
        return result
    
    def get_quarter_column_index(self, sheet_name: str, year: int, quarter: int) -> Optional[int]:
        """특정 연도/분기의 컬럼 인덱스 반환
        
        Args:
            sheet_name: 시트 이름
            year: 연도
            quarter: 분기 (1-4)
            
        Returns:
            컬럼 인덱스 또는 None
        """
        structure = self.parse_sheet_structure(sheet_name)
        quarter_key = f"{year} {quarter}/4"
        return structure['quarters'].get(quarter_key)
    
    def get_year_column_index(self, sheet_name: str, year: int) -> Optional[int]:
        """특정 연도의 컬럼 인덱스 반환
        
        Args:
            sheet_name: 시트 이름
            year: 연도
            
        Returns:
            컬럼 인덱스 또는 None
        """
        structure = self.parse_sheet_structure(sheet_name)
        return structure['years'].get(year)
    
    def extract_all_quarters(
        self, 
        sheet_name: str, 
        start_year: int = 2016, 
        start_quarter: int = 1,
        region_column: int = 1,
        classification_column: Optional[int] = None,
        classification_value: Optional[str] = None
    ) -> Dict[str, Dict[str, Any]]:
        """2016년 1분기부터 현재까지의 모든 분기 데이터 추출
        
        Args:
            sheet_name: 시트 이름
            start_year: 시작 연도 (기본값 2016)
            start_quarter: 시작 분기 (기본값 1)
            region_column: 지역명이 있는 컬럼 인덱스 (기본값 1)
            classification_column: 분류단계/구분 컬럼 인덱스
            classification_value: 분류단계/구분 값 (예: "0", "BCD", "계")
            
        Returns:
            {
                '2016.1/4': {'전국': value, '서울': value, ...},
                '2016.2/4': {...},
                ...
            }
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        structure = self.parse_sheet_structure(sheet_name)
        result = {}
        
        # 2016년 1분기부터 현재 분기까지 모든 분기 생성
        current_quarter_key = f"{self.current_year} {self.current_quarter}/4"
        
        for year in range(start_year, self.current_year + 1):
            start_q = start_quarter if year == start_year else 1
            end_q = self.current_quarter if year == self.current_year else 4
            
            for quarter in range(start_q, end_q + 1):
                quarter_key = f"{year} {quarter}/4"
                if year == self.current_year and quarter == self.current_quarter:
                    quarter_key = f"{year}.{quarter}/4p"  # 잠정치
                
                col_idx = structure['quarters'].get(f"{year} {quarter}/4")
                if col_idx is None:
                    continue
                
                quarter_data = {}
                
                # 각 행에서 데이터 추출
                for row_idx in range(len(df)):
                    # 지역명 추출
                    try:
                        region = str(df.iloc[row_idx, region_column]).strip()
                        if region not in self.ALL_REGIONS:
                            continue
                    except (IndexError, ValueError):
                        continue
                    
                    # 분류 단계 필터링
                    if classification_column is not None and classification_value is not None:
                        try:
                            classification = str(df.iloc[row_idx, classification_column]).strip()
                            if classification != classification_value:
                                continue
                        except (IndexError, ValueError):
                            continue
                    
                    # 값 추출
                    try:
                        value = df.iloc[row_idx, col_idx]
                        if pd.notna(value):
                            try:
                                quarter_data[region] = float(value)
                            except (ValueError, TypeError):
                                pass
                    except IndexError:
                        pass
                
                if quarter_data:
                    result[quarter_key] = quarter_data
        
        return result
    
    def extract_yearly_data(
        self,
        sheet_name: str,
        start_year: int = 2016,
        region_column: int = 1,
        classification_column: Optional[int] = None,
        classification_value: Optional[str] = None
    ) -> Dict[str, Dict[str, Any]]:
        """2016년부터 현재 연도까지의 모든 연도 데이터 추출
        
        Args:
            sheet_name: 시트 이름
            start_year: 시작 연도 (기본값 2016)
            region_column: 지역명이 있는 컬럼 인덱스 (기본값 1)
            classification_column: 분류단계/구분 컬럼 인덱스
            classification_value: 분류단계/구분 값
            
        Returns:
            {
                '2016': {'전국': value, '서울': value, ...},
                '2017': {...},
                ...
            }
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        structure = self.parse_sheet_structure(sheet_name)
        result = {}
        
        # 2016년부터 현재 연도까지
        for year in range(start_year, self.current_year + 1):
            col_idx = structure['years'].get(year)
            if col_idx is None:
                continue
            
            year_data = {}
            
            # 각 행에서 데이터 추출
            for row_idx in range(len(df)):
                # 지역명 추출
                try:
                    region = str(df.iloc[row_idx, region_column]).strip()
                    if region not in self.ALL_REGIONS:
                        continue
                except (IndexError, ValueError):
                    continue
                
                # 분류 단계 필터링
                if classification_column is not None and classification_value is not None:
                    try:
                        classification = str(df.iloc[row_idx, classification_column]).strip()
                        if classification != classification_value:
                            continue
                    except (IndexError, ValueError):
                        continue
                
                # 값 추출
                try:
                    value = df.iloc[row_idx, col_idx]
                    if pd.notna(value):
                        try:
                            year_data[region] = float(value)
                        except (ValueError, TypeError):
                            pass
                except IndexError:
                    pass
            
            if year_data:
                result[str(year)] = year_data
        
        return result
    
    def extract_by_region(
        self,
        sheet_name: str,
        region: str,
        start_year: int = 2016,
        start_quarter: int = 1,
        region_column: int = 1,
        classification_column: Optional[int] = None,
        classification_value: Optional[str] = None
    ) -> Dict[str, Any]:
        """특정 지역의 모든 데이터 추출
        
        Args:
            sheet_name: 시트 이름
            region: 지역명
            start_year: 시작 연도
            start_quarter: 시작 분기
            region_column: 지역명 컬럼 인덱스
            classification_column: 분류단계 컬럼 인덱스
            classification_value: 분류단계 값
            
        Returns:
            {
                'quarterly': {'2016.1/4': value, ...},
                'yearly': {'2016': value, ...}
            }
        """
        quarterly = self.extract_all_quarters(
            sheet_name, start_year, start_quarter,
            region_column, classification_column, classification_value
        )
        yearly = self.extract_yearly_data(
            sheet_name, start_year,
            region_column, classification_column, classification_value
        )
        
        result = {
            'quarterly': {q: data.get(region) for q, data in quarterly.items() if region in data},
            'yearly': {y: data.get(region) for y, data in yearly.items() if region in data}
        }
        
        return result
    
    def get_raw_sheet_name(self, report_name: str) -> Optional[str]:
        """보고서 이름에 대응하는 기초자료 시트 이름 반환"""
        return self.RAW_SHEET_MAPPING.get(report_name)
    
    def find_region_row(
        self,
        sheet_name: str,
        region: str,
        region_column: int = 1,
        classification_column: Optional[int] = None,
        classification_value: Optional[str] = None
    ) -> Optional[pd.Series]:
        """특정 지역의 행 찾기
        
        Args:
            sheet_name: 시트 이름
            region: 지역명
            region_column: 지역명 컬럼 인덱스
            classification_column: 분류단계 컬럼 인덱스
            classification_value: 분류단계 값
            
        Returns:
            해당 행의 Series 또는 None
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return None
        
        for row_idx in range(len(df)):
            try:
                row_region = str(df.iloc[row_idx, region_column]).strip()
                if row_region != region:
                    continue
                
                # 분류 단계 확인
                if classification_column is not None and classification_value is not None:
                    try:
                        classification = str(df.iloc[row_idx, classification_column]).strip()
                        if classification != classification_value:
                            continue
                    except (IndexError, ValueError):
                        continue
                
                return df.iloc[row_idx]
            except (IndexError, ValueError):
                continue
        
        return None

