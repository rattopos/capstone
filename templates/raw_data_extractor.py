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
        
        # 파일 존재 및 수정 시간 확인
        if not self.raw_excel_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {raw_excel_path}")
        
        try:
            self._file_mtime = self.raw_excel_path.stat().st_mtime
        except OSError:
            self._file_mtime = None
    
    def _check_file_modified(self) -> bool:
        """파일이 수정되었는지 확인 (캐시 무효화 판단)"""
        if self._file_mtime is None:
            return False  # 수정 시간을 확인할 수 없으면 변경되지 않은 것으로 간주
        
        try:
            if not self.raw_excel_path.exists():
                return True  # 파일이 없으면 무효화 필요
            
            current_mtime = self.raw_excel_path.stat().st_mtime
            if abs(current_mtime - self._file_mtime) > 1.0:  # 1초 이상 차이
                return True
            return False
        except OSError:
            return True  # 파일 접근 오류 시 안전하게 무효화
    
    def _clear_cache(self):
        """캐시 무효화"""
        self._sheet_cache.clear()
        self._header_cache.clear()
        if self._xl is not None:
            try:
                self._xl.close()
            except:
                pass
            self._xl = None
    
    def _get_excel_file(self) -> pd.ExcelFile:
        """ExcelFile 객체 가져오기 (캐시, 파일 수정 시간 확인)"""
        # 파일이 수정되었으면 캐시 무효화
        if self._check_file_modified():
            self._clear_cache()
            try:
                self._file_mtime = self.raw_excel_path.stat().st_mtime
            except OSError:
                pass
        
        if self._xl is None:
            try:
                self._xl = pd.ExcelFile(self.raw_excel_path)
            except Exception as e:
                raise RuntimeError(f"엑셀 파일을 열 수 없습니다: {self.raw_excel_path} - {e}")
        return self._xl
    
    def _load_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """시트 로드 (캐시, 파일 수정 시간 확인)"""
        # 파일이 수정되었으면 캐시 무효화
        if self._check_file_modified():
            self._clear_cache()
        
        if sheet_name not in self._sheet_cache:
            xl = self._get_excel_file()
            if sheet_name not in xl.sheet_names:
                return None
            try:
                # 시트 데이터 로드 (예외 발생 시 캐시에 저장하지 않음)
                df = pd.read_excel(
                    xl, sheet_name=sheet_name, header=None
                )
                # 정상 로드된 경우에만 캐시에 저장
                self._sheet_cache[sheet_name] = df
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
    
    def extract_yearly_difference(
        self,
        sheet_name: str,
        start_year: int = 2016,
        region_column: int = 1,
        classification_column: Optional[int] = None,
        classification_value: Optional[str] = None
    ) -> Dict[str, Dict[str, Any]]:
        """연도별 전년동기대비 차이(%p) 계산하여 추출
        
        고용률, 실업률 등 %p 단위 지표에 사용
        
        Args:
            sheet_name: 시트 이름
            start_year: 시작 연도
            region_column: 지역명 컬럼 인덱스
            classification_column: 분류단계 컬럼 인덱스
            classification_value: 분류단계 값
            
        Returns:
            {'2016': {'전국': 차이, ...}, ...}
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        structure = self.parse_sheet_structure(sheet_name)
        raw_data = {}
        
        for year in range(start_year - 1, self.current_year + 1):
            col_idx = structure['years'].get(year)
            if col_idx is None:
                continue
            
            year_data = {}
            for row_idx in range(len(df)):
                try:
                    region = str(df.iloc[row_idx, region_column]).strip()
                    if region not in self.ALL_REGIONS:
                        continue
                    
                    if region in year_data:
                        continue
                    
                    if classification_column is not None and classification_value is not None:
                        classification = str(df.iloc[row_idx, classification_column]).strip()
                        if classification != classification_value:
                            continue
                    
                    value = df.iloc[row_idx, col_idx]
                    if pd.notna(value):
                        year_data[region] = float(value)
                except (IndexError, ValueError):
                    continue
            
            if year_data:
                raw_data[year] = year_data
        
        # 차이 계산 (현재값 - 전년값)
        result = {}
        for year in range(start_year, self.current_year + 1):
            if year not in raw_data or (year - 1) not in raw_data:
                continue
            
            diff_data = {}
            for region in self.ALL_REGIONS:
                current = raw_data[year].get(region)
                previous = raw_data[year - 1].get(region)
                
                if current is not None and previous is not None:
                    diff = current - previous
                    diff_data[region] = round(diff, 1)
            
            if diff_data:
                result[str(year)] = diff_data
        
        return result
    
    def extract_quarterly_difference(
        self,
        sheet_name: str,
        start_year: int = 2016,
        start_quarter: int = 1,
        region_column: int = 1,
        classification_column: Optional[int] = None,
        classification_value: Optional[str] = None
    ) -> Dict[str, Dict[str, Any]]:
        """분기별 전년동분기대비 차이(%p) 계산하여 추출
        
        Args:
            sheet_name: 시트 이름
            start_year: 시작 연도
            start_quarter: 시작 분기
            region_column: 지역명 컬럼 인덱스
            classification_column: 분류단계 컬럼 인덱스
            classification_value: 분류단계 값
            
        Returns:
            {'2016.1/4': {'전국': 차이, ...}, ...}
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        structure = self.parse_sheet_structure(sheet_name)
        raw_data = {}
        
        for year in range(start_year - 1, self.current_year + 1):
            for quarter in range(1, 5):
                if year == self.current_year and quarter > self.current_quarter:
                    break
                
                quarter_key = f"{year} {quarter}/4"
                col_idx = structure['quarters'].get(quarter_key)
                if col_idx is None:
                    continue
                
                quarter_data = {}
                for row_idx in range(len(df)):
                    try:
                        region = str(df.iloc[row_idx, region_column]).strip()
                        if region not in self.ALL_REGIONS:
                            continue
                        
                        if region in quarter_data:
                            continue
                        
                        if classification_column is not None and classification_value is not None:
                            classification = str(df.iloc[row_idx, classification_column]).strip()
                            if classification != classification_value:
                                continue
                        
                        value = df.iloc[row_idx, col_idx]
                        if pd.notna(value):
                            quarter_data[region] = float(value)
                    except (IndexError, ValueError):
                        continue
                
                if quarter_data:
                    raw_data[quarter_key] = quarter_data
        
        # 차이 계산
        result = {}
        for year in range(start_year, self.current_year + 1):
            s_quarter = start_quarter if year == start_year else 1
            e_quarter = self.current_quarter if year == self.current_year else 4
            
            for quarter in range(s_quarter, e_quarter + 1):
                current_key = f"{year} {quarter}/4"
                previous_key = f"{year - 1} {quarter}/4"
                
                if current_key not in raw_data or previous_key not in raw_data:
                    continue
                
                diff_data = {}
                for region in self.ALL_REGIONS:
                    current = raw_data[current_key].get(region)
                    previous = raw_data[previous_key].get(region)
                    
                    if current is not None and previous is not None:
                        diff = current - previous
                        diff_data[region] = round(diff, 1)
                
                if diff_data:
                    if year == self.current_year and quarter == self.current_quarter:
                        result_key = f"{year}.{quarter}/4p"
                    else:
                        result_key = f"{year}.{quarter}/4"
                    result[result_key] = diff_data
        
        return result
    
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
    
    def extract_yearly_growth_rate(
        self,
        sheet_name: str,
        start_year: int = 2016,
        region_column: int = 1,
        classification_column: Optional[int] = None,
        classification_value: Optional[str] = None
    ) -> Dict[str, Dict[str, Any]]:
        """연도별 전년동기비(%) 계산하여 추출
        
        Args:
            sheet_name: 시트 이름
            start_year: 시작 연도 (기본값 2016)
            region_column: 지역명이 있는 컬럼 인덱스
            classification_column: 분류단계/구분 컬럼 인덱스
            classification_value: 분류단계/구분 값
            
        Returns:
            {
                '2016': {'전국': 증감률, '서울': 증감률, ...},
                '2017': {...},
                ...
            }
        """
        # 먼저 원지수 데이터 추출 (전년 데이터 포함을 위해 start_year-1부터)
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        structure = self.parse_sheet_structure(sheet_name)
        raw_data = {}  # {year: {region: value}}
        
        # 모든 연도의 원지수 추출 (각 지역별로 첫 번째로 발견된 행만 사용)
        for year in range(start_year - 1, self.current_year + 1):
            col_idx = structure['years'].get(year)
            if col_idx is None:
                continue
            
            year_data = {}
            for row_idx in range(len(df)):
                try:
                    region = str(df.iloc[row_idx, region_column]).strip()
                    if region not in self.ALL_REGIONS:
                        continue
                    
                    # 이미 해당 지역 데이터가 있으면 스킵 (첫 번째 행만 사용)
                    if region in year_data:
                        continue
                    
                    if classification_column is not None and classification_value is not None:
                        classification = str(df.iloc[row_idx, classification_column]).strip()
                        if classification != classification_value:
                            continue
                    
                    value = df.iloc[row_idx, col_idx]
                    if pd.notna(value):
                        year_data[region] = float(value)
                except (IndexError, ValueError):
                    continue
            
            if year_data:
                raw_data[year] = year_data
        
        # 전년동기비 계산
        result = {}
        for year in range(start_year, self.current_year + 1):
            if year not in raw_data or (year - 1) not in raw_data:
                continue
            
            growth_data = {}
            for region in self.ALL_REGIONS:
                current = raw_data[year].get(region)
                previous = raw_data[year - 1].get(region)
                
                if current is not None and previous is not None and previous != 0:
                    growth_rate = ((current / previous) - 1) * 100
                    growth_data[region] = round(growth_rate, 1)
            
            if growth_data:
                result[str(year)] = growth_data
        
        return result
    
    def extract_quarterly_growth_rate(
        self,
        sheet_name: str,
        start_year: int = 2016,
        start_quarter: int = 1,
        region_column: int = 1,
        classification_column: Optional[int] = None,
        classification_value: Optional[str] = None
    ) -> Dict[str, Dict[str, Any]]:
        """분기별 전년동분기비(%) 계산하여 추출
        
        Args:
            sheet_name: 시트 이름
            start_year: 시작 연도 (기본값 2016)
            start_quarter: 시작 분기 (기본값 1)
            region_column: 지역명이 있는 컬럼 인덱스
            classification_column: 분류단계/구분 컬럼 인덱스
            classification_value: 분류단계/구분 값
            
        Returns:
            {
                '2016.1/4': {'전국': 증감률, ...},
                '2016.2/4': {...},
                ...
            }
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        structure = self.parse_sheet_structure(sheet_name)
        
        # 모든 분기의 원지수 추출 (전년도 분기 포함)
        raw_data = {}  # {'2015 1/4': {region: value}, ...}
        
        for year in range(start_year - 1, self.current_year + 1):
            for quarter in range(1, 5):
                if year == self.current_year and quarter > self.current_quarter:
                    break
                
                quarter_key = f"{year} {quarter}/4"
                col_idx = structure['quarters'].get(quarter_key)
                if col_idx is None:
                    continue
                
                quarter_data = {}
                for row_idx in range(len(df)):
                    try:
                        region = str(df.iloc[row_idx, region_column]).strip()
                        if region not in self.ALL_REGIONS:
                            continue
                        
                        # 이미 해당 지역 데이터가 있으면 스킵 (첫 번째 행만 사용)
                        if region in quarter_data:
                            continue
                        
                        if classification_column is not None and classification_value is not None:
                            classification = str(df.iloc[row_idx, classification_column]).strip()
                            if classification != classification_value:
                                continue
                        
                        value = df.iloc[row_idx, col_idx]
                        if pd.notna(value):
                            quarter_data[region] = float(value)
                    except (IndexError, ValueError):
                        continue
                
                if quarter_data:
                    raw_data[quarter_key] = quarter_data
        
        # 전년동분기비 계산
        result = {}
        for year in range(start_year, self.current_year + 1):
            s_quarter = start_quarter if year == start_year else 1
            e_quarter = self.current_quarter if year == self.current_year else 4
            
            for quarter in range(s_quarter, e_quarter + 1):
                current_key = f"{year} {quarter}/4"
                previous_key = f"{year - 1} {quarter}/4"
                
                if current_key not in raw_data or previous_key not in raw_data:
                    continue
                
                growth_data = {}
                for region in self.ALL_REGIONS:
                    current = raw_data[current_key].get(region)
                    previous = raw_data[previous_key].get(region)
                    
                    if current is not None and previous is not None and previous != 0:
                        growth_rate = ((current / previous) - 1) * 100
                        growth_data[region] = round(growth_rate, 1)
                
                if growth_data:
                    # 현재 분기는 잠정치(p) 표시
                    if year == self.current_year and quarter == self.current_quarter:
                        result_key = f"{year}.{quarter}/4p"
                    else:
                        result_key = f"{year}.{quarter}/4"
                    result[result_key] = growth_data
        
        return result

