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
    
    # 기초자료 시트 → 보도자료 시트 매핑
    RAW_SHEET_MAPPING = {
        "광공업생산지수": "광공업생산",
        "서비스업생산지수": "서비스업생산",
        "소매판매액지수": "소비(소매, 추가)",
        "건설수주액": "건설 (공표자료)",
        "고용률": "고용률",
        "실업률": "실업자 수",
        "수출액": "수출",
        "수입액": "수입",
        "국내인구이동": "시도 간 이동",
        "물가": "품목성질별 물가"
    }
    
    # 기초자료 시트별 기본 설정 (열 인덱스는 동적으로 파악)
    # 분기 컬럼은 parse_sheet_structure()로 동적으로 찾음
    RAW_SHEET_CONFIG = {
        "광공업생산": {
            'region_col': 1, 'level_col': 2, 'name_col': 5, 'header_row': 2, 'total_code': '0'
        },
        "서비스업생산": {
            'region_col': 1, 'level_col': 2, 'name_col': 5, 'header_row': 2, 'total_code': '0'
        },
        "소비(소매, 추가)": {
            'region_col': 1, 'level_col': 2, 'name_col': 4, 'header_row': 2, 'total_code': '0'
        },
        "건설 (공표자료)": {
            'region_col': 1, 'level_col': 2, 'name_col': 4, 'header_row': 2, 'total_code': '0'
        },
        "수출": {
            'region_col': 1, 'level_col': 2, 'name_col': 5, 'header_row': 2, 'total_code': '계'
        },
        "수입": {
            'region_col': 1, 'level_col': 2, 'name_col': 5, 'header_row': 2, 'total_code': '계'
        },
        "품목성질별 물가": {
            'region_col': 0, 'level_col': 1, 'name_col': 3, 'header_row': 2, 'total_code': '총지수'
        },
        "고용": {
            'region_col': 1, 'level_col': 2, 'name_col': 3, 'header_row': 2, 'total_code': '계'
        },
        "연령별고용률": {
            'region_col': 1, 'level_col': 2, 'name_col': 3, 'header_row': 2, 'total_code': '계'
        },
        "실업자 수": {
            'region_col': 0, 'level_col': 1, 'name_col': 1, 'header_row': 2, 'total_code': '계'
        },
        "시도 간 이동": {
            'region_col': 1, 'level_col': 2, 'name_col': 1, 'header_row': 2, 'total_code': None
        },
    }
    
    # 레거시 호환성: RAW_SHEET_QUARTER_COLS는 RAW_SHEET_CONFIG를 참조
    @property
    def RAW_SHEET_QUARTER_COLS(self):
        """레거시 호환성 - 동적으로 config 반환"""
        return self.RAW_SHEET_CONFIG
    
    # 유효한 지역 목록
    ALL_REGIONS = [
        '전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
        '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주'
    ]
    
    # 지역명 전체 이름 → 단축명 매핑 (실업자 수 시트 등에서 사용)
    REGION_FULL_TO_SHORT = {
        '전국': '전국',
        '서울특별시': '서울', '서울': '서울',
        '부산광역시': '부산', '부산': '부산',
        '대구광역시': '대구', '대구': '대구',
        '인천광역시': '인천', '인천': '인천',
        '광주광역시': '광주', '광주': '광주',
        '대전광역시': '대전', '대전': '대전',
        '울산광역시': '울산', '울산': '울산',
        '세종특별자치시': '세종', '세종': '세종',
        '경기도': '경기', '경기': '경기',
        '강원특별자치도': '강원', '강원도': '강원', '강원': '강원',
        '충청북도': '충북', '충북': '충북',
        '충청남도': '충남', '충남': '충남',
        '전북특별자치도': '전북', '전라북도': '전북', '전북': '전북',
        '전라남도': '전남', '전남': '전남',
        '경상북도': '경북', '경북': '경북',
        '경상남도': '경남', '경남': '경남',
        '제주특별자치도': '제주', '제주': '제주',
    }
    
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
        """보도자료 이름에 대응하는 기초자료 시트 이름 반환"""
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
    
    # ========== 보도자료 데이터 추출 메서드 ==========
    
    def extract_mining_manufacturing_report_data(self) -> Dict[str, Any]:
        """광공업생산 보도자료 데이터 추출
        
        Returns:
            보도자료 템플릿에서 사용할 데이터 딕셔너리
        """
        sheet_name = '광공업생산'
        
        # 기본 정보
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '광공업생산지수',
            },
            'national_summary': {},
            'regional_data': {},
            'table_data': [],
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
            'summary_box': {},  # 템플릿에서 사용하는 요약 박스 데이터
            'nationwide_data': {},  # 전국 데이터
        }
        
        # 전년동분기비 증감률 추출 (분류단계 0 = 전체)
        quarterly_growth = self.extract_quarterly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=1,
            classification_column=2,
            classification_value='0'
        )
        
        # 전년동기비 증감률 추출
        yearly_growth = self.extract_yearly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=1,
            classification_column=2,
            classification_value='0'
        )
        
        # 현재 분기 데이터
        current_quarter_key = f"{self.current_year}.{self.current_quarter}/4p"
        current_data = quarterly_growth.get(current_quarter_key, {})
        
        if not current_data:
            # p 없이 다시 시도
            current_quarter_key = f"{self.current_year}.{self.current_quarter}/4"
            current_data = quarterly_growth.get(current_quarter_key, {})
        
        # 전국 데이터 (결측치는 None 유지)
        national_rate = current_data.get('전국', None)
        report_data['national_summary'] = {
            'growth_rate': national_rate,
            'direction': '증가' if national_rate is not None and national_rate > 0 else ('감소' if national_rate is not None and national_rate < 0 else ('보합' if national_rate == 0 else 'N/A')),
            'trend': '확대' if national_rate > 0 else ('축소' if national_rate < 0 else '유지'),
        }
        
        # nationwide_data (템플릿 호환성)
        report_data['nationwide_data'] = {
            'production_index': 100.0,  # 기준값
            'growth_rate': national_rate if national_rate is not None else 0.0,
            'main_industries': [],  # 가중치 필요하여 N/A
        }
        
        # 지역별 데이터 처리
        regional_list = []
        for region in self.ALL_REGIONS:
            if region == '전국':
                continue
            rate = current_data.get(region, None)
            if rate is None:
                continue
            regional_list.append({
                'region': region,
                'growth_rate': rate,
                'direction': '증가' if rate is not None and rate > 0 else ('감소' if rate is not None and rate < 0 else ('보합' if rate == 0 else 'N/A')),
                'industries': [],  # 가중치 필요
            })
        
        # 증가/감소 지역 분류
        increase_regions = sorted([r for r in regional_list if r.get('growth_rate') and r['growth_rate'] > 0], 
                                  key=lambda x: x['growth_rate'], reverse=True)
        decrease_regions = sorted([r for r in regional_list if r.get('growth_rate') and r['growth_rate'] < 0], 
                                  key=lambda x: x['growth_rate'])
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        
        # summary_box (템플릿에서 사용)
        report_data['summary_box'] = {
            'main_increase_regions': increase_regions[:3],
            'main_decrease_regions': decrease_regions[:3],
            'region_count': len(increase_regions),
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
        }
        
        # 테이블 데이터 (지역별 연도/분기 데이터)
        report_data['yearly_data'] = yearly_growth
        report_data['quarterly_data'] = quarterly_growth
        
        # summary_table 생성
        config = self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        report_data['summary_table'] = self._generate_production_summary_table(sheet_name, config, quarterly_growth, yearly_growth)
        
        return report_data
    
    def extract_service_industry_report_data(self) -> Dict[str, Any]:
        """서비스업생산 보도자료 데이터 추출"""
        sheet_name = '서비스업생산'
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '서비스업생산지수',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
            'summary_box': {},
            'nationwide_data': {},
        }
        
        # 전년동분기비 증감률 추출
        quarterly_growth = self.extract_quarterly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=1,
            classification_column=2,
            classification_value='0'
        )
        
        yearly_growth = self.extract_yearly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=1,
            classification_column=2,
            classification_value='0'
        )
        
        # 현재 분기 데이터
        current_quarter_key = f"{self.current_year}.{self.current_quarter}/4p"
        current_data = quarterly_growth.get(current_quarter_key, 
                       quarterly_growth.get(f"{self.current_year}.{self.current_quarter}/4", {}))
        
        national_rate = current_data.get('전국', None)
        report_data['national_summary'] = {
            'growth_rate': national_rate,
            'direction': '증가' if national_rate is not None and national_rate > 0 else ('감소' if national_rate is not None and national_rate < 0 else ('보합' if national_rate == 0 else 'N/A')),
        }
        
        # nationwide_data (템플릿 호환성)
        report_data['nationwide_data'] = {
            'production_index': 100.0,
            'growth_rate': national_rate if national_rate is not None else 0.0,
            'main_industries': [],
        }
        
        # 지역별 처리
        regional_list = []
        for region in self.ALL_REGIONS:
            if region == '전국':
                continue
            rate = current_data.get(region, None)
            if rate is None:
                continue
            regional_list.append({
                'region': region,
                'growth_rate': rate,
                'direction': '증가' if rate is not None and rate > 0 else ('감소' if rate is not None and rate < 0 else ('보합' if rate == 0 else 'N/A')),
                'industries': [],
            })
        
        increase_regions = sorted([r for r in regional_list if r.get('growth_rate') and r['growth_rate'] > 0], 
                                  key=lambda x: x['growth_rate'], reverse=True)
        decrease_regions = sorted([r for r in regional_list if r.get('growth_rate') and r['growth_rate'] < 0], 
                                  key=lambda x: x['growth_rate'])
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        
        # summary_box
        report_data['summary_box'] = {
            'main_increase_regions': increase_regions[:3],
            'main_decrease_regions': decrease_regions[:3],
            'region_count': len(increase_regions),
        }
        
        report_data['yearly_data'] = yearly_growth
        report_data['quarterly_data'] = quarterly_growth
        
        # summary_table 생성
        config = self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        report_data['summary_table'] = self._generate_production_summary_table(sheet_name, config, quarterly_growth, yearly_growth)
        
        return report_data
    
    def _extract_raw_indices_for_table(self, sheet_name: str, config: Dict) -> Dict[str, List[float]]:
        """요약 테이블용 원지수 데이터 추출 (전년동분기, 현재분기)
        
        Args:
            sheet_name: 시트 이름
            config: RAW_SHEET_QUARTER_COLS에서 가져온 설정
            
        Returns:
            {
                '전국': [전년동분기 원지수, 현재분기 원지수],
                '서울': [전년동분기 원지수, 현재분기 원지수],
                ...
            }
        """
        result = {}
        
        df = self._load_sheet(sheet_name)
        if df is None:
            return result
        
        # 전년동분기와 현재분기 컬럼 키
        prev_year_quarter_key = f"{self.current_year - 1}_{self.current_quarter}Q"
        curr_quarter_key = f"{self.current_year}_{self.current_quarter}Q"
        
        prev_col_idx = config.get(prev_year_quarter_key)
        curr_col_idx = config.get(curr_quarter_key)
        
        if prev_col_idx is None or curr_col_idx is None:
            # 설정에 없으면 parse_sheet_structure로 시도
            structure = self.parse_sheet_structure(sheet_name)
            prev_key = f"{self.current_year - 1} {self.current_quarter}/4"
            curr_key = f"{self.current_year} {self.current_quarter}/4"
            prev_col_idx = prev_col_idx or structure['quarters'].get(prev_key)
            curr_col_idx = curr_col_idx or structure['quarters'].get(curr_key)
        
        if prev_col_idx is None or curr_col_idx is None:
            # 컬럼을 찾을 수 없으면 빈 결과 반환
            return result
        
        # 지역 컬럼과 분류 컬럼 인덱스
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        
        # 각 지역의 원지수 추출 (분류단계 0 = 총지수)
        for row_idx in range(len(df)):
            try:
                region = str(df.iloc[row_idx, region_col]).strip()
                if region not in self.ALL_REGIONS:
                    continue
                
                # 이미 해당 지역 데이터가 있으면 스킵 (첫 번째 행만 사용)
                if region in result:
                    continue
                
                # 분류단계 확인 (0 = 총지수)
                level_val = df.iloc[row_idx, level_col]
                if pd.notna(level_val):
                    level_str = str(level_val).strip()
                    if level_str not in ['0', '0.0', '총지수', '계']:
                        continue
                
                # 원지수 값 추출
                prev_value = df.iloc[row_idx, prev_col_idx]
                curr_value = df.iloc[row_idx, curr_col_idx]
                
                prev_index = round(float(prev_value), 1) if pd.notna(prev_value) else None
                curr_index = round(float(curr_value), 1) if pd.notna(curr_value) else None
                
                result[region] = [prev_index, curr_index]
            except (IndexError, ValueError, TypeError):
                continue
        
        return result
    
    def _generate_production_summary_table(self, sheet_name: str, config: Dict, 
                                          quarterly_growth: Dict, yearly_growth: Dict) -> Dict[str, Any]:
        """광공업생산/서비스업생산 요약 테이블 데이터 생성"""
        # 테이블용 분기 컬럼들: 4개 분기의 전년동기비 증감률
        def get_quarter_key(year, q):
            return f"{year}_{q}Q"
        
        table_q_pairs = []
        # 1. 전년동분기
        table_q_pairs.append((
            get_quarter_key(self.current_year - 1, self.current_quarter),
            get_quarter_key(self.current_year - 2, self.current_quarter),
            f"{self.current_year - 1}.{self.current_quarter}/4"
        ))
        # 2. 2분기 전
        q2 = self.current_quarter + 1 if self.current_quarter < 4 else 1
        y2 = self.current_year - 1 if self.current_quarter < 4 else self.current_year
        table_q_pairs.append((
            get_quarter_key(y2, q2),
            get_quarter_key(y2 - 1, q2),
            f"{y2}.{q2}/4"
        ))
        # 3. 직전 분기
        q3 = self.current_quarter - 1 if self.current_quarter > 1 else 4
        y3 = self.current_year if self.current_quarter > 1 else self.current_year - 1
        table_q_pairs.append((
            get_quarter_key(y3, q3),
            get_quarter_key(y3 - 1, q3),
            f"{y3}.{q3}/4"
        ))
        # 4. 현재 분기
        table_q_pairs.append((
            get_quarter_key(self.current_year, self.current_quarter),
            get_quarter_key(self.current_year - 1, self.current_quarter),
            f"{self.current_year}.{self.current_quarter}/4p"
        ))
        
        # 지역별 데이터 추출
        region_data = {}
        for _, _, label in table_q_pairs:
            quarter_data = quarterly_growth.get(label, {})
            for region in self.ALL_REGIONS:
                if region not in region_data:
                    region_data[region] = {'growth_rates': [], 'indices': []}
                rate = quarter_data.get(region, None)
                region_data[region]['growth_rates'].append(rate)
        
        # 원지수 데이터 추출: 전년동분기, 현재분기
        raw_indices = self._extract_raw_indices_for_table(sheet_name, config)
        for region in self.ALL_REGIONS:
            if region in region_data:
                region_data[region]['indices'] = raw_indices.get(region, [None, None])
        
        # 테이블 행 생성
        rows = []
        region_display = {
            '전국': '전 국', '서울': '서 울', '부산': '부 산', '대구': '대 구', '인천': '인 천',
            '광주': '광 주', '대전': '대 전', '울산': '울 산', '세종': '세 종', '경기': '경 기',
            '강원': '강 원', '충북': '충 북', '충남': '충 남', '전북': '전 북', '전남': '전 남',
            '경북': '경 북', '경남': '경 남', '제주': '제 주'
        }
        
        REGION_GROUPS = {
            "수도권": ["서울", "인천", "경기"],
            "동남권": ["부산", "울산", "경남"],
            "대경권": ["대구", "경북"],
            "호남권": ["광주", "전북", "전남"],
            "충청권": ["대전", "세종", "충북", "충남"],
            "강원제주": ["강원", "제주"]
        }
        
        # 전국 행
        national = region_data.get('전국', {'growth_rates': [None]*4, 'indices': [None, None]})
        rows.append({
            'region': '전 국',
            'group': None,
            'growth_rates': national['growth_rates'][:4],
            'indices': national['indices'][:2],
        })
        
        # 권역별 시도
        for group_name in ['수도권', '동남권', '대경권', '호남권', '충청권', '강원제주']:
            sidos = REGION_GROUPS[group_name]
            for idx, sido in enumerate(sidos):
                sido_data = region_data.get(sido, {'growth_rates': [None]*4, 'indices': [None, None]})
                
                row_data = {
                    'region': region_display.get(sido, sido),
                    'growth_rates': sido_data['growth_rates'][:4],
                    'indices': sido_data['indices'][:2],
                }
                
                if idx == 0:
                    row_data['group'] = group_name
                    row_data['rowspan'] = len(sidos)
                else:
                    row_data['group'] = None
                
                rows.append(row_data)
        
        return {
            'base_year': 2020,
            'columns': {
                'growth_rate_columns': [label for _, _, label in table_q_pairs],
                'index_columns': [
                    f"{self.current_year - 1}.{self.current_quarter}/4",  # 전년동분기
                    f"{self.current_year}.{self.current_quarter}/4p"  # 현재분기(잠정)
                ],
            },
            'regions': rows,
        }
    
    def extract_consumption_report_data(self) -> Dict[str, Any]:
        """소비동향 보도자료 데이터 추출 (소매판매액지수)
        
        소비동향 템플릿에서 필요한 데이터 구조:
        - summary_box: 요약 박스
        - nationwide_data: 전국 데이터 (sales_index, growth_rate, main_businesses)
        - regional_data: 지역별 데이터
        - top3_increase_regions, top3_decrease_regions
        - summary_table: 지역별 테이블 데이터
        """
        sheet_name = '소비(소매, 추가)'
        
        # 시트 설정 가져오기
        config = self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        name_col = config.get('name_col', 4)
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '소매판매액지수',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
            'summary_box': {},
            'nationwide_data': {},
            'summary_table': {},
            'increase_businesses_text': '',
            'decrease_businesses_text': '',
        }
        
        # 분류단계가 0인 합계 행만 추출
        quarterly_growth = self.extract_quarterly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=region_col,
            classification_column=level_col,
            classification_value='0'
        )
        
        yearly_growth = self.extract_yearly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=region_col,
            classification_column=level_col,
            classification_value='0'
        )
        
        current_quarter_key = f"{self.current_year}.{self.current_quarter}/4p"
        current_data = quarterly_growth.get(current_quarter_key, 
                       quarterly_growth.get(f"{self.current_year}.{self.current_quarter}/4", {}))
        
        national_rate = current_data.get('전국', None)
        direction = '증가' if national_rate > 0 else ('감소' if national_rate < 0 else '보합')
        
        report_data['national_summary'] = {
            'growth_rate': national_rate,
            'direction': direction,
        }
        
        # nationwide_data (템플릿 호환성)
        report_data['nationwide_data'] = {
            'sales_index': 100.0,  # 기준값
            'growth_rate': national_rate if national_rate is not None else 0.0,
            'main_businesses': [],  # 가중치 필요하여 빈 배열 (업태별 데이터)
        }
        
        # 지역별 데이터 추출
        regional_list = []
        for region in self.ALL_REGIONS:
            if region == '전국':
                continue
            rate = current_data.get(region, None)
            regional_list.append({
                'region': region,
                'growth_rate': rate,
                'direction': '증가' if rate is not None and rate > 0 else ('감소' if rate is not None and rate < 0 else ('보합' if rate == 0 else 'N/A')),
                'businesses': [],  # 가중치 필요
            })
        
        increase_regions = sorted([r for r in regional_list if r.get('growth_rate') is not None and r['growth_rate'] > 0], 
                                  key=lambda x: x['growth_rate'], reverse=True)
        decrease_regions = sorted([r for r in regional_list if r.get('growth_rate') is not None and r['growth_rate'] < 0], 
                                  key=lambda x: x['growth_rate'])
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        
        # summary_box (템플릿에서 사용)
        main_decrease_regions = []
        for r in decrease_regions[:3]:
            main_decrease_regions.append({
                'region': r['region'],
                'main_business': '',  # 가중치 필요
            })
        
        report_data['summary_box'] = {
            'main_decrease_regions': main_decrease_regions,
            'main_increase_regions': increase_regions[:3],
            'region_count': len(decrease_regions),
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
        }
        
        # summary_table (지역별 테이블 데이터)
        report_data['summary_table'] = self._generate_consumption_summary_table(sheet_name, config)
        
        report_data['yearly_data'] = yearly_growth
        report_data['quarterly_data'] = quarterly_growth
        
        return report_data
    
    def _generate_consumption_summary_table(self, sheet_name: str, config: Dict) -> Dict[str, Any]:
        """소비동향 요약 테이블 데이터 생성"""
        df = self._load_sheet(sheet_name)
        if df is None:
            return self._get_empty_consumption_table()
        
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        
        # 컬럼 정보
        current_q_key = f"{self.current_year}_{self.current_quarter}Q"
        prev_year_q_key = f"{self.current_year - 1}_{self.current_quarter}Q"
        
        # 테이블용 분기 컬럼들: 4개 분기의 전년동기비 증감률 계산
        def get_quarter_key(year, q):
            return f"{year}_{q}Q"
        
        table_q_pairs = []
        # 1. 전년동분기 (2024.2/4)
        table_q_pairs.append((
            get_quarter_key(self.current_year - 1, self.current_quarter),
            get_quarter_key(self.current_year - 2, self.current_quarter),
            f"{self.current_year - 1}.{self.current_quarter}/4"
        ))
        # 2. 2분기 전
        q2 = self.current_quarter + 1 if self.current_quarter < 4 else 1
        y2 = self.current_year - 1 if self.current_quarter < 4 else self.current_year
        table_q_pairs.append((
            get_quarter_key(y2, q2),
            get_quarter_key(y2 - 1, q2),
            f"{y2}.{q2}/4"
        ))
        # 3. 직전 분기
        q3 = self.current_quarter - 1 if self.current_quarter > 1 else 4
        y3 = self.current_year if self.current_quarter > 1 else self.current_year - 1
        table_q_pairs.append((
            get_quarter_key(y3, q3),
            get_quarter_key(y3 - 1, q3),
            f"{y3}.{q3}/4"
        ))
        # 4. 현재 분기
        table_q_pairs.append((
            current_q_key,
            prev_year_q_key,
            f"{self.current_year}.{self.current_quarter}/4p"
        ))
        
        # 지역별 데이터 추출
        region_data = {}
        for i in range(3, len(df)):
            row = df.iloc[i]
            region = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
            level = str(row[level_col]).strip() if pd.notna(row[level_col]) else ''
            
            if region not in self.ALL_REGIONS or level != '0':
                continue
            
            if region not in region_data:
                region_data[region] = {
                    'growth_rates': [],
                    'indices': [],
                }
            
            # 4개 분기 증감률 계산
            for curr_key, prev_key, _ in table_q_pairs:
                curr_col = config.get(curr_key)
                prev_col = config.get(prev_key)
                
                if curr_col and prev_col:
                    try:
                        curr_v = float(row.iloc[curr_col]) if pd.notna(row.iloc[curr_col]) else 0.0
                        prev_v = float(row.iloc[prev_col]) if pd.notna(row.iloc[prev_col]) else 0.0
                        if prev_v != 0:
                            change = ((curr_v - prev_v) / prev_v) * 100
                        else:
                            change = 0.0
                        region_data[region]['growth_rates'].append(round(change, 1))
                    except:
                        region_data[region]['growth_rates'].append(None)
                else:
                    region_data[region]['growth_rates'].append(None)
            
            # 지수 데이터 (직전 분기, 현재 분기)
            for q_key in [table_q_pairs[2][0], table_q_pairs[3][0]]:
                col = config.get(q_key)
                if col:
                    try:
                        idx_val = float(row.iloc[col]) if pd.notna(row.iloc[col]) else None
                        region_data[region]['indices'].append(idx_val)
                    except:
                        region_data[region]['indices'].append(None)
                else:
                    region_data[region]['indices'].append(None)
        
        # 테이블 행 생성
        rows = []
        region_display = {
            '전국': '전 국', '서울': '서 울', '부산': '부 산', '대구': '대 구', '인천': '인 천',
            '광주': '광 주', '대전': '대 전', '울산': '울 산', '세종': '세 종', '경기': '경 기',
            '강원': '강 원', '충북': '충 북', '충남': '충 남', '전북': '전 북', '전남': '전 남',
            '경북': '경 북', '경남': '경 남', '제주': '제 주'
        }
        
        REGION_GROUPS = {
            "수도권": ["서울", "인천", "경기"],
            "동남권": ["부산", "울산", "경남"],
            "대경권": ["대구", "경북"],
            "호남권": ["광주", "전북", "전남"],
            "충청권": ["대전", "세종", "충북", "충남"],
            "강원제주": ["강원", "제주"]
        }
        
        # 전국 행
        national = region_data.get('전국', {'growth_rates': [None]*4, 'indices': [None]*2})
        rows.append({
            'region': '전 국',
            'group': None,
            'growth_rates': national['growth_rates'][:4],
            'indices': national['indices'][:2],
        })
        
        # 권역별 시도
        for group_name in ['수도권', '동남권', '대경권', '호남권', '충청권', '강원제주']:
            sidos = REGION_GROUPS[group_name]
            for idx, sido in enumerate(sidos):
                sido_data = region_data.get(sido, {'growth_rates': [None]*4, 'indices': [None]*2})
                
                row_data = {
                    'region': region_display.get(sido, sido),
                    'growth_rates': sido_data['growth_rates'][:4],
                    'indices': sido_data['indices'][:2],
                }
                
                if idx == 0:
                    row_data['group'] = group_name
                    row_data['rowspan'] = len(sidos)
                else:
                    row_data['group'] = None
                
                rows.append(row_data)
        
        return {
            'base_year': 2020,
            'columns': {
                'growth_rate_columns': [label for _, _, label in table_q_pairs],
                'index_columns': [
                    f"{self.current_year}.{self.current_quarter - 1 if self.current_quarter > 1 else 4}/4",
                    f"{self.current_year}.{self.current_quarter}/4"
                ],
            },
            'regions': rows,
        }
    
    def _get_empty_consumption_table(self) -> Dict[str, Any]:
        """빈 소비동향 테이블 반환"""
        return {
            'base_year': 2020,
            'columns': {
                'growth_rate_columns': [],
                'index_columns': [],
            },
            'regions': [],
        }
    
    def extract_construction_report_data(self) -> Dict[str, Any]:
        """건설동향 보도자료 데이터 추출"""
        sheet_name = '건설 (공표자료)'
        
        # 시트 설정 가져오기
        config = self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '건설수주',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
            'summary_box': {},
            'nationwide_data': {},
        }
        
        # 분류단계가 0인 합계 행만 추출
        quarterly_growth = self.extract_quarterly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=region_col,
            classification_column=level_col,
            classification_value='0'
        )
        
        yearly_growth = self.extract_yearly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=region_col,
            classification_column=level_col,
            classification_value='0'
        )
        
        current_quarter_key = f"{self.current_year}.{self.current_quarter}/4p"
        current_data = quarterly_growth.get(current_quarter_key, 
                       quarterly_growth.get(f"{self.current_year}.{self.current_quarter}/4", {}))
        
        national_rate = current_data.get('전국', None)
        direction = '증가' if national_rate > 0 else ('감소' if national_rate < 0 else '보합')
        
        report_data['national_summary'] = {
            'growth_rate': national_rate,
            'direction': direction,
        }
        
        # nationwide_data (템플릿 호환성)
        report_data['nationwide_data'] = {
            'amount': 0.0,  # 금액은 별도 계산 필요
            'growth_rate': national_rate if national_rate is not None else 0.0,
            'main_categories': [],  # 가중치 필요
        }
        
        regional_list = []
        for region in self.ALL_REGIONS:
            if region == '전국':
                continue
            rate = current_data.get(region, None)
            regional_list.append({
                'region': region,
                'growth_rate': rate,
                'direction': '증가' if rate is not None and rate > 0 else ('감소' if rate is not None and rate < 0 else ('보합' if rate == 0 else 'N/A')),
                'categories': [],  # 가중치 필요
            })
        
        increase_regions = sorted([r for r in regional_list if r.get('growth_rate') is not None and r['growth_rate'] > 0], 
                                  key=lambda x: x['growth_rate'], reverse=True)
        decrease_regions = sorted([r for r in regional_list if r.get('growth_rate') is not None and r['growth_rate'] < 0], 
                                  key=lambda x: x['growth_rate'])
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        
        # summary_box (템플릿에서 사용)
        report_data['summary_box'] = {
            'main_increase_regions': increase_regions[:3],
            'main_decrease_regions': decrease_regions[:3],
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
        }
        
        report_data['yearly_data'] = yearly_growth
        report_data['quarterly_data'] = quarterly_growth
        
        return report_data
    
    def extract_export_report_data(self) -> Dict[str, Any]:
        """수출 보도자료 데이터 추출"""
        sheet_name = '수출'
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '수출액',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
        }
        
        # 수출 시트: level='0'이 합계 행
        quarterly_growth = self.extract_quarterly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=1,
            classification_column=2,
            classification_value='0'
        )
        
        yearly_growth = self.extract_yearly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=1,
            classification_column=2,
            classification_value='0'
        )
        
        current_quarter_key = f"{self.current_year}.{self.current_quarter}/4p"
        current_data = quarterly_growth.get(current_quarter_key, 
                       quarterly_growth.get(f"{self.current_year}.{self.current_quarter}/4", {}))
        
        national_rate = current_data.get('전국', None)
        report_data['national_summary'] = {
            'growth_rate': national_rate,
            'direction': '증가' if national_rate is not None and national_rate > 0 else ('감소' if national_rate is not None and national_rate < 0 else ('보합' if national_rate == 0 else 'N/A')),
        }
        
        regional_list = []
        for region in self.ALL_REGIONS:
            if region == '전국':
                continue
            rate = current_data.get(region, None)
            regional_list.append({
                'region': region,
                'growth_rate': rate,
                'direction': '증가' if rate is not None and rate > 0 else ('감소' if rate is not None and rate < 0 else ('보합' if rate == 0 else 'N/A')),
            })
        
        increase_regions = sorted([r for r in regional_list if r['growth_rate'] > 0], 
                                  key=lambda x: x['growth_rate'], reverse=True)
        decrease_regions = sorted([r for r in regional_list if r['growth_rate'] < 0], 
                                  key=lambda x: x['growth_rate'])
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        report_data['yearly_data'] = yearly_growth
        report_data['quarterly_data'] = quarterly_growth
        
        # summary_table 생성 (수입과 동일한 구조)
        report_data['summary_table'] = self._generate_export_import_summary_table(sheet_name, quarterly_growth)
        
        return report_data
    
    def _extract_raw_amounts_for_table(self, sheet_name: str, config: Dict) -> Dict[str, List[float]]:
        """수출/수입 요약 테이블용 금액 데이터 추출 (전년동분기, 현재분기)
        
        Args:
            sheet_name: 시트 이름 (수출 또는 수입)
            config: RAW_SHEET_QUARTER_COLS에서 가져온 설정
            
        Returns:
            {
                '전국': [전년동분기 금액(억달러), 현재분기 금액(억달러)],
                '서울': [...],
                ...
            }
        """
        result = {}
        
        df = self._load_sheet(sheet_name)
        if df is None:
            return result
        
        # 전년동분기와 현재분기 컬럼 키
        prev_year_quarter_key = f"{self.current_year - 1}_{self.current_quarter}Q"
        curr_quarter_key = f"{self.current_year}_{self.current_quarter}Q"
        
        prev_col_idx = config.get(prev_year_quarter_key)
        curr_col_idx = config.get(curr_quarter_key)
        
        if prev_col_idx is None or curr_col_idx is None:
            return result
        
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        
        # 각 지역의 금액 추출 (분류단계 = 0, 합계 행)
        for row_idx in range(len(df)):
            try:
                region = str(df.iloc[row_idx, region_col]).strip()
                if region not in self.ALL_REGIONS:
                    continue
                
                # 이미 해당 지역 데이터가 있으면 스킵
                if region in result:
                    continue
                
                # 분류단계 확인 (0 = 합계)
                level_val = df.iloc[row_idx, level_col]
                if pd.notna(level_val):
                    level_str = str(level_val).strip()
                    if level_str not in ['0', '0.0']:
                        continue
                
                # 금액 추출 (백만달러 → 억달러 변환: /100)
                prev_value = df.iloc[row_idx, prev_col_idx]
                curr_value = df.iloc[row_idx, curr_col_idx]
                
                prev_amount = round(float(prev_value) / 100, 0) if pd.notna(prev_value) else None
                curr_amount = round(float(curr_value) / 100, 0) if pd.notna(curr_value) else None
                
                result[region] = [prev_amount, curr_amount]
            except (IndexError, ValueError, TypeError):
                continue
        
        return result
    
    def _generate_export_import_summary_table(self, sheet_name: str, quarterly_growth: Dict) -> Dict[str, Any]:
        """수출/수입 요약 테이블 데이터 생성"""
        # 테이블용 분기 컬럼들: 4개 분기의 전년동기비 증감률
        def get_quarter_key(year, q):
            return f"{year}_{q}Q"
        
        table_q_pairs = []
        # 1. 전년동분기
        table_q_pairs.append((
            get_quarter_key(self.current_year - 1, self.current_quarter),
            get_quarter_key(self.current_year - 2, self.current_quarter),
            f"{self.current_year - 1}.{self.current_quarter}/4"
        ))
        # 2. 2분기 전
        q2 = self.current_quarter + 1 if self.current_quarter < 4 else 1
        y2 = self.current_year - 1 if self.current_quarter < 4 else self.current_year
        table_q_pairs.append((
            get_quarter_key(y2, q2),
            get_quarter_key(y2 - 1, q2),
            f"{y2}.{q2}/4"
        ))
        # 3. 직전 분기
        q3 = self.current_quarter - 1 if self.current_quarter > 1 else 4
        y3 = self.current_year if self.current_quarter > 1 else self.current_year - 1
        table_q_pairs.append((
            get_quarter_key(y3, q3),
            get_quarter_key(y3 - 1, q3),
            f"{y3}.{q3}/4"
        ))
        # 4. 현재 분기
        table_q_pairs.append((
            get_quarter_key(self.current_year, self.current_quarter),
            get_quarter_key(self.current_year - 1, self.current_quarter),
            f"{self.current_year}.{self.current_quarter}/4p"
        ))
        
        # 지역별 데이터 추출
        region_data = {}
        for _, _, label in table_q_pairs:
            quarter_data = quarterly_growth.get(label, {})
            for region in self.ALL_REGIONS:
                if region not in region_data:
                    region_data[region] = {'changes': [], 'amounts': []}
                rate = quarter_data.get(region, None)
                region_data[region]['changes'].append(rate)
        
        # 수출/수입 금액 추출 (전년동분기, 현재분기)
        config = self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        raw_amounts = self._extract_raw_amounts_for_table(sheet_name, config)
        for region in self.ALL_REGIONS:
            if region in region_data:
                region_data[region]['amounts'] = raw_amounts.get(region, [None, None])
        
        # 테이블 행 생성
        rows = []
        region_display = {
            '전국': '전 국', '서울': '서 울', '부산': '부 산', '대구': '대 구', '인천': '인 천',
            '광주': '광 주', '대전': '대 전', '울산': '울 산', '세종': '세 종', '경기': '경 기',
            '강원': '강 원', '충북': '충 북', '충남': '충 남', '전북': '전 북', '전남': '전 남',
            '경북': '경 북', '경남': '경 남', '제주': '제 주'
        }
        
        REGION_GROUPS = {
            "경인": ["서울", "인천", "경기"],
            "충청": ["대전", "세종", "충북", "충남"],
            "호남": ["광주", "전북", "전남", "제주"],
            "동북": ["대구", "경북", "강원"],
            "동남": ["부산", "울산", "경남"]
        }
        
        # 전국 행
        national = region_data.get('전국', {'changes': [None]*4, 'amounts': [None, None]})
        rows.append({
            'region_group': None,
            'sido': '전 국',
            'changes': national['changes'][:4],
            'amounts': national['amounts'][:2],
        })
        
        # 권역별 시도
        for group_name in ['경인', '충청', '호남', '동북', '동남']:
            sidos = REGION_GROUPS[group_name]
            for idx, sido in enumerate(sidos):
                sido_data = region_data.get(sido, {'changes': [None]*4, 'amounts': [None, None]})
                
                row_data = {
                    'sido': region_display.get(sido, sido),
                    'changes': sido_data['changes'][:4],
                    'amounts': sido_data['amounts'][:2],
                }
                
                if idx == 0:
                    row_data['region_group'] = group_name
                    row_data['rowspan'] = len(sidos)
                else:
                    row_data['region_group'] = None
                
                rows.append(row_data)
        
        return {
            'rows': rows,
            'columns': {
                'growth_rate_columns': [label for _, _, label in table_q_pairs],
                'amount_columns': [
                    f"{self.current_year - 1}.{self.current_quarter}/4",
                    f"{self.current_year}.{self.current_quarter}/4"
                ],
            },
        }
    
    def extract_import_report_data(self) -> Dict[str, Any]:
        """수입 보도자료 데이터 추출 (품목, 수입액, summary_table 포함)"""
        sheet_name = '수입'
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '수입액',
            },
            'national_summary': {},
            'regional_data': {},
            'nationwide_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
            'summary_box': {},
            'summary_table': {},
            'increase_products_text': '',
            'decrease_products_text': '',
        }
        
        # 시트 로드
        df = self._load_sheet(sheet_name)
        if df is None:
            print(f"[WARNING] 수입 시트를 찾을 수 없습니다")
            return report_data
        
        # 시트 구조 정보
        config = self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        name_col = config.get('name_col', 5)
        
        # 현재 분기와 전년동분기 컬럼 찾기
        current_q_key = f"{self.current_year}_{self.current_quarter}Q"
        prev_year_q_key = f"{self.current_year - 1}_{self.current_quarter}Q"
        
        current_col = config.get(current_q_key)
        prev_col = config.get(prev_year_q_key)
        
        if current_col is None or prev_col is None:
            print(f"[WARNING] 분기 컬럼을 찾을 수 없습니다: {current_q_key}, {prev_year_q_key}")
            return report_data
        
        # 테이블용 분기 컬럼들: 4개 분기의 전년동기비 증감률 계산
        # 2025년 2분기 기준: 2024.2/4, 2025.3/4, 2025.1/4, 2025.2/4p
        # 각각의 전년동기비를 계산
        def get_quarter_key(year, q):
            return f"{year}_{q}Q"
        
        # 컬럼 순서: 전전년동분기, 전분기(3/4), 직전분기(1/4), 현재분기
        # 하지만 실제로는 config에 있는 분기만 사용 가능
        table_q_pairs = []
        
        # 1. 전년동분기 (2024.2/4)
        table_q_pairs.append((
            get_quarter_key(self.current_year - 1, self.current_quarter),
            get_quarter_key(self.current_year - 2, self.current_quarter),
            f"{self.current_year - 1}.{self.current_quarter}/4"
        ))
        
        # 2. 2분기 전 (2024.3/4 또는 2024.4/4)
        q2 = self.current_quarter + 1 if self.current_quarter < 4 else 1
        y2 = self.current_year - 1 if self.current_quarter < 4 else self.current_year
        table_q_pairs.append((
            get_quarter_key(y2, q2),
            get_quarter_key(y2 - 1, q2),
            f"{y2}.{q2}/4"
        ))
        
        # 3. 직전 분기 (2025.1/4)
        q3 = self.current_quarter - 1 if self.current_quarter > 1 else 4
        y3 = self.current_year if self.current_quarter > 1 else self.current_year - 1
        table_q_pairs.append((
            get_quarter_key(y3, q3),
            get_quarter_key(y3 - 1, q3),
            f"{y3}.{q3}/4"
        ))
        
        # 4. 현재 분기 (2025.2/4p)
        table_q_pairs.append((
            current_q_key,
            prev_year_q_key,
            f"{self.current_year}.{self.current_quarter}/4p"
        ))
        
        # 지역별, 품목별 데이터 추출
        region_data = {}  # {region: {total_curr, total_prev, products: [...]}}
        
        for i in range(3, len(df)):
            row = df.iloc[i]
            region = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
            level = str(row[level_col]).strip() if pd.notna(row[level_col]) else ''
            product_name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ''
            
            if region not in self.ALL_REGIONS:
                continue
            
            # 현재 분기와 전년동분기 값 가져오기
            try:
                curr_val = float(row[current_col]) if pd.notna(row[current_col]) else 0.0
                prev_val = float(row[prev_col]) if pd.notna(row[prev_col]) else 0.0
            except (ValueError, TypeError):
                curr_val, prev_val = 0.0, 0.0
            
            if region not in region_data:
                region_data[region] = {
                    'total_curr': 0.0,
                    'total_prev': 0.0,
                    'products': [],
                    'table_changes': [],
                    'table_amounts': [],
                }
            
            # 합계 행 (level='0')
            if level == '0':
                region_data[region]['total_curr'] = curr_val
                region_data[region]['total_prev'] = prev_val
                
                # 테이블용 증감률 데이터 추출
                for curr_key, prev_key, _ in table_q_pairs:
                    curr_col_idx = config.get(curr_key)
                    prev_col_idx = config.get(prev_key)
                    
                    if curr_col_idx and prev_col_idx:
                        try:
                            curr_v = float(row[curr_col_idx]) if pd.notna(row[curr_col_idx]) else 0.0
                            prev_v = float(row[prev_col_idx]) if pd.notna(row[prev_col_idx]) else 0.0
                            if prev_v != 0:
                                change = ((curr_v - prev_v) / prev_v) * 100
                            else:
                                change = 0.0
                            region_data[region]['table_changes'].append(round(change, 1))
                        except:
                            region_data[region]['table_changes'].append(None)
                    else:
                        region_data[region]['table_changes'].append(None)
                
                # 수입액 (백만달러 -> 억달러)
                region_data[region]['table_amounts'] = [
                    round(prev_val / 10, 1) if prev_val else None,
                    round(curr_val / 10, 1) if curr_val else None,
                ]
            
            # 품목 데이터 (level='2')
            elif level == '2' and product_name:
                if prev_val != 0:
                    change = ((curr_val - prev_val) / prev_val) * 100
                else:
                    change = 0.0 if curr_val == 0 else 100.0
                
                # 기여도 = 금액 변화량 (정렬용, 나중에 %p로 변환)
                amount_change = curr_val - prev_val
                
                region_data[region]['products'].append({
                    'name': product_name,
                    'change': round(change, 1),
                    'contribution': amount_change,  # 정렬용 금액 변화량
                    'amount_change': amount_change,  # 원본 금액 변화량 (백만달러)
                })
        
        # 전국 데이터 추출
        national_info = region_data.get('전국', {})
        total_curr = national_info.get('total_curr', 0.0)
        total_prev = national_info.get('total_prev', 0.0)
        
        if total_prev != 0:
            national_rate = ((total_curr - total_prev) / total_prev) * 100
        else:
            national_rate = 0.0
        
        # 전국 품목에 기여도(%p) 계산: (품목 금액 변화량 / 전체 전년동분기 금액) × 100
        national_products = national_info.get('products', [])
        for p in national_products:
            if total_prev != 0:
                p['contribution_pct'] = (p['amount_change'] / total_prev) * 100
            else:
                p['contribution_pct'] = 0.0
        
        # 전국 품목 순위 정렬 (금액 변화량 기준)
        positive_products = sorted([p for p in national_products if p['contribution'] > 0],
                                  key=lambda x: -x['contribution'])[:5]
        negative_products = sorted([p for p in national_products if p['contribution'] < 0],
                                  key=lambda x: x['contribution'])[:5]
        
        # nationwide_data 설정
        report_data['nationwide_data'] = {
            'amount': round(total_curr / 10, 1) if total_curr else 0.0,  # 억달러
            'change': round(national_rate, 1),
            'products': positive_products[:3] if national_rate >= 0 else negative_products[:3],
            'increase_products': positive_products[:3],
            'decrease_products': negative_products[:3],
        }
        
        report_data['national_summary'] = {
            'growth_rate': round(national_rate, 1),
            'direction': '증가' if national_rate is not None and national_rate > 0 else ('감소' if national_rate is not None and national_rate < 0 else ('보합' if national_rate == 0 else 'N/A')),
        }
        
        # 지역별 데이터 정리
        regional_list = []
        for region in self.ALL_REGIONS:
            if region == '전국':
                continue
            
            info = region_data.get(region, {})
            curr = info.get('total_curr', 0.0)
            prev = info.get('total_prev', 0.0)
            
            if prev != 0:
                rate = ((curr - prev) / prev) * 100
            else:
                rate = 0.0
            
            # 지역 품목에 기여도(%p) 계산: (품목 금액 변화량 / 지역 전년동분기 금액) × 100
            products = info.get('products', [])
            for p in products:
                if prev != 0:
                    p['contribution_pct'] = (p['amount_change'] / prev) * 100
                else:
                    p['contribution_pct'] = 0.0
            
            # 지역 품목 순위 정렬 (금액 변화량 기준)
            positive = sorted([p for p in products if p['contribution'] > 0],
                            key=lambda x: -x['contribution'])[:3]
            negative = sorted([p for p in products if p['contribution'] < 0],
                            key=lambda x: x['contribution'])[:3]
            
            regional_list.append({
                'region': region,
                'name': region,
                'growth_rate': round(rate, 1),
                'change': round(rate, 1),
                'direction': '증가' if rate is not None and rate > 0 else ('감소' if rate is not None and rate < 0 else ('보합' if rate == 0 else 'N/A')),
                'products': positive if rate >= 0 else negative,
                'increase_products': positive,
                'decrease_products': negative,
            })
        
        increase_regions = sorted([r for r in regional_list if r['growth_rate'] > 0],
                                  key=lambda x: x['growth_rate'], reverse=True)
        decrease_regions = sorted([r for r in regional_list if r['growth_rate'] < 0],
                                  key=lambda x: x['growth_rate'])
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        
        # 품목 텍스트 생성
        increase_prods = []
        for r in increase_regions[:3]:
            prods = r.get('products', [])
            if prods:
                increase_prods.append(prods[0]['name'])
        report_data['increase_products_text'] = ', '.join(increase_prods[:3]) if increase_prods else ''
        
        decrease_prods = []
        for r in decrease_regions[:3]:
            prods = r.get('products', [])
            if prods:
                decrease_prods.append(prods[0]['name'])
        report_data['decrease_products_text'] = ', '.join(decrease_prods[:3]) if decrease_prods else ''
        
        # summary_box 생성
        main_decrease_regions = []
        for r in decrease_regions[:3]:
            products = r.get('products', [])
            product_names = [p['name'] for p in products[:2]] if products else []
            main_decrease_regions.append({
                'region': r['region'],
                'products': product_names,
            })
        
        report_data['summary_box'] = {
            'main_increase_regions': increase_regions[:3],
            'main_decrease_regions': main_decrease_regions,
            'increase_count': len(increase_regions),
            'decrease_count': len(decrease_regions),
        }
        
        # summary_table 생성
        rows = []
        region_display = {
            '전국': '전 국', '서울': '서 울', '부산': '부 산', '대구': '대 구', '인천': '인 천',
            '광주': '광 주', '대전': '대 전', '울산': '울 산', '세종': '세 종', '경기': '경 기',
            '강원': '강 원', '충북': '충 북', '충남': '충 남', '전북': '전 북', '전남': '전 남',
            '경북': '경 북', '경남': '경 남', '제주': '제 주'
        }
        
        REGION_GROUPS = {
            "경인": ["서울", "인천", "경기"],
            "충청": ["대전", "세종", "충북", "충남"],
            "호남": ["광주", "전북", "전남", "제주"],
            "동북": ["대구", "경북", "강원"],
            "동남": ["부산", "울산", "경남"]
        }
        
        # 전국 행
        national_table = region_data.get('전국', {})
        rows.append({
            'region_group': None,
            'sido': '전 국',
            'changes': national_table.get('table_changes', [None, None, None, None])[:4],
            'amounts': national_table.get('table_amounts', [None, None])[:2],
        })
        
        # 권역별 시도
        for group_name in ['경인', '충청', '호남', '동북', '동남']:
            sidos = REGION_GROUPS[group_name]
            for idx, sido in enumerate(sidos):
                sido_data = region_data.get(sido, {})
                
                # 증감률 데이터
                changes = sido_data.get('table_changes', [])
                while len(changes) < 4:
                    changes.append(None)
                
                # 수입액
                amounts = sido_data.get('table_amounts', [])
                while len(amounts) < 2:
                    amounts.append(None)
                
                row_data = {
                    'sido': region_display.get(sido, sido),
                    'changes': changes[:4],
                    'amounts': amounts[:2],
                }
                
                if idx == 0:
                    row_data['region_group'] = group_name
                    row_data['rowspan'] = len(sidos)
                else:
                    row_data['region_group'] = None
                
                rows.append(row_data)
        
        # 테이블 컬럼 정보
        columns = {
            'growth_rate_columns': [label for _, _, label in table_q_pairs],
            'amount_columns': [
                f"{self.current_year - 1}.{self.current_quarter}/4",
                f"{self.current_year}.{self.current_quarter}/4"
            ],
        }
        
        report_data['summary_table'] = {
            'rows': rows,
            'columns': columns,
        }
        
        return report_data
    
    def extract_price_report_data(self) -> Dict[str, Any]:
        """물가동향 보도자료 데이터 추출"""
        sheet_name = '품목성질별 물가'
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '소비자물가지수',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
        }
        
        # 물가 시트: level='0'이 합계 행 (name='총지수')
        quarterly_growth = self.extract_quarterly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=0,
            classification_column=1,
            classification_value='0'
        )
        
        yearly_growth = self.extract_yearly_growth_rate(
            sheet_name,
            start_year=2020,
            region_column=0,
            classification_column=1,
            classification_value='0'
        )
        
        current_quarter_key = f"{self.current_year}.{self.current_quarter}/4p"
        current_data = quarterly_growth.get(current_quarter_key, 
                       quarterly_growth.get(f"{self.current_year}.{self.current_quarter}/4", {}))
        
        national_rate = current_data.get('전국', None)
        report_data['national_summary'] = {
            'growth_rate': national_rate,
            'direction': '상승' if national_rate is not None and national_rate > 0 else ('하락' if national_rate is not None and national_rate < 0 else ('보합' if national_rate == 0 else 'N/A')),
        }
        
        regional_list = []
        for region in self.ALL_REGIONS:
            if region == '전국':
                continue
            rate = current_data.get(region, None)
            regional_list.append({
                'region': region,
                'growth_rate': rate,
                'direction': '상승' if rate is not None and rate > 0 else ('하락' if rate is not None and rate < 0 else ('보합' if rate == 0 else 'N/A')),
            })
        
        increase_regions = sorted([r for r in regional_list if r['growth_rate'] > 0], 
                                  key=lambda x: x['growth_rate'], reverse=True)
        decrease_regions = sorted([r for r in regional_list if r['growth_rate'] < 0], 
                                  key=lambda x: x['growth_rate'])
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        report_data['yearly_data'] = yearly_growth
        report_data['quarterly_data'] = quarterly_growth
        
        # summary_table 생성
        report_data['summary_table'] = self._generate_price_summary_table(quarterly_growth)
        
        return report_data
    
    def _extract_price_indices_for_table(self, sheet_name: str, config: Dict) -> Dict[str, List[float]]:
        """물가동향 요약 테이블용 물가지수 추출 (전년동분기, 현재분기)
        
        Args:
            sheet_name: 시트 이름
            config: RAW_SHEET_QUARTER_COLS에서 가져온 설정
            
        Returns:
            {
                '전국': [전년동분기 지수, 현재분기 지수],
                '서울': [...],
                ...
            }
        """
        result = {}
        
        df = self._load_sheet(sheet_name)
        if df is None:
            return result
        
        # 전년동분기와 현재분기 컬럼 키
        prev_year_quarter_key = f"{self.current_year - 1}_{self.current_quarter}Q"
        curr_quarter_key = f"{self.current_year}_{self.current_quarter}Q"
        
        prev_col_idx = config.get(prev_year_quarter_key)
        curr_col_idx = config.get(curr_quarter_key)
        
        if prev_col_idx is None or curr_col_idx is None:
            return result
        
        region_col = config.get('region_col', 0)
        level_col = config.get('level_col', 1)
        total_code = config.get('total_code', '총지수')
        
        # 각 지역의 물가지수 추출 (분류단계 = 0, 총지수 행)
        for row_idx in range(len(df)):
            try:
                region = str(df.iloc[row_idx, region_col]).strip()
                if region not in self.ALL_REGIONS:
                    continue
                
                # 이미 해당 지역 데이터가 있으면 스킵
                if region in result:
                    continue
                
                # 분류단계 확인 (0 = 총지수)
                level_val = df.iloc[row_idx, level_col]
                if pd.notna(level_val):
                    # 숫자 0 또는 문자열 '0' 모두 처리
                    if level_val != 0 and str(level_val).strip() not in ['0', '0.0']:
                        continue
                
                # 지수 추출
                prev_value = df.iloc[row_idx, prev_col_idx]
                curr_value = df.iloc[row_idx, curr_col_idx]
                
                prev_index = round(float(prev_value), 1) if pd.notna(prev_value) else None
                curr_index = round(float(curr_value), 1) if pd.notna(curr_value) else None
                
                result[region] = [prev_index, curr_index]
            except (IndexError, ValueError, TypeError):
                continue
        
        return result
    
    def _generate_price_summary_table(self, quarterly_growth: Dict) -> Dict[str, Any]:
        """물가동향 요약 테이블 데이터 생성"""
        # 테이블용 분기 컬럼들
        table_q_pairs = [
            f"{self.current_year - 1}.{self.current_quarter}/4",  # 2024.2/4
            f"{self.current_year}.{self.current_quarter + 1 if self.current_quarter < 4 else 1}/4" if self.current_quarter < 4 else f"{self.current_year}.1/4",  # 2024.3/4 or 2025.1/4
            f"{self.current_year}.{self.current_quarter - 1 if self.current_quarter > 1 else 4}/4",  # 2025.1/4
            f"{self.current_year}.{self.current_quarter}/4p",  # 2025.2/4p
        ]
        
        # 지역별 데이터 추출
        region_data = {}
        for quarter_label in table_q_pairs:
            quarter_data = quarterly_growth.get(quarter_label, {})
            for region in self.ALL_REGIONS:
                if region not in region_data:
                    region_data[region] = {'changes': [], 'indices': []}
                rate = quarter_data.get(region, None)
                region_data[region]['changes'].append(rate)
        
        # 물가지수 추출 (전년동분기, 현재분기)
        sheet_name = '품목성질별 물가'
        config = self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        raw_indices = self._extract_price_indices_for_table(sheet_name, config)
        for region in self.ALL_REGIONS:
            if region in region_data:
                region_data[region]['indices'] = raw_indices.get(region, [None, None])
        
        # 테이블 행 생성
        rows = []
        region_display = {
            '전국': '전 국', '서울': '서 울', '부산': '부 산', '대구': '대 구', '인천': '인 천',
            '광주': '광 주', '대전': '대 전', '울산': '울 산', '세종': '세 종', '경기': '경 기',
            '강원': '강 원', '충북': '충 북', '충남': '충 남', '전북': '전 북', '전남': '전 남',
            '경북': '경 북', '경남': '경 남', '제주': '제 주'
        }
        
        REGION_GROUPS = {
            "수도권": ["서울", "인천", "경기"],
            "동남권": ["부산", "울산", "경남"],
            "대경권": ["대구", "경북"],
            "호남권": ["광주", "전북", "전남"],
            "충청권": ["대전", "세종", "충북", "충남"],
            "강원제주": ["강원", "제주"]
        }
        
        # 전국 행
        national = region_data.get('전국', {'changes': [None]*4, 'indices': [None, None]})
        rows.append({
            'region': '전 국',
            'group': None,
            'changes': national['changes'][:4],
            'indices': national['indices'][:2],
        })
        
        # 권역별 시도
        for group_name in ['수도권', '동남권', '대경권', '호남권', '충청권', '강원제주']:
            sidos = REGION_GROUPS[group_name]
            for idx, sido in enumerate(sidos):
                sido_data = region_data.get(sido, {'changes': [None]*4, 'indices': [None, None]})
                
                row_data = {
                    'region': region_display.get(sido, sido),
                    'changes': sido_data['changes'][:4],
                    'indices': sido_data['indices'][:2],
                }
                
                if idx == 0:
                    row_data['group'] = group_name
                    row_data['rowspan'] = len(sidos)
                else:
                    row_data['group'] = None
                
                rows.append(row_data)
        
        return {
            'base_year': 2020,
            'columns': {
                'change_columns': table_q_pairs,
                'index_columns': [
                    f"{self.current_year - 1}.{self.current_quarter}/4",
                    f"{self.current_year}.{self.current_quarter}/4"
                ],
            },
            'rows': rows,
        }
    
    def extract_employment_rate_report_data(self) -> Dict[str, Any]:
        """고용률 보도자료 데이터 추출 (전년동기비 %p 차이)"""
        sheet_name = '고용률'
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '고용률',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
        }
        
        # 고용률 시트: region_column=1, level='0'이 합계 행
        quarterly_diff = self.extract_quarterly_difference(
            sheet_name,
            start_year=2020,
            region_column=1,
            classification_column=2,
            classification_value='0'
        )
        
        yearly_diff = self.extract_yearly_difference(
            sheet_name,
            start_year=2020,
            region_column=1,
            classification_column=2,
            classification_value='0'
        )
        
        current_quarter_key = f"{self.current_year}.{self.current_quarter}/4p"
        current_data = quarterly_diff.get(current_quarter_key, 
                       quarterly_diff.get(f"{self.current_year}.{self.current_quarter}/4", {}))
        
        national_diff = current_data.get('전국', None)
        report_data['national_summary'] = {
            'change': national_diff,
            'direction': '상승' if national_diff > 0 else ('하락' if national_diff < 0 else '보합'),
        }
        
        regional_list = []
        for region in self.ALL_REGIONS:
            if region == '전국':
                continue
            diff = current_data.get(region, None)
            regional_list.append({
                'region': region,
                'change': diff,
                'direction': '상승' if diff > 0 else ('하락' if diff < 0 else '보합'),
            })
        
        increase_regions = sorted([r for r in regional_list if r['change'] > 0], 
                                  key=lambda x: x['change'], reverse=True)
        decrease_regions = sorted([r for r in regional_list if r['change'] < 0], 
                                  key=lambda x: x['change'])
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        report_data['yearly_data'] = yearly_diff
        report_data['quarterly_data'] = quarterly_diff
        
        return report_data
    
    def extract_unemployment_report_data(self) -> Dict[str, Any]:
        """실업률 보도자료 데이터 추출 (전년동기비 %p 차이)
        
        실업자 수 시트는 지역명이 전체 이름으로 되어 있고 (서울특별시, 부산광역시 등),
        연령계층별로 '계', '15-29세', '30-59세', '60세이상'으로 구분됨
        """
        sheet_name = '실업자 수'
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '실업률',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
        }
        
        # 실업자 수 시트 직접 파싱
        df = self._load_sheet(sheet_name)
        if df is None:
            return report_data
        
        config = self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        current_q_key = f"{self.current_year}_{self.current_quarter}Q"
        prev_q_key = f"{self.current_year - 1}_{self.current_quarter}Q"
        
        current_col = config.get(current_q_key)
        prev_col = config.get(prev_q_key)
        
        if current_col is None or prev_col is None:
            return report_data
        
        # 지역별 '계' 행에서 데이터 추출
        region_data = {}
        current_region = None
        
        for row_idx in range(3, len(df)):
            col0 = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ''
            col1 = str(df.iloc[row_idx, 1]).strip() if pd.notna(df.iloc[row_idx, 1]) else ''
            
            # 새 지역 시작
            if col0 and col0 != 'nan':
                current_region = self.REGION_FULL_TO_SHORT.get(col0, col0)
            
            # '계' 행 (전체 연령)
            if col1 == '계' and current_region in self.ALL_REGIONS:
                try:
                    current_val = float(df.iloc[row_idx, current_col]) if pd.notna(df.iloc[row_idx, current_col]) else None
                    prev_val = float(df.iloc[row_idx, prev_col]) if pd.notna(df.iloc[row_idx, prev_col]) else None
                    
                    if current_val is not None and prev_val is not None:
                        diff = round(current_val - prev_val, 1)
                    else:
                        diff = None
                    
                    region_data[current_region] = {
                        'change': diff,
                        'current_value': current_val,
                        'prev_value': prev_val,
                    }
                except (ValueError, TypeError):
                    pass
        
        # 전국 데이터
        national_data = region_data.get('전국', {})
        national_diff = national_data.get('change')
        report_data['national_summary'] = {
            'change': national_diff,
            'direction': '상승' if national_diff and national_diff > 0 else ('하락' if national_diff and national_diff < 0 else 'N/A'),
        }
        
        # 지역별 데이터
        regional_list = []
        for region in self.ALL_REGIONS:
            if region == '전국':
                continue
            data = region_data.get(region, {})
            diff = data.get('change')
            regional_list.append({
                'region': region,
                'change': diff,
                'direction': '상승' if diff and diff > 0 else ('하락' if diff and diff < 0 else 'N/A'),
            })
        
        # None이 아닌 값만 정렬
        increase_regions = sorted([r for r in regional_list if r['change'] is not None and r['change'] > 0], 
                                  key=lambda x: x['change'], reverse=True)
        decrease_regions = sorted([r for r in regional_list if r['change'] is not None and r['change'] < 0], 
                                  key=lambda x: x['change'])
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        
        # summary_table 생성
        report_data['summary_table'] = self._generate_unemployment_summary_table(df, config)
        
        return report_data
    
    def _generate_unemployment_summary_table(self, df, config: Dict) -> Dict[str, Any]:
        """실업률 요약 테이블 데이터 생성"""
        # 테이블용 분기 컬럼들
        quarter_keys = [
            (f"{self.current_year - 2}_{self.current_quarter}Q", f"{self.current_year - 2}.{self.current_quarter}/4"),
            (f"{self.current_year - 1}_{self.current_quarter}Q", f"{self.current_year - 1}.{self.current_quarter}/4"),
            (f"{self.current_year}_{self.current_quarter - 1 if self.current_quarter > 1 else 4}Q", 
             f"{self.current_year if self.current_quarter > 1 else self.current_year - 1}.{self.current_quarter - 1 if self.current_quarter > 1 else 4}/4"),
            (f"{self.current_year}_{self.current_quarter}Q", f"{self.current_year}.{self.current_quarter}/4p"),
        ]
        
        # 지역별 데이터 추출
        region_data = {}
        current_region = None
        
        for row_idx in range(3, len(df)):
            col0 = str(df.iloc[row_idx, 0]).strip() if pd.notna(df.iloc[row_idx, 0]) else ''
            col1 = str(df.iloc[row_idx, 1]).strip() if pd.notna(df.iloc[row_idx, 1]) else ''
            
            if col0 and col0 != 'nan':
                current_region = self.REGION_FULL_TO_SHORT.get(col0, col0)
            
            if col1 == '계' and current_region in self.ALL_REGIONS:
                if current_region not in region_data:
                    region_data[current_region] = {'changes': [], 'rates': []}
                
                # 각 분기별 전년동기비 차이 계산
                for q_key, _ in quarter_keys:
                    col = config.get(q_key)
                    prev_q_key = f"{int(q_key.split('_')[0]) - 1}_{q_key.split('_')[1]}"
                    prev_col = config.get(prev_q_key)
                    
                    if col and prev_col:
                        try:
                            current_val = float(df.iloc[row_idx, col]) if pd.notna(df.iloc[row_idx, col]) else None
                            prev_val = float(df.iloc[row_idx, prev_col]) if pd.notna(df.iloc[row_idx, prev_col]) else None
                            
                            if current_val is not None and prev_val is not None:
                                diff = round(current_val - prev_val, 1)
                            else:
                                diff = None
                            region_data[current_region]['changes'].append(diff)
                        except (ValueError, TypeError):
                            region_data[current_region]['changes'].append(None)
                    else:
                        region_data[current_region]['changes'].append(None)
                
                # 실업률 (현재, 청년)
                curr_q_col = config.get(f"{self.current_year}_{self.current_quarter}Q")
                if curr_q_col:
                    try:
                        rate = float(df.iloc[row_idx, curr_q_col]) if pd.notna(df.iloc[row_idx, curr_q_col]) else None
                        region_data[current_region]['rates'].append(rate)
                    except (ValueError, TypeError):
                        region_data[current_region]['rates'].append(None)
                else:
                    region_data[current_region]['rates'].append(None)
        
        # 테이블 행 생성
        rows = self._generate_region_rows(region_data, 'changes', 'rates')
        
        return {
            'columns': {
                'change_columns': [label for _, label in quarter_keys],
                'rate_columns': [
                    f"{self.current_year - 1}.{self.current_quarter}/4",
                    f"{self.current_year}.{self.current_quarter}/4",
                    "15-29세"
                ],
            },
            'rows': rows,
        }
    
    def extract_population_migration_report_data(self) -> Dict[str, Any]:
        """국내인구이동 보도자료 데이터 추출
        
        시도 간 이동 시트 구조:
        - col0: 지역코드
        - col1: 지역이름 (서울, 부산 등)
        - col2: 분류단계 (유입인구 수, 유출인구 수, 순인구이동 수)
        """
        sheet_name = '시도 간 이동'
        
        report_data = {
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'title': '국내 인구이동',
            },
            'national_summary': {},
            'regional_data': {},
            'top3_increase_regions': [],
            'top3_decrease_regions': [],
        }
        
        df = self._load_sheet(sheet_name)
        if df is None:
            return report_data
        
        config = self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        current_q_key = f"{self.current_year}_{self.current_quarter}Q"
        current_col = config.get(current_q_key)
        
        if current_col is None:
            return report_data
        
        # 지역별 순인구이동 데이터 추출
        region_data = {}
        
        for row_idx in range(3, len(df)):
            region = str(df.iloc[row_idx, 1]).strip() if pd.notna(df.iloc[row_idx, 1]) else ''
            category = str(df.iloc[row_idx, 2]).strip() if pd.notna(df.iloc[row_idx, 2]) else ''
            
            # 순인구이동 수 행만 추출
            if region in self.ALL_REGIONS and '순' in category and '이동' in category:
                try:
                    value = float(df.iloc[row_idx, current_col]) if pd.notna(df.iloc[row_idx, current_col]) else None
                    region_data[region] = value
                except (ValueError, TypeError):
                    region_data[region] = None
        
        # 전국은 시도간 이동 시트에 없음 (시도별 합계도 아님, 0으로 처리)
        # 순유입 지역과 순유출 지역의 합은 0이 됨
        report_data['national_summary'] = {
            'value': None,  # 전국 데이터 없음
            'direction': 'N/A',
        }
        
        # 지역별 데이터
        regional_list = []
        for region in self.ALL_REGIONS:
            if region == '전국':
                continue
            value = region_data.get(region)
            regional_list.append({
                'region': region,
                'value': value,
                'direction': '순유입' if value and value > 0 else ('순유출' if value and value < 0 else 'N/A'),
            })
        
        # None이 아닌 값만 정렬
        increase_regions = sorted([r for r in regional_list if r['value'] is not None and r['value'] > 0], 
                                  key=lambda x: x['value'], reverse=True)
        decrease_regions = sorted([r for r in regional_list if r['value'] is not None and r['value'] < 0], 
                                  key=lambda x: x['value'])
        
        report_data['top3_increase_regions'] = increase_regions[:3]
        report_data['top3_decrease_regions'] = decrease_regions[:3]
        report_data['regional_data'] = {
            'all': regional_list,
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
        }
        
        # summary_table 생성
        report_data['summary_table'] = self._generate_population_summary_table(df, config)
        
        return report_data
    
    def _generate_population_summary_table(self, df, config: Dict) -> Dict[str, Any]:
        """국내이동 요약 테이블 데이터 생성"""
        # 테이블용 분기 컬럼들
        quarter_keys = [
            (f"{self.current_year - 2}_{self.current_quarter}Q", f"{self.current_year - 2}.{self.current_quarter}/4"),
            (f"{self.current_year - 1}_{self.current_quarter}Q", f"{self.current_year - 1}.{self.current_quarter}/4"),
            (f"{self.current_year}_{self.current_quarter - 1 if self.current_quarter > 1 else 4}Q", 
             f"{self.current_year if self.current_quarter > 1 else self.current_year - 1}.{self.current_quarter - 1 if self.current_quarter > 1 else 4}/4"),
            (f"{self.current_year}_{self.current_quarter}Q", f"{self.current_year}.{self.current_quarter}/4p"),
        ]
        
        # 지역별 순인구이동 데이터 추출
        region_data = {}
        
        for row_idx in range(3, len(df)):
            region = str(df.iloc[row_idx, 1]).strip() if pd.notna(df.iloc[row_idx, 1]) else ''
            category = str(df.iloc[row_idx, 2]).strip() if pd.notna(df.iloc[row_idx, 2]) else ''
            
            if region in self.ALL_REGIONS and '순' in category and '이동' in category:
                if region not in region_data:
                    region_data[region] = {'migrations': [], 'amounts': []}
                
                # 각 분기별 데이터
                for q_key, _ in quarter_keys:
                    col = config.get(q_key)
                    if col:
                        try:
                            value = float(df.iloc[row_idx, col]) if pd.notna(df.iloc[row_idx, col]) else None
                            # 천명 단위로 변환
                            if value is not None:
                                value = round(value / 1000, 1)
                            region_data[region]['migrations'].append(value)
                        except (ValueError, TypeError):
                            region_data[region]['migrations'].append(None)
                    else:
                        region_data[region]['migrations'].append(None)
                
                # 현재 분기 금액 (전체, 20-29세)
                curr_col = config.get(f"{self.current_year}_{self.current_quarter}Q")
                if curr_col:
                    try:
                        total = float(df.iloc[row_idx, curr_col]) if pd.notna(df.iloc[row_idx, curr_col]) else None
                        if total is not None:
                            total = round(total / 1000, 1)
                        region_data[region]['amounts'] = [total, None]  # 20-29세는 별도 시트 필요
                    except (ValueError, TypeError):
                        region_data[region]['amounts'] = [None, None]
                else:
                    region_data[region]['amounts'] = [None, None]
        
        # 테이블 행 생성
        rows = self._generate_region_rows(region_data, 'migrations', 'amounts')
        
        return {
            'columns': {
                'quarter_columns': [label for _, label in quarter_keys],
                'amount_columns': ['전체', '20-29세'],
                'current_quarter': f"{self.current_year}.{self.current_quarter}/4",
            },
            'rows': rows,
        }
    
    def _generate_region_rows(self, region_data: Dict, data_key: str, secondary_key: str) -> List[Dict]:
        """공통 지역별 테이블 행 생성"""
        rows = []
        region_display = {
            '전국': '전 국', '서울': '서 울', '부산': '부 산', '대구': '대 구', '인천': '인 천',
            '광주': '광 주', '대전': '대 전', '울산': '울 산', '세종': '세 종', '경기': '경 기',
            '강원': '강 원', '충북': '충 북', '충남': '충 남', '전북': '전 북', '전남': '전 남',
            '경북': '경 북', '경남': '경 남', '제주': '제 주'
        }
        
        REGION_GROUPS = {
            "수도권": ["서울", "인천", "경기"],
            "동남권": ["부산", "울산", "경남"],
            "대경권": ["대구", "경북"],
            "호남권": ["광주", "전북", "전남"],
            "충청권": ["대전", "세종", "충북", "충남"],
            "강원제주": ["강원", "제주"]
        }
        
        # 전국 행 (있는 경우)
        if '전국' in region_data:
            national = region_data.get('전국', {data_key: [None]*4, secondary_key: [None, None]})
            rows.append({
                'region': '전 국',
                'group': None,
                data_key: national.get(data_key, [None]*4),
                secondary_key: national.get(secondary_key, [None, None]),
            })
        
        # 권역별 시도
        for group_name in ['수도권', '동남권', '대경권', '호남권', '충청권', '강원제주']:
            sidos = REGION_GROUPS[group_name]
            for idx, sido in enumerate(sidos):
                sido_data = region_data.get(sido, {data_key: [None]*4, secondary_key: [None, None]})
                
                row_data = {
                    'region': region_display.get(sido, sido),
                    data_key: sido_data.get(data_key, [None]*4),
                    secondary_key: sido_data.get(secondary_key, [None, None]),
                }
                
                if idx == 0:
                    row_data['group'] = group_name
                    row_data['rowspan'] = len(sidos)
                else:
                    row_data['group'] = None
                
                rows.append(row_data)
        
        return rows
    
    def extract_all_report_data(self) -> Dict[str, Dict[str, Any]]:
        """모든 부문별 보도자료 데이터 추출
        
        Returns:
            {
                'manufacturing': {...},
                'service': {...},
                'consumption': {...},
                ...
            }
        """
        return {
            'manufacturing': self.extract_mining_manufacturing_report_data(),
            'service': self.extract_service_industry_report_data(),
            'consumption': self.extract_consumption_report_data(),
            'construction': self.extract_construction_report_data(),
            'export': self.extract_export_report_data(),
            'import': self.extract_import_report_data(),
            'price': self.extract_price_report_data(),
            'employment': self.extract_employment_rate_report_data(),
            'unemployment': self.extract_unemployment_report_data(),
            'population': self.extract_population_migration_report_data(),
        }
    
    # ========== 시도별 보도자료 데이터 추출 메서드 ==========
    
    def extract_regional_data(self, region: str) -> Dict[str, Any]:
        """특정 지역의 시도별 보도자료 데이터 추출
        
        Args:
            region: 지역명 (예: '서울', '부산', '대구' 등)
            
        Returns:
            시도별 보도자료 템플릿에서 사용할 데이터 딕셔너리
        """
        # 템플릿이 기대하는 구조 생성 (regional_schema.json 준수)
        region_info = self._get_region_info(region)
        
        # 지역 코드 매핑
        REGION_CODES = {
            '서울': 11, '부산': 21, '대구': 22, '인천': 23, '광주': 24,
            '대전': 25, '울산': 26, '세종': 29, '경기': 31, '강원': 32,
            '충북': 33, '충남': 34, '전북': 35, '전남': 36, '경북': 37,
            '경남': 38, '제주': 39
        }
        
        report_data = {
            'region': region,
            'region_info': region_info,
            'report_info': {
                'year': self.current_year,
                'quarter': self.current_quarter,
                'region': region,
                'region_full_name': region_info.get('full_name', region),  # 스키마 필드명
                'region_name': region_info.get('full_name', region),  # 레거시 호환
                'region_code': REGION_CODES.get(region, 0),  # 지역 코드
                'region_index': region_info.get('index', 0),
                'page_number': 15 + region_info.get('index', 0),  # 기본 페이지 번호
            },
            'charts': {},  # 스키마 필드명 (chart_data → charts)
            'chart_data': {},  # 레거시 호환
            'summary_table': {'title': f'{region} 주요경제지표'},
        }
        
        # 생산 섹션 (광공업, 서비스업) + 상위/하위 산업
        mfg_data = self._extract_regional_indicator(region, '광공업생산', '0', 'growth_rate')
        svc_data = self._extract_regional_indicator(region, '서비스업생산', '0', 'growth_rate')
        
        mfg_industries = self.extract_regional_top_industries(region, '광공업생산', 3)
        svc_industries = self.extract_regional_top_industries(region, '서비스업생산', 3)
        
        mfg_template = self._to_template_format(mfg_data)
        svc_template = self._to_template_format(svc_data)
        
        # 상위/하위 산업 데이터 추가 (스키마: increase_industries, decrease_industries)
        mfg_template['increase_industries'] = [{'name': i['name'], 'growth_rate': i['value']} for i in mfg_industries.get('increase', [])]
        mfg_template['decrease_industries'] = [{'name': i['name'], 'growth_rate': i['value']} for i in mfg_industries.get('decrease', [])]
        svc_template['increase_industries'] = [{'name': i['name'], 'growth_rate': i['value']} for i in svc_industries.get('increase', [])]
        svc_template['decrease_industries'] = [{'name': i['name'], 'growth_rate': i['value']} for i in svc_industries.get('decrease', [])]
        
        report_data['production'] = {
            'manufacturing': mfg_template,
            'service': svc_template,
        }
        
        # 소비·건설 섹션 (분류단계 '0'이 합계 행) + 상위/하위 카테고리
        retail_data = self._extract_regional_indicator(region, '소비(소매, 추가)', '0', 'growth_rate')
        const_data = self._extract_regional_indicator(region, '건설 (공표자료)', '0', 'growth_rate')
        
        retail_categories = self.extract_regional_top_industries(region, '소비(소매, 추가)', 3)
        const_categories = self.extract_regional_top_industries(region, '건설 (공표자료)', 3)
        
        retail_template = self._to_template_format(retail_data)
        const_template = self._to_template_format(const_data)
        
        # 상위/하위 카테고리 데이터 추가 (스키마: increase_categories, decrease_categories)
        retail_template['increase_categories'] = [{'name': i['name'], 'growth_rate': i['value']} for i in retail_categories.get('increase', [])]
        retail_template['decrease_categories'] = [{'name': i['name'], 'growth_rate': i['value']} for i in retail_categories.get('decrease', [])]
        const_template['increase_categories'] = [{'name': i['name'], 'growth_rate': i['value']} for i in const_categories.get('increase', [])]
        const_template['decrease_categories'] = [{'name': i['name'], 'growth_rate': i['value']} for i in const_categories.get('decrease', [])]
        
        report_data['consumption_construction'] = {
            'retail': retail_template,
            'construction': const_template,
        }
        
        # 수출입·물가 섹션 + 상위/하위 품목
        export_data = self._extract_regional_indicator(region, '수출', '0', 'growth_rate')
        import_data = self._extract_regional_indicator(region, '수입', '0', 'growth_rate')
        price_data = self._extract_regional_indicator(region, '품목성질별 물가', '0', 'growth_rate')
        
        export_items = self.extract_regional_top_industries(region, '수출', 3)
        import_items = self.extract_regional_top_industries(region, '수입', 3)
        price_items = self.extract_regional_top_industries(region, '품목성질별 물가', 3)
        
        export_template = self._to_template_format(export_data)
        import_template = self._to_template_format(import_data)
        price_template = self._to_template_format(price_data, indicator_type='price')
        
        # 상위/하위 품목 데이터 추가 (스키마: increase_items, decrease_items)
        export_template['increase_items'] = [{'name': i['name'], 'growth_rate': i['value']} for i in export_items.get('increase', [])]
        export_template['decrease_items'] = [{'name': i['name'], 'growth_rate': i['value']} for i in export_items.get('decrease', [])]
        import_template['increase_items'] = [{'name': i['name'], 'growth_rate': i['value']} for i in import_items.get('increase', [])]
        import_template['decrease_items'] = [{'name': i['name'], 'growth_rate': i['value']} for i in import_items.get('decrease', [])]
        price_template['increase_items'] = [{'name': i['name'], 'growth_rate': i['value']} for i in price_items.get('increase', [])]
        price_template['decrease_items'] = [{'name': i['name'], 'growth_rate': i['value']} for i in price_items.get('decrease', [])]
        
        report_data['export_import_price'] = {
            'export': export_template,
            'import': import_template,
            'consumer_price': price_template,
        }
        
        # 고용·인구이동 섹션 (연령별고용률 시트 우선, 고용률 시트 폴백)
        emp_data = self._extract_regional_indicator(region, '연령별고용률', '0', 'difference')
        if emp_data.get('value') is None:
            emp_data = self._extract_regional_indicator(region, '고용률', '0', 'difference')
        pop_data = self._extract_regional_indicator(region, '시도 간 이동', None, 'absolute')
        
        report_data['employment_migration'] = {
            'employment_rate': self._to_template_format(emp_data, indicator_type='employment'),
            'population_migration': self._to_template_format(pop_data, indicator_type='population'),
        }
        
        # 차트용 시계열 데이터 (스키마: charts, 레거시: chart_data)
        chart_data = self._extract_regional_chart_data(region)
        report_data['charts'] = chart_data
        report_data['chart_data'] = chart_data  # 레거시 호환
        
        # 요약 표 데이터
        report_data['summary_table'] = self._extract_regional_summary_table(region)
        
        # 페이지 번호 템플릿 변수 (레거시 호환)
        report_data['page_number'] = report_data['report_info']['page_number']
        
        return report_data
    
    def _to_template_format(self, data: Dict[str, Any], indicator_type: str = 'default') -> Dict[str, Any]:
        """_extract_regional_indicator 결과를 템플릿 형식으로 변환
        
        템플릿이 기대하는 형식:
        - total_growth_rate: 총 증감률 (None이면 0.0으로, direction은 'N/A')
        - direction: 방향 (증가/감소/보합/N/A)
        - 각종 categories, items, industries 리스트들 (가중치 데이터 필요)
        
        ★ 중요: 
        - 결측치는 0.0으로 변환하여 템플릿에서 오류 방지 (| abs 필터 등)
        - direction='N/A'로 표시하여 데이터가 없음을 알림
        - is_missing=True로 표시하여 필요 시 템플릿에서 처리 가능
        """
        value = data.get('value')
        direction = data.get('direction', 'N/A')
        
        # None 값을 0.0으로 변환 (템플릿에서 | abs 등 필터 사용 시 오류 방지)
        safe_value = 0.0 if value is None else value
        
        result = {
            'total_growth_rate': safe_value,  # None -> 0.0 변환
            'direction': direction,
            'is_missing': value is None,  # 원래 데이터가 없음을 표시
            # 광공업/서비스업 생산용
            'increase_industries': [],  # 가중치 데이터 필요
            'decrease_industries': [],
            # 수출/수입용
            'increase_items': [],  # 가중치 데이터 필요
            'decrease_items': [],
            'increase_products': [],
            'decrease_products': [],
            # 소비용
            'increase_businesses': [],
            'decrease_businesses': [],
            'increase_categories': [],  # 소매판매 카테고리
            'decrease_categories': [],
            # 물가용
            'categories': [],
            # 고용·인구이동용
            'increase_age_groups': [],
            'decrease_age_groups': [],
            'inflow_age_groups': [],
            'outflow_age_groups': [],
        }
        
        # 지표 유형별 추가 필드
        if indicator_type == 'employment':
            result['change'] = safe_value  # 템플릿 오류 방지
            result['total_change'] = safe_value
            result['rate'] = None
        elif indicator_type == 'population':
            result['net_migration'] = int(value) if value is not None else 0  # 템플릿 오류 방지
        elif indicator_type == 'price':
            result['index'] = None  # 원지수는 별도 추출 필요
            result['change'] = value
            result['total_growth_rate'] = value
            # 물가는 주로 상승이므로 direction 조정
            if value is not None and value > 0:
                result['direction'] = '상승'
            elif value is not None and value < 0:
                result['direction'] = '하락'
        
        return result
    
    def _get_region_info(self, region: str) -> Dict[str, Any]:
        """지역 정보 반환"""
        REGION_INFO = {
            '서울': {'full_name': '서울특별시', 'index': 1},
            '부산': {'full_name': '부산광역시', 'index': 2},
            '대구': {'full_name': '대구광역시', 'index': 3},
            '인천': {'full_name': '인천광역시', 'index': 4},
            '광주': {'full_name': '광주광역시', 'index': 5},
            '대전': {'full_name': '대전광역시', 'index': 6},
            '울산': {'full_name': '울산광역시', 'index': 7},
            '세종': {'full_name': '세종특별자치시', 'index': 8},
            '경기': {'full_name': '경기도', 'index': 9},
            '강원': {'full_name': '강원특별자치도', 'index': 10},
            '충북': {'full_name': '충청북도', 'index': 11},
            '충남': {'full_name': '충청남도', 'index': 12},
            '전북': {'full_name': '전북특별자치도', 'index': 13},
            '전남': {'full_name': '전라남도', 'index': 14},
            '경북': {'full_name': '경상북도', 'index': 15},
            '경남': {'full_name': '경상남도', 'index': 16},
            '제주': {'full_name': '제주특별자치도', 'index': 17},
        }
        return REGION_INFO.get(region, {'full_name': region, 'index': 0})
    
    def _extract_regional_indicator(self, region: str, sheet_name: str, 
                                    level_value: Optional[str], value_type: str) -> Dict[str, Any]:
        """지역별 특정 지표 데이터 추출 (동적 헤더 파싱 지원)
        
        Args:
            region: 지역명
            sheet_name: 시트 이름
            level_value: 분류단계 값 (None이면 필터링 안함)
            value_type: 'growth_rate' (전년동기비), 'difference' (%p 차이), 'absolute' (절대값)
            
        Returns:
            {'value': 값, 'direction': 방향}
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            print(f"[지역지표] {sheet_name} 시트 로드 실패")
            return {'value': None, 'direction': 'N/A'}
        
        # 설정 가져오기 (기본값 포함)
        config = self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        header_row = config.get('header_row', 2)
        
        # 동적으로 분기 컬럼 찾기 (하드코딩 설정보다 우선)
        current_col = self.get_quarter_column_index(sheet_name, self.current_year, self.current_quarter)
        prev_col = self.get_quarter_column_index(sheet_name, self.current_year - 1, self.current_quarter)
        
        # 동적 파싱 실패 시 하드코딩 설정 사용
        if current_col is None:
            current_quarter_key = f"{self.current_year}_{self.current_quarter}Q"
            current_col = config.get(current_quarter_key)
        if prev_col is None:
            prev_quarter_key = f"{self.current_year - 1}_{self.current_quarter}Q"
            prev_col = config.get(prev_quarter_key)
        
        if current_col is None:
            print(f"[지역지표] {sheet_name}: {self.current_year}년 {self.current_quarter}분기 컬럼을 찾을 수 없음")
            return {'value': None, 'direction': 'N/A'}
        
        print(f"[지역지표] {sheet_name} {region}: 현재분기 열={current_col}, 전년동기 열={prev_col}")
        
        # 해당 지역/분류 행 찾기
        for row_idx in range(header_row + 1, len(df)):
            try:
                row_region = str(df.iloc[row_idx, region_col]).strip()
                # 지역명 정규화
                row_region = self.REGION_FULL_TO_SHORT.get(row_region, row_region)
                
                if row_region != region:
                    continue
                
                if level_value is not None:
                    row_level = str(df.iloc[row_idx, level_col]).strip()
                    # '0.0' -> '0' 등 정규화
                    if row_level.replace('.0', '') != level_value.replace('.0', ''):
                        continue
                
                current_val = df.iloc[row_idx, current_col]
                if pd.isna(current_val):
                    continue
                
                current_val = float(current_val)
                
                if value_type == 'absolute':
                    # 절대값 (인구이동)
                    result_val = current_val
                    direction = '순유입' if result_val > 0 else ('순유출' if result_val < 0 else '균형')
                elif value_type == 'difference':
                    # 차이 (고용률 %p)
                    if prev_col is None:
                        return {'value': None, 'direction': 'N/A'}
                    prev_val = df.iloc[row_idx, prev_col]
                    if pd.isna(prev_val):
                        return {'value': None, 'direction': 'N/A'}
                    result_val = round(float(current_val) - float(prev_val), 1)
                    direction = '상승' if result_val > 0 else ('하락' if result_val < 0 else '보합')
                else:
                    # 증감률 (growth_rate)
                    if prev_col is None:
                        return {'value': None, 'direction': 'N/A'}
                    prev_val = df.iloc[row_idx, prev_col]
                    if pd.isna(prev_val) or float(prev_val) == 0:
                        return {'value': None, 'direction': 'N/A'}
                    result_val = round((float(current_val) / float(prev_val) - 1) * 100, 1)
                    direction = '증가' if result_val > 0 else ('감소' if result_val < 0 else '보합')
                
                print(f"[지역지표] {sheet_name} {region}: {result_val} ({direction})")
                return {'value': result_val, 'direction': direction}
                
            except (IndexError, ValueError, TypeError) as e:
                continue
        
        print(f"[지역지표] {sheet_name} {region}: 데이터를 찾을 수 없음")
        return {'value': None, 'direction': 'N/A'}
    
    def _extract_regional_chart_data(self, region: str) -> Dict[str, List[Dict[str, Any]]]:
        """지역별 차트용 시계열 데이터 추출"""
        chart_data = {}
        
        # 각 지표별 차트 데이터 추출 (분류단계 '0'이 합계 행)
        indicators = [
            ('manufacturing', '광공업생산', '0'),
            ('service', '서비스업생산', '0'),
            ('retail', '소비(소매, 추가)', '0'),
            ('construction', '건설 (공표자료)', '0'),
            ('export', '수출', '0'),
            ('import', '수입', '0'),
            ('price', '품목성질별 물가', '0'),
            ('employment', '고용률', '0'),
        ]
        
        for key, sheet_name, level_value in indicators:
            chart_data[key] = self._get_time_series_data(region, sheet_name, level_value)
        
        return chart_data
    
    def _get_time_series_data(self, region: str, sheet_name: str, 
                              level_value: Optional[str]) -> List[Dict[str, Any]]:
        """시계열 데이터 추출 (전년동기비 증감률) - 동적 헤더 파싱 사용"""
        config = self.RAW_SHEET_CONFIG.get(sheet_name)
        if not config:
            return []
        
        df = self._load_sheet(sheet_name)
        if df is None:
            return []
        
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        
        # 해당 지역/분류 행 찾기
        row_data = None
        for row_idx in range(len(df)):
            try:
                row_region = str(df.iloc[row_idx, region_col]).strip()
                if row_region != region:
                    continue
                
                if level_value is not None:
                    row_level = str(df.iloc[row_idx, level_col]).strip()
                    if row_level != level_value:
                        continue
                
                row_data = df.iloc[row_idx]
                break
            except (IndexError, ValueError):
                continue
        
        if row_data is None:
            return []
        
        # 동적 헤더 파싱으로 분기 컬럼 찾기
        structure = self.parse_sheet_structure(sheet_name, config.get('header_row', 2))
        quarters_info = structure.get('quarters', {})
        
        result = []
        
        # 현재 분기 기준으로 최근 16개 분기 데이터 추출
        # (현재 분기 포함하여 과거 4년간)
        for year in range(self.current_year - 3, self.current_year + 1):
            for q in range(1, 5):
                # 현재보다 미래 분기는 제외
                if year > self.current_year or (year == self.current_year and q > self.current_quarter):
                    continue
                
                quarter_key = f"{year} {q}/4"
                prev_quarter_key = f"{year - 1} {q}/4"
                
                current_col = quarters_info.get(quarter_key)
                prev_col = quarters_info.get(prev_quarter_key)
                
                # 라벨: "'23.1/4" 형식
                q_label = f"'{str(year)[-2:]}.{q}/4"
                
                if current_col is None or prev_col is None:
                    continue  # 컬럼이 없으면 건너뜀
                
                try:
                    current_val = row_data.iloc[current_col]
                    prev_val = row_data.iloc[prev_col]
                    
                    if pd.isna(current_val) or pd.isna(prev_val) or float(prev_val) == 0:
                        result.append({'period': q_label, 'value': 0.0})
                    else:
                        growth_rate = round((float(current_val) / float(prev_val) - 1) * 100, 1)
                        result.append({'period': q_label, 'value': growth_rate})
                except (IndexError, ValueError, TypeError):
                    result.append({'period': q_label, 'value': 0.0})
        
        return result
    
    def _extract_regional_summary_table(self, region: str) -> Dict[str, Any]:
        """지역별 요약 표 데이터 추출 (템플릿 형식: rows 배열)
        
        Returns:
            {
                'title': '서울 주요경제지표',
                'rows': [
                    {
                        'period': '2023.2/4',
                        'manufacturing': -8.3,
                        'service': 2.3,
                        'retail': -1.8,
                        'construction': 29.2,
                        'export': 1.6,
                        'import': -1.7,
                        'consumer_price': 2.0,
                        'employment_rate': -0.2,
                        'migration_total': 98.0,
                        'migration_20_29': 10.0
                    },
                    ...
                ]
            }
        """
        # 각 지표별 시트 및 설정
        indicators = [
            ('manufacturing', '광공업생산', '0', 'growth_rate'),
            ('service', '서비스업생산', '0', 'growth_rate'),
            ('retail', '소비(소매, 추가)', '0', 'growth_rate'),
            ('construction', '건설 (공표자료)', '0', 'growth_rate'),
            ('export', '수출', '0', 'growth_rate'),
            ('import', '수입', '0', 'growth_rate'),
            ('price', '품목성질별 물가', '0', 'growth_rate'),
            ('employment', '고용률', '0', 'difference'),
            ('population', '시도 간 이동', None, 'absolute'),  # 인구순이동은 분류 단계 필터링 안함
        ]
        
        # 여러 분기의 데이터 추출 (2023.2/4, 2024.2/4, 2025.1/4, 2025.2/4p)
        quarter_labels = [
            f"{self.current_year - 2}.{self.current_quarter}/4",  # 2023.2/4
            f"{self.current_year - 1}.{self.current_quarter}/4",  # 2024.2/4
            f"{self.current_year}.{self.current_quarter - 1 if self.current_quarter > 1 else 4}/4",  # 2025.1/4
            f"{self.current_year}.{self.current_quarter}/4p",  # 2025.2/4p
        ]
        
        rows = []
        
        for quarter_label in quarter_labels:
            row_data = {'period': quarter_label}
            
            # 각 지표별 데이터 추출
            for key, sheet_name, level_value, value_type in indicators:
                value = None
                
                # 분기별 데이터 추출
                if value_type == 'growth_rate':
                    # 전년동기비 증감률 (키 형식: "2023.2/4" 또는 "2025.2/4p")
                    quarterly_growth = self.extract_quarterly_growth_rate(
                        sheet_name,
                        start_year=2020,
                        region_column=self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {}).get('region_col', 1),
                        classification_column=self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {}).get('level_col', 2),
                        classification_value=level_value
                    )
                    # 키 형식: 이미 "2023.2/4" 형식이므로 그대로 사용
                    quarter_data = quarterly_growth.get(quarter_label, {})
                    value = quarter_data.get(region, None)
                    
                elif value_type == 'difference':
                    # %p 차이 (고용률) (키 형식: "2023.2/4" 또는 "2025.2/4p")
                    quarterly_diff = self.extract_quarterly_difference(
                        sheet_name,
                        start_year=2020,
                        region_column=self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {}).get('region_col', 1),
                        classification_column=self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {}).get('level_col', 2),
                        classification_value=level_value
                    )
                    # 키 형식: 이미 "2023.2/4" 형식이므로 그대로 사용
                    quarter_data = quarterly_diff.get(quarter_label, {})
                    value = quarter_data.get(region, None)
                    
                elif value_type == 'absolute':
                    # 절대값 (인구이동) (키 형식: "2023 2/4" 또는 "2025.2/4p")
                    # 인구순이동은 분류 단계 필터링 안함 (level_value가 None)
                    quarterly_data = self.extract_all_quarters(
                        sheet_name,
                        start_year=2020,
                        region_column=self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {}).get('region_col', 1),
                        classification_column=None if level_value is None else self.RAW_SHEET_QUARTER_COLS.get(sheet_name, {}).get('level_col', 2),
                        classification_value=level_value
                    )
                    # 키 형식 변환: "2023.2/4" -> "2023 2/4", "2025.2/4p" -> "2025.2/4p"
                    if quarter_label.endswith('p'):
                        lookup_key = quarter_label
                    else:
                        lookup_key = quarter_label.replace('.', ' ')
                    quarter_data = quarterly_data.get(lookup_key, {})
                    value = quarter_data.get(region, None)
                
                # 결측치는 None 유지 (템플릿에서 N/A로 표시)
                
                # 키 매핑
                if key == 'price':
                    row_data['consumer_price'] = value
                elif key == 'employment':
                    row_data['employment_rate'] = value
                elif key == 'population':
                    # 인구이동: 값을 그대로 사용 (천명 단위)
                    row_data['migration_total'] = value
                    row_data['migration_20_29'] = None  # 20-29세 데이터는 별도 추출 필요
                else:
                    row_data[key] = value
            
            rows.append(row_data)
        
        return {
            'title': f'{region} 주요경제지표',
            'rows': rows,
        }
    
    def extract_regional_top_industries(self, region: str, sheet_name: str, 
                                        top_n: int = 3) -> Dict[str, List[Dict[str, Any]]]:
        """지역별 상위/하위 업종 추출
        
        Args:
            region: 지역명
            sheet_name: 시트 이름 (광공업생산, 서비스업생산 등)
            top_n: 상위/하위 개수
            
        Returns:
            {'increase': [업종 목록], 'decrease': [업종 목록]}
        """
        config = self.RAW_SHEET_QUARTER_COLS.get(sheet_name)
        if not config:
            return {'increase': [], 'decrease': []}
        
        df = self._load_sheet(sheet_name)
        if df is None:
            return {'increase': [], 'decrease': []}
        
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        name_col = config.get('name_col', 5)
        
        current_quarter = f"{self.current_year}_{self.current_quarter}Q"
        prev_quarter = f"{self.current_year - 1}_{self.current_quarter}Q"
        
        current_col = config.get(current_quarter)
        prev_col = config.get(prev_quarter)
        
        if current_col is None or prev_col is None:
            return {'increase': [], 'decrease': []}
        
        industries = []
        
        for row_idx in range(len(df)):
            try:
                row_region = str(df.iloc[row_idx, region_col]).strip()
                if row_region != region:
                    continue
                
                row_level = str(df.iloc[row_idx, level_col]).strip()
                # 하위 분류만 (분류단계가 0이 아닌 것)
                if row_level in ['0', '계', '총지수', '', 'nan']:
                    continue
                
                industry_name = str(df.iloc[row_idx, name_col]).strip()
                if not industry_name or industry_name == 'nan':
                    continue
                
                current_val = df.iloc[row_idx, current_col]
                prev_val = df.iloc[row_idx, prev_col]
                
                if pd.isna(current_val) or pd.isna(prev_val) or float(prev_val) == 0:
                    continue
                
                growth_rate = round((float(current_val) / float(prev_val) - 1) * 100, 1)
                
                industries.append({
                    'name': industry_name,
                    'value': growth_rate,
                })
            except (IndexError, ValueError, TypeError):
                continue
        
        # 정렬
        increase = sorted([i for i in industries if i['value'] > 0], 
                         key=lambda x: x['value'], reverse=True)[:top_n]
        decrease = sorted([i for i in industries if i['value'] < 0], 
                         key=lambda x: x['value'])[:top_n]
        
        return {'increase': increase, 'decrease': decrease}

