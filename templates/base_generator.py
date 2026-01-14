#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Base Generator 클래스

모든 Generator가 공통으로 사용하는 기능을 제공하는 기본 클래스입니다.
"""

import pandas as pd
from pathlib import Path
from typing import Optional, Tuple, List, Any, Dict
from abc import ABC, abstractmethod


class BaseGenerator(ABC):
    """모든 Generator의 기본 클래스"""
    
    def __init__(
        self, 
        excel_path: str, 
        year: Optional[int] = None, 
        quarter: Optional[int] = None,
        excel_file: Optional[pd.ExcelFile] = None
    ):
        """
        초기화
        
        Args:
            excel_path: 엑셀 파일 경로
            year: 연도 (선택사항)
            quarter: 분기 (선택사항)
            excel_file: 캐시된 ExcelFile 객체 (선택사항, 있으면 재사용)
        """
        self.excel_path = excel_path
        self.year = year
        self.quarter = quarter
        self.xl = excel_file  # 캐시된 ExcelFile 객체 재사용
        self.df_cache = {}  # 시트별 DataFrame 캐시
        self._xl_owner = excel_file is not None  # 외부에서 전달된 경우 소유권 없음
        
    def load_excel(self) -> pd.ExcelFile:
        """엑셀 파일 로드 (캐싱)"""
        if self.xl is None:
            # 캐시에서 가져오기 시도
            try:
                import sys
                from pathlib import Path
                # 절대 import 시도
                cache_module_path = Path(__file__).parent.parent / 'services' / 'excel_cache.py'
                if cache_module_path.exists():
                    import importlib.util
                    spec = importlib.util.spec_from_file_location('excel_cache', str(cache_module_path))
                    excel_cache = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(excel_cache)
                    self.xl = excel_cache.get_excel_file(self.excel_path, use_data_only=True)
                    if self.xl is None:
                        # 캐시 실패 시 직접 로드
                        self.xl = pd.ExcelFile(self.excel_path)
                        self._xl_owner = True
                else:
                    # excel_cache 모듈이 없으면 직접 로드
                    self.xl = pd.ExcelFile(self.excel_path)
                    self._xl_owner = True
            except (ImportError, Exception) as e:
                # excel_cache 모듈이 없거나 오류 발생 시 직접 로드
                self.xl = pd.ExcelFile(self.excel_path)
                self._xl_owner = True
        return self.xl
    
    def get_sheet(self, sheet_name: str, use_cache: bool = True) -> Optional[pd.DataFrame]:
        """
        시트를 읽어서 DataFrame 반환
        
        Args:
            sheet_name: 시트 이름
            use_cache: 캐시 사용 여부
            
        Returns:
            DataFrame 또는 None
        """
        if use_cache and sheet_name in self.df_cache:
            return self.df_cache[sheet_name]
        
        xl = self.load_excel()
        if sheet_name not in xl.sheet_names:
            return None
        
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        
        if use_cache:
            self.df_cache[sheet_name] = df
        
        return df
    
    def find_sheet_with_fallback(
        self, 
        primary_sheets: List[str], 
        fallback_sheets: List[str]
    ) -> Tuple[Optional[str], bool]:
        """
        시트 찾기 - 기본 시트가 없으면 대체 시트 사용
        
        Args:
            primary_sheets: 우선 시트 이름 목록
            fallback_sheets: 대체 시트 이름 목록
            
        Returns:
            (시트 이름, 기초자료 사용 여부) 튜플
        """
        xl = self.load_excel()
        
        for sheet in primary_sheets:
            if sheet in xl.sheet_names:
                return sheet, False
        
        for sheet in fallback_sheets:
            if sheet in xl.sheet_names:
                print(f"[시트 대체] '{primary_sheets[0]}' → '{sheet}' (기초자료)")
                return sheet, True
        
        return None, False
    
    @staticmethod
    def safe_float(value: Any, default: Optional[float] = None) -> Optional[float]:
        """
        안전한 float 변환 함수 (NaN, '-' 체크 포함)
        
        Args:
            value: 변환할 값
            default: 변환 실패 시 반환할 기본값
            
        Returns:
            float 값 또는 default
        """
        if value is None:
            return default
        try:
            if pd.isna(value):
                return default
            if isinstance(value, str):
                value = value.strip()
                if value == '-' or value == '' or value == 'N/A':
                    return default
            result = float(value)
            if pd.isna(result):
                return default
            return result
        except (ValueError, TypeError):
            return default
    
    @staticmethod
    def safe_round(value: Any, decimals: int = 1, default: Optional[float] = None) -> Optional[float]:
        """
        안전한 반올림 함수
        
        Args:
            value: 반올림할 값
            decimals: 소수점 자릿수
            default: 변환 실패 시 반환할 기본값
            
        Returns:
            반올림된 float 값 또는 default
        """
        result = BaseGenerator.safe_float(value, default)
        if result is None:
            return default
        return round(result, decimals)
    
    @staticmethod
    def safe_int(value: Any, default: Optional[int] = None) -> Optional[int]:
        """
        안전한 int 변환 함수
        
        Args:
            value: 변환할 값
            default: 변환 실패 시 반환할 기본값
            
        Returns:
            int 값 또는 default
        """
        if value is None:
            return default
        try:
            if pd.isna(value):
                return default
            if isinstance(value, str):
                value = value.strip()
                if value == '-' or value == '' or value == 'N/A':
                    return default
            result = int(float(value))  # float을 거쳐서 변환 (소수점 처리)
            return result
        except (ValueError, TypeError):
            return default
    
    @staticmethod
    def safe_str(value: Any, default: str = '') -> str:
        """
        안전한 문자열 변환 함수
        
        Args:
            value: 변환할 값
            default: 변환 실패 시 반환할 기본값
            
        Returns:
            문자열 값 또는 default
        """
        if value is None:
            return default
        try:
            if pd.isna(value):
                return default
            result = str(value).strip()
            return result if result else default
        except (ValueError, TypeError):
            return default
    
    def get_cell_value(
        self, 
        df: pd.DataFrame, 
        row: int, 
        col: int, 
        default: Any = None
    ) -> Any:
        """
        DataFrame에서 특정 셀 값 안전하게 추출
        
        Args:
            df: DataFrame
            row: 행 인덱스 (0-based)
            col: 열 인덱스 (0-based)
            default: 값이 없을 때 반환할 기본값
            
        Returns:
            셀 값 또는 default
        """
        try:
            if row < 0 or row >= len(df):
                return default
            if col < 0 or col >= len(df.columns):
                return default
            value = df.iloc[row, col]
            if pd.isna(value):
                return default
            return value
        except (IndexError, KeyError):
            return default
    
    def find_row_by_value(
        self, 
        df: pd.DataFrame, 
        col: int, 
        value: Any, 
        start_row: int = 0
    ) -> Optional[int]:
        """
        특정 컬럼에서 값을 찾아 행 인덱스 반환
        
        Args:
            df: DataFrame
            col: 검색할 컬럼 인덱스
            value: 찾을 값
            start_row: 검색 시작 행
            
        Returns:
            행 인덱스 또는 None
        """
        try:
            for i in range(start_row, len(df)):
                cell_value = self.get_cell_value(df, i, col)
                if cell_value == value or str(cell_value).strip() == str(value).strip():
                    return i
        except Exception:
            pass
        return None
    
    def find_rows_by_condition(
        self,
        df: pd.DataFrame,
        conditions: Dict[int, Any],
        start_row: int = 0
    ) -> List[int]:
        """
        여러 조건으로 행 찾기
        
        Args:
            df: DataFrame
            conditions: {컬럼 인덱스: 값} 딕셔너리
            start_row: 검색 시작 행
            
        Returns:
            조건에 맞는 행 인덱스 목록
        """
        result = []
        try:
            for i in range(start_row, len(df)):
                match = True
                for col, value in conditions.items():
                    cell_value = self.get_cell_value(df, i, col)
                    if cell_value != value and str(cell_value).strip() != str(value).strip():
                        match = False
                        break
                if match:
                    result.append(i)
        except Exception:
            pass
        return result
    
    def check_sheet_has_data(
        self, 
        df: pd.DataFrame, 
        test_conditions: Dict[int, Any],
        max_empty_cells: int = 20
    ) -> bool:
        """
        시트에 실제 데이터가 있는지 확인 (수식 미계산 체크)
        
        Args:
            df: DataFrame
            test_conditions: {컬럼 인덱스: 값} 딕셔너리 (테스트 조건)
            max_empty_cells: 최대 빈 셀 개수
            
        Returns:
            데이터가 있으면 True, 없으면 False
        """
        try:
            rows = self.find_rows_by_condition(df, test_conditions)
            if not rows:
                return False
            
            # 첫 번째 매칭 행의 빈 셀 개수 확인
            test_row = df.iloc[rows[0]]
            empty_count = test_row.isna().sum()
            return empty_count <= max_empty_cells
        except Exception:
            return False
    
    @abstractmethod
    def extract_all_data(self) -> Dict[str, Any]:
        """
        모든 데이터 추출 (자식 클래스에서 구현)
        
        Returns:
            추출된 데이터 딕셔너리
        """
        pass
    
    def get_report_info(self) -> Dict[str, Any]:
        """
        report_info 기본 구조 반환
        
        Returns:
            report_info 딕셔너리
        """
        return {
            "year": self.year or 2025,
            "quarter": self.quarter or 2,
            "organization": "국가데이터처",
            "department": "경제통계심의관"
        }
