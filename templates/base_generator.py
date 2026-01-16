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
        근본 수정: parse_excel_value와 동일한 로직으로 콤마, 특수문자 처리
        
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
                # 근본 수정: '-' (통계표의 0) -> 0.0
                if value == '-' or value == '' or value == 'N/A':
                    return 0.0 if default is None else default
                # 근본 수정: 콤마 제거 후 실수 변환 ('1,234.5' -> 1234.5)
                # 정규표현식으로 숫자, 점(.), 음수 부호(-), 지수 표기(e, E)만 남기기
                import re
                # 숫자, 점(.), 음수 부호(-), 지수 표기(e, E)만 남기기 (콤마는 자동 제거됨)
                cleaned = re.sub(r'[^\d\.\-\+\eE]', '', value)
                # 빈 문자열이면 default 반환
                if not cleaned:
                    return 0.0 if default is None else default
                value = cleaned
            result = float(value)
            if pd.isna(result):
                return default
            return result
        except (ValueError, TypeError):
            # 근본 수정: 변환 불가능한 문자(예: '...')가 있어도 에러 내지 말고 0 처리
            return 0.0 if default is None else default
    
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
            "department": "경제동향통계심의관 지역경제동향과",
            "contact_phone": "042-481-xxxx",
            "page_number": ""  # 페이지 번호는 더 이상 사용하지 않음 (목차 생성 중단)
        }
    
    def find_target_col_index(
        self,
        header_row: Any,
        target_year: int,
        target_quarter: int
    ) -> int:
        """
        [Robust Dynamic Parsing System - 1단계]
        헤더에서 연도와 분기를 찾아 인덱스를 반환 (모든 포맷 대응)
        
        목표: 엑셀 파일의 구조가 어떻게 변하든 정확한 데이터를 찾아내는 유연한 탐색기
        - 입력: header_row(엑셀 행), 2025, 3
        - 매칭 대상: "2025 3/4", "2025.3/4", "2025년 3분기", "25년 3/4" 등
        - 정규표현식(Regex)에 준하는 유연성 제공
        
        Args:
            header_row: 헤더 행 (pandas Series, 리스트, 튜플, openpyxl Cell 객체 등)
            target_year: 찾을 연도 (예: 2025)
            target_quarter: 찾을 분기 (예: 3)
            
        Returns:
            컬럼 인덱스 (0-based)
            
        Raises:
            ValueError: 헤더에서 해당 연도/분기 열을 찾을 수 없을 때
        """
        # 1. 연도 문자열 준비 (전체 연도 + 단축형)
        y_str = str(target_year)  # "2025"
        # 2025가 안되면 25로도 찾을 수 있게 대비
        y_short = y_str[2:] if len(y_str) == 4 else y_str  # "25"
        
        # 2. 분기 표현식 다양화
        q_markers = [
            f"{target_quarter}/4",      # "3/4" (가장 일반적)
            f"{target_quarter}분기",    # "3분기"
            f"{target_quarter}Q",       # "3Q"
            f"Q{target_quarter}",       # "Q3"
        ]
        
        print(f"[SmartSearch] {y_str}년 {target_quarter}분기 데이터 열 탐색 시작...")
        
        # 3. header_row를 순회 가능한 형태로 변환
        # pandas Series, 리스트, 튜플, openpyxl Cell 객체 등 모든 형태 지원
        header_items = []
        if isinstance(header_row, pd.Series):
            header_items = header_row.tolist()
        elif isinstance(header_row, (list, tuple)):
            header_items = list(header_row)
        elif hasattr(header_row, '__iter__') and not isinstance(header_row, str):
            # openpyxl Cell 객체나 다른 iterable 객체
            header_items = [cell.value if hasattr(cell, 'value') else cell for cell in header_row]
        else:
            header_items = [header_row]
        
        # 4. 헤더 행을 순회하며 연도와 분기가 모두 포함된 셀 찾기
        for idx, cell in enumerate(header_items):
            # 셀 값 추출 (다양한 타입 지원)
            if pd.isna(cell):
                continue
            
            # cell.value 속성이 있는 경우 (openpyxl 스타일)
            if hasattr(cell, 'value'):
                cell_value = cell.value
            else:
                cell_value = cell
            
            # 셀 값이 None이거나 빈 값이면 스킵
            if cell_value is None:
                continue
            
            # 5. 값 정규화: 공백 제거 후 비교 (정규표현식 수준의 유연성)
            val = str(cell_value).strip().replace(" ", "")  # 모든 공백 제거
            
            # 6. 조건 검사: 연도가 있고(2025 or 25) AND 분기 마커 중 하나가 있어야 함
            has_year = (y_str in val) or (y_short in val)
            has_quarter = any(q.replace(" ", "") in val for q in q_markers)
            
            if has_year and has_quarter:
                print(f"[SmartSearch] ✅ 발견! Index {idx}: '{cell_value}'")
                return idx
        
        # 7. 못 찾으면 에러 발생 (어설프게 진행 금지)
        # 헤더 덤프 생성 (디버깅용)
        header_dump = []
        for item in header_items[:50]:  # 최대 50개만 표시
            if pd.isna(item):
                header_dump.append("NaN")
            elif hasattr(item, 'value'):
                header_dump.append(str(item.value))
            else:
                header_dump.append(str(item))
        
        error_msg = (
            f"❌ 헤더에서 {target_year}년 {target_quarter}분기 열을 찾을 수 없습니다.\n"
            f"   (Header dump: {header_dump})"
        )
        raise ValueError(error_msg)
