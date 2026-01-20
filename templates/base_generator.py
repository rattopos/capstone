#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Base Generator 클래스

모든 Generator가 공통으로 사용하는 기능을 제공하는 기본 클래스입니다.
"""

import pandas as pd
import re
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
        self._calculated_excel_path = None
        
    def load_excel(self) -> pd.ExcelFile:
        """엑셀 파일 로드 (캐싱)"""
        if self.xl is None:
            self.xl = pd.ExcelFile(self.excel_path)
            self._xl_owner = True
        return self.xl

    def _get_calculated_excel_path(self) -> str:
        """수식 계산 로직으로 계산된 임시 파일 경로 반환 (전역 캐시 사용)."""
        from services.excel_processor import preprocess_excel
        from services.excel_cache import get_cached_calculated_path, set_cached_calculated_path
        from config.settings import TEMP_CALCULATED_DIR

        if self._calculated_excel_path and Path(self._calculated_excel_path).exists():
            return self._calculated_excel_path

        cached_path = get_cached_calculated_path(self.excel_path)
        if cached_path:
            self._calculated_excel_path = cached_path
            return cached_path

        TEMP_CALCULATED_DIR.mkdir(parents=True, exist_ok=True)
        output_path = TEMP_CALCULATED_DIR / f"{Path(self.excel_path).stem}_calculated.xlsx"
        result_path, success, _ = preprocess_excel(
            self.excel_path,
            str(output_path),
            force_calculation=True
        )

        if success and result_path:
            set_cached_calculated_path(self.excel_path, result_path)
            self._calculated_excel_path = result_path

        return result_path
    
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
        
        if '분석' in sheet_name:
            calculated_path = self._get_calculated_excel_path()
            df = pd.read_excel(calculated_path, sheet_name=sheet_name, header=None)
        else:
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
            "page_number": ""  # 페이지 번호는 더 이상 사용하지 않음 (목차 생성 중단)
            # footer info는 나중에 추가 예정
        }
    
    def find_target_col_index(
        self,
        header_row: Any,
        target_year: int,
        target_quarter: int,
        require_type_match: bool = True,
        max_header_rows: Optional[int] = None
    ) -> Optional[int]:
        """
        [스마트 헤더 탐색기] 병합된 셀(Merged Header) 문제를 완벽하게 돌파하여 정확한 열 위치를 찾아냅니다.
        (유연성 강화: 다양한 연도/분기/타입 표기법 robust 지원)
        
        Args:
            header_row: DataFrame 또는 Series/list (단일 행)
            target_year: 대상 연도
            target_quarter: 대상 분기
            require_type_match: 타입 매칭 필수 여부
            max_header_rows: 헤더 행 수 (DataFrame인 경우만 사용, None이면 기본값 5)
        """
        import re
        # 1. DataFrame인 경우: 병합된 셀 처리 로직 사용
        if isinstance(header_row, pd.DataFrame):
            return self._find_target_col_index_from_df(
                header_row, target_year, target_quarter, require_type_match, max_header_rows
            )
        
        # 2. Series나 단일 행인 경우: 기존 로직 사용 (하위 호환성)
        # 하지만 가능하면 DataFrame을 전달하는 것이 좋습니다.
        # require_type_match는 단일 행에서는 항상 True로 사용 (하위 호환성)
        # --- 유연성 강화: 다양한 헤더 표기법 robust 지원 ---
        y_str = str(target_year)
        y_short = y_str[2:] if len(y_str) == 4 else y_str
        q = str(target_quarter)
        # 다양한 분기 표기
        q_patterns = [
            f"{q}/4", f"{q}분기", f"{q}Q", f"Q{q}", f"{q}-4", f"{q} - 4", f"{q}분", f"{q}q"
        ]
        # 다양한 연도+분기 결합 패턴 (띄어쓰기, 괄호, 하이픈 등)
        year_quarter_patterns = [
            rf"{y_str}\s*[\(\-]{{0,1}}\s*{q}/4[\)\-]{{0,1}}",  # 2025(3/4), 2025-3/4
            rf"{y_str}\s*{q}분기", rf"{y_str}\s*{q}Q", rf"{y_str}\s*Q{q}",
            rf"{y_str}{q}/4", rf"{y_str}{q}Q", rf"{y_str}{q}분기",
            rf"{y_str}\s*{q}-4", rf"{y_str}\s*{q} - 4",
            rf"{y_str}\s*\({q}/4\)", rf"{y_str}\s*\({q}Q\)",
            rf"{y_str}\s*\({q}분기\)",
            rf"{y_str}\s*{q}q", rf"{y_str}{q}q",
            rf"{y_short}\s*[\(\-]{{0,1}}\s*{q}/4[\)\-]{{0,1}}",  # 25(3/4)
            rf"{y_short}\s*{q}분기", rf"{y_short}\s*{q}Q", rf"{y_short}\s*Q{q}",
            rf"{y_short}{q}/4", rf"{y_short}{q}Q", rf"{y_short}{q}분기",
            rf"{y_short}\s*{q}-4", rf"{y_short}\s*{q} - 4",
            rf"{y_short}\s*\({q}/4\)", rf"{y_short}\s*\({q}Q\)",
            rf"{y_short}\s*\({q}분기\)",
            rf"{y_short}\s*{q}q", rf"{y_short}{q}q",
        ]
        # 타입 키워드
        type_keywords = ["증감", "등락", "증감률", "지수", "고용률", "실업률", "실업자"]
        header_items = []
        if isinstance(header_row, pd.Series):
            header_items = header_row.tolist()
        elif isinstance(header_row, (list, tuple)):
            header_items = list(header_row)
        elif hasattr(header_row, '__iter__') and not isinstance(header_row, str):
            header_items = [cell.value if hasattr(cell, 'value') else cell for cell in header_row]
        else:
            header_items = [header_row]
        for idx, cell in enumerate(header_items):
            if pd.isna(cell):
                continue
            val = str(cell).strip().replace(" ", "")
            # 연도/분기 패턴 매칭
            has_yq = any(re.search(pat.replace(" ", ""), val) for pat in year_quarter_patterns)
            year_tokens = re.findall(r"\d{4}", val)
            has_year = str(target_year) in year_tokens
            if not has_year and len(y_short) == 2:
                if re.search(rf"(?<!\d){re.escape(y_short)}(?!\d)", val):
                    has_year = True
            has_quarter = any(qp.replace(" ", "") in val for qp in q_patterns)
            # 타입 키워드 매칭
            is_growth = any(k in val for k in type_keywords[:4])
            is_rate = any(k in val for k in type_keywords[4:])
            type_match = is_growth or (is_rate and not require_type_match)
            # 최종 매칭: 연도+분기 패턴 or (연도 and 분기)
            if (has_yq or (has_year and has_quarter)) and (type_match if require_type_match else True):
                print(f"[SmartSearch-유연] ✅ 발견! Index {idx}: '{cell}' (패턴매칭)")
                return idx
        # 못 찾음
        print(f"[SmartSearch-유연] ❌ 탐색 실패: 컬럼을 찾을 수 없습니다.")
        return None
    
    def _find_target_col_index_from_df(
        self,
        df: pd.DataFrame,
        target_year: int,
        target_quarter: int,
        require_type_match: bool = True,
        max_header_rows: Optional[int] = None
    ) -> Optional[int]:
        """
        DataFrame에서 병합된 셀을 고려하여 열 인덱스를 찾습니다.
        
        핵심: pandas의 ffill을 사용하여 병합된 셀을 채웁니다.
        
        Args:
            df: DataFrame
            target_year: 대상 연도
            target_quarter: 대상 분기
            require_type_match: 타입 매칭 필수 여부
            max_header_rows: 헤더 행 수 (None이면 기본값 5 사용)
        """
        # 1. 헤더 영역 추출 (config 설정 우선, 없으면 기본값 5)
        if max_header_rows is None:
            max_header_rows = 5
        max_header_rows = min(max_header_rows, len(df))
        header_df = df.iloc[:max_header_rows].copy()
        
        # 2. 정규화된 타겟 문자열 생성
        target_year_str = str(target_year).replace(" ", "")
        target_year_short = target_year_str[2:] if len(target_year_str) == 4 else target_year_str
        
        target_quarter_strs = [
            f"{target_quarter}/4",
            f"{target_quarter}분기",
            f"{target_quarter}Q",
            f"Q{target_quarter}",
            f"{target_quarter}q",
        ]
        
        print(f"\n{'='*80}")
        print(f"[SmartHeader] ═══ 헤더 탐색 시작 ═══")
        print(f"  타겟: {target_year}년 {target_quarter}분기")
        print(f"  검색 조건:")
        print(f"    - 연도: {target_year_str} 또는 {target_year_short}")
        print(f"    - 분기: {target_quarter_strs}")
        if require_type_match:
            print(f"    - 타입: '지수' 또는 '증감률'/'등락' (필수)")
        else:
            print(f"    - 타입: 선택적 (고용률/실업률 등)")
        print(f"  DataFrame 크기: {len(df)}행 × {len(df.columns)}열")
        print(f"  헤더 영역: 상위 {max_header_rows}행 사용")
        print(f"{'='*80}\n")
        
        # 3. 헤더 재구성: 병합된 셀 채우기 (강화된 버전)
        print(f"[디버그] 병합 셀 처리 시작...")
        
        # 원본 헤더 저장 (디버깅용)
        headers_original = header_df.copy()
        
        # 모든 값을 문자열로 변환하고 빈 값 처리
        # FutureWarning 방지: object 타입으로 변환
        headers = header_df.copy().astype(object)
        
        # 문자열로 변환 (NaN, None 처리)
        empty_count = 0
        for col_idx in range(len(headers.columns)):
            for row_idx in range(len(headers)):
                val = headers.iloc[row_idx, col_idx]
                if pd.isna(val) or val == '' or str(val).strip() == '':
                    headers.iloc[row_idx, col_idx] = None
                    empty_count += 1
                else:
                    headers.iloc[row_idx, col_idx] = str(val).strip()
        
        print(f"[디버그] 빈 셀 개수 (처리 전): {empty_count}개")
        
        # 좌우 병합 채우기 (axis=1): [2025년][빈칸][빈칸] -> [2025년][2025년][2025년]
        headers = headers.ffill(axis=1, limit=None)
        filled_horizontal = sum(1 for col_idx in range(len(headers.columns)) 
                                for row_idx in range(len(headers)) 
                                if headers.iloc[row_idx, col_idx] is not None and 
                                   (headers_original.iloc[row_idx, col_idx] is None or 
                                    pd.isna(headers_original.iloc[row_idx, col_idx])))
        
        # 상하 병합 채우기 (axis=0): 위 행의 값이 없으면 아래 행에서 채움
        headers = headers.ffill(axis=0, limit=None)
        
        # 역방향 채우기 (bfill): 아래에서 위로도 채우기 (상하 병합 대응)
        headers = headers.bfill(axis=0, limit=None)
        
        final_empty_count = sum(1 for col_idx in range(len(headers.columns)) 
                                for row_idx in range(len(headers)) 
                                if headers.iloc[row_idx, col_idx] is None)
        
        print(f"[디버그] 병합 처리 완료: 빈 셀 {empty_count}개 → {final_empty_count}개 (채움: {empty_count - final_empty_count}개)")
        print(f"[디버그] 좌우 병합으로 채운 셀: 약 {filled_horizontal}개\n")
        
        # 4. 전수 조사: 열 단위로 스캔 (개선된 버전)
        print(f"[디버그] 컬럼 전수 조사 시작 (총 {len(headers.columns)}개 컬럼)...\n")
        
        target_col = None
        candidate_cols = []  # 매칭 후보 컬럼 저장
        
        # 안전한 컬럼 범위 체크
        max_cols = len(headers.columns) if hasattr(headers, 'columns') else 0
        max_rows = len(headers) if hasattr(headers, '__len__') else 0
        
        if max_cols == 0 or max_rows == 0:
            print(f"[ERROR] 헤더 DataFrame이 비어있습니다: {max_rows}행 × {max_cols}열")
            return None
        
        for col_idx in range(max_cols):
            # 해당 열의 상위 행들의 텍스트를 모두 수집 (안전한 인덱스 접근)
            col_values = []
            for row_idx in range(max_rows):
                try:
                    if row_idx < len(headers) and col_idx < len(headers.columns):
                        val = headers.iloc[row_idx, col_idx]
                        if pd.notna(val) and str(val).strip() != '':
                            col_values.append(str(val).strip())
                except (IndexError, KeyError) as e:
                    # 인덱스 오류는 무시하고 계속 진행
                    continue
            
            # 열의 모든 값을 하나의 문자열로 합침 (공백 제거)
            col_text = "".join(col_values).replace(" ", "")
            
            # 원본 헤더 값도 저장 (비교용) - 안전한 인덱스 접근
            col_values_orig = []
            try:
                max_orig_rows = len(headers_original) if hasattr(headers_original, '__len__') else 0
                for row_idx in range(max_orig_rows):
                    try:
                        if row_idx < len(headers_original) and col_idx < len(headers_original.columns):
                            val = headers_original.iloc[row_idx, col_idx]
                            if pd.notna(val) and str(val).strip() != '':
                                col_values_orig.append(str(val).strip())
                    except (IndexError, KeyError):
                        continue
            except Exception as e:
                print(f"[WARNING] 원본 헤더 값 수집 중 오류: {e}")
            
            # 5. 조건 검사 (개선된 버전) - 상세 디버깅
            # A. 연도 확인 (전체 연도 또는 단축형, 경계 고려)
            has_year = False
            year_match_reason = []
            
            # 4자리 연도 토큰 확인
            years_in_text = re.findall(r'\d{4}', col_text)
            if years_in_text and any(str(target_year) == y for y in years_in_text):
                has_year = True
                year_match_reason.append(f"4자리 숫자 '{target_year}' 발견")
            # 2자리 연도는 숫자 경계가 있는 경우만 허용
            if not has_year and len(target_year_short) == 2:
                if re.search(rf"(?<!\d){re.escape(target_year_short)}(?!\d)", col_text):
                    has_year = True
                    year_match_reason.append(f"'{target_year_short}' 경계 포함")
            
            # B. 분기 확인 (더 유연한 매칭)
            has_quarter = False
            quarter_match_reason = []
            
            for q_str in target_quarter_strs:
                token = q_str.replace(" ", "")
                if token and token in col_text:
                    has_quarter = True
                    quarter_match_reason.append(f"'{q_str}' 형식 매칭")
                    break
            
            # C. 컬럼 타입 확인 (더 유연한 매칭)
            # 고용률/실업률 등은 타입 필터링을 선택적으로 적용
            is_growth = "증감" in col_text or "등락" in col_text or "증감률" in col_text
            is_index = "지수" in col_text
            is_rate = "고용률" in col_text or "실업률" in col_text or "실업자" in col_text
            type_match = is_growth or is_index or (is_rate and not require_type_match)
            type_match_reason = []
            
            if "증감" in col_text:
                type_match_reason.append("'증감' 포함")
            if "등락" in col_text:
                type_match_reason.append("'등락' 포함")
            if "증감률" in col_text:
                type_match_reason.append("'증감률' 포함")
            if "지수" in col_text:
                type_match_reason.append("'지수' 포함")
            if is_rate and not require_type_match:
                if "고용률" in col_text:
                    type_match_reason.append("'고용률' 포함 (타입 필터링 선택적)")
                    type_match = True  # 명시적으로 True 설정
                if "실업률" in col_text or "실업자" in col_text:
                    type_match_reason.append("'실업률/실업자' 포함 (타입 필터링 선택적)")
                    type_match = True  # 명시적으로 True 설정
            
            # 매칭 결과 출력 (처음 10개 컬럼 또는 매칭 조건 일부 만족 시)
            should_print = (col_idx < 10) or has_year or has_quarter or type_match
            
            if should_print:
                # require_type_match가 False면 타입 매칭 없이도 연도+분기만 맞으면 매칭
                final_match_check = has_year and has_quarter and (type_match if require_type_match else True)
                match_status = "✅ 매칭" if final_match_check else "❌ 불일치"
                print(f"  컬럼 {col_idx:3d}: {match_status}")
                print(f"    원본: {' | '.join(col_values_orig) if col_values_orig else '(빈)'}")
                print(f"    병합처리: {' | '.join(col_values) if col_values else '(빈)'}")
                print(f"    텍스트: '{col_text[:60]}{'...' if len(col_text) > 60 else ''}'")
                print(f"    연도: {has_year} {year_match_reason if year_match_reason else '(불일치)'}")
                print(f"    분기: {has_quarter} {quarter_match_reason if quarter_match_reason else '(불일치)'}")
                print(f"    타입: {type_match} ({type_match_reason if type_match_reason else '불일치'})")
                if has_year or has_quarter or type_match:
                    candidate_cols.append({
                        'col_idx': col_idx,
                        'has_year': has_year,
                        'has_quarter': has_quarter,
                        'type_match': type_match,
                        'text': col_text[:100],
                        'values': col_values
                    })
                print()
            
            # 최종 매칭 확인
            # require_type_match가 False면 타입 매칭 없이도 연도+분기만 맞으면 OK
            final_match = has_year and has_quarter and (type_match if require_type_match else True)
            if final_match:
                target_col = col_idx
                col_display = " | ".join(col_values[:3]) if col_values else "빈값"
                print(f"\n{'='*80}")
                print(f"✅ [동적탐색 성공] 컬럼 {target_col} 매칭!")
                print(f"   원본 헤더: {' | '.join(col_values_orig) if col_values_orig else '(빈)'}")
                print(f"   병합 처리 후: {' | '.join(col_values) if col_values else '(빈)'}")
                print(f"   전체 텍스트: '{col_text}'")
                print(f"   매칭 이유:")
                print(f"     - 연도: {', '.join(year_match_reason)}")
                print(f"     - 분기: {', '.join(quarter_match_reason)}")
                print(f"     - 타입: {', '.join(type_match_reason)}")
                print(f"{'='*80}\n")
                return target_col
        
        # 6. 실패 시 (최대 강화된 디버깅 정보 출력)
        if target_col is None:
            print(f"\n{'='*80}")
            print(f"❌❌❌ [탐색 실패] {target_year}년 {target_quarter}분기 데이터 열을 찾을 수 없습니다! ❌❌❌")
            print(f"{'='*80}\n")
            
            # 검색 조건 재확인
            print(f"[검색 조건 재확인]")
            print(f"  - 타겟 연도: {target_year} ({target_year_str} 또는 {target_year_short})")
            print(f"  - 타겟 분기: {target_quarter} (형식: {target_quarter_strs})")
            if require_type_match:
                print(f"  - 필수 타입: '지수' 또는 '증감률'/'등락'/'증감'")
            else:
                print(f"  - 타입: 선택적 (고용률/실업률 등은 타입 필터링 없이 연도+분기만 확인)")
            print()
            
            # 후보 컬럼 분석
            print(f"[후보 컬럼 분석] (부분 매칭된 컬럼들)")
            if candidate_cols:
                print(f"  총 {len(candidate_cols)}개 컬럼이 부분적으로 매칭되었습니다:\n")
                for i, cand in enumerate(candidate_cols[:10], 1):  # 최대 10개만 출력
                    match_summary = []
                    if cand['has_year']:
                        match_summary.append("✅연도")
                    else:
                        match_summary.append("❌연도")
                    if cand['has_quarter']:
                        match_summary.append("✅분기")
                    else:
                        match_summary.append("❌분기")
                    if cand['type_match']:
                        match_summary.append("✅타입")
                    else:
                        match_summary.append("❌타입")
                    
                    print(f"  [{i}] 컬럼 {cand['col_idx']:3d}: {' | '.join(match_summary)}")
                    print(f"      텍스트: '{cand['text'][:80]}{'...' if len(cand['text']) > 80 else ''}'")
                    print(f"      값: {cand['values'][:5]}")
                    print()
            else:
                print(f"  ⚠️ 부분 매칭된 컬럼이 하나도 없습니다!")
                print(f"  → 헤더 구조가 예상과 완전히 다를 수 있습니다.\n")
            
            # 원본 헤더 vs 병합 처리 후 비교 (상위 20개 컬럼)
            print(f"[헤더 비교 분석] (상위 20개 컬럼)")
            print(f"{'='*80}")
            print(f"{'컬럼':<6} {'원본 헤더 (병합 전)':<50} {'병합 처리 후':<50}")
            print(f"{'-'*6} {'-'*50} {'-'*50}")
            
            # 안전한 범위 체크
            max_compare_cols = min(20, len(headers.columns) if hasattr(headers, 'columns') else 0)
            max_compare_rows = len(headers_original) if hasattr(headers_original, '__len__') else 0
            
            for col_idx in range(max_compare_cols):
                # 원본 헤더 값 (안전한 인덱스 접근)
                orig_values = []
                try:
                    for row_idx in range(max_compare_rows):
                        try:
                            if row_idx < len(headers_original) and col_idx < len(headers_original.columns):
                                val = headers_original.iloc[row_idx, col_idx]
                                if pd.notna(val) and str(val).strip() != '':
                                    orig_values.append(str(val).strip())
                        except (IndexError, KeyError):
                            continue
                except Exception as e:
                    print(f"[WARNING] 원본 헤더 비교 중 오류 (컬럼 {col_idx}): {e}")
                
                orig_display = ' | '.join(orig_values[:2]) if orig_values else '(빈)'
                
                # 병합 처리 후 헤더 값
                merged_values = []
                for row_idx in range(len(headers)):
                    val = headers.iloc[row_idx, col_idx]
                    if pd.notna(val) and str(val).strip() != '':
                        merged_values.append(str(val).strip())
                merged_display = ' | '.join(merged_values[:2]) if merged_values else '(빈)'
                
                # 매칭 상태 표시
                col_text_check = "".join(merged_values).replace(" ", "")
                has_y = (target_year_str in col_text_check or target_year_short in col_text_check)
                has_q = str(target_quarter) in col_text_check
                # 분기 형식도 확인 ("3/4" 형식)
                for q_str in target_quarter_strs:
                    if "/" in q_str:
                        q_normalized = q_str.replace(" ", "")
                        if q_normalized in col_text_check:
                            has_q = True
                            break
                    else:
                        q_clean = q_str.replace(" ", "").replace("/", "")
                        if q_clean in col_text_check:
                            has_q = True
                            break
                has_t = ("지수" in col_text_check or "증감" in col_text_check or "등락" in col_text_check)
                # 고용률/실업률도 타입으로 인정 (require_type_match가 False인 경우)
                if not require_type_match:
                    has_t = has_t or ("고용률" in col_text_check or "실업률" in col_text_check or "실업자" in col_text_check)
                
                status = ""
                # require_type_match가 False면 타입 없이도 연도+분기만 맞으면 완전매칭
                if require_type_match:
                    if has_y and has_q and has_t:
                        status = " ✅ 완전매칭"
                    elif has_y or has_q or has_t:
                        status = " ⚠️ 부분매칭"
                else:
                    if has_y and has_q:
                        status = " ✅ 완전매칭"
                    elif has_y or has_q:
                        status = " ⚠️ 부분매칭"
                
                print(f"{col_idx:4d}  {orig_display[:48]:<50} {merged_display[:48]:<50}{status}")
            
            print(f"{'='*80}\n")
            
            # 원본 헤더 상세 덤프 (상위 30개 컬럼, 모든 행)
            print(f"[원본 헤더 상세 덤프] (상위 30개 컬럼, 모든 {len(headers_original)}행)")
            print(f"{'='*80}")
            print(f"컬럼 번호: ", end="")
            for col_idx in range(min(30, len(headers_original.columns))):
                print(f"{col_idx:3d} ", end="")
            print()
            print(f"{'-'*120}")
            
            for row_idx in range(len(headers_original)):
                print(f"행 {row_idx:2d}:  ", end="")
                for col_idx in range(min(30, len(headers_original.columns))):
                    val = headers_original.iloc[row_idx, col_idx]
                    if pd.notna(val) and str(val).strip() != '':
                        val_str = str(val).strip()[:8]  # 최대 8자
                        print(f"{val_str:>8} ", end="")
                    else:
                        print(f"{'[빈]':>8} ", end="")
                print()
            print(f"{'='*80}\n")
            
            # 병합 처리 후 헤더 상세 덤프
            print(f"[병합 처리 후 헤더 상세 덤프] (상위 30개 컬럼, 모든 {len(headers)}행)")
            print(f"{'='*80}")
            print(f"컬럼 번호: ", end="")
            for col_idx in range(min(30, len(headers.columns))):
                print(f"{col_idx:3d} ", end="")
            print()
            print(f"{'-'*120}")
            
            for row_idx in range(len(headers)):
                print(f"행 {row_idx:2d}:  ", end="")
                for col_idx in range(min(30, len(headers.columns))):
                    val = headers.iloc[row_idx, col_idx]
                    if pd.notna(val) and str(val).strip() != '':
                        val_str = str(val).strip()[:8]  # 최대 8자
                        print(f"{val_str:>8} ", end="")
                    else:
                        print(f"{'[빈]':>8} ", end="")
                print()
            print(f"{'='*80}\n")
            
            # 원인 분석 리포트
            print(f"[원인 분석 리포트]")
            print(f"{'='*80}")
            
            # 1. 연도만 매칭된 컬럼 찾기
            year_only_cols = [c for c in candidate_cols if c['has_year'] and not c['has_quarter']]
            if year_only_cols:
                print(f"⚠️  연도만 매칭된 컬럼 {len(year_only_cols)}개 발견:")
                for cand in year_only_cols[:5]:
                    print(f"   - 컬럼 {cand['col_idx']}: '{cand['text'][:60]}' (분기 불일치)")
                print()
            
            # 2. 분기만 매칭된 컬럼 찾기
            quarter_only_cols = [c for c in candidate_cols if c['has_quarter'] and not c['has_year']]
            if quarter_only_cols:
                print(f"⚠️  분기만 매칭된 컬럼 {len(quarter_only_cols)}개 발견:")
                for cand in quarter_only_cols[:5]:
                    print(f"   - 컬럼 {cand['col_idx']}: '{cand['text'][:60]}' (연도 불일치)")
                print()
            
            # 3. 타입만 매칭된 컬럼 찾기
            type_only_cols = [c for c in candidate_cols if c['type_match'] and not (c['has_year'] and c['has_quarter'])]
            if type_only_cols:
                print(f"⚠️  타입만 매칭된 컬럼 {len(type_only_cols)}개 발견:")
                for cand in type_only_cols[:5]:
                    print(f"   - 컬럼 {cand['col_idx']}: '{cand['text'][:60]}' (연도/분기 불일치)")
                print()
            
            # 4. 완전 불일치인 경우
            if not candidate_cols:
                print(f"❌  모든 컬럼이 완전 불일치입니다.")
                print(f"   가능한 원인:")
                print(f"   1. 헤더 구조가 예상과 완전히 다름")
                print(f"   2. 시트 선택이 잘못됨 (집계 시트 vs 분석 시트)")
                print(f"   3. 연도/분기 형식이 예상과 다름")
                print(f"   4. 병합 셀 처리가 제대로 되지 않음")
                print()
            
            # 5. 헤더 통계
            print(f"[헤더 통계]")
            total_cells = len(headers_original) * len(headers_original.columns)
            empty_cells_before = sum(1 for col_idx in range(len(headers_original.columns)) 
                                     for row_idx in range(len(headers_original))
                                     if pd.isna(headers_original.iloc[row_idx, col_idx]) or 
                                        str(headers_original.iloc[row_idx, col_idx]).strip() == '')
            empty_cells_after = sum(1 for col_idx in range(len(headers.columns)) 
                                   for row_idx in range(len(headers))
                                   if headers.iloc[row_idx, col_idx] is None)
            
            print(f"  - 총 셀 수: {total_cells}개")
            print(f"  - 빈 셀 (처리 전): {empty_cells_before}개 ({empty_cells_before/total_cells*100:.1f}%)")
            print(f"  - 빈 셀 (처리 후): {empty_cells_after}개 ({empty_cells_after/total_cells*100:.1f}%)")
            print(f"  - 채운 셀: {empty_cells_before - empty_cells_after}개")
            print()
            
            print(f"{'='*80}")
            print(f"❌ 탐색 실패: 컬럼을 찾을 수 없습니다.")
            print(f"{'='*80}\n")
            
            return None
    
    def _find_target_col_index_from_row(
        self,
        header_row: Any,
        target_year: int,
        target_quarter: int
    ) -> Optional[int]:
        """
        단일 행에서 열 인덱스를 찾습니다 (하위 호환성용).
        
        주의: 병합된 셀을 제대로 처리하지 못할 수 있으므로,
        가능하면 DataFrame을 전달하는 것을 권장합니다.
        """
        # 연도 문자열 준비
        y_str = str(target_year)
        y_short = y_str[2:] if len(y_str) == 4 else y_str
        
        # 분기 표현식
        q_markers = [
            f"{target_quarter}/4",
            f"{target_quarter}분기",
            f"{target_quarter}Q",
            f"Q{target_quarter}",
        ]
        
        # header_row를 순회 가능한 형태로 변환
        header_items = []
        if isinstance(header_row, pd.Series):
            header_items = header_row.tolist()
        elif isinstance(header_row, (list, tuple)):
            header_items = list(header_row)
        elif hasattr(header_row, '__iter__') and not isinstance(header_row, str):
            header_items = [cell.value if hasattr(cell, 'value') else cell for cell in header_row]
        else:
            header_items = [header_row]
        
        # 헤더 행을 순회하며 연도와 분기가 모두 포함된 셀 찾기
        for idx, cell in enumerate(header_items):
            if pd.isna(cell):
                continue
            
            if hasattr(cell, 'value'):
                cell_value = cell.value
            else:
                cell_value = cell
            
            if cell_value is None:
                continue
            
            val = str(cell_value).strip().replace(" ", "")
            
            has_year = (y_str in val) or (y_short in val)
            has_quarter = any(q.replace(" ", "") in val for q in q_markers)
            is_growth = "증감" in val or "등락" in val
            
            if has_year and has_quarter and is_growth:
                print(f"[SmartSearch] ✅ 발견! Index {idx}: '{cell_value}'")
                return idx
        
        # 못 찾음
        return None
    
    # ============================================================================
    # [문서 2] 나레이션 생성 엔진: 4가지 패턴 (Pattern Engine)
    # ============================================================================
    
    def select_narrative_pattern(
        self,
        growth_rate: float,
        prev_rate: Optional[float] = None,
        has_contrast_industries: bool = False
    ) -> str:
        """
        [문서 2] 4가지 나레이션 패턴 중 하나 선택
        
        패턴 우선순위:
        1. 패턴 C: 보합 (growth_rate == 0.0)
        2. 패턴 D: 방향 전환 (전분기와 부호 다름)
        3. 패턴 B: 역접 (상반된 업종 혼재)
        4. 패턴 A: 순접 (기본)
        
        Args:
            growth_rate: 현재 분기 증감률
            prev_rate: 전분기 증감률 (선택)
            has_contrast_industries: 상반된 업종 혼재 여부
        
        Returns:
            'pattern_a' | 'pattern_b' | 'pattern_c' | 'pattern_d'
        """
        # 패턴 C: 보합 (최우선)
        if abs(growth_rate) < 0.01:
            return 'pattern_c'
        
        # 패턴 D: 방향 전환 (전분기 데이터 있을 때)
        if prev_rate is not None and abs(prev_rate) > 0.01:
            # 부호 변경 체크: (+ → -) 또는 (- → +)
            if (prev_rate > 0 and growth_rate < 0) or \
               (prev_rate < 0 and growth_rate > 0):
                return 'pattern_d'
        
        # 패턴 B: 역접 (상반된 업종 혼재)
        if has_contrast_industries:
            return 'pattern_b'
        
        # 패턴 A: 순접 (기본)
        return 'pattern_a'
    
    def generate_narrative(
        self,
        pattern: str,
        region: str,
        growth_rate: float,
        prev_rate: Optional[float],
        main_industries: List[str],
        contrast_industries: Optional[List[str]] = None,
        report_id: str = 'manufacturing'
    ) -> str:
        """
        [문서 2] 선택된 패턴에 따라 나레이션 생성
        
        엄격한 어휘 매핑 준수:
        - Type A (물량): 증가/감소, 늘어/줄어
        - Type B (가격): 상승/하락, 올라/내려
        
        Args:
            pattern: 'pattern_a' | 'pattern_b' | 'pattern_c' | 'pattern_d'
            region: 지역명
            growth_rate: 증감률
            prev_rate: 전분기 증감률
            main_industries: 주요 업종 리스트
            contrast_industries: 반대 업종 리스트 (패턴 B, C용)
            report_id: 보고서 ID (어휘 타입 결정)
        
        Returns:
            생성된 나레이션 문자열
        """
        # 1. 어휘 선택 (엄격한 매핑)
        try:
            from utils.text_utils import get_terms, get_josa
        except ImportError:
            # 상대 import 시도
            import sys
            from pathlib import Path
            sys.path.insert(0, str(Path(__file__).parent.parent))
            from utils.text_utils import get_terms, get_josa
        
        # 조사 선택
        josa = get_josa(region, "은/는")
        
        # 주요 업종 문자열
        industries_str = ', '.join(main_industries[:3]) if main_industries else "주요 업종"
        
        # 어휘 선택
        cause_verb, result_noun = get_terms(report_id, growth_rate)
        
        # 2. 패턴별 문장 생성
        if pattern == 'pattern_a':
            # 패턴 A (순접): "[지역]은 [업종] 등이 늘어 전년동분기대비 5.2% 증가."
            if cause_verb is None:  # 보합 예외 처리
                return f"{region}{josa} 전년동분기대비 {result_noun}."
            
            return (
                f"{region}{josa} {industries_str} 등이 {cause_verb} "
                f"전년동분기대비 {abs(growth_rate):.1f}% {result_noun}."
            )
        
        elif pattern == 'pattern_b':
            # 패턴 B (역접): "[지역]은 [반대업종] 등이 줄었으나, [주요업종] 등이 늘어 5.2% 증가."
            contrast_str = ', '.join(contrast_industries[:2]) if contrast_industries else "일부 업종"
            
            # 반대 방향 어휘
            if growth_rate > 0:
                opposite_verb = "줄었으나"
            else:
                opposite_verb = "늘었으나"
            
            return (
                f"{region}{josa} {contrast_str} 등이 {opposite_verb}, "
                f"{industries_str} 등이 {cause_verb} "
                f"전년동분기대비 {abs(growth_rate):.1f}% {result_noun}."
            )
        
        elif pattern == 'pattern_c':
            # 패턴 C (보합): "[지역]은 [증가업종] 등은 늘었으나, [감소업종] 등이 줄어 보합."
            inc_str = ', '.join(main_industries[:2]) if main_industries else "일부 업종"
            dec_str = ', '.join(contrast_industries[:2]) if contrast_industries else "일부 업종"
            
            return (
                f"{region}{josa} {inc_str} 등은 늘었으나, "
                f"{dec_str} 등이 줄어 전년동분기대비 {result_noun}."
            )
        
        elif pattern == 'pattern_d':
            # 패턴 D (추세): "[지역]은 전분기 증가하였으나, 이번 분기 [업종] 등이 줄어 3.1% 감소."
            if prev_rate is not None:
                _, prev_result = get_terms(report_id, prev_rate)
            else:
                prev_result = "변동"
            
            return (
                f"{region}{josa} 전분기 {prev_result}하였으나, "
                f"이번 분기 {industries_str} 등이 {cause_verb} "
                f"{abs(growth_rate):.1f}% {result_noun}."
            )
        
        # Fallback (도달 불가능)
        return f"{region}{josa} [패턴 오류: {pattern}]"
    
    # ============================================================================
    # [문서 3] 기여도 기반 업종 정렬 (Weighted Contribution)
    # ============================================================================
    
    def rank_by_contribution(
        self,
        industries: List[Dict[str, Any]],
        top_n: int = 3
    ) -> List[Dict[str, Any]]:
        """
        [문서 3] 기여도 = |증감률 × 가중치| 순으로 업종 정렬
        
        Args:
            industries: [{'name': str, 'change_rate': float, 'weight': float}, ...]
            top_n: 상위 N개 선정
        
        Returns:
            기여도 순 정렬된 리스트
        
        Note:
            가중치 없으면 주요 업종 여부로 가산점 (Fallback)
        """
        # 주요 업종 목록 (가중치 없을 때 fallback)
        MAJOR_INDUSTRIES = [
            '반도체', '전자부품', '자동차', '기계', 
            '화학', '전기장비', '1차금속', '의약품'
        ]
        
        for ind in industries:
            weight = self.safe_float(ind.get('weight'), 0)
            change_rate = self.safe_float(ind.get('change_rate'), 0)
            
            # Fallback: 가중치 없으면 주요 업종 여부로 가산점
            if weight == 0 or pd.isna(weight):
                is_major = any(major in ind.get('name', '') for major in MAJOR_INDUSTRIES)
                weight = 10 if is_major else 1
            
            # 기여도 = |증감률 × 가중치|
            ind['contribution'] = abs(change_rate * weight)
        
        # 기여도 순 정렬 (내림차순)
        ranked = sorted(
            industries, 
            key=lambda x: x.get('contribution', 0), 
            reverse=True
        )
        
        return ranked[:top_n]
