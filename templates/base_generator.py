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
        self.df_cache: Dict[str, pd.DataFrame] = {}  # 시트별 DataFrame 캐시
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
            df = pd.read_excel(calculated_path, sheet_name=sheet_name, header=None)  # type: ignore
        else:
            df = pd.read_excel(xl, sheet_name=sheet_name, header=None)  # type: ignore
        
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
                cleaned = re.sub(r'[^\d\.\-\+eE]', '', value)
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
        result: List[int] = []
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
        동적 데이터 탐색 기능은 제거되었습니다.
        컬럼 인덱스는 설정에서 명시적으로 제공해야 합니다.
        """
        raise NotImplementedError(
            "동적 데이터 탐색 기능이 제거되었습니다. 설정에서 컬럼 인덱스를 지정하세요."
        )
    
    def _find_target_col_index_from_df(
        self,
        df: pd.DataFrame,
        target_year: int,
        target_quarter: int,
        require_type_match: bool = True,
        max_header_rows: Optional[int] = None
    ) -> Optional[int]:
        """
        동적 데이터 탐색 기능은 제거되었습니다.
        """
        raise NotImplementedError(
            "동적 데이터 탐색 기능이 제거되었습니다. 설정에서 컬럼 인덱스를 지정하세요."
        )
    
    def _find_target_col_index_from_row(
        self,
        header_row: Any,
        target_year: int,
        target_quarter: int
    ) -> Optional[int]:
        """
        동적 데이터 탐색 기능은 제거되었습니다.
        """
        raise NotImplementedError(
            "동적 데이터 탐색 기능이 제거되었습니다. 설정에서 컬럼 인덱스를 지정하세요."
        )
    
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
        cause_verb, result_noun, _ = get_terms(report_id, growth_rate)
        
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
                _, prev_result, _ = get_terms(report_id, prev_rate)
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
            if weight is None or weight == 0 or pd.isna(weight):
                is_major = any(major in ind.get('name', '') for major in MAJOR_INDUSTRIES)
                weight = 10 if is_major else 1
            if change_rate is None or pd.isna(change_rate):
                change_rate = 0.0
            # 기여도 = |증감률 × 가중치|
            ind['contribution'] = abs(float(change_rate) * float(weight))
        
        # 기여도 순 정렬 (내림차순)
        ranked = sorted(
            industries, 
            key=lambda x: x.get('contribution', 0), 
            reverse=True
        )
        
        return ranked[:top_n]
