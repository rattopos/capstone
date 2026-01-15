# -*- coding: utf-8 -*-
"""
엑셀 파일 관련 유틸리티 함수
"""

import importlib.util
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Optional
import pandas as pd

from config.settings import TEMPLATES_DIR


def get_previous_quarter():
    """현재 시점의 바로 이전 분기를 계산하여 반환
    
    Returns:
        (year, quarter) 튜플
        예: 오늘이 2026년 1월이면 → (2025, 4)
        예: 오늘이 2026년 5월이면 → (2026, 1)
    """
    now = datetime.now()
    current_year = now.year
    current_month = now.month
    current_quarter = (current_month - 1) // 3 + 1
    
    # 이전 분기 계산
    if current_quarter == 1:
        # 1분기면 작년 4분기
        return current_year - 1, 4
    else:
        # 2, 3, 4분기면 올해의 이전 분기
        return current_year, current_quarter - 1


def get_period_context(target_year: int, target_quarter: int) -> Dict[str, Any]:
    """
    분석 대상 연도/분기를 기준으로 비교 시점들을 수학적으로 계산하는 함수
    
    입력: 분석 대상 연도(target_year), 분기(target_quarter)
    반환: 모든 비교 시점 정보를 담은 딕셔너리
    
    Args:
        target_year: 분석 대상 연도 (예: 2025)
        target_quarter: 분석 대상 분기 (1-4)
    
    Returns:
        {
            'target_year': int,           # 분석 대상 연도
            'target_quarter': int,        # 분석 대상 분기
            'target_period': str,         # "2025.3/4" 형식
            'prev_q_year': int,           # 전분기 연도
            'prev_q': int,                # 전분기 (1-4)
            'prev_q_period': str,         # "2025.2/4" 형식
            'prev_y_year': int,           # 전년동기 연도
            'prev_y_quarter': int,        # 전년동기 분기
            'prev_y_period': str,         # "2024.3/4" 형식
            'recent_5_years': list,       # [2021, 2022, 2023, 2024, 2025]
            'recent_5_quarters': list,    # ["2024.1/4", "2024.2/4", "2024.3/4", "2024.4/4", "2025.1/4"] (역순)
            'quarter_keys': dict,         # {'2025_3Q': '2025.3/4', '2024_3Q': '2024.3/4', ...}
            'target_key': str,            # '2025_3Q' (현재 분기)
            'prev_quarter_key': str,      # '2025_2Q' (전분기)
            'prev_year_key': str,         # '2024_3Q' (전년동기)
        }
    """
    # 전분기(QoQ) 계산: 분기가 1이면 작년 4분기, 아니면 올해 (분기-1)
    if target_quarter == 1:
        prev_q_year = target_year - 1
        prev_q = 4
    else:
        prev_q_year = target_year
        prev_q = target_quarter - 1
    
    # 전년동기(YoY) 계산: (연도-1) 동일 분기
    prev_y_year = target_year - 1
    prev_y_quarter = target_quarter
    
    # 최근 5년 리스트: [target_year-4, ..., target_year]
    recent_5_years = list(range(target_year - 4, target_year + 1))
    
    # 최근 5분기 리스트: 역순으로 계산
    recent_5_quarters = []
    for i in range(4, -1, -1):  # 4, 3, 2, 1, 0
        q_offset = i
        if q_offset == 0:
            # 현재 분기
            q_year = target_year
            q_num = target_quarter
        else:
            # 과거 분기
            total_quarters_back = q_offset
            q_year = target_year
            q_num = target_quarter
            
            # 분기를 역산
            for _ in range(total_quarters_back):
                if q_num == 1:
                    q_num = 4
                    q_year -= 1
                else:
                    q_num -= 1
        
        recent_5_quarters.append(f"{q_year}.{q_num}/4")
    
    # Quarter Key 형식 생성 (예: '2025_3Q')
    def make_quarter_key(year: int, quarter: int) -> str:
        return f"{year}_{quarter}Q"
    
    target_key = make_quarter_key(target_year, target_quarter)
    prev_q_key = make_quarter_key(prev_q_year, prev_q)
    prev_y_key = make_quarter_key(prev_y_year, prev_y_quarter)
    
    # Quarter Keys 딕셔너리 (최근 5분기)
    quarter_keys = {}
    for period_str in recent_5_quarters:
        year_str, q_str = period_str.split('.')
        q_num = int(q_str.split('/')[0])
        year = int(year_str)
        quarter_keys[make_quarter_key(year, q_num)] = period_str
    
    return {
        'target_year': target_year,
        'target_quarter': target_quarter,
        'target_period': f"{target_year}.{target_quarter}/4",
        'prev_q_year': prev_q_year,
        'prev_q': prev_q,
        'prev_q_period': f"{prev_q_year}.{prev_q}/4",
        'prev_y_year': prev_y_year,
        'prev_y_quarter': prev_y_quarter,
        'prev_y_period': f"{prev_y_year}.{prev_y_quarter}/4",
        'recent_5_years': recent_5_years,
        'recent_5_quarters': recent_5_quarters,
        'quarter_keys': quarter_keys,
        'target_key': target_key,
        'prev_quarter_key': prev_q_key,
        'prev_year_key': prev_y_key,
    }


def find_column_by_header(df: pd.DataFrame, header_pattern: str, search_rows: int = 5) -> Optional[int]:
    """
    엑셀 헤더를 동적으로 찾아 컬럼 인덱스를 반환
    
    엑셀 파일의 컬럼 헤더 위치가 바뀔 수 있으므로, 인덱스 번호를 하드코딩하는 대신
    헤더 텍스트를 찾아 매핑합니다.
    
    Args:
        df: pandas DataFrame (엑셀 시트)
        header_pattern: 찾을 헤더 패턴 (예: "2025.3/4", "2025", "3/4")
        search_rows: 헤더를 찾을 상단 행 수 (기본값: 5)
    
    Returns:
        컬럼 인덱스 (0-based) 또는 None (찾지 못한 경우)
    """
    import re
    
    # 상단 행들을 스캔
    for row_idx in range(min(search_rows, len(df))):
        row = df.iloc[row_idx]
        
        # 모든 컬럼을 순회하며 패턴 찾기
        for col_idx in range(len(row)):
            cell_value = row.iloc[col_idx] if hasattr(row, 'iloc') else row[col_idx]
            
            if pd.notna(cell_value):
                cell_str = str(cell_value).strip()
                
                # 정확한 매칭 또는 패턴 포함 확인
                if header_pattern in cell_str:
                    return col_idx
                
                # 정규식 패턴 매칭 (예: "2025.3/4", "2025 3/4" 등)
                pattern_escaped = re.escape(header_pattern)
                if re.search(pattern_escaped, cell_str, re.IGNORECASE):
                    return col_idx
    
    return None


def find_columns_by_period(df: pd.DataFrame, period_context: Dict[str, Any], search_rows: int = 5) -> Dict[str, Optional[int]]:
    """
    period_context를 사용하여 여러 컬럼을 한 번에 찾기
    
    Args:
        df: pandas DataFrame (엑셀 시트)
        period_context: get_period_context()로 생성한 딕셔너리
        search_rows: 헤더를 찾을 상단 행 수
    
    Returns:
        {
            'target_col': int,      # 현재 분기 컬럼 인덱스
            'prev_q_col': int,      # 전분기 컬럼 인덱스
            'prev_y_col': int,      # 전년동기 컬럼 인덱스
        }
    """
    target_period = period_context['target_period']
    prev_q_period = period_context['prev_q_period']
    prev_y_period = period_context['prev_y_period']
    
    # 여러 패턴으로 시도 (더 유연한 매칭)
    target_patterns = [
        target_period,  # "2025.3/4"
        f"{period_context['target_year']}.{period_context['target_quarter']}/4",
        f"{period_context['target_year']} {period_context['target_quarter']}/4",
    ]
    
    prev_q_patterns = [
        prev_q_period,
        f"{period_context['prev_q_year']}.{period_context['prev_q']}/4",
    ]
    
    prev_y_patterns = [
        prev_y_period,
        f"{period_context['prev_y_year']}.{period_context['prev_y_quarter']}/4",
    ]
    
    result = {
        'target_col': None,
        'prev_q_col': None,
        'prev_y_col': None,
    }
    
    # 현재 분기 찾기
    for pattern in target_patterns:
        col_idx = find_column_by_header(df, pattern, search_rows)
        if col_idx is not None:
            result['target_col'] = col_idx
            break
    
    # 전분기 찾기
    for pattern in prev_q_patterns:
        col_idx = find_column_by_header(df, pattern, search_rows)
        if col_idx is not None:
            result['prev_q_col'] = col_idx
            break
    
    # 전년동기 찾기
    for pattern in prev_y_patterns:
        col_idx = find_column_by_header(df, pattern, search_rows)
        if col_idx is not None:
            result['prev_y_col'] = col_idx
            break
    
    return result


def load_generator_module(generator_name):
    """동적으로 generator 모듈 로드"""
    generator_path = TEMPLATES_DIR / generator_name
    if not generator_path.exists():
        return None
    
    spec = importlib.util.spec_from_file_location(
        generator_name.replace('.py', ''),
        str(generator_path)
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def extract_year_quarter_from_data(excel_path, default_year=None, default_quarter=None):
    """파일명에서 연도/분기 추출
    
    파일명에서만 연도/분기 정보를 추출합니다.
    추출 실패 시 기본값을 반환하거나 예외를 발생시킵니다.
    
    Args:
        excel_path: 엑셀 파일 경로
        default_year: 추출 실패 시 사용할 기본 연도 (None이면 예외 발생)
        default_quarter: 추출 실패 시 사용할 기본 분기 (None이면 예외 발생)
    
    Returns:
        (year, quarter) 튜플
    """
    import re
    from pathlib import Path
    
    try:
        # 파일명에서 추출
        filename = Path(excel_path).stem
        print(f"[데이터에서 연도/분기 추출] 파일명: {filename}")
        
        filename_patterns = [
            r'(\d{4})년[_\s-]*(\d)분기',      # 2025년 3분기, 2025년_3분기, 2025년-3분기
            r'(\d{2})년[_\s-]*(\d)분기',      # 25년 3분기, 25년_3분기
            r'(\d{4})[년\s_-]+(\d)분기',     # 2025년_3분기, 2025 3분기
            r'(\d{2})[년\s_-]+(\d)분기',     # 25년_3분기, 25 3분기
            r'(\d{4})[_\s-](\d)[분]',        # 2025_3, 2025 3분
            r'(\d{2})[_\s-](\d)[분]',        # 25_3, 25 3분
        ]
        
        for pattern in filename_patterns:
            match = re.search(pattern, filename)
            if match:
                year_str = match.group(1)
                quarter = int(match.group(2))
                
                if len(year_str) == 2:
                    year = 2000 + int(year_str)
                else:
                    year = int(year_str)
                
                print(f"[데이터에서 연도/분기 추출] ✅ 파일명에서 추출: {year}년 {quarter}분기")
                return year, quarter
        
        # 파일명에서 찾지 못한 경우
        if default_year is not None and default_quarter is not None:
            print(f"[경고] 파일명에서 연도/분기 정보를 찾을 수 없습니다. 기본값 사용: {default_year}년 {default_quarter}분기")
            return default_year, default_quarter
        else:
            raise ValueError(f"파일명에서 연도/분기 정보를 찾을 수 없습니다. 파일명에 '2025년 3분기' 형식의 정보가 포함되어 있는지 확인하세요.")
    
    except ValueError:
        # 기본값이 있으면 반환, 없으면 예외 전파
        if default_year is not None and default_quarter is not None:
            print(f"[경고] 연도/분기 추출 실패. 기본값 사용: {default_year}년 {default_quarter}분기")
            return default_year, default_quarter
        raise
    except Exception as e:
        print(f"[오류] 데이터에서 연도/분기 추출 오류: {e}")
        if default_year is not None and default_quarter is not None:
            print(f"[경고] 기본값 사용: {default_year}년 {default_quarter}분기")
            return default_year, default_quarter
        raise ValueError(f"데이터에서 연도/분기 추출 중 오류: {str(e)}")


def extract_year_quarter_from_excel(filepath, default_year=None, default_quarter=None):
    """엑셀 파일에서 최신 연도와 분기 추출 (분석표용) - 전면 재작성 버전
    
    우선순위:
    1. 파일명 분석 (가장 정확)
    2. '이용관련' 시트의 K17(연도), M17(분기) 셀 확인
    3. 데이터 시트 헤더(2~3행)에서 '20xx' 패턴 찾기
    4. 안전장치: 현재 날짜 기준 이전 분기 또는 기본값 반환
    
    Args:
        filepath: 엑셀 파일 경로
        default_year: 추출 실패 시 사용할 기본 연도 (None이면 안전장치 사용)
        default_quarter: 추출 실패 시 사용할 기본 분기 (None이면 안전장치 사용)
    
    Returns:
        (year, quarter) 튜플
    """
    import re
    from datetime import datetime
    
    # ========================================
    # 1순위: 파일명 분석 (가장 정확)
    # ========================================
    try:
        filename = Path(filepath).stem
        print(f"[연도/분기 추출] 1순위: 파일명 분석 시작 - '{filename}'")
        
        # 파일명 패턴 (다양한 형식 지원)
        filename_patterns = [
            r'(\d{4})년[_\s-]*(\d)분기',      # 2025년 3분기, 2025년_3분기, 2025년-3분기
            r'(\d{2})년[_\s-]*(\d)분기',      # 25년 3분기, 25년_3분기
            r'(\d{4})[년\s_-]+(\d)분기',     # 2025년_3분기, 2025 3분기
            r'(\d{2})[년\s_-]+(\d)분기',     # 25년_3분기, 25 3분기
            r'(\d{4})[_\s-](\d)[Qq]',        # 2025_3Q, 2025 3q, 2025-3Q
            r'(\d{2})[_\s-](\d)[Qq]',        # 25_3Q, 25 3q
            r'(\d{4})[_\s-](\d)[분]',        # 2025_3, 2025 3분
            r'(\d{2})[_\s-](\d)[분]',        # 25_3, 25 3분
            r'(\d{4})[_\s-](\d)',            # 2025_3, 2025 3
            r'(\d{2})[_\s-](\d)',            # 25_3, 25 3
        ]
        
        for pattern in filename_patterns:
            match = re.search(pattern, filename)
            if match:
                year_str = match.group(1)
                quarter = int(match.group(2))
                
                # 2자리 연도 처리
                if len(year_str) == 2:
                    year = 2000 + int(year_str)
                else:
                    year = int(year_str)
                
                # 분기 유효성 검사
                if 1 <= quarter <= 4:
                    print(f"[연도/분기 추출] ✅ 1순위 성공: 파일명에서 추출 - {year}년 {quarter}분기")
                    return year, quarter
        
        print(f"[연도/분기 추출] 1순위 실패: 파일명에서 패턴을 찾지 못했습니다.")
    except Exception as e:
        print(f"[연도/분기 추출] 1순위 오류: {e}")
    
    # ========================================
    # 2순위: 엑셀 시트 정밀 탐색
    # ========================================
    try:
        xl = pd.ExcelFile(filepath)
        print(f"[연도/분기 추출] 2순위: 엑셀 시트 탐색 시작")
        
        # 2-1. '이용관련' 시트의 K17(연도), M17(분기) 확인
        if '이용관련' in xl.sheet_names:
            print(f"[연도/분기 추출] 2-1: '이용관련' 시트 확인 중...")
            try:
                df_info = pd.read_excel(xl, sheet_name='이용관련', header=None)
                
                # K17 셀 (0-based: 행 16, 열 10)
                if len(df_info) > 16 and len(df_info.columns) > 10:
                    year_cell = df_info.iloc[16, 10]  # K17
                    if pd.notna(year_cell):
                        year_str = str(year_cell).strip()
                        # 연도 추출 (숫자만)
                        year_match = re.search(r'(\d{4})', year_str)
                        if year_match:
                            year = int(year_match.group(1))
        
                            # M17 셀 (0-based: 행 16, 열 12)
                            if len(df_info.columns) > 12:
                                quarter_cell = df_info.iloc[16, 12]  # M17
                                if pd.notna(quarter_cell):
                                    quarter_str = str(quarter_cell).strip()
                                    # 분기 추출 (1-4 숫자)
                                    quarter_match = re.search(r'(\d)', quarter_str)
                                    if quarter_match:
                                        quarter = int(quarter_match.group(1))
                                        if 1 <= quarter <= 4:
                                            print(f"[연도/분기 추출] ✅ 2-1 성공: '이용관련' 시트 K17/M17에서 추출 - {year}년 {quarter}분기")
                                            return year, quarter
                
                print(f"[연도/분기 추출] 2-1 실패: '이용관련' 시트의 K17/M17에서 추출 실패")
            except Exception as e:
                print(f"[연도/분기 추출] 2-1 오류: {e}")
        
        # 2-2. 데이터 시트 헤더(2~3행)에서 '20xx' 패턴 찾기
        print(f"[연도/분기 추출] 2-2: 데이터 시트 헤더 탐색 중...")
        target_sheets = ['A 분석', 'B 분석', 'A(광공업생산)집계', 'B(서비스업생산)집계']
        
        for sheet_name in target_sheets:
            if sheet_name not in xl.sheet_names:
                continue
            
            try:
                # 헤더 행만 읽기 (2-3행, 0-based: 1-2행)
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=3)
                
                if len(df) < 2:
                    continue
        
                # 헤더 행(1-2행, 0-based)에서 연도/분기 패턴 찾기
        for row_idx in [1, 2]:
                    if row_idx >= len(df):
                        continue
                    
                    # 모든 컬럼을 순회하며 패턴 찾기 (최대 30개 컬럼)
                    for col_idx in range(min(30, len(df.columns))):
                        cell_value = df.iloc[row_idx, col_idx]
                if pd.notna(cell_value):
                    cell_str = str(cell_value).strip()
                            
                            # 연도/분기 패턴
                    patterns = [
                                r'(\d{4})\s*\.?\s*(\d)/4',      # 2025 2/4, 2025.2/4
                                r'(\d{2})\s*\.?\s*(\d)/4',       # 25 2/4, 25.2/4
                        r'(\d{4})년[_\s]*(\d)분기',      # 2025년 2분기
                        r'(\d{2})년[_\s]*(\d)분기',      # 25년 2분기
                                r'(\d{4})[_\s-](\d)[Qq]',        # 2025_2Q, 2025 2q
                                r'(\d{2})[_\s-](\d)[Qq]',        # 25_2Q, 25 2q
                                r'(\d{4})[_\s-](\d)',            # 2025_2, 2025 2
                                r'(\d{2})[_\s-](\d)',            # 25_2, 25 2
                    ]
                    
                    for pattern in patterns:
                        match = re.search(pattern, cell_str)
                        if match:
                            year_str = match.group(1)
                            quarter = int(match.group(2))
                            
                            if len(year_str) == 2:
                                year = 2000 + int(year_str)
                            else:
                                year = int(year_str)
                            
                                    # 분기 유효성 검사
                                    if 1 <= quarter <= 4:
                                        print(f"[연도/분기 추출] ✅ 2-2 성공: '{sheet_name}' 시트 헤더에서 추출 - {year}년 {quarter}분기")
                            return year, quarter
            except Exception as e:
                print(f"[연도/분기 추출] 2-2 오류 ({sheet_name}): {e}")
                continue
        
        print(f"[연도/분기 추출] 2순위 실패: 모든 시트 탐색 실패")
    except Exception as e:
        print(f"[연도/분기 추출] 2순위 오류: {e}")
    
    # ========================================
    # 3순위: 안전장치 (Fallback)
    # ========================================
    print(f"[연도/분기 추출] 3순위: 안전장치 실행")
    
    # 기본값이 제공된 경우 사용
        if default_year is not None and default_quarter is not None:
        print(f"[연도/분기 추출] ⚠️ 안전장치: 제공된 기본값 사용 - {default_year}년 {default_quarter}분기")
            return default_year, default_quarter
    
    # 현재 날짜 기준 이전 분기 계산
    try:
        fallback_year, fallback_quarter = get_previous_quarter()
        now = datetime.now()
        current_quarter = (now.month - 1) // 3 + 1
        print(f"[연도/분기 추출] ⚠️ 안전장치: 현재 날짜 기준 이전 분기 사용 - {fallback_year}년 {fallback_quarter}분기 (현재: {now.year}년 {current_quarter}분기)")
        return fallback_year, fallback_quarter
    except Exception as e:
        print(f"[연도/분기 추출] ⚠️ 안전장치 오류: {e}, 최종 기본값 계산 시도")
        # 최종 안전장치: get_previous_quarter 재시도
        try:
            return get_previous_quarter()
        except:
            # 모든 방법 실패 시 현재 연도 1분기 반환 (최소한의 안전장치)
            return datetime.now().year, 1


def extract_year_quarter_from_raw(filepath):
    """기초자료 수집표에서 연도와 분기 추출"""
    import re
    
    try:
        latest_year = 0
        latest_quarter = 0
        
        # 먼저 파일명에서 추출 시도
        filename = Path(filepath).stem
        print(f"[기초자료 연도/분기 추출] 파일명: {filename}")
        
        # 파일명 패턴: "기초자료 수집표_2025년 2분기" 또는 "25년_2분기" 등
        patterns = [
            r'(\d{4})년\s*(\d)분기',  # 2025년 2분기
            r'(\d{2})년\s*(\d)분기',   # 25년 2분기
            r'(\d{4})_(\d)',           # 2025_2
        ]
        
        for pattern in patterns:
            match = re.search(pattern, filename)
            if match:
                year_str = match.group(1)
                quarter = int(match.group(2))
                
                if len(year_str) == 2:
                    year = 2000 + int(year_str)
                else:
                    year = int(year_str)
                
                print(f"[기초자료 연도/분기 추출] 파일명에서 추출: {year}년 {quarter}분기")
                return year, quarter
        
        # 시트에서 연도/분기 정보 추출 시도
        xl = pd.ExcelFile(filepath)
        
        # 기초자료 수집표의 주요 시트에서 헤더 확인
        target_sheets = ['광공업생산', '서비스업생산', '소비(소매, 추가)', '고용률']
        
        for sheet_name in target_sheets:
            if sheet_name not in xl.sheet_names:
                continue
                
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=10)
                
                for row_idx in range(min(10, len(df))):
                    for col_idx in range(min(30, len(df.columns))):
                        cell = str(df.iloc[row_idx, col_idx])
                        if pd.isna(df.iloc[row_idx, col_idx]):
                            continue
                        
                        # "2025 2/4", "2025.2/4", "25.2/4", "2025년 2분기" 등 패턴 찾기
                        patterns = [
                            r'(\d{4})\s*\.?\s*(\d)/4',  # 2025 2/4, 2025.2/4
                            r'(\d{2})\s*\.?\s*(\d)/4',   # 25 2/4, 25.2/4
                            r'(\d{4})년\s*(\d)분기',      # 2025년 2분기
                            r'(\d{2})년\s*(\d)분기',       # 25년 2분기
                        ]
                        
                        for pattern in patterns:
                            match = re.search(pattern, cell)
                            if match:
                                year_str = match.group(1)
                                quarter = int(match.group(2))
                                
                                if len(year_str) == 2:
                                    year = 2000 + int(year_str)
                                else:
                                    year = int(year_str)
                                
                                # 가장 최신 연도/분기 저장
                                if year > latest_year or (year == latest_year and quarter > latest_quarter):
                                    latest_year = year
                                    latest_quarter = quarter
                                    print(f"[기초자료 연도/분기 추출] {sheet_name}에서 발견: {year}년 {quarter}분기")
            except Exception as e:
                print(f"[경고] {sheet_name} 시트 읽기 실패: {e}")
                continue
        
        if latest_year > 0 and latest_quarter > 0:
            print(f"[기초자료 연도/분기 추출] 최종 결과: {latest_year}년 {latest_quarter}분기")
            return latest_year, latest_quarter
        
        # 기본값 (실패 시): 현재 날짜 기준 이전 분기
        try:
            fallback_year, fallback_quarter = get_previous_quarter()
            print(f"[경고] 기초자료 연도/분기 추출 실패, 현재 날짜 기준 이전 분기 사용: {fallback_year}년 {fallback_quarter}분기")
            return fallback_year, fallback_quarter
        except Exception as e2:
            print(f"[경고] 이전 분기 계산 실패: {e2}, 최소 안전장치 사용")
            return datetime.now().year, 1
    except Exception as e:
        print(f"[오류] 기초자료 연도/분기 추출 오류: {e}")
        try:
            return get_previous_quarter()
        except:
            return datetime.now().year, 1


def detect_file_type(filepath: str) -> str:
    """엑셀 파일 유형 자동 감지 (기초자료 수집표 vs 분석표) - 최적화 버전
    
    빠른 판정을 위해 다음 순서로 확인:
    1. 파일명 확인 (가장 빠름)
    2. 시트명만 빠르게 읽어서 핵심 시트 확인
    3. 첫 매칭 시 즉시 반환
    
    Returns:
        'raw': 기초자료 수집표
        'analysis': 분석표
        'unknown': 알 수 없는 형식
    """
    try:
        print(f"[파일 유형 분석] 시작: {Path(filepath).name}")
        
        # ========================================
        # 1단계: 파일명으로 빠른 판정 (가장 빠름)
        # ========================================
        filename = Path(filepath).stem.lower()
        print(f"[파일 유형 분석] 1단계: 파일명 확인 - {filename}")
        if '기초' in filename or '수집' in filename or 'raw' in filename:
            print(f"[파일 유형 분석] 파일명으로 판정: raw")
            return 'raw'
        elif '분석' in filename or 'analysis' in filename:
            print(f"[파일 유형 분석] 파일명으로 판정: analysis")
            return 'analysis'
        
        # ========================================
        # 2단계: 시트명만 빠르게 읽기 (openpyxl 사용 - pandas보다 빠름)
        # ========================================
        print(f"[파일 유형 분석] 2단계: 시트명 읽기 시작...")
        try:
            import openpyxl
            print(f"[파일 유형 분석] openpyxl로 파일 열기 시도...")
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=False)
            sheet_names = set(wb.sheetnames)
            print(f"[파일 유형 분석] 시트명 읽기 완료: {len(sheet_names)}개 시트")
            wb.close()
        except Exception as e:
            print(f"[파일 유형 분석] openpyxl 실패, pandas 사용: {e}")
            # openpyxl 실패 시 pandas 사용 (fallback)
            xl = pd.ExcelFile(filepath)
            sheet_names = set(xl.sheet_names)
            print(f"[파일 유형 분석] pandas로 시트명 읽기 완료: {len(sheet_names)}개 시트")
        
        sheet_count = len(sheet_names)
        
        # ========================================
        # 3단계: 핵심 시트만 확인 (빠른 판정)
        # ========================================
        # 분석표의 가장 특징적인 시트들 (우선순위 높음)
        ANALYSIS_KEY_SHEETS = {
            '본청', '이용관련',  # 분석표에만 있는 확실한 시트
            'A(광공업생산)집계', 'A 분석',  # 가장 흔한 분석표 시트
            'B(서비스업생산)집계', 'B 분석',
        }
        
        # 기초자료의 가장 특징적인 시트들 (우선순위 높음)
        RAW_KEY_SHEETS = {
            '광공업생산', '서비스업생산',  # 가장 흔한 기초자료 시트
            '완료체크',  # 기초자료에만 있는 시트
            '고용률', '실업자 수',
        }
        
        # 핵심 시트 매칭 확인 (첫 매칭 시 즉시 반환)
        print(f"[파일 유형 분석] 3단계: 핵심 시트 매칭 확인...")
        for sheet in sheet_names:
            if sheet in ANALYSIS_KEY_SHEETS:
                print(f"[파일 유형 분석] 분석표 핵심 시트 발견: {sheet} → analysis")
                return 'analysis'
            if sheet in RAW_KEY_SHEETS:
                print(f"[파일 유형 분석] 기초자료 핵심 시트 발견: {sheet} → raw")
                return 'raw'
        
        # ========================================
        # 4단계: 패턴 매칭 (빠른 확인)
        # ========================================
        print(f"[파일 유형 분석] 4단계: 패턴 매칭 확인...")
        # 분석표 패턴: '집계', '분석', '참고' 키워드
        analysis_pattern_count = 0
        raw_pattern_count = 0
        
        for sheet in sheet_names:
            if '집계' in sheet or '분석' in sheet or '참고' in sheet:
                analysis_pattern_count += 1
                if analysis_pattern_count >= 2:  # 2개만 찾으면 충분
                    print(f"[파일 유형 분석] 분석표 패턴 매칭 ({analysis_pattern_count}개) → analysis")
                    return 'analysis'
            # 기초자료는 한글 시트명이 많고 특별한 패턴이 없으므로 패턴 매칭 생략
        
        # ========================================
        # 5단계: 시트 개수로 추정 (빠른 판정)
        # ========================================
        print(f"[파일 유형 분석] 5단계: 시트 개수로 추정 (시트 수: {sheet_count})...")
        if sheet_count <= 20:
            # 기초자료일 가능성 높음 (~17개)
            print(f"[파일 유형 분석] 시트 개수로 판정: raw")
            return 'raw'
        elif sheet_count >= 30:
            # 분석표일 가능성 높음 (~42개)
            print(f"[파일 유형 분석] 시트 개수로 판정: analysis")
            return 'analysis'
        
        # 기본값: 분석표 (더 복잡한 처리가 필요하므로)
        print(f"[파일 유형 분석] 기본값으로 판정: analysis")
        return 'analysis'
            
    except Exception as e:
        import traceback
        print(f"[오류] 파일 유형 감지 실패: {e}")
        traceback.print_exc()
        return 'unknown'


def get_file_type_details(filepath: str) -> dict:
    """파일 유형과 상세 정보를 함께 반환
    
    Returns:
        {
            'type': 'raw' | 'analysis' | 'unknown',
            'sheet_count': int,
            'matched_raw_sheets': list,
            'matched_analysis_sheets': list,
            'confidence': 'high' | 'medium' | 'low',
            'reason': str
        }
    """
    try:
        xl = pd.ExcelFile(filepath)
        sheet_names = set(xl.sheet_names)
        sheet_count = len(sheet_names)
        
        # 기초자료 전용 시트
        RAW_ONLY_SHEETS = {
            '광공업생산', '서비스업생산', '소비(소매, 추가)', 
            '고용', '고용(kosis)', '고용률', '실업자 수',
            '지출목적별 물가', '품목성질별 물가', '건설 (공표자료)',
            '수출', '수입', '분기 GRDP', '완료체크'
        }
        
        # 분석표 전용 시트
        ANALYSIS_ONLY_SHEETS = {
            '본청', '시도별 현황', '이용관련',
            'A(광공업생산)집계', 'A 분석',
            'B(서비스업생산)집계', 'B 분석',
            "F'(건설)집계", "F'분석",
            'G(수출)집계', 'G 분석',
        }
        
        matched_raw = sorted(sheet_names & RAW_ONLY_SHEETS)
        matched_analysis = sorted(sheet_names & ANALYSIS_ONLY_SHEETS)
        
        result = {
            'sheet_count': sheet_count,
            'sheet_names': sorted(sheet_names),
            'matched_raw_sheets': matched_raw,
            'matched_analysis_sheets': matched_analysis,
        }
        
        # 판정
        if len(matched_analysis) >= 5:
            result['type'] = 'analysis'
            result['confidence'] = 'high'
            result['reason'] = f"분석표 전용 시트 {len(matched_analysis)}개 발견"
        elif len(matched_raw) >= 5:
            result['type'] = 'raw'
            result['confidence'] = 'high'
            result['reason'] = f"기초자료 전용 시트 {len(matched_raw)}개 발견"
        elif len(matched_analysis) >= 2:
            result['type'] = 'analysis'
            result['confidence'] = 'medium'
            result['reason'] = f"분석표 전용 시트 {len(matched_analysis)}개 발견"
        elif len(matched_raw) >= 2:
            result['type'] = 'raw'
            result['confidence'] = 'medium'
            result['reason'] = f"기초자료 전용 시트 {len(matched_raw)}개 발견"
        else:
            result['type'] = 'unknown'
            result['confidence'] = 'low'
            result['reason'] = "시트 구조를 명확히 판단할 수 없음"
        
        return result
        
    except Exception as e:
        return {
            'type': 'unknown',
            'confidence': 'low',
            'reason': f"파일 읽기 오류: {e}",
            'sheet_count': 0,
            'matched_raw_sheets': [],
            'matched_analysis_sheets': []
        }


def validate_sheet_structure(filepath: str, expected_type: str) -> dict:
    """파일의 시트 구조가 예상 유형과 일치하는지 검증
    
    Args:
        filepath: 엑셀 파일 경로
        expected_type: 'raw' 또는 'analysis'
        
    Returns:
        {
            'valid': bool,
            'missing_sheets': list,  # 필수 시트 중 누락된 것
            'extra_sheets': list,    # 예상치 못한 추가 시트
            'warnings': list
        }
    """
    try:
        xl = pd.ExcelFile(filepath)
        sheet_names = set(xl.sheet_names)
        
        result = {
            'valid': True,
            'missing_sheets': [],
            'extra_sheets': [],
            'warnings': []
        }
        
        if expected_type == 'raw':
            # 기초자료 필수 시트
            REQUIRED_RAW_SHEETS = {
                '광공업생산', '서비스업생산', '고용률', '실업자 수',
                '수출', '수입', '시도 간 이동'
            }
            OPTIONAL_RAW_SHEETS = {
                '소비(소매, 추가)', '지출목적별 물가', '품목성질별 물가',
                '건설 (공표자료)', '분기 GRDP', '완료체크',
                '고용', '고용(kosis)', '연령별 인구이동', '시군구인구이동'
            }
            
            missing = REQUIRED_RAW_SHEETS - sheet_names
            if missing:
                result['valid'] = False
                result['missing_sheets'] = sorted(missing)
            
            # 예상치 못한 시트 (분석표 시트가 섞여있는 경우)
            unexpected = [s for s in sheet_names if '분석' in s or '집계' in s]
            if unexpected:
                result['warnings'].append(f"기초자료에 분석표 시트가 포함됨: {unexpected}")
                
        elif expected_type == 'analysis':
            # 분석표 필수 시트
            REQUIRED_ANALYSIS_SHEETS = {
                'A(광공업생산)집계', 'A 분석',
                'B(서비스업생산)집계', 'B 분석',
                'C(소비)집계', 'C 분석',
                'D(고용률)집계', 'D(고용률)분석',
                '이용관련'
            }
            
            missing = REQUIRED_ANALYSIS_SHEETS - sheet_names
            if missing:
                result['valid'] = False
                result['missing_sheets'] = sorted(missing)
            
            # GRDP 시트 체크 (선택이지만 경고)
            grdp_sheets = {'GRDP', 'I GRDP', '분기 GRDP'}
            if not (sheet_names & grdp_sheets):
                result['warnings'].append("GRDP 시트가 없음 - 별도 업로드 필요")
        
        return result
        
    except Exception as e:
        return {
            'valid': False,
            'missing_sheets': [],
            'extra_sheets': [],
            'warnings': [f"파일 읽기 오류: {e}"]
        }

