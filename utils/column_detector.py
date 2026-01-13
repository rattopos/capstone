#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
동적 컬럼 감지 유틸리티
기초자료 형식과 분석표 형식을 자동으로 감지하고 올바른 컬럼 인덱스를 반환합니다.
"""

import pandas as pd
from typing import Dict, Optional, Tuple, List


def detect_column_structure(df: pd.DataFrame, header_row_idx: Optional[int] = None) -> Dict[str, int]:
    """
    데이터프레임의 컬럼 구조를 자동으로 감지합니다.
    
    Args:
        df: 데이터프레임
        header_row_idx: 헤더 행 인덱스 (None이면 자동 감지)
    
    Returns:
        컬럼 인덱스 딕셔너리:
        {
            'is_raw_format': bool,
            'region_col': int,
            'class_col': int,
            'weight_col': int,
            'code_col': int,
            'name_col': int,
            'header_row_idx': int
        }
    """
    # 헤더 행 찾기
    if header_row_idx is None:
        header_row_idx = 2  # 기본값
        for i in range(min(10, len(df))):
            row = df.iloc[i]
            row_str = ' '.join([str(v) for v in row.values[:10] if pd.notna(v)])
            if '지역' in row_str:
                header_row_idx = i
                break
    
    # 기초자료 형식 감지 (헤더 행에서 "지역"이 열 1에 있음)
    is_raw_format = False
    if len(df.columns) > 1 and header_row_idx < len(df):
        header_cell = df.iloc[header_row_idx, 1]
        if pd.notna(header_cell) and '지역' in str(header_cell):
            is_raw_format = True
    
    if is_raw_format:
        # 기초자료 형식
        return {
            'is_raw_format': True,
            'region_col': 1,
            'class_col': 2,
            'weight_col': 3,
            'code_col': 4,
            'name_col': 5,
            'header_row_idx': header_row_idx
        }
    else:
        # 분석표 형식
        return {
            'is_raw_format': False,
            'region_col': 4,
            'class_col': 5,
            'weight_col': 6,
            'code_col': 7,
            'name_col': 8,
            'header_row_idx': header_row_idx
        }


def find_quarter_columns(df: pd.DataFrame, header_row_idx: int, 
                        year: Optional[int] = None, quarter: Optional[int] = None) -> Dict[str, Optional[int]]:
    """
    분기 데이터 컬럼을 헤더에서 동적으로 찾습니다.
    
    Args:
        df: 데이터프레임
        header_row_idx: 헤더 행 인덱스
        year: 현재 연도 (선택사항)
        quarter: 현재 분기 (선택사항)
    
    Returns:
        분기 컬럼 인덱스 딕셔너리:
        {
            'col_2024_2q': Optional[int],
            'col_2025_2q': Optional[int],
            'col_2025_3q': Optional[int],
            'current_quarter_col': int
        }
    """
    if header_row_idx >= len(df):
        header_row_idx = 0
    
    header_row = df.iloc[header_row_idx]
    
    col_2024_2q = None
    col_2025_2q = None
    col_2025_3q = None
    
    # 헤더에서 분기 컬럼 찾기
    for col_idx in range(len(header_row)):
        val = str(header_row[col_idx]) if pd.notna(header_row[col_idx]) else ''
        val_clean = val.strip().replace('.', ' ').replace('p', '').replace('P', '')
        
        if '2024' in val_clean and '2/4' in val_clean:
            col_2024_2q = col_idx
        if '2025' in val_clean and '2/4' in val_clean:
            col_2025_2q = col_idx
        if '2025' in val_clean and '3/4' in val_clean:
            col_2025_3q = col_idx
    
    # 현재 분기 컬럼 선택
    if quarter == 3 and col_2025_3q is not None:
        current_quarter_col = col_2025_3q
    elif col_2025_2q is not None:
        current_quarter_col = col_2025_2q
    else:
        # 기본값: 마지막 분기 컬럼
        current_quarter_col = len(header_row) - 1
    
    # 2024 2/4를 찾지 못한 경우, 현재 분기의 전년동분기 찾기
    if col_2024_2q is None:
        target_quarter = quarter if quarter else 2
        for col_idx in range(len(header_row)):
            val = str(header_row[col_idx]) if pd.notna(header_row[col_idx]) else ''
            val_clean = val.strip().replace('.', ' ').replace('p', '').replace('P', '')
            if '2024' in val_clean and f'{target_quarter}/4' in val_clean:
                col_2024_2q = col_idx
                break
    
    return {
        'col_2024_2q': col_2024_2q,
        'col_2025_2q': col_2025_2q,
        'col_2025_3q': col_2025_3q,
        'current_quarter_col': current_quarter_col
    }


def get_column_mapping(df: pd.DataFrame, year: Optional[int] = None, 
                       quarter: Optional[int] = None) -> Dict[str, any]:
    """
    데이터프레임의 전체 컬럼 매핑을 반환합니다.
    
    Args:
        df: 데이터프레임
        year: 현재 연도 (선택사항)
        quarter: 현재 분기 (선택사항)
    
    Returns:
        전체 컬럼 매핑 딕셔너리
    """
    column_structure = detect_column_structure(df)
    quarter_columns = find_quarter_columns(
        df, 
        column_structure['header_row_idx'],
        year,
        quarter
    )
    
    return {
        **column_structure,
        **quarter_columns
    }
