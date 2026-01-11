#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
동적 컬럼 찾기 유틸리티

PM 요구사항: 하드코딩된 열 번호를 제거하고 헤더 텍스트를 검색하여 열 번호를 찾도록 수정
"""

import pandas as pd
from typing import Optional


def find_column_by_header_text(df: pd.DataFrame, header_row: int, year: int, quarter: int) -> Optional[int]:
    """헤더 텍스트를 검색하여 열 번호를 찾기 (PM 가이드 준수)
    
    Args:
        df: DataFrame
        header_row: 헤더가 있는 행 인덱스
        year: 대상 연도
        quarter: 대상 분기 (1-4)
        
    Returns:
        컬럼 인덱스 또는 None
        
    Example:
        # [수정 전]
        # current_data = row[64]  <-- 위험!
        
        # [수정 후]
        header_row = 2
        target_col_idx = find_column_by_header_text(df, header_row, 2025, 3)
        if target_col_idx is not None:
            current_data = row[target_col_idx]
    """
    if header_row >= len(df):
        return None
    
    target_col_idx = -1
    header_row_data = df.iloc[header_row]
    
    # PM 가이드: "2025"와 "3/4" (또는 "3분기")가 모두 포함된 열 찾기
    for idx, val in enumerate(header_row_data):
        if pd.isna(val):
            continue
        
        val_str = str(val).strip()
        
        has_year = str(year) in val_str
        has_quarter = (f"{quarter}/4" in val_str or f"{quarter}분기" in val_str)
        
        if has_year and has_quarter:
            target_col_idx = idx
            break
    
    return target_col_idx if target_col_idx != -1 else None


def find_columns_by_header_text(df: pd.DataFrame, header_row: int, 
                                year: int, quarter: int, 
                                prev_year: Optional[int] = None) -> tuple[Optional[int], Optional[int]]:
    """현재 분기와 전년동기 분기 컬럼을 동시에 찾기
    
    Args:
        df: DataFrame
        header_row: 헤더가 있는 행 인덱스
        year: 현재 연도
        quarter: 현재 분기 (1-4)
        prev_year: 전년 연도 (None이면 year - 1)
        
    Returns:
        (현재 분기 컬럼 인덱스, 전년동기 분기 컬럼 인덱스) 튜플
    """
    if prev_year is None:
        prev_year = year - 1
    
    curr_col = find_column_by_header_text(df, header_row, year, quarter)
    prev_col = find_column_by_header_text(df, header_row, prev_year, quarter)
    
    return curr_col, prev_col
