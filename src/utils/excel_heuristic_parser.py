#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Heuristic Parser
방어적이고 휴리스틱한 Excel 파싱 유틸리티
하드코딩된 시트 이름이나 행 인덱스 없이 동적으로 데이터를 찾습니다.
"""

import pandas as pd
from typing import List, Dict, Optional, Tuple, Set
from pathlib import Path
import re


class ExcelHeuristicParser:
    """방어적 Excel 파서 - 휴리스틱 기반 시트 및 테이블 탐색"""
    
    def __init__(self, excel_path: str):
        """
        초기화
        
        Args:
            excel_path: Excel 파일 경로
        """
        self.excel_path = Path(excel_path)
        self._xl: Optional[pd.ExcelFile] = None
        self._sheet_cache: Dict[str, pd.DataFrame] = {}
    
    @property
    def xl(self) -> pd.ExcelFile:
        """ExcelFile 객체 (lazy loading)"""
        if self._xl is None:
            self._xl = pd.ExcelFile(self.excel_path)
        return self._xl
    
    def find_target_sheet(
        self, 
        keywords: List[str],
        required_columns: Optional[List[str]] = None,
        required_row_labels: Optional[List[str]] = None,
        max_sheets_to_check: int = 10
    ) -> Optional[Tuple[str, pd.DataFrame]]:
        """
        키워드 기반으로 타겟 시트를 찾습니다.
        
        Args:
            keywords: 시트 이름에 포함되어야 할 키워드 리스트 (우선순위 순)
            required_columns: 시트에 있어야 할 컬럼 이름들 (선택사항)
            required_row_labels: 시트에 있어야 할 행 레이블들 (선택사항)
            max_sheets_to_check: 확인할 최대 시트 수
        
        Returns:
            (시트 이름, DataFrame) 튜플 또는 None
        """
        sheet_names = self.xl.sheet_names
        
        # 1단계: 키워드 기반 점수 계산
        scored_sheets = []
        for sheet_name in sheet_names:
            score = self._calculate_sheet_score(sheet_name, keywords)
            if score > 0:
                scored_sheets.append((score, sheet_name))
        
        # 점수 순으로 정렬
        scored_sheets.sort(key=lambda x: x[0], reverse=True)
        
        # 2단계: 상위 후보들 검증
        for score, sheet_name in scored_sheets[:max_sheets_to_check]:
            try:
                df = self._load_sheet(sheet_name)
                
                # 컬럼 검증
                if required_columns:
                    if not self._validate_columns(df, required_columns):
                        continue
                
                # 행 레이블 검증
                if required_row_labels:
                    if not self._validate_row_labels(df, required_row_labels):
                        continue
                
                # 모든 검증 통과
                return (sheet_name, df)
            
            except Exception as e:
                # 시트 로드 실패 시 다음 후보로
                continue
        
        return None
    
    def _calculate_sheet_score(self, sheet_name: str, keywords: List[str]) -> float:
        """
        시트 이름의 키워드 매칭 점수를 계산합니다.
        
        Args:
            sheet_name: 시트 이름
            keywords: 키워드 리스트 (우선순위 순)
        
        Returns:
            점수 (0.0 ~ 1.0)
        """
        sheet_lower = sheet_name.lower()
        total_score = 0.0
        
        for idx, keyword in enumerate(keywords):
            keyword_lower = keyword.lower()
            
            # 정확한 매칭 (가장 높은 점수)
            if keyword_lower == sheet_lower:
                total_score += 10.0 * (len(keywords) - idx)
            # 부분 매칭
            elif keyword_lower in sheet_lower:
                # 키워드가 시트 이름의 시작 부분에 있으면 더 높은 점수
                if sheet_lower.startswith(keyword_lower):
                    total_score += 5.0 * (len(keywords) - idx)
                else:
                    total_score += 2.0 * (len(keywords) - idx)
            # 단어 단위 매칭 (공백, 언더스코어, 하이픈으로 구분)
            else:
                # 키워드를 단어로 분리하여 매칭
                keyword_words = re.split(r'[_\-\s]+', keyword_lower)
                sheet_words = re.split(r'[_\-\s]+', sheet_lower)
                
                matched_words = sum(1 for kw in keyword_words if kw in sheet_words)
                if matched_words > 0:
                    total_score += (matched_words / len(keyword_words)) * 1.0 * (len(keywords) - idx)
        
        return total_score
    
    def _load_sheet(self, sheet_name: str) -> pd.DataFrame:
        """시트를 로드합니다 (캐싱)"""
        if sheet_name not in self._sheet_cache:
            self._sheet_cache[sheet_name] = pd.read_excel(
                self.xl, 
                sheet_name=sheet_name, 
                header=None
            )
        return self._sheet_cache[sheet_name].copy()
    
    def _validate_columns(
        self, 
        df: pd.DataFrame, 
        required_columns: List[str],
        max_rows_to_check: int = 10
    ) -> bool:
        """
        DataFrame에 필요한 컬럼이 있는지 검증합니다.
        
        Args:
            df: 검증할 DataFrame
            required_columns: 필요한 컬럼 이름 리스트
            max_rows_to_check: 헤더를 찾기 위해 확인할 최대 행 수
        
        Returns:
            검증 통과 여부
        """
        if len(df) == 0:
            return False
        
        # 처음 몇 행에서 컬럼 이름 찾기
        found_columns = set()
        for row_idx in range(min(max_rows_to_check, len(df))):
            row = df.iloc[row_idx]
            row_str = ' '.join([str(v).lower() for v in row.values[:20] if pd.notna(v)])
            
            for col_name in required_columns:
                col_lower = col_name.lower()
                if col_lower in row_str:
                    found_columns.add(col_name)
        
        # 모든 필수 컬럼이 발견되었는지 확인
        return len(found_columns) >= len(required_columns) * 0.7  # 70% 이상 매칭
    
    def _validate_row_labels(
        self, 
        df: pd.DataFrame, 
        required_labels: List[str],
        max_rows_to_check: int = 100
    ) -> bool:
        """
        DataFrame에 필요한 행 레이블이 있는지 검증합니다.
        
        Args:
            df: 검증할 DataFrame
            required_labels: 필요한 행 레이블 리스트
            max_rows_to_check: 확인할 최대 행 수
        
        Returns:
            검증 통과 여부
        """
        if len(df) == 0:
            return False
        
        found_labels = set()
        for row_idx in range(min(max_rows_to_check, len(df))):
            row = df.iloc[row_idx]
            row_str = ' '.join([str(v) for v in row.values[:10] if pd.notna(v)])
            
            for label in required_labels:
                if label in row_str:
                    found_labels.add(label)
        
        # 모든 필수 레이블이 발견되었는지 확인
        return len(found_labels) >= len(required_labels) * 0.7  # 70% 이상 매칭
    
    def locate_table_start(
        self, 
        df: pd.DataFrame,
        anchor_keywords: List[str],
        max_rows_to_check: int = 10
    ) -> Optional[int]:
        """
        테이블 시작 행을 동적으로 찾습니다.
        
        Args:
            df: 검색할 DataFrame
            anchor_keywords: 헤더 행에 있어야 할 키워드 리스트
            max_rows_to_check: 확인할 최대 행 수
        
        Returns:
            헤더 행 인덱스 또는 None
        """
        if len(df) == 0:
            return None
        
        best_row = None
        best_score = 0
        
        for row_idx in range(min(max_rows_to_check, len(df))):
            row = df.iloc[row_idx]
            
            # 행의 모든 값에서 키워드 검색
            row_values = [str(v).lower() if pd.notna(v) else '' for v in row.values[:30]]
            row_text = ' '.join(row_values)
            
            # 키워드 매칭 점수 계산
            score = 0
            for keyword in anchor_keywords:
                keyword_lower = keyword.lower()
                if keyword_lower in row_text:
                    # 정확한 매칭 (컬럼 이름으로 사용되는 경우)
                    if any(keyword_lower == val.strip() for val in row_values):
                        score += 10
                    # 부분 매칭
                    else:
                        score += 5
            
            if score > best_score:
                best_score = score
                best_row = row_idx
        
        # 최소 점수 이상이면 반환
        if best_score >= len(anchor_keywords) * 3:
            return best_row
        
        return None
    
    def find_multiple_sheets(
        self,
        keywords: List[str],
        required_columns: Optional[List[str]] = None,
        required_row_labels: Optional[List[str]] = None
    ) -> List[Tuple[str, pd.DataFrame]]:
        """
        키워드와 매칭되는 여러 시트를 찾습니다 (분할된 데이터셋용).
        
        Args:
            keywords: 시트 이름에 포함되어야 할 키워드 리스트
            required_columns: 시트에 있어야 할 컬럼 이름들 (선택사항)
            required_row_labels: 시트에 있어야 할 행 레이블들 (선택사항)
        
        Returns:
            (시트 이름, DataFrame) 튜플 리스트
        """
        sheet_names = self.xl.sheet_names
        matched_sheets = []
        
        for sheet_name in sheet_names:
            score = self._calculate_sheet_score(sheet_name, keywords)
            if score > 0:
                try:
                    df = self._load_sheet(sheet_name)
                    
                    # 검증
                    if required_columns:
                        if not self._validate_columns(df, required_columns):
                            continue
                    
                    if required_row_labels:
                        if not self._validate_row_labels(df, required_row_labels):
                            continue
                    
                    matched_sheets.append((sheet_name, df, score))
                
                except Exception:
                    continue
        
        # 점수 순으로 정렬
        matched_sheets.sort(key=lambda x: x[2], reverse=True)
        
        return [(name, df) for name, df, _ in matched_sheets]
    
    def get_sheet_by_fallback(
        self,
        primary_keywords: List[str],
        fallback_keywords: List[str],
        required_columns: Optional[List[str]] = None,
        required_row_labels: Optional[List[str]] = None
    ) -> Optional[Tuple[str, pd.DataFrame, bool]]:
        """
        기본 키워드로 시트를 찾고, 실패하면 대체 키워드로 찾습니다.
        
        Args:
            primary_keywords: 기본 키워드 리스트
            fallback_keywords: 대체 키워드 리스트
            required_columns: 필요한 컬럼 이름들 (선택사항)
            required_row_labels: 필요한 행 레이블들 (선택사항)
        
        Returns:
            (시트 이름, DataFrame, is_fallback) 튜플 또는 None
        """
        # 기본 키워드로 시도
        result = self.find_target_sheet(
            primary_keywords,
            required_columns,
            required_row_labels
        )
        
        if result:
            return (result[0], result[1], False)
        
        # 대체 키워드로 시도
        result = self.find_target_sheet(
            fallback_keywords,
            required_columns,
            required_row_labels
        )
        
        if result:
            return (result[0], result[1], True)
        
        return None
    
    def close(self):
        """리소스 정리"""
        self._sheet_cache.clear()
        if self._xl is not None:
            self._xl.close()
            self._xl = None
