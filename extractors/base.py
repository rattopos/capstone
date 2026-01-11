#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
기본 데이터 추출기 클래스
모든 보도자료 추출기의 기반이 되는 공통 로직 제공
"""

import pandas as pd
import re
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple

from .config import ALL_REGIONS, RAW_SHEET_MAPPING, RAW_SHEET_QUARTER_COLS, SHEET_KEYWORDS


class BaseExtractor:
    """기본 데이터 추출기 클래스"""
    
    # 지역명 전체 이름 → 단축명 매핑
    REGION_FULL_TO_SHORT: Dict[str, str] = {
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
        self._xl: Optional[pd.ExcelFile] = None
        self._sheet_cache: Dict[str, pd.DataFrame] = {}
        self._header_cache: Dict[str, Dict] = {}
        
        if not self.raw_excel_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {raw_excel_path}")
        
        try:
            self._file_mtime = self.raw_excel_path.stat().st_mtime
        except OSError:
            self._file_mtime = None
    
    # =========================================================================
    # 캐시 관리
    # =========================================================================
    
    def _check_file_modified(self) -> bool:
        """파일이 수정되었는지 확인"""
        if self._file_mtime is None:
            return False
        try:
            if not self.raw_excel_path.exists():
                return True
            current_mtime = self.raw_excel_path.stat().st_mtime
            return abs(current_mtime - self._file_mtime) > 1.0
        except OSError:
            return True
    
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
        """ExcelFile 객체 가져오기 (캐싱)"""
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
    
    def find_sheet_by_keywords(self, keywords: List[str], exact_match: Optional[str] = None) -> Optional[str]:
        """키워드 기반 fuzzy matching으로 시트 찾기
        
        Args:
            keywords: 찾을 시트의 키워드 리스트 (예: ['고용률', '연령별'])
            exact_match: 정확한 시트명 (우선순위 1)
            
        Returns:
            찾은 시트명 또는 None
        """
        xl = self._get_excel_file()
        sheet_names = xl.sheet_names
        
        # 1. 정확한 매칭 (최우선)
        if exact_match and exact_match in sheet_names:
            return exact_match
        
        # 2. 키워드 기반 부분 매칭
        # 모든 키워드가 포함된 시트 찾기
        best_match = None
        best_score = 0
        
        for sheet_name in sheet_names:
            normalized_sheet = self._normalize_sheet_name(sheet_name)
            score = 0
            
            # 각 키워드가 포함되어 있는지 확인
            matched_keywords = 0
            for keyword in keywords:
                normalized_keyword = self._normalize_sheet_name(keyword)
                if normalized_keyword in normalized_sheet:
                    matched_keywords += 1
                    # 키워드 길이에 비례하여 점수 부여
                    score += len(normalized_keyword) / len(normalized_sheet)
            
            # 모든 키워드가 매칭되고 점수가 높으면 선택
            if matched_keywords == len(keywords) and score > best_score:
                best_score = score
                best_match = sheet_name
        
        if best_match:
            return best_match
        
        # 3. 일부 키워드만 매칭되는 경우 (fallback)
        for sheet_name in sheet_names:
            normalized_sheet = self._normalize_sheet_name(sheet_name)
            for keyword in keywords:
                normalized_keyword = self._normalize_sheet_name(keyword)
                if normalized_keyword in normalized_sheet:
                    return sheet_name
        
        return None
    
    def _normalize_sheet_name(self, name: str) -> str:
        """시트명 정규화 (공백, 괄호, 특수문자 제거)"""
        if not name:
            return ""
        # 소문자 변환, 공백 제거, 괄호 제거
        normalized = re.sub(r'[()（）\s\-_\.]', '', str(name).lower())
        return normalized
    
    def _load_sheet(self, sheet_name: str, keywords: Optional[List[str]] = None) -> Optional[pd.DataFrame]:
        """시트 로드 (캐싱 및 fuzzy matching 지원)
        
        Args:
            sheet_name: 시트명 (정확한 이름 또는 기본값)
            keywords: 키워드 리스트 (시트를 찾지 못할 경우 사용, None이면 자동 추출)
        """
        if self._check_file_modified():
            self._clear_cache()
        
        # 캐시 확인
        if sheet_name in self._sheet_cache:
            return self._sheet_cache[sheet_name]
        
        xl = self._get_excel_file()
        actual_sheet_name = sheet_name
        
        # 시트를 찾지 못한 경우 fuzzy matching 시도
        if sheet_name not in xl.sheet_names:
            # 키워드가 제공되지 않으면 config에서 찾기
            if keywords is None:
                keywords = SHEET_KEYWORDS.get(sheet_name, [])
                # 키워드가 없으면 시트명에서 추출
                if not keywords:
                    # 괄호 제거 후 단어 추출
                    cleaned = re.sub(r'[()（）\s]', ' ', sheet_name)
                    keywords = [k.strip() for k in cleaned.split() if k.strip() and len(k.strip()) > 1]
                    if not keywords:
                        keywords = [sheet_name]
            
            actual_sheet_name = self.find_sheet_by_keywords(keywords, exact_match=sheet_name)
            if not actual_sheet_name:
                print(f"[BaseExtractor] 시트를 찾을 수 없습니다: {sheet_name} (키워드: {keywords})")
                return None
            if actual_sheet_name != sheet_name:
                print(f"[BaseExtractor] 시트명 매칭: '{sheet_name}' → '{actual_sheet_name}' (키워드: {keywords})")
        
        # 시트 로드
        try:
            df = pd.read_excel(xl, sheet_name=actual_sheet_name, header=None)
            self._sheet_cache[sheet_name] = df  # 원래 이름으로 캐싱
            self._sheet_cache[actual_sheet_name] = df  # 실제 이름으로도 캐싱
            return df
        except Exception as e:
            print(f"[BaseExtractor] 시트 로드 실패: {actual_sheet_name} - {e}")
            return None
    
    # =========================================================================
    # 시트 구조 분석
    # =========================================================================
    
    def parse_sheet_structure(self, sheet_name: str, header_row: int = 2) -> Dict:
        """시트 헤더에서 연도/분기별 컬럼 인덱스 매핑 생성"""
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
            
            # 연도 패턴
            if isinstance(val, (int, float)) and 2000 <= int(val) <= 2100:
                year_cols[int(val)] = col_idx
            elif re.match(r'^(\d{4})\.?0?$', val_str):
                year = int(re.match(r'^(\d{4})\.?0?$', val_str).group(1))
                year_cols[year] = col_idx
            
            # 분기 패턴
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
        """특정 연도/분기의 컬럼 인덱스 반환"""
        structure = self.parse_sheet_structure(sheet_name)
        quarter_key = f"{year} {quarter}/4"
        return structure['quarters'].get(quarter_key)
    
    def get_year_column_index(self, sheet_name: str, year: int) -> Optional[int]:
        """특정 연도의 컬럼 인덱스 반환"""
        structure = self.parse_sheet_structure(sheet_name)
        return structure['years'].get(year)
    
    def get_raw_sheet_name(self, report_name: str) -> Optional[str]:
        """보도자료 이름에서 기초자료 시트 이름 반환"""
        return RAW_SHEET_MAPPING.get(report_name)
    
    def get_sheet_config(self, sheet_name: str) -> Dict:
        """시트별 설정 정보 반환"""
        return RAW_SHEET_QUARTER_COLS.get(sheet_name, {})
    
    def find_and_load_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """시트 찾기 및 로드 (fuzzy matching 포함)
        
        Args:
            sheet_name: 시트명 또는 기본 시트명
            
        Returns:
            DataFrame 또는 None
        """
        xl = self._get_excel_file()
        
        # 1. 정확한 매칭 시도
        if sheet_name in xl.sheet_names:
            return self._load_sheet(sheet_name)
        
        # 2. 키워드 기반 fuzzy matching
        keywords = SHEET_KEYWORDS.get(sheet_name, [sheet_name])
        actual_sheet = self.find_sheet_by_keywords(keywords, exact_match=sheet_name)
        
        if actual_sheet:
            return self._load_sheet(actual_sheet)
        
        # 3. 시트명에서 키워드 추출하여 재시도
        # 예: "연령별고용률" → ["연령별", "고용률"]
        if '(' in sheet_name or ')' in sheet_name:
            # 괄호 제거 후 키워드 추출
            cleaned = re.sub(r'[()（）]', ' ', sheet_name)
            keywords = [k.strip() for k in cleaned.split() if k.strip()]
            actual_sheet = self.find_sheet_by_keywords(keywords)
            if actual_sheet:
                return self._load_sheet(actual_sheet)
        
        print(f"[BaseExtractor] 시트를 찾을 수 없습니다: {sheet_name}")
        return None
    
    # =========================================================================
    # 지역 처리
    # =========================================================================
    
    def normalize_region(self, region: str) -> Optional[str]:
        """지역명 정규화"""
        if not region:
            return None
        region = str(region).strip()
        return self.REGION_FULL_TO_SHORT.get(region, region if region in ALL_REGIONS else None)
    
    def find_region_row(
        self,
        df: pd.DataFrame,
        region: str,
        region_col: int,
        level_col: int,
        level_value: str = '0',
        start_row: int = 3
    ) -> Optional[int]:
        """특정 지역의 총지수 행 찾기"""
        normalized_region = self.normalize_region(region)
        if not normalized_region:
            return None
        
        for row_idx in range(start_row, len(df)):
            try:
                row_region = str(df.iloc[row_idx, region_col]).strip()
                norm_row_region = self.normalize_region(row_region)
                
                if norm_row_region == normalized_region:
                    level_val = df.iloc[row_idx, level_col]
                    if pd.isna(level_val):
                        continue
                    if str(level_val).strip() == level_value:
                        return row_idx
            except (IndexError, ValueError):
                continue
        return None
    
    # =========================================================================
    # 데이터 추출 유틸리티
    # =========================================================================
    
    def safe_float(self, value: Any) -> Optional[float]:
        """안전한 float 변환"""
        if value is None or pd.isna(value):
            return None
        try:
            return float(value)
        except (ValueError, TypeError):
            return None
    
    def calculate_growth_rate(self, current: Optional[float], previous: Optional[float]) -> Optional[float]:
        """전년동기비 증감률 계산"""
        if current is None or previous is None:
            return None
        if previous == 0:
            return None
        return round((current - previous) / previous * 100, 1)
    
    def calculate_difference(self, current: Optional[float], previous: Optional[float]) -> Optional[float]:
        """전년동기 대비 차이 계산 (%p)"""
        if current is None or previous is None:
            return None
        return round(current - previous, 1)
    
    def get_quarter_key(self, year: int, quarter: int) -> str:
        """분기 키 생성 (예: 2025_2Q)"""
        return f"{year}_{quarter}Q"
    
    def format_quarter_label(self, year: int, quarter: int) -> str:
        """분기 레이블 포맷 (예: 2025.2/4)"""
        return f"{year}.{quarter}/4"
    
    # =========================================================================
    # 분기 데이터 추출
    # =========================================================================
    
    def extract_quarterly_data(
        self,
        sheet_name: str,
        region: str,
        region_col: int,
        level_col: int,
        level_value: str,
        quarters: List[Tuple[int, int]]
    ) -> Dict[str, Optional[float]]:
        """여러 분기의 데이터 추출
        
        Args:
            sheet_name: 시트 이름
            region: 지역명
            region_col: 지역 컬럼 인덱스
            level_col: 분류단계 컬럼 인덱스
            level_value: 총지수 분류값
            quarters: [(연도, 분기), ...] 리스트
            
        Returns:
            {분기키: 값, ...}
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        config = self.get_sheet_config(sheet_name)
        row_idx = self.find_region_row(df, region, region_col, level_col, level_value)
        if row_idx is None:
            return {}
        
        result = {}
        for year, quarter in quarters:
            key = self.get_quarter_key(year, quarter)
            col_idx = config.get(key)
            if col_idx is not None:
                try:
                    value = df.iloc[row_idx, col_idx]
                    result[key] = self.safe_float(value)
                except (IndexError, ValueError):
                    result[key] = None
            else:
                result[key] = None
        
        return result
    
    def extract_growth_rates(
        self,
        sheet_name: str,
        region: str,
        config: Dict,
        quarters: List[Tuple[int, int]]
    ) -> Dict[str, Optional[float]]:
        """여러 분기의 전년동기비 증감률 계산
        
        Args:
            sheet_name: 시트 이름
            region: 지역명
            config: 시트 설정
            quarters: 비교할 분기 [(연도, 분기), ...]
            
        Returns:
            {분기키: 증감률, ...}
        """
        df = self._load_sheet(sheet_name)
        if df is None:
            return {}
        
        region_col = config.get('region_col', 1)
        level_col = config.get('level_col', 2)
        level_value = config.get('total_code', '0')
        
        row_idx = self.find_region_row(df, region, region_col, level_col, level_value)
        if row_idx is None:
            return {}
        
        result = {}
        for year, quarter in quarters:
            current_key = self.get_quarter_key(year, quarter)
            prev_key = self.get_quarter_key(year - 1, quarter)
            
            current_col = config.get(current_key)
            prev_col = config.get(prev_key)
            
            if current_col is None or prev_col is None:
                result[current_key] = None
                continue
            
            try:
                current_val = self.safe_float(df.iloc[row_idx, current_col])
                prev_val = self.safe_float(df.iloc[row_idx, prev_col])
                result[current_key] = self.calculate_growth_rate(current_val, prev_val)
            except (IndexError, ValueError):
                result[current_key] = None
        
        return result
    
    # =========================================================================
    # 보도자료 공통 구조 생성
    # =========================================================================
    
    def create_report_info(self) -> Dict[str, Any]:
        """보도자료 기본 정보 생성"""
        return {
            'year': self.current_year,
            'quarter': self.current_quarter,
            'period': f"{self.current_year}년 {self.current_quarter}분기"
        }
    
    def create_empty_summary_table(self, columns: Dict[str, List[str]]) -> Dict[str, Any]:
        """빈 요약 테이블 생성"""
        return {
            'columns': columns,
            'national': None,
            'rows': []
        }
    
    def sort_regions_by_value(
        self,
        region_data: Dict[str, Optional[float]],
        descending: bool = True
    ) -> List[Dict[str, Any]]:
        """지역을 값 기준으로 정렬"""
        items = []
        for region, value in region_data.items():
            if region == '전국' or value is None:
                continue
            items.append({'region': region, 'value': value})
        
        items.sort(key=lambda x: x['value'], reverse=descending)
        return items
