"""
유연한 매핑 모듈
시트명과 컬럼명이 뒤죽박죽으로 섞여 있어도 자동으로 매핑하는 기능 제공
"""

import re
from typing import Dict, List, Optional, Tuple
from difflib import SequenceMatcher


class FlexibleMapper:
    """유연한 시트 및 컬럼 매핑 클래스"""
    
    def __init__(self, excel_extractor):
        """
        매퍼 초기화
        
        Args:
            excel_extractor: ExcelExtractor 인스턴스
        """
        self.excel_extractor = excel_extractor
        self.sheet_cache = {}  # 시트 정보 캐시
        self.header_cache = {}  # 헤더 정보 캐시
    
    def find_sheet_by_name(self, target_sheet_name: str, similarity_threshold: float = 0.3) -> Optional[str]:
        """
        시트명을 유연하게 찾습니다. 키워드 기반 자연어 처리 사용.
        
        Args:
            target_sheet_name: 찾고자 하는 시트명 또는 키워드
            similarity_threshold: 유사도 임계값 (0.0 ~ 1.0, 기본값 0.3으로 낮춤)
            
        Returns:
            실제 시트명 또는 None
        """
        if not self.excel_extractor.workbook:
            self.excel_extractor.load_workbook()
        
        available_sheets = self.excel_extractor.get_sheet_names()
        
        # 정확한 매칭
        for sheet in available_sheets:
            if sheet == target_sheet_name:
                return sheet
        
        # 부분 매칭 (포함 관계)
        target_normalized = self._normalize_name(target_sheet_name)
        for sheet in available_sheets:
            sheet_normalized = self._normalize_name(sheet)
            if target_normalized in sheet_normalized or sheet_normalized in target_normalized:
                return sheet
        
        # 키워드 기반 매칭 (새로운 방식)
        target_keywords = self._extract_keywords(target_sheet_name)
        best_match = None
        best_score = 0.0
        
        for sheet in available_sheets:
            sheet_keywords = self._extract_keywords(sheet)
            
            # 키워드 교집합 점수
            if target_keywords and sheet_keywords:
                common_keywords = set(target_keywords) & set(sheet_keywords)
                if common_keywords:
                    keyword_score = len(common_keywords) / max(len(target_keywords), len(sheet_keywords))
                    if keyword_score > best_score:
                        best_score = keyword_score
                        best_match = sheet
        
        # 키워드 매칭이 실패하면 유사도 기반 매칭
        if not best_match:
            for sheet in available_sheets:
                score = self._calculate_similarity(target_sheet_name, sheet)
                if score > best_score and score >= similarity_threshold:
                    best_score = score
                    best_match = sheet
        
        return best_match
    
    def find_column_by_header(
        self, 
        sheet_name: str, 
        target_header: str, 
        header_row: int = 1,
        similarity_threshold: float = 0.6
    ) -> Optional[str]:
        """
        헤더 텍스트로 컬럼을 유연하게 찾습니다.
        
        Args:
            sheet_name: 시트 이름
            target_header: 찾고자 하는 헤더 텍스트
            header_row: 헤더가 있는 행 번호
            similarity_threshold: 유사도 임계값
            
        Returns:
            컬럼 문자 (예: 'A', 'B') 또는 None
        """
        sheet = self.excel_extractor.get_sheet(sheet_name)
        
        # 헤더 캐시 확인
        cache_key = f"{sheet_name}_{header_row}"
        if cache_key not in self.header_cache:
            self.header_cache[cache_key] = self._parse_headers(sheet, header_row)
        
        headers = self.header_cache[cache_key]
        
        # 정확한 매칭
        target_normalized = self._normalize_name(target_header)
        for col_info in headers:
            if col_info['normalized'] == target_normalized:
                return col_info['letter']
        
        # 부분 매칭
        for col_info in headers:
            header_text = col_info['header']
            header_normalized = col_info['normalized']
            
            # 양방향 포함 관계 확인
            if (target_normalized in header_normalized or 
                header_normalized in target_normalized or
                target_header in header_text or
                header_text in target_header):
                return col_info['letter']
        
        # 키워드 기반 매칭
        target_keywords = self._extract_keywords(target_header)
        for col_info in headers:
            header_keywords = self._extract_keywords(col_info['header'])
            # 키워드가 일치하는 경우
            if target_keywords and any(kw in header_keywords for kw in target_keywords):
                return col_info['letter']
        
        # 유사도 기반 매칭
        best_match = None
        best_score = 0.0
        
        for col_info in headers:
            score = self._calculate_similarity(target_header, col_info['header'])
            if score > best_score and score >= similarity_threshold:
                best_score = score
                best_match = col_info['letter']
        
        return best_match
    
    def find_row_by_value(
        self,
        sheet_name: str,
        column: str,
        target_value: str,
        start_row: int = 2,
        similarity_threshold: float = 0.7
    ) -> Optional[int]:
        """
        특정 컬럼에서 값을 유연하게 찾아 행 번호를 반환합니다.
        
        Args:
            sheet_name: 시트 이름
            column: 컬럼 문자 (예: 'A', 'B')
            target_value: 찾고자 하는 값
            start_row: 검색 시작 행
            similarity_threshold: 유사도 임계값
            
        Returns:
            행 번호 또는 None
        """
        sheet = self.excel_extractor.get_sheet(sheet_name)
        col_num = self._column_letter_to_number(column)
        
        target_normalized = self._normalize_name(target_value)
        
        # 정확한 매칭
        for row in range(start_row, min(start_row + 200, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=col_num)
            if cell.value:
                cell_value = str(cell.value).strip()
                if cell_value == target_value:
                    return row
                if self._normalize_name(cell_value) == target_normalized:
                    return row
        
        # 부분 매칭
        for row in range(start_row, min(start_row + 200, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=col_num)
            if cell.value:
                cell_value = str(cell.value).strip()
                cell_normalized = self._normalize_name(cell_value)
                if (target_normalized in cell_normalized or 
                    cell_normalized in target_normalized):
                    return row
        
        # 유사도 기반 매칭
        best_match = None
        best_score = 0.0
        
        for row in range(start_row, min(start_row + 200, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=col_num)
            if cell.value:
                cell_value = str(cell.value).strip()
                score = self._calculate_similarity(target_value, cell_value)
                if score > best_score and score >= similarity_threshold:
                    best_score = score
                    best_match = row
        
        return best_match
    
    def resolve_marker(
        self,
        marker_sheet_name: str,
        marker_key: str,
        header_row: int = 1
    ) -> Optional[Tuple[str, str]]:
        """
        마커를 실제 시트명과 컬럼 주소로 해석합니다.
        
        Args:
            marker_sheet_name: 마커에 있는 시트명
            marker_key: 마커 키 (컬럼 주소 또는 헤더 기반 키)
            header_row: 헤더가 있는 행 번호
            
        Returns:
            (실제_시트명, 컬럼_주소) 튜플 또는 None
        """
        # 시트명 찾기
        actual_sheet = self.find_sheet_by_name(marker_sheet_name)
        if not actual_sheet:
            return None
        
        # 컬럼 주소 형식인지 확인 (A1, B2 등)
        if re.match(r'^[A-Z]+\d+$', marker_key):
            # 이미 컬럼 주소 형식
            return (actual_sheet, marker_key)
        
        # 헤더 기반 키인 경우
        # 형식: "전국_증감률", "서울_2023_3_4" 등
        parts = marker_key.split('_')
        
        # 행 헤더와 컬럼 헤더 분리 시도
        # 첫 번째 부분이 지역명일 가능성이 높음
        if len(parts) >= 2:
            # 마지막 부분이 컬럼 헤더일 가능성
            col_header = '_'.join(parts[-2:]) if len(parts) >= 2 else parts[-1]
            
            # 컬럼 찾기
            col_letter = self.find_column_by_header(actual_sheet, col_header, header_row)
            if col_letter:
                # 행도 찾기
                row_header = parts[0] if len(parts) > 2 else None
                if row_header:
                    row_num = self.find_row_by_value(actual_sheet, 'A', row_header)
                    if row_num:
                        return (actual_sheet, f"{col_letter}{row_num}")
                    # 행을 찾지 못하면 컬럼만 반환
                    return (actual_sheet, col_letter)
                return (actual_sheet, col_letter)
        
        # 전체 키로 컬럼 찾기
        col_letter = self.find_column_by_header(actual_sheet, marker_key, header_row)
        if col_letter:
            return (actual_sheet, col_letter)
        
        return None
    
    def _parse_headers(self, sheet, header_row: int) -> List[Dict]:
        """시트의 헤더를 파싱합니다."""
        headers = []
        max_col = sheet.max_column
        
        for col in range(1, min(max_col + 1, 200)):  # 최대 200개 컬럼
            cell = sheet.cell(row=header_row, column=col)
            header_value = cell.value
            
            if header_value is None:
                continue
            
            header_value = str(header_value).strip()
            if not header_value:
                continue
            
            col_letter = self._number_to_column_letter(col)
            normalized = self._normalize_name(header_value)
            
            headers.append({
                'index': col,
                'letter': col_letter,
                'header': header_value,
                'normalized': normalized
            })
        
        return headers
    
    def _normalize_name(self, name: str) -> str:
        """이름을 정규화합니다."""
        if not name:
            return ""
        
        # 공백, 특수문자 제거 및 언더스코어로 대체
        normalized = re.sub(r'[^\w가-힣]', '_', str(name))
        # 연속된 언더스코어를 하나로
        normalized = re.sub(r'_+', '_', normalized)
        # 앞뒤 언더스코어 제거
        normalized = normalized.strip('_').lower()
        
        return normalized
    
    def _extract_keywords(self, text: str) -> List[str]:
        """텍스트에서 키워드를 추출합니다."""
        if not text:
            return []
        
        # 숫자, 한글, 영문 단어 추출
        keywords = re.findall(r'[가-힣]+|[a-zA-Z]+|\d+', text)
        # 1자리 숫자나 너무 짧은 단어 제외
        keywords = [kw for kw in keywords if len(kw) > 1 or kw.isdigit()]
        
        return keywords
    
    def _calculate_similarity(self, str1: str, str2: str) -> float:
        """두 문자열의 유사도를 계산합니다."""
        if not str1 or not str2:
            return 0.0
        
        # SequenceMatcher를 사용한 유사도 계산
        return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()
    
    def _number_to_column_letter(self, col_num: int) -> str:
        """열 번호를 열 문자로 변환"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(65 + (col_num % 26)) + result
            col_num //= 26
        return result
    
    def _column_letter_to_number(self, col_letter: str) -> int:
        """열 문자를 열 번호로 변환"""
        col_num = 0
        for char in col_letter.upper():
            col_num = col_num * 26 + (ord(char) - ord('A') + 1)
        return col_num

