"""
엑셀 헤더 파서 모듈
엑셀 파일의 헤더를 읽어서 의미있는 객체로 변환하는 기능 제공
"""

import re
from typing import Dict, List, Optional, Any
from pathlib import Path
from .excel_extractor import ExcelExtractor


class ExcelHeaderParser:
    """엑셀 파일의 헤더를 파싱하여 객체로 만드는 클래스"""
    
    def __init__(self, excel_path: str):
        """
        헤더 파서 초기화
        
        Args:
            excel_path: 엑셀 파일 경로
        """
        self.excel_path = excel_path
        self.extractor = ExcelExtractor(excel_path)
        self.extractor.load_workbook()
        self.header_cache = {}  # 시트별 헤더 캐시
    
    def parse_sheet_headers(self, sheet_name: str, header_row: int = 1) -> Dict[str, Any]:
        """
        시트의 헤더를 파싱하여 객체로 만듭니다.
        
        Args:
            sheet_name: 시트 이름
            header_row: 헤더가 있는 행 번호 (기본값: 1)
            
        Returns:
            헤더 정보 딕셔너리:
            {
                'columns': [
                    {'index': 1, 'letter': 'A', 'header': '시·도', 'normalized': '시도'},
                    {'index': 2, 'letter': 'B', 'header': '2023 3/4', 'normalized': '2023_3_4'},
                    ...
                ],
                'header_map': {
                    '시도': 'A',
                    '2023_3_4': 'B',
                    ...
                },
                'row_map': {
                    '전국': 2,
                    '서울': 3,
                    ...
                }
            }
        """
        if sheet_name in self.header_cache:
            return self.header_cache[sheet_name]
        
        sheet = self.extractor.get_sheet(sheet_name)
        
        # 헤더 행 읽기
        headers = []
        header_map = {}
        
        max_col = sheet.max_column
        for col in range(1, max_col + 1):
            cell = sheet.cell(row=header_row, column=col)
            header_value = cell.value
            
            if header_value is None:
                header_value = f"Column{col}"
            else:
                header_value = str(header_value).strip()
            
            # 열 문자 계산 (A, B, C, ...)
            col_letter = self._number_to_column_letter(col)
            
            # 헤더 정규화 (마커 이름으로 사용 가능하도록)
            normalized = self._normalize_header(header_value)
            
            header_info = {
                'index': col,
                'letter': col_letter,
                'header': header_value,
                'normalized': normalized
            }
            
            headers.append(header_info)
            header_map[normalized] = col_letter
        
        # 행 헤더도 찾기 (첫 번째 열의 값들을 행 식별자로 사용)
        row_map = {}
        max_row = sheet.max_row
        
        # 첫 번째 열의 값들을 읽어서 행 매핑 생성
        for row in range(header_row + 1, min(header_row + 50, max_row + 1)):  # 최대 50개 행만
            cell = sheet.cell(row=row, column=1)
            row_value = cell.value
            
            if row_value is None:
                continue
            
            row_value = str(row_value).strip()
            if row_value:
                normalized_row = self._normalize_header(row_value)
                row_map[normalized_row] = row
        
        result = {
            'columns': headers,
            'header_map': header_map,
            'row_map': row_map,
            'header_row': header_row
        }
        
        self.header_cache[sheet_name] = result
        return result
    
    def _number_to_column_letter(self, col_num: int) -> str:
        """열 번호를 열 문자로 변환 (1 -> A, 2 -> B, ...)"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(65 + (col_num % 26)) + result
            col_num //= 26
        return result
    
    def _normalize_header(self, header: str) -> str:
        """
        헤더를 정규화하여 마커 이름으로 사용 가능하게 만듭니다.
        
        Args:
            header: 원본 헤더 문자열
            
        Returns:
            정규화된 헤더 문자열
        """
        # 공백, 특수문자 제거 및 언더스코어로 대체
        normalized = re.sub(r'[^\w가-힣]', '_', header)
        
        # 연속된 언더스코어를 하나로
        normalized = re.sub(r'_+', '_', normalized)
        
        # 앞뒤 언더스코어 제거
        normalized = normalized.strip('_')
        
        # 빈 문자열이면 기본값
        if not normalized:
            normalized = 'column'
        
        return normalized
    
    def get_marker_name(self, sheet_name: str, col_header: str, row_header: Optional[str] = None) -> str:
        """
        헤더 기반으로 마커 이름을 생성합니다.
        
        Args:
            sheet_name: 시트 이름
            col_header: 열 헤더 (정규화된 형태 또는 원본)
            row_header: 행 헤더 (선택사항, 정규화된 형태 또는 원본)
            
        Returns:
            마커 이름 (예: '{광공업생산:전국_증감률}' 또는 '{광공업생산:서울_2023_3_4}')
        """
        headers_info = self.parse_sheet_headers(sheet_name)
        
        # 열 헤더 정규화
        normalized_col = self._normalize_header(col_header)
        
        # 행 헤더가 있으면 조합
        if row_header:
            normalized_row = self._normalize_header(row_header)
            marker_name = f"{normalized_row}_{normalized_col}"
        else:
            marker_name = normalized_col
        
        return f"{{{sheet_name}:{marker_name}}}"
    
    def find_column_by_header(self, sheet_name: str, header_text: str) -> Optional[str]:
        """
        헤더 텍스트로 열을 찾습니다.
        
        Args:
            sheet_name: 시트 이름
            header_text: 헤더 텍스트 (부분 일치 가능)
            
        Returns:
            열 문자 (예: 'A', 'B') 또는 None
        """
        headers_info = self.parse_sheet_headers(sheet_name)
        
        # 정확한 매칭 시도
        normalized = self._normalize_header(header_text)
        if normalized in headers_info['header_map']:
            return headers_info['header_map'][normalized]
        
        # 부분 매칭 시도
        for col_info in headers_info['columns']:
            if header_text in col_info['header'] or col_info['header'] in header_text:
                return col_info['letter']
        
        return None
    
    def find_row_by_header(self, sheet_name: str, row_text: str) -> Optional[int]:
        """
        행 헤더 텍스트로 행을 찾습니다.
        
        Args:
            sheet_name: 시트 이름
            row_text: 행 헤더 텍스트
            
        Returns:
            행 번호 또는 None
        """
        headers_info = self.parse_sheet_headers(sheet_name)
        
        # 정규화된 행 헤더로 찾기
        normalized = self._normalize_header(row_text)
        if normalized in headers_info['row_map']:
            return headers_info['row_map'][normalized]
        
        # 부분 매칭 시도
        for row_name, row_num in headers_info['row_map'].items():
            if row_text in row_name or row_name in row_text:
                return row_num
        
        return None
    
    def get_all_headers(self, sheet_name: str) -> Dict[str, Any]:
        """
        시트의 모든 헤더 정보를 반환합니다.
        
        Args:
            sheet_name: 시트 이름
            
        Returns:
            헤더 정보 딕셔너리
        """
        return self.parse_sheet_headers(sheet_name)
    
    def close(self):
        """리소스 정리"""
        if self.extractor:
            self.extractor.close()

