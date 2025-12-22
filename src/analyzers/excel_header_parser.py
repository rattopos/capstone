"""
엑셀 헤더 파서 모듈
엑셀 파일의 헤더를 읽어서 의미있는 객체로 변환하는 기능 제공
"""

import re
from typing import Dict, List, Optional, Any
from pathlib import Path
from ..core.excel_extractor import ExcelExtractor


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
        
        # 성능 최적화: 헤더 캐시
        self._headers_cache: Dict[str, Dict[str, Any]] = {}
    
    def parse_sheet_headers(self, sheet_name: str, header_row: int = 1, max_header_rows: int = 3) -> Dict[str, Any]:
        """
        시트의 헤더를 파싱하여 객체로 만듭니다.
        여러 행의 헤더도 고려합니다.
        
        성능 최적화: 결과를 캐시하여 반복 호출 시 재파싱하지 않습니다.
        
        Args:
            sheet_name: 시트 이름
            header_row: 헤더가 시작하는 행 번호 (기본값: 1)
            max_header_rows: 최대 헤더 행 수 (기본값: 3)
            
        Returns:
            헤더 정보 딕셔너리:
            {
                'columns': [
                    {'index': 1, 'letter': 'A', 'header': '시·도', 'normalized': '시도', 'full_header': '시·도'},
                    {'index': 2, 'letter': 'B', 'header': '2023 3/4', 'normalized': '2023_3_4', 'full_header': '2023 3/4'},
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
        # 캐시 키 생성
        cache_key = f"{sheet_name}_{header_row}_{max_header_rows}"
        if cache_key in self._headers_cache:
            return self._headers_cache[cache_key]
        
        sheet = self.extractor.get_sheet(sheet_name)
        
        # 헤더 행 읽기 (여러 행 고려)
        headers = []
        header_map = {}
        
        max_col = min(sheet.max_column, 200)  # 최대 200개 컬럼
        
        # 여러 행의 헤더를 조합
        for col in range(1, max_col + 1):
            header_parts = []
            
            # 여러 행의 헤더 읽기
            for row_offset in range(max_header_rows):
                row = header_row + row_offset
                if row > sheet.max_row:
                    break
                
                cell = sheet.cell(row=row, column=col)
                cell_value = cell.value
                
                if cell_value is not None:
                    cell_value = str(cell_value).strip()
                    if cell_value:
                        header_parts.append(cell_value)
            
            # 헤더 조합
            if header_parts:
                full_header = ' '.join(header_parts)
                primary_header = header_parts[0]  # 첫 번째 행의 헤더를 주 헤더로
            else:
                full_header = f"Column{col}"
                primary_header = full_header
            
            # 열 문자 계산 (A, B, C, ...)
            col_letter = self._number_to_column_letter(col)
            
            # 헤더 정규화 (마커 이름으로 사용 가능하도록)
            normalized = self._normalize_header(full_header)
            normalized_primary = self._normalize_header(primary_header)
            
            header_info = {
                'index': col,
                'letter': col_letter,
                'header': primary_header,
                'full_header': full_header,
                'normalized': normalized,
                'normalized_primary': normalized_primary
            }
            
            headers.append(header_info)
            # 여러 키로 매핑 (전체 헤더와 주 헤더 모두)
            header_map[normalized] = col_letter
            if normalized_primary != normalized:
                header_map[normalized_primary] = col_letter
        
        # 행 헤더도 찾기 (첫 번째 열의 값들을 행 식별자로 사용)
        row_map = {}
        max_row = sheet.max_row
        
        # 첫 번째 열의 값들을 읽어서 행 매핑 생성
        # 지역명이 있는 열 찾기 (보통 2번째 열)
        region_col = 2  # 기본값: B열
        for col in range(1, min(6, max_col + 1)):  # 처음 5개 열 중에서 찾기
            cell = sheet.cell(row=header_row, column=col)
            if cell.value:
                header_text = str(cell.value).strip().lower()
                if any(keyword in header_text for keyword in ['시도', '지역', 'region', 'area']):
                    region_col = col
                    break
        
        for row in range(header_row + 1, min(header_row + 200, max_row + 1)):
            cell = sheet.cell(row=row, column=region_col)
            row_value = cell.value
            
            if row_value is None:
                continue
            
            row_value = str(row_value).strip()
            if row_value:
                normalized_row = self._normalize_header(row_value)
                if normalized_row not in row_map:  # 중복 방지
                    row_map[normalized_row] = row
        
        result = {
            'columns': headers,
            'header_map': header_map,
            'row_map': row_map,
            'header_row': header_row,
            'region_column': region_col
        }
        
        # 캐시에 저장
        self._headers_cache[cache_key] = result
        
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
        # 캐시 초기화
        self._headers_cache.clear()
    
    def clear_cache(self):
        """캐시만 초기화 (리소스는 유지)"""
        self._headers_cache.clear()

