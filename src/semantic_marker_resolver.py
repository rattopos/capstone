"""
의미 기반 마커 해석 모듈
마커의 의미를 파악하여 엑셀 데이터와 유기적으로 연결
"""

import re
from typing import Dict, List, Optional, Tuple, Any
from difflib import SequenceMatcher


class SemanticMarkerResolver:
    """의미 기반 마커를 해석하여 실제 데이터 위치를 찾는 클래스"""
    
    # 데이터 의미 키워드 매핑
    DATA_SEMANTIC_MAPPING = {
        # 지역 관련
        '전국': ['전국', 'national', '전체', '합계'],
        '서울': ['서울', 'seoul'],
        '부산': ['부산', 'busan'],
        '대구': ['대구', 'daegu'],
        '인천': ['인천', 'incheon'],
        '광주': ['광주', 'gwangju'],
        '대전': ['대전', 'daejeon'],
        '울산': ['울산', 'ulsan'],
        '세종': ['세종', 'sejong'],
        '경기': ['경기', 'gyeonggi'],
        '강원': ['강원', 'gangwon'],
        '충북': ['충북', 'chungbuk'],
        '충남': ['충남', 'chungnam'],
        '전북': ['전북', 'jeonbuk'],
        '전남': ['전남', 'jeonnam'],
        '경북': ['경북', 'gyeongbuk'],
        '경남': ['경남', 'gyeongnam'],
        '제주': ['제주', 'jeju'],
        
        # 데이터 타입 관련
        '증감률': ['증감률', '증가율', '감소율', '변화율', 'growth', 'rate', '변동률'],
        '증감액': ['증감액', '증가액', '감소액', '변화액', 'change', 'amount'],
        '값': ['값', 'value', '수치', '데이터'],
        '지수': ['지수', 'index'],
        '총지수': ['총지수', '계', 'total', '합계'],
    }
    
    def __init__(self, excel_extractor):
        """
        의미 기반 마커 해석기 초기화
        
        Args:
            excel_extractor: ExcelExtractor 인스턴스
        """
        self.excel_extractor = excel_extractor
        self.sheet_cache = {}
        self.header_cache = {}
    
    def resolve_semantic_marker(
        self,
        marker_text: str,
        year: int = None,
        quarter: int = None
    ) -> Optional[Any]:
        """
        의미 기반 마커를 해석하여 실제 데이터를 반환합니다.
        
        Args:
            marker_text: 마커 텍스트 (예: "{경제지표:전국_증감률}", "{전국_증감률}")
            year: 연도 (선택적)
            quarter: 분기 (선택적)
            
        Returns:
            찾은 데이터 값 또는 None
        """
        # 마커 형식 파싱
        marker_info = self._parse_marker(marker_text)
        if not marker_info:
            return None
        
        sheet_keyword = marker_info.get('sheet_keyword')
        data_keyword = marker_info.get('data_keyword')
        
        # 1. 시트 찾기 (키워드 기반)
        from src.semantic_sheet_matcher import SemanticSheetMatcher
        semantic_matcher = SemanticSheetMatcher(self.excel_extractor)
        
        if sheet_keyword:
            # 시트 키워드가 있는 경우
            sheet_name = semantic_matcher.find_sheet_by_semantic_keywords(sheet_keyword)
        else:
            # 시트 키워드가 없는 경우, 데이터 키워드에서 추론
            sheet_name = semantic_matcher.find_sheet_by_semantic_keywords(data_keyword)
        
        if not sheet_name:
            return None
        
        # 2. 데이터 위치 찾기
        data_location = self._find_data_location(sheet_name, data_keyword, year, quarter)
        if not data_location:
            return None
        
        # 3. 데이터 추출
        return self._extract_data(sheet_name, data_location)
    
    def _parse_marker(self, marker_text: str) -> Optional[Dict[str, str]]:
        """
        마커 텍스트를 파싱합니다.
        
        지원 형식:
        - {시트키워드:데이터키워드}
        - {데이터키워드} (시트키워드 생략)
        - {전국_증감률}
        - {경제지표:전국_증감률}
        """
        # 중괄호 제거
        marker_text = marker_text.strip()
        if marker_text.startswith('{') and marker_text.endswith('}'):
            marker_text = marker_text[1:-1]
        
        # 콜론으로 분리
        if ':' in marker_text:
            parts = marker_text.split(':', 1)
            sheet_keyword = parts[0].strip()
            data_keyword = parts[1].strip()
        else:
            # 시트 키워드가 없는 경우
            sheet_keyword = None
            data_keyword = marker_text.strip()
        
        return {
            'sheet_keyword': sheet_keyword,
            'data_keyword': data_keyword,
            'original': marker_text
        }
    
    def _find_data_location(
        self,
        sheet_name: str,
        data_keyword: str,
        year: int = None,
        quarter: int = None
    ) -> Optional[Dict[str, Any]]:
        """
        데이터 키워드를 기반으로 실제 데이터 위치를 찾습니다.
        
        Returns:
            데이터 위치 정보 딕셔너리:
            - 'type': 'cell' (셀 주소) 또는 'header' (헤더 기반)
            - 'location': 셀 주소 또는 헤더 정보
        """
        sheet = self.excel_extractor.get_sheet(sheet_name)
        
        # 데이터 키워드 파싱
        parts = data_keyword.split('_')
        
        # 지역명과 데이터 타입 분리
        region_keyword = None
        data_type_keyword = None
        
        if len(parts) >= 2:
            # 첫 번째 부분이 지역명일 가능성
            potential_region = parts[0]
            if self._is_region_keyword(potential_region):
                region_keyword = potential_region
                data_type_keyword = '_'.join(parts[1:])
            else:
                # 지역명이 없는 경우
                data_type_keyword = data_keyword
        else:
            # 단일 키워드
            if self._is_region_keyword(data_keyword):
                region_keyword = data_keyword
            else:
                data_type_keyword = data_keyword
        
        # 헤더 기반 검색
        header_location = self._find_by_headers(
            sheet, region_keyword, data_type_keyword, year, quarter
        )
        if header_location:
            return header_location
        
        # 행/열 기반 검색
        row_col_location = self._find_by_row_col(
            sheet, region_keyword, data_type_keyword, year, quarter
        )
        if row_col_location:
            return row_col_location
        
        return None
    
    def _find_by_headers(
        self,
        sheet,
        region_keyword: Optional[str],
        data_type_keyword: Optional[str],
        year: int = None,
        quarter: int = None
    ) -> Optional[Dict[str, Any]]:
        """헤더를 기반으로 데이터 위치 찾기"""
        from src.flexible_mapper import FlexibleMapper
        flexible_mapper = FlexibleMapper(self.excel_extractor)
        
        # 헤더 파싱
        cache_key = f"{sheet.title}_1"
        if cache_key not in self.header_cache:
            headers = flexible_mapper._parse_headers(sheet, 1)
            self.header_cache[cache_key] = headers
        
        headers = self.header_cache[cache_key]
        
        # 연도/분기 컬럼 찾기
        target_col = None
        if year and quarter:
            # 연도/분기 패턴으로 컬럼 찾기
            period_pattern = f"{year}.*{quarter}"
            for header in headers:
                header_text = str(header['header'])
                if re.search(period_pattern, header_text, re.IGNORECASE):
                    target_col = header['letter']
                    break
        
        # 데이터 타입 컬럼 찾기
        if data_type_keyword and not target_col:
            for header in headers:
                header_text = str(header['header'])
                if self._keyword_match(data_type_keyword, header_text):
                    target_col = header['letter']
                    break
        
        if not target_col:
            return None
        
        # 지역 행 찾기
        target_row = None
        if region_keyword:
            # 첫 번째 열에서 지역명 찾기
            for row in range(1, min(100, sheet.max_row + 1)):
                cell_value = sheet.cell(row=row, column=1).value
                if cell_value and self._keyword_match(region_keyword, str(cell_value)):
                    target_row = row
                    break
        
        if target_col:
            if target_row:
                return {
                    'type': 'cell',
                    'location': f"{target_col}{target_row}"
                }
            else:
                # 컬럼만 있는 경우 (전체 합계 등)
                return {
                    'type': 'column',
                    'location': target_col
                }
        
        return None
    
    def _find_by_row_col(
        self,
        sheet,
        region_keyword: Optional[str],
        data_type_keyword: Optional[str],
        year: int = None,
        quarter: int = None
    ) -> Optional[Dict[str, Any]]:
        """행/열 구조를 기반으로 데이터 위치 찾기"""
        # 일반적인 엑셀 구조:
        # - 행 1-3: 헤더
        # - 행 4부터: 데이터
        # - 열 1: 지역 코드
        # - 열 2: 지역명
        # - 열 3: 분류 단계
        # - 열 4+: 데이터
        
        # 지역 행 찾기
        target_row = None
        if region_keyword:
            for row in range(4, min(200, sheet.max_row + 1)):
                # 열 2 (지역명) 확인
                cell_value = sheet.cell(row=row, column=2).value
                if cell_value and self._keyword_match(region_keyword, str(cell_value)):
                    target_row = row
                    break
        
        # 연도/분기 열 찾기
        target_col = None
        if year and quarter:
            # 행 3에서 연도/분기 정보 찾기
            for col in range(50, min(200, sheet.max_column + 1)):
                cell_value = sheet.cell(row=3, column=col).value
                if cell_value:
                    cell_str = str(cell_value)
                    # "2025 2/4" 형식 확인
                    if f"{year}" in cell_str and f"{quarter}" in cell_str:
                        target_col = col
                        break
        
        if target_row and target_col:
            from src.flexible_mapper import FlexibleMapper
            flexible_mapper = FlexibleMapper(self.excel_extractor)
            col_letter = flexible_mapper._number_to_column_letter(target_col)
            return {
                'type': 'cell',
                'location': f"{col_letter}{target_row}"
            }
        
        return None
    
    def _extract_data(self, sheet_name: str, location: Dict[str, Any]) -> Optional[Any]:
        """데이터 위치 정보를 기반으로 실제 데이터 추출"""
        if location['type'] == 'cell':
            try:
                return self.excel_extractor.get_cell_value(sheet_name, location['location'])
            except:
                return None
        elif location['type'] == 'column':
            # 컬럼 전체는 지원하지 않음
            return None
        
        return None
    
    def _is_region_keyword(self, keyword: str) -> bool:
        """키워드가 지역명인지 확인"""
        keyword_lower = keyword.lower()
        for region, aliases in self.DATA_SEMANTIC_MAPPING.items():
            if region in ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', 
                         '세종', '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']:
                if keyword_lower in [r.lower() for r in aliases]:
                    return True
        return False
    
    def _keyword_match(self, keyword: str, text: str) -> bool:
        """키워드와 텍스트가 매칭되는지 확인"""
        keyword_lower = keyword.lower()
        text_lower = text.lower()
        
        # 정확한 매칭
        if keyword_lower == text_lower:
            return True
        
        # 포함 관계
        if keyword_lower in text_lower or text_lower in keyword_lower:
            return True
        
        # 의미 기반 매칭
        for semantic_key, aliases in self.DATA_SEMANTIC_MAPPING.items():
            if keyword_lower in [a.lower() for a in aliases]:
                if any(alias.lower() in text_lower for alias in aliases):
                    return True
        
        # 유사도 기반 매칭
        similarity = SequenceMatcher(None, keyword_lower, text_lower).ratio()
        if similarity > 0.7:
            return True
        
        return False

