"""
자동 매핑 모듈
이미지에서 추출한 텍스트와 엑셀 시트/열 자동 매칭
"""

import re
from typing import Dict, List, Optional, Tuple, Any
from .excel_extractor import ExcelExtractor
from .image_analyzer import ImageAnalyzer


class AutoMapper:
    """이미지 텍스트와 엑셀 데이터를 자동으로 매칭하는 클래스"""
    
    def __init__(self, excel_extractor: ExcelExtractor):
        """
        자동 매퍼 초기화
        
        Args:
            excel_extractor: 엑셀 추출기 인스턴스
        """
        self.excel_extractor = excel_extractor
    
    def suggest_mappings(
        self,
        image_analysis: Dict[str, Any],
        sheet_names: Optional[List[str]] = None
    ) -> Dict[str, List[Dict[str, Any]]]:
        """
        이미지 분석 결과를 기반으로 엑셀 매핑 제안을 생성합니다.
        
        Args:
            image_analysis: 이미지 분석 결과
            sheet_names: 검색할 시트 이름 리스트 (None이면 모든 시트)
            
        Returns:
            필드별 매핑 제안 딕셔너리. 각 필드에 대해 여러 제안이 있을 수 있음
        """
        if sheet_names is None:
            sheet_names = self.excel_extractor.get_sheet_names()
        
        suggestions = {}
        
        # 이미지에서 추출한 필드들
        fields = image_analysis.get('fields', [])
        
        for field in fields:
            field_id = field.get('field_id', '')
            text_in_image = field.get('text_in_image', '')
            
            # 각 시트에서 매핑 제안 생성
            field_suggestions = []
            
            for sheet_name in sheet_names:
                sheet_suggestions = self._suggest_sheet_mapping(
                    sheet_name, text_in_image, field
                )
                field_suggestions.extend(sheet_suggestions)
            
            # 신뢰도 점수 기준으로 정렬
            field_suggestions.sort(key=lambda x: x.get('confidence', 0), reverse=True)
            
            suggestions[field_id] = field_suggestions
        
        return suggestions
    
    def _suggest_sheet_mapping(
        self,
        sheet_name: str,
        text_in_image: str,
        field: Dict[str, Any]
    ) -> List[Dict[str, Any]]:
        """
        특정 시트에서 매핑 제안 생성
        
        Args:
            sheet_name: 시트 이름
            text_in_image: 이미지의 텍스트
            field: 필드 정보
            
        Returns:
            매핑 제안 리스트
        """
        suggestions = []
        
        try:
            sheet = self.excel_extractor.get_sheet(sheet_name)
            
            # 시트의 헤더 행 찾기 (보통 1-4행)
            header_rows = self._find_header_rows(sheet)
            
            # 텍스트에서 키워드 추출
            keywords = self._extract_keywords(text_in_image)
            
            # 각 헤더 행에서 키워드 매칭
            for row_idx in header_rows:
                row_data = self._get_row_data(sheet, row_idx)
                
                # 키워드 매칭 점수 계산
                match_score = self._calculate_keyword_match(keywords, row_data)
                
                if match_score > 0:
                    # 열 찾기
                    col_idx = self._find_matching_column(sheet, row_idx, keywords)
                    
                    if col_idx:
                        suggestion = {
                            'sheet_name': sheet_name,
                            'row': row_idx,
                            'column': col_idx,
                            'cell_address': self._col_row_to_cell(col_idx, row_idx),
                            'confidence': match_score,
                            'match_type': 'keyword',
                            'matched_keywords': keywords
                        }
                        suggestions.append(suggestion)
            
            # 지역명 매칭 (특수 케이스)
            region_match = self._match_region_name(text_in_image, sheet)
            if region_match:
                suggestions.append(region_match)
            
        except Exception as e:
            print(f"시트 {sheet_name} 매핑 제안 생성 중 오류: {e}")
        
        return suggestions
    
    def _find_header_rows(self, sheet, max_rows: int = 10) -> List[int]:
        """시트에서 헤더 행 찾기"""
        header_rows = []
        
        for row in range(1, min(max_rows + 1, sheet.max_row + 1)):
            # 행에 텍스트가 있는지 확인
            has_text = False
            for col in range(1, min(20, sheet.max_column + 1)):
                cell = sheet.cell(row=row, column=col)
                if cell.value and isinstance(cell.value, str):
                    has_text = True
                    break
            
            if has_text:
                header_rows.append(row)
        
        return header_rows[:5]  # 최대 5개 행
    
    def _get_row_data(self, sheet, row: int) -> List[str]:
        """행의 데이터를 문자열 리스트로 가져오기"""
        row_data = []
        for col in range(1, min(50, sheet.max_column + 1)):
            cell = sheet.cell(row=row, column=col)
            if cell.value:
                row_data.append(str(cell.value).strip())
        return row_data
    
    def _extract_keywords(self, text: str) -> List[str]:
        """텍스트에서 키워드 추출"""
        # 한글 키워드 패턴
        keywords = []
        
        # 지역명
        regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', 
                   '세종', '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        for region in regions:
            if region in text:
                keywords.append(region)
        
        # 일반 키워드
        common_keywords = ['증감률', '증가', '감소', '상위', '하위', '시도', '업태', '산업', 
                          '총지수', '계', '이름', '명']
        for keyword in common_keywords:
            if keyword in text:
                keywords.append(keyword)
        
        return keywords
    
    def _calculate_keyword_match(self, keywords: List[str], row_data: List[str]) -> float:
        """키워드 매칭 점수 계산 (0-1)"""
        if not keywords or not row_data:
            return 0.0
        
        row_text = ' '.join(row_data)
        matches = sum(1 for keyword in keywords if keyword in row_text)
        
        return matches / len(keywords) if keywords else 0.0
    
    def _find_matching_column(self, sheet, row: int, keywords: List[str]) -> Optional[int]:
        """키워드와 매칭되는 열 찾기"""
        for col in range(1, min(100, sheet.max_column + 1)):
            cell = sheet.cell(row=row, column=col)
            if cell.value:
                cell_text = str(cell.value).strip()
                for keyword in keywords:
                    if keyword in cell_text:
                        return col
        
        return None
    
    def _col_row_to_cell(self, col: int, row: int) -> str:
        """열 번호와 행 번호를 셀 주소로 변환"""
        # 열 번호를 문자로 변환 (A=1, B=2, ..., Z=26, AA=27, ...)
        col_letters = ''
        col_num = col
        while col_num > 0:
            col_num -= 1
            col_letters = chr(65 + (col_num % 26)) + col_letters
            col_num //= 26
        
        return f"{col_letters}{row}"
    
    def _match_region_name(self, text: str, sheet) -> Optional[Dict[str, Any]]:
        """지역명 매칭 (특수 케이스)"""
        regions = {
            '전국': '00',
            '서울': '11', '부산': '21', '대구': '22', '인천': '23',
            '광주': '24', '대전': '25', '울산': '26', '세종': '29',
            '경기': '31', '강원': '32', '충북': '33', '충남': '34',
            '전북': '35', '전남': '36', '경북': '37', '경남': '38', '제주': '39'
        }
        
        for region_name, region_code in regions.items():
            if region_name in text:
                # 시트에서 지역 코드로 행 찾기
                for row in range(1, min(1000, sheet.max_row + 1)):
                    cell = sheet.cell(row=row, column=1)  # A열에 지역 코드
                    if cell.value and str(cell.value).strip() == region_code:
                        return {
                            'sheet_name': sheet.title,
                            'row': row,
                            'column': 2,  # B열에 지역 이름
                            'cell_address': f"B{row}",
                            'confidence': 0.9,
                            'match_type': 'region_code',
                            'matched_keywords': [region_name]
                        }
        
        return None
    
    def calculate_mapping_quality(
        self,
        mapping: Dict[str, Any],
        image_analysis: Dict[str, Any]
    ) -> float:
        """
        매핑 품질 점수 계산 (0-1)
        
        Args:
            mapping: 매핑 정보
            image_analysis: 이미지 분석 결과
            
        Returns:
            품질 점수 (0-1)
        """
        score = 0.0
        
        # 키워드 매칭 점수
        if 'matched_keywords' in mapping:
            keywords = mapping.get('matched_keywords', [])
            if keywords:
                score += 0.4
        
        # 신뢰도 점수
        confidence = mapping.get('confidence', 0.0)
        score += confidence * 0.4
        
        # 매핑 타입 점수
        match_type = mapping.get('match_type', '')
        if match_type == 'region_code':
            score += 0.2
        elif match_type == 'keyword':
            score += 0.1
        
        return min(score, 1.0)

