"""
동적 시트 파서 모듈
헤더를 기반으로 시트 구조를 동적으로 파싱하여 데이터를 추출
가중치 및 분류단계 필터링 지원
"""

import re
from typing import Dict, List, Optional, Tuple, Any
from .excel_extractor import ExcelExtractor
from .excel_header_parser import ExcelHeaderParser
from .schema_loader import SchemaLoader


class DynamicSheetParser:
    """헤더 기반으로 시트를 동적으로 파싱하는 클래스"""
    
    def __init__(self, excel_extractor: ExcelExtractor, schema_loader: Optional[SchemaLoader] = None):
        """
        동적 시트 파서 초기화
        
        Args:
            excel_extractor: 엑셀 추출기 인스턴스
            schema_loader: 스키마 로더 인스턴스 (기본값: 새로 생성)
        """
        self.excel_extractor = excel_extractor
        self.schema_loader = schema_loader if schema_loader is not None else SchemaLoader()
        self.header_parser = ExcelHeaderParser(excel_extractor.excel_path)
        self.sheet_structure_cache = {}  # 시트별 구조 캐시
    
    def parse_sheet_structure(self, sheet_name: str) -> Dict[str, Any]:
        """
        시트의 구조를 동적으로 파싱합니다.
        스키마에서 가중치 및 분류단계 설정을 가져옵니다.
        
        Args:
            sheet_name: 시트 이름
            
        Returns:
            시트 구조 정보:
            {
                'region_column': 지역 이름이 있는 열 번호,
                'category_column': 산업/업태/품목 이름이 있는 열 번호,
                'classification_column': 분류 단계가 있는 열 번호 (선택),
                'weight_column': 가중치 열 번호 (선택),
                'weight_default': 가중치 기본값,
                'max_classification_level': 최대 분류단계,
                'quarter_columns': {
                    (2023, 3): 열 번호,
                    (2023, 4): 열 번호,
                    ...
                },
                'header_row': 헤더가 있는 행 번호,
                'data_start_row': 데이터가 시작하는 행 번호
            }
        """
        if sheet_name in self.sheet_structure_cache:
            return self.sheet_structure_cache[sheet_name]
        
        sheet = self.excel_extractor.get_sheet(sheet_name)
        headers_info = self.header_parser.parse_sheet_headers(sheet_name, header_row=1, max_header_rows=5)
        
        # 스키마에서 가중치 설정 가져오기
        weight_config = self.schema_loader.get_weight_config(sheet_name)
        
        # 지역 이름 열 찾기
        region_column = self._find_region_column(sheet, headers_info)
        
        # 산업/업태/품목 이름 열 찾기
        category_column = self._find_category_column(sheet, headers_info)
        
        # 분류 단계 열: 스키마에서 가져오거나 동적으로 찾기
        classification_column = weight_config.get('classification_column')
        if classification_column is None:
            classification_column = self._find_classification_column(sheet, headers_info)
        
        # 가중치 열: 스키마에서 가져오거나 동적으로 찾기
        weight_column = weight_config.get('weight_column')
        if weight_column is None:
            weight_column = self._find_weight_column(sheet, headers_info)
        
        # 분기별 열 찾기
        quarter_columns = self._find_quarter_columns(sheet, headers_info)
        
        # 헤더 행 찾기
        header_row = self._find_header_row(sheet)
        
        # 데이터 시작 행 찾기
        data_start_row = self._find_data_start_row(sheet, header_row, region_column)
        
        structure = {
            'region_column': region_column,
            'category_column': category_column,
            'classification_column': classification_column,
            'weight_column': weight_column,
            'weight_default': weight_config.get('weight_default', 1),
            'max_classification_level': weight_config.get('max_classification_level', 2),
            'use_weighted_ranking': weight_config.get('use_weighted_ranking', True),
            'quarter_columns': quarter_columns,
            'header_row': header_row,
            'data_start_row': data_start_row,
            'headers_info': headers_info
        }
        
        self.sheet_structure_cache[sheet_name] = structure
        return structure
    
    def _find_region_column(self, sheet, headers_info: Dict) -> int:
        """지역 이름이 있는 열을 찾습니다."""
        # 헤더에서 "시도", "지역" 등의 키워드로 찾기
        for col_info in headers_info['columns']:
            header_lower = col_info['header'].lower()
            if any(keyword in header_lower for keyword in ['시도', '지역', 'region', 'area', '시·도']):
                return col_info['index']
        
        # 기본값: 2번째 열 (B열)
        return 2
    
    def _find_category_column(self, sheet, headers_info: Dict) -> int:
        """산업/업태/품목 이름이 있는 열을 찾습니다."""
        # 헤더에서 "산업", "업태", "품목" 등의 키워드로 찾기
        for col_info in headers_info['columns']:
            header_lower = col_info['header'].lower()
            if any(keyword in header_lower for keyword in ['산업', '업태', '품목', 'category', 'industry', 'item']):
                return col_info['index']
        
        # 기본값: 6번째 열 (F열)
        return 6
    
    def _find_classification_column(self, sheet, headers_info: Dict) -> Optional[int]:
        """분류 단계가 있는 열을 찾습니다."""
        # 헤더에서 "분류", "단계", "level" 등의 키워드로 찾기
        for col_info in headers_info['columns']:
            header_lower = col_info['header'].lower()
            if any(keyword in header_lower for keyword in ['분류', '단계', 'level', 'classification']):
                return col_info['index']
        
        # 기본값: 3번째 열 (C열)
        return 3
    
    def _find_weight_column(self, sheet, headers_info: Dict) -> Optional[int]:
        """가중치 열을 찾습니다."""
        # 헤더에서 "가중치", "weight" 등의 키워드로 찾기
        for col_info in headers_info['columns']:
            header_lower = col_info['header'].lower()
            if any(keyword in header_lower for keyword in ['가중치', 'weight']):
                return col_info['index']
        
        # 기본값: 4번째 열 (D열)
        return 4
    
    def _find_quarter_columns(self, sheet, headers_info: Dict) -> Dict[Tuple[int, int], int]:
        """분기별 열을 찾습니다."""
        quarter_columns = {}
        
        for col_info in headers_info['columns']:
            header = col_info['header']
            # "2023 3/4", "2024 1/4", "2025 2/4p" 등의 형식 파싱
            period = self._parse_period_from_header(header)
            if period:
                year, quarter = period
                quarter_columns[(year, quarter)] = col_info['index']
        
        return quarter_columns
    
    def _parse_period_from_header(self, header: str) -> Optional[Tuple[int, int]]:
        """헤더에서 연도와 분기를 파싱합니다."""
        # "2023 3/4", "2024 1/4", "2025 2/4p" 등의 패턴 (분기별만 인식)
        # 주의: "2024. 1" 같은 월별 패턴은 분기로 잘못 해석될 수 있으므로 제외
        patterns = [
            r'(\d{4})\s+(\d)/4[pP]?',  # "2023 3/4", "2025 2/4p"
            r'(\d{4})\s*(\d)/4[pP]?',  # "2023 3/4" (공백 없음)
        ]
        
        for pattern in patterns:
            match = re.search(pattern, header)
            if match:
                year = int(match.group(1))
                quarter = int(match.group(2))
                if 2000 <= year <= 2100 and 1 <= quarter <= 4:
                    return (year, quarter)
        
        return None
    
    def _find_header_row(self, sheet) -> int:
        """헤더가 있는 행을 찾습니다."""
        # 처음 10개 행 중에서 헤더 찾기
        for row in range(1, min(11, sheet.max_row + 1)):
            # 헤더 행의 특징: 숫자가 아닌 텍스트가 많고, "시도", "지역", "2023" 등의 키워드 포함
            text_count = 0
            keyword_count = 0
            
            for col in range(1, min(20, sheet.max_column + 1)):
                cell = sheet.cell(row=row, column=col)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    # 숫자가 아닌 텍스트
                    if not re.match(r'^-?\d+\.?\d*$', cell_str):
                        text_count += 1
                        # 키워드 확인
                        if any(keyword in cell_str for keyword in ['시도', '지역', '2023', '2024', '2025', '분기', '3/4', '1/4']):
                            keyword_count += 1
            
            # 텍스트가 많고 키워드가 있으면 헤더 행
            if text_count >= 3 and keyword_count >= 2:
                return row
        
        # 기본값: 1행
        return 1
    
    def _find_data_start_row(self, sheet, header_row: int, region_column: int) -> int:
        """데이터가 시작하는 행을 찾습니다."""
        # 헤더 행 다음부터 확인
        for row in range(header_row + 1, min(header_row + 20, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=region_column)
            if cell.value:
                cell_str = str(cell.value).strip()
                # 지역 이름이 있는 행 (예: "전국", "서울" 등)
                if any(region in cell_str for region in ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']):
                    return row
        
        # 기본값: 헤더 행 + 1
        return header_row + 1
    
    def find_region_row(self, sheet_name: str, region_name: str, category_filter: Optional[str] = None) -> Optional[int]:
        """
        특정 지역의 행을 찾습니다.
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름
            category_filter: 카테고리 필터 (예: '총지수', '계') - None이면 모든 행 허용
            
        Returns:
            행 번호 또는 None
        """
        structure = self.parse_sheet_structure(sheet_name)
        sheet = self.excel_extractor.get_sheet(sheet_name)
        region_col = structure['region_column']
        data_start_row = structure['data_start_row']
        category_col = structure.get('category_column')
        classification_col = structure.get('classification_column')
        
        # 지역 이름으로 행 찾기
        for row in range(data_start_row, min(data_start_row + 500, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=region_col)
            if not cell.value:
                continue
            
            cell_str = str(cell.value).strip()
            if cell_str != region_name:
                continue
            
            # 카테고리 필터 적용
            if category_filter and category_col:
                cat_cell = sheet.cell(row=row, column=category_col)
                if cat_cell.value:
                    cat_str = str(cat_cell.value).strip()
                    if cat_str != category_filter:
                        continue
                else:
                    # 카테고리가 비어있으면 스킵 (필터가 있는 경우)
                    continue
            
            # 수출/수입 시트의 경우: 카테고리나 분류 단계 체크 완화
            if sheet_name in ['수출', '수입']:
                # 카테고리 열이 비어있거나 None이면 OK
                if category_col and not category_filter:
                    cat_cell = sheet.cell(row=row, column=category_col)
                    if cat_cell.value:
                        cat_str = str(cat_cell.value).strip()
                        # "총지수"나 "계"가 아니면 스킵할 수도 있지만, 일단 허용
                        if cat_str and cat_str not in ['총지수', '계', '   계']:
                            # 품목별 데이터일 수 있으므로 계속 확인
                            pass
                return row
            else:
                # 다른 시트: 분류 단계가 0이거나 카테고리가 "총지수"/"계"인 경우
                if classification_col:
                    class_cell = sheet.cell(row=row, column=classification_col)
                    if class_cell.value and str(class_cell.value).strip() not in ['0', '0.0']:
                        continue
                
                if category_col and not category_filter:
                    cat_cell = sheet.cell(row=row, column=category_col)
                    if cat_cell.value:
                        cat_str = str(cat_cell.value).strip()
                        if cat_str in ['총지수', '계', '   계']:
                            return row
                    else:
                        # 카테고리가 비어있으면 스킵
                        continue
                else:
                    return row
        
        return None
    
    def get_quarter_column(self, sheet_name: str, year: int, quarter: int) -> Optional[int]:
        """
        특정 분기의 열 번호를 반환합니다.
        
        Args:
            sheet_name: 시트 이름
            year: 연도
            quarter: 분기
            
        Returns:
            열 번호 또는 None
        """
        structure = self.parse_sheet_structure(sheet_name)
        quarter_columns = structure['quarter_columns']
        
        return quarter_columns.get((year, quarter))
    
    def get_quarter_value(self, sheet_name: str, region_name: str, year: int, quarter: int) -> Optional[float]:
        """
        특정 지역의 특정 분기 값을 가져옵니다.
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름
            year: 연도
            quarter: 분기
            
        Returns:
            값 또는 None
        """
        structure = self.parse_sheet_structure(sheet_name)
        region_row = self.find_region_row(sheet_name, region_name)
        
        if region_row is None:
            return None
        
        quarter_col = self.get_quarter_column(sheet_name, year, quarter)
        if quarter_col is None:
            return None
        
        sheet = self.excel_extractor.get_sheet(sheet_name)
        cell = sheet.cell(row=region_row, column=quarter_col)
        
        # 스키마에서 가중치 설정 가져오기
        weight_col = structure.get('weight_column')
        weight_default = structure.get('weight_default', 1)
        
        # 가중치 확인
        if weight_col:
            weight_value = sheet.cell(row=region_row, column=weight_col).value
            weight = self.schema_loader.get_weight_value(sheet_name, weight_value)
        else:
            weight = weight_default
        
        # 값이 비어있고 가중치가 100이면 기본값 100 반환
        if cell.value is None or (isinstance(cell.value, str) and not cell.value.strip()):
            if weight == 100:
                return 100.0
            return None
        
        try:
            return float(cell.value)
        except (ValueError, TypeError):
            # 변환 실패 시 가중치가 100이면 기본값 반환
            if weight == 100:
                return 100.0
            return None
    
    def calculate_growth_rate(self, sheet_name: str, region_name: str, year: int, quarter: int) -> Optional[float]:
        """
        특정 지역의 특정 분기 증감률을 계산합니다.
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름
            year: 연도
            quarter: 분기
            
        Returns:
            증감률 (퍼센트) 또는 None
        """
        # 현재 분기 값
        current_value = self.get_quarter_value(sheet_name, region_name, year, quarter)
        if current_value is None:
            return None
        
        # 전년 동분기 값
        prev_year = year - 1
        prev_value = self.get_quarter_value(sheet_name, region_name, prev_year, quarter)
        if prev_value is None or prev_value == 0:
            return None
        
        # 증감률 계산
        try:
            growth_rate = ((current_value / prev_value) - 1) * 100
            import math
            if math.isnan(growth_rate) or math.isinf(growth_rate):
                return None
            return growth_rate
        except (ZeroDivisionError, OverflowError):
            return None
    
    def close(self):
        """리소스 정리"""
        if self.header_parser:
            self.header_parser.close()

