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
        
        # 성능 최적화: 시트 구조 캐시
        self._structure_cache: Dict[str, Dict[str, Any]] = {}
        # 성능 최적화: 지역 행 캐시 (시트명_지역명_카테고리 -> 행 번호)
        self._region_row_cache: Dict[str, Optional[int]] = {}
    
    def parse_sheet_structure(self, sheet_name: str) -> Dict[str, Any]:
        """
        시트의 구조를 동적으로 파싱합니다.
        스키마에서 가중치 및 분류단계 설정을 가져옵니다.
        
        성능 최적화: 결과를 캐시하여 반복 호출 시 재계산하지 않습니다.
        
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
        # 캐시 확인
        if sheet_name in self._structure_cache:
            return self._structure_cache[sheet_name]
        
        sheet = self.excel_extractor.get_sheet(sheet_name)
        headers_info = self.header_parser.parse_sheet_headers(sheet_name, header_row=1, max_header_rows=5)
        
        # 스키마에서 가중치 설정 가져오기
        weight_config = self.schema_loader.get_weight_config(sheet_name)
        
        # 지역 이름 열: 스키마에서 가져오거나 동적으로 찾기
        region_column = weight_config.get('region_column')
        if region_column is None:
            region_column = self._find_region_column(sheet, headers_info)
        
        # 산업/업태/품목 이름 열: 스키마에서 가져오거나 동적으로 찾기
        category_column = weight_config.get('category_column')
        if category_column is None:
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
        
        # 캐시에 저장
        self._structure_cache[sheet_name] = structure
        
        return structure
    
    def _find_region_column(self, sheet, headers_info: Dict) -> int:
        """
        지역 이름이 있는 열을 찾습니다.
        여러 전략을 시도하여 최적의 열을 찾습니다.
        """
        # 전략 1: 헤더에서 "시도", "지역" 등의 키워드로 찾기
        region_keywords = ['시도', '지역', 'region', 'area', '시·도', '시 도', '광역시', '도']
        for col_info in headers_info['columns']:
            header_lower = col_info['header'].lower()
            if any(keyword in header_lower for keyword in region_keywords):
                return col_info['index']
        
        # 전략 2: 데이터에서 지역명이 포함된 열 찾기 (처음 10개 열, 5-15행 검사)
        region_names = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', 
                       '세종', '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        best_col = None
        best_count = 0
        
        for col in range(1, min(11, sheet.max_column + 1)):
            region_count = 0
            for row in range(1, min(20, sheet.max_row + 1)):
                cell = sheet.cell(row=row, column=col)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    if any(region in cell_str for region in region_names):
                        region_count += 1
            
            if region_count > best_count:
                best_count = region_count
                best_col = col
        
        if best_col and best_count >= 2:
            return best_col
        
        # 전략 3: 첫 번째 열이 코드이고 두 번째 열이 이름인 패턴 확인
        for row in range(1, min(10, sheet.max_row + 1)):
            cell_a = sheet.cell(row=row, column=1)
            cell_b = sheet.cell(row=row, column=2)
            if cell_a.value and cell_b.value:
                # A열이 숫자이고 B열이 지역명이면 B열 반환
                try:
                    int(str(cell_a.value).strip())
                    if any(region in str(cell_b.value).strip() for region in region_names):
                        return 2
                except (ValueError, TypeError):
                    pass
        
        # 폴백: 동적 탐색 결과 없을 경우 경고 후 2번째 열 반환
        print(f"[WARNING] 지역 열을 동적으로 찾지 못했습니다. 기본값(열 2) 사용")
        return 2
    
    def _find_category_column(self, sheet, headers_info: Dict) -> int:
        """
        산업/업태/품목 이름이 있는 열을 찾습니다.
        여러 전략을 시도하여 최적의 열을 찾습니다.
        """
        # 전략 1: 헤더에서 "산업", "업태", "품목" 등의 키워드로 찾기
        category_keywords = ['산업', '업태', '품목', 'category', 'industry', 'item', '업종', '분류명', '항목']
        for col_info in headers_info['columns']:
            header_lower = col_info['header'].lower()
            if any(keyword in header_lower for keyword in category_keywords):
                return col_info['index']
        
        # 전략 2: 데이터에서 일반적인 카테고리 값이 포함된 열 찾기
        category_values = ['총지수', '계', '합계', '반도체', '전자', '서비스', '제조업', '백화점', '마트']
        
        best_col = None
        best_count = 0
        
        for col in range(3, min(12, sheet.max_column + 1)):  # 3~11열 검사
            cat_count = 0
            for row in range(2, min(30, sheet.max_row + 1)):
                cell = sheet.cell(row=row, column=col)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    if any(cat in cell_str for cat in category_values):
                        cat_count += 1
            
            if cat_count > best_count:
                best_count = cat_count
                best_col = col
        
        if best_col and best_count >= 2:
            return best_col
        
        # 전략 3: 텍스트가 많은 열 찾기 (숫자 열 제외)
        for col in range(4, min(10, sheet.max_column + 1)):
            text_count = 0
            for row in range(4, min(20, sheet.max_row + 1)):
                cell = sheet.cell(row=row, column=col)
                if cell.value:
                    cell_str = str(cell.value).strip()
                    # 숫자가 아닌 텍스트인지 확인
                    try:
                        float(cell_str.replace(',', ''))
                    except ValueError:
                        if len(cell_str) >= 2:  # 2자 이상 텍스트
                            text_count += 1
            
            if text_count >= 5:
                return col
        
        # 폴백: 동적 탐색 결과 없을 경우 경고 후 6번째 열 반환
        print(f"[WARNING] 카테고리 열을 동적으로 찾지 못했습니다. 기본값(열 6) 사용")
        return 6
    
    def _find_classification_column(self, sheet, headers_info: Dict) -> Optional[int]:
        """
        분류 단계가 있는 열을 찾습니다.
        여러 전략을 시도하여 최적의 열을 찾습니다.
        """
        # 전략 1: 헤더에서 "분류", "단계", "level" 등의 키워드로 찾기
        class_keywords = ['분류', '단계', 'level', 'classification', '레벨', '차수']
        for col_info in headers_info['columns']:
            header_lower = col_info['header'].lower()
            if any(keyword in header_lower for keyword in class_keywords):
                return col_info['index']
        
        # 전략 2: 0, 1, 2 같은 작은 정수만 있는 열 찾기 (분류 단계 특성)
        for col in range(2, min(8, sheet.max_column + 1)):
            valid_count = 0
            invalid_count = 0
            
            for row in range(4, min(30, sheet.max_row + 1)):
                cell = sheet.cell(row=row, column=col)
                if cell.value is not None:
                    try:
                        val = float(cell.value)
                        if val in [0, 1, 2, 3]:  # 분류 단계는 보통 0-3
                            valid_count += 1
                        else:
                            invalid_count += 1
                    except (ValueError, TypeError):
                        invalid_count += 1
            
            # 대부분이 0-3 값이면 분류 열로 판단
            if valid_count >= 5 and valid_count > invalid_count * 2:
                return col
        
        # 전략 3: None 반환 (분류 열이 없는 시트일 수 있음)
        return None
    
    def _find_weight_column(self, sheet, headers_info: Dict) -> Optional[int]:
        """
        가중치 열을 찾습니다.
        여러 전략을 시도하여 최적의 열을 찾습니다.
        """
        # 전략 1: 헤더에서 "가중치", "weight" 등의 키워드로 찾기
        weight_keywords = ['가중치', 'weight', '비중', '구성비']
        for col_info in headers_info['columns']:
            header_lower = col_info['header'].lower()
            if any(keyword in header_lower for keyword in weight_keywords):
                return col_info['index']
        
        # 전략 2: 0-1000 범위의 소수점 숫자가 있는 열 찾기 (가중치 특성)
        for col in range(3, min(8, sheet.max_column + 1)):
            weight_like_count = 0
            total_count = 0
            
            for row in range(4, min(30, sheet.max_row + 1)):
                cell = sheet.cell(row=row, column=col)
                if cell.value is not None:
                    try:
                        val = float(cell.value)
                        total_count += 1
                        # 가중치는 보통 0-1000 범위
                        if 0 <= val <= 1000:
                            weight_like_count += 1
                    except (ValueError, TypeError):
                        pass
            
            # 대부분이 가중치 범위 내 값이면 가중치 열로 판단
            if weight_like_count >= 5 and weight_like_count > total_count * 0.7:
                # 분류 열과 구분: 분류 열은 0-3만 있음
                has_large_values = False
                for row in range(4, min(30, sheet.max_row + 1)):
                    cell = sheet.cell(row=row, column=col)
                    if cell.value is not None:
                        try:
                            val = float(cell.value)
                            if val > 10:
                                has_large_values = True
                                break
                        except (ValueError, TypeError):
                            pass
                
                if has_large_values:
                    return col
        
        # 전략 3: None 반환 (가중치 열이 없는 시트일 수 있음)
        return None
    
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
            else:
                # primary_header에서 못 찾은 경우, full_header에서도 시도
                # (예: Row 2에 "(2020=100)", Row 3에 "2025 2/4p"인 경우)
                full_header = col_info.get('full_header', '')
                if full_header:
                    period = self._parse_period_from_header(full_header)
                    if period:
                        year, quarter = period
                        # 이미 해당 분기가 없는 경우에만 추가
                        if (year, quarter) not in quarter_columns:
                            quarter_columns[(year, quarter)] = col_info['index']
        
        return quarter_columns
    
    def _parse_period_from_header(self, header: str) -> Optional[Tuple[int, int]]:
        """
        헤더에서 연도와 분기를 파싱합니다.
        다양한 형식의 분기 표기를 지원합니다.
        """
        # 지원하는 형식들:
        # - "2023 3/4", "2025 2/4p", "2025 2/4P" (한국 표준)
        # - "2024.Q1", "2024 Q2" (Q 표기)
        # - "2024-1Q", "2024 1Q" (분기 뒤 Q)
        # - "2024년 1분기", "2024년1분기" (한글)
        # - "'23 3/4", "'24 1Q" (년도 축약)
        # - "1/4분기 2024" (순서 뒤집힘)
        
        patterns = [
            # 한국 표준 형식
            (r'(\d{4})\s*(\d)/4[pP]?', lambda m: (int(m.group(1)), int(m.group(2)))),
            # Q 표기 형식
            (r'(\d{4})[\s\.]*[Qq](\d)', lambda m: (int(m.group(1)), int(m.group(2)))),
            # 분기 뒤 Q
            (r'(\d{4})[\s\-]*(\d)[Qq]', lambda m: (int(m.group(1)), int(m.group(2)))),
            # 한글 형식
            (r'(\d{4})년?\s*(\d)\s*분기', lambda m: (int(m.group(1)), int(m.group(2)))),
            # 년도 축약 (2자리)
            (r"'(\d{2})\s*(\d)/4[pP]?", lambda m: (2000 + int(m.group(1)), int(m.group(2)))),
            (r"'(\d{2})\s*[Qq](\d)", lambda m: (2000 + int(m.group(1)), int(m.group(2)))),
            (r"'(\d{2})\s*(\d)[Qq]", lambda m: (2000 + int(m.group(1)), int(m.group(2)))),
            # 순서 뒤집힌 형식
            (r'(\d)/4\s*분기?\s*(\d{4})', lambda m: (int(m.group(2)), int(m.group(1)))),
            (r'[Qq](\d)\s*(\d{4})', lambda m: (int(m.group(2)), int(m.group(1)))),
            # 반기 형식 (1반기=1Q+2Q, 2반기=3Q+4Q) - 참고용
            # 영문 형식
            (r'(\d{4})\s*(\d)(?:st|nd|rd|th)?\s*[Qq](?:uarter)?', lambda m: (int(m.group(1)), int(m.group(2)))),
        ]
        
        for pattern, extractor in patterns:
            match = re.search(pattern, header)
            if match:
                try:
                    year, quarter = extractor(match)
                    if 2000 <= year <= 2100 and 1 <= quarter <= 4:
                        return (year, quarter)
                except (ValueError, TypeError, IndexError):
                    continue
        
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
        
        성능 최적화: 결과를 캐시하여 반복 호출 시 재검색하지 않습니다.
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름
            category_filter: 카테고리 필터 (예: '총지수', '계') - None이면 모든 행 허용
            
        Returns:
            행 번호 또는 None
        """
        # 캐시 키 생성
        cache_key = f"{sheet_name}_{region_name}_{category_filter or 'none'}"
        if cache_key in self._region_row_cache:
            return self._region_row_cache[cache_key]
        
        structure = self.parse_sheet_structure(sheet_name)
        sheet = self.excel_extractor.get_sheet(sheet_name)
        region_col = structure['region_column']
        data_start_row = structure['data_start_row']
        category_col = structure.get('category_column')
        classification_col = structure.get('classification_column')
        
        # 지역 이름으로 행 찾기
        # 공백을 무시하고 비교하기 위한 정규화 함수
        def normalize_name(name: str) -> str:
            return name.replace(' ', '').replace('　', '')  # 일반 공백과 전각 공백 제거
        
        normalized_region = normalize_name(region_name)
        
        result = None
        for row in range(data_start_row, min(data_start_row + 500, sheet.max_row + 1)):
            cell = sheet.cell(row=row, column=region_col)
            if not cell.value:
                continue
            
            cell_str = str(cell.value).strip()
            cell_normalized = normalize_name(cell_str)
            
            # 정확한 매칭 또는 정규화된 이름 비교
            if cell_str != region_name and cell_normalized != normalized_region:
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
            
            # 수출/수입/물가 시트의 경우: 카테고리나 분류 단계 체크 완화
            if sheet_name in ['수출', '수입'] or '물가' in sheet_name:
                # 카테고리 열이 비어있거나 None이면 OK
                if category_col and not category_filter:
                    cat_cell = sheet.cell(row=row, column=category_col)
                    if cat_cell.value:
                        cat_str = str(cat_cell.value).strip()
                        # "총지수"나 "계"가 아니면 스킵할 수도 있지만, 일단 허용
                        if cat_str and cat_str not in ['총지수', '계', '   계']:
                            # 품목별 데이터일 수 있으므로 계속 확인
                            pass
                result = row
                break
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
                            result = row
                            break
                    else:
                        # 카테고리가 비어있으면 스킵
                        continue
                else:
                    result = row
                    break
        
        # 캐시에 저장 (None도 저장하여 재검색 방지)
        self._region_row_cache[cache_key] = result
        return result
    
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
        
        # 값이 비어있거나 "-"이면 기본값 1 반환
        if cell.value is None or (isinstance(cell.value, str) and (not cell.value.strip() or cell.value.strip() == '-')):
            return 1.0
        
        try:
            return float(cell.value)
        except (ValueError, TypeError):
            # 변환 실패 시 기본값 1 반환
            return 1.0
    
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
        # 캐시 초기화
        self._structure_cache.clear()
        self._region_row_cache.clear()
    
    def clear_cache(self):
        """캐시만 초기화 (리소스는 유지)"""
        self._structure_cache.clear()
        self._region_row_cache.clear()

