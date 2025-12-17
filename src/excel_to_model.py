"""
엑셀 데이터를 데이터 모델로 변환하는 모듈
"""

import re
from typing import Dict, List, Optional, Any, Tuple
from .excel_extractor import ExcelExtractor
from .data_model import SheetData, RegionData, CategoryData, DocumentData
from .period_detector import PeriodDetector


class ExcelToModelConverter:
    """엑셀 데이터를 데이터 모델로 변환"""
    
    # 지역 코드 매핑
    REGION_CODE_MAP = {
        '00': '전국',
        '11': '서울',
        '21': '부산',
        '22': '대구',
        '23': '인천',
        '24': '광주',
        '25': '대전',
        '26': '울산',
        '29': '세종',
        '31': '경기',
        '32': '강원',
        '33': '충북',
        '34': '충남',
        '35': '전북',
        '36': '전남',
        '37': '경북',
        '38': '경남',
        '39': '제주'
    }
    
    def __init__(self, excel_extractor: ExcelExtractor):
        """
        변환기 초기화
        
        Args:
            excel_extractor: ExcelExtractor 인스턴스
        """
        self.excel_extractor = excel_extractor
        self.period_detector = PeriodDetector(excel_extractor)
    
    def convert_sheet(self, sheet_name: str, year: int, quarter: int) -> SheetData:
        """
        시트를 SheetData 객체로 변환
        
        Args:
            sheet_name: 시트명
            year: 연도
            quarter: 분기
            
        Returns:
            SheetData 객체
        """
        sheet = self.excel_extractor.get_sheet(sheet_name)
        
        # 시트 제목 찾기 (보통 1행)
        title = ""
        if sheet.cell(row=1, column=1).value:
            title = str(sheet.cell(row=1, column=1).value).strip()
        
        sheet_data = SheetData(
            sheet_name=sheet_name,
            title=title,
            header_row=3
        )
        
        # 헤더 행에서 연도 컬럼 찾기
        self._parse_year_columns(sheet, sheet_data, year, quarter)
        
        # 데이터 행 파싱
        self._parse_data_rows(sheet, sheet_data, year, quarter)
        
        return sheet_data
    
    def _parse_year_columns(self, sheet, sheet_data: SheetData, year: int, quarter: int):
        """헤더에서 연도/분기 컬럼 찾기"""
        header_row = sheet_data.header_row
        import re
        
        # 헤더 행에서 분기 컬럼 찾기 (50열부터 시작, period_detector와 동일한 로직)
        for col_idx in range(50, min(200, sheet.max_column + 1)):
            cell = sheet.cell(row=header_row, column=col_idx)
            if not cell.value:
                continue
            
            header_text = str(cell.value).strip()
            
            # 분기 문자열 파싱 (period_detector와 동일한 로직)
            patterns = [
                r'(\d{4})\s+(\d)/4[pP]?',  # "2023 3/4", "2025 2/4p"
                r'(\d{4})\s*(\d)/4[pP]?',  # "2023 3/4" (공백 없음)
                r'(\d{4})\.\s*(\d)',  # "2024. 1"
            ]
            
            period = None
            for pattern in patterns:
                match = re.search(pattern, header_text)
                if match:
                    found_year = int(match.group(1))
                    found_quarter = int(match.group(2))
                    if 2000 <= found_year <= 2100 and 1 <= found_quarter <= 4:
                        period = (found_year, found_quarter)
                        break
            
            if period:
                found_year, found_quarter = period
                sheet_data.quarter_columns[f"{found_year}_{found_quarter}"] = col_idx
                
                # 연도 컬럼도 저장
                if found_year not in sheet_data.year_columns:
                    sheet_data.year_columns[found_year] = col_idx
    
    def _find_quarter_column(self, sheet, header_row: int, start_col: int, 
                             year: int, quarter: int) -> Optional[int]:
        """특정 연도/분기의 컬럼 찾기"""
        # 같은 행에서 연도 다음 컬럼들 확인
        for offset in range(5):  # 최대 5개 컬럼 확인
            col_idx = start_col + offset
            if col_idx > sheet.max_column:
                break
            
            cell = sheet.cell(row=header_row, column=col_idx)
            if not cell.value:
                continue
            
            header_text = str(cell.value).strip()
            # 분기 패턴 확인
            if re.search(rf'{year}\s*[\./]?\s*{quarter}\s*[/분]?', header_text):
                return col_idx
        
        return None
    
    def _parse_data_rows(self, sheet, sheet_data: SheetData, year: int, quarter: int):
        """데이터 행 파싱"""
        header_row = sheet_data.header_row
        period_key = f"{year}_{quarter}"
        
        # 카테고리 컬럼 번호
        category_col = self._get_category_column(sheet_data.sheet_name)
        
        # 데이터 시작 행 (헤더 다음 행)
        for row_idx in range(header_row + 1, min(sheet.max_row + 1, 1000)):  # 최대 1000행까지
            row = sheet[row_idx]
            
            # 지역 코드 (1열, 인덱스 0)
            if len(row) < 2:
                continue
            
            region_code_cell = row[0]
            if not region_code_cell or not region_code_cell.value:
                continue
            
            region_code = str(region_code_cell.value).strip()
            
            # 지역 이름 (2열, 인덱스 1)
            region_name_cell = row[1]
            if not region_name_cell or not region_name_cell.value:
                continue
            
            region_name = str(region_name_cell.value).strip()
            
            # 빈 행 스킵
            if not region_code or not region_name or region_name in ['지역이름', '지역 이름']:
                continue
            
            # RegionData 가져오기 또는 생성
            if region_code not in sheet_data.regions:
                sheet_data.regions[region_code] = RegionData(
                    region_code=region_code,
                    region_name=region_name
                )
            
            region_data = sheet_data.regions[region_code]
            
            # 현재 분기 값 찾기
            if period_key in sheet_data.quarter_columns:
                col_idx = sheet_data.quarter_columns[period_key]
                if col_idx <= len(row):
                    value_cell = row[col_idx - 1]  # 0-based 인덱스
                    if value_cell.value is not None:
                        try:
                            value = float(value_cell.value)
                            region_data.values[period_key] = value
                        except (ValueError, TypeError):
                            pass
            
            # 모든 분기 값 저장
            for period, col_idx in sheet_data.quarter_columns.items():
                if col_idx <= len(row):
                    value_cell = row[col_idx - 1]
                    if value_cell.value is not None:
                        try:
                            value = float(value_cell.value)
                            region_data.values[period] = value
                        except (ValueError, TypeError):
                            pass
            
            # 카테고리 정보 (산업, 업태 등)
            if category_col and category_col <= len(row):
                category_cell = row[category_col - 1]
                if category_cell and category_cell.value:
                    category_name = str(category_cell.value).strip()
                    # 헤더 텍스트 제외
                    if category_name and category_name not in ['산업 이름', '업태 종류', '공정 이름', '산업코드', '상품 이름']:
                        # 카테고리별 값 저장
                        for period, col_idx in sheet_data.quarter_columns.items():
                            if col_idx <= len(row):
                                value_cell = row[col_idx - 1]
                                if value_cell.value is not None:
                                    try:
                                        value = float(value_cell.value)
                                        if category_name not in region_data.categories:
                                            region_data.categories[category_name] = {}
                                        region_data.categories[category_name][period] = value
                                    except (ValueError, TypeError):
                                        pass
    
    def _get_category_column(self, sheet_name: str) -> Optional[int]:
        """시트별 카테고리 컬럼 번호 반환"""
        # 일반적으로 6열 (F열) 또는 5열 (E열)
        if '소비' in sheet_name or '소매' in sheet_name:
            return 5  # E열
        elif '건설' in sheet_name:
            return 5  # E열
        else:
            return 6  # F열
    
    def convert_document(self, year: int, quarter: int, 
                        sheet_names: Optional[List[str]] = None) -> DocumentData:
        """
        전체 문서를 DocumentData로 변환
        
        Args:
            year: 연도
            quarter: 분기
            sheet_names: 변환할 시트 목록 (None이면 모든 시트)
            
        Returns:
            DocumentData 객체
        """
        if sheet_names is None:
            sheet_names = self.excel_extractor.get_sheet_names()
            # '완료체크' 시트 제외
            sheet_names = [s for s in sheet_names if s != '완료체크']
        
        document_data = DocumentData(year=year, quarter=quarter)
        
        for sheet_name in sheet_names:
            try:
                sheet_data = self.convert_sheet(sheet_name, year, quarter)
                document_data.sheets[sheet_name] = sheet_data
            except Exception as e:
                print(f"시트 '{sheet_name}' 변환 중 오류: {e}")
                continue
        
        return document_data

