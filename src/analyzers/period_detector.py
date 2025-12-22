"""
연도 및 분기 자동 감지 모듈
엑셀 파일에서 사용 가능한 연도와 분기 범위를 자동으로 감지
"""

from typing import List, Tuple, Optional, Dict
from ..core.excel_extractor import ExcelExtractor


class PeriodDetector:
    """엑셀 파일에서 연도와 분기 정보를 자동으로 감지하는 클래스"""
    
    def __init__(self, excel_extractor: ExcelExtractor):
        """
        PeriodDetector 초기화
        
        Args:
            excel_extractor: 엑셀 추출기 인스턴스
        """
        self.excel_extractor = excel_extractor
    
    def detect_available_periods(self, sheet_name: str) -> Dict[str, any]:
        """
        시트에서 사용 가능한 연도와 분기 범위를 감지합니다.
        
        Args:
            sheet_name: 시트 이름
            
        Returns:
            {
                'min_year': 최소 연도,
                'max_year': 최대 연도,
                'min_quarter': 최소 분기,
                'max_quarter': 최대 분기,
                'available_periods': [(year, quarter), ...],
                'default_year': 기본 연도,
                'default_quarter': 기본 분기
            }
        """
        sheet = self.excel_extractor.get_sheet(sheet_name)
        
        # 헤더 행에서 분기 정보 찾기 (보통 3행에 있음)
        periods = []
        preliminary_period = None
        
        # Column 50부터 100까지 확인 (분기 데이터가 있는 범위)
        for col in range(50, min(150, sheet.max_column + 1)):
            # Row 3에 분기 정보가 있을 수 있음
            cell_value = sheet.cell(row=3, column=col).value
            
            if cell_value:
                cell_str = str(cell_value).strip()
                # "2023 3/4", "2024 1/4", "2025 2/4p" 등의 형식 파싱
                period = self._parse_period_string(cell_str)
                if period:
                    periods.append(period)
                    # 'p' 또는 'P'가 있으면 예비 데이터로 표시
                    if 'p' in cell_str.lower():
                        preliminary_period = period
        
        if not periods:
            # 기본값 반환
            result = {
                'min_year': 2023,
                'max_year': 2025,
                'min_quarter': 1,
                'max_quarter': 4,
                'available_periods': [(2023, 1), (2023, 2), (2023, 3), (2023, 4),
                                     (2024, 1), (2024, 2), (2024, 3), (2024, 4),
                                     (2025, 1), (2025, 2)],
                'default_year': 2025,
                'default_quarter': 2
            }
        else:
            # 감지된 기간에서 최소/최대 추출
            years = [p[0] for p in periods]
            quarters = [p[1] for p in periods]
            
            min_year = min(years)
            max_year = max(years)
            
            # 최신 연도의 최신 분기를 기본값으로
            latest_periods = [p for p in periods if p[0] == max_year]
            if latest_periods:
                default_quarter = max([p[1] for p in latest_periods])
            else:
                default_quarter = max(quarters) if quarters else 2
            
            result = {
                'min_year': min_year,
                'max_year': max_year,
                'min_quarter': min(quarters),
                'max_quarter': max(quarters),
                'available_periods': sorted(set(periods)),
                'preliminary_period': preliminary_period,  # 예비 데이터 분기
                'default_year': max_year,
                'default_quarter': default_quarter
            }
        
        return result
    
    def _parse_period_string(self, period_str: str) -> Optional[Tuple[int, int]]:
        """
        분기 문자열을 파싱합니다.
        다양한 형식의 분기 표기를 지원합니다.
        
        Args:
            period_str: 다양한 형식의 분기 문자열
            
        Returns:
            (연도, 분기) 튜플 또는 None
        """
        import re
        
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
            # 영문 형식
            (r'(\d{4})\s*(\d)(?:st|nd|rd|th)?\s*[Qq](?:uarter)?', lambda m: (int(m.group(1)), int(m.group(2)))),
        ]
        
        for pattern, extractor in patterns:
            match = re.search(pattern, period_str)
            if match:
                try:
                    year, quarter = extractor(match)
                    if 2000 <= year <= 2100 and 1 <= quarter <= 4:
                        return (year, quarter)
                except (ValueError, TypeError, IndexError):
                    continue
        
        return None
    
    def get_quarter_headers(self, sheet_name: str, start_year: int = None, 
                            start_quarter: int = None, count: int = 8) -> List[str]:
        """
        분기 헤더 리스트를 생성합니다.
        
        Args:
            sheet_name: 시트 이름
            start_year: 시작 연도 (None이면 최신 연도에서 역순)
            start_quarter: 시작 분기 (None이면 최신 분기)
            count: 생성할 헤더 개수
            
        Returns:
            분기 헤더 문자열 리스트 (예: ['2023 3/4', '2023 4/4', ...])
        """
        periods_info = self.detect_available_periods(sheet_name)
        
        if start_year is None:
            start_year = periods_info['default_year']
        if start_quarter is None:
            start_quarter = periods_info['default_quarter']
        
        headers = []
        current_year = start_year
        current_quarter = start_quarter
        
        # 역순으로 헤더 생성 (최신 분기부터 과거로)
        for i in range(count):
            # P 표시 (예비)는 첫 번째(최신) 분기에만
            is_latest = (i == 0 and current_year == periods_info['max_year'] and 
                        current_quarter == periods_info['max_quarter'])
            suffix = 'P' if is_latest else ''
            
            header = f"{current_year} {current_quarter}/4{suffix}"
            headers.append(header)
            
            # 이전 분기로 이동
            current_quarter -= 1
            if current_quarter < 1:
                current_quarter = 4
                current_year -= 1
        
        reversed_headers = list(reversed(headers))  # 과거부터 최신 순서로 정렬
        
        # 마지막 헤더(최신 분기)에 P 표시 추가
        preliminary_period = periods_info.get('preliminary_period')
        if reversed_headers:
            last_header = reversed_headers[-1]
            # 예비 데이터 분기이거나 최신 분기인 경우 P 표시
            is_preliminary = (preliminary_period and 
                            start_year == preliminary_period[0] and 
                            start_quarter == preliminary_period[1])
            is_latest = (start_year == periods_info['max_year'] and 
                        start_quarter == periods_info['max_quarter'])
            
            if (is_preliminary or is_latest) and not last_header.endswith('P'):
                reversed_headers[-1] = last_header + 'P'
        
        return reversed_headers
    
    def validate_period(self, sheet_name: str, year: int, quarter: int) -> Tuple[bool, Optional[str]]:
        """
        연도와 분기가 유효한지 검증합니다.
        
        Args:
            sheet_name: 시트 이름
            year: 연도
            quarter: 분기
            
        Returns:
            (유효 여부, 에러 메시지) 튜플
        """
        periods_info = self.detect_available_periods(sheet_name)
        
        if year < periods_info['min_year'] or year > periods_info['max_year']:
            return False, f"연도는 {periods_info['min_year']}년부터 {periods_info['max_year']}년까지 가능합니다."
        
        if quarter < 1 or quarter > 4:
            return False, "분기는 1부터 4까지 가능합니다."
        
        if (year, quarter) not in periods_info['available_periods']:
            return False, f"{year}년 {quarter}분기 데이터가 엑셀 파일에 없습니다."
        
        return True, None

