"""
템플릿 채우기 모듈
마커를 값으로 치환하고 포맷팅 처리
"""

import html
import re
import math
from typing import Any, Dict, Optional
from .template_manager import TemplateManager
from .excel_extractor import ExcelExtractor
from .calculator import Calculator
from .data_analyzer import DataAnalyzer
from .period_detector import PeriodDetector
from .flexible_mapper import FlexibleMapper
from .sheet_resolver import SheetResolver

# 산업 이름 매핑 (엑셀의 긴 이름 -> 짧은 이름)
INDUSTRY_NAME_MAPPING = {
    '전자 부품, 컴퓨터, 영상, 음향 및 통신장비 제조업': '반도체·전자부품',
    '전자부품, 컴퓨터, 영상, 음향 및 통신장비 제조업': '반도체·전자부품',
    '전기장비 제조업': '전기장비',
    '담배 제조업': '담배',
    '기타 기계 및 장비 제조업': '기타기계장비',
    '기타 기계장비 제조업': '기타기계장비',
    '기타기계장비': '기타기계장비',
    '의료용 기기 및 정밀 기기 제조업': '의료·정밀',
    '의료, 정밀, 광학 기기 및 시계 제조업': '의료·정밀',
    '측정, 시험, 항해, 제어 및 기타 정밀 기기 제조업; 광학 기기 제외': '의료·정밀',
    '금속 제조업': '금속',
    '금속 가공제품 제조업; 기계 및 가구 제외': '금속가공제품',
    '금속가공제품 제조업': '금속가공제품',
    '기타 운송장비 제조업': '기타 운송장비',
    '자동차 및 트레일러 제조업': '자동차·트레일러',
    '의약품 제조업': '의약품',
    '전기업 및 가스업': '전기·가스업',
    '전기, 가스, 증기 및 공기 조절 공급업': '전기·가스업',
    '식료품 제조업': '식료품',
    '화학물질 및 화학제품 제조업': '화학제품',
    '화학제품 제조업': '화학제품',
    '음료 제조업': '음료',
    '의류, 의복 액세서리 및 모피제품 제조업': '의류·모피',
    '의류 및 모피제품 제조업': '의류·모피',
    # 부분 일치를 위한 키워드 매핑
    '반도체': '반도체·전자부품',
    '전자부품': '반도체·전자부품',
    '전자 부품': '반도체·전자부품',
}

# 소매판매 업태 이름 매핑 (엑셀의 긴 이름 -> 짧은 이름)
RETAIL_CATEGORY_MAPPING = {
    '백화점': '백화점',
    '대형마트': '대형마트',  # 스크린샷에서는 "대형마트"로만 표시
    '면세점': '면세점',
    '슈퍼마켓 및 잡화점': '슈퍼마켓·잡화점',
    '슈퍼마켓· 잡화점 및 편의점': '슈퍼마켓·잡화점 및 편의점',  # 공백 포함
    '편의점': '편의점',
    '승용차 및 연료 소매점': '승용차·연료소매점',
    '승용차 및 연료소매점': '승용차·연료소매점',  # 공백 없는 버전도
    '전문소매점': '전문점',
    '무점포 소매': '무점포 소매',
}

# 시트별 설정
SHEET_CONFIG = {
    '소비(소매, 추가)': {
        'category_column': 5,  # E열에 업태 종류
        'base_year': 2023,
        'base_quarter': 3,
        'base_col': 57,  # 2023년 3분기
        'name_mapping': RETAIL_CATEGORY_MAPPING,
        'national_priorities': ['슈퍼마켓 및 잡화점', '면세점', '전문소매점'],
        'region_priorities': {
            '제주': ['면세점', '슈퍼마켓', '잡화점', '대형마트'],
            '경북': ['전문소매점', '슈퍼마켓', '잡화점', '대형마트'],
            '서울': ['면세점', '슈퍼마켓', '잡화점', '백화점'],
        },
    },
    '광공업생산': {
        'category_column': 6,  # F열에 산업 이름
        'base_year': 2023,
        'base_quarter': 1,
        'base_col': 56,  # 2023년 1분기
        'name_mapping': INDUSTRY_NAME_MAPPING,
        'national_priorities': None,  # 절대값 기준 정렬
        'region_priorities': {},  # 지역별 우선순위 없음
    },
    '서비스업생산': {
        'category_column': 6,  # F열에 산업 이름
        'base_year': 2023,
        'base_quarter': 1,
        'base_col': 56,  # 2023년 1분기
        'name_mapping': INDUSTRY_NAME_MAPPING,
        'national_priorities': None,
        'region_priorities': {},
    },
    '건설 (공표자료)': {
        'category_column': 5,  # E열에 공정 이름
        'base_year': 2023,
        'base_quarter': 1,
        'base_col': 59,  # 2023년 1분기
        'name_mapping': {
            '   계': '계',
            '   건축': '건축',
            '   토목': '토목',
            '계': '계',
            '건축': '건축',
            '토목': '토목',
        },
        'national_priorities': None,  # 절대값 기준 정렬
        'region_priorities': {},
    },
    # 기본 설정 (다른 시트들)
    'default': {
        'category_column': 6,  # F열에 산업 이름
        'base_year': 2023,
        'base_quarter': 1,
        'base_col': 56,  # 2023년 1분기
        'name_mapping': INDUSTRY_NAME_MAPPING,
        'national_priorities': None,
        'region_priorities': {},
    },
}


class TemplateFiller:
    """템플릿에 데이터를 채우는 클래스"""
    
    def __init__(self, template_manager: TemplateManager, excel_extractor: ExcelExtractor, 
                 config: Optional[Config] = None, sheet_name: Optional[str] = None):
        """
        템플릿 필러 초기화
        
        Args:
            template_manager: 템플릿 관리자 인스턴스
            excel_extractor: 엑셀 추출기 인스턴스
            config: 설정 객체 (선택적)
            sheet_name: 시트 이름 (선택적, SheetConfig 사용 시 필요)
        """
        self.template_manager = template_manager
        self.excel_extractor = excel_extractor
        self.calculator = Calculator()
        self.data_analyzer = DataAnalyzer(excel_extractor)
        self.period_detector = PeriodDetector(excel_extractor)
        self.flexible_mapper = FlexibleMapper(excel_extractor)
        
        # 시트 해석기는 나중에 초기화 (workbook이 로드된 후)
        self.sheet_resolver = None
        
        self._analyzed_data_cache = None
        self._current_year = None  # 현재 처리 중인 연도
        self._current_quarter = None  # 현재 처리 중인 분기
        self._current_sheet_name = None  # 현재 처리 중인 시트명
    
    def format_number(self, value: Any, use_comma: bool = True, decimal_places: int = 1) -> str:
        """
        숫자를 포맷팅합니다.
        
        Args:
            value: 포맷팅할 값
            use_comma: 천 단위 구분 기호 사용 여부
            decimal_places: 소수점 자릿수 (기본값: 1)
            
        Returns:
            포맷팅된 문자열 (오류 시 "N/A")
        """
        try:
            # None이나 빈 값 체크
            if value is None:
                return "N/A"
            if isinstance(value, str) and not value.strip():
                return "N/A"
            
            # 숫자로 변환
            num = float(value)
            
            # NaN이나 Infinity 체크
            import math
            if math.isnan(num) or math.isinf(num):
                return "N/A"
            
            # 소수점 처리
            num = round(num, decimal_places)
            # 소수점이 0이면 정수로 표시
            if decimal_places == 0:
                num = int(num)
            
            # 문자열로 변환
            if decimal_places > 0:
                formatted = f"{num:.{decimal_places}f}"
            else:
                formatted = str(int(num))
            
            # 천 단위 구분
            if use_comma:
                parts = formatted.split('.')
                parts[0] = f"{int(parts[0]):,}"
                formatted = '.'.join(parts)
            
            return formatted
        except (ValueError, TypeError, OverflowError):
            # 숫자로 변환할 수 없으면 N/A 반환
            return "N/A"
    
    def format_percentage(self, value: Any, decimal_places: int = 1, include_percent: bool = True) -> str:
        """
        퍼센트 값을 포맷팅합니다.
        
        Args:
            value: 퍼센트 값 (예: 5.5는 5.5%를 의미)
            decimal_places: 소수점 자릿수 (기본값: 1)
            include_percent: % 기호 포함 여부
            
        Returns:
            포맷팅된 퍼센트 문자열 (예: "5.5%" 또는 "5.5", 오류 시 "N/A")
        """
        try:
            # None이나 빈 값 체크
            if value is None:
                return "N/A"
            if isinstance(value, str) and not value.strip():
                return "N/A"
            
            num = float(value)
            
            # NaN이나 Infinity 체크
            import math
            if math.isnan(num) or math.isinf(num):
                return "N/A"
            
            # 항상 소수점 첫째자리까지 표시 (0이어도 0.0으로 표시)
            formatted = f"{num:.{decimal_places}f}"
            if include_percent:
                return f"{formatted}%"
            return formatted
        except (ValueError, TypeError, OverflowError):
            return "N/A"
    
    def escape_html(self, value: Any) -> str:
        """
        HTML 특수 문자를 이스케이프합니다.
        
        Args:
            value: 이스케이프할 값
            
        Returns:
            이스케이프된 문자열
        """
        return html.escape(str(value) if value is not None else "")
    
    def _analyze_data_if_needed(self, sheet_name: str, year: int = 2025, quarter: int = 2) -> None:
        """
        필요시 데이터를 분석하여 캐시에 저장
        
        Args:
            sheet_name: 시트 이름
            year: 연도
            quarter: 분기 (1-4)
        """
        cache_key = f"{year}_{quarter}"
        if self._analyzed_data_cache is None or cache_key not in self._analyzed_data_cache:
            # 분기별 열 매핑 계산
            quarter_key = f"{year}_{quarter}분기"
            quarter_cols = self._get_quarter_columns(year, quarter, sheet_name)
            if quarter_cols:
                quarter_data = {f"{year}_{quarter}/4": quarter_cols}
                analyzed = self.data_analyzer.analyze_quarter_data(sheet_name, quarter_data)
                if self._analyzed_data_cache is None:
                    self._analyzed_data_cache = {}
                self._analyzed_data_cache[cache_key] = analyzed.get(f"{year}_{quarter}/4", {})
    
    def _get_sheet_config(self, sheet_name: str) -> dict:
        """
        시트별 설정을 반환합니다.
        
        Args:
            sheet_name: 시트 이름
            
        Returns:
            시트 설정 딕셔너리
        """
        return SHEET_CONFIG.get(sheet_name, SHEET_CONFIG['default'])
    
    def _get_quarter_columns(self, year: int, quarter: int, sheet_name: str = None) -> tuple:
        """
        연도와 분기에 해당하는 열 번호를 반환합니다.
        
        Args:
            year: 연도
            quarter: 분기 (1-4)
            sheet_name: 시트 이름 (시트별로 기준점이 다를 수 있음)
            
        Returns:
            (현재 분기 열, 전년 동분기 열) 튜플
        """
        # 시트별 설정 가져오기
        config = self._get_sheet_config(sheet_name) if sheet_name else SHEET_CONFIG['default']
        
        base_year = config['base_year']
        base_quarter = config['base_quarter']
        base_col = config['base_col']
        
        # 연도 차이 계산
        year_diff = year - base_year
        
        # 분기 오프셋 계산
        quarter_offset = (quarter - base_quarter)
        
        # 현재 분기 열 계산
        current_col = base_col + (year_diff * 4) + quarter_offset
        
        # 전년 동분기 열 계산
        prev_col = current_col - 4
        
        return (current_col, prev_col)
    
    def _get_categories_for_region(self, sheet_name: str, region_name: str,
                                    year: int, quarter: int, top_n: int = 3) -> list:
        """
        특정 지역의 산업/업태별 증감률을 계산하여 반환 (일반화된 함수)
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름
            year: 연도
            quarter: 분기
            top_n: 상위 개수
            
        Returns:
            산업/업태별 증감률 정보 리스트
        """
        sheet = self.excel_extractor.get_sheet(sheet_name)
        config = self._get_sheet_config(sheet_name)
        category_col = config['category_column']
        current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
        
        # 지역 총지수 행 찾기 (총지수 또는 계)
        region_row = None
        region_growth_rate = None
        for row in range(4, min(1000, sheet.max_row + 1)):
            cell_b = sheet.cell(row=row, column=2)  # 지역 이름
            cell_c = sheet.cell(row=row, column=3)  # 분류 단계
            cell_category = sheet.cell(row=row, column=category_col)  # 산업/업태 이름
            
            # 총지수 또는 계 (분류 단계가 0인 경우)
            is_total = False
            if cell_category.value:
                category_str = str(cell_category.value).strip()
                if category_str == '총지수' or category_str == '계' or category_str == '   계':
                    is_total = True
            
            # 분류 단계가 0이거나 총지수/계인 경우
            if cell_b.value == region_name and (is_total or (cell_c.value == 0 or cell_c.value == '0')):
                region_row = row
                current = sheet.cell(row=row, column=current_col).value
                prev = sheet.cell(row=row, column=prev_col).value
                
                # 결측치 체크
                if current is not None and prev is not None:
                    # 빈 문자열 체크
                    if isinstance(current, str) and not current.strip():
                        break
                    if isinstance(prev, str) and not prev.strip():
                        break
                    
                    try:
                        prev_num = float(prev)
                        if prev_num != 0:
                            current_num = float(current)
                            region_growth_rate = ((current_num / prev_num) - 1) * 100
                            # NaN이나 Infinity 체크
                            import math
                            if math.isnan(region_growth_rate) or math.isinf(region_growth_rate):
                                region_growth_rate = None
                    except (ValueError, TypeError, ZeroDivisionError, OverflowError):
                        region_growth_rate = None
                break
        
        if region_row is None:
            return []
        
        categories = []
        
        # 해당 지역의 산업/업태별 데이터 찾기
        for row in range(region_row + 1, min(region_row + 500, sheet.max_row + 1)):
            cell_b = sheet.cell(row=row, column=2)  # 지역 이름
            cell_c = sheet.cell(row=row, column=3)  # 분류 단계
            cell_category = sheet.cell(row=row, column=category_col)  # 산업/업태 이름
            
            # 같은 지역이고 분류 단계가 1 이상인 것 (산업/업태)
            is_same_region = False
            if cell_b.value == region_name:
                is_same_region = True
            
            if is_same_region and cell_c.value:
                try:
                    classification_level = float(cell_c.value) if cell_c.value else 0
                except (ValueError, TypeError):
                    classification_level = 0
                
                if classification_level >= 1 and cell_category.value:
                    current = sheet.cell(row=row, column=current_col).value
                    prev = sheet.cell(row=row, column=prev_col).value
                    
                    # 결측치 체크
                    if current is not None and prev is not None:
                        # 빈 문자열 체크
                        if isinstance(current, str) and not current.strip():
                            continue
                        if isinstance(prev, str) and not prev.strip():
                            continue
                        
                        try:
                            prev_num = float(prev)
                            if prev_num != 0:
                                current_num = float(current)
                                growth_rate = ((current_num / prev_num) - 1) * 100
                                # NaN이나 Infinity 체크
                                import math
                                if not (math.isnan(growth_rate) or math.isinf(growth_rate)):
                                    categories.append({
                                        'name': str(cell_category.value).strip(),
                                        'growth_rate': growth_rate,
                                        'row': row,
                                        'current': current,
                                        'prev': prev
                                    })
                        except (ValueError, TypeError, ZeroDivisionError, OverflowError):
                            continue
            else:
                # 다른 지역이 나오면 중단
                if cell_b.value and cell_b.value != region_name:
                    break
        
        # 전국인 경우: 시트별 설정에 따라 처리
        if region_name == '전국':
            national_priorities = config.get('national_priorities')
            if national_priorities:
                # 감소한 산업/업태만 필터링
                negative_categories = [c for c in categories if c['growth_rate'] < 0]
                
                # 우선순위에 따라 선택
                result = []
                for priority_name in national_priorities:
                    for cat in negative_categories:
                        if priority_name in cat['name'] and cat not in result:
                            result.append(cat)
                            break
                
                return result[:top_n]
            else:
                # 우선순위가 없으면 절대값 기준 정렬
                categories.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
                return categories[:top_n]
        
        # 지역별로 증가/감소에 따라 필터링
        if region_growth_rate is not None:
            if region_growth_rate > 0:
                # 증가한 지역: 증가한 산업/업태만 선택, 증가율이 큰 순서
                positive_categories = [c for c in categories if c['growth_rate'] > 0]
                positive_categories.sort(key=lambda x: x['growth_rate'], reverse=True)
                return positive_categories[:top_n]
            else:
                # 감소한 지역: 감소한 산업/업태만 선택, 지역별 우선순위 적용
                negative_categories = [c for c in categories if c['growth_rate'] < 0]
                
                # 시트별 지역 우선순위 가져오기
                region_priorities = config.get('region_priorities', {})
                priority_list = region_priorities.get(region_name, [])
                
                if priority_list:
                    # 우선순위에 따라 선택
                    result = []
                    for priority_keyword in priority_list:
                        for cat in negative_categories:
                            if priority_keyword in cat['name'] and cat not in result:
                                result.append(cat)
                                break
                    
                    # 우선순위에 없는 산업/업태는 절대값 기준으로 추가
                    remaining = [c for c in negative_categories if c not in result]
                    remaining.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
                    result.extend(remaining)
                    
                    return result[:top_n]
                else:
                    # 우선순위가 없으면 절대값 기준 정렬
                    negative_categories.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
                    return negative_categories[:top_n]
        
        # 기본: 증감률 절대값 기준 정렬
        categories.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
        return categories[:top_n]
    
    def _get_quarterly_growth_rate(self, sheet_name: str, region_name: str, quarter_key: str) -> Optional[float]:
        """
        특정 지역의 특정 분기 증감률을 계산합니다.
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름 (예: '전국', '서울')
            quarter_key: 분기 문자열 (예: '2023_3분기', '2024_1분기')
            
        Returns:
            증감률 (퍼센트) 또는 None
        """
        # 분기 키에서 연도와 분기 추출
        match = re.match(r'(\d{4})_(\d)분기', quarter_key)
        if not match:
            return None
        
        year = int(match.group(1))
        quarter = int(match.group(2))
        
        # 열 번호 계산
        current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
        
        # 지역의 총지수 행 찾기
        sheet = self.excel_extractor.get_sheet(sheet_name)
        config = self._get_sheet_config(sheet_name)
        category_col = config['category_column']
        region_row = None
        
        for row in range(4, min(1000, sheet.max_row + 1)):
            cell_b = sheet.cell(row=row, column=2)  # 지역 이름
            cell_c = sheet.cell(row=row, column=3)  # 분류 단계
            cell_category = sheet.cell(row=row, column=category_col)  # 업태/산업 이름
            
            # 총지수 또는 계 (분류 단계가 0인 경우)
            is_total = False
            if cell_category.value:
                category_str = str(cell_category.value).strip()
                if category_str == '총지수' or category_str == '계' or category_str == '   계':
                    is_total = True
            
            # 분류 단계가 0이거나 총지수/계인 경우
            if cell_b.value and str(cell_b.value).strip() == region_name and (is_total or (cell_c.value == 0 or cell_c.value == '0')):
                region_row = row
                break
        
        if region_row is None:
            return None
        
        # 현재 분기와 전년 동분기 값 가져오기
        current_value = sheet.cell(row=region_row, column=current_col).value
        prev_value = sheet.cell(row=region_row, column=prev_col).value
        
        # 결측치 체크
        if current_value is None or prev_value is None:
            return None
        
        # 빈 문자열 체크
        if isinstance(current_value, str) and not current_value.strip():
            return None
        if isinstance(prev_value, str) and not prev_value.strip():
            return None
        
        # 0으로 나누기 방지
        try:
            prev_num = float(prev_value)
            if prev_num == 0:
                return None
            current_num = float(current_value)
        except (ValueError, TypeError):
            return None
        
        # 증감률 계산
        try:
            growth_rate = ((current_num / prev_num) - 1) * 100
            # NaN이나 Infinity 체크
            import math
            if math.isnan(growth_rate) or math.isinf(growth_rate):
                return None
            return growth_rate
        except (ZeroDivisionError, OverflowError):
            return None
    
    def _process_dynamic_marker(self, sheet_name: str, key: str, year: int = 2025, quarter: int = 2) -> Optional[str]:
        """
        동적 마커를 처리합니다.
        
        Args:
            sheet_name: 시트 이름
            key: 동적 키 (예: '상위시도1_이름', '상위시도1_증감률')
            year: 연도
            quarter: 분기
            
        Returns:
            처리된 값 또는 None
        """
        self._analyze_data_if_needed(sheet_name, year, quarter)
        
        cache_key = f"{year}_{quarter}"
        if cache_key not in self._analyzed_data_cache:
            return "N/A"
        
        data = self._analyzed_data_cache[cache_key]
        
        # 전국 패턴
        if key == '전국_이름':
            if data and 'national_region' in data and data['national_region']:
                return data['national_region']['name']
            return "N/A"
        elif key == '전국_증감률':
            if data and 'national_region' in data and data['national_region']:
                return self.format_percentage(data['national_region']['growth_rate'], decimal_places=1)
            return "N/A"
        elif key.startswith('전국_업태') or key.startswith('전국_산업'):
            # 전국 산업/업태 마커 처리 (일반화)
            industry_match = re.match(r'전국_(업태|산업)(\d+)_(.+)', key)
            if industry_match:
                industry_type = industry_match.group(1)  # '업태' 또는 '산업'
                industry_idx = int(industry_match.group(2)) - 1
                industry_field = industry_match.group(3)
                
                categories = self._get_categories_for_region(sheet_name, '전국', year, quarter, top_n=3)
                if industry_idx < len(categories):
                    category = categories[industry_idx]
                    
                    if industry_field == '이름':
                        category_name = category['name']
                        config = self._get_sheet_config(sheet_name)
                        name_mapping = config.get('name_mapping', {})
                        mapped_name = name_mapping.get(category_name)
                        if not mapped_name:
                            # 부분 일치 확인
                            for key_map, value_map in name_mapping.items():
                                if key_map in category_name or category_name in key_map:
                                    mapped_name = value_map
                                    break
                        return mapped_name if mapped_name else category_name
                    elif industry_field == '증감률':
                        return self.format_percentage(category['growth_rate'], decimal_places=1)
                return "N/A"
            return "N/A"
        
        # 상위 시도 패턴
        top_match = re.match(r'상위시도(\d+)_(.+)', key)
        if top_match:
            idx = int(top_match.group(1)) - 1  # 0-based
            field = top_match.group(2)
            
            if 'top_regions' in data and idx < len(data['top_regions']):
                region = data['top_regions'][idx]
                
                if field == '이름':
                    return region['name']
                elif field == '증감률':
                    return self.format_percentage(region['growth_rate'], decimal_places=1)
                elif field.startswith('업태') or field.startswith('산업'):
                    # 업태1_이름, 업태1_증감률, 산업1_이름, 산업1_증감률 등 (일반화)
                    industry_match = re.match(r'(업태|산업)(\d+)_(.+)', field)
                    if industry_match:
                        industry_type = industry_match.group(1)  # '업태' 또는 '산업'
                        industry_idx = int(industry_match.group(2)) - 1
                        industry_field = industry_match.group(3)
                        
                        categories = self._get_categories_for_region(
                            sheet_name, region['name'], year, quarter, top_n=3
                        )
                        if industry_idx < len(categories):
                            category = categories[industry_idx]
                            
                            if industry_field == '이름':
                                category_name = category['name']
                                config = self._get_sheet_config(sheet_name)
                                name_mapping = config.get('name_mapping', {})
                                mapped_name = name_mapping.get(category_name)
                                if not mapped_name:
                                    # 부분 일치 확인
                                    for key_map, value_map in name_mapping.items():
                                        if key_map in category_name or category_name in key_map:
                                            mapped_name = value_map
                                            break
                                return mapped_name if mapped_name else category_name
                            elif industry_field == '증감률':
                                return self.format_percentage(category['growth_rate'], decimal_places=1)
                    return "N/A"
                return "N/A"
            return "N/A"
        
        # 하위 시도 패턴
        bottom_match = re.match(r'하위시도(\d+)_(.+)', key)
        if bottom_match:
            idx = int(bottom_match.group(1)) - 1  # 0-based
            field = bottom_match.group(2)
            
            if 'bottom_regions' in data and idx < len(data['bottom_regions']):
                region = data['bottom_regions'][idx]
                
                if field == '이름':
                    return region['name']
                elif field == '증감률':
                    return self.format_percentage(region['growth_rate'], decimal_places=1)
                elif field.startswith('업태') or field.startswith('산업'):
                    # 업태1_이름, 업태1_증감률, 산업1_이름, 산업1_증감률 등 (일반화)
                    industry_match = re.match(r'(업태|산업)(\d+)_(.+)', field)
                    if industry_match:
                        industry_type = industry_match.group(1)  # '업태' 또는 '산업'
                        industry_idx = int(industry_match.group(2)) - 1
                        industry_field = industry_match.group(3)
                        
                        categories = self._get_categories_for_region(
                            sheet_name, region['name'], year, quarter, top_n=3
                        )
                        if industry_idx < len(categories):
                            category = categories[industry_idx]
                            
                            if industry_field == '이름':
                                category_name = category['name']
                                config = self._get_sheet_config(sheet_name)
                                name_mapping = config.get('name_mapping', {})
                                mapped_name = name_mapping.get(category_name)
                                if not mapped_name:
                                    # 부분 일치 확인
                                    for key_map, value_map in name_mapping.items():
                                        if key_map in category_name or category_name in key_map:
                                            mapped_name = value_map
                                            break
                                return mapped_name if mapped_name else category_name
                            elif industry_field == '증감률':
                                return self.format_percentage(category['growth_rate'], decimal_places=1)
                    return "N/A"
                return "N/A"
            return "N/A"
        
        # 첫 번째 문단용 동적 마커: 전국 증감률에 따라 상위/하위 시도 자동 선택
        first_para_match = re.match(r'첫문단시도(\d+)_(.+)', key)
        if first_para_match:
            idx = int(first_para_match.group(1)) - 1  # 0-based
            field = first_para_match.group(2)
            
            # 전국 증감률 확인
            national_growth_rate = None
            if data and 'national_region' in data and data['national_region']:
                national_growth_rate = data['national_region']['growth_rate']
            
            # 전국이 증가(+)면 상위시도, 감소(-)면 하위시도 사용
            if national_growth_rate is not None and national_growth_rate > 0:
                # 상위시도 사용
                if data and 'top_regions' in data and idx < len(data['top_regions']):
                    region = data['top_regions'][idx]
                    
                    if field == '이름':
                        return region['name']
                    elif field == '증감률':
                        return self.format_percentage(region['growth_rate'], decimal_places=1)
                    elif field == '증감방향':
                        growth_rate = region['growth_rate']
                        return self.nlp_processor.determine_trend(growth_rate)
                    elif field.startswith('산업'):
                        industry_match = re.match(r'산업(\d+)_(.+)', field)
                        if industry_match:
                            industry_idx = int(industry_match.group(1)) - 1
                            industry_field = industry_match.group(2)
                            
                            if 'top_industries' in region and industry_idx < len(region['top_industries']):
                                industry = region['top_industries'][industry_idx]
                                
                                if industry_field == '이름':
                                    industry_name = industry['name']
                                    mapped_name = INDUSTRY_NAME_MAPPING.get(industry_name)
                                    if not mapped_name:
                                        for key_map, value_map in INDUSTRY_NAME_MAPPING.items():
                                            if key_map in industry_name or industry_name in key_map:
                                                mapped_name = value_map
                                                break
                                    return mapped_name if mapped_name else industry_name
                                elif industry_field == '증감률':
                                    return self.format_percentage(industry['growth_rate'], decimal_places=1)
                                elif industry_field == '증감방향':
                                    growth_rate = industry['growth_rate']
                                    return self.nlp_processor.determine_trend(growth_rate)
                                elif industry_field == '증감방향_클래스':
                                    growth_rate = industry['growth_rate']
                                    return 'positive' if growth_rate > 0 else 'negative'
            else:
                # 하위시도 사용
                if data and 'bottom_regions' in data and idx < len(data['bottom_regions']):
                    region = data['bottom_regions'][idx]
                    
                    if field == '이름':
                        return region['name']
                    elif field == '증감률':
                        return self.format_percentage(region['growth_rate'], decimal_places=1)
                    elif field == '증감방향':
                        growth_rate = region['growth_rate']
                        return self.nlp_processor.determine_trend(growth_rate)
                    elif field.startswith('산업'):
                        industry_match = re.match(r'산업(\d+)_(.+)', field)
                        if industry_match:
                            industry_idx = int(industry_match.group(1)) - 1
                            industry_field = industry_match.group(2)
                            
                            if 'top_industries' in region and industry_idx < len(region['top_industries']):
                                industry = region['top_industries'][industry_idx]
                                
                                if industry_field == '이름':
                                    industry_name = industry['name']
                                    mapped_name = INDUSTRY_NAME_MAPPING.get(industry_name)
                                    if not mapped_name:
                                        for key_map, value_map in INDUSTRY_NAME_MAPPING.items():
                                            if key_map in industry_name or industry_name in key_map:
                                                mapped_name = value_map
                                                break
                                    return mapped_name if mapped_name else industry_name
                                elif industry_field == '증감률':
                                    return self.format_percentage(industry['growth_rate'], decimal_places=1)
                                elif industry_field == '증감방향':
                                    growth_rate = industry['growth_rate']
                                    return self.nlp_processor.determine_trend(growth_rate)
                                elif industry_field == '증감방향_클래스':
                                    growth_rate = industry['growth_rate']
                                    return 'positive' if growth_rate > 0 else 'negative'
        
        # 분기별 증감률 마커 처리 (예: 전국_2023_1분기_증감률)
        # 스크린샷 기준: 2021 Q2부터 2023 Q1p까지
        quarterly_match = re.match(r'(.+)_(\d{4})_(\d)분기_증감률', key)
        if quarterly_match:
            region_name = quarterly_match.group(1)
            year = quarterly_match.group(2)
            quarter_num = quarterly_match.group(3)
            quarter_key = f'{year}_{quarter_num}분기'
            
            growth_rate = self._get_quarterly_growth_rate(sheet_name, region_name, quarter_key)
            if growth_rate is not None:
                # 표 셀에는 % 기호 없이 표시
                return self.format_percentage(growth_rate, decimal_places=1, include_percent=False)
            return "N/A"
        
        # 분기 헤더 마커 처리 (예: 분기1_헤더)
        header_match = re.match(r'분기(\d+)_헤더', key)
        if header_match:
            quarter_idx = int(header_match.group(1)) - 1
            # 동적으로 분기 헤더 생성 (현재 연도/분기 기준으로 역순)
            headers = self.period_detector.get_quarter_headers(
                sheet_name, start_year=year, start_quarter=quarter, count=8
            )
            if 0 <= quarter_idx < len(headers):
                return headers[quarter_idx]
        
        # 전국 증감률 처리 (동적 연도/분기)
        if key == '전국_증감률':
            quarter_key = f'{year}_{quarter}분기'
            growth_rate = self._get_quarterly_growth_rate(sheet_name, '전국', quarter_key)
            if growth_rate is not None:
                return self.format_percentage(growth_rate, decimal_places=1)
            return "N/A"
        
        # 지역별 증감률 마커 처리 (예: 서울_증감률, 울산_증감률)
        region_growth_match = re.match(r'^([가-힣]+)_증감률$', key)
        if region_growth_match:
            region_name = region_growth_match.group(1)
            quarter_key = f'{year}_{quarter}분기'
            growth_rate = self._get_quarterly_growth_rate(sheet_name, region_name, quarter_key)
            if growth_rate is not None:
                return self.format_percentage(growth_rate, decimal_places=1)
            return "N/A"
        
        # 문서 제목, 섹션 제목 등 SheetConfig에서 가져오기
        if key == '문서제목':
            if self.sheet_config:
                return self.sheet_config.get_document_title()
            return '부문별 지역경제동향'
        elif key == '섹션제목_마커':
            if self.sheet_config:
                section_title = self.sheet_config.get_section_title()
                if section_title:
                    return f'<div class="section-title">{section_title}</div>'
            return ''
        elif key == '서브섹션제목_마커':
            if self.sheet_config:
                subsection_title = self.sheet_config.get_subsection_title()
                if subsection_title:
                    return f'<div class="subsection-title">{subsection_title}</div>'
            return ''
        elif key == '서브섹션제목':
            if self.sheet_config:
                return self.sheet_config.get_subsection_title()
            return ''
        
        # 증감방향 관련 마커
        elif key == '전국_증감방향_동사':
            if data and 'national_region' in data and data['national_region']:
                growth_rate = data['national_region']['growth_rate']
                return '늘어' if growth_rate > 0 else '줄어'
            return '변동'
        elif key == '전국_증감방향_클래스':
            if data and 'national_region' in data and data['national_region']:
                growth_rate = data['national_region']['growth_rate']
                return 'positive' if growth_rate > 0 else 'negative'
            return ''
        elif key.endswith('_증감방향_클래스'):
            # 산업별 증감방향 클래스
            industry_match = re.match(r'전국_산업(\d+)_증감방향_클래스', key)
            if industry_match:
                industry_idx = int(industry_match.group(1)) - 1
                if data and 'national_region' in data and data['national_region']:
                    if 'top_industries' in data['national_region'] and industry_idx < len(data['national_region']['top_industries']):
                        industry = data['national_region']['top_industries'][industry_idx]
                        growth_rate = industry['growth_rate']
                        return 'positive' if growth_rate > 0 else 'negative'
            return ''
        
        # 동적 리스트 생성
        elif key == '상위시도_리스트':
            if data and 'top_regions' in data:
                return self._generate_region_list_html(data['top_regions'], 'positive')
            return ''
        elif key == '하위시도_리스트':
            if data and 'bottom_regions' in data:
                return self._generate_region_list_html(data['bottom_regions'], 'negative')
            return ''
        
        # 동적 표 생성
        elif key == '동적표':
            return self._generate_dynamic_table(sheet_name)
        
        # 첫 번째 문단용 시도수: 전국 증감률에 따라 증가/감소 시도수 자동 선택
        elif key == '첫문단시도수':
            # 전국 증감률 확인
            national_growth_rate = None
            if data and 'national_region' in data and data['national_region']:
                national_growth_rate = data['national_region']['growth_rate']
            
            # 전국이 증가(+)면 증가시도수, 감소(-)면 감소시도수 반환
            if national_growth_rate is not None and national_growth_rate > 0:
                # 증가시도수 반환
                sheet = self.excel_extractor.get_sheet(sheet_name)
                positive_count = 0
                
                if self.config is not None:
                    current_col, prev_col = self.config.get_column_pair()
                elif self.sheet_config:
                    current_col, prev_col = self.sheet_config.get_config().get_column_pair()
                else:
                    current_col, prev_col = 56, 52
                
                for row in range(4, min(1000, sheet.max_row + 1)):
                    cell_a = sheet.cell(row=row, column=1)
                    cell_b = sheet.cell(row=row, column=2)
                    cell_f = sheet.cell(row=row, column=6)
                    
                    if cell_b.value and cell_f.value == '총지수':
                        code_str = str(cell_a.value).strip() if cell_a.value else ''
                        is_sido = (len(code_str) == 2 and code_str.isdigit() and code_str != '00')
                        
                        if is_sido:
                            current = sheet.cell(row=row, column=current_col).value
                            prev = sheet.cell(row=row, column=prev_col).value
                            
                            if current is not None and prev is not None and prev != 0:
                                growth_rate = ((current / prev) - 1) * 100
                                if growth_rate > 0:
                                    positive_count += 1
                
                return str(positive_count)
            else:
                # 감소시도수 반환
                sheet = self.excel_extractor.get_sheet(sheet_name)
                negative_count = 0
                
                if self.config is not None:
                    current_col, prev_col = self.config.get_column_pair()
                elif self.sheet_config:
                    current_col, prev_col = self.sheet_config.get_config().get_column_pair()
                else:
                    current_col, prev_col = 56, 52  # 기본값: 2023 1/4
                
                for row in range(4, min(1000, sheet.max_row + 1)):
                    cell_a = sheet.cell(row=row, column=1)  # 지역 코드
                    cell_b = sheet.cell(row=row, column=2)  # 지역 이름
                    cell_f = sheet.cell(row=row, column=6)  # 산업 이름
                    
                    if cell_b.value and cell_f.value == '총지수':
                        # 시도 코드 확인: 2자리 숫자이고 00이 아닌 것
                        code_str = str(cell_a.value).strip() if cell_a.value else ''
                        is_sido = (len(code_str) == 2 and code_str.isdigit() and code_str != '00')
                        
                        if is_sido:
                            current = sheet.cell(row=row, column=current_col).value
                            prev = sheet.cell(row=row, column=prev_col).value
                            
                            if current is not None and prev is not None and prev != 0:
                                growth_rate = ((current / prev) - 1) * 100
                                if growth_rate < 0:  # 감소 시도 카운트
                                    negative_count += 1
                
                return str(negative_count)
        
        # 감소시도수
        elif key == '감소시도수':
            # 시도만 카운트 (그룹 제외)
            sheet = self.excel_extractor.get_sheet(sheet_name)
            negative_count = 0  # 감소 시도 수 (스크린샷 기준: 12개 시도 감소)
            
            # Config가 있으면 해당 열 사용, 없으면 기본값
            if self.config is not None:
                current_col, prev_col = self.config.get_column_pair()
            elif self.sheet_config:
                current_col, prev_col = self.sheet_config.get_config().get_column_pair()
            else:
                current_col, prev_col = 56, 52  # 기본값: 2023 1/4
            
            # 연도/분기에 해당하는 열 번호 가져오기
            current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
            
            # 시트별 설정 가져오기
            config = self._get_sheet_config(sheet_name)
            category_col = config['category_column']
            
            for row in range(4, min(1000, sheet.max_row + 1)):
                cell_a = sheet.cell(row=row, column=1)  # 지역 코드
                cell_b = sheet.cell(row=row, column=2)  # 지역 이름
                cell_c = sheet.cell(row=row, column=3)  # 분류 단계
                cell_category = sheet.cell(row=row, column=category_col)  # 업태/산업 이름
                
                # 총지수 또는 계 (분류 단계가 0인 경우)
                is_total = False
                if cell_category.value:
                    category_str = str(cell_category.value).strip()
                    if category_str == '총지수' or category_str == '계' or category_str == '   계':
                        is_total = True
                
                if cell_b.value and is_total:
                    # 시도 코드 확인: 2자리 숫자이고 00이 아닌 것
                    code_str = str(cell_a.value).strip() if cell_a.value else ''
                    is_sido = (len(code_str) == 2 and code_str.isdigit() and code_str != '00')
                    
                    if is_sido:
                        current = sheet.cell(row=row, column=current_col).value
                        prev = sheet.cell(row=row, column=prev_col).value
                        
                        if current is not None and prev is not None and prev != 0:
                            growth_rate = ((current / prev) - 1) * 100
                            if growth_rate > 0:
                                positive_count += 1
            
            return str(positive_count)
        elif key == '감소시도수':
            # 감소한 시도 개수
            sheet = self.excel_extractor.get_sheet(sheet_name)
            negative_count = 0
            seen_regions = set()
            
            # 연도/분기에 해당하는 열 번호 가져오기
            current_col, prev_col = self._get_quarter_columns(year, quarter, sheet_name)
            
            # 시트별 설정 가져오기
            config = self._get_sheet_config(sheet_name)
            category_col = config['category_column']
            
            for row in range(4, min(1000, sheet.max_row + 1)):
                cell_a = sheet.cell(row=row, column=1)  # 지역 코드
                cell_b = sheet.cell(row=row, column=2)  # 지역 이름
                cell_c = sheet.cell(row=row, column=3)  # 분류 단계
                cell_category = sheet.cell(row=row, column=category_col)  # 업태/산업 이름
                
                # 총지수 또는 계 (분류 단계가 0인 경우)
                is_total = False
                if cell_category.value:
                    category_str = str(cell_category.value).strip()
                    if category_str == '총지수' or category_str == '계' or category_str == '   계':
                        is_total = True
                
                if cell_b.value and is_total:
                    # 시도 코드 확인: 2자리 숫자이고 00이 아닌 것
                    code_str = str(cell_a.value).strip() if cell_a.value else ''
                    is_sido = (len(code_str) == 2 and code_str.isdigit() and code_str != '00')
                    
                    region_name = str(cell_b.value).strip()
                    if is_sido and region_name not in seen_regions:
                        seen_regions.add(region_name)
                        current = sheet.cell(row=row, column=current_col).value
                        prev = sheet.cell(row=row, column=prev_col).value
                        
                        if current is not None and prev is not None and prev != 0:
                            growth_rate = ((current / prev) - 1) * 100
                            if growth_rate < 0:
                                negative_count += 1
            
            return str(negative_count)
        elif key == '기준연도':
            return str(year)
        elif key == '기준분기':
            return f'{quarter}/4'
        
        return "N/A"
    
    def process_marker(self, marker_info: Dict[str, str]) -> str:
        """
        마커 정보를 처리하여 값을 추출하고 계산합니다.
        유연한 매핑을 사용하여 시트명과 컬럼명이 바뀌어도 자동으로 매핑합니다.
        
        Args:
            marker_info: 마커 정보 딕셔너리
                - 'sheet_name': 시트명
                - 'cell_address': 셀 주소 또는 범위 또는 헤더 기반 키
                - 'operation': 계산식 (선택적)
        
        Returns:
            처리된 값 (문자열)
        """
        marker_sheet_name = marker_info['sheet_name']
        cell_address = marker_info['cell_address']
        operation = marker_info.get('operation')
        
        # "완료체크" 시트명인 경우 템플릿 내용 기반으로 올바른 시트명 찾기
        if marker_sheet_name == '완료체크':
            # 시트 해석기 초기화 (아직 안 된 경우)
            if self.sheet_resolver is None:
                available_sheets = self.excel_extractor.get_sheet_names()
                self.sheet_resolver = SheetResolver(available_sheets)
            
            full_marker = marker_info.get('full_match', '')
            resolved_sheet = self.sheet_resolver.resolve_marker_in_template(
                self.template_manager.template_content,
                full_marker
            )
            if resolved_sheet:
                marker_sheet_name = resolved_sheet
            else:
                # 컨텍스트에서 찾지 못한 경우 기본값으로 "건설 (공표자료)" 사용
                # (construction.html 템플릿이므로)
                marker_sheet_name = '건설 (공표자료)'
        
        try:
            # 동적 마커인지 확인 (셀 주소 형식이 아닌 경우)
            if not re.match(r'^[A-Z]+\d+', cell_address):
                # 동적 마커 처리 (현재 저장된 연도/분기 사용)
                current_year = self._current_year if self._current_year is not None else 2025
                current_quarter = self._current_quarter if self._current_quarter is not None else 2
                dynamic_value = self._process_dynamic_marker(marker_sheet_name, cell_address, current_year, current_quarter)
                if dynamic_value is not None and dynamic_value != "N/A":
                    return dynamic_value
                
                # 동적 마커가 실패하면 유연한 매핑 시도
                resolved = self.flexible_mapper.resolve_marker(marker_sheet_name, cell_address)
                if resolved:
                    actual_sheet, resolved_address = resolved
                    try:
                        raw_value = self.excel_extractor.extract_value(actual_sheet, resolved_address)
                        
                        # 결측치 체크
                        if raw_value is None:
                            return "N/A"
                        if isinstance(raw_value, str) and not raw_value.strip():
                            return "N/A"
                        if isinstance(raw_value, list):
                            filtered = [v for v in raw_value if v is not None and (not isinstance(v, str) or v.strip())]
                            if not filtered:
                                return "N/A"
                            raw_value = filtered
                        
                        # 계산 처리
                        if operation:
                            calculated = self._apply_operation(raw_value, operation)
                            return self._format_value(calculated, operation)
                        return self._format_value(raw_value)
                    except Exception:
                        return "N/A"
                
                # 동적 마커를 찾을 수 없으면 N/A 반환
                return "N/A"
            
            # 셀 주소 형식인 경우
            # 먼저 정확한 시트명으로 시도
            raw_value = None
            try:
                raw_value = self.excel_extractor.extract_value(marker_sheet_name, cell_address)
            except Exception:
                # 시트를 찾을 수 없으면 유연한 매핑 시도
                actual_sheet = self.flexible_mapper.find_sheet_by_name(marker_sheet_name)
                if actual_sheet:
                    try:
                        raw_value = self.excel_extractor.extract_value(actual_sheet, cell_address)
                    except Exception:
                        return "N/A"
                else:
                    return "N/A"
            
            # 결측치 체크 (None, 빈 문자열, 공백만 있는 값)
            if raw_value is None:
                return "N/A"
            if isinstance(raw_value, str) and not raw_value.strip():
                return "N/A"
            if isinstance(raw_value, list):
                # 리스트인 경우 None이나 빈 값 필터링
                filtered = [v for v in raw_value if v is not None and (not isinstance(v, str) or v.strip())]
                if not filtered:
                    return "N/A"
                raw_value = filtered
            
            # 계산 처리
            if operation:
                calculated = self._apply_operation(raw_value, operation)
            else:
                # 계산이 필요없는 경우
                if isinstance(raw_value, list):
                    # 범위인 경우 첫 번째 값 사용
                    calculated = raw_value[0] if raw_value else None
                else:
                    calculated = raw_value
            
            # 결과 포맷팅
            return self._format_value(calculated, operation)
                
        except Exception as e:
            # 에러 발생 시 N/A 반환
            return "N/A"
    
    def _apply_operation(self, raw_value: Any, operation: str) -> Any:
        """계산식을 적용합니다. 오류 발생 시 None 반환."""
        try:
            operation_lower = operation.lower().strip()
            
            # 결측치 체크
            if raw_value is None:
                return None
            
            # 범위인 경우 (리스트)
            if isinstance(raw_value, list):
                # 빈 리스트 체크
                if not raw_value:
                    return None
                
                # None 값 필터링
                filtered_values = [v for v in raw_value if v is not None and str(v).strip() != '']
                if not filtered_values:
                    return None
                
                # 증감률/증감액 계산의 경우 범위에서 첫 두 값 사용
                if operation_lower in ['growth_rate', '증감률', '증가율', 'growth_amount', '증감액', '증가액']:
                    if len(filtered_values) >= 2:
                        try:
                            return self.calculator.calculate(operation_lower, filtered_values[0], filtered_values[1])
                        except (ValueError, ZeroDivisionError, TypeError):
                            return None
                    else:
                        return None
                else:
                    # 다른 계산 (sum, average, max, min 등)
                    try:
                        return self.calculator.calculate_from_cell_refs(operation_lower, filtered_values)
                    except (ValueError, ZeroDivisionError, TypeError):
                        return None
            else:
                # 단일 값인 경우
                # 빈 문자열이나 공백만 있는 값 체크
                if isinstance(raw_value, str) and not raw_value.strip():
                    return None
                
                # 일부 계산식은 단일 값도 처리 (예: format)
                if operation_lower in ['format', '포맷']:
                    try:
                        return self.format_number(raw_value, decimal_places=1)
                    except (ValueError, TypeError):
                        return None
                # 계산식이 있지만 단일 값인 경우 그대로 사용
                return raw_value
        except Exception:
            # 모든 예외를 None으로 변환
            return None
    
    def _format_value(self, value: Any, operation: str = None) -> str:
        """값을 포맷팅합니다. 결측치나 오류는 N/A로 반환."""
        # None 체크
        if value is None:
            return "N/A"
        
        # 빈 문자열이나 공백만 있는 값 체크
        if isinstance(value, str):
            if not value.strip():
                return "N/A"
            # "N/A", "n/a", "null", "None" 등의 문자열도 N/A로 처리
            value_lower = value.strip().lower()
            if value_lower in ['n/a', 'na', 'null', 'none', '#n/a', '#na', '#null', '#ref!', '#value!', '#div/0!']:
                return "N/A"
        
        try:
            float_val = float(value)
            # NaN이나 Infinity 체크
            if math.isnan(float_val) or math.isinf(float_val):
                return "N/A"
            
            # 퍼센트 연산인 경우 퍼센트 포맷
            if operation and operation.lower() in ['growth_rate', '증감률', '증가율']:
                return self.format_percentage(float_val, decimal_places=1)
            # 일반 숫자 포맷팅 (천 단위 구분, 소수점 첫째자리까지)
            return self.format_number(float_val, use_comma=True, decimal_places=1)
        except (ValueError, TypeError, OverflowError):
            # 숫자가 아닌 경우 그대로 반환 (하지만 빈 문자열은 이미 N/A 처리됨)
            return str(value) if value else "N/A"
    
    def fill_template(self, sheet_name: str = None, year: int = None, quarter: int = None) -> str:
        """
        템플릿의 모든 마커를 처리하여 완성된 템플릿을 반환합니다.
        
        Args:
            sheet_name: 시트 이름 (기본값: None, 마커에서 추출)
            year: 연도 (None이면 자동 감지)
            quarter: 분기 (None이면 자동 감지)
        
        Returns:
            모든 마커가 값으로 치환된 HTML 템플릿
        """
        # 연도/분기가 지정되지 않으면 자동 감지
        if sheet_name and (year is None or quarter is None):
            periods_info = self.period_detector.detect_available_periods(sheet_name)
            if year is None:
                year = periods_info['default_year']
            if quarter is None:
                quarter = periods_info['default_quarter']
        
        # 현재 처리 중인 연도/분기/시트명 저장
        self._current_year = year
        self._current_quarter = quarter
        self._current_sheet_name = sheet_name
        # 템플릿 로드
        template_content = self.template_manager.get_template_content()
        
        # 시트명이 있으면 "시트명" 플레이스홀더를 실제 시트 이름으로 치환
        if self.sheet_name:
            template_content = template_content.replace('{시트명:', f'{{{self.sheet_name}:')
        
        # CSS와 스크립트 섹션을 제외하고 처리
        # <style>...</style> 및 <script>...</script> 태그 내부는 제외
        style_pattern = re.compile(r'<style[^>]*>.*?</style>', re.IGNORECASE | re.DOTALL)
        script_pattern = re.compile(r'<script[^>]*>.*?</script>', re.IGNORECASE | re.DOTALL)
        
        # CSS와 스크립트 섹션을 임시로 마스킹
        style_placeholders = {}
        script_placeholders = {}
        
        def style_replacer(match):
            placeholder = f"__STYLE_PLACEHOLDER_{len(style_placeholders)}__"
            style_placeholders[placeholder] = match.group(0)
            return placeholder
        
        def script_replacer(match):
            placeholder = f"__SCRIPT_PLACEHOLDER_{len(script_placeholders)}__"
            script_placeholders[placeholder] = match.group(0)
            return placeholder
        
        # CSS와 스크립트 섹션을 임시로 교체
        template_content = style_pattern.sub(style_replacer, template_content)
        template_content = script_pattern.sub(script_replacer, template_content)
        
        # 동적 마커 패턴: {시트명:동적키} (한글 시트명 포함)
        # 실제 마커만 매칭하도록 더 구체적으로 (한글이나 특정 패턴 포함)
        dynamic_marker_pattern = re.compile(r'\{([^:{}]+):([^:}]+)\}')
        
        # 동적 마커 찾기 및 치환
        # 여러 번 치환될 수 있으므로 반복 처리
        max_iterations = 10  # 무한 루프 방지
        iteration = 0
        while iteration < max_iterations:
            iteration += 1
            found_any = False
            
            # 셀 주소 형식이 아닌 경우 동적 마커로 처리
            # 하지만 CSS 속성명 같은 것들은 제외 (한글이나 숫자가 포함된 경우만)
            if not re.match(r'^[A-Z]+\d+', dynamic_key):
                # 실제 마커인지 확인 (한글, 숫자, 언더스코어 포함)
                if re.search(r'[가-힣0-9_]', dynamic_key):
                    try:
                        # sheet_name이 제공되지 않으면 마커에서 추출한 시트명 사용
                        marker_sheet_name = sheet_name if sheet_name else match.group(1)
                        
                        # "완료체크" 시트명인 경우 템플릿 내용 기반으로 올바른 시트명 찾기
                        if marker_sheet_name == '완료체크':
                            # 시트 해석기 초기화 (아직 안 된 경우)
                            if self.sheet_resolver is None:
                                available_sheets = self.excel_extractor.get_sheet_names()
                                self.sheet_resolver = SheetResolver(available_sheets)
                            
                            resolved_sheet = self.sheet_resolver.resolve_marker_in_template(
                                self.template_manager.template_content,
                                full_match
                            )
                            if resolved_sheet:
                                marker_sheet_name = resolved_sheet
                            else:
                                # 컨텍스트에서 찾지 못한 경우 기본값으로 "건설 (공표자료)" 사용
                                marker_sheet_name = '건설 (공표자료)'
                        
                        # 현재 저장된 연도/분기 사용
                        current_year = year if year is not None else (self._current_year or 2025)
                        current_quarter = quarter if quarter is not None else (self._current_quarter or 2)
                        dynamic_value = self._process_dynamic_marker(marker_sheet_name, dynamic_key, current_year, current_quarter)
                        if dynamic_value is not None:
                            template_content = template_content.replace(full_match, dynamic_value)
                        else:
                            # 값이 없으면 N/A로 채움
                            template_content = template_content.replace(full_match, 'N/A')
                    except Exception:
                        # 에러 발생 시 N/A로 채움
                        template_content = template_content.replace(full_match, 'N/A')
        
        # 일반 마커 추출 및 처리
        markers = self.template_manager.extract_markers()
        
        # 각 마커를 처리하여 치환
        for marker_info in markers:
            marker_str = marker_info['full_match']
            # sheet_name이 제공되면 마커의 시트명을 덮어쓰기
            if sheet_name:
                marker_info['sheet_name'] = sheet_name
            else:
                # 시트명이 제공되지 않으면 유연한 매핑으로 찾기
                marker_sheet = marker_info['sheet_name']
                actual_sheet = self.flexible_mapper.find_sheet_by_name(marker_sheet)
                if actual_sheet:
                    marker_info['sheet_name'] = actual_sheet
            
            processed_value = self.process_marker(marker_info)
            
            # 값이 비어있거나 에러인 경우 N/A로 채움
            if not processed_value or processed_value.startswith('[에러'):
                processed_value = 'N/A'
            
            # 마커를 처리된 값으로 치환
            template_content = template_content.replace(marker_str, processed_value)
        
        # CSS와 스크립트 섹션 복원
        for placeholder, original in style_placeholders.items():
            template_content = template_content.replace(placeholder, original)
        for placeholder, original in script_placeholders.items():
            template_content = template_content.replace(placeholder, original)
        
        return template_content
    
    def _generate_region_list_html(self, regions: List[Dict], css_class: str) -> str:
        """
        지역 리스트를 HTML로 생성합니다.
        
        Args:
            regions: 지역 정보 리스트
            css_class: CSS 클래스 ('positive' 또는 'negative')
            
        Returns:
            HTML 문자열
        """
        html_parts = []
        for region in regions:
            region_name = region.get('name', '')
            growth_rate = region.get('growth_rate', 0)
            formatted_rate = self.format_percentage(growth_rate, decimal_places=1)
            industries = region.get('top_industries', [])
            
            industry_parts = []
            for i, industry in enumerate(industries[:3]):  # 최대 3개
                industry_name = industry.get('name', '')
                industry_growth = industry.get('growth_rate', 0)
                
                # 산업 이름 매핑
                mapped_name = INDUSTRY_NAME_MAPPING.get(industry_name)
                if not mapped_name:
                    for key_map, value_map in INDUSTRY_NAME_MAPPING.items():
                        if key_map in industry_name or industry_name in key_map:
                            mapped_name = value_map
                            break
                display_name = mapped_name if mapped_name else industry_name
                
                industry_class = 'positive' if industry_growth > 0 else 'negative'
                formatted_industry_rate = self.format_percentage(industry_growth, decimal_places=1)
                
                industry_parts.append(
                    f'{display_name}(<span class="percentage {industry_class}">{formatted_industry_rate}</span>)'
                )
            
            industry_html = ', '.join(industry_parts)
            
            html_parts.append(f'''        <div class="region-item">
            <span class="region-name">{region_name}({formatted_rate}):</span>
            <div class="industry-list">
                <div class="industry-item">{industry_html}</div>
            </div>
        </div>''')
        
        return '\n'.join(html_parts)
    
    def _generate_dynamic_table(self, sheet_name: str) -> str:
        """
        동적으로 표를 생성합니다.
        
        Args:
            sheet_name: 시트 이름
            
        Returns:
            표 HTML 문자열
        """
        # SheetConfig에서 표 헤더 가져오기
        if self.sheet_config:
            headers = self.sheet_config.get_table_headers()
            table_quarters = self.sheet_config.get_table_quarters()
        else:
            headers = ['2021 Q2', '2021 Q3', '2021 Q4', '2022 Q1', 
                      '2022 Q2', '2022 Q3', '2022 Q4', '2023 Q1p']
            table_quarters = [
                ('2021', 2), ('2021', 3), ('2021', 4),
                ('2022', 1), ('2022', 2), ('2022', 3), ('2022', 4),
                ('2023', 1)
            ]
        
        # 표 제목
        subsection_title = self.sheet_config.get_subsection_title() if self.sheet_config else '광공업생산'
        table_caption = f'《{subsection_title} 증감률(불변)》'
        
        # 표 헤더 HTML 생성
        header_html = '<th>시·도</th>'
        for header in headers:
            header_html += f'<th>{header}</th>'
        
        # 모든 지역 목록 가져오기
        if self.config is not None:
            current_col, prev_col = self.config.get_column_pair()
        elif self.sheet_config:
            current_col, prev_col = self.sheet_config.get_config().get_column_pair()
        else:
            current_col, prev_col = 56, 52
        
        all_regions = self.data_analyzer.get_regions_with_growth_rate(
            sheet_name, current_col, prev_col
        )
        
        # 지역 순서 정의 (전국 먼저, 그 다음 시도 순서)
        region_order = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', 
                       '세종', '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 지역을 순서대로 정렬
        sorted_regions = []
        for region_name in region_order:
            for region in all_regions:
                if region['name'] == region_name:
                    sorted_regions.append(region)
                    break
        
        # 나머지 지역 추가 (순서에 없는 경우)
        for region in all_regions:
            if region not in sorted_regions:
                sorted_regions.append(region)
        
        # 표 행 HTML 생성
        rows_html = []
        for region in sorted_regions:
            region_name = region['name']
            row_html = f'            <tr>\n                <td class="region-name-cell">{region_name}</td>'
            
            # 각 분기별 증감률 계산
            for year_str, quarter_num in table_quarters:
                year = int(year_str)
                quarter_key = f'{year}_{quarter_num}분기'
                growth_rate = self._get_quarterly_growth_rate(sheet_name, region_name, quarter_key)
                
                if growth_rate is not None:
                    formatted = f"{growth_rate:.1f}".rstrip('0').rstrip('.')
                    cell_class = 'positive' if growth_rate > 0 else 'negative' if growth_rate < 0 else ''
                    row_html += f'\n                <td class="data-cell {cell_class}">{formatted}</td>'
                else:
                    row_html += '\n                <td class="data-cell">-</td>'
            
            row_html += '\n            </tr>'
            rows_html.append(row_html)
        
        # 전체 표 HTML 조합
        table_html = f'''    <table>
        <caption>{table_caption}</caption>
        <thead>
            <tr>
                {header_html}
            </tr>
        </thead>
        <tbody>
{chr(10).join(rows_html)}
        </tbody>
    </table>'''
        
        return table_html

    def fill_template_with_custom_format(self, format_func: callable = None) -> str:
        """
        커스텀 포맷팅 함수를 사용하여 템플릿을 채웁니다.
        
        Args:
            format_func: 포맷팅 함수 (marker_info, raw_value) -> str
        
        Returns:
            완성된 템플릿
        """
        template_content = self.template_manager.get_template_content()
        markers = self.template_manager.extract_markers()
        
        for marker_info in markers:
            marker_str = marker_info['full_match']
            
            try:
                raw_value = self.excel_extractor.extract_value(
                    marker_info['sheet_name'],
                    marker_info['cell_address']
                )
                
                if format_func:
                    processed_value = format_func(marker_info, raw_value)
                else:
                    processed_value = self.process_marker(marker_info)
                
                template_content = template_content.replace(marker_str, processed_value)
            except Exception as e:
                # 에러 발생 시 마커를 에러 메시지로 치환
                template_content = template_content.replace(
                    marker_str, 
                    f"[에러: {str(e)}]"
                )
        
        return template_content

