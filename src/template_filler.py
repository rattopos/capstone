"""
템플릿 채우기 모듈
마커를 값으로 치환하고 포맷팅 처리
"""

import html
import re
from typing import Any, Dict, Optional, List
from .template_manager import TemplateManager
from .excel_extractor import ExcelExtractor
from .calculator import Calculator
from .data_analyzer import DataAnalyzer
from .config import Config
from .sheet_config import SheetConfig
from .nlp_processor import NLPProcessor

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
        self.config = config
        self.sheet_name = sheet_name
        
        # SheetConfig 초기화 (시트명이 있으면 사용)
        if sheet_name:
            try:
                if config:
                    self.sheet_config = SheetConfig(sheet_name, config.year, config.quarter)
                else:
                    self.sheet_config = SheetConfig(sheet_name)
                # SheetConfig의 Config를 사용
                self.config = self.sheet_config.get_config()
            except Exception as e:
                # SheetConfig 초기화 실패 시 기본값 사용
                print(f"경고: SheetConfig 초기화 실패 ({sheet_name}): {str(e)}")
                self.sheet_config = None
        else:
            self.sheet_config = None
        
        self.data_analyzer = DataAnalyzer(excel_extractor, self.config, sheet_config=self.sheet_config)
        self.nlp_processor = NLPProcessor()
        self._analyzed_data_cache = None
        self._all_regions_cache = None  # 모든 지역 목록 캐시
    
    def format_number(self, value: Any, use_comma: bool = True, decimal_places: int = None) -> str:
        """
        숫자를 포맷팅합니다.
        
        Args:
            value: 포맷팅할 값
            use_comma: 천 단위 구분 기호 사용 여부
            decimal_places: 소수점 자릿수 (None이면 원래대로)
            
        Returns:
            포맷팅된 문자열
        """
        try:
            # 숫자로 변환
            num = float(value)
            
            # 소수점 처리
            if decimal_places is not None:
                num = round(num, decimal_places)
                # 소수점이 0이면 정수로 표시
                if decimal_places == 0:
                    num = int(num)
            
            # 문자열로 변환
            if decimal_places is not None and decimal_places > 0:
                formatted = f"{num:.{decimal_places}f}"
            else:
                formatted = str(int(num) if num.is_integer() else num)
            
            # 천 단위 구분
            if use_comma:
                parts = formatted.split('.')
                parts[0] = f"{int(parts[0]):,}"
                formatted = '.'.join(parts)
            
            return formatted
        except (ValueError, TypeError):
            # 숫자로 변환할 수 없으면 원본 반환
            return str(value) if value is not None else ""
    
    def format_percentage(self, value: Any, decimal_places: int = 1, include_percent: bool = True) -> str:
        """
        퍼센트 값을 포맷팅합니다.
        
        Args:
            value: 퍼센트 값 (예: 5.5는 5.5%를 의미)
            decimal_places: 소수점 자릿수
            include_percent: % 기호 포함 여부
            
        Returns:
            포맷팅된 퍼센트 문자열 (예: "5.5%" 또는 "5.5")
        """
        try:
            num = float(value)
            formatted = f"{num:.{decimal_places}f}".rstrip('0').rstrip('.')
            if include_percent:
                return f"{formatted}%"
            return formatted
        except (ValueError, TypeError):
            return str(value) if value is not None else ""
    
    def escape_html(self, value: Any) -> str:
        """
        HTML 특수 문자를 이스케이프합니다.
        
        Args:
            value: 이스케이프할 값
            
        Returns:
            이스케이프된 문자열
        """
        return html.escape(str(value) if value is not None else "")
    
    def _analyze_data_if_needed(self, sheet_name: str) -> None:
        """
        필요시 데이터를 분석하여 캐시에 저장
        
        Args:
            sheet_name: 시트 이름
        """
        if self._analyzed_data_cache is None:
            try:
                if self.config is not None:
                    # Config를 사용하여 분석
                    self._analyzed_data_cache = self.data_analyzer.analyze_quarter_data(sheet_name)
                elif self.sheet_config:
                    # SheetConfig 사용
                    config_obj = self.sheet_config.get_config()
                    current_col, prev_col = config_obj.get_column_pair()
                    quarter_name = config_obj.get_quarter_name()
                    quarter_data = {quarter_name: (current_col, prev_col)}
                    self._analyzed_data_cache = self.data_analyzer.analyze_quarter_data(
                        sheet_name, quarter_data
                    )
                else:
                    # 기본값: 2023 1/4 분기 데이터 분석 (Col 56: 현재, Col 52: 전년 동분기)
                    quarter_data = {'2023_1/4': (56, 52)}
                    self._analyzed_data_cache = self.data_analyzer.analyze_quarter_data(
                        sheet_name, quarter_data
                    )
            except Exception as e:
                # 분석 실패 시 빈 딕셔너리로 초기화
                print(f"경고: 데이터 분석 실패 ({sheet_name}): {str(e)}")
                self._analyzed_data_cache = {}
    
    def _get_quarterly_growth_rate(self, sheet_name: str, region_name: str, quarter: str) -> Optional[float]:
        """
        특정 지역의 특정 분기 증감률을 계산합니다.
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름 (예: '전국', '서울')
            quarter: 분기 문자열 (예: '2021_2분기', '2023_1분기')
            
        Returns:
            증감률 (퍼센트) 또는 None
        """
        # 분기별 열 매핑 (현재 분기 열, 전년 동분기 열)
        # Config 클래스를 사용하여 동적으로 열 계산
        quarter_cols = {}
        
        # SheetConfig에서 분기 목록 가져오기
        if self.sheet_config:
            table_quarters = self.sheet_config.get_table_quarters()
            
            for year_str, quarter_num in table_quarters:
                year = int(year_str)
                # Config 클래스를 사용하여 현재 분기 열 계산
                current_config = Config(year, quarter_num)
                current_col = current_config.get_current_column()
                
                # 전년 동분기 열 계산
                prev_col = current_config.get_previous_year_column()
                
                quarter_key = f'{year}_{quarter_num}분기'
                quarter_cols[quarter_key] = (current_col, prev_col)
        else:
            # SheetConfig가 없는 경우에도 Config 클래스 사용
            # 기본값: 2021 Q2부터 2023 Q1까지
            for year in [2021, 2022, 2023]:
                for quarter_num in [1, 2, 3, 4]:
                    # 2021년은 Q2부터 시작
                    if year == 2021 and quarter_num < 2:
                        continue
                    # 2023년은 Q1까지만
                    if year == 2023 and quarter_num > 1:
                        continue
                    
                    current_config = Config(year, quarter_num)
                    current_col = current_config.get_current_column()
                    prev_col = current_config.get_previous_year_column()
                    
                    quarter_key = f'{year}_{quarter_num}분기'
                    quarter_cols[quarter_key] = (current_col, prev_col)
        
        if quarter not in quarter_cols:
            return None
        
        current_col, prev_col = quarter_cols[quarter]
        
        # 지역의 총지수 행 찾기
        sheet = self.excel_extractor.get_sheet(sheet_name)
        region_row = None
        
        for row in range(4, min(1000, sheet.max_row + 1)):
            cell_b = sheet.cell(row=row, column=2)  # 지역 이름
            cell_f = sheet.cell(row=row, column=6)  # 산업 이름
            
            if cell_b.value and str(cell_b.value).strip() == region_name and cell_f.value == '총지수':
                region_row = row
                break
        
        if region_row is None:
            return None
        
        # 현재 분기와 전년 동분기 값 가져오기
        current_value = sheet.cell(row=region_row, column=current_col).value
        prev_value = sheet.cell(row=region_row, column=prev_col).value
        
        if current_value is None or prev_value is None or prev_value == 0:
            return None
        
        # 증감률 계산: 전년 동분기 대비
        # 공식: ((현재값 / 전년동분기값) - 1) * 100
        growth_rate = ((current_value / prev_value) - 1) * 100
        return growth_rate
    
    def _process_dynamic_marker(self, sheet_name: str, key: str) -> Optional[str]:
        """
        동적 마커를 처리합니다.
        
        Args:
            sheet_name: 시트 이름
            key: 동적 키 (예: '상위시도1_이름', '상위시도1_증감률')
            
        Returns:
            처리된 값 또는 None
        """
        # 데이터 분석이 필요하면 수행
        self._analyze_data_if_needed(sheet_name)
        
        # 캐시가 없거나 비어있으면 None 반환
        if self._analyzed_data_cache is None:
            return None
        
        # 캐시에서 데이터 가져오기 (Config가 있으면 해당 분기, 없으면 기본값)
        if self.config is not None:
            quarter_name = self.config.get_quarter_name()
        elif self.sheet_config:
            quarter_name = self.sheet_config.get_config().get_quarter_name()
        else:
            quarter_name = '2023_1/4'
        
        # 캐시가 딕셔너리가 아니거나 해당 분기 데이터가 없으면 None 반환
        if not isinstance(self._analyzed_data_cache, dict):
            return None
        
        if quarter_name not in self._analyzed_data_cache:
            return None
        
        data = self._analyzed_data_cache[quarter_name]
        
        # data가 None이거나 딕셔너리가 아니면 None 반환
        if data is None or not isinstance(data, dict):
            return None
        
        # 전국 패턴
        if key == '전국_이름':
            if data and 'national_region' in data and data['national_region']:
                return data['national_region']['name']
        elif key == '전국_증감률':
            if data and 'national_region' in data and data['national_region']:
                return self.format_percentage(data['national_region']['growth_rate'], decimal_places=1)
        elif key == '전국_증감방향':
            if data and 'national_region' in data and data['national_region']:
                growth_rate = data['national_region']['growth_rate']
                return self.nlp_processor.determine_trend(growth_rate)
        elif key.startswith('전국_산업'):
            if data and 'national_region' in data and data['national_region']:
                industry_match = re.match(r'전국_산업(\d+)_(.+)', key)
                if industry_match:
                    industry_idx = int(industry_match.group(1)) - 1
                    industry_field = industry_match.group(2)
                    
                    if 'top_industries' in data['national_region'] and industry_idx < len(data['national_region']['top_industries']):
                        industry = data['national_region']['top_industries'][industry_idx]
                        
                        if industry_field == '이름':
                            # 산업 이름 매핑 적용
                            industry_name = industry['name']
                            mapped_name = INDUSTRY_NAME_MAPPING.get(industry_name)
                            if not mapped_name:
                                # 부분 일치 확인
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
        
        # 상위 시도 패턴
        top_match = re.match(r'상위시도(\d+)_(.+)', key)
        if top_match:
            idx = int(top_match.group(1)) - 1  # 0-based
            field = top_match.group(2)
            
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
                    # 산업1_이름, 산업1_증감률 등
                    industry_match = re.match(r'산업(\d+)_(.+)', field)
                    if industry_match:
                        industry_idx = int(industry_match.group(1)) - 1
                        industry_field = industry_match.group(2)
                        
                        if 'top_industries' in region and industry_idx < len(region['top_industries']):
                            industry = region['top_industries'][industry_idx]
                            
                            if industry_field == '이름':
                                # 산업 이름 매핑 적용
                                industry_name = industry['name']
                                mapped_name = INDUSTRY_NAME_MAPPING.get(industry_name)
                                if not mapped_name:
                                    # 부분 일치 확인
                                    for key, value in INDUSTRY_NAME_MAPPING.items():
                                        if key in industry_name or industry_name in key:
                                            mapped_name = value
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
        
        # 하위 시도 패턴
        bottom_match = re.match(r'하위시도(\d+)_(.+)', key)
        if bottom_match:
            idx = int(bottom_match.group(1)) - 1  # 0-based
            field = bottom_match.group(2)
            
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
                    # 산업1_이름, 산업1_증감률 등
                    industry_match = re.match(r'산업(\d+)_(.+)', field)
                    if industry_match:
                        industry_idx = int(industry_match.group(1)) - 1
                        industry_field = industry_match.group(2)
                        
                        if 'top_industries' in region and industry_idx < len(region['top_industries']):
                            industry = region['top_industries'][industry_idx]
                            
                            if industry_field == '이름':
                                # 산업 이름 매핑 적용
                                industry_name = industry['name']
                                mapped_name = INDUSTRY_NAME_MAPPING.get(industry_name)
                                if not mapped_name:
                                    # 부분 일치 확인
                                    for key, value in INDUSTRY_NAME_MAPPING.items():
                                        if key in industry_name or industry_name in key:
                                            mapped_name = value
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
                # 표 셀에는 % 기호 없이 표시, 소수점 1자리
                formatted = f"{growth_rate:.1f}".rstrip('0').rstrip('.')
                return formatted
            return ""
        
        # 분기 헤더 마커 처리 (예: 분기1_헤더)
        header_match = re.match(r'분기(\d+)_헤더', key)
        if header_match:
            quarter_idx = int(header_match.group(1)) - 1
            # SheetConfig에서 헤더 가져오기
            if self.sheet_config:
                headers = self.sheet_config.get_table_headers()
            else:
                # 기본값
                headers = ['2021 Q2', '2021 Q3', '2021 Q4', '2022 Q1', 
                          '2022 Q2', '2022 Q3', '2022 Q4', '2023 Q1p']
            if 0 <= quarter_idx < len(headers):
                return headers[quarter_idx]
        
        # 전국 증감률 처리 (2023 1/4 분기)
        if key == '전국_증감률':
            growth_rate = self._get_quarterly_growth_rate(sheet_name, '전국', '2023_1분기')
            if growth_rate is not None:
                return self.format_percentage(growth_rate, decimal_places=1)
            return ""
        
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
            # 시도 코드: 2자리 숫자, 00(전국) 제외, 11-39 범위
            # 그룹 코드: 1자리 숫자 또는 다른 형식 (수도, 대경, 호남, 충청 등)
            sheet = self.excel_extractor.get_sheet(sheet_name)
            negative_count = 0  # 감소 시도 수 (스크린샷 기준: 12개 시도 감소)
            
            # Config가 있으면 해당 열 사용, 없으면 기본값
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
        elif key == '증가시도수':
            # 증가시도수도 지원 (하위 호환성)
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
        elif key == '기준연도':
            if self.config is not None:
                return str(self.config.year)
            elif self.sheet_config:
                return str(self.sheet_config.year)
            return '2023'
        elif key == '기준분기':
            if self.config is not None:
                return f'{self.config.quarter}/4'
            elif self.sheet_config:
                return f'{self.sheet_config.quarter}/4'
            return '1/4'
        
        return None
    
    def process_marker(self, marker_info: Dict[str, str]) -> str:
        """
        마커 정보를 처리하여 값을 추출하고 계산합니다.
        
        Args:
            marker_info: 마커 정보 딕셔너리
                - 'sheet_name': 시트명
                - 'cell_address': 셀 주소 또는 범위
                - 'operation': 계산식 (선택적)
        
        Returns:
            처리된 값 (문자열)
        """
        sheet_name = marker_info['sheet_name']
        cell_address = marker_info['cell_address']
        operation = marker_info.get('operation')
        
        try:
            # 동적 마커인지 확인 (셀 주소 형식이 아닌 경우)
            if not re.match(r'^[A-Z]+\d+', cell_address):
                # 동적 마커 처리
                dynamic_value = self._process_dynamic_marker(sheet_name, cell_address)
                if dynamic_value is not None:
                    return dynamic_value
                # 동적 마커가 아니면 에러
                return f"[에러: 동적 마커를 찾을 수 없음: {cell_address}]"
            
            # 엑셀에서 값 추출
            raw_value = self.excel_extractor.extract_value(sheet_name, cell_address)
            
            # 계산이 필요한 경우
            if operation:
                operation_lower = operation.lower().strip()
                
                # 범위인 경우 (리스트)
                if isinstance(raw_value, list):
                    # 증감률/증감액 계산의 경우 범위에서 첫 두 값 사용
                    if operation_lower in ['growth_rate', '증감률', '증가율', 'growth_amount', '증감액', '증가액']:
                        if len(raw_value) >= 2:
                            calculated = self.calculator.calculate(operation_lower, raw_value[0], raw_value[1])
                        else:
                            raise ValueError(f"{operation} 계산에는 최소 2개의 값이 필요합니다.")
                    else:
                        # 다른 계산 (sum, average, max, min 등)
                        calculated = self.calculator.calculate_from_cell_refs(operation_lower, raw_value)
                else:
                    # 단일 값인 경우
                    # 일부 계산식은 단일 값도 처리 (예: format)
                    if operation_lower in ['format', '포맷']:
                        return self.format_number(raw_value)
                    # 계산식이 있지만 단일 값인 경우 그대로 사용
                    calculated = raw_value
            else:
                # 계산이 필요없는 경우
                if isinstance(raw_value, list):
                    # 범위인 경우 첫 번째 값 사용
                    calculated = raw_value[0] if raw_value else None
                else:
                    calculated = raw_value
            
            # 결과 포맷팅
            if calculated is not None:
                try:
                    float_val = float(calculated)
                    # 퍼센트 연산인 경우 퍼센트 포맷
                    if operation and operation.lower() in ['growth_rate', '증감률', '증가율']:
                        return self.format_percentage(float_val, decimal_places=1)
                    # 일반 숫자 포맷팅 (천 단위 구분)
                    return self.format_number(float_val, use_comma=True)
                except (ValueError, TypeError):
                    # 숫자가 아닌 경우 그대로 반환 (HTML 이스케이프는 하지 않음 - HTML 내에서 사용)
                    return str(calculated)
            else:
                return ""
                
        except Exception as e:
            # 에러 발생 시 에러 메시지 반환 (디버깅용)
            return f"[에러: {str(e)}]"
    
    def fill_template(self) -> str:
        """
        템플릿의 모든 마커를 처리하여 완성된 템플릿을 반환합니다.
        
        Returns:
            모든 마커가 값으로 치환된 HTML 템플릿
        """
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
            
            for match in dynamic_marker_pattern.finditer(template_content):
                full_match = match.group(0)
                sheet_name = match.group(1)
                dynamic_key = match.group(2)
                
                # 셀 주소 형식이 아닌 경우 동적 마커로 처리
                if not re.match(r'^[A-Z]+\d+', dynamic_key):
                    # 실제 마커인지 확인 (한글, 숫자, 언더스코어 포함)
                    if re.search(r'[가-힣0-9_]', dynamic_key):
                        try:
                            dynamic_value = self._process_dynamic_marker(sheet_name, dynamic_key)
                            if dynamic_value is not None:
                                template_content = template_content.replace(full_match, str(dynamic_value), 1)
                                found_any = True
                            else:
                                # 값이 None이면 빈 문자열로 치환 (디버깅용으로 마커 키 표시)
                                # 실제 운영에서는 빈 문자열로 치환
                                template_content = template_content.replace(full_match, "", 1)
                                found_any = True
                        except Exception as e:
                            # 에러 발생 시 에러 메시지 표시 (디버깅용)
                            error_msg = f"[에러:{dynamic_key}:{str(e)}]"
                            template_content = template_content.replace(full_match, error_msg, 1)
                            found_any = True
            
            # 더 이상 치환할 마커가 없으면 종료
            if not found_any:
                break
        
        # 일반 마커 추출 및 처리
        markers = self.template_manager.extract_markers()
        
        # 각 마커를 처리하여 치환
        for marker_info in markers:
            marker_str = marker_info['full_match']
            processed_value = self.process_marker(marker_info)
            
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

