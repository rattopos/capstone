"""
템플릿 채우기 모듈
마커를 값으로 치환하고 포맷팅 처리
"""

import html
import re
from typing import Any, Dict, Optional
from .template_manager import TemplateManager
from .excel_extractor import ExcelExtractor
from .calculator import Calculator
from .data_analyzer import DataAnalyzer
from .config import Config
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
                 config: Optional[Config] = None):
        """
        템플릿 필러 초기화
        
        Args:
            template_manager: 템플릿 관리자 인스턴스
            excel_extractor: 엑셀 추출기 인스턴스
            config: 설정 객체 (선택적)
        """
        self.template_manager = template_manager
        self.excel_extractor = excel_extractor
        self.calculator = Calculator()
        self.config = config
        self.data_analyzer = DataAnalyzer(excel_extractor, config)
        self.nlp_processor = NLPProcessor()
        self._analyzed_data_cache = None
    
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
            if self.config is not None:
                # Config를 사용하여 분석
                self._analyzed_data_cache = self.data_analyzer.analyze_quarter_data(sheet_name)
            else:
                # 기본값: 2023 1/4 분기 데이터 분석 (Col 56: 현재, Col 52: 전년 동분기)
                # 2023 Q1 = Col 58 - 2 = 56, 2022 Q1 = Col 58 - 6 = 52
                quarter_data = {'2023_1/4': (56, 52)}
                self._analyzed_data_cache = self.data_analyzer.analyze_quarter_data(
                    sheet_name, quarter_data
                )
    
    def _get_quarterly_growth_rate(self, sheet_name: str, region_name: str, quarter: str) -> Optional[float]:
        """
        특정 지역의 특정 분기 증감률을 계산합니다.
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름 (예: '전국', '서울')
            quarter: 분기 문자열 (예: '2023_3분기', '2024_1분기')
            
        Returns:
            증감률 (퍼센트) 또는 None
        """
        # 분기별 열 매핑 (현재 분기 열, 전년 동분기 열)
        # 스크린샷 기준: 2021 Q2부터 2023 Q1p까지
        # 엑셀 열 번호 계산 필요 (2021 Q2 = 2020 Q2 대비, 2023 Q1p = 2022 Q1 대비)
        # 기준: 2023년 3/4분기 = Col 58 (BD)
        # 2021 Q2 = 2020 Q2 대비 → Col 계산 필요
        # 2021 Q3 = 2020 Q3 대비
        # 2021 Q4 = 2020 Q4 대비
        # 2022 Q1 = 2021 Q1 대비
        # 2022 Q2 = 2021 Q2 대비
        # 2022 Q3 = 2021 Q3 대비
        # 2022 Q4 = 2021 Q4 대비
        # 2023 Q1p = 2022 Q1 대비
        # Config 기준: 2023년 3/4분기 = Col 58
        # 2021 Q2 = Col 58 - (2년 * 4 + 1분기) = 58 - 9 = 49
        # 2021 Q3 = Col 58 - (2년 * 4 + 0분기) = 58 - 8 = 50
        # 2021 Q4 = Col 58 - (2년 * 4 - 1분기) = 58 - 7 = 51
        # 2022 Q1 = Col 58 - (1년 * 4 + 2분기) = 58 - 6 = 52
        # 2022 Q2 = Col 58 - (1년 * 4 + 1분기) = 58 - 5 = 53
        # 2022 Q3 = Col 58 - (1년 * 4 + 0분기) = 58 - 4 = 54
        # 2022 Q4 = Col 58 - (1년 * 4 - 1분기) = 58 - 3 = 55
        # 2023 Q1p = Col 58 - (0년 * 4 + 2분기) = 58 - 2 = 56
        quarter_cols = {
            '2021_2분기': (49, 45),  # 2021 Q2 vs 2020 Q2 (전년 동분기)
            '2021_3분기': (50, 46),  # 2021 Q3 vs 2020 Q3
            '2021_4분기': (51, 47),  # 2021 Q4 vs 2020 Q4
            '2022_1분기': (52, 48),  # 2022 Q1 vs 2021 Q1
            '2022_2분기': (53, 49),  # 2022 Q2 vs 2021 Q2
            '2022_3분기': (54, 50),  # 2022 Q3 vs 2021 Q3
            '2022_4분기': (55, 51),  # 2022 Q4 vs 2021 Q4
            '2023_1분기': (56, 52),  # 2023 Q1p vs 2022 Q1
        }
        
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
        
        # 증감률 계산
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
        self._analyze_data_if_needed(sheet_name)
        
        # 캐시에서 데이터 가져오기 (Config가 있으면 해당 분기, 없으면 기본값)
        if self.config is not None:
            quarter_name = self.config.get_quarter_name()
        else:
            quarter_name = '2023_1/4'
        
        if quarter_name not in self._analyzed_data_cache:
            return None
        
        data = self._analyzed_data_cache[quarter_name]
        
        # 전국 패턴
        if key == '전국_이름':
            if 'national_region' in data and data['national_region']:
                return data['national_region']['name']
        elif key == '전국_증감률':
            if 'national_region' in data and data['national_region']:
                return self.format_percentage(data['national_region']['growth_rate'], decimal_places=1)
        elif key == '전국_증감방향':
            if 'national_region' in data and data['national_region']:
                growth_rate = data['national_region']['growth_rate']
                return self.nlp_processor.determine_trend(growth_rate)
        elif key.startswith('전국_산업'):
            if 'national_region' in data and data['national_region']:
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
        
        # 상위 시도 패턴
        top_match = re.match(r'상위시도(\d+)_(.+)', key)
        if top_match:
            idx = int(top_match.group(1)) - 1  # 0-based
            field = top_match.group(2)
            
            if idx < len(data['top_regions']):
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
        
        # 하위 시도 패턴
        bottom_match = re.match(r'하위시도(\d+)_(.+)', key)
        if bottom_match:
            idx = int(bottom_match.group(1)) - 1  # 0-based
            field = bottom_match.group(2)
            
            if idx < len(data['bottom_regions']):
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
        
        # 기타 동적 키 처리
        if key == '증가시도수':
            # 시도만 카운트 (그룹 제외)
            # 시도 코드: 2자리 숫자, 00(전국) 제외, 11-39 범위
            # 그룹 코드: 1자리 숫자 또는 다른 형식 (수도, 대경, 호남, 충청 등)
            sheet = self.excel_extractor.get_sheet(sheet_name)
            negative_count = 0  # 감소 시도 수 (스크린샷 기준: 12개 시도 감소)
            
            # Config가 있으면 해당 열 사용, 없으면 기본값
            if self.config is not None:
                current_col, prev_col = self.config.get_column_pair()
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
        elif key == '기준연도':
            if self.config is not None:
                return str(self.config.year)
            return '2023'
        elif key == '기준분기':
            if self.config is not None:
                return f'{self.config.quarter}/4'
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
        for match in dynamic_marker_pattern.finditer(template_content):
            full_match = match.group(0)
            sheet_name = match.group(1)
            dynamic_key = match.group(2)
            
            # 셀 주소 형식이 아닌 경우 동적 마커로 처리
            # 하지만 CSS 속성명 같은 것들은 제외 (한글이나 숫자가 포함된 경우만)
            if not re.match(r'^[A-Z]+\d+', dynamic_key):
                # 실제 마커인지 확인 (한글, 숫자, 언더스코어 포함)
                if re.search(r'[가-힣0-9_]', dynamic_key):
                    try:
                        dynamic_value = self._process_dynamic_marker(sheet_name, dynamic_key)
                        if dynamic_value is not None:
                            template_content = template_content.replace(full_match, dynamic_value)
                    except Exception:
                        # 에러 발생 시 원본 유지
                        pass
        
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

