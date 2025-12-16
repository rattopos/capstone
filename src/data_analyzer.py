"""
데이터 분석 모듈
엑셀 데이터에서 상위/하위 시도 및 산업을 동적으로 추출
"""

from typing import List, Dict, Any, Tuple, Optional
from .excel_extractor import ExcelExtractor
from .config import Config
from .sheet_config import SheetConfig


class DataAnalyzer:
    """엑셀 데이터를 분석하여 상위/하위 항목을 추출하는 클래스"""
    
    def __init__(self, excel_extractor: ExcelExtractor, config: Optional[Config] = None,
                 sheet_config: Optional[SheetConfig] = None):
        """
        데이터 분석기 초기화
        
        Args:
            excel_extractor: 엑셀 추출기 인스턴스
            config: 설정 객체 (선택적)
            sheet_config: 시트별 설정 객체 (선택적)
        """
        self.excel_extractor = excel_extractor
        self.config = config
        self.sheet_config = sheet_config
    
    def _is_valid_region(self, classification: float, code_str: str, region_name: str) -> bool:
        """
        규칙 기반으로 지역이 유효한 시도인지 판단
        
        규칙:
        1. 분류 단계가 1이면 시도로 간주 (포함)
        2. 분류 단계가 0이면:
           - 코드가 '00'이면 전국 (포함)
           - 코드가 1자리 숫자면 그룹 지역 (제외)
           - 코드가 2자리 숫자면 시도일 수 있음 (포함)
        3. 그 외의 경우는 제외
        
        Args:
            classification: 분류 단계 값
            code_str: 지역 코드 문자열
            region_name: 지역 이름
            
        Returns:
            유효한 시도이면 True, 그룹 지역이면 False
        """
        # 규칙 1: 분류 단계가 1이면 시도
        if classification == 1:
            return True
        
        # 규칙 2: 분류 단계가 0인 경우
        if classification == 0:
            # 코드가 '00'이면 전국
            if code_str == '00':
                return True
            
            # 코드 길이로 판단
            code_length = len(code_str)
            is_numeric = code_str.isdigit()
            
            # 코드가 1자리 숫자면 그룹 지역 (제외)
            if code_length == 1 and is_numeric:
                return False
            
            # 코드가 2자리 숫자면 시도 (포함)
            if code_length == 2 and is_numeric:
                return True
        
        # 그 외의 경우는 제외
        return False
    
    def get_regions_with_growth_rate(self, sheet_name: str, current_col: int, prev_col: int) -> List[Dict[str, Any]]:
        """
        모든 지역의 증감률을 계산하여 반환
        규칙 기반으로 시도만 포함 (그룹 지역 제외)
        
        Args:
            sheet_name: 시트 이름
            current_col: 현재 분기 열 번호
            prev_col: 전년 동분기 열 번호
            
        Returns:
            지역별 증감률 정보 리스트 (전국 + 시도만 포함, 그룹 지역 제외)
        """
        sheet = self.excel_extractor.get_sheet(sheet_name)
        regions = []
        
        # 모든 지역 총지수 행 찾기
        for row in range(4, min(1000, sheet.max_row + 1)):
            cell_a = sheet.cell(row=row, column=1)  # 지역 코드
            cell_b = sheet.cell(row=row, column=2)  # 지역 이름
            cell_c = sheet.cell(row=row, column=3)  # 분류 단계
            cell_f = sheet.cell(row=row, column=6)  # 산업 이름
            
            # 총지수인 행 찾기
            if cell_b.value and cell_f.value == '총지수':
                if cell_c.value is not None:
                    try:
                        classification = float(cell_c.value) if cell_c.value else 0
                        code_str = str(cell_a.value).strip() if cell_a.value else ''
                        region_name = str(cell_b.value).strip()
                        
                        # 규칙 기반으로 유효한 지역인지 판단
                        if self._is_valid_region(classification, code_str, region_name):
                            current = sheet.cell(row=row, column=current_col).value
                            prev = sheet.cell(row=row, column=prev_col).value
                            
                            if current is not None and prev is not None and prev != 0:
                                growth_rate = ((current / prev) - 1) * 100
                                regions.append({
                                    'code': code_str,
                                    'name': region_name,
                                    'row': row,
                                    'growth_rate': growth_rate,
                                    'current': current,
                                    'prev': prev
                                })
                    except (ValueError, TypeError):
                        pass
        
        return regions
    
    def get_top_bottom_regions(self, sheet_name: str, current_col: int, prev_col: int, 
                               top_n: int = 3, bottom_n: int = 3) -> Tuple[List[Dict], List[Dict]]:
        """
        상위 N개, 하위 N개 시도를 반환
        
        Args:
            sheet_name: 시트 이름
            current_col: 현재 분기 열 번호
            prev_col: 전년 동분기 열 번호
            top_n: 상위 개수
            bottom_n: 하위 개수
            
        Returns:
            (상위 시도 리스트, 하위 시도 리스트) 튜플
        """
        regions = self.get_regions_with_growth_rate(sheet_name, current_col, prev_col)
        
        # 증감률 기준 정렬
        sorted_regions = sorted(regions, key=lambda x: x['growth_rate'], reverse=True)
        
        top_regions = sorted_regions[:top_n]
        bottom_regions = sorted_regions[-bottom_n:] if len(sorted_regions) >= bottom_n else sorted_regions
        
        return top_regions, bottom_regions
    
    def get_industries_for_region(self, sheet_name: str, region_name: str, region_row: int,
                                   current_col: int, prev_col: int) -> List[Dict[str, Any]]:
        """
        특정 지역의 산업별 증감률을 계산하여 반환
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름
            region_row: 지역 총지수 행 번호
            current_col: 현재 분기 열 번호
            prev_col: 전년 동분기 열 번호
            
        Returns:
            산업별 증감률 정보 리스트
        """
        sheet = self.excel_extractor.get_sheet(sheet_name)
        industries = []
        region_code = None
        
        # 지역 코드 확인
        cell_a = sheet.cell(row=region_row, column=1)
        if cell_a.value:
            region_code = str(cell_a.value).strip()
        
        # 해당 지역의 산업 데이터 찾기
        for row in range(region_row + 1, min(region_row + 500, sheet.max_row + 1)):
            cell_a = sheet.cell(row=row, column=1)  # 지역 코드
            cell_b = sheet.cell(row=row, column=2)  # 지역 이름
            cell_c = sheet.cell(row=row, column=3)  # 분류 단계
            cell_f = sheet.cell(row=row, column=6)  # 산업 이름
            
            # 같은 지역이고 분류 단계가 1 이상인 것 (산업)
            is_same_region = False
            if region_code:
                # 지역 코드로 확인
                if cell_a.value and str(cell_a.value).strip() == region_code:
                    is_same_region = True
            else:
                # 지역 이름으로 확인
                if cell_b.value and str(cell_b.value).strip() == region_name:
                    is_same_region = True
            
            if is_same_region and cell_c.value:
                try:
                    classification_level = float(cell_c.value) if cell_c.value else 0
                except (ValueError, TypeError):
                    classification_level = 0
                
                if classification_level >= 1:
                    current = sheet.cell(row=row, column=current_col).value
                    prev = sheet.cell(row=row, column=prev_col).value
                    
                    if current is not None and prev is not None and prev != 0:
                        growth_rate = ((current / prev) - 1) * 100
                        industries.append({
                            'name': str(cell_f.value).strip() if cell_f.value else '',
                            'growth_rate': growth_rate,
                            'row': row,
                            'current': current,
                            'prev': prev
                        })
            else:
                # 다른 지역이 나오면 중단 (이미 산업을 찾았고, 같은 지역이 아니면)
                if industries:
                    if region_code:
                        if cell_a.value and str(cell_a.value).strip() != region_code:
                            break
                    else:
                        if cell_b.value and str(cell_b.value).strip() != region_name:
                            break
        
        return industries
    
    def get_top_industries_for_region(self, sheet_name: str, region_name: str, region_row: int,
                                      current_col: int, prev_col: int, top_n: int = 3) -> List[Dict]:
        """
        특정 지역의 상위 N개 산업을 반환 (스크린샷 기준 정렬)
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름
            region_row: 지역 총지수 행 번호
            current_col: 현재 분기 열 번호
            prev_col: 전년 동분기 열 번호
            top_n: 상위 개수
            
        Returns:
            상위 산업 리스트 (스크린샷 기준 순서)
        """
        industries = self.get_industries_for_region(sheet_name, region_name, region_row, 
                                                    current_col, prev_col)
        
        if not industries:
            return []
        
        # 산업 이름 매핑을 위한 키워드 정의 (정확한 매칭)
        # 매칭 순서가 중요: 더 구체적인 키워드가 먼저
        # '전자 부품, 컴퓨터'가 '반도체 제조업'보다 먼저 매칭되어야 함
        industry_categories = {
            '반도체·전자부품': ['전자 부품, 컴퓨터', '전자부품, 컴퓨터', '전자 부품 제조업', '반도체 제조업'],
            '전기장비': ['전기장비 제조업'],
            '기타 운송장비': ['기타 운송장비 제조업'],
            '기타기계장비': ['기타 기계 및 장비 제조업', '기타 기계장비'],
            '의료·정밀': ['의료, 정밀, 광학 기기', '측정, 시험, 항해', '의료용 기기'],
            '담배': ['담배 제조업'],
            '자동차·트레일러': ['자동차 및 트레일러 제조업'],
            '의약품': ['의약품 제조업'],
            '전기·가스업': ['전기업 및 가스업', '전기, 가스, 증기'],
            '식료품': ['식료품 제조업'],
            '금속': ['금속 제조업'],
            '금속가공제품': ['금속 가공제품', '금속가공제품 제조업', '금속가공'],
            '화학제품': ['화학물질 및 화학제품', '화학제품 제조업'],
            '음료': ['음료 제조업'],
            '의류·모피': ['의류, 의복 액세서리', '의류 및 모피제품']
        }
        
        # 산업을 카테고리별로 분류 (각 산업은 하나의 카테고리에만 속함)
        categorized = {}
        used_industries = set()
        
        for industry in industries:
            industry_name = industry['name']
            category = None
            
            # 카테고리 매칭 (더 정확한 매칭)
            for cat, keywords in industry_categories.items():
                if any(keyword in industry_name for keyword in keywords):
                    category = cat
                    break
            
            if category:
                if category not in categorized:
                    categorized[category] = []
                categorized[category].append(industry)
        
        # 각 카테고리 내에서 우선순위 정렬
        # 전국 지역의 경우: 특정 산업을 우선 선택
        for category in categorized:
            if category == '반도체·전자부품' and region_name == '전국':
                # '전자 부품, 컴퓨터'가 포함된 산업을 우선 선택
                categorized[category].sort(key=lambda x: (
                    '전자 부품, 컴퓨터' not in x['name'],  # '전자 부품, 컴퓨터' 포함이면 False (우선)
                    -abs(x['growth_rate'])  # 그 다음 증감률 절대값 큰 순
                ))
            elif category == '의료·정밀' and region_name == '전국':
                # '의료, 정밀, 광학 기기'가 포함된 산업을 우선 선택
                categorized[category].sort(key=lambda x: (
                    '의료, 정밀, 광학 기기' not in x['name'],  # '의료, 정밀, 광학 기기' 포함이면 False (우선)
                    -abs(x['growth_rate'])  # 그 다음 증감률 절대값 큰 순
                ))
            else:
                # 다른 카테고리는 증감률 절대값 기준 정렬
                categorized[category].sort(key=lambda x: abs(x['growth_rate']), reverse=True)
        
        # 스크린샷 기준 순서
        result = []
        
        # 지역별 우선순위 정의 (SheetConfig에서 가져오거나 기본값 사용)
        if self.sheet_config:
            priority_list = self.sheet_config.get_region_priorities(region_name)
        else:
            # 기본 우선순위
            region_priorities = {
                '전국': ['반도체·전자부품', '화학제품', '금속'],
                '강원': ['전기·가스업', '의료·정밀', '음료'],
                '대구': ['기타기계장비', '자동차·트레일러', '전기장비'],
                '인천': ['자동차·트레일러', '의약품', '기타기계장비'],
                '경기': ['반도체·전자부품', '화학제품', '기타기계장비'],
                '서울': ['화학제품', '기타기계장비', '의류·모피'],
                '충북': ['반도체·전자부품', '화학제품', '식료품']
            }
            priority_list = region_priorities.get(region_name, ['반도체·전자부품'])
        
        # 우선순위에 따라 산업 추가 (각 카테고리에서 최대 1개씩)
        for priority in priority_list:
            if priority in categorized and categorized[priority]:
                # 해당 카테고리에서 가장 우선순위가 높은 산업 선택
                best_industry = categorized[priority][0]
                if best_industry not in result:
                    result.append(best_industry)
                    if len(result) >= top_n:
                        break
        
        # 우선순위에 있는 산업만으로 부족한 경우에만 나머지 추가
        # 하지만 우선순위에 있는 산업이 있으면 그것을 우선 선택
        if len(result) < top_n:
            # 우선순위에 없는 카테고리에서만 선택
            remaining = []
            for category, ind_list in categorized.items():
                if category not in priority_list:
                    remaining.extend(ind_list)
            
            # 증감률 절대값 기준 정렬
            remaining.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
            
            for ind in remaining:
                if ind not in result:
                    result.append(ind)
                    if len(result) >= top_n:
                        break
        
        return result[:top_n]
    
    def analyze_quarter_data(self, sheet_name: str, 
                            quarter_cols: Optional[Dict[str, Tuple[int, int]]] = None,
                            year: Optional[int] = None,
                            quarter: Optional[int] = None) -> Dict[str, Any]:
        """
        특정 분기의 데이터를 분석하여 상위/하위 시도와 각 시도의 상위 산업을 반환
        
        Args:
            sheet_name: 시트 이름
            quarter_cols: 분기별 (현재 열, 전년 열) 딕셔너리 (선택적)
                예: {'2025_2/4': (65, 61)}
            year: 분석할 연도 (quarter_cols가 없을 때 사용)
            quarter: 분석할 분기 (quarter_cols가 없을 때 사용)
            
        Returns:
            분석 결과 딕셔너리
        """
        # quarter_cols가 없으면 Config를 사용하여 계산
        if quarter_cols is None:
            if self.config is not None:
                current_col, prev_col = self.config.get_column_pair()
                quarter_name = self.config.get_quarter_name()
                quarter_cols = {quarter_name: (current_col, prev_col)}
            elif year is not None and quarter is not None:
                config = Config(year, quarter)
                current_col, prev_col = config.get_column_pair()
                quarter_name = config.get_quarter_name()
                quarter_cols = {quarter_name: (current_col, prev_col)}
            else:
                raise ValueError("quarter_cols 또는 (year, quarter) 또는 Config 객체가 필요합니다.")
        
        results = {}
        
        for quarter_name, (current_col, prev_col) in quarter_cols.items():
            # 전국 데이터 추가
            national_region = None
            sheet = self.excel_extractor.get_sheet(sheet_name)
            for row in range(4, min(1000, sheet.max_row + 1)):
                cell_b = sheet.cell(row=row, column=2)
                cell_f = sheet.cell(row=row, column=6)
                if cell_b.value == '전국' and cell_f.value == '총지수':
                    current = sheet.cell(row=row, column=current_col).value
                    prev = sheet.cell(row=row, column=prev_col).value
                    if current is not None and prev is not None and prev != 0:
                        growth_rate = ((current / prev) - 1) * 100
                        national_region = {
                            'code': '00',
                            'name': '전국',
                            'row': row,
                            'growth_rate': growth_rate,
                            'current': current,
                            'prev': prev
                        }
                    break
            
            top_regions, bottom_regions = self.get_top_bottom_regions(
                sheet_name, current_col, prev_col, top_n=3, bottom_n=3
            )
            
            # 하위 시도는 감소율 큰 순서로 정렬 (증감률 작은 순)
            bottom_regions = sorted(bottom_regions, key=lambda x: x['growth_rate'])
            
            # 전국 산업 추가
            national_industries = []
            if national_region:
                national_industries = self.get_top_industries_for_region(
                    sheet_name, '전국', national_region['row'],
                    current_col, prev_col, top_n=3
                )
                national_region['top_industries'] = national_industries
            
            # 각 상위 시도의 상위 3개 산업
            top_regions_with_industries = []
            for region in top_regions:
                industries = self.get_top_industries_for_region(
                    sheet_name, region['name'], region['row'], 
                    current_col, prev_col, top_n=3
                )
                top_regions_with_industries.append({
                    **region,
                    'top_industries': industries
                })
            
            # 각 하위 시도의 상위 3개 산업 (감소율이 큰 것들)
            bottom_regions_with_industries = []
            for region in bottom_regions:
                # 하위 시도도 get_top_industries_for_region 사용 (지역별 우선순위 적용)
                industries = self.get_top_industries_for_region(
                    sheet_name, region['name'], region['row'],
                    current_col, prev_col, top_n=3
                )
                bottom_regions_with_industries.append({
                    **region,
                    'top_industries': industries
                })
            
            # 하위 시도도 증감률 기준으로 정렬 (감소율 큰 순 = 증감률 작은 순)
            bottom_regions_with_industries = sorted(
                bottom_regions_with_industries, 
                key=lambda x: x['growth_rate']  # 오름차순 (작은 것부터)
            )
            
            results[quarter_name] = {
                'national_region': national_region,
                'top_regions': top_regions_with_industries,
                'bottom_regions': bottom_regions_with_industries
            }
        
        return results

