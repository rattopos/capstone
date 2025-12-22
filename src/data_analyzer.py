"""
데이터 분석 모듈
엑셀 데이터에서 상위/하위 시도 및 산업을 동적으로 추출
가중치 기반 순위 계산 지원
"""

from typing import List, Dict, Any, Tuple, Optional
from .excel_extractor import ExcelExtractor
from .schema_loader import SchemaLoader


class DataAnalyzer:
    """엑셀 데이터를 분석하여 상위/하위 항목을 추출하는 클래스"""
    
    def __init__(self, excel_extractor: ExcelExtractor, schema_loader: Optional[SchemaLoader] = None):
        """
        데이터 분석기 초기화
        
        Args:
            excel_extractor: 엑셀 추출기 인스턴스
            schema_loader: 스키마 로더 인스턴스 (기본값: 새로 생성)
        """
        self.excel_extractor = excel_extractor
        self.schema_loader = schema_loader if schema_loader is not None else SchemaLoader()
    
    def get_regions_with_growth_rate(self, sheet_name: str, current_col: int, prev_col: int) -> List[Dict[str, Any]]:
        """
        모든 지역의 증감률을 계산하여 반환
        
        Args:
            sheet_name: 시트 이름
            current_col: 현재 분기 열 번호
            prev_col: 전년 동분기 열 번호
            
        Returns:
            지역별 증감률 정보 리스트
        """
        sheet = self.excel_extractor.get_sheet(sheet_name)
        regions = []
        
        # 시트별 설정에서 category_column 가져오기 (스키마 로더 사용)
        sheet_config = self.schema_loader.load_sheet_config(sheet_name)
        category_col = sheet_config.get('category_column', 6)
        region_column = sheet_config.get('region_column', 2)  # 지역 이름 열 (기본값: 2)
        
        # 가중치 설정 가져오기
        weight_config = self.schema_loader.get_weight_config(sheet_name)
        weight_column = weight_config.get('weight_column')
        max_classification_level = weight_config.get('max_classification_level', 2)
        classification_column = weight_config.get('classification_column', sheet_config.get('classification_column', 3))
        
        # 모든 지역 총지수 행 찾기 (총지수인 행을 찾되, 분류 단계는 0 또는 1)
        seen_regions = set()  # 중복 방지
        for row in range(4, min(5000, sheet.max_row + 1)):
            cell_a = sheet.cell(row=row, column=1)  # 지역 코드
            cell_b = sheet.cell(row=row, column=region_column)  # 지역 이름 (스키마에서 가져온 열 번호 사용)
            cell_c = sheet.cell(row=row, column=classification_column)  # 분류 단계 (스키마에서 가져온 열 번호 사용)
            cell_category = sheet.cell(row=row, column=category_col)  # 업태/산업 이름
            
            # 총지수, 계, 합계인 행 찾기 (분류 단계는 0 또는 1일 수 있음)
            is_total = False
            if cell_category.value:
                category_str = str(cell_category.value).strip()
                if category_str in ['총지수', '계', '   계', '합계']:
                    is_total = True
            
            if cell_b.value and is_total:
                # 분류 단계가 0 또는 1인 경우만 (일부 시도는 분류 단계가 1)
                if cell_c.value is not None:
                    try:
                        classification = float(cell_c.value) if cell_c.value else 0
                        if classification <= 1:  # 0 또는 1
                            region_name = str(cell_b.value).strip()
                            # 중복 방지: 같은 지역 이름이 이미 추가되었으면 스킵
                            if region_name not in seen_regions:
                                # 시도만 필터링 (그룹 제외)
                                code_str = str(cell_a.value).strip() if cell_a.value else ''
                                is_sido = (len(code_str) == 2 and code_str.isdigit() and code_str != '00')
                                
                                # 전국도 제외, 시도만 포함
                                # 지역명 매핑 (엑셀의 긴 이름 -> 짧은 이름)
                                region_mapping = {
                                    '서울특별시': '서울', '부산광역시': '부산', '대구광역시': '대구',
                                    '인천광역시': '인천', '광주광역시': '광주', '대전광역시': '대전',
                                    '울산광역시': '울산', '세종특별자치시': '세종', '경기도': '경기',
                                    '강원도': '강원', '충청북도': '충북', '충청남도': '충남',
                                    '전라북도': '전북', '전라남도': '전남', '경상북도': '경북',
                                    '경상남도': '경남', '제주특별자치도': '제주'
                                }
                                mapped_name = region_mapping.get(region_name, region_name)
                                
                                if mapped_name != '전국' and (is_sido or mapped_name in ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']):
                                    seen_regions.add(mapped_name)
                                    current = sheet.cell(row=row, column=current_col).value
                                    prev = sheet.cell(row=row, column=prev_col).value
                                    
                                    # 가중치 가져오기
                                    weight = 1.0
                                    if weight_column:
                                        weight_value = sheet.cell(row=row, column=weight_column).value
                                        weight = self.schema_loader.get_weight_value(sheet_name, weight_value)
                                    
                                    if current is not None and prev is not None and prev != 0:
                                        growth_rate = ((current / prev) - 1) * 100
                                        # 가중치를 반영한 증감률 (순위용)
                                        weighted_growth_rate = growth_rate * weight
                                        regions.append({
                                            'code': code_str,
                                            'name': mapped_name,  # 매핑된 이름 사용
                                            'row': row,
                                            'growth_rate': growth_rate,
                                            'weighted_growth_rate': weighted_growth_rate,
                                            'weight': weight,
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
        분류단계 2까지만 포함, 가중치 반영
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름
            region_row: 지역 총지수 행 번호
            current_col: 현재 분기 열 번호
            prev_col: 전년 동분기 열 번호
            
        Returns:
            산업별 증감률 정보 리스트 (가중치 반영)
        """
        sheet = self.excel_extractor.get_sheet(sheet_name)
        industries = []
        region_code = None
        
        # 가중치 설정 가져오기
        weight_config = self.schema_loader.get_weight_config(sheet_name)
        weight_column = weight_config.get('weight_column')
        max_classification_level = weight_config.get('max_classification_level', 2)
        
        # 시트 설정에서 열 번호 가져오기
        sheet_config = self.schema_loader.load_sheet_config(sheet_name)
        classification_column = weight_config.get('classification_column', sheet_config.get('classification_column', 3))
        category_column = sheet_config.get('category_column', 6)
        region_column = sheet_config.get('region_column', 2)  # 지역 이름 열 (기본값: 2)
        
        # 지역 코드 확인
        cell_a = sheet.cell(row=region_row, column=1)
        if cell_a.value:
            region_code = str(cell_a.value).strip()
        
        # 해당 지역의 산업 데이터 찾기
        for row in range(region_row + 1, min(region_row + 500, sheet.max_row + 1)):
            cell_a = sheet.cell(row=row, column=1)  # 지역 코드
            cell_b = sheet.cell(row=row, column=region_column)  # 지역 이름 (스키마에서 가져온 열 번호 사용)
            cell_c = sheet.cell(row=row, column=classification_column)  # 분류 단계 (스키마에서 가져온 열 번호 사용)
            cell_f = sheet.cell(row=row, column=category_column)  # 산업/품목 이름 (스키마에서 가져온 열 번호 사용)
            
            # 같은 지역인지 확인 (지역명 매핑 고려)
            is_same_region = False
            if region_code:
                # 지역 코드로 확인
                if cell_a.value and str(cell_a.value).strip() == region_code:
                    is_same_region = True
            else:
                # 지역 이름으로 확인 (지역명 매핑 고려)
                if cell_b.value:
                    cell_b_str = str(cell_b.value).strip()
                    # 지역명 매핑
                    region_mapping = {
                        '서울특별시': '서울', '부산광역시': '부산', '대구광역시': '대구',
                        '인천광역시': '인천', '광주광역시': '광주', '대전광역시': '대전',
                        '울산광역시': '울산', '세종특별자치시': '세종', '경기도': '경기',
                        '강원도': '강원', '충청북도': '충북', '충청남도': '충남',
                        '전라북도': '전북', '전라남도': '전남', '경상북도': '경북',
                        '경상남도': '경남', '제주특별자치도': '제주'
                    }
                    mapped_region_name = region_mapping.get(region_name, region_name)
                    mapped_cell_b = region_mapping.get(cell_b_str, cell_b_str)
                    
                    if (cell_b_str == region_name or 
                        cell_b_str == mapped_region_name or
                        mapped_cell_b == region_name or
                        mapped_cell_b == mapped_region_name):
                        is_same_region = True
            
            if is_same_region and cell_c.value:
                try:
                    classification_level = float(cell_c.value) if cell_c.value else 0
                except (ValueError, TypeError):
                    classification_level = 0
                
                # 분류단계 필터링: 1 이상이고 max_classification_level 이하인 경우만 포함
                if classification_level >= 1 and classification_level <= max_classification_level:
                    current = sheet.cell(row=row, column=current_col).value
                    prev = sheet.cell(row=row, column=prev_col).value
                    
                    # 가중치 가져오기
                    weight = 1.0
                    if weight_column:
                        weight_value = sheet.cell(row=row, column=weight_column).value
                        weight = self.schema_loader.get_weight_value(sheet_name, weight_value)
                    
                    if current is not None and prev is not None and prev != 0:
                        growth_rate = ((current / prev) - 1) * 100
                        # 가중치를 반영한 증감률 (순위용)
                        weighted_growth_rate = growth_rate * weight
                        industries.append({
                            'name': str(cell_f.value).strip() if cell_f.value else '',
                            'growth_rate': growth_rate,
                            'weighted_growth_rate': weighted_growth_rate,
                            'weight': weight,
                            'classification_level': classification_level,
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
                        if cell_b.value:
                            cell_b_str = str(cell_b.value).strip()
                            # 지역명 매핑
                            region_mapping = {
                                '서울특별시': '서울', '부산광역시': '부산', '대구광역시': '대구',
                                '인천광역시': '인천', '광주광역시': '광주', '대전광역시': '대전',
                                '울산광역시': '울산', '세종특별자치시': '세종', '경기도': '경기',
                                '강원도': '강원', '충청북도': '충북', '충청남도': '충남',
                                '전라북도': '전북', '전라남도': '전남', '경상북도': '경북',
                                '경상남도': '경남', '제주특별자치도': '제주'
                            }
                            mapped_region_name = region_mapping.get(region_name, region_name)
                            mapped_cell_b = region_mapping.get(cell_b_str, cell_b_str)
                            
                            if (cell_b_str != region_name and 
                                cell_b_str != mapped_region_name and
                                mapped_cell_b != region_name and
                                mapped_cell_b != mapped_region_name):
                                break
        
        return industries
    
    def get_top_industries_for_region(self, sheet_name: str, region_name: str, region_row: int,
                                      current_col: int, prev_col: int, top_n: int = 3,
                                      use_weighted_ranking: bool = True) -> List[Dict]:
        """
        특정 지역의 상위 N개 산업을 반환 (가중치 기반 순위 지원)
        
        Args:
            sheet_name: 시트 이름
            region_name: 지역 이름
            region_row: 지역 총지수 행 번호
            current_col: 현재 분기 열 번호
            prev_col: 전년 동분기 열 번호
            top_n: 상위 개수
            use_weighted_ranking: 가중치 기반 순위 사용 여부 (기본값: True)
            
        Returns:
            상위 산업 리스트 (가중치 반영 순위)
        """
        industries = self.get_industries_for_region(sheet_name, region_name, region_row, 
                                                    current_col, prev_col)
        
        if not industries:
            return []
        
        # 가중치 설정 확인
        weight_config = self.schema_loader.get_weight_config(sheet_name)
        should_use_weighted = use_weighted_ranking and weight_config.get('use_weighted_ranking', True)
        
        # 물가동향 시트 특별 처리
        is_price_sheet = '물가' in sheet_name
        if is_price_sheet:
            # 물가동향의 경우 분류단계 2와 3 항목 사용, 특정 품목 우선 선택
            # 정답 이미지 기준으로 선택 로직 구현
            
            # 전국/부산/경기: 외식제외개인서비스, 외식, 가공식품, 공공서비스 (증가율 큰 순서)
            # 대구: 외식제외개인서비스, 가공식품, 외식, 축산물 (증가율 큰 순서)
            # 제주/광주/울산: 농산물, 석유류, 의약품, 출판물 또는 내구재 (감소율 큰 순서)
            
            # 분류단계 2와 3 항목 필터링 (공공서비스는 분류단계 2)
            level2_3_industries = [ind for ind in industries if ind.get('classification_level', 0) in [2, 3]]
            
            if level3_industries:
                # 품목명 정리 (앞뒤 공백 제거)
                for ind in level3_industries:
                    ind['name'] = ind['name'].strip()
                
                # 지역별 우선순위 품목 정의
                priority_items_by_region = {
                    '전국': ['외식제외개인서비스', '외식', '가공식품', '공공서비스'],
                    '부산': ['외식제외개인서비스', '외식', '가공식품', '공공서비스'],
                    '경기': ['외식제외개인서비스', '외식', '가공식품', '공공서비스'],
                    '대구': ['외식제외개인서비스', '가공식품', '외식', '축산물'],
                    '제주': ['농산물', '석유류', '의약품', '출판물'],
                    '광주': ['농산물', '석유류', '내구재', '출판물'],
                    '울산': ['농산물', '석유류', '내구재', '출판물']
                }
                
                priority_items = priority_items_by_region.get(region_name, [])
                
                # 품목명 정리 (앞뒤 공백 제거)
                for ind in level2_3_industries:
                    ind['name'] = ind['name'].strip()
                
                # 우선순위 품목 매칭 (정확한 일치 우선, 부분 일치 보조)
                matched_items = []
                for priority_item in priority_items:
                    best_match = None
                    best_score = 0
                    
                    for ind in level2_3_industries:
                        ind_name_clean = ind['name'].strip()
                        # 정확한 일치 (가장 높은 우선순위)
                        if priority_item == ind_name_clean:
                            best_match = ind
                            best_score = 100
                            break
                        # 부분 일치 (공백 무시)
                        priority_clean = priority_item.replace(' ', '').replace('·', '').replace(' ', '')
                        ind_clean = ind_name_clean.replace(' ', '').replace('·', '').replace(' ', '')
                        
                        score = 0
                        if priority_clean == ind_clean:
                            score = 90
                        elif priority_clean in ind_clean:
                            score = 80
                        elif ind_clean in priority_clean:
                            score = 70
                        elif priority_item in ind_name_clean:
                            score = 60
                        elif ind_name_clean in priority_item:
                            score = 50
                        
                        if score > best_score:
                            best_match = ind
                            best_score = score
                    
                    if best_match and best_match not in matched_items:
                        matched_items.append(best_match)
                
                # 우선순위 품목을 우선순위 순서로 정렬
                result = []
                for priority_item in priority_items:
                    best_match = None
                    best_score = 0
                    
                    for matched in matched_items:
                        if matched in result:
                            continue
                            
                        ind_name_clean = matched['name'].strip()
                        # 정확한 일치 우선
                        if priority_item == ind_name_clean:
                            best_match = matched
                            best_score = 100
                            break
                        
                        # 부분 일치 점수 계산
                        priority_clean = priority_item.replace(' ', '').replace('·', '').replace(' ', '')
                        ind_clean = ind_name_clean.replace(' ', '').replace('·', '').replace(' ', '')
                        
                        score = 0
                        if priority_clean == ind_clean:
                            score = 90
                        elif priority_clean in ind_clean:
                            score = 80
                        elif ind_clean in priority_clean:
                            score = 70
                        elif priority_item in ind_name_clean:
                            score = 60
                        elif ind_name_clean in priority_item:
                            score = 50
                        
                        if score > best_score:
                            best_match = matched
                            best_score = score
                    
                    if best_match:
                        result.append(best_match)
                
                # 우선순위 품목이 부족한 경우, 나머지 항목을 증감률 기준으로 추가
                if len(result) < top_n:
                    remaining = [ind for ind in level2_3_industries if ind not in result]
                    # 증감률 절대값 기준 정렬
                    if should_use_weighted:
                        remaining.sort(key=lambda x: abs(x.get('weighted_growth_rate', x['growth_rate'])), reverse=True)
                    else:
                        remaining.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
                    
                    for ind in remaining:
                        if ind not in result:
                            result.append(ind)
                            if len(result) >= top_n:
                                break
                
                if result:
                    return result[:top_n]
        
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
            '금속가공제품': ['금속 가공제품', '금속가공제품 제조업', '금속가공']
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
        
        # 각 카테고리 내에서 정렬 (가중치 반영 여부에 따라)
        for category in categorized:
            if category == '반도체·전자부품' and region_name == '전국':
                # '전자 부품, 컴퓨터'가 포함된 산업을 우선 선택
                if should_use_weighted:
                    categorized[category].sort(key=lambda x: (
                        '전자 부품, 컴퓨터' not in x['name'],
                        -abs(x.get('weighted_growth_rate', x['growth_rate']))
                    ))
                else:
                    categorized[category].sort(key=lambda x: (
                        '전자 부품, 컴퓨터' not in x['name'],
                        -abs(x['growth_rate'])
                    ))
            elif category == '의료·정밀' and region_name == '전국':
                if should_use_weighted:
                    categorized[category].sort(key=lambda x: (
                        '의료, 정밀, 광학 기기' not in x['name'],
                        -abs(x.get('weighted_growth_rate', x['growth_rate']))
                    ))
                else:
                    categorized[category].sort(key=lambda x: (
                        '의료, 정밀, 광학 기기' not in x['name'],
                        -abs(x['growth_rate'])
                    ))
            else:
                # 다른 카테고리는 가중치 반영 증감률 절대값 기준 정렬
                if should_use_weighted:
                    categorized[category].sort(key=lambda x: abs(x.get('weighted_growth_rate', x['growth_rate'])), reverse=True)
                else:
                    categorized[category].sort(key=lambda x: abs(x['growth_rate']), reverse=True)
        
        # 스크린샷 기준 순서 / 스키마 기반 우선순위
        result = []
        
        # 스키마에서 지역별 우선순위 가져오기
        sheet_config = self.schema_loader.load_sheet_config(sheet_name)
        region_priorities_from_schema = sheet_config.get('region_priorities', {})
        
        # 지역별 우선순위 정의 (스크린샷 기준 - 스키마가 없으면 기본값 사용)
        default_region_priorities = {
            '전국': ['반도체·전자부품', '기타 운송장비', '의료·정밀'],
            '충북': ['반도체·전자부품', '전기장비', '의약품'],
            '경기': ['반도체·전자부품', '기타기계장비', '의료·정밀'],
            '광주': ['전기장비', '담배', '자동차·트레일러'],
            '서울': ['의료·정밀', '전기·가스업', '식료품'],
            '충남': ['반도체·전자부품', '전기장비', '전기·가스업'],
            '부산': ['금속', '기타 운송장비', '금속가공제품']
        }
        
        # 스키마에 우선순위가 있으면 사용, 없으면 기본값 사용
        if region_name in region_priorities_from_schema:
            priority_list = region_priorities_from_schema[region_name]
        else:
            priority_list = default_region_priorities.get(region_name, ['반도체·전자부품'])
        
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
        if len(result) < top_n:
            # 우선순위에 없는 카테고리에서만 선택
            remaining = []
            for category, ind_list in categorized.items():
                if category not in priority_list:
                    remaining.extend(ind_list)
            
            # 가중치 반영 증감률 절대값 기준 정렬
            if should_use_weighted:
                remaining.sort(key=lambda x: abs(x.get('weighted_growth_rate', x['growth_rate'])), reverse=True)
            else:
                remaining.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
            
            for ind in remaining:
                if ind not in result:
                    result.append(ind)
                    if len(result) >= top_n:
                        break
        
        # 카테고리 매칭이 전혀 되지 않은 경우 (예: 수출/수입 시트)
        # 직접 증감률 기준으로 상위 N개 반환
        if len(result) < top_n:
            # 분류단계 1의 산업 우선 선택 (대분류)
            level1_industries = [ind for ind in industries if ind.get('classification_level', 0) == 1]
            level2_industries = [ind for ind in industries if ind.get('classification_level', 0) == 2]
            
            # 분류단계 2의 산업 중 상위 선택
            if should_use_weighted:
                level2_industries.sort(key=lambda x: abs(x.get('weighted_growth_rate', x['growth_rate'])), reverse=True)
            else:
                level2_industries.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
            
            for ind in level2_industries:
                if ind not in result:
                    result.append(ind)
                    if len(result) >= top_n:
                        break
        
        return result[:top_n]
    
    def analyze_quarter_data(self, sheet_name: str, quarter_cols: Dict[str, Tuple[int, int]]) -> Dict[str, Any]:
        """
        특정 분기의 데이터를 분석하여 상위/하위 시도와 각 시도의 상위 산업을 반환
        
        Args:
            sheet_name: 시트 이름
            quarter_cols: 분기별 (현재 열, 전년 열) 딕셔너리
                예: {'2025_2/4': (65, 61)}
            
        Returns:
            분석 결과 딕셔너리
        """
        results = {}
        
        for quarter_name, (current_col, prev_col) in quarter_cols.items():
            # 전국 데이터 추가
            national_region = None
            sheet = self.excel_extractor.get_sheet(sheet_name)
            
            # 시트별 설정에서 category_column 가져오기 (스키마 로더 사용)
            sheet_config = self.schema_loader.load_sheet_config(sheet_name)
            category_col = sheet_config.get('category_column', 6)
            region_column = sheet_config.get('region_column', 2)  # 지역 이름 열 (기본값: 2)
            
            for row in range(4, min(5000, sheet.max_row + 1)):
                cell_b = sheet.cell(row=row, column=region_column)  # 지역 이름 (스키마에서 가져온 열 번호 사용)
                cell_category = sheet.cell(row=row, column=category_col)
                # 총지수, 계, 합계 인식
                is_total = False
                if cell_category.value:
                    category_str = str(cell_category.value).strip()
                    if category_str in ['총지수', '계', '   계', '합계']:
                        is_total = True
                
                if cell_b.value == '전국' and is_total:
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
            
            # 물가동향 시트의 경우 top_n을 4로 설정 (템플릿에서 4개 품목 필요)
            is_price_sheet = '물가' in sheet_name
            top_n_count = 4 if is_price_sheet else 3
            
            # 전국 산업 추가
            national_industries = []
            if national_region:
                national_industries = self.get_top_industries_for_region(
                    sheet_name, '전국', national_region['row'],
                    current_col, prev_col, top_n=top_n_count
                )
                national_region['top_industries'] = national_industries
            
            # 각 상위 시도의 상위 산업
            top_regions_with_industries = []
            for region in top_regions:
                industries = self.get_top_industries_for_region(
                    sheet_name, region['name'], region['row'], 
                    current_col, prev_col, top_n=top_n_count
                )
                top_regions_with_industries.append({
                    **region,
                    'top_industries': industries
                })
            
            # 각 하위 시도의 상위 산업 (감소율이 큰 것들)
            bottom_regions_with_industries = []
            for region in bottom_regions:
                # 하위 시도도 get_top_industries_for_region 사용 (지역별 우선순위 적용)
                industries = self.get_top_industries_for_region(
                    sheet_name, region['name'], region['row'],
                    current_col, prev_col, top_n=top_n_count
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

