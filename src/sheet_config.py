"""
시트별 설정 관리 모듈
각 시트별로 다른 설정(기준 분기, 표 헤더, 산업 우선순위 등)을 관리
"""

from typing import Dict, List, Tuple, Optional
from .config import Config


class SheetConfig:
    """시트별 설정을 관리하는 클래스"""
    
    # 시트별 기본 설정
    # 각 시트는 고유한 기준 분기, 표 헤더 범위, 산업 우선순위 등을 가질 수 있음
    SHEET_CONFIGS: Dict[str, Dict] = {
        '광공업생산': {
            'default_year': 2023,
            'default_quarter': 1,
            'table_quarters': [
                ('2021', 2), ('2021', 3), ('2021', 4),
                ('2022', 1), ('2022', 2), ('2022', 3), ('2022', 4),
                ('2023', 1)
            ],
            'table_headers': [
                '2021 Q2', '2021 Q3', '2021 Q4',
                '2022 Q1', '2022 Q2', '2022 Q3', '2022 Q4',
                '2023 Q1p'
            ],
            'document_title': 'I. 부문별 지역경제동향',
            'section_title': '1. 생산 동향',
            'subsection_title': '가. 광공업생산',
            'region_priorities': {
                '전국': ['반도체·전자부품', '화학제품', '금속'],
                '강원': ['전기·가스업', '의료·정밀', '음료'],
                '대구': ['기타기계장비', '자동차·트레일러', '전기장비'],
                '인천': ['자동차·트레일러', '의약품', '기타기계장비'],
                '경기': ['반도체·전자부품', '화학제품', '기타기계장비'],
                '서울': ['화학제품', '기타기계장비', '의류·모피'],
                '충북': ['반도체·전자부품', '화학제품', '식료품']
            }
        }
        # 다른 시트 설정을 여기에 추가할 수 있음
        # 예:
        # '건설기성액': {
        #     'default_year': 2023,
        #     'default_quarter': 1,
        #     ...
        # }
    }
    
    def __init__(self, sheet_name: str, year: Optional[int] = None, quarter: Optional[int] = None):
        """
        시트별 설정 초기화
        
        Args:
            sheet_name: 시트 이름
            year: 분석할 연도 (None이면 시트 기본값 사용)
            quarter: 분석할 분기 (None이면 시트 기본값 사용)
        """
        self.sheet_name = sheet_name
        
        # 시트 설정 가져오기 (없으면 기본값 사용)
        if sheet_name in self.SHEET_CONFIGS:
            self.config = self.SHEET_CONFIGS[sheet_name].copy()
        else:
            # 기본 설정
            self.config = {
                'default_year': 2023,
                'default_quarter': 1,
                'table_quarters': [],
                'table_headers': [],
                'document_title': '부문별 지역경제동향',
                'section_title': '',
                'subsection_title': '',
                'region_priorities': {}
            }
        
        # 연도/분기 설정
        self.year = year if year is not None else self.config['default_year']
        self.quarter = quarter if quarter is not None else self.config['default_quarter']
        
        # Config 객체 생성
        self.config_obj = Config(self.year, self.quarter)
    
    def get_config(self) -> Config:
        """Config 객체 반환"""
        return self.config_obj
    
    def get_table_headers(self) -> List[str]:
        """표 헤더 리스트 반환"""
        return self.config.get('table_headers', [])
    
    def get_table_quarters(self) -> List[Tuple[str, int]]:
        """표 분기 리스트 반환 (연도, 분기) 튜플"""
        return self.config.get('table_quarters', [])
    
    def get_region_priorities(self, region_name: str) -> List[str]:
        """지역별 산업 우선순위 반환"""
        priorities = self.config.get('region_priorities', {})
        return priorities.get(region_name, ['반도체·전자부품'])
    
    def get_document_title(self) -> str:
        """문서 제목 반환"""
        return self.config.get('document_title', '부문별 지역경제동향')
    
    def get_section_title(self) -> str:
        """섹션 제목 반환"""
        return self.config.get('section_title', '')
    
    def get_subsection_title(self) -> str:
        """서브섹션 제목 반환"""
        return self.config.get('subsection_title', '')
    
    @classmethod
    def register_sheet_config(cls, sheet_name: str, config: Dict):
        """
        새로운 시트 설정을 등록합니다.
        
        Args:
            sheet_name: 시트 이름
            config: 시트 설정 딕셔너리
        """
        cls.SHEET_CONFIGS[sheet_name] = config
    
    @classmethod
    def get_available_sheets(cls) -> List[str]:
        """등록된 시트 목록 반환"""
        return list(cls.SHEET_CONFIGS.keys())

