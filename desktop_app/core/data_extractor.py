# -*- coding: utf-8 -*-
"""
데이터 추출 래퍼
기존 RawDataExtractor를 활용하여 시도별 데이터를 추출합니다.
"""

import sys
import os
from pathlib import Path
from typing import Dict, List, Optional, Any

# 상위 디렉토리를 경로에 추가
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

try:
    from templates.raw_data_extractor import RawDataExtractor
    HAS_EXTRACTOR = True
except ImportError:
    HAS_EXTRACTOR = False
    RawDataExtractor = None


class DataExtractorWrapper:
    """기존 RawDataExtractor를 데스크톱 앱용으로 래핑"""
    
    # 시도 목록
    SIDO_LIST = [
        "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종",
        "경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"
    ]
    
    # 지표 매핑 (내부 키 -> 플레이스홀더 키)
    INDICATOR_MAPPING = {
        "manufacturing": "manufacturing",
        "service": "service",
        "retail": "retail",
        "construction": "construction",
        "export": "export",
        "import": "import",
        "price": "price",
        "employment": "employment",
        "migration": "migration",
    }
    
    def __init__(self, excel_path: str, year: int, quarter: int):
        """
        Args:
            excel_path: 기초자료 수집표 엑셀 파일 경로
            year: 기준 연도
            quarter: 기준 분기
        """
        self.excel_path = excel_path
        self.year = year
        self.quarter = quarter
        self.extractor = None
        
        if HAS_EXTRACTOR:
            self.extractor = RawDataExtractor(excel_path, year, quarter)
    
    def extract_all_sido_data(self) -> Dict[str, Dict]:
        """모든 시도의 주요지표 데이터 추출
        
        Returns:
            시도별 데이터 딕셔너리
            {
                "서울": {
                    "DATA_23_2Q4_manufacturing": -9.7,
                    "DATA_23_2Q4_service": 6.6,
                    ...
                },
                "부산": {...},
                ...
            }
        """
        all_data = {}
        
        for sido in self.SIDO_LIST:
            all_data[sido] = self.extract_sido_data(sido)
        
        return all_data
    
    def extract_sido_data(self, sido_name: str) -> Dict[str, Any]:
        """특정 시도의 주요지표 데이터 추출
        
        Args:
            sido_name: 시도명 (예: "서울", "부산")
            
        Returns:
            해당 시도의 지표 데이터 딕셔너리
        """
        data = {}
        
        if not HAS_EXTRACTOR or self.extractor is None:
            # 더미 데이터 반환
            return self._get_dummy_data(sido_name)
        
        try:
            # 기존 extract_regional_data 활용
            regional_data = self.extractor.extract_regional_data(sido_name)
            
            if regional_data:
                data = self._convert_to_placeholder_format(regional_data)
        except Exception as e:
            print(f"[DataExtractor] {sido_name} 데이터 추출 오류: {e}")
            data = self._get_dummy_data(sido_name)
        
        return data
    
    def _convert_to_placeholder_format(self, regional_data: Dict) -> Dict[str, Any]:
        """기존 데이터 형식을 플레이스홀더 형식으로 변환
        
        Args:
            regional_data: RawDataExtractor에서 추출한 데이터
            
        Returns:
            플레이스홀더 키 형식의 데이터
        """
        result = {}
        
        # 분기 키 매핑 (예: "23.2/4" -> "23_2Q4")
        quarter_mappings = {
            f"{self.year - 2}.{self.quarter}/4": f"{str(self.year - 2)[2:]}_{{q}}Q4",
            f"{self.year - 1}.{self.quarter}/4": f"{str(self.year - 1)[2:]}_{{q}}Q4",
            f"{self.year}.{self.quarter - 1}/4" if self.quarter > 1 else f"{self.year - 1}.4/4": f"{str(self.year)[2:]}_{self.quarter - 1 if self.quarter > 1 else 4}Q4",
            f"{self.year}.{self.quarter}/4": f"{str(self.year)[2:]}_{self.quarter}Q4",
        }
        
        # 각 지표별 데이터 변환
        indicators = [
            ("manufacturing", "광공업생산", "growth_rate"),
            ("service", "서비스업생산", "growth_rate"),
            ("retail", "소매판매", "growth_rate"),
            ("construction", "건설수주", "growth_rate"),
            ("export", "수출", "growth_rate"),
            ("import", "수입", "growth_rate"),
            ("price", "소비자물가", "growth_rate"),
            ("employment", "고용률", "change"),
            ("migration", "인구순이동", "value"),
        ]
        
        for ind_key, ind_name, value_key in indicators:
            section_data = regional_data.get(ind_name, {})
            
            # 각 분기별 데이터
            for orig_q, new_q in quarter_mappings.items():
                q_key = new_q.format(q=self.quarter)
                placeholder_key = f"DATA_{q_key}_{ind_key}"
                
                # 분기별 데이터 찾기
                quarterly_data = section_data.get("quarterly_data", {})
                value = quarterly_data.get(orig_q, {}).get(value_key)
                
                if value is None:
                    value = section_data.get(value_key)
                
                result[placeholder_key] = value if value is not None else "N/A"
        
        return result
    
    def _get_dummy_data(self, sido_name: str) -> Dict[str, Any]:
        """더미 데이터 생성 (테스트용)
        
        Args:
            sido_name: 시도명
            
        Returns:
            더미 데이터 딕셔너리
        """
        import random
        
        quarters = ["23_2Q4", "24_2Q4", "25_1Q4", "25_2Q4"]
        indicators = [
            "manufacturing", "service", "retail", "construction",
            "export", "import", "price", "employment", "migration"
        ]
        
        data = {}
        for q in quarters:
            for ind in indicators:
                key = f"DATA_{q}_{ind}"
                
                # 지표별 적절한 범위의 더미 값 생성
                if ind == "employment":
                    data[key] = round(random.uniform(-1.5, 1.5), 1)
                elif ind == "migration":
                    data[key] = round(random.uniform(-20, 20), 1)
                elif ind in ["construction", "export", "import"]:
                    data[key] = round(random.uniform(-50, 100), 1)
                else:
                    data[key] = round(random.uniform(-10, 10), 1)
        
        return data
    
    def get_report_info(self) -> Dict[str, Any]:
        """보도자료 기본 정보 반환"""
        return {
            "year": self.year,
            "quarter": self.quarter,
            "organization": "국가데이터처",
            "department": "경제통계심의관",
        }


# 테스트 코드
if __name__ == "__main__":
    # 테스트
    wrapper = DataExtractorWrapper(
        "/path/to/test.xlsx",
        2025,
        3
    )
    
    # 더미 데이터 테스트
    seoul_data = wrapper.extract_sido_data("서울")
    print("서울 데이터 샘플:")
    for key, value in list(seoul_data.items())[:5]:
        print(f"  {key}: {value}")
