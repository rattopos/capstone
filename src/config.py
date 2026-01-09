"""
설정 관리 모듈
연도/분기 설정 관리 및 열 번호 계산
"""

from typing import Tuple, Optional


class Config:
    """연도/분기 설정 및 열 번호 계산을 관리하는 클래스"""
    
    # 기준: 2023년 3/4분기 = Col 58 (BD)
    BASE_YEAR = 2023
    BASE_QUARTER = 3
    BASE_COL = 58
    
    def __init__(self, year: int, quarter: int):
        """
        설정 초기화
        
        Args:
            year: 분석할 연도 (예: 2025)
            quarter: 분석할 분기 (1-4)
            
        Raises:
            ValueError: 잘못된 연도 또는 분기 값일 때
        """
        if not isinstance(year, int) or year < 2020 or year > 2100:
            raise ValueError(f"잘못된 연도: {year}. 2020-2100 사이의 값을 입력하세요.")
        
        if not isinstance(quarter, int) or quarter < 1 or quarter > 4:
            raise ValueError(f"잘못된 분기: {quarter}. 1-4 사이의 값을 입력하세요.")
        
        self.year = year
        self.quarter = quarter
    
    def get_current_column(self) -> int:
        """
        현재 분기의 열 번호를 계산합니다.
        
        Returns:
            현재 분기 열 번호
            
        계산 공식:
        base_col = 58 + (year - 2023) * 4 + (quarter - 3)
        """
        year_diff = self.year - self.BASE_YEAR
        quarter_diff = self.quarter - self.BASE_QUARTER
        
        current_col = self.BASE_COL + (year_diff * 4) + quarter_diff
        return current_col
    
    def get_previous_year_column(self) -> int:
        """
        전년 동분기의 열 번호를 계산합니다.
        
        Returns:
            전년 동분기 열 번호
        """
        prev_year = self.year - 1
        prev_config = Config(prev_year, self.quarter)
        return prev_config.get_current_column()
    
    def get_column_pair(self) -> Tuple[int, int]:
        """
        현재 분기와 전년 동분기의 열 번호 쌍을 반환합니다.
        
        Returns:
            (현재 분기 열 번호, 전년 동분기 열 번호) 튜플
        """
        return (self.get_current_column(), self.get_previous_year_column())
    
    def get_quarter_name(self) -> str:
        """
        분기 이름을 반환합니다 (예: "2025_2/4").
        
        Returns:
            분기 이름 문자열
        """
        return f"{self.year}_{self.quarter}/4"
    
    def get_quarter_key(self) -> str:
        """
        분기 키를 반환합니다 (예: "2025_2분기").
        
        Returns:
            분기 키 문자열
        """
        return f"{self.year}_{self.quarter}분기"
    
    def __repr__(self) -> str:
        """문자열 표현"""
        return f"Config(year={self.year}, quarter={self.quarter})"

