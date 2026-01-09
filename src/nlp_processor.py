"""
자연어 처리 모듈
증감률에 따라 적절한 한국어 표현 생성
"""

from typing import Optional


class NLPProcessor:
    """증감률 기반 자연어 처리를 수행하는 클래스"""
    
    def __init__(self):
        """NLP 프로세서 초기화"""
        pass
    
    def determine_trend(self, growth_rate: float) -> str:
        """
        증감률에 따라 증감 방향을 판단합니다.
        
        Args:
            growth_rate: 증감률 (퍼센트)
            
        Returns:
            "증가", "감소", 또는 "변동 없음"
        """
        if growth_rate > 0:
            return "증가"
        elif growth_rate < 0:
            return "감소"
        else:
            return "변동 없음"
    
    def format_trend_text(self, growth_rate: float, context: str = "") -> str:
        """
        맥락에 맞는 증감 텍스트를 생성합니다.
        
        Args:
            growth_rate: 증감률 (퍼센트)
            context: 맥락 정보 (현재는 사용하지 않지만 향후 확장용)
            
        Returns:
            증감 방향 텍스트 ("증가", "감소", "변동 없음")
        """
        return self.determine_trend(growth_rate)
    
    def get_trend_with_magnitude(self, growth_rate: float) -> str:
        """
        증감률의 크기에 따른 표현을 반환합니다 (향후 확장용).
        
        Args:
            growth_rate: 증감률 (퍼센트)
            
        Returns:
            증감 방향 텍스트 (현재는 단순 판단만 수행)
        """
        # 향후 확장: "소폭 증가", "급격히 감소" 등
        return self.determine_trend(growth_rate)

