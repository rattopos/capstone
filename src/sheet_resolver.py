"""
시트명 자동 해석 모듈
템플릿 내용을 기반으로 올바른 시트명을 찾는 기능
"""

import re
from typing import Optional, List


class SheetResolver:
    """템플릿 내용 기반으로 올바른 시트명을 찾는 클래스"""
    
    # 키워드 -> 시트명 매핑
    KEYWORD_SHEET_MAPPING = {
        # 생산 관련
        '광공업생산': '광공업생산',
        '광공업': '광공업생산',
        '제조업': '광공업생산',
        '광업': '광공업생산',
        
        '서비스업생산': '서비스업생산',
        '서비스업': '서비스업생산',
        '서비스': '서비스업생산',
        
        # 소비 관련
        '소매판매': '소비(소매, 추가)',
        '소매': '소비(소매, 추가)',
        '소비': '소비(소매, 추가)',
        '판매': '소비(소매, 추가)',
        '백화점': '소비(소매, 추가)',
        '대형마트': '소비(소매, 추가)',
        '면세점': '소비(소매, 추가)',
        
        # 건설 관련
        '건설수주': '건설 (공표자료)',
        '건설': '건설 (공표자료)',
        '수주': '건설 (공표자료)',
        '건축': '건설 (공표자료)',
        '토목': '건설 (공표자료)',
        '공정': '건설 (공표자료)',
        
        # 수출입 관련
        '수출': '수출',
        '수입': '수입',
        '무역': '수출',  # 기본적으로 수출
        
        # 고용 관련
        '고용': '고용',
        '취업자': '고용',
        '취업': '고용',
        '고용률': '고용률',
        '실업': '실업자 수',
        '실업자': '실업자 수',
        
        # 물가 관련
        '물가': '지출목적별 물가',
        '소비자물가': '지출목적별 물가',
        '지출목적': '지출목적별 물가',
        '품목성질': '품목성질별 물가',
        
        # 인구이동 관련
        '인구이동': '연령별 인구이동',
        '순인구이동': '연령별 인구이동',
        '시도간': '시도 간 이동',
        '시도 간': '시도 간 이동',
        '시군구': '시군구인구이동',
    }
    
    def __init__(self, available_sheets: List[str]):
        """
        시트 해석기 초기화
        
        Args:
            available_sheets: 사용 가능한 시트명 리스트
        """
        self.available_sheets = available_sheets
        self._normalized_sheets = {self._normalize(s): s for s in available_sheets}
    
    def resolve_sheet(self, marker_sheet_name: str, context_text: str = "") -> Optional[str]:
        """
        마커의 시트명을 실제 시트명으로 해석
        
        Args:
            marker_sheet_name: 마커에 있는 시트명 (예: "완료체크")
            context_text: 마커 주변의 텍스트 (선택적)
            
        Returns:
            실제 시트명 또는 None
        """
        # "완료체크"가 아니면 그대로 반환
        if marker_sheet_name != '완료체크':
            # 정확한 시트명인지 확인
            if marker_sheet_name in self.available_sheets:
                return marker_sheet_name
            
            # 정규화해서 찾기
            normalized = self._normalize(marker_sheet_name)
            if normalized in self._normalized_sheets:
                return self._normalized_sheets[normalized]
            
            return marker_sheet_name
        
        # "완료체크"인 경우 컨텍스트 기반으로 찾기
        if context_text:
            return self._find_sheet_from_context(context_text)
        
        return None
    
    def _find_sheet_from_context(self, context_text: str) -> Optional[str]:
        """
        컨텍스트 텍스트에서 시트명 찾기
        
        Args:
            context_text: 마커 주변 텍스트
            
        Returns:
            시트명 또는 None
        """
        context_lower = context_text.lower()
        
        # 키워드 매핑을 사용해서 찾기
        for keyword, sheet_name in self.KEYWORD_SHEET_MAPPING.items():
            if keyword in context_lower:
                # 시트가 실제로 존재하는지 확인
                if sheet_name in self.available_sheets:
                    return sheet_name
        
        # 직접 시트명이 포함되어 있는지 확인
        for sheet_name in self.available_sheets:
            # 시트명의 주요 단어가 컨텍스트에 있는지 확인
            sheet_words = self._extract_keywords(sheet_name)
            if any(word in context_lower for word in sheet_words if len(word) > 1):
                return sheet_name
        
        return None
    
    def _normalize(self, text: str) -> str:
        """텍스트 정규화 (공백 제거, 소문자 변환 등)"""
        if not text:
            return ""
        # 공백, 특수문자 제거
        normalized = re.sub(r'[^\w가-힣]', '', text.lower())
        return normalized
    
    def _extract_keywords(self, text: str) -> List[str]:
        """텍스트에서 키워드 추출"""
        # 한글 단어, 영문 단어 추출
        keywords = re.findall(r'[가-힣]+|[a-zA-Z]+', text)
        # 1자리 키워드 제외
        keywords = [kw for kw in keywords if len(kw) > 1]
        return keywords
    
    def resolve_marker_in_template(self, template_content: str, marker: str, 
                                   context_range: int = 200) -> Optional[str]:
        """
        템플릿 내용에서 마커의 컨텍스트를 추출하고 시트명 해석
        
        Args:
            template_content: 템플릿 전체 내용
            marker: 마커 문자열 (예: "{완료체크:Column1}")
            context_range: 마커 앞뒤로 확인할 문자 수
            
        Returns:
            해석된 시트명 또는 None
        """
        # 마커 위치 찾기
        marker_pos = template_content.find(marker)
        if marker_pos == -1:
            return None
        
        # 컨텍스트 추출
        start = max(0, marker_pos - context_range)
        end = min(len(template_content), marker_pos + len(marker) + context_range)
        context = template_content[start:end]
        
        # HTML 태그 제거
        context = re.sub(r'<[^>]+>', ' ', context)
        
        # 마커에서 시트명 추출
        match = re.search(r'\{([^:{}]+):', marker)
        if not match:
            return None
        
        marker_sheet_name = match.group(1).strip()
        
        # 시트명 해석
        return self.resolve_sheet(marker_sheet_name, context)

