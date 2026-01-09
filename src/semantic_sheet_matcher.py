"""
의미 기반 시트 매칭 모듈
키워드 기반 자연어 처리를 통한 시트 자동 매칭
"""

import re
from typing import Dict, List, Optional, Tuple
from difflib import SequenceMatcher


class SemanticSheetMatcher:
    """키워드 기반 시트 매칭 클래스"""
    
    # 시트명 키워드 매핑 (의미 기반)
    SHEET_KEYWORD_MAPPING = {
        # 경제 지표 관련
        '경제지표': ['경제', '지표', 'gdp', '생산', '소비', '투자'],
        '생산': ['생산', '제조', 'manufacturing', 'production'],
        '소비': ['소비', '소매', 'consumption', 'retail'],
        '건설': ['건설', 'construction', '공정'],
        '광업': ['광업', 'mining'],
        '제조업': ['제조', 'manufacturing'],
        '서비스업': ['서비스', 'service'],
        '소매업': ['소매', 'retail'],
        
        # 지역 관련
        '지역': ['지역', 'region', '시도', '시군구'],
        '전국': ['전국', 'national', '전체'],
        '서울': ['서울', 'seoul'],
        '부산': ['부산', 'busan'],
        
        # 산업 관련
        '산업': ['산업', 'industry', '업종'],
        '업종': ['업종', 'industry'],
        
        # 기타
        '완료체크': ['완료', '체크', 'check'],
        '요약': ['요약', 'summary'],
    }
    
    def __init__(self, excel_extractor):
        """
        의미 기반 매퍼 초기화
        
        Args:
            excel_extractor: ExcelExtractor 인스턴스
        """
        self.excel_extractor = excel_extractor
        self.sheet_cache = {}  # 시트 정보 캐시
        self._build_sheet_keyword_index()
    
    def _build_sheet_keyword_index(self):
        """사용 가능한 시트들의 키워드 인덱스 구축"""
        if not self.excel_extractor.workbook:
            self.excel_extractor.load_workbook()
        
        available_sheets = self.excel_extractor.get_sheet_names()
        self.sheet_keywords = {}
        
        for sheet_name in available_sheets:
            # 시트명에서 키워드 추출
            keywords = self._extract_keywords_from_sheet_name(sheet_name)
            self.sheet_keywords[sheet_name] = keywords
    
    def _extract_keywords_from_sheet_name(self, sheet_name: str) -> List[str]:
        """시트명에서 의미 있는 키워드 추출"""
        keywords = []
        
        # 한글 단어 추출
        korean_words = re.findall(r'[가-힣]+', sheet_name)
        keywords.extend(korean_words)
        
        # 영문 단어 추출
        english_words = re.findall(r'[a-zA-Z]+', sheet_name.lower())
        keywords.extend(english_words)
        
        # 숫자 제거 (의미 없는 경우)
        # 하지만 "2023년", "1분기" 같은 의미 있는 숫자는 유지
        
        return [kw.lower() for kw in keywords if len(kw) > 1]
    
    def find_sheet_by_semantic_keywords(
        self, 
        target_keywords: str,
        similarity_threshold: float = 0.3
    ) -> Optional[str]:
        """
        키워드 기반으로 시트를 찾습니다.
        
        Args:
            target_keywords: 찾고자 하는 키워드 문자열 (예: "경제지표", "생산", "전국_증감률")
            similarity_threshold: 유사도 임계값 (0.0 ~ 1.0)
            
        Returns:
            실제 시트명 또는 None
        """
        if not self.excel_extractor.workbook:
            self.excel_extractor.load_workbook()
        
        available_sheets = self.excel_extractor.get_sheet_names()
        
        # 타겟 키워드 추출
        target_keywords_list = self._extract_keywords_from_text(target_keywords)
        
        # 키워드 매핑 확장 (의미 기반)
        expanded_keywords = self._expand_keywords_by_semantic_mapping(target_keywords_list)
        
        best_match = None
        best_score = 0.0
        
        for sheet_name in available_sheets:
            sheet_keywords = self.sheet_keywords.get(sheet_name, [])
            
            # 키워드 매칭 점수 계산
            score = self._calculate_keyword_match_score(expanded_keywords, sheet_keywords, sheet_name)
            
            if score > best_score and score >= similarity_threshold:
                best_score = score
                best_match = sheet_name
        
        return best_match
    
    def _extract_keywords_from_text(self, text: str) -> List[str]:
        """텍스트에서 키워드 추출"""
        keywords = []
        
        # 한글 단어 추출
        korean_words = re.findall(r'[가-힣]+', text)
        keywords.extend(korean_words)
        
        # 영문 단어 추출
        english_words = re.findall(r'[a-zA-Z]+', text.lower())
        keywords.extend(english_words)
        
        # 언더스코어로 구분된 단어들도 분리
        parts = text.replace('_', ' ').replace('-', ' ').split()
        for part in parts:
            if part and len(part) > 1:
                keywords.append(part.lower())
        
        # 중복 제거 및 정규화
        keywords = list(set([kw.lower() for kw in keywords if len(kw) > 1]))
        
        return keywords
    
    def _expand_keywords_by_semantic_mapping(self, keywords: List[str]) -> List[str]:
        """의미 기반 키워드 매핑을 통해 키워드 확장"""
        expanded = set(keywords)
        
        for keyword in keywords:
            keyword_lower = keyword.lower()
            
            # 키워드 매핑에서 찾기
            for semantic_key, related_keywords in self.SHEET_KEYWORD_MAPPING.items():
                if keyword_lower in related_keywords or any(kw in keyword_lower for kw in related_keywords):
                    # 관련 키워드 모두 추가
                    expanded.update(related_keywords)
                    expanded.add(semantic_key)
        
        return list(expanded)
    
    def _calculate_keyword_match_score(
        self, 
        target_keywords: List[str], 
        sheet_keywords: List[str],
        sheet_name: str
    ) -> float:
        """키워드 매칭 점수 계산"""
        if not target_keywords or not sheet_keywords:
            return 0.0
        
        # 정확한 키워드 매칭
        exact_matches = set(target_keywords) & set(sheet_keywords)
        if exact_matches:
            return 1.0
        
        # 부분 매칭 (키워드가 시트명에 포함되는지)
        sheet_name_lower = sheet_name.lower()
        partial_matches = sum(1 for kw in target_keywords if kw in sheet_name_lower)
        if partial_matches > 0:
            return 0.7 + (partial_matches / len(target_keywords)) * 0.2
        
        # 유사도 기반 매칭
        max_similarity = 0.0
        for target_kw in target_keywords:
            for sheet_kw in sheet_keywords:
                similarity = SequenceMatcher(None, target_kw, sheet_kw).ratio()
                max_similarity = max(max_similarity, similarity)
        
        return max_similarity * 0.5
    
    def find_sheets_by_context(self, context_text: str) -> List[Tuple[str, float]]:
        """
        컨텍스트 텍스트를 기반으로 관련 시트들을 찾습니다.
        
        Args:
            context_text: 컨텍스트 텍스트 (예: "경제지표 전국 증감률")
            
        Returns:
            (시트명, 점수) 튜플 리스트 (점수 순으로 정렬)
        """
        if not self.excel_extractor.workbook:
            self.excel_extractor.load_workbook()
        
        available_sheets = self.excel_extractor.get_sheet_names()
        context_keywords = self._extract_keywords_from_text(context_text)
        expanded_keywords = self._expand_keywords_by_semantic_mapping(context_keywords)
        
        results = []
        for sheet_name in available_sheets:
            sheet_keywords = self.sheet_keywords.get(sheet_name, [])
            score = self._calculate_keyword_match_score(expanded_keywords, sheet_keywords, sheet_name)
            
            if score > 0.0:
                results.append((sheet_name, score))
        
        # 점수 순으로 정렬
        results.sort(key=lambda x: x[1], reverse=True)
        
        return results
    
    def infer_sheet_from_marker(self, marker_text: str) -> Optional[str]:
        """
        마커 텍스트에서 시트를 추론합니다.
        
        Args:
            marker_text: 마커 텍스트 (예: "{경제지표:전국_증감률}", "경제지표:전국_증감률")
            
        Returns:
            추론된 시트명 또는 None
        """
        # 마커 형식에서 시트명 부분 추출
        # {시트명:키} 또는 시트명:키 형식
        match = re.match(r'\{?([^:{}]+):', marker_text)
        if match:
            sheet_part = match.group(1).strip()
            
            # 키워드 기반으로 시트 찾기
            return self.find_sheet_by_semantic_keywords(sheet_part)
        
        return None

