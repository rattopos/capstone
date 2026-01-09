"""
템플릿 관리 모듈
HTML 템플릿 파일 로드 및 마커 파싱 기능 제공
"""

import re
from pathlib import Path
from typing import List, Dict, Tuple, Optional


class TemplateManager:
    """HTML 템플릿을 관리하고 마커를 파싱하는 클래스"""
    
    # 마커 패턴: {시트명:셀주소} 또는 {시트명:셀주소:계산식} 또는 {시트명:헤더기반키}
    # CSS 중괄호와 구분: 시트명 다음에 콜론(:)이 있고, 셀주소는 영문+숫자 형식 또는 헤더 기반 키
    # 시트명: 한글, 영문, 숫자, 공백, 언더스코어 등 가능
    # 셀주소: 영문+숫자 형식 (A1, B5, AA100 등) 또는 범위 (A1:A5)
    # 헤더 기반 키: 한글, 영문, 숫자, 언더스코어 조합 (예: 전국_증감률, 서울_2023_3_4)
    MARKER_PATTERN = re.compile(r'\{([^:{}]+):([^:}]+)(?::([^}]+))?\}')
    
    def __init__(self, template_path: Optional[str] = None):
        """
        템플릿 매니저 초기화
        
        Args:
            template_path: HTML 템플릿 파일 경로 (None이면 메모리에서만 작업)
        """
        self.template_path = Path(template_path) if template_path else None
        self.template_content = ""
        self.markers = []
        
    def load_template(self) -> str:
        """
        HTML 템플릿 파일을 로드합니다.
        
        Returns:
            템플릿 내용 문자열
            
        Raises:
            ValueError: template_path가 None일 때
            FileNotFoundError: 템플릿 파일이 존재하지 않을 때
            IOError: 파일 읽기 실패 시
        """
        if self.template_path is None:
            raise ValueError("template_path가 None입니다. template_content를 직접 설정하거나 template_path를 제공해야 합니다.")
        
        if not self.template_path.exists():
            raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {self.template_path}")
        
        try:
            with open(self.template_path, 'r', encoding='utf-8') as f:
                self.template_content = f.read()
            return self.template_content
        except IOError as e:
            raise IOError(f"템플릿 파일 읽기 실패: {e}")
    
    def extract_markers(self) -> List[Dict[str, str]]:
        """
        템플릿에서 모든 마커를 추출합니다.
        
        Returns:
            마커 정보 딕셔너리 리스트. 각 딕셔너리는 다음 키를 포함:
            - 'full_match': 전체 마커 문자열 (예: '{시트1:A1:sum}')
            - 'sheet_name': 시트명
            - 'cell_address': 셀 주소 (또는 범위)
            - 'operation': 계산식 (선택적, 없으면 None)
        """
        if not self.template_content and self.template_path:
            self.load_template()
        
        self.markers = []
        matches = self.MARKER_PATTERN.finditer(self.template_content)
        
        for match in matches:
            sheet_name = match.group(1).strip()
            cell_address = match.group(2).strip()
            operation = match.group(3).strip() if match.group(3) else None
            
            marker_info = {
                'full_match': match.group(0),
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'operation': operation
            }
            
            # 중복 제거 (동일한 마커가 여러 번 나와도 한 번만 추가)
            if marker_info not in self.markers:
                self.markers.append(marker_info)
        
        return self.markers
    
    def parse_markers(self) -> List[Dict[str, str]]:
        """
        extract_markers의 별칭 (하위 호환성)
        """
        return self.extract_markers()
    
    def validate_template(self, required_markers: List[str] = None) -> Tuple[bool, List[str]]:
        """
        템플릿의 유효성을 검증합니다.
        
        Args:
            required_markers: 필수 마커 리스트 (전체 마커 문자열 형식)
            
        Returns:
            (유효성 여부, 누락된 마커 리스트) 튜플
        """
        if not self.template_content and self.template_path:
            self.load_template()
        
        extracted = self.extract_markers()
        extracted_full_matches = [m['full_match'] for m in extracted]
        
        if required_markers is None:
            # 필수 마커가 지정되지 않으면 모든 마커가 존재하는지만 확인
            return True, []
        
        missing_markers = [
            marker for marker in required_markers 
            if marker not in extracted_full_matches
        ]
        
        is_valid = len(missing_markers) == 0
        return is_valid, missing_markers
    
    def get_template_content(self) -> str:
        """
        현재 로드된 템플릿 내용을 반환합니다.
        
        Returns:
            템플릿 내용 문자열
        """
        if not self.template_content and self.template_path:
            self.load_template()
        return self.template_content
    
    def replace_marker(self, marker: str, value: str) -> str:
        """
        템플릿에서 특정 마커를 값으로 치환합니다.
        
        Args:
            marker: 치환할 마커 (전체 문자열, 예: '{시트1:A1}')
            value: 치환할 값
            
        Returns:
            치환된 템플릿 내용
        """
        if not self.template_content and self.template_path:
            self.load_template()
        
        # 정확히 일치하는 마커만 치환 (부분 일치 방지)
        self.template_content = self.template_content.replace(marker, str(value))
        return self.template_content

