"""
템플릿 관리 모듈
HTML 템플릿 파일 로드 및 마커 파싱 기능 제공
"""

import re
from pathlib import Path
from typing import List, Dict, Tuple


class TemplateManager:
    """HTML 템플릿을 관리하고 마커를 파싱하는 클래스"""
    
    # 마커 패턴: {시트명:셀주소} 또는 {시트명:셀주소:계산식} 또는 {시트명:헤더기반키}
    # CSS 중괄호와 구분: 시트명 다음에 콜론(:)이 있고, 셀주소는 영문+숫자 형식 또는 헤더 기반 키
    # 시트명: 한글, 영문, 숫자, 공백, 언더스코어 등 가능
    # 셀주소: 영문+숫자 형식 (A1, B5, AA100 등) 또는 범위 (A1:A5)
    # 헤더 기반 키: 한글, 영문, 숫자, 언더스코어 조합 (예: 전국_증감률, 서울_2023_3_4)
    MARKER_PATTERN = re.compile(r'\{([^:{}]+):([^:}]+)(?::([^}]+))?\}')
    
    def __init__(self, template_path: str):
        """
        템플릿 매니저 초기화
        
        Args:
            template_path: HTML 템플릿 파일 경로
        """
        self.template_path = Path(template_path)
        self.template_content = ""
        self.markers = []
        
    def load_template(self) -> str:
        """
        HTML 템플릿 파일을 로드합니다.
        
        Returns:
            템플릿 내용 문자열
            
        Raises:
            FileNotFoundError: 템플릿 파일이 존재하지 않을 때
            IOError: 파일 읽기 실패 시
        """
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
        CSS 스타일 블록과 스크립트 블록은 제외합니다.
        
        Returns:
            마커 정보 딕셔너리 리스트. 각 딕셔너리는 다음 키를 포함:
            - 'full_match': 전체 마커 문자열 (예: '{시트1:A1:sum}')
            - 'sheet_name': 시트명
            - 'cell_address': 셀 주소 (또는 범위)
            - 'operation': 계산식 (선택적, 없으면 None)
        """
        if not self.template_content:
            self.load_template()
        
        # CSS 스타일 블록과 스크립트 블록 제거
        # <style>...</style> 및 <script>...</script> 태그 내용 제외
        content_without_style_script = self.template_content
        
        # <style> 태그 내용 제거 (재귀적으로 중첩된 태그도 처리)
        style_pattern = re.compile(r'<style[^>]*>.*?</style>', re.DOTALL | re.IGNORECASE)
        content_without_style_script = style_pattern.sub('', content_without_style_script)
        
        # <script> 태그 내용 제거
        script_pattern = re.compile(r'<script[^>]*>.*?</script>', re.DOTALL | re.IGNORECASE)
        content_without_style_script = script_pattern.sub('', content_without_style_script)
        
        self.markers = []
        matches = self.MARKER_PATTERN.finditer(content_without_style_script)
        
        for match in matches:
            sheet_name = match.group(1).strip()
            cell_address = match.group(2).strip()
            operation = match.group(3).strip() if match.group(3) else None
            
            # CSS 속성명이나 일반적인 CSS 값이 아닌지 추가 검증
            # 시트명이 일반적인 CSS 속성명이 아닌지 확인
            css_properties = {
                'width', 'height', 'margin', 'padding', 'font-size', 'font-family',
                'background-color', 'color', 'border', 'display', 'position',
                'top', 'left', 'right', 'bottom', 'z-index', 'opacity', 'overflow',
                'text-align', 'line-height', 'font-weight', 'box-sizing', 'max-width'
            }
            
            # 시트명이 CSS 속성명이면 건너뛰기
            if sheet_name.lower() in css_properties:
                continue
            
            # 셀 주소가 CSS 값 형식(px, %, em 등으로 끝나는)이면 건너뛰기
            if re.match(r'^[\d.]+(px|%|em|rem|pt|vh|vw|cm|mm|in)$', cell_address, re.IGNORECASE):
                continue
            
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
    
    def validate_template(self, required_markers: List[str] = None) -> Tuple[bool, List[str]]:
        """
        템플릿의 유효성을 검증합니다.
        
        Args:
            required_markers: 필수 마커 리스트 (전체 마커 문자열 형식)
            
        Returns:
            (유효성 여부, 누락된 마커 리스트) 튜플
        """
        if not self.template_content:
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
        if not self.template_content:
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
        if not self.template_content:
            self.load_template()
        
        # 정확히 일치하는 마커만 치환 (부분 일치 방지)
        self.template_content = self.template_content.replace(marker, str(value))
        return self.template_content
    
    def validate_markers_against_excel(self, excel_sheets: List[str]) -> Dict:
        """
        템플릿의 마커가 엑셀 시트와 호환되는지 검증합니다.
        
        Args:
            excel_sheets: 엑셀 파일에서 사용 가능한 시트 목록
            
        Returns:
            검증 결과 딕셔너리:
            {
                'valid': 유효성 여부,
                'missing_sheets': 누락된 시트 목록,
                'warnings': 경고 메시지 목록,
                'marker_count': 전체 마커 수,
                'sheet_summary': 시트별 마커 수
            }
        """
        if not self.template_content:
            self.load_template()
        
        markers = self.extract_markers()
        
        # 시트별 마커 수 집계
        sheet_summary = {}
        for marker in markers:
            sheet_name = marker['sheet_name']
            if sheet_name not in sheet_summary:
                sheet_summary[sheet_name] = 0
            sheet_summary[sheet_name] += 1
        
        # 누락된 시트 찾기
        missing_sheets = []
        warnings = []
        
        for sheet_name in sheet_summary.keys():
            # 정확한 매칭
            if sheet_name in excel_sheets:
                continue
            
            # 부분 매칭 시도
            found = False
            for excel_sheet in excel_sheets:
                if sheet_name in excel_sheet or excel_sheet in sheet_name:
                    found = True
                    warnings.append(f"시트 '{sheet_name}'가 '{excel_sheet}'로 매핑될 수 있습니다.")
                    break
            
            if not found:
                missing_sheets.append(sheet_name)
        
        # 마커 형식 검증
        for marker in markers:
            cell_address = marker['cell_address']
            
            # 동적 마커 형식 검사 (예: 전국_증감률, 상위시도1_이름)
            if '_' in cell_address:
                # 동적 마커는 OK
                continue
            
            # 셀 주소 형식 검사 (예: A1, B5:C10)
            if not re.match(r'^[A-Z]+\d+(:[A-Z]+\d+)?$', cell_address):
                warnings.append(f"마커 '{marker['full_match']}'의 형식이 표준과 다릅니다.")
        
        return {
            'valid': len(missing_sheets) == 0,
            'missing_sheets': missing_sheets,
            'warnings': warnings,
            'marker_count': len(markers),
            'sheet_summary': sheet_summary
        }
    
    def get_marker_statistics(self) -> Dict:
        """
        템플릿의 마커 통계를 반환합니다.
        
        Returns:
            통계 딕셔너리:
            {
                'total_markers': 전체 마커 수,
                'unique_sheets': 고유 시트 수,
                'sheets': 시트 목록,
                'markers_with_operations': 계산식이 있는 마커 수,
                'dynamic_markers': 동적 마커 수 (헤더 기반),
                'cell_address_markers': 셀 주소 기반 마커 수
            }
        """
        if not self.template_content:
            self.load_template()
        
        markers = self.extract_markers()
        
        sheets = set()
        markers_with_ops = 0
        dynamic_markers = 0
        cell_address_markers = 0
        
        for marker in markers:
            sheets.add(marker['sheet_name'])
            
            if marker['operation']:
                markers_with_ops += 1
            
            cell_address = marker['cell_address']
            if '_' in cell_address or not re.match(r'^[A-Z]+\d+', cell_address):
                dynamic_markers += 1
            else:
                cell_address_markers += 1
        
        return {
            'total_markers': len(markers),
            'unique_sheets': len(sheets),
            'sheets': list(sheets),
            'markers_with_operations': markers_with_ops,
            'dynamic_markers': dynamic_markers,
            'cell_address_markers': cell_address_markers
        }

