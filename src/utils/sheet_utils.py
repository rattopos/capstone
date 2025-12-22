"""
시트 관련 유틸리티 모듈
시트명 매핑 및 변환 기능 제공
"""

# 시트명 매핑 (가상 시트명 -> 실제 시트명)
# 실업률 템플릿은 "실업자 수" 시트의 데이터를 사용
SHEET_NAME_MAPPING = {
    '실업률': '실업자 수'
}


def get_actual_sheet_name(sheet_name: str) -> str:
    """
    가상 시트명을 실제 시트명으로 변환합니다.
    
    Args:
        sheet_name: 템플릿에서 사용하는 시트명
        
    Returns:
        실제 엑셀 시트명
    """
    if not sheet_name:
        return sheet_name
    return SHEET_NAME_MAPPING.get(sheet_name, sheet_name)


def add_sheet_mapping(virtual_name: str, actual_name: str) -> None:
    """
    시트명 매핑을 추가합니다.
    
    Args:
        virtual_name: 가상 시트명 (템플릿에서 사용)
        actual_name: 실제 시트명 (엑셀에서 사용)
    """
    SHEET_NAME_MAPPING[virtual_name] = actual_name


def get_all_mappings() -> dict:
    """
    모든 시트명 매핑을 반환합니다.
    
    Returns:
        시트명 매핑 딕셔너리
    """
    return SHEET_NAME_MAPPING.copy()

