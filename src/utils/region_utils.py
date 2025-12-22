"""
지역 관련 유틸리티 모듈
지역명 매핑, 정규화, 비교 기능 제공
"""

from typing import List

# 지역명 매핑 상수 (긴 이름 -> 짧은 이름)
REGION_MAPPING = {
    '서울특별시': '서울', '부산광역시': '부산', '대구광역시': '대구',
    '인천광역시': '인천', '광주광역시': '광주', '대전광역시': '대전',
    '울산광역시': '울산', '세종특별자치시': '세종', '경기도': '경기',
    '강원도': '강원', '충청북도': '충북', '충청남도': '충남',
    '전라북도': '전북', '전라남도': '전남', '경상북도': '경북',
    '경상남도': '경남', '제주특별자치도': '제주'
}

# 역방향 매핑 (짧은 이름 -> 긴 이름)
REGION_MAPPING_REVERSE = {v: k for k, v in REGION_MAPPING.items()}

# 17개 시도 리스트 (전국 제외)
SIDO_LIST = [
    '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
    '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주'
]


def normalize_region_name(region_name: str) -> str:
    """
    지역명을 짧은 형식으로 정규화합니다.
    
    Args:
        region_name: 정규화할 지역명
        
    Returns:
        정규화된 짧은 지역명 (예: '서울특별시' -> '서울')
    """
    if not region_name:
        return ''
    region_name = str(region_name).strip()
    return REGION_MAPPING.get(region_name, region_name)


def get_full_region_name(short_name: str) -> str:
    """
    짧은 지역명을 긴 형식으로 변환합니다.
    
    Args:
        short_name: 짧은 지역명 (예: '서울')
        
    Returns:
        긴 형식의 지역명 (예: '서울특별시')
    """
    if not short_name:
        return ''
    short_name = str(short_name).strip()
    return REGION_MAPPING_REVERSE.get(short_name, short_name)


def is_same_region(region1: str, region2: str) -> bool:
    """
    두 지역명이 같은 지역인지 확인합니다.
    
    Args:
        region1: 첫 번째 지역명
        region2: 두 번째 지역명
        
    Returns:
        같은 지역이면 True, 아니면 False
    """
    if not region1 or not region2:
        return False
    norm1 = normalize_region_name(region1)
    norm2 = normalize_region_name(region2)
    return norm1 == norm2


def is_sido(region_name: str) -> bool:
    """
    지역명이 17개 시도인지 확인합니다.
    
    Args:
        region_name: 확인할 지역명
        
    Returns:
        17개 시도이면 True, 아니면 False
    """
    normalized = normalize_region_name(region_name)
    return normalized in SIDO_LIST


def get_sido_list() -> List[str]:
    """
    17개 시도 리스트를 반환합니다.
    
    Returns:
        시도 리스트
    """
    return SIDO_LIST.copy()

