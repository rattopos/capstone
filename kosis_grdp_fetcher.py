#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KOSIS API를 통한 GRDP 데이터 자동 수집 모듈

현재 비활성화 상태입니다. 향후 활성화하려면:
1. KOSIS API 키를 설정 파일 또는 환경 변수에 추가
2. ENABLE_KOSIS_GRDP_FETCH 플래그를 True로 변경
3. 필요한 경우 API 엔드포인트 및 데이터 파싱 로직 수정
"""

import json
from typing import Dict, Optional, Any
from pathlib import Path
import os

# ===== 비활성화 플래그 =====
# 이 플래그를 True로 변경하면 KOSIS API를 통한 GRDP 데이터 수집이 활성화됩니다
ENABLE_KOSIS_GRDP_FETCH = False

# KOSIS API 키 (환경 변수 또는 설정 파일에서 로드)
# 실제 사용 시 환경 변수나 보안 설정 파일에서 읽어오도록 수정 필요
KOSIS_API_KEY = os.environ.get('KOSIS_API_KEY', '')


class KOSISGRDPFetcher:
    """KOSIS API를 통한 GRDP 데이터 수집 클래스"""
    
    # 지역 매핑 (KOSIS 코드 → 내부 지역명)
    REGION_MAPPING = {
        '서울특별시': '서울',
        '부산광역시': '부산',
        '대구광역시': '대구',
        '인천광역시': '인천',
        '광주광역시': '광주',
        '대전광역시': '대전',
        '울산광역시': '울산',
        '세종특별자치시': '세종',
        '경기도': '경기',
        '강원특별자치도': '강원',
        '충청북도': '충북',
        '충청남도': '충남',
        '전북특별자치도': '전북',
        '전라남도': '전남',
        '경상북도': '경북',
        '경상남도': '경남',
        '제주특별자치도': '제주',
    }
    
    def __init__(self, api_key: Optional[str] = None):
        """
        Args:
            api_key: KOSIS API 키 (None이면 환경 변수 또는 기본값 사용)
        """
        self.api_key = api_key or KOSIS_API_KEY
        self.enabled = ENABLE_KOSIS_GRDP_FETCH and bool(self.api_key)
        
        if not self.enabled:
            print("[KOSIS GRDP Fetcher] 현재 비활성화 상태입니다. 활성화하려면 ENABLE_KOSIS_GRDP_FETCH = True로 설정하세요.")
    
    def fetch_grdp_data(
        self, 
        start_year: int = 2016, 
        start_quarter: int = 1,
        end_year: int = 2025,
        end_quarter: int = 2
    ) -> Optional[Dict[str, Any]]:
        """
        KOSIS API에서 GRDP 데이터를 가져옵니다
        
        Args:
            start_year: 시작 연도
            start_quarter: 시작 분기
            end_year: 종료 연도
            end_quarter: 종료 분기
            
        Returns:
            GRDP 데이터 딕셔너리 또는 None (비활성화 또는 실패 시)
            
        데이터 구조:
            {
                'national_summary': {
                    'growth_rate': float,
                    'direction': str,
                    'contributions': {...}
                },
                'regional_data': [
                    {
                        'region': str,
                        'growth_rate': float,
                        'manufacturing': float,
                        'service': float,
                        'construction': float,
                        'other': float
                    },
                    ...
                ]
            }
        """
        if not self.enabled:
            print("[KOSIS GRDP Fetcher] 비활성화 상태 - 데이터를 가져오지 않습니다.")
            return None
        
        try:
            print(f"[KOSIS GRDP Fetcher] GRDP 데이터 수집 시작: {start_year}Q{start_quarter} ~ {end_year}Q{end_quarter}")
            
            # TODO: KOSIS API 엔드포인트 및 파라미터 설정
            # 실제 API URL과 파라미터는 KOSIS 공유서비스에서 확인 필요
            # 예시 구조만 제공
            
            # API URL 예시 (실제 사용 시 kosis_config.json 또는 별도 설정 파일에서 로드)
            # api_url = f"https://kosis.kr/openapi/Param/statisticsParameterData.do?method=getList&apiKey={self.api_key}&..."
            
            # TODO: API 호출 및 데이터 파싱
            # import requests
            # response = requests.get(api_url)
            # data = response.json()
            
            # TODO: 데이터 변환 (KOSIS 형식 → 내부 형식)
            # grdp_data = self._parse_api_response(data, start_year, start_quarter, end_year, end_quarter)
            
            print("[KOSIS GRDP Fetcher] 데이터 수집 완료")
            return None  # 현재는 비활성화이므로 None 반환
            
        except Exception as e:
            print(f"[KOSIS GRDP Fetcher] 데이터 수집 실패: {e}")
            return None
    
    def _parse_api_response(
        self, 
        api_data: Dict, 
        start_year: int, 
        start_quarter: int,
        end_year: int,
        end_quarter: int
    ) -> Dict[str, Any]:
        """
        KOSIS API 응답을 내부 데이터 형식으로 변환
        
        Args:
            api_data: KOSIS API 응답 데이터
            start_year: 시작 연도
            start_quarter: 시작 분기
            end_year: 종료 연도
            end_quarter: 종료 분기
            
        Returns:
            변환된 GRDP 데이터 딕셔너리
        """
        # TODO: API 응답 구조에 맞게 파싱 로직 구현
        # KOSIS API의 실제 응답 구조를 확인 후 구현 필요
        
        # 예시 구조:
        # - API 응답에서 연도/분기별, 지역별 데이터 추출
        # - 지역명을 내부 형식으로 변환 (REGION_MAPPING 사용)
        # - 필요한 지표 추출 (전체, 제조업, 서비스업, 건설업, 기타)
        
        return {
            'national_summary': {
                'growth_rate': 0.0,
                'direction': '증가',
                'contributions': {
                    'manufacturing': 0.0,
                    'service': 0.0,
                    'construction': 0.0,
                    'other': 0.0
                }
            },
            'regional_data': []
        }
    
    def _convert_region_name(self, kosis_region_name: str) -> Optional[str]:
        """KOSIS 지역명을 내부 지역명으로 변환"""
        return self.REGION_MAPPING.get(kosis_region_name)


def fetch_grdp_from_kosis(
    start_year: int = 2016,
    start_quarter: int = 1,
    end_year: int = 2025,
    end_quarter: int = 2,
    api_key: Optional[str] = None
) -> Optional[Dict[str, Any]]:
    """
    KOSIS API에서 GRDP 데이터를 가져오는 편의 함수
    
    Args:
        start_year: 시작 연도
        start_quarter: 시작 분기
        end_year: 종료 연도
        end_quarter: 종료 분기
        api_key: KOSIS API 키 (선택사항)
        
    Returns:
        GRDP 데이터 딕셔너리 또는 None
    """
    fetcher = KOSISGRDPFetcher(api_key=api_key)
    return fetcher.fetch_grdp_data(start_year, start_quarter, end_year, end_quarter)


# ===== 활성화 가이드 =====
"""
향후 활성화 절차:

1. KOSIS API 키 발급
   - KOSIS 공유서비스 (https://kosis.kr) 접속
   - API 키 발급 받기
   - 환경 변수 또는 설정 파일에 저장

2. 활성화
   - 이 파일의 ENABLE_KOSIS_GRDP_FETCH를 True로 변경
   - KOSIS_API_KEY를 설정 (환경 변수 또는 코드 내 설정)

3. API 엔드포인트 설정
   - KOSIS 공유서비스에서 GRDP 관련 통계표 찾기
   - API URL 및 필요한 파라미터 확인
   - fetch_grdp_data() 메서드에 실제 API 호출 로직 추가

4. 데이터 파싱
   - KOSIS API 응답 구조 확인
   - _parse_api_response() 메서드에 실제 파싱 로직 구현
   - 지역명 매핑 및 데이터 형식 변환

5. 테스트
   - 소규모 데이터로 테스트
   - 에러 처리 및 로깅 확인
   - 데이터 검증
"""

