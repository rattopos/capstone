#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
통계표 생성기

분석표 엑셀에서 데이터를 추출하여 통계표 HTML을 생성합니다.
스키마를 통해 분석표의 구조 변경에 유연하게 대응할 수 있습니다.
"""

import pandas as pd
import json
import re
from pathlib import Path
from jinja2 import Environment, FileSystemLoader
from typing import Dict, List, Optional, Any

# RawDataExtractor 임포트
try:
    from raw_data_extractor import RawDataExtractor
except ImportError:
    RawDataExtractor = None


class StatisticsTableGenerator:
    """통계표 생성 클래스
    
    데이터 로딩 전략:
    1. 과거 데이터 (statistics_historical_data.json)에서 먼저 로드
    2. JSON에 없는 데이터 (현재/최신 분기)는 기초자료 또는 분석표에서 동적 추출
    3. 새로 추출한 데이터는 옵션에 따라 JSON에 저장 가능
    
    과거 데이터 JSON 구조:
    {
        "_metadata": {
            "quarterly_range": {"start": "2016.1/4", "end": "2025.1/4"},
            "yearly_range": {"start": "2016", "end": "2024"}
        },
        "광공업생산지수": {
            "yearly": {"2016": {"전국": 2.2, ...}, ...},
            "quarterly": {"2016.1/4": {"전국": 1.5, ...}, ...}
        },
        ...
    }
    """
    
    # 과거 데이터 JSON 파일 경로
    HISTORICAL_DATA_FILE = "statistics_historical_data.json"
    
    # 통계표 항목 정의 (집계 시트 기준, 전년동기비 계산 필요)
    TABLE_CONFIG = {
        "광공업생산지수": {
            "집계_시트": "A(광공업생산)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 4, "값": "BCD"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 9, "2018": 9, "2019": 9, "2020": 9, "2021": 10, "2022": 11, "2023": 12, "2024": 13},
            "분기_시작_컬럼": 14,  # 2022 2/4부터
            "계산방식": "growth_rate"
        },
        "서비스업생산지수": {
            "집계_시트": "B(서비스업생산)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 4, "값": "E~S"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 8, "2018": 8, "2019": 8, "2020": 8, "2021": 9, "2022": 10, "2023": 11, "2024": 12},
            "분기_시작_컬럼": 13,
            "계산방식": "growth_rate"
        },
        "소매판매액지수": {
            "집계_시트": "C(소비)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 5, "값": "A0"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 7, "2018": 7, "2019": 7, "2020": 7, "2021": 8, "2022": 9, "2023": 10, "2024": 11},
            "분기_시작_컬럼": 12,
            "계산방식": "growth_rate"
        },
        "건설수주액": {
            "집계_시트": "F'(건설)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 2, "값": "0"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 5, "2018": 5, "2019": 5, "2020": 5, "2021": 6, "2022": 7, "2023": 8, "2024": 9},
            "분기_시작_컬럼": 10,
            "계산방식": "growth_rate"
        },
        "고용률": {
            "단위": "[%]",  # 절대값 (JSON에서 로드)
            "계산방식": "absolute"
        },
        "실업률": {
            "단위": "[%]",  # 절대값 (JSON에서 로드)
            "계산방식": "absolute"
        },
        "국내인구이동": {
            "단위": "[천 명]",  # 절대값 순이동 수 (JSON에서 로드)
            "계산방식": "absolute",
            "전국_제외": True  # 전국 데이터 없음
        },
        "수출액": {
            "집계_시트": "G(수출)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 2, "값": "0"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 9, "2018": 9, "2019": 9, "2020": 9, "2021": 10, "2022": 11, "2023": 12, "2024": 13},
            "분기_시작_컬럼": 14,
            "계산방식": "growth_rate"
        },
        "수입액": {
            "집계_시트": "H(수입)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 2, "값": "0"},
            "지역_컬럼": 1,
            "분류단계_컬럼": 2,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 9, "2018": 9, "2019": 9, "2020": 9, "2021": 10, "2022": 11, "2023": 12, "2024": 13},
            "분기_시작_컬럼": 14,
            "계산방식": "growth_rate"
        },
        "소비자물가지수": {
            "집계_시트": "E(품목성질물가)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 1, "값": 0},
            "지역_컬럼": 0,
            "분류단계_컬럼": 1,
            "데이터_시작_행": 3,
            "연도_컬럼": {"2017": 7, "2018": 7, "2019": 7, "2020": 7, "2021": 8, "2022": 9, "2023": 10, "2024": 11},
            "분기_시작_컬럼": 12,
            "계산방식": "growth_rate"
        }
    }
    
    # 기초자료 시트 매핑
    RAW_SHEET_MAPPING = {
        "광공업생산지수": "광공업생산",
        "서비스업생산지수": "서비스업생산",
        "소매판매액지수": "소비(소매, 추가)",
        "건설수주액": "건설 (공표자료)",
        "고용률": "고용률",
        "실업률": "실업자 수",
        "수출액": "수출",
        "수입액": "수입",
        "국내인구이동": "시도 간 이동",
        "소비자물가지수": "품목성질별 물가"
    }
    
    # 기초자료 시트별 컬럼 매핑 (분석표와 다른 구조 대응)
    # 각 시트에서 지역이름과 분류단계의 컬럼 인덱스
    RAW_COLUMN_MAPPING = {
        "광공업생산": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "서비스업생산": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "소비(소매, 추가)": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "건설 (공표자료)": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "고용률": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "difference"},  # %p 단위
        "실업자 수": {"지역_컬럼": 0, "분류단계_컬럼": 1, "분류값": "계", "계산방식": "growth_rate"},
        "수출": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "수입": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "0", "계산방식": "growth_rate"},
        "시도 간 이동": {"지역_컬럼": 1, "분류단계_컬럼": 2, "분류값": "순인구이동 수", "계산방식": "growth_rate"},
        "품목성질별 물가": {"지역_컬럼": 0, "분류단계_컬럼": 1, "분류값": "0", "계산방식": "growth_rate"},
    }
    
    # 지역 목록 (페이지별)
    PAGE1_REGIONS = ["전국", "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종"]
    PAGE2_REGIONS = ["경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"]
    ALL_REGIONS = PAGE1_REGIONS + PAGE2_REGIONS
    
    def __init__(self, excel_path: str, historical_data_path: Optional[str] = None, 
                 raw_excel_path: Optional[str] = None, current_year: int = 2025, current_quarter: int = 2,
                 auto_update_json: bool = False):
        """
        초기화
        
        Args:
            excel_path: 분석표 엑셀 파일 경로
            historical_data_path: 과거 데이터 JSON 파일 경로 (사용하지 않음, 자동 로드)
            raw_excel_path: 기초자료 엑셀 파일 경로 (최신 데이터 추출용)
            current_year: 현재 연도
            current_quarter: 현재 분기
            auto_update_json: 새로 추출한 데이터를 JSON에 자동 저장할지 여부
        """
        self.excel_path = excel_path
        self.raw_excel_path = raw_excel_path
        self.current_year = current_year
        self.current_quarter = current_quarter
        self.historical_data_path = historical_data_path
        self.historical_data = {}
        self.historical_metadata = {}  # JSON 메타데이터
        self.cached_sheets = {}
        self.auto_update_json = auto_update_json
        self.newly_extracted_data = {}  # 새로 추출한 데이터 (JSON 업데이트용)
        
        # 파일 존재 및 수정 시간 확인
        excel_path_obj = Path(excel_path)
        if not excel_path_obj.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {excel_path}")
        
        try:
            self._file_mtime = excel_path_obj.stat().st_mtime
        except OSError:
            self._file_mtime = None
        
        # 과거 데이터 JSON 자동 로드
        self._load_historical_data()
        
        # RawDataExtractor 초기화 (최신 데이터 추출용)
        self.raw_extractor = None
        if raw_excel_path and RawDataExtractor and Path(raw_excel_path).exists():
            try:
                self.raw_extractor = RawDataExtractor(raw_excel_path, current_year, current_quarter)
                print(f"[통계표] 기초자료 추출기 초기화 완료: {raw_excel_path}")
            except Exception as e:
                print(f"[통계표] 기초자료 추출기 초기화 실패: {e}")
    
    def _load_historical_data(self):
        """과거 데이터 JSON 로드 (메타데이터 포함)"""
        historical_path = Path(__file__).parent / self.HISTORICAL_DATA_FILE
        if historical_path.exists():
            try:
                with open(historical_path, 'r', encoding='utf-8') as f:
                    raw_data = json.load(f)
                
                # 메타데이터 분리
                self.historical_metadata = raw_data.pop('_metadata', {})
                self.historical_data = raw_data
                
                # 메타데이터에서 범위 정보 확인
                q_range = self.historical_metadata.get('quarterly_range', {})
                y_range = self.historical_metadata.get('yearly_range', {})
                
                print(f"[통계표] 과거 데이터 로드 완료: {len(self.historical_data)}개 지표")
                print(f"[통계표] - 연도별 범위: {y_range.get('start', '?')} ~ {y_range.get('end', '?')}")
                print(f"[통계표] - 분기별 범위: {q_range.get('start', '?')} ~ {q_range.get('end', '?')}")
                
            except Exception as e:
                print(f"[통계표] 과거 데이터 로드 실패: {e}")
                self.historical_data = {}
                self.historical_metadata = {}
        else:
            print(f"[통계표] 과거 데이터 파일 없음: {historical_path}")
            self.historical_data = {}
            self.historical_metadata = {}
    
    def _is_quarter_in_json(self, quarter_key: str) -> bool:
        """해당 분기 데이터가 JSON에 있는지 확인"""
        q_range = self.historical_metadata.get('quarterly_range', {})
        json_end = q_range.get('end', '')
        
        if not json_end:
            return False
        
        # 분기 키 비교 (예: "2025.1/4" vs "2025.2/4p")
        quarter_key_clean = quarter_key.rstrip('p')
        json_end_clean = json_end.rstrip('p')
        
        # 연도.분기 형식 파싱
        def parse_quarter(q):
            match = re.match(r'(\d{4})\.(\d)/4', q)
            if match:
                return int(match.group(1)), int(match.group(2))
            return None, None  # PM 요구사항: 데이터 없음 → N/A
        
        current_year, current_q = parse_quarter(quarter_key_clean)
        end_year, end_q = parse_quarter(json_end_clean)
        
        # 현재 분기가 JSON 범위 내에 있는지 확인
        if current_year < end_year:
            return True
        elif current_year == end_year and current_q <= end_q:
            return True
        return False
    
    def _get_missing_quarters(self) -> List[str]:
        """JSON에 없는 분기 목록 반환"""
        missing = []
        q_range = self.historical_metadata.get('quarterly_range', {})
        json_end = q_range.get('end', '2016.1/4')
        
        # JSON 끝 분기 이후부터 현재 분기까지
        def parse_quarter(q):
            match = re.match(r'(\d{4})\.(\d)/4', q.rstrip('p'))
            if match:
                return int(match.group(1)), int(match.group(2))
            return 2016, 1
        
        end_year, end_q = parse_quarter(json_end)
        
        # JSON 다음 분기부터 시작
        if end_q == 4:
            start_year, start_q = end_year + 1, 1
        else:
            start_year, start_q = end_year, end_q + 1
        
        # 현재 분기까지 목록 생성
        year, q = start_year, start_q
        while year < self.current_year or (year == self.current_year and q <= self.current_quarter):
            is_current = (year == self.current_year and q == self.current_quarter)
            quarter_key = f"{year}.{q}/4"
            if is_current:
                quarter_key += "p"
            missing.append(quarter_key)
            
            if q == 4:
                year, q = year + 1, 1
            else:
                q += 1
        
        return missing
    
    def _check_file_modified(self) -> bool:
        """파일이 수정되었는지 확인 (캐시 무효화 판단)"""
        if self._file_mtime is None:
            return False
        
        try:
            excel_path_obj = Path(self.excel_path)
            if not excel_path_obj.exists():
                return True  # 파일이 없으면 무효화 필요
            
            current_mtime = excel_path_obj.stat().st_mtime
            if abs(current_mtime - self._file_mtime) > 1.0:  # 1초 이상 차이
                return True
            return False
        except OSError:
            return True  # 파일 접근 오류 시 안전하게 무효화
    
    def _clear_cache(self):
        """캐시 무효화"""
        self.cached_sheets.clear()
    
    def _load_sheet(self, sheet_name: str) -> pd.DataFrame:
        """엑셀 시트 로드 (캐싱, 파일 수정 시간 확인)"""
        # 파일이 수정되었으면 캐시 무효화
        if self._check_file_modified():
            self._clear_cache()
            try:
                excel_path_obj = Path(self.excel_path)
                self._file_mtime = excel_path_obj.stat().st_mtime
            except OSError:
                pass
        
        if sheet_name not in self.cached_sheets:
            try:
                # 시트 데이터 로드 (예외 발생 시 캐시에 저장하지 않음)
                df = pd.read_excel(
                    self.excel_path,
                    sheet_name=sheet_name,
                    header=None
                )
                # 정상 로드된 경우에만 캐시에 저장
                self.cached_sheets[sheet_name] = df
            except Exception as e:
                print(f"시트 로드 실패: {sheet_name} - {e}")
                return None
        return self.cached_sheets[sheet_name]
    
    def _extract_from_raw_data(self, raw_sheet_name: str, config: dict) -> Optional[Dict[str, Any]]:
        """기초자료에서 데이터 추출
        
        Args:
            raw_sheet_name: 기초자료 시트 이름
            config: 통계표 설정
            
        Returns:
            {
                'yearly': {'2016': {'전국': value, ...}, ...},
                'quarterly': {'2016.1/4': {'전국': value, ...}, ...}
            }
        """
        if not self.raw_extractor:
            return None
        
        # 기초자료 시트별 컬럼 매핑 사용 (분석표 설정이 아닌 기초자료 구조에 맞춤)
        raw_col_config = self.RAW_COLUMN_MAPPING.get(raw_sheet_name, {})
        
        region_column = raw_col_config.get("지역_컬럼", 1)
        classification_column = raw_col_config.get("분류단계_컬럼", 2)
        classification_value = raw_col_config.get("분류값", "0")
        
        calculation_method = raw_col_config.get("계산방식", "growth_rate")
        
        print(f"[통계표] 기초자료 추출 - 시트: {raw_sheet_name}, 지역컬럼: {region_column}, 분류컬럼: {classification_column}, 분류값: {classification_value}, 계산방식: {calculation_method}")
        
        # 계산방식에 따라 다른 함수 호출
        if calculation_method == "difference":
            # 차이 계산 (%p 단위)
            yearly_data = self.raw_extractor.extract_yearly_difference(
                raw_sheet_name,
                start_year=2016,
                region_column=region_column,
                classification_column=classification_column,
                classification_value=classification_value if classification_value else None
            )
            quarterly_data = self.raw_extractor.extract_quarterly_difference(
                raw_sheet_name,
                start_year=2016,
                start_quarter=1,
                region_column=region_column,
                classification_column=classification_column,
                classification_value=classification_value if classification_value else None
            )
        else:
            # 전년동기비 계산 (% 단위)
            yearly_data = self.raw_extractor.extract_yearly_growth_rate(
                raw_sheet_name,
                start_year=2016,
                region_column=region_column,
                classification_column=classification_column,
                classification_value=classification_value if classification_value else None
            )
            quarterly_data = self.raw_extractor.extract_quarterly_growth_rate(
                raw_sheet_name,
                start_year=2016,
                start_quarter=1,
                region_column=region_column,
                classification_column=classification_column,
                classification_value=classification_value if classification_value else None
            )
        
        print(f"[통계표] 기초자료 추출 결과 - 연도: {len(yearly_data)}개, 분기: {len(quarterly_data)}개")
        
        # 데이터 형식 변환 (분기 키 형식 통일)
        # raw_data_extractor가 이미 "2025.2/4p" 형식으로 반환하므로 그대로 사용
        quarterly_formatted = {}
        for quarter_key, data in quarterly_data.items():
            # 공백 형식("2016 1/4")인 경우 점 형식("2016.1/4")으로 변환
            formatted_key = quarter_key.replace(" ", ".")
            quarterly_formatted[formatted_key] = data
        
        return {
            'yearly': yearly_data,
            'quarterly': quarterly_formatted
        }
    
    # 지역명 매핑 (실업률 시트용)
    REGION_NAME_MAPPING = {
        "전국": ["전국"],
        "서울": ["서울", "서울특별시"],
        "부산": ["부산", "부산광역시"],
        "대구": ["대구", "대구광역시"],
        "인천": ["인천", "인천광역시"],
        "광주": ["광주", "광주광역시"],
        "대전": ["대전", "대전광역시"],
        "울산": ["울산", "울산광역시"],
        "세종": ["세종", "세종특별자치시"],
        "경기": ["경기", "경기도"],
        "강원": ["강원", "강원도", "강원특별자치도"],
        "충북": ["충북", "충청북도"],
        "충남": ["충남", "충청남도"],
        "전북": ["전북", "전라북도", "전북특별자치도"],
        "전남": ["전남", "전라남도"],
        "경북": ["경북", "경상북도"],
        "경남": ["경남", "경상남도"],
        "제주": ["제주", "제주도", "제주특별자치도"]
    }
    
    def _get_total_row(self, df: pd.DataFrame, region: str, config: dict) -> Optional[pd.Series]:
        """특정 지역의 총지수 행 가져오기 (시트별 설정 사용)"""
        region_col = config["지역_컬럼"]
        id_col = config["총지수_식별"]["컬럼"]
        id_val = config["총지수_식별"]["값"]
        use_region_mapping = config.get("지역_매핑", False)
        
        # 값 타입에 따라 비교
        try:
            df_region = df[region_col].astype(str).str.strip()
            region_clean = str(region).strip()
            
            # 지역명 매핑 사용 여부
            if use_region_mapping:
                region_variants = self.REGION_NAME_MAPPING.get(region_clean, [region_clean])
                mask = df_region.isin(region_variants) & (df[id_col].astype(str).str.strip() == str(id_val).strip())
            else:
                mask = (df_region == region_clean) & (df[id_col].astype(str).str.strip() == str(id_val).strip())
            
            result = df[mask]
            
            if not result.empty:
                return result.iloc[0]
        except Exception as e:
            print(f"행 검색 오류: {region} - {e}")
        
        return None
    
    def _calculate_yoy_growth(self, current: float, previous: float) -> Optional[float]:
        """전년동기비 증감률 계산"""
        if pd.isna(current) or pd.isna(previous) or previous == 0:
            return None
        try:
            return ((float(current) - float(previous)) / float(previous)) * 100
        except (ValueError, TypeError):
            return None
    
    def _calculate_difference(self, current: float, previous: float) -> Optional[float]:
        """전년동기비 차이 계산 (%p)"""
        if pd.isna(current) or pd.isna(previous):
            return None
        try:
            return float(current) - float(previous)
        except (ValueError, TypeError):
            return None
    
    def _format_value(self, value, decimals: int = 1) -> str:
        """값 포맷팅"""
        if pd.isna(value):
            return "-"
        try:
            val = float(value)
            return f"{val:.{decimals}f}"
        except (ValueError, TypeError):
            return str(value)
    
    def _create_empty_table_data(self) -> Dict[str, Any]:
        """모든 연도/분기/지역에 기본값 '-'가 채워진 데이터 구조 생성
        
        원본 이미지 기준 범위:
        - 연도별: 현재년도 제외 최근 8년 (예: 2025년 기준 → 2017~2024)
        - 분기별: (현재년도-9)년 4분기부터 현재 분기까지 (예: 2016.4/4 ~ 2025.2/4p)
        """
        # 연도별: 현재년도 제외 최근 8년
        yearly_start = self.current_year - 8  # 예: 2025 - 8 = 2017
        yearly_end = self.current_year - 1    # 예: 2025 - 1 = 2024
        yearly_years = [str(year) for year in range(yearly_start, yearly_end + 1)]
        
        # 분기별: (현재년도-9)년 4분기부터 현재 분기까지
        quarterly_keys = []
        q_start_year = self.current_year - 9  # 예: 2025 - 9 = 2016
        
        # 시작: q_start_year의 4분기
        quarterly_keys.append(f"{q_start_year}.4/4")
        
        # 그 다음 해부터 현재 분기까지
        for year in range(q_start_year + 1, self.current_year + 1):
            end_q = self.current_quarter if year == self.current_year else 4
            for quarter in range(1, end_q + 1):
                if year == self.current_year and quarter == self.current_quarter:
                    quarterly_keys.append(f"{year}.{quarter}/4p")
                else:
                    quarterly_keys.append(f"{year}.{quarter}/4")
        
        # 모든 연도에 대해 모든 지역의 기본값 '-' 설정
        yearly = {}
        for year in yearly_years:
            yearly[year] = {region: "-" for region in self.ALL_REGIONS}
        
        # 모든 분기에 대해 모든 지역의 기본값 '-' 설정
        quarterly = {}
        for quarter in quarterly_keys:
            quarterly[quarter] = {region: "-" for region in self.ALL_REGIONS}
        
        return {
            "yearly": yearly,
            "quarterly": quarterly,
            "yearly_years": yearly_years,
            "quarterly_keys": quarterly_keys
        }
    
    def extract_table_data(self, table_name: str) -> Dict[str, Any]:
        """특정 통계표의 데이터 추출 (하이브리드: JSON + 동적 추출)
        
        데이터 로딩 전략:
        1. 과거 데이터 JSON에서 먼저 로드 (확정된 데이터)
        2. JSON에 없는 분기만 기초자료/분석표에서 동적 추출
        3. auto_update_json이 True면 새 데이터를 JSON에 저장
        """
        config = self.TABLE_CONFIG.get(table_name)
        if not config:
            raise ValueError(f"알 수 없는 통계표: {table_name}")
        
        # 데이터 구조 초기화 - 모든 연도/분기/지역에 기본값 '-' 설정
        data = self._create_empty_table_data()
        
        # 1. 과거 데이터 JSON에서 로드 (우선순위 1)
        json_loaded = False
        if table_name in self.historical_data:
            historical = self.historical_data[table_name]
            
            # 연도별 데이터 로드
            for year, regions in historical.get('yearly', {}).items():
                if year in data['yearly']:
                    for region, value in regions.items():
                        if value is not None:
                            data['yearly'][year][region] = value
            
            # 분기별 데이터 로드
            for quarter, regions in historical.get('quarterly', {}).items():
                if quarter in data['quarterly']:
                    for region, value in regions.items():
                        if value is not None:
                            data['quarterly'][quarter][region] = value
            
            json_loaded = True
            print(f"[통계표] JSON에서 과거 데이터 로드: {table_name}")
        
        # 2. JSON에 없는 분기 확인
        missing_quarters = self._get_missing_quarters()
        
        if missing_quarters:
            print(f"[통계표] JSON에 없는 분기 동적 추출 필요: {missing_quarters}")
            
            # 기초자료에서 최신 데이터 추출
            if self.raw_extractor:
                raw_sheet_name = self.RAW_SHEET_MAPPING.get(table_name)
                if raw_sheet_name:
                    try:
                        print(f"[통계표] 기초자료에서 데이터 추출: {table_name}")
                        raw_data = self._extract_from_raw_data(raw_sheet_name, config)
                        if raw_data:
                            # JSON에 없는 분기만 업데이트
                            for quarter in missing_quarters:
                                quarter_clean = quarter.rstrip('p')
                                # raw_data에서 해당 분기 찾기
                                quarter_data = raw_data.get('quarterly', {}).get(quarter) or \
                                              raw_data.get('quarterly', {}).get(quarter_clean)
                                
                                if quarter_data and quarter in data['quarterly']:
                                    for region, value in quarter_data.items():
                                        if value is not None:
                                            data['quarterly'][quarter][region] = value
                                    print(f"[통계표] 분기 데이터 추가: {quarter}")
                                    
                                    # 새로 추출한 데이터 저장 (JSON 업데이트용)
                                    if table_name not in self.newly_extracted_data:
                                        self.newly_extracted_data[table_name] = {'yearly': {}, 'quarterly': {}}
                                    self.newly_extracted_data[table_name]['quarterly'][quarter] = quarter_data
                            
                            # 연도 데이터도 필요시 업데이트
                            for year, regions in raw_data.get('yearly', {}).items():
                                if year not in self.historical_data.get(table_name, {}).get('yearly', {}):
                                    if year in data['yearly']:
                                        for region, value in regions.items():
                                            if value is not None:
                                                data['yearly'][year][region] = value
                                        
                                        # 새로 추출한 데이터 저장
                                        if table_name not in self.newly_extracted_data:
                                            self.newly_extracted_data[table_name] = {'yearly': {}, 'quarterly': {}}
                                        self.newly_extracted_data[table_name]['yearly'][year] = regions
                    except Exception as e:
                        print(f"[통계표] 기초자료 추출 실패: {table_name} - {e}")
            
            # 기초자료 없으면 분석표 집계 시트에서 추출
            if not self.raw_extractor:
                self._extract_from_aggregate_sheet(table_name, config, data)
        
        # JSON 데이터가 없고 동적 추출도 실패한 경우 -> 전체 동적 추출 시도
        if not json_loaded and not missing_quarters:
            print(f"[통계표] JSON 없음, 전체 동적 추출: {table_name}")
            self._extract_all_dynamic(table_name, config, data)
        
        # 모든 지역이 "-"인 분기 제거 (해당 통계에서 데이터를 제공하지 않는 분기)
        quarters_to_remove = []
        for q_key, regions in data['quarterly'].items():
            if all(v == '-' or v is None for v in regions.values()):
                quarters_to_remove.append(q_key)
        
        for q_key in quarters_to_remove:
            del data['quarterly'][q_key]
            if q_key in data['quarterly_keys']:
                data['quarterly_keys'].remove(q_key)
        
        return data
    
    def _extract_all_dynamic(self, table_name: str, config: dict, data: dict):
        """모든 데이터를 동적으로 추출 (JSON 없는 경우)"""
        # 기초자료에서 직접 추출 (우선순위 1)
        if self.raw_extractor:
            raw_sheet_name = self.RAW_SHEET_MAPPING.get(table_name)
            if raw_sheet_name:
                try:
                    print(f"[통계표] 기초자료에서 전체 추출: {table_name}")
                    raw_data = self._extract_from_raw_data(raw_sheet_name, config)
                    if raw_data:
                        for year, regions in raw_data.get('yearly', {}).items():
                            if year in data['yearly']:
                                data['yearly'][year].update({k: v for k, v in regions.items() if v is not None})
                        
                        for quarter, regions in raw_data.get('quarterly', {}).items():
                            if quarter in data['quarterly']:
                                data['quarterly'][quarter].update({k: v for k, v in regions.items() if v is not None})
                        return
                except Exception as e:
                    print(f"[통계표] 기초자료 추출 실패: {table_name} - {e}")
        
        # 분석표 집계 시트에서 추출 (fallback)
        self._extract_from_aggregate_sheet(table_name, config, data)
    
    def _extract_from_aggregate_sheet(self, table_name: str, config: dict, data: dict):
        """분석표 집계 시트에서 데이터 추출"""
        sheet_name = config.get("집계_시트")
        if not sheet_name:
            print(f"[통계표] 집계_시트 설정 없음: {table_name}")
            return
            
        df = self._load_sheet(sheet_name)
        if df is None:
            print(f"[통계표] 시트를 찾을 수 없음: {sheet_name}")
            return
        
        print(f"[통계표] 집계 시트에서 추출: {table_name} (시트: {sheet_name})")
        
        # 헤더 행에서 분기 컬럼 위치 파악
        header_row = df.iloc[2]
        quarter_col_map = {}
        for col_idx, header in enumerate(header_row):
            if pd.notna(header):
                header_str = str(header).strip()
                # "2022  2/4" 형식 파싱
                match = re.match(r'(\d{4})\s+(\d)/4', header_str)
                if match:
                    year = match.group(1)
                    q = match.group(2)
                    quarter_key = f"{year}.{q}/4"
                    quarter_col_map[quarter_key] = col_idx
        
        calculation_method = config.get("계산방식", "growth_rate")
        
        # 각 지역에 대해 데이터 추출
        for region in self.ALL_REGIONS:
            row = self._get_total_row(df, region, config)
            if row is None:
                continue
            
            # 연도별 데이터 (전년동기비 계산)
            for year in data["yearly_years"]:
                # 이미 값이 있으면 스킵 (JSON에서 로드됨)
                if data["yearly"].get(year, {}).get(region, "-") != "-":
                    continue
                
                year_int = int(year)
                prev_year = str(year_int - 1)
                
                curr_col = config["연도_컬럼"].get(year)
                prev_col = config["연도_컬럼"].get(prev_year)
                
                if curr_col is not None and curr_col < len(row):
                    curr_val = row.iloc[curr_col]
                    
                    if calculation_method == "absolute":
                        # 절대값: 전년동기비 계산 없이 현재 값 그대로 사용
                        if not pd.isna(curr_val):
                            try:
                                data["yearly"][year][region] = round(float(curr_val), 1)
                            except (ValueError, TypeError):
                                pass
                    elif prev_col is not None and prev_col < len(row):
                        prev_val = row.iloc[prev_col]
                        
                        if calculation_method == "difference":
                            result = self._calculate_difference(curr_val, prev_val)
                        else:
                            result = self._calculate_yoy_growth(curr_val, prev_val)
                        
                        if result is not None:
                            data["yearly"][year][region] = round(result, 1)
            
            # 분기별 데이터 (전년동기비 계산 또는 절대값)
            for quarter_key in data["quarterly_keys"]:
                # 이미 값이 있으면 스킵 (JSON에서 로드됨)
                if data["quarterly"].get(quarter_key, {}).get(region, "-") != "-":
                    continue
                
                match = re.match(r'(\d{4})\.(\d)/4(p?)', quarter_key)
                if not match:
                    continue
                
                year = int(match.group(1))
                q = int(match.group(2))
                prev_year_quarter_key = f"{year - 1}.{q}/4"
                
                curr_col = quarter_col_map.get(quarter_key.rstrip('p'))
                prev_col = quarter_col_map.get(prev_year_quarter_key)
                
                if curr_col is not None and curr_col < len(row):
                    curr_val = row.iloc[curr_col]
                    
                    if calculation_method == "absolute":
                        # 절대값: 전년동기비 계산 없이 현재 값 그대로 사용
                        if not pd.isna(curr_val):
                            try:
                                data["quarterly"][quarter_key][region] = round(float(curr_val), 1)
                            except (ValueError, TypeError):
                                pass
                    elif prev_col is not None and prev_col < len(row):
                        prev_val = row.iloc[prev_col]
                        
                        if calculation_method == "difference":
                            result = self._calculate_difference(curr_val, prev_val)
                        else:
                            result = self._calculate_yoy_growth(curr_val, prev_val)
                        
                        if result is not None:
                            data["quarterly"][quarter_key][region] = round(result, 1)
    
    def extract_all_tables(self, year: Optional[int] = None, quarter: Optional[int] = None) -> Dict[str, Any]:
        """모든 통계표 데이터 추출"""
        if year is None:
            year = self.current_year
        if quarter is None:
            quarter = self.current_quarter
        tables = []
        # 통계표 순서: 광공업생산지수-서비스업생산지수-소매판매액지수-건설수주액-고용률-실업률-국내인구이동-수출액-수입액-소비자물가지수
        table_order = [
            "광공업생산지수",
            "서비스업생산지수",
            "소매판매액지수",
            "건설수주액",
            "고용률",
            "실업률",
            "국내인구이동",
            "수출액",
            "수입액",
            "소비자물가지수"
        ]
        
        page_num = 22  # 시작 페이지 번호
        
        for idx, table_name in enumerate(table_order, 1):
            config = self.TABLE_CONFIG[table_name]
            data = self.extract_table_data(table_name)
            
            if data:
                tables.append({
                    "id": idx,
                    "title": table_name,
                    "unit": config["단위"],
                    "data": data,
                    "page_number_1": page_num,
                    "page_number_2": page_num + 1
                })
                page_num += 2
        
        # 부록 용어 정의
        appendix = {
            "terms": [
                {
                    "term": "불변지수",
                    "definition": "불변지수는 가격 변동분이 제외된 수량 변동분만 포함되어 있음을 의미하며, 성장 수준 분석(전년동분기비)에 활용됨"
                },
                {
                    "term": "광공업생산지수",
                    "definition": "한국표준산업분류 상의 3개 대분류(B, C, D)를 대상으로 광업제조업동향조사의 월별 품목별 생산·출하(내수 및 수출)·재고 및 생산능력·가동률지수를 기초로 작성됨"
                },
                {
                    "term": "서비스업생산지수",
                    "definition": "한국표준산업분류 상의 13개 대분류(E, G, H, I, J, K, L, M, N, P, Q, R, S)를 대상으로 서비스업동향조사의 월별 매출액을 기초로 작성됨"
                },
                {
                    "term": "소매판매액지수",
                    "definition": "한국표준산업분류 상의 '자동차 판매업 중 승용차'와 '소매업'을 대상으로 서비스업동향조사의 월별 상품판매액을 기초로 작성됨"
                },
                {
                    "term": "건설수주",
                    "definition": "종합건설업 등록업체 중 전전년 「건설업조사」 결과를 기준으로 기성액 순위 상위 기업체(대표도: 54%)의 국내공사에 대한 건설수주액임"
                },
                {
                    "term": "소비자물가지수",
                    "definition": "가구에서 일상생활을 영위하기 위해 구입하는 상품과 서비스의 평균적인 가격변동을 측정한 지수임"
                },
                {
                    "term": "지역내총생산",
                    "definition": "일정 기간 동안에 일정 지역 내에서 새로이 창출된 최종생산물을 시장가격으로 평가한 가치의 합임"
                }
            ]
        }
        
        # GRDP 데이터 (현재 N/A로 처리)
        grdp_data = self._create_grdp_placeholder()
        
        return {
            "report_info": {
                "year": year,
                "quarter": quarter
            },
            "tables": tables,
            "grdp": grdp_data,
            "appendix": appendix,
            "page_numbers": {
                "toc": 21,
                "appendix_1": page_num + 2,  # GRDP 2페이지 후
                "appendix_2": page_num + 3
            }
        }
    
    def _create_grdp_placeholder(self) -> Dict[str, Any]:
        """GRDP 데이터 생성 - grdp_historical_data.json 및 grdp_extracted.json에서 데이터 로드
        
        원본 이미지 기준 범위:
        - 연도별: 현재년도 제외 최근 8년 (예: 2025년 기준 → 2017~2024)
          - 마지막 2년에 'p' 표시 (2023p, 2024p)
        - 분기별: (현재년도-8)년 3분기부터 현재 분기 이전까지
          - GRDP는 1분기 늦게 발표되므로 현재 분기가 2/4이면 1/4까지만 표시
          - 예: 2025년 2분기 기준 → 2017.3/4 ~ 2025.1/4p
        """
        import json
        
        # 연도별: 현재년도 제외 최근 8년
        yearly_start = self.current_year - 8  # 예: 2025 - 8 = 2017
        yearly_end = self.current_year - 1    # 예: 2025 - 1 = 2024
        yearly_years = []
        for year in range(yearly_start, yearly_end + 1):
            # 마지막 2년에 'p' 표시 (잠정치)
            if year >= self.current_year - 2:
                yearly_years.append(f"{year}p")
            else:
                yearly_years.append(str(year))
        
        # 분기별: (현재년도-8)년 3분기부터 현재 분기 이전까지
        # GRDP는 1분기 늦게 발표됨
        quarterly_keys = []
        q_start_year = self.current_year - 8  # 예: 2025 - 8 = 2017
        
        # GRDP 최신 분기 계산 (현재 분기보다 1분기 이전)
        if self.current_quarter == 1:
            grdp_end_year = self.current_year - 1
            grdp_end_quarter = 4
        else:
            grdp_end_year = self.current_year
            grdp_end_quarter = self.current_quarter - 1
        
        for year in range(q_start_year, grdp_end_year + 1):
            # 시작 분기: 첫 해는 3분기부터
            start_q = 3 if year == q_start_year else 1
            # 끝 분기
            end_q = grdp_end_quarter if year == grdp_end_year else 4
            
            for quarter in range(start_q, end_q + 1):
                # 마지막 2년의 분기에 'p' 표시
                if year >= self.current_year - 2:
                    quarterly_keys.append(f"{year}p.{quarter}/4")
                else:
                    quarterly_keys.append(f"{year}.{quarter}/4")
        
        # 기본값으로 플레이스홀더 (편집 가능) 생성
        yearly = {}
        for year in yearly_years:
            yearly[year] = {region: "-" for region in self.ALL_REGIONS}
        
        quarterly = {}
        for qk in quarterly_keys:
            quarterly[qk] = {region: "-" for region in self.ALL_REGIONS}
        
        # grdp_historical_data.json에서 과거 데이터 로드
        try:
            hist_json_path = Path(__file__).parent / 'grdp_historical_data.json'
            if hist_json_path.exists():
                with open(hist_json_path, 'r', encoding='utf-8') as f:
                    hist_data = json.load(f)
                
                # 연도별 데이터 로드 ('p' 표시 키와 원본 키 매핑)
                if 'yearly' in hist_data:
                    for display_key in yearly_years:
                        # 'p' 제거하여 원본 키 생성
                        source_key = display_key.rstrip('p')
                        if source_key in hist_data['yearly']:
                            for region, value in hist_data['yearly'][source_key].items():
                                if region in self.ALL_REGIONS:
                                    yearly[display_key][region] = value
                
                # 분기별 데이터 로드 ('p' 표시 키와 원본 키 매핑)
                if 'quarterly' in hist_data:
                    for display_key in quarterly_keys:
                        # '2023p.1/4' → '2023.1/4' 형태로 변환
                        source_key = display_key.replace('p.', '.')
                        if source_key in hist_data['quarterly']:
                            for region, value in hist_data['quarterly'][source_key].items():
                                if region in self.ALL_REGIONS:
                                    quarterly[display_key][region] = value
                
                print(f"[통계표] GRDP 과거 데이터 로드 완료")
        except Exception as e:
            print(f"[통계표] GRDP 과거 데이터 로드 실패: {e}")
        
        # grdp_extracted.json에서 현재 분기 데이터 로드 시도 (최신 데이터로 덮어쓰기)
        try:
            grdp_json_path = Path(__file__).parent / 'grdp_extracted.json'
            if grdp_json_path.exists():
                with open(grdp_json_path, 'r', encoding='utf-8') as f:
                    grdp_data = json.load(f)
                
                # 현재 분기 키
                current_key = f"{self.current_year}.{self.current_quarter}/4p"
                
                if current_key in quarterly and 'regional_data' in grdp_data:
                    for item in grdp_data['regional_data']:
                        region = item.get('region', '')
                        growth_rate = item.get('growth_rate', 0)
                        # placeholder가 명시적으로 True가 아니면 데이터 사용
                        if region in self.ALL_REGIONS and not item.get('placeholder', False):
                            quarterly[current_key][region] = round(growth_rate, 1)
                    
                    print(f"[통계표] GRDP JSON에서 {current_key} 데이터 로드 완료")
        except Exception as e:
            print(f"[통계표] GRDP JSON 로드 실패: {e}")
        
        return {
            "title": "분기 지역내총생산(GRDP)",
            "unit": "[전년동기비, %]",
            "page_number_1": 42,
            "page_number_2": 43,
            "data": {
                "yearly": yearly,
                "quarterly": quarterly,
                "yearly_years": yearly_years,
                "quarterly_keys": quarterly_keys
            }
        }
    
    def render_html(self, output_path: str, year: int = 2025, quarter: int = 2) -> str:
        """HTML 렌더링"""
        data = self.extract_all_tables(year, quarter)
        
        # Jinja2 환경 설정
        template_dir = Path(__file__).parent
        env = Environment(loader=FileSystemLoader(str(template_dir)))
        
        # 커스텀 필터 추가
        def round_value(value):
            if value is None or value == "-":
                return "-"
            if value == "N/A":
                return "N/A"
            try:
                return f"{float(value):.1f}"
            except (ValueError, TypeError):
                return str(value)
        
        env.filters['round_value'] = round_value
        
        template = env.get_template("statistics_table_template.html")
        html_content = template.render(**data)
        
        # 파일 저장
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"통계표가 생성되었습니다: {output_path}")
        return html_content
    
    def export_data_json(self, output_path: str, year: int = 2025, quarter: int = 2):
        """데이터를 JSON으로 내보내기"""
        data = self.extract_all_tables(year, quarter)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        print(f"데이터가 저장되었습니다: {output_path}")
    
    def create_historical_template(self, output_path: str):
        """과거 데이터 템플릿 JSON 생성"""
        template = {}
        
        for table_name in self.TABLE_CONFIG.keys():
            template[table_name] = {
                "yearly": {},
                "quarterly": {}
            }
            
            # 과거 연도 (2017-2020)
            for year in ["2017", "2018", "2019", "2020"]:
                template[table_name]["yearly"][year] = {
                    region: None for region in self.ALL_REGIONS
                }
            
            # 과거 분기 (2016.4/4 ~ 2023.1/4)
            quarters = []
            for y in range(2016, 2024):
                for q in ["1/4", "2/4", "3/4", "4/4"]:
                    if y == 2016 and q != "4/4":
                        continue
                    if y == 2023 and q == "2/4":
                        break
                    quarters.append(f"{y}.{q}")
            
            for quarter in quarters:
                template[table_name]["quarterly"][quarter] = {
                    region: None for region in self.ALL_REGIONS
                }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(template, f, ensure_ascii=False, indent=2)
        
        print(f"과거 데이터 템플릿이 생성되었습니다: {output_path}")
    
    def update_historical_json(self, save_path: Optional[str] = None):
        """새로 추출한 데이터를 JSON 파일에 업데이트
        
        Args:
            save_path: 저장할 경로 (기본: 기존 JSON 파일)
        """
        if not self.newly_extracted_data:
            print("[통계표] 업데이트할 새 데이터가 없습니다.")
            return
        
        # 기존 JSON 로드
        historical_path = Path(__file__).parent / self.HISTORICAL_DATA_FILE
        
        try:
            with open(historical_path, 'r', encoding='utf-8') as f:
                existing_data = json.load(f)
        except Exception as e:
            print(f"[통계표] 기존 JSON 로드 실패: {e}")
            existing_data = {'_metadata': {}}
        
        # 메타데이터 분리
        metadata = existing_data.pop('_metadata', {})
        
        # 새 데이터 병합
        new_quarters_added = []
        new_years_added = []
        
        for table_name, new_data in self.newly_extracted_data.items():
            if table_name not in existing_data:
                existing_data[table_name] = {'yearly': {}, 'quarterly': {}}
            
            # 분기별 데이터 병합
            for quarter, regions in new_data.get('quarterly', {}).items():
                if quarter not in existing_data[table_name].get('quarterly', {}):
                    existing_data[table_name]['quarterly'][quarter] = regions
                    if quarter not in new_quarters_added:
                        new_quarters_added.append(quarter)
            
            # 연도별 데이터 병합
            for year, regions in new_data.get('yearly', {}).items():
                if year not in existing_data[table_name].get('yearly', {}):
                    existing_data[table_name]['yearly'][year] = regions
                    if year not in new_years_added:
                        new_years_added.append(year)
        
        # 메타데이터 업데이트
        if new_quarters_added:
            # 분기 범위 업데이트
            all_quarters = []
            for table_data in existing_data.values():
                if isinstance(table_data, dict) and 'quarterly' in table_data:
                    all_quarters.extend(table_data['quarterly'].keys())
            
            if all_quarters:
                sorted_quarters = sorted(set(all_quarters), key=lambda x: (
                    int(re.match(r'(\d{4})', x).group(1)),
                    int(re.match(r'\d{4}\.(\d)', x).group(1))
                ))
                metadata['quarterly_range'] = {
                    'start': sorted_quarters[0].rstrip('p'),
                    'end': sorted_quarters[-1].rstrip('p')
                }
        
        if new_years_added:
            # 연도 범위 업데이트
            all_years = []
            for table_data in existing_data.values():
                if isinstance(table_data, dict) and 'yearly' in table_data:
                    all_years.extend(table_data['yearly'].keys())
            
            if all_years:
                sorted_years = sorted(set(all_years))
                metadata['yearly_range'] = {
                    'start': sorted_years[0],
                    'end': sorted_years[-1]
                }
        
        # 업데이트 일시 기록
        from datetime import datetime
        metadata['last_updated'] = datetime.now().strftime('%Y-%m-%d')
        
        # 저장
        output_data = {'_metadata': metadata, **existing_data}
        target_path = save_path or historical_path
        
        with open(target_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)
        
        print(f"[통계표] JSON 업데이트 완료: {target_path}")
        if new_quarters_added:
            print(f"[통계표] - 추가된 분기: {sorted(new_quarters_added)}")
        if new_years_added:
            print(f"[통계표] - 추가된 연도: {sorted(new_years_added)}")
        
        # 업데이트 완료 후 클리어
        self.newly_extracted_data = {}
    
    def extract_and_save_all(self, year: int = None, quarter: int = None):
        """모든 통계표 데이터 추출 후 JSON 업데이트
        
        새로운 분기 데이터가 있으면 JSON 파일에 자동 저장
        """
        if year is None:
            year = self.current_year
        if quarter is None:
            quarter = self.current_quarter
        
        # 모든 테이블 데이터 추출
        result = self.extract_all_tables(year, quarter)
        
        # 새 데이터가 있으면 JSON 업데이트
        if self.newly_extracted_data:
            self.update_historical_json()
        
        return result


def main():
    """메인 실행 함수"""
    import argparse
    
    parser = argparse.ArgumentParser(description='통계표 생성기')
    parser.add_argument('--excel', '-e', required=True, help='분석표 엑셀 파일 경로')
    parser.add_argument('--output', '-o', help='출력 HTML 파일 경로')
    parser.add_argument('--json', '-j', help='데이터 JSON 출력 경로')
    parser.add_argument('--historical', help='과거 데이터 JSON 파일 경로')
    parser.add_argument('--create-historical-template', help='과거 데이터 템플릿 생성')
    parser.add_argument('--year', '-y', type=int, default=2025, help='보도자료 연도')
    parser.add_argument('--quarter', '-q', type=int, default=2, help='보도자료 분기')
    
    args = parser.parse_args()
    
    generator = StatisticsTableGenerator(args.excel, args.historical)
    
    if args.create_historical_template:
        generator.create_historical_template(args.create_historical_template)
        return
    
    if args.json:
        generator.export_data_json(args.json, args.year, args.quarter)
    
    if args.output:
        generator.render_html(args.output, args.year, args.quarter)


if __name__ == '__main__':
    main()

