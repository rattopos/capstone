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


class 통계표Generator:
    """통계표 생성 클래스"""
    
    # 통계표 항목 정의 (시트별 컬럼 매핑 포함)
    TABLE_CONFIG = {
        "광공업생산지수": {
            "분석_시트": "A 분석",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 6, "값": "BCD"},
            "지역_컬럼": 3,
            "분류단계_컬럼": 4,
            "연도_컬럼": {"2021": 9, "2022": 10, "2023": 11, "2024": 12},
            "분기_컬럼": {
                "2023_2Q": 13, "2023_3Q": 14, "2023_4Q": 15,
                "2024_1Q": 16, "2024_2Q": 17, "2024_3Q": 18, "2024_4Q": 19,
                "2025_1Q": 20, "2025_2Q": 21
            }
        },
        "서비스업생산지수": {
            "분석_시트": "B 분석",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 6, "값": "E~S"},
            "지역_컬럼": 3,
            "분류단계_컬럼": 4,
            "연도_컬럼": {"2021": 8, "2022": 9, "2023": 10, "2024": 11},
            "분기_컬럼": {
                "2023_2Q": 12, "2023_3Q": 13, "2023_4Q": 14,
                "2024_1Q": 15, "2024_2Q": 16, "2024_3Q": 17, "2024_4Q": 18,
                "2025_1Q": 19, "2025_2Q": 20
            }
        },
        "소매판매액지수": {
            "분석_시트": "C 분석",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 6, "값": "A0"},
            "지역_컬럼": 3,
            "분류단계_컬럼": 4,
            "연도_컬럼": {"2021": 8, "2022": 9, "2023": 10, "2024": 11},
            "분기_컬럼": {
                "2023_2Q": 12, "2023_3Q": 13, "2023_4Q": 14,
                "2024_1Q": 15, "2024_2Q": 16, "2024_3Q": 17, "2024_4Q": 18,
                "2025_1Q": 19, "2025_2Q": 20
            }
        },
        "건설수주액": {
            "분석_시트": "F'분석",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 3, "값": "0"},  # 분류단계가 0인 행
            "지역_컬럼": 2,
            "분류단계_컬럼": 3,
            "연도_컬럼": {"2021": 7, "2022": 8, "2023": 9, "2024": 10},
            "분기_컬럼": {
                "2023_2Q": 11, "2023_3Q": 12, "2023_4Q": 13,
                "2024_1Q": 14, "2024_2Q": 15, "2024_3Q": 16, "2024_4Q": 17,
                "2025_1Q": 18, "2025_2Q": 19
            }
        },
        "고용률": {
            "분석_시트": "D(고용률)분석",
            "단위": "[전년동기비, %p]",
            "총지수_식별": {"컬럼": 3, "값": "0"},
            "지역_컬럼": 2,
            "분류단계_컬럼": 3,
            "연도_컬럼": {"2021": 6, "2022": 7, "2023": 8, "2024": 9},
            "분기_컬럼": {
                "2023_2Q": 10, "2023_3Q": 11, "2023_4Q": 12,
                "2024_1Q": 13, "2024_2Q": 14, "2024_3Q": 15, "2024_4Q": 16,
                "2025_1Q": 17, "2025_2Q": 18
            }
        },
        "실업률": {
            "분석_시트": "D(실업)분석",
            "단위": "[전년동기비, %p]",
            "총지수_식별": {"컬럼": 1, "값": "계"},
            "지역_컬럼": 0,
            "분류단계_컬럼": None,
            "연도_컬럼": {"2021": 2, "2022": 3, "2023": 4, "2024": 5},
            "분기_컬럼": {
                "2023_2Q": 6, "2023_3Q": 7, "2023_4Q": 8,
                "2024_1Q": 9, "2024_2Q": 10, "2024_3Q": 11, "2024_4Q": 12,
                "2025_1Q": 13, "2025_2Q": 14
            }
        },
        "국내인구이동": {
            "분석_시트": "I(순인구이동)집계",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 4, "값": "0"},
            "지역_컬럼": 3,
            "분류단계_컬럼": 4,
            "연도_컬럼": {"2021": 6, "2022": 7, "2023": 8, "2024": 9},
            "분기_컬럼": {
                "2023_2Q": 12, "2023_3Q": 13, "2023_4Q": 14,
                "2024_1Q": 15, "2024_2Q": 16, "2024_3Q": 17, "2024_4Q": 18,
                "2025_1Q": 19, "2025_2Q": 20
            }
        },
        "수출액": {
            "분석_시트": "G 분석",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 4, "값": "0"},
            "지역_컬럼": 3,
            "분류단계_컬럼": 4,
            "연도_컬럼": {"2021": 10, "2022": 11, "2023": 12, "2024": 13},
            "분기_컬럼": {
                "2023_2Q": 14, "2023_3Q": 15, "2023_4Q": 16,
                "2024_1Q": 17, "2024_2Q": 18, "2024_3Q": 19, "2024_4Q": 20,
                "2025_1Q": 21, "2025_2Q": 22
            }
        },
        "수입액": {
            "분석_시트": "H 분석",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 4, "값": "0"},
            "지역_컬럼": 3,
            "분류단계_컬럼": 4,
            "연도_컬럼": {"2021": 10, "2022": 11, "2023": 12, "2024": 13},
            "분기_컬럼": {
                "2023_2Q": 14, "2023_3Q": 15, "2023_4Q": 16,
                "2024_1Q": 17, "2024_2Q": 18, "2024_3Q": 19, "2024_4Q": 20,
                "2025_1Q": 21, "2025_2Q": 22
            }
        },
        "소비자물가지수": {
            "분석_시트": "E(지출목적물가) 분석",
            "단위": "[전년동기비, %]",
            "총지수_식별": {"컬럼": 4, "값": 0},
            "지역_컬럼": 3,
            "분류단계_컬럼": 4,
            "연도_컬럼": {"2021": 9, "2022": 10, "2023": 11, "2024": 12},
            "분기_컬럼": {
                "2023_2Q": 13, "2023_3Q": 14, "2023_4Q": 15,
                "2024_1Q": 16, "2024_2Q": 17, "2024_3Q": 18, "2024_4Q": 19,
                "2025_1Q": 20, "2025_2Q": 21
            }
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
        "국내인구이동": "시도 간 이동"
    }
    
    # 지역 목록 (페이지별)
    PAGE1_REGIONS = ["전국", "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종"]
    PAGE2_REGIONS = ["경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"]
    ALL_REGIONS = PAGE1_REGIONS + PAGE2_REGIONS
    
    def __init__(self, excel_path: str, historical_data_path: Optional[str] = None, 
                 raw_excel_path: Optional[str] = None, current_year: int = 2025, current_quarter: int = 2):
        """
        초기화
        
        Args:
            excel_path: 분석표 엑셀 파일 경로
            historical_data_path: 과거 데이터 JSON 파일 경로 (사용하지 않음, 호환성 유지)
            raw_excel_path: 기초자료 엑셀 파일 경로 (우선 사용)
            current_year: 현재 연도
            current_quarter: 현재 분기
        """
        self.excel_path = excel_path
        self.raw_excel_path = raw_excel_path
        self.current_year = current_year
        self.current_quarter = current_quarter
        self.historical_data_path = historical_data_path
        self.historical_data = {}
        self.cached_sheets = {}
        
        # RawDataExtractor 초기화
        self.raw_extractor = None
        if raw_excel_path and RawDataExtractor and Path(raw_excel_path).exists():
            try:
                self.raw_extractor = RawDataExtractor(raw_excel_path, current_year, current_quarter)
                print(f"[통계표] 기초자료 추출기 초기화 완료: {raw_excel_path}")
            except Exception as e:
                print(f"[통계표] 기초자료 추출기 초기화 실패: {e}")
        
        # 과거 데이터 로드 (하위 호환성, 사용하지 않음)
        if historical_data_path and Path(historical_data_path).exists():
            with open(historical_data_path, 'r', encoding='utf-8') as f:
                self.historical_data = json.load(f)
    
    def _load_sheet(self, sheet_name: str) -> pd.DataFrame:
        """엑셀 시트 로드 (캐싱)"""
        if sheet_name not in self.cached_sheets:
            try:
                self.cached_sheets[sheet_name] = pd.read_excel(
                    self.excel_path,
                    sheet_name=sheet_name,
                    header=None
                )
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
        
        # 분류 단계/구분 설정 확인
        classification_column = config.get("분류단계_컬럼")
        classification_value = str(config.get("총지수_식별", {}).get("값", ""))
        region_column = config.get("지역_컬럼", 1)
        
        # 연도 데이터 추출
        yearly_data = self.raw_extractor.extract_yearly_data(
            raw_sheet_name,
            start_year=2016,
            region_column=region_column,
            classification_column=classification_column,
            classification_value=classification_value if classification_value else None
        )
        
        # 분기 데이터 추출
        quarterly_data = self.raw_extractor.extract_all_quarters(
            raw_sheet_name,
            start_year=2016,
            start_quarter=1,
            region_column=region_column,
            classification_column=classification_column,
            classification_value=classification_value if classification_value else None
        )
        
        # 데이터 형식 변환 (분기 키 형식 통일: "2016 1/4" -> "2016.1/4")
        quarterly_formatted = {}
        for quarter_key, data in quarterly_data.items():
            # "2016 1/4" -> "2016.1/4" 형식으로 변환
            formatted_key = quarter_key.replace(" ", ".")
            quarterly_formatted[formatted_key] = data
        
        return {
            'yearly': yearly_data,
            'quarterly': quarterly_formatted
        }
    
    def _get_total_row(self, df: pd.DataFrame, region: str, config: dict) -> Optional[pd.Series]:
        """특정 지역의 총지수 행 가져오기 (시트별 설정 사용)"""
        region_col = config["지역_컬럼"]
        id_col = config["총지수_식별"]["컬럼"]
        id_val = config["총지수_식별"]["값"]
        
        # 값 타입에 따라 비교
        try:
            mask = (df[region_col] == region) & (df[id_col].astype(str) == str(id_val))
            result = df[mask]
            
            if not result.empty:
                return result.iloc[0]
        except Exception as e:
            print(f"행 검색 오류: {region} - {e}")
        
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
        """모든 연도/분기/지역에 기본값 '-'가 채워진 데이터 구조 생성 (2016년부터)"""
        yearly_years = ["2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"]
        
        # 2016년 1분기부터 현재 분기까지 생성
        quarterly_keys = []
        for year in range(2016, self.current_year + 1):
            start_q = 1 if year > 2016 else 1
            end_q = self.current_quarter if year == self.current_year else 4
            for quarter in range(start_q, end_q + 1):
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
        """특정 통계표의 데이터 추출"""
        config = self.TABLE_CONFIG.get(table_name)
        if not config:
            raise ValueError(f"알 수 없는 통계표: {table_name}")
        
        # 데이터 구조 초기화 - 모든 연도/분기/지역에 기본값 '-' 설정
        data = self._create_empty_table_data()
        
        # 기초자료에서 직접 추출 (우선순위 1)
        if self.raw_extractor:
            raw_sheet_name = self.RAW_SHEET_MAPPING.get(table_name)
            if raw_sheet_name:
                try:
                    print(f"[통계표] 기초자료에서 추출: {table_name} (시트: {raw_sheet_name})")
                    raw_data = self._extract_from_raw_data(raw_sheet_name, config)
                    if raw_data:
                        # 기초자료에서 추출한 데이터로 업데이트
                        for year in raw_data.get('yearly', {}):
                            if year in data['yearly']:
                                data['yearly'][year].update(raw_data['yearly'][year])
                        
                        for quarter in raw_data.get('quarterly', {}):
                            if quarter in data['quarterly']:
                                data['quarterly'][quarter].update(raw_data['quarterly'][quarter])
                        
                        print(f"[통계표] 기초자료 추출 완료: {table_name}")
                        return data
                except Exception as e:
                    print(f"[통계표] 기초자료 추출 실패: {table_name} - {e}")
                    # 실패 시 분석표로 fallback
        
        # 분석표에서 추출 (fallback)
        df = self._load_sheet(config["분석_시트"])
        if df is None:
            print(f"[통계표] 시트를 찾을 수 없음: {config['분석_시트']}")
            # 시트가 없어도 기본 구조 반환 (모든 값 '-')
            return data
        
        # 분석표에 있는 연도별 데이터 추출
        analysis_years = list(config["연도_컬럼"].keys())
        
        for year in data["yearly_years"]:
            for region in self.ALL_REGIONS:
                # 분석표에 있는 데이터
                if year in analysis_years:
                    row = self._get_total_row(df, region, config)
                    if row is not None:
                        col_idx = config["연도_컬럼"].get(year)
                        if col_idx is not None and col_idx < len(row):
                            value = row.iloc[col_idx]
                            if pd.notna(value):
                                try:
                                    data["yearly"][year][region] = round(float(value), 1)
                                except (ValueError, TypeError):
                                    pass  # 기본값 '-' 유지
                
                # 과거 데이터 (historical_data에서)
                elif year in self.historical_data.get(table_name, {}).get("yearly", {}):
                    hist_value = self.historical_data[table_name]["yearly"][year].get(region)
                    if hist_value is not None:
                        data["yearly"][year][region] = hist_value
        
        # 분석표에 있는 연도별 데이터 추출
        analysis_years = list(config["연도_컬럼"].keys())
        
        for year in data["yearly_years"]:
            for region in self.ALL_REGIONS:
                # 분석표에 있는 데이터
                if year in analysis_years:
                    row = self._get_total_row(df, region, config)
                    if row is not None:
                        col_idx = config["연도_컬럼"].get(year)
                        if col_idx is not None and col_idx < len(row):
                            value = row.iloc[col_idx]
                            if pd.notna(value):
                                try:
                                    data["yearly"][year][region] = round(float(value), 1)
                                except (ValueError, TypeError):
                                    pass  # 기본값 '-' 유지
        
        # 분기별 데이터 추출 - 분석표에서 사용 가능한 분기만
        quarterly_mapping = {}
        for quarter_key in data["quarterly_keys"]:
            # "2016.1/4" 형식을 "2016_1Q" 형식으로 변환
            match = re.match(r'(\d{4})\.(\d)/4(p?)', quarter_key)
            if match:
                year = int(match.group(1))
                quarter = int(match.group(2))
                is_provisional = match.group(3) == 'p'
                col_key = f"{year}_{quarter}Q"
                
                # 분석표에 해당 분기가 있는지 확인
                if col_key in config["분기_컬럼"]:
                    quarterly_mapping[quarter_key] = col_key
                else:
                    quarterly_mapping[quarter_key] = None
        
        for quarter_key, col_key in quarterly_mapping.items():
            for region in self.ALL_REGIONS:
                if col_key:
                    # 분석표에 있는 데이터
                    row = self._get_total_row(df, region, config)
                    if row is not None:
                        col_idx = config["분기_컬럼"].get(col_key)
                        if col_idx is not None and col_idx < len(row):
                            value = row.iloc[col_idx]
                            if pd.notna(value):
                                try:
                                    data["quarterly"][quarter_key][region] = round(float(value), 1)
                                except (ValueError, TypeError):
                                    pass  # 기본값 '-' 유지
        
        # quarterly_keys 정리 (데이터가 있는 분기만 유지)
        data["quarterly_keys"] = [
            q for q in data["quarterly_keys"] 
            if any(v != "-" for v in data["quarterly"].get(q, {}).values())
        ]
        
        return data
    
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
        """GRDP 데이터 플레이스홀더 생성 (N/A) - 2016년부터"""
        yearly_years = ["2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024"]
        
        # 2016년 1분기부터 현재 분기까지 생성
        quarterly_keys = []
        for year in range(2016, self.current_year + 1):
            start_q = 1 if year > 2016 else 1
            end_q = self.current_quarter if year == self.current_year else 4
            for quarter in range(start_q, end_q + 1):
                if year == self.current_year and quarter == self.current_quarter:
                    quarterly_keys.append(f"{year}.{quarter}/4p")
                else:
                    quarterly_keys.append(f"{year}.{quarter}/4")
        
        # 모든 값을 N/A로 설정
        yearly = {}
        for year in yearly_years:
            yearly[year] = {region: "N/A" for region in self.ALL_REGIONS}
        
        quarterly = {}
        for qk in quarterly_keys:
            quarterly[qk] = {region: "N/A" for region in self.ALL_REGIONS}
        
        return {
            "title": "분기 지역내총생산(GRDP)",
            "unit": "[전년동기비, %]",
            "page_number_1": 42,  # 10개 통계표 * 2페이지 + 목차 + 1
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
        
        template = env.get_template("통계표_template.html")
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


def main():
    """메인 실행 함수"""
    import argparse
    
    parser = argparse.ArgumentParser(description='통계표 생성기')
    parser.add_argument('--excel', '-e', required=True, help='분석표 엑셀 파일 경로')
    parser.add_argument('--output', '-o', help='출력 HTML 파일 경로')
    parser.add_argument('--json', '-j', help='데이터 JSON 출력 경로')
    parser.add_argument('--historical', help='과거 데이터 JSON 파일 경로')
    parser.add_argument('--create-historical-template', help='과거 데이터 템플릿 생성')
    parser.add_argument('--year', '-y', type=int, default=2025, help='보고서 연도')
    parser.add_argument('--quarter', '-q', type=int, default=2, help='보고서 분기')
    
    args = parser.parse_args()
    
    generator = 통계표Generator(args.excel, args.historical)
    
    if args.create_historical_template:
        generator.create_historical_template(args.create_historical_template)
        return
    
    if args.json:
        generator.export_data_json(args.json, args.year, args.quarter)
    
    if args.output:
        generator.render_html(args.output, args.year, args.quarter)


if __name__ == '__main__':
    main()

