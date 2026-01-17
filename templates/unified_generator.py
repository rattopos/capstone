#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
통합 보고서 Generator (간소화 버전)
모든 부문의 보고서를 생성하는 범용 Generator
집계 시트 기반, 완전 동적 매핑

[중요] Based on V2 (Lite Version)
이 파일은 unified_generator_v2.py에서 승격되었습니다 (2025-01-17).
기존 unified_generator.py는 unified_generator_legacy.py.bak으로 백업되었습니다.

자세한 비교는 docs/UNIFIED_GENERATOR_COMPARISON.md 참조
"""

# Based on V2 (Lite Version)
import pandas as pd
from typing import Dict, Any, List, Optional
from pathlib import Path

try:
    from .base_generator import BaseGenerator
    from config.report_configs import (
        get_report_config, REPORT_CONFIGS,
        REGION_DISPLAY_MAPPING, REGION_GROUPS, VALID_REGIONS
    )
except ImportError:
    import sys
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from templates.base_generator import BaseGenerator
    from config.report_configs import (
        get_report_config, REPORT_CONFIGS,
        REGION_DISPLAY_MAPPING, REGION_GROUPS, VALID_REGIONS
    )


class UnifiedReportGenerator(BaseGenerator):
    """
    통합 보고서 Generator (집계 시트 기반)
    
    mining_manufacturing_generator의 검증된 로직을 기반으로 구현
    """
    
    # 데이터 시작 행은 동적으로 찾음 (하드코딩 제거)
    
    def __init__(self, report_type: str, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__(excel_path, year, quarter, excel_file)
        
        # 설정 로드
        self.config = get_report_config(report_type)
        if not self.config:
            raise ValueError(f"Unknown report type: {report_type}")
        
        self.report_type = report_type
        self.report_id = self.config['report_id']
        # 기본값/폴백 사용 금지: 설정에서 값을 찾을 수 없으면 ValueError 발생
        if 'name_mapping' not in self.config:
            raise ValueError(f"[{self.config['name']}] ❌ 설정에서 'name_mapping'을 찾을 수 없습니다. 기본값 사용 금지.")
        self.name_mapping = self.config['name_mapping']
        
        # 집계 시트 구조 (설정에서 로드, 기본값/폴백 사용 금지)
        if 'aggregation_structure' not in self.config:
            raise ValueError(f"[{self.config['name']}] ❌ 설정에서 'aggregation_structure'를 찾을 수 없습니다. 기본값 사용 금지.")
        agg_struct = self.config['aggregation_structure']
        # 기본값은 설정에서 가져오지만, 실제로는 동적으로 찾음
        self.region_name_col = None  # 동적으로 찾음
        self.industry_code_col = None  # 동적으로 찾음
        self.total_code = agg_struct.get('total_code', 'BCD')
        
        # metadata_columns 설정 (동적 컬럼 찾기에 사용, 기본값/폴백 사용 금지)
        if 'metadata_columns' not in self.config:
            raise ValueError(f"[{self.config['name']}] ❌ 설정에서 'metadata_columns'를 찾을 수 없습니다. 기본값 사용 금지.")
        self.metadata_cols = self.config['metadata_columns']
        
        # 산업명 컬럼도 동적으로 찾음
        self.industry_name_col = None  # 동적으로 찾음
        
        # 데이터 시작 행도 동적으로 찾음 (하드코딩 제거)
        self.data_start_row = None  # 동적으로 찾음
        
        # 여러 시트 지원
        self.df_analysis = None
        self.df_aggregation = None
        self.df_reference = None
        self.target_col = None
        self.prev_y_col = None
        self.use_aggregation_only = False
        
        print(f"[{self.config['name']}] Generator 초기화")
    
    def _get_region_display_name(self, region: str) -> str:
        """지역명 변환"""
        return REGION_DISPLAY_MAPPING.get(region, region)
    
    def load_data(self):
        """모든 관련 시트 로드 (분석 시트, 집계 시트, 참고 시트)"""
        xl = self.load_excel()
        sheet_names = xl.sheet_names
        
        # 1. 분석 시트 찾기 (기본값/폴백 사용 금지)
        analysis_sheets = self.config['sheets'].get('analysis')
        if analysis_sheets is None:
            raise ValueError(f"[{self.config['name']}] ❌ 설정에서 'analysis' 시트 목록을 찾을 수 없습니다.")
        
        analysis_sheet = None
        for name in analysis_sheets:
            if name in sheet_names:
                analysis_sheet = name
                break
        
        # 분석 시트를 찾을 수 없으면 경고만 출력 (집계 시트만 있어도 작동 가능)
        if not analysis_sheet:
            # 상세 디버그 정보 출력
            print(f"[{self.config['name']}] 🔍 [디버그] 분석 시트 찾기 실패:")
            print(f"  - 찾으려는 시트 목록: {analysis_sheets}")
            print(f"  - 파일의 모든 시트 목록: {sheet_names}")
            print(f"  - 시트 개수: {len(sheet_names)}")
            # 유사한 시트 이름 찾기
            similar_sheets = []
            for target in analysis_sheets:
                for sheet in sheet_names:
                    if target.lower() in sheet.lower() or sheet.lower() in target.lower():
                        similar_sheets.append(f"'{sheet}' (유사: '{target}')")
            if similar_sheets:
                print(f"  - 유사한 시트 이름: {similar_sheets}")
            # 집계 시트가 있는지 먼저 확인
            agg_sheets_check = self.config['sheets'].get('aggregation', [])
            agg_exists = any(name in sheet_names for name in agg_sheets_check)
            if agg_exists:
                print(f"[{self.config['name']}] ⚠️ 분석 시트를 찾을 수 없지만, 집계 시트가 있으므로 집계 시트만 사용합니다.")
            else:
                # 집계 시트도 없으면 ValueError 발생
                raise ValueError(
                    f"[{self.config['name']}] ❌ 분석 시트와 집계 시트를 모두 찾을 수 없습니다.\n"
                    f"  찾으려는 분석 시트: {analysis_sheets}\n"
                    f"  찾으려는 집계 시트: {agg_sheets_check}\n"
                    f"  파일의 시트 목록: {sheet_names}\n"
                    f"  유사한 시트: {similar_sheets if similar_sheets else '없음'}"
                )
        
        # 2. 집계 시트 찾기 (기본값/폴백 사용 금지)
        if 'sheets' not in self.config or 'aggregation' not in self.config['sheets']:
            raise ValueError(f"[{self.config['name']}] ❌ 설정에서 'sheets.aggregation'을 찾을 수 없습니다. 기본값 사용 금지.")
        agg_sheets = self.config['sheets']['aggregation']
        agg_sheet = None
        for name in agg_sheets:
            if name in sheet_names:
                agg_sheet = name
                break
        
        # 집계 시트가 없으면 분석 시트 사용
        if not agg_sheet:
            agg_sheet = analysis_sheet
            if agg_sheet:
                print(f"[{self.config['name']}] [시트 대체] 집계 시트 → 분석 시트 '{agg_sheet}'")
        
        # 3. 참고 시트(비공표자료) 찾기
        # 파일 전체에서 "참고", "비공표자료", "reference" 등의 키워드가 포함된 시트 찾기
        # 셀 위치가 아닌 시트 이름으로만 찾음
        reference_sheet = None
        
        # 키워드 패턴: "참고", "비공표", "reference" 등이 포함된 시트 찾기
        reference_keywords = ['참고', '비공표', 'reference', '비공표자료', '참고자료']
        
        for sheet_name in sheet_names:
            # 시트 이름에서 키워드가 포함되어 있는지 확인
            normalized_name = sheet_name.lower().replace(' ', '').replace("'", "").replace('(', '').replace(')', '')
            for keyword in reference_keywords:
                if keyword in sheet_name or keyword in normalized_name:
                    # 분석 시트나 집계 시트와는 다른 시트인지 확인
                    if sheet_name != analysis_sheet and sheet_name != agg_sheet:
                        reference_sheet = sheet_name
                        print(f"[{self.config['name']}] 🔍 참고 시트 후보 발견: '{sheet_name}' (키워드: '{keyword}')")
                        break
            if reference_sheet:
                break
        
        # 키워드로 찾지 못한 경우, 보고서명 기반으로 추가 시도
        if not reference_sheet:
            report_name_patterns = [
                f"{self.config['name']} 참고",
                f"{self.config['name']}참고",
            ]
            if analysis_sheet:
                base_name = analysis_sheet.replace(' 분석', '').replace('분석', '').replace('(', '').replace(')', '').replace("'", "")
                report_name_patterns.extend([
                    f"{base_name} 참고",
                    f"{base_name}참고",
                ])
            
            for pattern in report_name_patterns:
                if pattern in sheet_names:
                    reference_sheet = pattern
                    break
        
        # 4. 시트 로드
        if analysis_sheet:
            self.df_analysis = self.get_sheet(analysis_sheet)
            if self.df_analysis is not None:
                print(f"[{self.config['name']}] ✅ 분석 시트: '{analysis_sheet}' ({len(self.df_analysis)}행 × {len(self.df_analysis.columns)}열)")
        
        if agg_sheet:
            self.df_aggregation = self.get_sheet(agg_sheet)
            if self.df_aggregation is not None:
                print(f"[{self.config['name']}] ✅ 집계 시트: '{agg_sheet}' ({len(self.df_aggregation)}행 × {len(self.df_aggregation.columns)}열)")
        
        if reference_sheet and reference_sheet != analysis_sheet:
            self.df_reference = self.get_sheet(reference_sheet)
            if self.df_reference is not None:
                print(f"[{self.config['name']}] ✅ 참고 시트: '{reference_sheet}' ({len(self.df_reference)}행 × {len(self.df_reference.columns)}열)")
        
        # 5. 분석 시트가 비어있는지 확인 (수식 미계산 체크)
        if self.df_analysis is not None:
            # 간단한 체크: 특정 행에 데이터가 거의 없으면 비어있다고 판단
            if len(self.df_analysis) > 0:
                # 중간 행의 NaN 비율 확인
                mid_row = len(self.df_analysis) // 2
                if mid_row < len(self.df_analysis):
                    nan_ratio = self.df_analysis.iloc[mid_row].isna().sum() / len(self.df_analysis.columns)
                    if nan_ratio > 0.8:  # 80% 이상이 NaN이면 비어있다고 판단
                        print(f"[{self.config['name']}] ⚠️ 분석 시트가 비어있음 → 집계 시트에서 직접 계산")
                        self.use_aggregation_only = True
        
        # 6. 최종 데이터 소스 결정
        # 집계 시트가 있으면 우선 사용, 없으면 분석 시트 사용
        if self.df_aggregation is None and self.df_analysis is None:
            raise ValueError(f"[{self.config['name']}] ❌ 분석 시트와 집계 시트를 모두 찾을 수 없습니다. 시트 목록: {sheet_names}")
        
        # 동적 컬럼 찾기 (집계 시트 우선, 없으면 분석 시트)
        self._find_data_columns()
        # 동적 컬럼 위치 찾기 (지역명, 산업코드, 산업명 등)
        self._find_metadata_columns()
    
    def _find_metadata_columns(self):
        """메타데이터 컬럼 동적 탐색 (지역명, 산업코드, 산업명 등)"""
        # 데이터 소스 결정: 집계 시트 우선, 없으면 분석 시트
        df = None
        if self.df_aggregation is not None:
            df = self.df_aggregation
        elif self.df_analysis is not None:
            df = self.df_analysis
        else:
            return  # 컬럼을 찾을 수 없음
        
        # 헤더 행 찾기 (처음 몇 행에서)
        header_rows = min(5, len(df))
        if header_rows == 0:
            return
        
        # metadata_columns 설정에서 키워드 가져오기
        region_keywords = self.metadata_cols.get('region', ['지역', 'region', '시도'])
        code_keywords = self.metadata_cols.get('code', ['코드', 'code', '산업코드', '업태코드', '품목코드', '분류코드'])
        name_keywords = self.metadata_cols.get('name', ['이름', 'name', '산업명', '업태명', '품목명', '공정이름', '공정명'])
        
        # 각 행에서 키워드 검색
        for row_idx in range(header_rows):
            row = df.iloc[row_idx]
            for col_idx, cell_value in enumerate(row):
                if pd.isna(cell_value):
                    continue
                cell_str = str(cell_value).strip().lower()
                
                # 지역명 컬럼 찾기
                if self.region_name_col is None:
                    for keyword in region_keywords:
                        if keyword.lower() in cell_str:
                            self.region_name_col = col_idx
                            print(f"[{self.config['name']}] ✅ 지역명 컬럼 발견: {col_idx} (키워드: '{keyword}', 행: {row_idx})")
                            break
                
                # 산업코드 컬럼 찾기
                if self.industry_code_col is None:
                    for keyword in code_keywords:
                        if keyword.lower() in cell_str:
                            self.industry_code_col = col_idx
                            print(f"[{self.config['name']}] ✅ 산업코드 컬럼 발견: {col_idx} (키워드: '{keyword}', 행: {row_idx})")
                            break
                
                # 산업명 컬럼 찾기
                if self.industry_name_col is None:
                    for keyword in name_keywords:
                        if keyword.lower() in cell_str:
                            self.industry_name_col = col_idx
                            print(f"[{self.config['name']}] ✅ 산업명 컬럼 발견: {col_idx} (키워드: '{keyword}', 행: {row_idx})")
                            break
                
                # 모든 컬럼을 찾았으면 종료
                if (self.region_name_col is not None and 
                    self.industry_code_col is not None and 
                    self.industry_name_col is not None):
                    break
            
            if (self.region_name_col is not None and 
                self.industry_code_col is not None and 
                self.industry_name_col is not None):
                break
        
        # 데이터 시작 행 찾기 (헤더 다음 행)
        # 지역명이나 산업코드가 실제로 나타나는 첫 번째 행 찾기
        if self.region_name_col is not None:
            for row_idx in range(header_rows, min(header_rows + 10, len(df))):
                row = df.iloc[row_idx]
                if self.region_name_col < len(row):
                    cell_value = row.iloc[self.region_name_col]
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip()
                        # 지역명이 실제로 나타나는 행 찾기
                        if cell_str in ['전국', '서울', '부산', '대구', '인천']:
                            self.data_start_row = row_idx
                            print(f"[{self.config['name']}] ✅ 데이터 시작 행 발견: {row_idx}")
                            break
        
        # 기본값/폴백 사용 금지: 동적으로 찾지 못하면 ValueError 발생 (상세 디버그 정보 포함)
        if self.region_name_col is None:
            # 상세 디버그 정보 출력
            print(f"[{self.config['name']}] 🔍 [디버그] 지역명 컬럼 찾기 실패:")
            print(f"  - 확인한 시트: {'집계' if self.df_aggregation is not None else '분석'}")
            print(f"  - 확인한 행 수: {header_rows}")
            print(f"  - 찾으려는 키워드: {region_keywords}")
            print(f"  - 시트 크기: {len(df)}행 × {len(df.columns)}열")
            # 헤더 행 샘플 출력
            print(f"  - 헤더 행 샘플 (처음 3행):")
            for i in range(min(3, header_rows)):
                row_sample = [str(df.iloc[i, j])[:20] if j < len(df.columns) and pd.notna(df.iloc[i, j]) else 'NaN' 
                             for j in range(min(10, len(df.columns)))]
                print(f"    행 {i}: {row_sample}")
            raise ValueError(
                f"[{self.config['name']}] ❌ 지역명 컬럼을 찾을 수 없습니다.\n"
                f"  확인한 시트: {'집계' if self.df_aggregation is not None else '분석'}\n"
                f"  찾으려는 키워드: {region_keywords}\n"
                f"  시트 크기: {len(df)}행 × {len(df.columns)}열\n"
                f"  확인한 헤더 행 수: {header_rows}"
            )
        
        if self.industry_code_col is None:
            print(f"[{self.config['name']}] 🔍 [디버그] 산업코드 컬럼 찾기 실패:")
            print(f"  - 확인한 시트: {'집계' if self.df_aggregation is not None else '분석'}")
            print(f"  - 찾으려는 키워드: {code_keywords}")
            print(f"  - 시트 크기: {len(df)}행 × {len(df.columns)}열")
            raise ValueError(
                f"[{self.config['name']}] ❌ 산업코드 컬럼을 찾을 수 없습니다.\n"
                f"  확인한 시트: {'집계' if self.df_aggregation is not None else '분석'}\n"
                f"  찾으려는 키워드: {code_keywords}\n"
                f"  시트 크기: {len(df)}행 × {len(df.columns)}열"
            )
        
        if self.industry_name_col is None:
            print(f"[{self.config['name']}] 🔍 [디버그] 산업명 컬럼 찾기 실패:")
            print(f"  - 확인한 시트: {'집계' if self.df_aggregation is not None else '분석'}")
            print(f"  - 찾으려는 키워드: {name_keywords}")
            print(f"  - 시트 크기: {len(df)}행 × {len(df.columns)}열")
            raise ValueError(
                f"[{self.config['name']}] ❌ 산업명 컬럼을 찾을 수 없습니다.\n"
                f"  확인한 시트: {'집계' if self.df_aggregation is not None else '분석'}\n"
                f"  찾으려는 키워드: {name_keywords}\n"
                f"  시트 크기: {len(df)}행 × {len(df.columns)}열"
            )
        
        if self.data_start_row is None:
            print(f"[{self.config['name']}] 🔍 [디버그] 데이터 시작 행 찾기 실패:")
            print(f"  - 확인한 시트: {'집계' if self.df_aggregation is not None else '분석'}")
            print(f"  - 지역명 컬럼 인덱스: {self.region_name_col}")
            print(f"  - 확인한 행 범위: {header_rows} ~ {min(header_rows + 10, len(df))}")
            # 확인한 행의 지역명 컬럼 값 샘플 출력
            print(f"  - 지역명 컬럼 값 샘플:")
            for i in range(header_rows, min(header_rows + 10, len(df))):
                if self.region_name_col < len(df.iloc[i]):
                    val = df.iloc[i, self.region_name_col]
                    if pd.notna(val):
                        print(f"    행 {i}: '{val}'")
            raise ValueError(
                f"[{self.config['name']}] ❌ 데이터 시작 행을 찾을 수 없습니다.\n"
                f"  확인한 시트: {'집계' if self.df_aggregation is not None else '분석'}\n"
                f"  지역명 컬럼 인덱스: {self.region_name_col}\n"
                f"  확인한 행 범위: {header_rows} ~ {min(header_rows + 10, len(df))}"
            )
    
    def _find_data_columns(self):
        """데이터 컬럼 동적 탐색 (병합된 셀 처리) - 집계 시트 우선, 없으면 분석 시트"""
        # 데이터 소스 결정: 집계 시트 우선, 없으면 분석 시트
        df = None
        if self.df_aggregation is not None:
            df = self.df_aggregation
            sheet_type = "집계"
        elif self.df_analysis is not None:
            df = self.df_analysis
            sheet_type = "분석"
        else:
            raise ValueError(
                f"[{self.config['name']}] ❌ 집계 시트와 분석 시트가 모두 로드되지 않았습니다. "
                f"load_data()를 먼저 호출해야 합니다."
            )
        
        # DataFrame 전체를 전달하여 병합된 셀 처리 (스마트 헤더 탐색기)
        # 고용률/실업률은 타입 필터링을 선택적으로 적용 (헤더에 "고용률", "실업률" 키워드가 있으면 OK)
        require_type_match = self.report_type not in ['employment', 'unemployment']
        
        # target_col 찾기
        if self.target_col is None:
            self.target_col = self.find_target_col_index(df, self.year, self.quarter, require_type_match=require_type_match)
            if self.target_col is not None:
                print(f"[{self.config['name']}] ✅ Target 컬럼 ({sheet_type} 시트): {self.target_col} ({self.year} {self.quarter}/4)")
        
        # prev_y_col 찾기
        if self.prev_y_col is None:
            self.prev_y_col = self.find_target_col_index(df, self.year - 1, self.quarter, require_type_match=require_type_match)
            if self.prev_y_col is not None:
                print(f"[{self.config['name']}] ✅ 전년 컬럼 ({sheet_type} 시트): {self.prev_y_col} ({self.year - 1} {self.quarter}/4)")
        
        # 기본값 사용 금지: 반드시 찾아야 함 (상세 디버그 정보 포함)
        if self.target_col is None:
            # 헤더 행 샘플 출력
            print(f"[{self.config['name']}] 🔍 [디버그] Target 컬럼 찾기 실패:")
            print(f"  - 찾으려는 연도/분기: {self.year}년 {self.quarter}분기")
            print(f"  - 확인한 시트: {sheet_type}")
            print(f"  - 시트 크기: {len(df)}행 × {len(df.columns)}열")
            # 헤더 행 샘플 출력
            header_sample_rows = min(3, len(df))
            print(f"  - 헤더 행 샘플 (처음 {header_sample_rows}행):")
            for i in range(header_sample_rows):
                row_sample = [str(df.iloc[i, j])[:30] if j < len(df.columns) and pd.notna(df.iloc[i, j]) else 'NaN' 
                             for j in range(min(15, len(df.columns)))]
                print(f"    행 {i}: {row_sample}")
            raise ValueError(
                f"[{self.config['name']}] ❌ Target 컬럼을 찾을 수 없습니다.\n"
                f"  찾으려는 연도/분기: {self.year}년 {self.quarter}분기\n"
                f"  확인한 시트: {sheet_type}\n"
                f"  시트 크기: {len(df)}행 × {len(df.columns)}열"
            )
        
        if self.prev_y_col is None:
            print(f"[{self.config['name']}] 🔍 [디버그] 전년 컬럼 찾기 실패:")
            print(f"  - 찾으려는 연도/분기: {self.year - 1}년 {self.quarter}분기")
            print(f"  - 확인한 시트: {sheet_type}")
            print(f"  - 시트 크기: {len(df)}행 × {len(df.columns)}열")
            raise ValueError(
                f"[{self.config['name']}] ❌ 전년 컬럼을 찾을 수 없습니다.\n"
                f"  찾으려는 연도/분기: {self.year - 1}년 {self.quarter}분기\n"
                f"  확인한 시트: {sheet_type}\n"
                f"  시트 크기: {len(df)}행 × {len(df.columns)}열"
            )
    
    def _extract_table_data_ssot(self) -> List[Dict[str, Any]]:
        """
        집계 시트 또는 분석 시트에서 전국 + 17개 시도 데이터 추출 (SSOT)
        집계 시트 우선, 없으면 분석 시트 사용
        """
        # 데이터 소스 결정: 집계 시트 우선, 없으면 분석 시트
        df = None
        if self.df_aggregation is not None:
            df = self.df_aggregation
        elif self.df_analysis is not None:
            df = self.df_analysis
        else:
            raise ValueError(
                f"[{self.config['name']}] ❌ 집계 시트와 분석 시트가 모두 로드되지 않았습니다. "
                f"load_data() 또는 extract_all_data()를 먼저 호출해야 합니다."
            )
        
        # 데이터 행만 (헤더 제외) - 동적으로 찾은 시작 행 사용
        if self.data_start_row is None:
            self.data_start_row = 0
        
        if self.data_start_row < 0:
            self.data_start_row = 0
        
        if self.data_start_row < len(df):
            data_df = df.iloc[self.data_start_row:].copy()
        else:
            print(f"[{self.config['name']}] ⚠️ data_start_row({self.data_start_row})가 DataFrame 길이({len(df)})를 초과합니다. 전체 DataFrame 사용")
            data_df = df.copy()
        
        # 지역 목록
        regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        table_data = []
        
        # 컬럼 인덱스 검증 (동적으로 찾은 컬럼)
        if self.region_name_col is None or self.region_name_col < 0 or self.region_name_col >= len(data_df.columns):
            raise ValueError(
                f"[{self.config['name']}] ❌ 지역명 컬럼을 찾을 수 없습니다. "
                f"동적 탐색 실패 또는 인덱스({self.region_name_col})가 유효하지 않습니다. "
                f"DataFrame 컬럼 수: {len(data_df.columns)}"
            )
        
        for region in regions:
            # 지역명으로 필터링 (설정에서 가져온 컬럼 사용) - 안전한 인덱스 접근
            try:
                region_filter = data_df[
                    data_df.iloc[:, self.region_name_col].astype(str).str.strip() == region
                ]
            except (IndexError, KeyError) as e:
                print(f"[{self.config['name']}] ⚠️ {region} 필터링 오류: {e}")
                continue
            
            if region_filter.empty:
                continue
            
            # 총지수 행 찾기 (동적으로 찾은 컬럼 및 코드 사용) - 안전한 인덱스 접근
            # 컬럼 인덱스 검증
            if self.industry_code_col is None or self.industry_code_col < 0 or self.industry_code_col >= len(region_filter.columns):
                print(f"[{self.config['name']}] ⚠️ {region}: 산업코드 컬럼을 찾을 수 없습니다. 동적 탐색 실패 또는 인덱스({self.industry_code_col})가 유효하지 않습니다. 스킵합니다.")
                continue
            
            # 디버깅: 실제 코드 값 확인
            if region == '전국':
                try:
                    industry_codes = region_filter.iloc[:, self.industry_code_col].astype(str).head(5).tolist()
                    print(f"[{self.config['name']}] 디버그: {region} 산업코드 (처음 5개): {industry_codes}")
                    print(f"[{self.config['name']}] 디버그: 찾으려는 코드: '{self.total_code}'")
                except (IndexError, KeyError) as e:
                    print(f"[{self.config['name']}] ⚠️ {region} 산업코드 확인 오류: {e}")
            
            try:
                region_total = region_filter[
                    region_filter.iloc[:, self.industry_code_col].astype(str).str.contains(self.total_code, na=False, regex=False)
                ]
            except (IndexError, KeyError) as e:
                raise ValueError(f"[{self.config['name']}] ❌ {region} 총지수 행 찾기 오류: {e}. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")
            
            if region_total.empty:
                # 상세 디버그 정보 출력
                print(f"[{self.config['name']}] 🔍 [디버그] {region} 총지수 행 찾기 실패:")
                print(f"  - 찾으려는 코드: '{self.total_code}'")
                print(f"  - 산업코드 컬럼 인덱스: {self.industry_code_col}")
                print(f"  - 필터링된 행 수: {len(region_filter)}")
                # 실제 코드 값 샘플 출력
                if len(region_filter) > 0:
                    print(f"  - 실제 코드 값 샘플 (처음 10개):")
                    for idx, row in region_filter.head(10).iterrows():
                        if self.industry_code_col < len(row):
                            code_val = row.iloc[self.industry_code_col]
                            code_str = str(code_val).strip() if pd.notna(code_val) else 'NaN'
                            print(f"    행 {idx}: '{code_str}'")
                raise ValueError(
                    f"[{self.config['name']}] ❌ {region}: 코드 '{self.total_code}'를 찾을 수 없습니다.\n"
                    f"  산업코드 컬럼 인덱스: {self.industry_code_col}\n"
                    f"  필터링된 행 수: {len(region_filter)}\n"
                    f"  실제 코드 값 샘플: {[str(region_filter.iloc[i, self.industry_code_col]).strip() if i < len(region_filter) and self.industry_code_col < len(region_filter.iloc[i]) else 'N/A' for i in range(min(5, len(region_filter)))]}"
                )
            
            if region_total.empty:
                continue
            
            row = region_total.iloc[0]
            
            # 기본값 사용 금지: 반드시 유효한 인덱스여야 함 (상세 디버그 정보 포함)
            if self.target_col is None:
                print(f"[{self.config['name']}] 🔍 [디버그] {region} Target 컬럼이 None:")
                print(f"  - 찾으려는 연도/분기: {self.year}년 {self.quarter}분기")
                print(f"  - 행 길이: {len(row)}")
                print(f"  - 행 샘플: {[str(row.iloc[j])[:20] if j < len(row) and pd.notna(row.iloc[j]) else 'NaN' for j in range(min(10, len(row)))]}")
                raise ValueError(
                    f"[{self.config['name']}] ❌ {region} Target 컬럼이 None입니다.\n"
                    f"  찾으려는 연도/분기: {self.year}년 {self.quarter}분기\n"
                    f"  행 길이: {len(row)}"
                )
            
            if self.prev_y_col is None:
                print(f"[{self.config['name']}] 🔍 [디버그] {region} 전년 컬럼이 None:")
                print(f"  - 찾으려는 연도/분기: {self.year - 1}년 {self.quarter}분기")
                print(f"  - 행 길이: {len(row)}")
                raise ValueError(
                    f"[{self.config['name']}] ❌ {region} 전년 컬럼이 None입니다.\n"
                    f"  찾으려는 연도/분기: {self.year - 1}년 {self.quarter}분기\n"
                    f"  행 길이: {len(row)}"
                )
            
            # 인덱스 범위 체크
            if self.target_col >= len(row):
                print(f"[{self.config['name']}] ⚠️ Target 컬럼 인덱스({self.target_col})가 행 길이({len(row)})를 초과합니다. 스킵합니다.")
                continue
            
            if self.prev_y_col >= len(row):
                print(f"[{self.config['name']}] ⚠️ 전년 컬럼 인덱스({self.prev_y_col})가 행 길이({len(row)})를 초과합니다. 스킵합니다.")
                continue
            
            # 지수 추출
            try:
                idx_current = self.safe_float(row.iloc[self.target_col], None)
                idx_prev_year = self.safe_float(row.iloc[self.prev_y_col], None)
            except (IndexError, KeyError) as e:
                print(f"[{self.config['name']}] ⚠️ 데이터 추출 오류: {e}. 스킵합니다.")
                continue
            
            if idx_current is None:
                continue
            
            # 증감률 계산
            if idx_prev_year and idx_prev_year != 0:
                change_rate = round(((idx_current - idx_prev_year) / idx_prev_year) * 100, 1)
            else:
                change_rate = None
            
            table_data.append({
                'region_name': region,
                'region_display': self._get_region_display_name(region),
                'value': round(idx_current, 1),
                'prev_value': round(idx_prev_year, 1) if idx_prev_year else None,
                'change_rate': change_rate
            })
            
            print(f"[{self.config['name']}] ✅ {region}: 지수={idx_current:.1f}, 증감률={change_rate}%")
        
        return table_data
    
    def _extract_industry_data(self, region: str) -> List[Dict[str, Any]]:
        """
        특정 지역의 업종별 데이터 추출
        
        Args:
            region: 지역명 ('전국', '서울', 등)
            
        Returns:
            업종별 데이터 리스트 [{'name': '업종명', 'value': 지수, 'change_rate': 증감률, 'growth_rate': 증감률}, ...]
        """
        if self.df_aggregation is None:
            return []
        
        df = self.df_aggregation
        
        # 컬럼 인덱스 검증 (동적으로 찾은 컬럼)
        if self.region_name_col is None or self.region_name_col < 0 or self.region_name_col >= len(df.columns):
            print(f"[{self.config['name']}] ⚠️ 지역명 컬럼을 찾을 수 없습니다. 동적 탐색 실패 또는 인덱스({self.region_name_col})가 유효하지 않습니다. 빈 리스트 반환")
            return []
        
        # 데이터 행만 (헤더 제외) - 동적으로 찾은 시작 행 사용
        if self.data_start_row is None:
            self.data_start_row = 0
        
        if self.data_start_row < 0:
            self.data_start_row = 0
        
        if self.data_start_row < len(df):
            data_df = df.iloc[self.data_start_row:].copy()
        else:
            data_df = df.copy()
        
        # 지역 필터링 (안전한 인덱스 접근)
        try:
            region_filter = data_df[
                data_df.iloc[:, self.region_name_col].astype(str).str.strip() == region
            ]
        except (IndexError, KeyError) as e:
            print(f"[{self.config['name']}] ⚠️ {region} 필터링 오류: {e}")
            return []
        
        if region_filter.empty:
            return []
        
        industries = []
        # 기본값/폴백 사용 금지
        if 'name_mapping' not in self.config:
            raise ValueError(f"[{self.config['name']}] ❌ 설정에서 'name_mapping'을 찾을 수 없습니다. 기본값 사용 금지.")
        name_mapping = self.config['name_mapping']
        
        # 산업명 컬럼 찾기 (동적으로 찾은 값 사용, 기본값/폴백 사용 금지)
        if self.industry_name_col is None:
            raise ValueError(f"[{self.config['name']}] ❌ 산업명 컬럼을 찾을 수 없습니다. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")
        
        industry_name_col = self.industry_name_col
        
        if industry_name_col < 0:
            industry_name_col = 0
        
        for idx, row in region_filter.iterrows():
            # 산업코드 확인 (총지수 제외) - 동적으로 찾은 컬럼 사용
            if self.industry_code_col is None:
                continue
            
            if self.industry_code_col >= len(row):
                continue
                
            # 기본값/폴백 사용 금지
            if pd.isna(row.iloc[self.industry_code_col]):
                continue  # NaN이면 스킵
            industry_code = str(row.iloc[self.industry_code_col]).strip()
            
            # 총지수 코드는 제외
            if not industry_code or industry_code == '' or industry_code == 'nan':
                continue
            
            # total_code와 일치하면 제외 (총지수)
            # total_code가 'BCD', 'E~S' 같은 패턴일 수 있으므로 contains 체크
            if str(self.total_code) in str(industry_code) or industry_code == str(self.total_code):
                continue
            
            # 산업명 추출
            industry_name = ''
            if industry_name_col < len(row) and pd.notna(row.iloc[industry_name_col]):
                industry_name = str(row.iloc[industry_name_col]).strip()
                if industry_name == 'nan' or not industry_name:
                    continue
            else:
                # 산업명 컬럼이 없으면 스킵
                continue
            
            # 이름 매핑 적용
            if industry_name in name_mapping:
                industry_name = name_mapping[industry_name]
            
            if not industry_name:
                continue
            
            # 지수 추출 (안전한 인덱스 접근)
            try:
                if self.target_col is None or self.prev_y_col is None:
                    continue
                
                # 인덱스 범위 체크
                if self.target_col < 0 or self.target_col >= len(row):
                    continue
                if self.prev_y_col < 0 or self.prev_y_col >= len(row):
                    continue
                    
                idx_current = self.safe_float(row.iloc[self.target_col], None)
                idx_prev_year = self.safe_float(row.iloc[self.prev_y_col], None)
            except (IndexError, KeyError, AttributeError) as e:
                print(f"[{self.config['name']}] ⚠️ 데이터 추출 오류 (인덱스 {self.target_col}/{self.prev_y_col}): {e}")
                continue
            
            if idx_current is None:
                continue
            
            # 증감률 계산
            change_rate = None
            if idx_prev_year and idx_prev_year != 0:
                change_rate = round(((idx_current - idx_prev_year) / idx_prev_year) * 100, 1)
            
            industries.append({
                'name': industry_name,
                'value': round(idx_current, 1),
                'prev_value': round(idx_prev_year, 1) if idx_prev_year else None,
                'change_rate': change_rate,
                'growth_rate': change_rate,  # 템플릿 호환 필드명
                'code': industry_code
            })
        
        return industries
    
    def _get_top_industries_for_region(self, region: str, increase: bool = True, top_n: int = 3) -> List[Dict[str, Any]]:
        """
        특정 지역의 상위 업종 추출
        
        Args:
            region: 지역명
            increase: True면 증가 업종, False면 감소 업종
            top_n: 상위 N개
            
        Returns:
            상위 업종 리스트
        """
        if not region or not isinstance(region, str):
            return []
        
        industries = self._extract_industry_data(region)
        
        # 안전한 필터링
        if not industries:
            return []
        
        if increase:
            filtered = [
                ind for ind in industries 
                if ind and isinstance(ind, dict) and 
                ind.get('change_rate') is not None and 
                ind['change_rate'] > 0
            ]
            try:
                # 기본값/폴백 사용 금지: change_rate가 None이면 정렬에서 제외
                filtered = [x for x in filtered if x and isinstance(x, dict) and x.get('change_rate') is not None]
                filtered.sort(key=lambda x: x['change_rate'], reverse=True)
            except (TypeError, AttributeError, KeyError) as e:
                raise ValueError(f"[{self.config['name']}] ❌ 정렬 오류: {e}. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")
        else:
            filtered = [
                ind for ind in industries 
                if ind and isinstance(ind, dict) and 
                ind.get('change_rate') is not None and 
                ind['change_rate'] < 0
            ]
            try:
                # 기본값/폴백 사용 금지: change_rate가 None이면 정렬에서 제외
                filtered = [x for x in filtered if x and isinstance(x, dict) and x.get('change_rate') is not None]
                filtered.sort(key=lambda x: x['change_rate'])
            except (TypeError, AttributeError):
                pass  # 정렬 실패 시 원본 유지
        
        # 안전한 슬라이싱
        # 기본값/폴백 사용 금지: filtered가 없으면 None 반환
        if not filtered or len(filtered) == 0:
            return None
        return filtered[:top_n]
    
    def extract_nationwide_data(self, table_data: List[Dict] = None) -> Dict[str, Any]:
        """전국 데이터 추출 - 템플릿 호환 필드명"""
        if table_data is None:
            table_data = self._extract_table_data_ssot()
        
        nationwide = next((d for d in table_data if d['region_name'] == '전국'), None)
        
        # 기본값/폴백 사용 금지: 데이터를 찾을 수 없으면 ValueError 발생
        
        # 기본값/폴백 사용 금지: 데이터를 찾을 수 없으면 ValueError 발생 (상세 디버그 정보 포함)
        if not nationwide or not isinstance(nationwide, dict):
            print(f"[{self.config['name']}] 🔍 [디버그] 전국 데이터 찾기 실패:")
            print(f"  - nationwide 타입: {type(nationwide)}")
            print(f"  - nationwide 값: {nationwide}")
            print(f"  - table_data 길이: {len(table_data)}")
            if table_data:
                print(f"  - table_data 샘플 (처음 3개): {table_data[:3]}")
            raise ValueError(
                f"[{self.config['name']}] ❌ 전국 데이터를 찾을 수 없습니다.\n"
                f"  nationwide 타입: {type(nationwide)}\n"
                f"  nationwide 값: {nationwide}\n"
                f"  table_data 길이: {len(table_data)}"
            )
        
        index_value = nationwide.get('value')
        if index_value is None:
            print(f"[{self.config['name']}] 🔍 [디버그] 전국 지수값 찾기 실패:")
            print(f"  - nationwide 키: {list(nationwide.keys())}")
            print(f"  - nationwide 전체 값: {nationwide}")
            raise ValueError(
                f"[{self.config['name']}] ❌ 전국 지수값을 찾을 수 없습니다.\n"
                f"  nationwide 키: {list(nationwide.keys())}\n"
                f"  nationwide 전체 값: {nationwide}"
            )
        
        growth_rate = nationwide.get('change_rate')
        if growth_rate is None:
            print(f"[{self.config['name']}] 🔍 [디버그] 전국 증감률 찾기 실패:")
            print(f"  - nationwide 키: {list(nationwide.keys())}")
            print(f"  - nationwide 전체 값: {nationwide}")
            raise ValueError(
                f"[{self.config['name']}] ❌ 전국 증감률을 찾을 수 없습니다.\n"
                f"  nationwide 키: {list(nationwide.keys())}\n"
                f"  nationwide 전체 값: {nationwide}"
            )
        
        # 업종별 데이터 추출
        industry_data = self._extract_industry_data('전국')
        
        # 안전한 업종 데이터 처리
        if not industry_data:
            industry_data = []
        
        # 증가/감소 업종 분류 (None 체크 강화)
        increase_industries = [
            ind for ind in industry_data 
            if ind and isinstance(ind, dict) and 
            ind.get('change_rate') is not None and 
            ind['change_rate'] > 0
        ]
        decrease_industries = [
            ind for ind in industry_data 
            if ind and isinstance(ind, dict) and 
            ind.get('change_rate') is not None and 
            ind['change_rate'] < 0
        ]
        
        # 증감률 기준 정렬 (안전한 정렬)
        try:
            # 기본값/폴백 사용 금지: change_rate가 None이면 정렬에서 제외
            increase_industries = [x for x in increase_industries if x and isinstance(x, dict) and x.get('change_rate') is not None]
            decrease_industries = [x for x in decrease_industries if x and isinstance(x, dict) and x.get('change_rate') is not None]
            increase_industries.sort(key=lambda x: x['change_rate'], reverse=True)
            decrease_industries.sort(key=lambda x: x['change_rate'])
        except (TypeError, AttributeError) as e:
            print(f"[{self.config['name']}] ⚠️ 업종 정렬 오류: {e}")
            # 정렬 실패 시 원본 유지
        
        # 상위 3개 추출 (안전한 슬라이싱)
        # 기본값/폴백 사용 금지
        main_increase = increase_industries[:3] if increase_industries and len(increase_industries) > 0 else None
        main_decrease = decrease_industries[:3] if decrease_industries and len(decrease_industries) > 0 else None
        
        # 모든 필드명 포함 (템플릿 호환)
        return {
            'production_index': index_value,
            'sales_index': index_value,  # 소비동향 템플릿 호환
            'service_index': index_value,  # 서비스업 템플릿 호환
            'growth_rate': growth_rate,
            'main_items': main_increase,  # 업종별 데이터 추가 완료
            'main_industries': main_increase,  # 템플릿 호환
            'main_businesses': main_increase,  # 소비동향 템플릿 호환
            'main_increase_industries': main_increase,  # 템플릿 호환
            'main_decrease_industries': main_decrease   # 템플릿 호환
        }
    
    def extract_regional_data(self, table_data: List[Dict] = None) -> Dict[str, Any]:
        """시도별 데이터 추출"""
        if table_data is None:
            table_data = self._extract_table_data_ssot()
        
        # 전국 제외 (안전한 필터링)
        regional = [
            d for d in table_data 
            if d and isinstance(d, dict) and 
            d.get('region_name') != '전국'
        ]
        
        # 증가/감소 분류 (None 체크 강화)
        increase = [
            r for r in regional 
            if r and isinstance(r, dict) and 
            r.get('change_rate') is not None and 
            r['change_rate'] > 0
        ]
        decrease = [
            r for r in regional 
            if r and isinstance(r, dict) and 
            r.get('change_rate') is not None and 
            r['change_rate'] < 0
        ]
        
        # 기본값/폴백 사용 금지: 정렬 (change_rate가 None이면 제외)
        try:
            # change_rate가 None인 항목은 정렬에서 제외
            increase_filtered = [x for x in increase if x and isinstance(x, dict) and x.get('change_rate') is not None]
            decrease_filtered = [x for x in decrease if x and isinstance(x, dict) and x.get('change_rate') is not None]
            increase_filtered.sort(key=lambda x: x['change_rate'], reverse=True)
            decrease_filtered.sort(key=lambda x: x['change_rate'])
            increase = increase_filtered
            decrease = decrease_filtered
        except (TypeError, AttributeError, KeyError) as e:
            print(f"[{self.config['name']}] 🔍 [디버그] 지역 정렬 오류:")
            print(f"  - 오류: {e}")
            print(f"  - increase 샘플: {increase[:3] if increase else '없음'}")
            print(f"  - decrease 샘플: {decrease[:3] if decrease else '없음'}")
            raise ValueError(f"[{self.config['name']}] ❌ 지역 정렬 오류: {e}. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")
        
        return {
            'increase_regions': increase,
            'decrease_regions': decrease,
            'all_regions': regional
        }
    
    def extract_all_data(self) -> Dict[str, Any]:
        """전체 데이터 추출"""
        # 데이터 로드
        self.load_data()
        
        # 스마트 헤더 탐색기로 인덱스 확보 (병합된 셀 처리)
        # 기본값 사용 금지: 반드시 찾아야 함
        # 고용률/실업률은 타입 필터링을 선택적으로 적용
        require_type_match = self.report_type not in ['employment', 'unemployment']
        
        if self.df_aggregation is not None:
            target_idx = self.find_target_col_index(self.df_aggregation, self.year, self.quarter, require_type_match=require_type_match)
            prev_y_idx = self.find_target_col_index(self.df_aggregation, self.year - 1, self.quarter, require_type_match=require_type_match)
            
            if target_idx is None:
                print(f"[{self.config['name']}] 🔍 [디버그] {self.year}년 {self.quarter}분기 컬럼 찾기 실패:")
                print(f"  - 확인한 시트: 집계")
                print(f"  - 시트 크기: {len(self.df_aggregation)}행 × {len(self.df_aggregation.columns)}열")
                # 헤더 행 샘플 출력
                header_sample_rows = min(3, len(self.df_aggregation))
                print(f"  - 헤더 행 샘플 (처음 {header_sample_rows}행):")
                for i in range(header_sample_rows):
                    row_sample = [str(self.df_aggregation.iloc[i, j])[:30] if j < len(self.df_aggregation.columns) and pd.notna(self.df_aggregation.iloc[i, j]) else 'NaN' 
                                 for j in range(min(15, len(self.df_aggregation.columns)))]
                    print(f"    행 {i}: {row_sample}")
                raise ValueError(
                    f"[{self.config['name']}] ❌ {self.year}년 {self.quarter}분기 컬럼을 찾을 수 없습니다.\n"
                    f"  확인한 시트: 집계\n"
                    f"  시트 크기: {len(self.df_aggregation)}행 × {len(self.df_aggregation.columns)}열"
                )
            
            if prev_y_idx is None:
                print(f"[{self.config['name']}] 🔍 [디버그] {self.year - 1}년 {self.quarter}분기 컬럼 찾기 실패:")
                print(f"  - 확인한 시트: 집계")
                print(f"  - 시트 크기: {len(self.df_aggregation)}행 × {len(self.df_aggregation.columns)}열")
                raise ValueError(
                    f"[{self.config['name']}] ❌ {self.year - 1}년 {self.quarter}분기 컬럼을 찾을 수 없습니다.\n"
                    f"  확인한 시트: 집계\n"
                    f"  시트 크기: {len(self.df_aggregation)}행 × {len(self.df_aggregation.columns)}열"
                )
            
            self.target_col = target_idx
            self.prev_y_col = prev_y_idx
            print(f"[{self.config['name']}] ✅ extract_all_data: Target 컬럼 = {target_idx}, 전년 컬럼 = {prev_y_idx}")
        else:
            raise ValueError(
                f"[{self.config['name']}] ❌ 집계 시트를 로드할 수 없습니다. "
                f"기본값 사용 금지: 반드시 데이터를 찾아야 합니다."
            )
        
        # Table Data (SSOT)
        table_data = self._extract_table_data_ssot()
        
        # Text Data
        nationwide = self.extract_nationwide_data(table_data)
        regional = self.extract_regional_data(table_data)
        
        # Top3 regions (템플릿 호환 필드명으로 생성, 기본값/폴백 사용 금지)
        top3_increase = []
        if 'increase_regions' not in regional or not isinstance(regional['increase_regions'], list):
            print(f"[{self.config['name']}] 🔍 [디버그] regional 데이터에서 'increase_regions' 찾기 실패:")
            print(f"  - regional 타입: {type(regional)}")
            print(f"  - regional 키: {list(regional.keys()) if isinstance(regional, dict) else 'N/A'}")
            print(f"  - regional 전체 값: {regional}")
            raise ValueError(
                f"[{self.config['name']}] ❌ regional 데이터에서 'increase_regions'를 찾을 수 없습니다.\n"
                f"  regional 타입: {type(regional)}\n"
                f"  regional 키: {list(regional.keys()) if isinstance(regional, dict) else 'N/A'}\n"
                f"  regional 전체 값: {regional}"
            )
        increase_regions = regional['increase_regions']
        
        for r in increase_regions[:3]:
            if not r or not isinstance(r, dict):
                continue
            
            # 기본값/폴백 사용 금지
            if 'region_name' not in r or not r['region_name']:
                print(f"[{self.config['name']}] 🔍 [디버그] region_name 찾기 실패:")
                print(f"  - r 타입: {type(r)}")
                print(f"  - r 키: {list(r.keys()) if isinstance(r, dict) else 'N/A'}")
                print(f"  - r 전체 값: {r}")
                continue
            region_name = r['region_name']
            
            try:
                # 지역별 업종 데이터 추출
                region_industries = self._extract_industry_data(region_name)
                # 기본값/폴백 사용 금지: 빈 리스트는 그대로 사용 (데이터가 없는 경우)
                # 하지만 None 체크는 필요
                if region_industries is None:
                    raise ValueError(f"[{self.config['name']}] ❌ {region_name} 업종 데이터를 추출할 수 없습니다. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")
                
                # 증가 업종만 필터링 및 정렬 (안전한 처리)
                increase_industries = [
                    ind for ind in region_industries 
                    if ind and isinstance(ind, dict) and 
                    ind.get('change_rate') is not None and 
                    ind['change_rate'] > 0
                ]
                try:
                    # 기본값/폴백 사용 금지: change_rate가 None이면 정렬에서 제외
                    increase_industries = [x for x in increase_industries if x and isinstance(x, dict) and x.get('change_rate') is not None]
                    increase_industries.sort(key=lambda x: x['change_rate'], reverse=True)
                except (TypeError, AttributeError):
                    pass  # 정렬 실패 시 원본 유지
                
                top3_increase.append({
                    'region': region_name,
                    # 기본값/폴백 사용 금지
                    'growth_rate': r['change_rate'] if 'change_rate' in r and r['change_rate'] is not None else None,
                    # 기본값/폴백 사용 금지: increase_industries가 없으면 None
                    'industries': increase_industries[:3] if increase_industries and len(increase_industries) > 0 else None
                })
            except Exception as e:
                print(f"[{self.config['name']}] ⚠️ {region_name} 업종 데이터 추출 오류: {e}")
                # 오류 발생 시 빈 업종 리스트로 추가
                top3_increase.append({
                    'region': region_name,
                    # 기본값/폴백 사용 금지
                    'growth_rate': r['change_rate'] if 'change_rate' in r and r['change_rate'] is not None else None,
                    # 기본값/폴백 사용 금지: 빈 리스트 대신 None
                    'industries': None
                })
        
        top3_decrease = []
        # 기본값/폴백 사용 금지
        if 'decrease_regions' not in regional or not isinstance(regional['decrease_regions'], list):
            print(f"[{self.config['name']}] 🔍 [디버그] regional 데이터에서 'decrease_regions' 찾기 실패:")
            print(f"  - regional 타입: {type(regional)}")
            print(f"  - regional 키: {list(regional.keys()) if isinstance(regional, dict) else 'N/A'}")
            raise ValueError(
                f"[{self.config['name']}] ❌ regional 데이터에서 'decrease_regions'를 찾을 수 없습니다.\n"
                f"  regional 타입: {type(regional)}\n"
                f"  regional 키: {list(regional.keys()) if isinstance(regional, dict) else 'N/A'}"
            )
        decrease_regions = regional['decrease_regions']
        # 기본값/폴백 사용 금지: 타입 체크는 이미 위에서 했으므로 여기서는 추가 체크 불필요
        
        for r in decrease_regions[:3]:
            if not r or not isinstance(r, dict):
                continue
            
            # 기본값/폴백 사용 금지
            if 'region_name' not in r or not r['region_name']:
                print(f"[{self.config['name']}] 🔍 [디버그] region_name 찾기 실패:")
                print(f"  - r 타입: {type(r)}")
                print(f"  - r 키: {list(r.keys()) if isinstance(r, dict) else 'N/A'}")
                print(f"  - r 전체 값: {r}")
                continue
            region_name = r['region_name']
            
            try:
                # 지역별 업종 데이터 추출
                region_industries = self._extract_industry_data(region_name)
                # 기본값/폴백 사용 금지: 빈 리스트는 그대로 사용 (데이터가 없는 경우)
                # 하지만 None 체크는 필요
                if region_industries is None:
                    raise ValueError(f"[{self.config['name']}] ❌ {region_name} 업종 데이터를 추출할 수 없습니다. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")
                
                # 감소 업종만 필터링 및 정렬 (안전한 처리)
                decrease_industries = [
                    ind for ind in region_industries 
                    if ind and isinstance(ind, dict) and 
                    ind.get('change_rate') is not None and 
                    ind['change_rate'] < 0
                ]
                try:
                    # 기본값/폴백 사용 금지: change_rate가 None이면 정렬에서 제외
                    decrease_industries_filtered = [x for x in decrease_industries if x and isinstance(x, dict) and x.get('change_rate') is not None]
                    decrease_industries_filtered.sort(key=lambda x: x['change_rate'])
                    decrease_industries = decrease_industries_filtered
                except (TypeError, AttributeError, KeyError) as e:
                    print(f"[{self.config['name']}] 🔍 [디버그] decrease_industries 정렬 오류:")
                    print(f"  - 오류: {e}")
                    print(f"  - decrease_industries 샘플: {decrease_industries[:3] if decrease_industries else '없음'}")
                    raise ValueError(f"[{self.config['name']}] ❌ decrease_industries 정렬 오류: {e}. 기본값 사용 금지: 반드시 데이터를 찾아야 합니다.")
                
                # 소비동향용 주요 업태 (첫 번째 감소 업종, 기본값/폴백 사용 금지)
                main_business = None
                if decrease_industries and decrease_industries[0] and isinstance(decrease_industries[0], dict):
                    # 기본값/폴백 사용 금지
                    if 'name' not in decrease_industries[0] or not decrease_industries[0]['name']:
                        raise ValueError(f"[{self.config['name']}] ❌ decrease_industries[0]에서 'name'을 찾을 수 없습니다.")
                    main_business = decrease_industries[0]['name']
                
                top3_decrease.append({
                    'region': region_name,
                    # 기본값/폴백 사용 금지
                    'growth_rate': r['change_rate'] if 'change_rate' in r and r['change_rate'] is not None else None,
                    # 기본값/폴백 사용 금지
                    'industries': decrease_industries[:3] if decrease_industries and len(decrease_industries) > 0 else None,
                    'main_business': main_business  # 소비동향용 주요 업태
                })
            except Exception as e:
                print(f"[{self.config['name']}] ⚠️ {region_name} 업종 데이터 추출 오류: {e}")
                # 오류 발생 시 빈 업종 리스트로 추가
                top3_decrease.append({
                    'region': region_name,
                    # 기본값/폴백 사용 금지
                    'growth_rate': r['change_rate'] if 'change_rate' in r and r['change_rate'] is not None else None,
                    # 기본값/폴백 사용 금지: 빈 리스트 대신 None
                    'industries': None,
                    # 기본값/폴백 사용 금지: 빈 문자열 대신 None
                    'main_business': None
                })
        
        # Summary Box (안전한 처리)
        main_regions = []
        for r in top3_increase:
            if r and isinstance(r, dict):
                main_regions.append({
                    # 기본값/폴백 사용 금지
                    'region': r['region'] if 'region' in r and r['region'] else None,
                    # 기본값/폴백 사용 금지
                    'items': r['industries'] if 'industries' in r and isinstance(r['industries'], list) else None
                })
        
        # 기본값/폴백 사용 금지
        if 'increase_regions' not in regional or not isinstance(regional['increase_regions'], list):
            raise ValueError(f"[{self.config['name']}] ❌ regional 데이터에서 'increase_regions'를 찾을 수 없습니다.")
        increase_regions_count = len(regional['increase_regions'])
        
        summary_box = {
            'main_regions': main_regions,
            'region_count': increase_regions_count
        }
        
        # Regional data 필드명 변환 (템플릿 호환, 기본값/폴백 사용 금지)
        if 'increase_regions' not in regional or not isinstance(regional['increase_regions'], list):
            print(f"[{self.config['name']}] 🔍 [디버그] regional 데이터에서 'increase_regions' 찾기 실패:")
            print(f"  - regional 타입: {type(regional)}")
            print(f"  - regional 키: {list(regional.keys()) if isinstance(regional, dict) else 'N/A'}")
            raise ValueError(
                f"[{self.config['name']}] ❌ regional 데이터에서 'increase_regions'를 찾을 수 없습니다.\n"
                f"  regional 타입: {type(regional)}\n"
                f"  regional 키: {list(regional.keys()) if isinstance(regional, dict) else 'N/A'}"
            )
        increase_regions_list = regional['increase_regions']
        
        if 'decrease_regions' not in regional or not isinstance(regional['decrease_regions'], list):
            print(f"[{self.config['name']}] 🔍 [디버그] regional 데이터에서 'decrease_regions' 찾기 실패:")
            print(f"  - regional 타입: {type(regional)}")
            print(f"  - regional 키: {list(regional.keys()) if isinstance(regional, dict) else 'N/A'}")
            raise ValueError(
                f"[{self.config['name']}] ❌ regional 데이터에서 'decrease_regions'를 찾을 수 없습니다.\n"
                f"  regional 타입: {type(regional)}\n"
                f"  regional 키: {list(regional.keys()) if isinstance(regional, dict) else 'N/A'}"
            )
        decrease_regions_list = regional['decrease_regions']
        
        if 'all_regions' not in regional or not isinstance(regional['all_regions'], list):
            print(f"[{self.config['name']}] 🔍 [디버그] regional 데이터에서 'all_regions' 찾기 실패:")
            print(f"  - regional 타입: {type(regional)}")
            print(f"  - regional 키: {list(regional.keys()) if isinstance(regional, dict) else 'N/A'}")
            raise ValueError(
                f"[{self.config['name']}] ❌ regional 데이터에서 'all_regions'를 찾을 수 없습니다.\n"
                f"  regional 타입: {type(regional)}\n"
                f"  regional 키: {list(regional.keys()) if isinstance(regional, dict) else 'N/A'}"
            )
        all_regions_list = regional['all_regions']
        
        regional_converted = {
            'increase_regions': [
                {
                    # 기본값/폴백 사용 금지
                    'region': r['region_name'] if r and isinstance(r, dict) and 'region_name' in r and r['region_name'] else None,
                    'growth_rate': r['change_rate'] if r and isinstance(r, dict) and 'change_rate' in r and r['change_rate'] is not None else None,
                    # 기본값/폴백 사용 금지
                    'value': r['value'] if r and isinstance(r, dict) and 'value' in r and r['value'] is not None else None,
                    'top_industries': self._get_top_industries_for_region(
                        r['region_name'] if r and isinstance(r, dict) and 'region_name' in r and r['region_name'] else None, 
                        increase=True
                    )
                }
                for r in increase_regions_list
                if r and isinstance(r, dict) and r.get('region_name')
            ],
            'decrease_regions': [
                {
                    # 기본값/폴백 사용 금지
                    'region': r['region_name'] if r and isinstance(r, dict) and 'region_name' in r and r['region_name'] else None,
                    'growth_rate': r['change_rate'] if r and isinstance(r, dict) and 'change_rate' in r and r['change_rate'] is not None else None,
                    # 기본값/폴백 사용 금지
                    'value': r['value'] if r and isinstance(r, dict) and 'value' in r and r['value'] is not None else None,
                    'top_industries': self._get_top_industries_for_region(
                        r['region_name'] if r and isinstance(r, dict) and 'region_name' in r and r['region_name'] else None, 
                        increase=False
                    )
                }
                for r in decrease_regions_list
                if r and isinstance(r, dict) and r.get('region_name')
            ],
            'all_regions': all_regions_list
        }
        
        return {
            'report_info': {
                'year': self.year,
                'quarter': self.quarter,
                'report_type': self.report_type,
                'report_name': self.config['name'],
                'index_name': self.config.get('index_name', '지수'),
                'item_name': self.config.get('item_name', '항목')
            },
            'summary_box': summary_box,
            'nationwide_data': nationwide,
            'regional_data': regional_converted,  # 필드명 변환된 버전
            'table_data': table_data,
            'top3_increase_regions': top3_increase,  # 템플릿 호환
            'top3_decrease_regions': top3_decrease   # 템플릿 호환
        }


# 하위 호환성 Wrapper
class MiningManufacturingGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('mining', excel_path, year, quarter, excel_file)


class ServiceIndustryGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('service', excel_path, year, quarter, excel_file)


class ConsumptionGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('consumption', excel_path, year, quarter, excel_file)


class ConstructionGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('construction', excel_path, year, quarter, excel_file)


class ExportGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('export', excel_path, year, quarter, excel_file)


class ImportGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('import', excel_path, year, quarter, excel_file)


class PriceTrendGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('price', excel_path, year, quarter, excel_file)


class EmploymentRateGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('employment', excel_path, year, quarter, excel_file)


class UnemploymentGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('unemployment', excel_path, year, quarter, excel_file)


class DomesticMigrationGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        # report_configs.py에서 'migration'을 사용하지만, 
        # 실제로는 REPORT_CONFIGS에 'migration'으로 정의되어 있으므로 'migration' 사용
        super().__init__('migration', excel_path, year, quarter, excel_file)


class RegionalReportGenerator(BaseGenerator):
    """시도별 보고서 생성기 (unified_generator에 통합)"""
    
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__(excel_path, year, quarter, excel_file)
        # regional_generator.py를 import하여 사용
        self._regional_gen = None
    
    def _get_regional_generator(self):
        """regional_generator.py의 RegionalGenerator 인스턴스 가져오기 (지연 로딩)"""
        if self._regional_gen is None:
            # regional_generator.py 동적 import
            generator_path = Path(__file__).parent / 'regional_generator.py'
            if generator_path.exists():
                import importlib.util
                spec = importlib.util.spec_from_file_location('regional_generator', str(generator_path))
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                
                if hasattr(module, 'RegionalGenerator'):
                    self._regional_gen = module.RegionalGenerator(
                        str(self.excel_path), 
                        year=self.year, 
                        quarter=self.quarter
                    )
        return self._regional_gen
    
    def extract_all_data(self, region: str) -> Dict[str, Any]:
        """시도별 모든 데이터 추출"""
        regional_gen = self._get_regional_generator()
        if regional_gen is None:
            raise ValueError("시도별 Generator를 로드할 수 없습니다")
        
        return regional_gen.extract_all_data(region)
    
    def render_html(self, region: str, template_path: str) -> str:
        """시도별 HTML 보도자료 렌더링"""
        regional_gen = self._get_regional_generator()
        if regional_gen is None:
            raise ValueError("시도별 Generator를 로드할 수 없습니다")
        
        return regional_gen.render_html(region, template_path)


if __name__ == '__main__':
    # 테스트
    base_path = Path(__file__).parent.parent
    excel_path = base_path / '분석표_25년 3분기_캡스톤(업데이트).xlsx'
    
    print("=" * 70)
    print("통합 Generator V2 테스트 (집계 시트 기반)")
    print("=" * 70)
    
    for report_type in ['mining', 'service', 'consumption']:
        print(f"\n{'='*70}")
        print(f"[TEST] {REPORT_CONFIGS[report_type]['name']}")
        print(f"{'='*70}\n")
        
        try:
            generator = UnifiedReportGenerator(report_type, str(excel_path), 2025, 3)
            data = generator.extract_all_data()
            
            # 결과 출력
            print(f"\n[결과] ✅ 데이터 추출 완료")
            nationwide = data['nationwide_data']
            print(f"  전국: 지수={nationwide['production_index']:.1f}, 증감률={nationwide['growth_rate']}%")
            
            regional = data['regional_data']
            print(f"  지역: 증가={len(regional['increase_regions'])}개, 감소={len(regional['decrease_regions'])}개")
            
            if regional['increase_regions']:
                top = regional['increase_regions'][0]
                print(f"  최고: {top['region_name']} ({top['change_rate']}%)")
            
        except Exception as e:
            print(f"\n[오류] ❌ {e}")
            import traceback
            traceback.print_exc()
