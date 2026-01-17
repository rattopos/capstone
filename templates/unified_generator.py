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
        self.name_mapping = self.config.get('name_mapping', {})
        
        # 집계 시트 구조 (설정에서 로드 - 기본값은 나중에 동적으로 찾을 때 fallback으로만 사용)
        agg_struct = self.config.get('aggregation_structure', {})
        # 기본값은 설정에서 가져오지만, 실제로는 동적으로 찾음
        self.region_name_col = None  # 동적으로 찾음
        self.industry_code_col = None  # 동적으로 찾음
        self.total_code = agg_struct.get('total_code', 'BCD')
        
        # metadata_columns 설정 (동적 컬럼 찾기에 사용)
        self.metadata_cols = self.config.get('metadata_columns', {})
        
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
        
        # 1. 분석 시트 찾기
        analysis_sheets = self.config['sheets'].get('analysis', [])
        fallback_sheets = self.config['sheets'].get('fallback', [])
        
        analysis_sheet = None
        for name in analysis_sheets:
            if name in sheet_names:
                analysis_sheet = name
                break
        
        # 분석 시트가 없으면 fallback 시트에서 찾기
        if not analysis_sheet:
            for name in fallback_sheets:
                if name in sheet_names:
                    analysis_sheet = name
                    print(f"[{self.config['name']}] [시트 대체] 분석 시트 → '{name}' (기초자료)")
                    break
        
        # 2. 집계 시트 찾기
        agg_sheets = self.config['sheets'].get('aggregation', [])
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
        
        # Fallback: 설정에서 기본값 사용 (하드코딩이지만 최후의 수단)
        if self.region_name_col is None:
            agg_struct = self.config.get('aggregation_structure', {})
            self.region_name_col = agg_struct.get('region_name_col')
            if self.region_name_col is not None:
                print(f"[{self.config['name']}] ⚠️ 지역명 컬럼을 동적으로 찾지 못해 설정값 사용: {self.region_name_col}")
        
        if self.industry_code_col is None:
            agg_struct = self.config.get('aggregation_structure', {})
            self.industry_code_col = agg_struct.get('industry_code_col')
            if self.industry_code_col is not None:
                print(f"[{self.config['name']}] ⚠️ 산업코드 컬럼을 동적으로 찾지 못해 설정값 사용: {self.industry_code_col}")
        
        if self.industry_name_col is None and self.industry_code_col is not None:
            # 산업명은 보통 산업코드 다음 컬럼
            self.industry_name_col = self.industry_code_col + 1
            print(f"[{self.config['name']}] ⚠️ 산업명 컬럼을 동적으로 찾지 못해 산업코드+1 사용: {self.industry_name_col}")
        
        if self.data_start_row is None:
            # 기본값 사용 (하드코딩이지만 최후의 수단)
            self.data_start_row = 3
            print(f"[{self.config['name']}] ⚠️ 데이터 시작 행을 동적으로 찾지 못해 기본값 사용: {self.data_start_row}")
    
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
        # target_col 찾기
        if self.target_col is None:
            self.target_col = self.find_target_col_index(df, self.year, self.quarter)
            if self.target_col is not None:
                print(f"[{self.config['name']}] ✅ Target 컬럼 ({sheet_type} 시트): {self.target_col} ({self.year} {self.quarter}/4)")
        
        # prev_y_col 찾기
        if self.prev_y_col is None:
            self.prev_y_col = self.find_target_col_index(df, self.year - 1, self.quarter)
            if self.prev_y_col is not None:
                print(f"[{self.config['name']}] ✅ 전년 컬럼 ({sheet_type} 시트): {self.prev_y_col} ({self.year - 1} {self.quarter}/4)")
        
        # 기본값 사용 금지: 반드시 찾아야 함
        if self.target_col is None:
            raise ValueError(
                f"[{self.config['name']}] ❌ Target 컬럼을 찾을 수 없습니다. "
                f"기본값 사용 금지: 반드시 데이터를 찾아야 합니다."
            )
        
        if self.prev_y_col is None:
            raise ValueError(
                f"[{self.config['name']}] ❌ 전년 컬럼을 찾을 수 없습니다. "
                f"기본값 사용 금지: 반드시 데이터를 찾아야 합니다."
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
                print(f"[{self.config['name']}] ⚠️ {region} 총지수 행 찾기 오류: {e}")
                region_total = region_filter.head(1)  # Fallback
            
            if region_total.empty:
                # Fallback: 첫 번째 행
                print(f"[{self.config['name']}] ⚠️ {region}: 코드 '{self.total_code}' 찾기 실패, 첫 번째 행 사용")
                region_total = region_filter.head(1)
            
            if region_total.empty:
                continue
            
            row = region_total.iloc[0]
            
            # 기본값 사용 금지: 반드시 유효한 인덱스여야 함
            if self.target_col is None:
                raise ValueError(
                    f"[{self.config['name']}] ❌ Target 컬럼이 None입니다. "
                    f"기본값 사용 금지: 반드시 데이터를 찾아야 합니다."
                )
            
            if self.prev_y_col is None:
                raise ValueError(
                    f"[{self.config['name']}] ❌ 전년 컬럼이 None입니다. "
                    f"기본값 사용 금지: 반드시 데이터를 찾아야 합니다."
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
        name_mapping = self.config.get('name_mapping', {})
        
        # 산업명 컬럼 찾기 (동적으로 찾은 값 사용)
        if self.industry_name_col is None:
            # 산업명 컬럼을 찾지 못한 경우, 산업코드 다음 컬럼 사용 (fallback)
            if self.industry_code_col is not None:
                industry_name_col = self.industry_code_col + 1
            else:
                print(f"[{self.config['name']}] ⚠️ 산업명 컬럼과 산업코드 컬럼을 모두 찾을 수 없습니다.")
                return []
        else:
            industry_name_col = self.industry_name_col
        
        if industry_name_col < 0:
            industry_name_col = 0
        
        for idx, row in region_filter.iterrows():
            # 산업코드 확인 (총지수 제외) - 동적으로 찾은 컬럼 사용
            if self.industry_code_col is None:
                continue
            
            if self.industry_code_col >= len(row):
                continue
                
            industry_code = str(row.iloc[self.industry_code_col]).strip() if pd.notna(row.iloc[self.industry_code_col]) else ''
            
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
                filtered.sort(key=lambda x: x.get('change_rate', 0) if x and isinstance(x, dict) else 0, reverse=True)
            except (TypeError, AttributeError):
                pass  # 정렬 실패 시 원본 유지
        else:
            filtered = [
                ind for ind in industries 
                if ind and isinstance(ind, dict) and 
                ind.get('change_rate') is not None and 
                ind['change_rate'] < 0
            ]
            try:
                filtered.sort(key=lambda x: x.get('change_rate', 0) if x and isinstance(x, dict) else 0)
            except (TypeError, AttributeError):
                pass  # 정렬 실패 시 원본 유지
        
        # 안전한 슬라이싱
        return filtered[:top_n] if filtered else []
    
    def extract_nationwide_data(self, table_data: List[Dict] = None) -> Dict[str, Any]:
        """전국 데이터 추출 - 템플릿 호환 필드명"""
        if table_data is None:
            table_data = self._extract_table_data_ssot()
        
        nationwide = next((d for d in table_data if d['region_name'] == '전국'), None)
        
        if not nationwide:
            return {
                'production_index': 100.0,
                'sales_index': 100.0,  # 소비동향 템플릿 호환
                'service_index': 100.0,  # 서비스업 템플릿 호환
                'growth_rate': 0.0,
                'main_items': [],
                'main_industries': [],  # 템플릿 호환
                'main_businesses': [],  # 소비동향 템플릿 호환
                'main_increase_industries': [],  # 템플릿 호환
                'main_decrease_industries': []   # 템플릿 호환
            }
        
        # 안전한 값 추출
        index_value = nationwide.get('value', 100.0) if nationwide and isinstance(nationwide, dict) else 100.0
        growth_rate = nationwide.get('change_rate', 0.0) if nationwide and isinstance(nationwide, dict) and nationwide.get('change_rate') is not None else 0.0
        
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
            increase_industries.sort(key=lambda x: x.get('change_rate', 0) if x and isinstance(x, dict) else 0, reverse=True)
            decrease_industries.sort(key=lambda x: x.get('change_rate', 0) if x and isinstance(x, dict) else 0)
        except (TypeError, AttributeError) as e:
            print(f"[{self.config['name']}] ⚠️ 업종 정렬 오류: {e}")
            # 정렬 실패 시 원본 유지
        
        # 상위 3개 추출 (안전한 슬라이싱)
        main_increase = increase_industries[:3] if increase_industries else []
        main_decrease = decrease_industries[:3] if decrease_industries else []
        
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
        
        # 안전한 정렬
        try:
            increase.sort(key=lambda x: x.get('change_rate', 0) if x and isinstance(x, dict) else 0, reverse=True)
            decrease.sort(key=lambda x: x.get('change_rate', 0) if x and isinstance(x, dict) else 0)
        except (TypeError, AttributeError) as e:
            print(f"[{self.config['name']}] ⚠️ 지역 정렬 오류: {e}")
            # 정렬 실패 시 원본 유지
        
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
        if self.df_aggregation is not None:
            target_idx = self.find_target_col_index(self.df_aggregation, self.year, self.quarter)
            prev_y_idx = self.find_target_col_index(self.df_aggregation, self.year - 1, self.quarter)
            
            if target_idx is None:
                raise ValueError(
                    f"[{self.config['name']}] ❌ {self.year}년 {self.quarter}분기 컬럼을 찾을 수 없습니다. "
                    f"기본값 사용 금지: 반드시 데이터를 찾아야 합니다."
                )
            
            if prev_y_idx is None:
                raise ValueError(
                    f"[{self.config['name']}] ❌ {self.year - 1}년 {self.quarter}분기 컬럼을 찾을 수 없습니다. "
                    f"기본값 사용 금지: 반드시 데이터를 찾아야 합니다."
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
        
        # Top3 regions (템플릿 호환 필드명으로 생성) - 안전한 처리
        top3_increase = []
        increase_regions = regional.get('increase_regions', [])
        if not isinstance(increase_regions, list):
            increase_regions = []
        
        for r in increase_regions[:3]:
            if not r or not isinstance(r, dict):
                continue
            
            region_name = r.get('region_name', '')
            if not region_name:
                continue
            
            try:
                # 지역별 업종 데이터 추출
                region_industries = self._extract_industry_data(region_name)
                if not region_industries:
                    region_industries = []
                
                # 증가 업종만 필터링 및 정렬 (안전한 처리)
                increase_industries = [
                    ind for ind in region_industries 
                    if ind and isinstance(ind, dict) and 
                    ind.get('change_rate') is not None and 
                    ind['change_rate'] > 0
                ]
                try:
                    increase_industries.sort(key=lambda x: x.get('change_rate', 0) if x and isinstance(x, dict) else 0, reverse=True)
                except (TypeError, AttributeError):
                    pass  # 정렬 실패 시 원본 유지
                
                top3_increase.append({
                    'region': region_name,
                    'growth_rate': r.get('change_rate', 0.0) if r.get('change_rate') is not None else 0.0,
                    'industries': increase_industries[:3] if increase_industries else []  # 상위 3개 업종
                })
            except Exception as e:
                print(f"[{self.config['name']}] ⚠️ {region_name} 업종 데이터 추출 오류: {e}")
                # 오류 발생 시 빈 업종 리스트로 추가
                top3_increase.append({
                    'region': region_name,
                    'growth_rate': r.get('change_rate', 0.0) if r.get('change_rate') is not None else 0.0,
                    'industries': []
                })
        
        top3_decrease = []
        decrease_regions = regional.get('decrease_regions', [])
        if not isinstance(decrease_regions, list):
            decrease_regions = []
        
        for r in decrease_regions[:3]:
            if not r or not isinstance(r, dict):
                continue
            
            region_name = r.get('region_name', '')
            if not region_name:
                continue
            
            try:
                # 지역별 업종 데이터 추출
                region_industries = self._extract_industry_data(region_name)
                if not region_industries:
                    region_industries = []
                
                # 감소 업종만 필터링 및 정렬 (안전한 처리)
                decrease_industries = [
                    ind for ind in region_industries 
                    if ind and isinstance(ind, dict) and 
                    ind.get('change_rate') is not None and 
                    ind['change_rate'] < 0
                ]
                try:
                    decrease_industries.sort(key=lambda x: x.get('change_rate', 0) if x and isinstance(x, dict) else 0)
                except (TypeError, AttributeError):
                    pass  # 정렬 실패 시 원본 유지
                
                # 소비동향용 주요 업태 (첫 번째 감소 업종)
                main_business = ''
                if decrease_industries and decrease_industries[0] and isinstance(decrease_industries[0], dict):
                    main_business = decrease_industries[0].get('name', '')
                
                top3_decrease.append({
                    'region': region_name,
                    'growth_rate': r.get('change_rate', 0.0) if r.get('change_rate') is not None else 0.0,
                    'industries': decrease_industries[:3] if decrease_industries else [],  # 상위 3개 업종
                    'main_business': main_business  # 소비동향용 주요 업태
                })
            except Exception as e:
                print(f"[{self.config['name']}] ⚠️ {region_name} 업종 데이터 추출 오류: {e}")
                # 오류 발생 시 빈 업종 리스트로 추가
                top3_decrease.append({
                    'region': region_name,
                    'growth_rate': r.get('change_rate', 0.0) if r.get('change_rate') is not None else 0.0,
                    'industries': [],
                    'main_business': ''
                })
        
        # Summary Box (안전한 처리)
        main_regions = []
        for r in top3_increase:
            if r and isinstance(r, dict):
                main_regions.append({
                    'region': r.get('region', ''),
                    'items': r.get('industries', []) if isinstance(r.get('industries'), list) else []
                })
        
        increase_regions_count = len(regional.get('increase_regions', [])) if isinstance(regional.get('increase_regions'), list) else 0
        
        summary_box = {
            'main_regions': main_regions,
            'region_count': increase_regions_count
        }
        
        # Regional data 필드명 변환 (템플릿 호환) - 안전한 처리
        increase_regions_list = regional.get('increase_regions', [])
        if not isinstance(increase_regions_list, list):
            increase_regions_list = []
        
        decrease_regions_list = regional.get('decrease_regions', [])
        if not isinstance(decrease_regions_list, list):
            decrease_regions_list = []
        
        all_regions_list = regional.get('all_regions', [])
        if not isinstance(all_regions_list, list):
            all_regions_list = []
        
        regional_converted = {
            'increase_regions': [
                {
                    'region': r.get('region_name', '') if r and isinstance(r, dict) else '',
                    'growth_rate': r.get('change_rate', 0.0) if r and isinstance(r, dict) and r.get('change_rate') is not None else 0.0,
                    'value': r.get('value', 0.0) if r and isinstance(r, dict) else 0.0,
                    'top_industries': self._get_top_industries_for_region(
                        r.get('region_name', '') if r and isinstance(r, dict) else '', 
                        increase=True
                    )
                }
                for r in increase_regions_list
                if r and isinstance(r, dict) and r.get('region_name')
            ],
            'decrease_regions': [
                {
                    'region': r.get('region_name', '') if r and isinstance(r, dict) else '',
                    'growth_rate': r.get('change_rate', 0.0) if r and isinstance(r, dict) and r.get('change_rate') is not None else 0.0,
                    'value': r.get('value', 0.0) if r and isinstance(r, dict) else 0.0,
                    'top_industries': self._get_top_industries_for_region(
                        r.get('region_name', '') if r and isinstance(r, dict) else '', 
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
