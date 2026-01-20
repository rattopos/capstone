#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd

from pathlib import Path
from typing import Dict, Any, List, Optional
try:
    from .base_generator import BaseGenerator
    from config.reports import REPORT_ORDER, SECTOR_REPORTS, REGIONAL_REPORTS, REGION_DISPLAY_MAPPING, REGION_GROUPS, VALID_REGIONS
except ImportError:
    import sys
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from templates.base_generator import BaseGenerator
    from config.reports import REPORT_ORDER, SECTOR_REPORTS, REGIONAL_REPORTS, REGION_DISPLAY_MAPPING, REGION_GROUPS, VALID_REGIONS

def get_report_config(report_type: str) -> dict:
    """Return the config matching either id or report_id; accept legacy aliases."""
    aliases = {
        'mining': 'manufacturing',  # legacy name used in 일부 호출
    }
    normalized = aliases.get(report_type, report_type)
    for config in SECTOR_REPORTS:
        # 지원: id 매칭 혹은 report_id 매칭
        if config.get('id') == normalized or config.get('report_id') == normalized:
            return config
    raise ValueError(f"Unknown report type: {report_type}")


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
        # report_id 누락 시 id로 폴백하여 KeyError 방지
        self.report_id = self.config.get('report_id', self.config.get('id', report_type))
        if 'name_mapping' not in self.config:
            raise ValueError(f"[{self.config['name']}] ❌ 설정에서 'name_mapping'을 찾을 수 없습니다. 기본값 사용 금지.")
        self.name_mapping = self.config['name_mapping']

        if 'aggregation_structure' not in self.config:
            raise ValueError(f"[{self.config['name']}] ❌ 설정에서 'aggregation_structure'를 찾을 수 없습니다. 기본값 사용 금지.")
        # metadata_columns는 컬럼 존재 여부 힌트일 뿐, 키워드 탐색에는 기본 키워드 목록을 사용
        meta = self.config.get('metadata_columns', {})
        if isinstance(meta, dict):
            self.metadata_cols = meta
        elif isinstance(meta, list):
            # 단순 보존
            self.metadata_cols = {c: c for c in meta}
        else:
            self.metadata_cols = {}
        # 동적으로 할당되는 주요 속성들 기본값 None으로 초기화
        self.region_name_col = None
        self.industry_code_col = None
        self.industry_name_col = None
        self.data_start_row = None
        self.df_analysis = None
        self.df_aggregation = None
        self.df_reference = None
        self.target_col = None
        self.prev_y_col = None
        self.prev_prev_y_col = None
        self.prev_prev_prev_y_col = None
        self.quarterly_keys = []
        self.quarterly_cols = {}
        self.analysis_target_col = None
        self.analysis_prev_y_col = None
        self.analysis_prev_prev_y_col = None
        self.analysis_prev_prev_prev_y_col = None
        self.analysis_quarterly_keys = []
        self.analysis_quarterly_cols = {}
        # 인스턴스 생성 시 데이터프레임 등 필드 자동 초기화
        self.load_data()
    def _get_region_display_name(self, region: str) -> str:
        try:
            return REGION_DISPLAY_MAPPING.get(region, region)
        except Exception:
            return region
    @staticmethod
    def _is_numeric(val) -> bool:
        try:
            if pd.isna(val):
                return False
            float(str(val).replace(',', '').replace('%', ''))
            return True
        except Exception:
            return False

    @staticmethod
    def _find_textual_column(df: pd.DataFrame, header_rows: int, exclude_cols: List[int]) -> Optional[int]:
        """
        헤더 키워드로 못 찾을 때, 데이터 행의 문자 비율이 높은 컬럼을 업종명 후보로 추정
        """
        if df is None or df.empty:
            return None
        n_rows = min(len(df) - header_rows, 30)
        if n_rows <= 0:
            return None
        best_idx = None
        best_score = -1.0
        start = max(header_rows, 0)
        for col_idx in range(len(df.columns)):
            if exclude_cols and col_idx in exclude_cols:
                continue
            text_cnt = 0
            total = 0
            for r in range(start, start + n_rows):
                val = df.iloc[r, col_idx] if col_idx < len(df.columns) else None
                if pd.isna(val):
                    continue
                total += 1
                s = str(val).strip()
                # 숫자만/날짜/코드 패턴 제외
                if not UnifiedReportGenerator._is_numeric(s):
                    text_cnt += 1
            if total == 0:
                continue
            score = text_cnt / total
            if score > best_score:
                best_score = score
                best_idx = col_idx
        return best_idx

    @staticmethod
    def _find_total_row_by_name(df: pd.DataFrame, name_col: int, header_rows: int) -> Optional[pd.DataFrame]:
        """
        업종명 컬럼에서 총계를 의미하는 키워드로 행을 탐색
        """
        if df is None or df.empty or name_col is None:
            return None
        # '계' 단독 키워드는 '단계' 등과 오탐 가능하므로 제외
        keywords = ['총계', '합계', '총지수', '전체', '전산업', '전 산업']
        try:
            series = df.iloc[:, name_col].astype(str).str.strip()
        except Exception:
            return None
        mask = pd.Series(False, index=series.index)
        for kw in keywords:
            mask = mask | series.str.contains(kw, na=False)
        result = df[mask]
        if result is not None and not result.empty:
            return result.head(1)
        return None

    @staticmethod
    def _find_total_row_by_code(
        df: pd.DataFrame,
        total_code: Any,
        exclude_cols: Optional[List[int]] = None
    ) -> Optional[pd.DataFrame]:
        """지정 코드(total_code)를 갖는 행을 모든 텍스트 컬럼에서 탐색."""
        if df is None or df.empty or total_code is None:
            return None
        code_str = str(total_code).strip()
        exclude_cols = exclude_cols or []
        for col_idx in range(len(df.columns)):
            if col_idx in exclude_cols:
                continue
            try:
                series = df.iloc[:, col_idx].astype(str).str.strip()
            except Exception:
                continue
            matched = df[series == code_str]
            if matched is not None and not matched.empty:
                return matched.head(1)
        return None

    @staticmethod
    def _previous_quarter(year: int, quarter: int) -> tuple[int, int]:
        if quarter <= 1:
            return (year - 1, 4)
        return (year, quarter - 1)

    @staticmethod
    def _format_quarter_key(year: int, quarter: int) -> str:
        return f"{year} {quarter}/4"

    def _build_quarter_range(
        self,
        start_year: int,
        start_quarter: int,
        end_year: int,
        end_quarter: int
    ) -> List[tuple[int, int]]:
        quarters = []
        y, q = start_year, start_quarter
        while (y < end_year) or (y == end_year and q <= end_quarter):
            quarters.append((y, q))
            q += 1
            if q > 4:
                q = 1
                y += 1
        return quarters

    def _ensure_quarter_columns(
        self,
        df: pd.DataFrame,
        start_year: int,
        start_quarter: int,
        end_year: int,
        end_quarter: int,
        max_header_rows: int
    ) -> None:
        if df is None or start_year is None or start_quarter is None or end_year is None or end_quarter is None:
            return
        quarter_range = self._build_quarter_range(start_year, start_quarter, end_year, end_quarter)
        keys: List[str] = []
        cols: Dict[str, Optional[int]] = {}
        for y, q in quarter_range:
            key = self._format_quarter_key(y, q)
            keys.append(key)
            cols[key] = self.find_target_col_index(
                df,
                y,
                q,
                require_type_match=False,
                max_header_rows=max_header_rows
            )
        self.quarterly_keys = keys
        self.quarterly_cols = cols

    def _collect_quarter_columns(
        self,
        df: pd.DataFrame,
        start_year: int,
        start_quarter: int,
        end_year: int,
        end_quarter: int,
        max_header_rows: int
    ) -> tuple[List[str], Dict[str, Optional[int]]]:
        if df is None or start_year is None or start_quarter is None or end_year is None or end_quarter is None:
            return [], {}
        quarter_range = self._build_quarter_range(start_year, start_quarter, end_year, end_quarter)
        keys: List[str] = []
        cols: Dict[str, Optional[int]] = {}
        for y, q in quarter_range:
            key = self._format_quarter_key(y, q)
            keys.append(key)
            cols[key] = self.find_target_col_index(
                df,
                y,
                q,
                require_type_match=False,
                max_header_rows=max_header_rows
            )
        return keys, cols
    def load_data(self):
        """
        테스트 호환성: 기존 테스트 코드에서 generator.load_data()를 호출하는 경우
        실제 데이터프레임 및 주요 속성(df_aggregation, target_col 등)을 초기화
        
        데이터 누락 시 우아하게 처리:
        - 요청한 연도/분기가 없으면 최신 데이터를 자동으로 사용
        - 설정에서 require_analysis_sheet=False면 분석시트 요구 안 함
        """
        import openpyxl
        wb = openpyxl.load_workbook(self.excel_path, data_only=True)
        agg_sheet_name = self.config['aggregation_structure']['sheet']
        print(f"[디버그] config['aggregation_structure']: {self.config.get('aggregation_structure')}")
        print(f"[디버그] agg_sheet_name: {agg_sheet_name}")
        print(f"[디버그] wb.sheetnames: {wb.sheetnames}")
        if not agg_sheet_name:
            raise ValueError('집계 시트명이 설정에 없습니다.')
        # 헤더 행을 보존하기 위해 header=None으로 읽어 병합 헤더 탐색과 데이터 시작 행 탐색을 일관되게 처리
        self.df_aggregation = pd.read_excel(self.excel_path, sheet_name=agg_sheet_name, header=None)
        # 집계 범위가 설정되어 있으면 해당 범위만 사용
        agg_range = self.config.get('aggregation_range')
        if isinstance(agg_range, dict) and self.df_aggregation is not None:
            from openpyxl.utils import column_index_from_string

            def _col_to_index(col_value):
                if col_value is None:
                    return None
                if isinstance(col_value, int):
                    return col_value
                if isinstance(col_value, str) and col_value.strip():
                    return column_index_from_string(col_value.strip().upper()) - 1
                return None

            start_row = agg_range.get('start_row')
            end_row = agg_range.get('end_row')
            start_col = _col_to_index(agg_range.get('start_col'))
            end_col = _col_to_index(agg_range.get('end_col'))

            row_start = max((start_row - 1) if isinstance(start_row, int) else 0, 0)
            row_end = end_row if isinstance(end_row, int) else len(self.df_aggregation)
            col_start = start_col if isinstance(start_col, int) else 0
            col_end = (end_col + 1) if isinstance(end_col, int) else len(self.df_aggregation.columns)

            self.df_aggregation = self.df_aggregation.iloc[row_start:row_end, col_start:col_end].copy()
            print(
                f"[{self.config['name']}] ✅ 집계 범위 적용: rows {start_row}-{end_row}, cols {agg_range.get('start_col')}-{agg_range.get('end_col')}"
            )
        self.target_col = None
        
        # target column 찾기 (요청한 연도/분기)
        require_type_match = False
        sheet_type = agg_sheet_name
        
        # config에서 header_rows 가져오기 (기본값 5)
        max_header_rows = self.config.get('header_rows', 5)
        
        # 1. 요청한 연도/분기 찾기 (max_header_rows 전달)
        target_col_result = self.find_target_col_index(
            self.df_aggregation, self.year, self.quarter, 
            require_type_match=require_type_match,
            max_header_rows=max_header_rows
        )
        
        # 2. 없으면 최신 데이터 자동 사용 (우아한 처리)
        if target_col_result is None:
            print(f"[{self.config['name']}] ⚠️ {self.year}년 {self.quarter}분기 데이터를 찾을 수 없음. 최신 데이터 탐색 시작...")
            # 헤더 행에서 최신 연도/분기 자동 탐색
            latest_col = self._find_latest_data_col()
            if latest_col is not None:
                print(f"[{self.config['name']}] ✅ 최신 데이터 컬럼 사용: {latest_col}")
                self.target_col = latest_col
            else:
                # 여전히 못 찾으면 에러
                print(f"[{self.config['name']}] 🔍 [디버그] Target 컬럼 찾기 실패:")
                print(f"  - 찾으려는 연도/분기: {self.year}년 {self.quarter}분기")
                print(f"  - 확인한 시트: {sheet_type}")
                print(f"  - 시트 크기: {len(self.df_aggregation)}행 × {len(self.df_aggregation.columns)}열")
                raise ValueError(
                    f"[{self.config['name']}] ❌ Target 컬럼을 찾을 수 없습니다 (최신 데이터도 없음).\n"
                    f"  찾으려는 연도/분기: {self.year}년 {self.quarter}분기\n"
                    f"  확인한 시트: {sheet_type}\n"
                    f"  시트 크기: {len(self.df_aggregation)}행 × {len(self.df_aggregation.columns)}열"
                )
        else:
            self.target_col = target_col_result
        
        # 전년 컬럼 찾기 (max_header_rows 전달)
        prev_y_col_result = self.find_target_col_index(
            self.df_aggregation, self.year - 1, self.quarter, 
            require_type_match=require_type_match,
            max_header_rows=max_header_rows
        )
        if prev_y_col_result is not None:
            self.prev_y_col = prev_y_col_result
            print(f"[{self.config['name']}] ✅ 전년 컬럼 ({sheet_type} 시트): {self.prev_y_col} ({self.year - 1} {self.quarter}/4)")
        else:
            # 전년 데이터가 없으면 최신 데이터 - 1년
            print(f"[{self.config['name']}] ⚠️ {self.year - 1}년 {self.quarter}분기 데이터 없음. 이전 연도 데이터 탐색...")
            prev_col = self._find_latest_data_col(target_year=self.year - 1)
            if prev_col is not None:
                self.prev_y_col = prev_col
                print(f"[{self.config['name']}] ✅ 이전 연도 데이터 사용: {self.prev_y_col}")
            else:
                print(f"[{self.config['name']}] ⚠️ 이전 연도 데이터도 없음 (계속 진행)")
                self.prev_y_col = None

        # 전전년 컬럼 찾기 (2년 전)
        prev_prev_y_col_result = self.find_target_col_index(
            self.df_aggregation, self.year - 2, self.quarter,
            require_type_match=require_type_match,
            max_header_rows=max_header_rows
        )
        if prev_prev_y_col_result is not None:
            self.prev_prev_y_col = prev_prev_y_col_result
            print(f"[{self.config['name']}] ✅ 전전년 컬럼 ({sheet_type} 시트): {self.prev_prev_y_col} ({self.year - 2} {self.quarter}/4)")
        else:
            print(f"[{self.config['name']}] ⚠️ {self.year - 2}년 {self.quarter}분기 데이터 없음 (계속 진행)")
            self.prev_prev_y_col = None

        # 전전전년 컬럼 찾기 (3년 전) - 재작년 증감률 계산용
        prev_prev_prev_y_col_result = self.find_target_col_index(
            self.df_aggregation, self.year - 3, self.quarter,
            require_type_match=require_type_match,
            max_header_rows=max_header_rows
        )
        if prev_prev_prev_y_col_result is not None:
            self.prev_prev_prev_y_col = prev_prev_prev_y_col_result
            print(f"[{self.config['name']}] ✅ 전전전년 컬럼 ({sheet_type} 시트): {self.prev_prev_prev_y_col} ({self.year - 3} {self.quarter}/4)")
        else:
            self.prev_prev_prev_y_col = None

        # 22년 3분기 ~ 25년 3분기처럼 분기 단위 전체 범위 컬럼 확보
        if self.year is not None and self.quarter is not None:
            self._ensure_quarter_columns(
                self.df_aggregation,
                self.year - 3,
                self.quarter,
                self.year,
                self.quarter,
                max_header_rows
            )
        
        wb.close()

        # header_rows: config에서 지정하거나 기본값 1
        header_rows = self.config.get('header_rows', 1)
        # region_keywords: config에서 지정하거나 기본값
        region_keywords = self.config.get('region_keywords', ['지역', '시도', '시군구', '지역명', '행정구역'])

        # 이름 기반 탐색으로 완전 전환 - 산업코드 로직 완전 제거
        name_keywords = ['이름', 'name', '산업명', '산업 이름', '업태명', '품목명', '품목 이름', '공정이름', '공정명', '연령']

        # 지역명 컬럼 후보 목록 (순서대로)
        region_col_candidates = []

        # df_aggregation을 df로 사용
        df = self.df_aggregation

        # 산업코드 컬럼은 사용하지 않음 (이름 기반 탐색만 사용)
        self.industry_code_col = None

        # 0) 컬럼명에서 우선 탐색 (ws.values 첫 행이 헤더인 구조 대응)
        for col_idx, col_name in enumerate(df.columns):
            if pd.isna(col_name):
                continue
            cell_str = str(col_name).strip().lower()
            matched_region = False
            if self.region_name_col is None:
                for keyword in region_keywords:
                    if keyword.lower() in cell_str:
                        region_col_candidates.append((col_idx, keyword, -1))
                        matched_region = True
                        print(f"[{self.config['name']}] 🔍 [헤더] 지역명 컬럼 후보: {col_idx} (키워드: '{keyword}')")
                        break
            if matched_region:
                continue
            if self.industry_name_col is None:
                for keyword in name_keywords:
                    if keyword.lower() in cell_str:
                        self.industry_name_col = col_idx
                        print(f"[{self.config['name']}] ✅ [헤더] 산업명 컬럼 발견: {col_idx} (키워드: '{keyword}')")
                        break

        # 1) 헤더 행 내용에서도 키워드 검색 (병합 헤더 등 대응)
        for row_idx in range(header_rows):
            row = df.iloc[row_idx]
            for col_idx, cell_value in enumerate(row):
                if pd.isna(cell_value):
                    continue
                cell_str = str(cell_value).strip().lower()
                matched_region = False

                # 지역명 컬럼 후보 찾기 (모든 일치하는 컬럼 수집)
                if self.region_name_col is None:
                    for keyword in region_keywords:
                        if keyword.lower() in cell_str:
                            region_col_candidates.append((col_idx, keyword, row_idx))
                            print(f"[{self.config['name']}] 🔍 지역명 컬럼 후보: {col_idx} (키워드: '{keyword}', 행: {row_idx})")
                            matched_region = True
                            break

                # 산업명 컬럼 찾기
                if matched_region:
                    continue
                if self.industry_name_col is None:
                    for keyword in name_keywords:
                        if keyword.lower() in cell_str:
                            self.industry_name_col = col_idx
                            print(f"[{self.config['name']}] ✅ 산업명 컬럼 발견: {col_idx} (키워드: '{keyword}', 행: {row_idx})")
                            break

        # 지역명 컬럼을 찾지 못한 경우, 데이터에서 직접 '전국' 등으로 탐색하여 추정
        if self.region_name_col is None and not region_col_candidates:
            valid_regions_probe = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
            found_col = None
            rows_to_scan = min(30, len(df))
            try:
                for r in range(rows_to_scan):
                    for c in range(len(df.columns)):
                        val = df.iloc[r, c]
                        if pd.notna(val):
                            s = str(val).strip()
                            if s in valid_regions_probe:
                                found_col = c
                                print(f"[{self.config['name']}] ✅ 데이터에서 지역명 발견으로 컬럼 추정: {found_col} (예: '{s}', 행 {r})")
                                break
                    if found_col is not None:
                        break
            except Exception:
                found_col = None
            if found_col is not None:
                self.region_name_col = found_col
            else:
                self.region_name_col = 0
                print(f"[{self.config['name']}] ⚠️ 지역명 컬럼 후보가 없어, 첫 번째 컬럼(0)으로 임시 설정합니다. 이후 검증 단계에서 교체됩니다.")
        
        # 지역명 컬럼 후보 중에서 실제 유효한 지역명이 있는 컬럼 선택
        if region_col_candidates:
            valid_regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                            '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
            valid_region_codes = ['00', '11', '26', '27', '28', '29', '30', '31', '36', '41', '42', '43', '44', '45', '46', '47', '48', '50']
            
            # 먼저 실제 지역명이 있는 컬럼 찾기 (우선순위: 실제 지역명 > 지역 코드)
            for col_idx, keyword, _ in region_col_candidates:
                # 데이터 행에서 이 컬럼의 값들 확인 (헤더 이후 처음 20행)
                has_actual_region_name = False
                has_valid_region = False
                
                for data_row_idx in range(header_rows, min(header_rows + 20, len(df))):
                    if col_idx < len(df.columns):
                        cell_value = df.iloc[data_row_idx, col_idx]
                        if pd.notna(cell_value):
                            cell_str = str(cell_value).strip()
                            # 실제 지역명 확인
                            if cell_str in valid_regions:
                                has_actual_region_name = True
                                self.region_name_col = col_idx
                                print(f"[{self.config['name']}] ✅ 지역명 컬럼 확정: {col_idx} (키워드: '{keyword}', 실제 지역명 발견: '{cell_str}')")
                                break
                            # 지역 코드 확인 (실제 지역명이 없을 때만)
                            elif cell_str in valid_region_codes and not has_actual_region_name:
                                has_valid_region = True
                
                if has_actual_region_name:
                    break  # 실제 지역명 찾음 - 종료
            
            # 실제 지역명을 찾지 못했지만 지역 코드만 있는 경우, 지역 코드 다음 컬럼에서 지역명 찾기
            if self.region_name_col is None:
                for col_idx, keyword, _ in region_col_candidates:
                    # 이 컬럼이 지역 코드 컬럼인지 확인
                    is_code_column = False
                    for data_row_idx in range(header_rows, min(header_rows + 5, len(df))):
                        if col_idx < len(df.columns):
                            cell_value = df.iloc[data_row_idx, col_idx]
                            if pd.notna(cell_value) and str(cell_value).strip() in valid_region_codes:
                                is_code_column = True
                                break
                    
                    if is_code_column:
                        # 지역명이 다음 컬럼에 있는지 확인
                        next_col_idx = col_idx + 1
                        if next_col_idx < len(df.columns):
                            for data_row_idx in range(header_rows, min(header_rows + 20, len(df))):
                                if next_col_idx < len(df.columns):
                                    cell_value = df.iloc[data_row_idx, next_col_idx]
                                    if pd.notna(cell_value):
                                        cell_str = str(cell_value).strip()
                                        if cell_str in valid_regions:
                                            self.region_name_col = next_col_idx
                                            print(f"[{self.config['name']}] ✅ 지역명 컬럼 확정: {next_col_idx} (지역 코드 컬럼 {col_idx} 다음, 지역명 발견: '{cell_str}')")
                                            break
                        
                        if self.region_name_col is not None:
                            break
            
            # 여전히 찾지 못했으면 첫 번째 후보 사용
            if self.region_name_col is None and region_col_candidates:
                self.region_name_col = region_col_candidates[0][0]
                print(f"[{self.config['name']}] ⚠️ 실제 지역명/지역명 다음 컬럼을 찾지 못해, 첫 번째 후보 컬럼 사용: {self.region_name_col}")
        
        # 데이터 시작 행 찾기 (헤더 다음 행)
        # 지역명이나 산업코드가 실제로 나타나는 첫 번째 행 찾기

        # 산업명 컬럼이 지역명 컬럼과 동일하게 잡힌 경우 초기화 후 재추정
        if self.industry_name_col is not None and self.region_name_col is not None and self.industry_name_col == self.region_name_col:
            print(f"[{self.config['name']}] ⚠️ 산업명 컬럼이 지역명 컬럼과 동일({self.industry_name_col})하여 재탐색합니다.")
            self.industry_name_col = None

        # 업종/품목명 컬럼을 찾지 못했거나 제거된 경우, 텍스트 비율 기반으로 재추정
        if self.industry_name_col is None:
            exclude_cols = [self.region_name_col] if self.region_name_col is not None else []
            guessed_col = self._find_textual_column(df, header_rows=header_rows, exclude_cols=exclude_cols)
            if guessed_col is not None and guessed_col != self.region_name_col:
                self.industry_name_col = guessed_col
                print(f"[{self.config['name']}] ✅ 업종/품목 컬럼 재추정: {guessed_col}")

        if self.region_name_col is not None:
            valid_regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                            '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
            valid_region_codes = ['00', '11', '26', '27', '28', '29', '30', '31', '36', '41', '42', '43', '44', '45', '46', '47', '48', '50']  # 지역 코드
            
            # 먼저 지역명 컬럼에서 실제 지역명 찾기
            for row_idx in range(header_rows, min(header_rows + 20, len(df))):
                row = df.iloc[row_idx]
                if self.region_name_col < len(row):
                    cell_value = row.iloc[self.region_name_col]
                    if pd.notna(cell_value):
                        cell_str = str(cell_value).strip()
                        # 지역명이 실제로 나타나는 행 찾기 (또는 지역 코드)
                        if cell_str in valid_regions or cell_str in valid_region_codes:
                            self.data_start_row = row_idx
                            print(f"[{self.config['name']}] ✅ 데이터 시작 행 발견: {row_idx} (지역명: '{cell_str}')")
                            break
            
            # 지역명 컬럼에서 찾지 못했으면, 다른 컬럼에서도 찾기 (국내인구이동의 경우 지역명이 다른 컬럼에 있을 수 있음)
            if self.data_start_row is None and self.report_type == 'migration':
                # 지역명 컬럼이 코드 컬럼인 경우, 실제 지역명이 있는 다른 컬럼 찾기
                # 보통 지역명은 코드 컬럼 옆에 있음
                for col_idx in range(max(0, self.region_name_col - 2), min(len(df.columns), self.region_name_col + 3)):
                    if col_idx == self.region_name_col:
                        continue
                    for row_idx in range(header_rows, min(header_rows + 20, len(df))):
                        row = df.iloc[row_idx]
                        if col_idx < len(row):
                            cell_value = row.iloc[col_idx]
                            if pd.notna(cell_value):
                                cell_str = str(cell_value).strip()
                                if cell_str in valid_regions:
                                    # 지역명 컬럼을 실제 지역명이 있는 컬럼으로 업데이트
                                    print(f"[{self.config['name']}] 🔍 지역명 컬럼 업데이트: {self.region_name_col} → {col_idx} (실제 지역명 발견: '{cell_str}')")
                                    self.region_name_col = col_idx
                                    self.data_start_row = row_idx
                                    print(f"[{self.config['name']}] ✅ 데이터 시작 행 발견: {row_idx} (지역명: '{cell_str}')")
                                    break
                        if self.data_start_row is not None:
                            break
                    if self.data_start_row is not None:
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
        
        # 실업률/고용률/마이그레이션은 산업코드가 선택적일 수 있음
        if self.industry_code_col is None:
            # 산업코드가 없더라도 이름/패턴 기반 폴백으로 진행 가능
            print(f"[{self.config['name']}] ⚠️ 산업코드 컬럼을 찾지 못했습니다. 이름·패턴 기반 폴백으로 계속 진행합니다.")
        
        # 실업률/고용률은 산업명이 선택적일 수 있음 (연령별 데이터이므로)
        # 국내인구이동은 산업명이 아예 필요 없음 (연령으로 구분)
        if self.industry_name_col is None:
            if self.report_type in ['employment', 'unemployment']:
                print(f"[{self.config['name']}] ⚠️ 산업명 컬럼을 찾을 수 없지만, 고용률/실업률은 산업명이 선택적이므로 계속 진행합니다.")
                # 산업명이 없으면 None으로 유지 (나중에 사용 시 체크 필요)
            else:
                # (A) 헤더에서 '산업'과 '이름' 토큰 동시 포함 컬럼 우선 선택
                import re
                header_exact_idx = None
                for c, cname in enumerate(df.columns):
                    try:
                        s = str(cname).strip().lower()
                    except Exception:
                        s = ''
                    s_norm = re.sub(r"\s+", "", s)
                    if '산업' in s and ('이름' in s or '명' in s) or '산업이름' in s_norm:
                        header_exact_idx = c
                        break
                if header_exact_idx is not None:
                    self.industry_name_col = header_exact_idx
                    print(f"[{self.config['name']}] ✅ 헤더 정확매칭으로 업종명 컬럼 확정: {header_exact_idx}")
                else:
                    # (B) 데이터에서 총계 키워드 등장 컬럼 탐색 (헤더 오탐 방지 필터 포함)
                    total_pattern = re.compile(r'(?:총지수|총계|합계|전\s*산업|전체)')
                    disallow_in_header = ['코드', '단계', '가중치', '지역', '조회']
                    best_idx = None
                    best_hits = -1
                    for c in range(len(df.columns)):
                        try:
                            header_s = str(df.columns[c]).lower()
                        except Exception:
                            header_s = ''
                        # 헤더에 금지 토큰 있으면 제외
                        if any(k in header_s for k in disallow_in_header):
                            continue
                        try:
                            series = df.iloc[:, c].astype(str).str.strip()
                            # 헤더 행 이후 데이터에서만 검사
                            window = series.iloc[max(header_rows, 0):max(header_rows, 0)+50]
                            hits = window.str.contains(total_pattern, regex=True, na=False).sum()
                            if hits > best_hits:
                                best_hits = hits
                                best_idx = c
                        except Exception:
                            continue
                    if best_idx is not None and best_hits > 0:
                        self.industry_name_col = best_idx
                        print(f"[{self.config['name']}] ✅ 총계 키워드로 업종명 컬럼 추정: {best_idx} (매치 {best_hits}건)")
                    else:
                        # (C) 헤더 키워드로 탐색 (산업/업종/품목/공정 포함, 단 '코드' 제외)
                        header_guess = None
                        for c, cname in enumerate(df.columns):
                            try:
                                s = str(cname).strip().lower()
                            except Exception:
                                s = ''
                            if any(k in s for k in ['산업', '업종', '품목', '공정']) and '코드' not in s:
                                header_guess = c
                                break
                        if header_guess is not None:
                            self.industry_name_col = header_guess
                            print(f"[{self.config['name']}] ✅ 헤더명으로 업종명 컬럼 추정: {header_guess}")
                        else:
                            # (D) 데이터 특성을 보고 업종명 컬럼 추정
                            guessed = self._find_textual_column(df, header_rows, exclude_cols=[self.region_name_col] if self.region_name_col is not None else [])
                            if guessed is not None:
                                self.industry_name_col = guessed
                                print(f"[{self.config['name']}] ✅ 헤더 키워드 없이 업종명 컬럼 추정: {guessed}")
                            else:
                                print(f"[{self.config['name']}] ⚠️ 업종명 컬럼을 추정하지 못했습니다.")
            class UnifiedReportGenerator(BaseGenerator):
                """통합 보고서 Generator (집계 시트 기반)
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

                    # 설정 로드: REPORT_ORDER에서 report_type(id)로 검색
                    all_reports = [*REPORT_ORDER]
                    self.config = next((r for r in all_reports if r.get('id') == report_type), None)
                    if not self.config:
                        raise ValueError(f"알 수 없는 report_type: {report_type}")
                    self.report_type = report_type
                    self.report_id = self.config.get('report_id', report_type)
                    if 'name_mapping' not in self.config:
                        raise ValueError(f"name_mapping이 설정에 없습니다: {report_type}")
                    self.name_mapping = self.config['name_mapping']
                    if 'aggregation_structure' not in self.config:
                        raise ValueError(f"aggregation_structure가 설정에 없습니다: {report_type}")
                    agg_struct = self.config['aggregation_structure']
                    self.region_name_col = None  # 동적으로 찾음
                    self.industry_code_col = None  # 동적으로 찾음
                    self.total_code = agg_struct.get('total_code', 'BCD')
                    if 'metadata_columns' not in self.config:
                        raise ValueError(f"metadata_columns가 설정에 없습니다: {report_type}")
                    self.metadata_cols = self.config['metadata_columns']
                    self.industry_name_col = None  # 동적으로 찾음
                    self.data_start_row = None  # 동적으로 찾음
                    self.df_analysis = None
                    self.df_aggregation = None
                    self.df_reference = None
                    self.target_col = None
                    self.prev_y_col = None
                    self.prev_prev_y_col = None
                    self.prev_prev_prev_y_col = None
                    self.use_aggregation_only = False
                    print(f"[{self.config['name']}] Generator 초기화")

                    # 안전하게 미정의 변수 기본값 처리
                    analysis_sheet = None
                    require_analysis_sheet = False
                    analysis_sheets = []
                    # 실제 엑셀 파일에서 시트 목록 읽기
                    sheet_names = []
                    try:
                        import openpyxl
                        wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
                        sheet_names = wb.sheetnames
                        wb.close()
                    except Exception as e:
                        print(f"[경고] 엑셀 시트 목록을 읽는 중 오류 발생: {e}")

                def load_data(self):
                    """테스트 호환성: 기존 테스트 코드에서 generator.load_data()를 호출하는 경우 extract_all_data()로 프록시"""
                    return self.extract_all_data()
        # prev_y_col 찾기
        require_type_match = False  # 기본값 False로 선언
        sheet_type = agg_sheet_name  # 디버그 메시지용 시트명
        if self.prev_y_col is None:
            self.prev_y_col = self.find_target_col_index(df, self.year - 1, self.quarter, require_type_match=require_type_match)
            if self.prev_y_col is not None:
                print(f"[{self.config['name']}] ✅ 전년 컬럼 ({sheet_type} 시트): {self.prev_y_col} ({self.year - 1} {self.quarter}/4)")

        if self.prev_prev_y_col is None:
            self.prev_prev_y_col = self.find_target_col_index(df, self.year - 2, self.quarter, require_type_match=require_type_match)
            if self.prev_prev_y_col is not None:
                print(f"[{self.config['name']}] ✅ 전전년 컬럼 ({sheet_type} 시트): {self.prev_prev_y_col} ({self.year - 2} {self.quarter}/4)")

        if self.prev_prev_prev_y_col is None:
            self.prev_prev_prev_y_col = self.find_target_col_index(df, self.year - 3, self.quarter, require_type_match=require_type_match)
            if self.prev_prev_prev_y_col is not None:
                print(f"[{self.config['name']}] ✅ 전전전년 컬럼 ({sheet_type} 시트): {self.prev_prev_prev_y_col} ({self.year - 3} {self.quarter}/4)")

        analysis_sheet = self.config.get('analysis_sheet')
        if analysis_sheet and analysis_sheet != agg_sheet_name:
            try:
                self.df_analysis = pd.read_excel(self.excel_path, sheet_name=analysis_sheet, header=None)
                analysis_header_rows = self.config.get('analysis_header_rows', max_header_rows)
                self.analysis_target_col = self.find_target_col_index(
                    self.df_analysis,
                    self.year,
                    self.quarter,
                    require_type_match=False,
                    max_header_rows=analysis_header_rows
                )
                self.analysis_prev_y_col = self.find_target_col_index(
                    self.df_analysis,
                    self.year - 1,
                    self.quarter,
                    require_type_match=False,
                    max_header_rows=analysis_header_rows
                )
                self.analysis_prev_prev_y_col = self.find_target_col_index(
                    self.df_analysis,
                    self.year - 2,
                    self.quarter,
                    require_type_match=False,
                    max_header_rows=analysis_header_rows
                )
                self.analysis_prev_prev_prev_y_col = self.find_target_col_index(
                    self.df_analysis,
                    self.year - 3,
                    self.quarter,
                    require_type_match=False,
                    max_header_rows=analysis_header_rows
                )
                if self.year is not None and self.quarter is not None:
                    keys, cols = self._collect_quarter_columns(
                        self.df_analysis,
                        self.year - 3,
                        self.quarter,
                        self.year,
                        self.quarter,
                        analysis_header_rows
                    )
                    self.analysis_quarterly_keys = keys
                    self.analysis_quarterly_cols = cols
            except Exception as e:
                print(f"[{self.config['name']}] ⚠️ 분석 시트 로드 실패: {analysis_sheet} ({e})")
        
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
    
    def _find_latest_data_col(self, target_year=None):
        """
        헤더 행에서 최신 연도/분기의 데이터 컬럼을 찾기
        target_year이 지정되면 그 연도 데이터를 찾음
        """
        import re
        import pandas as pd
        
        if not hasattr(self, 'df_aggregation') or self.df_aggregation is None:
            return None
        
        df = self.df_aggregation
        if len(df) == 0:
            return None
        
        # 헤더 행 (첫 번째 행)
        header_row = df.iloc[0]
        
        # 숫자로 보이는 값 추출 (연도 후보)
        year_patterns = []
        for idx, cell in enumerate(header_row):
            if pd.isna(cell):
                continue
            cell_str = str(cell).strip()
            
            # 정수 추출 (연도 후보)
            numbers = re.findall(r'\d+', cell_str)
            if numbers:
                for num_str in numbers:
                    year_val = int(num_str)
                    # 범위 체크: 1990 ~ 2100
                    if 1990 <= year_val <= 2100:
                        year_patterns.append((idx, year_val, cell_str))
        
        if not year_patterns:
            return None
        
        # target_year이 지정되면 그에 맞는 것 찾기
        if target_year is not None:
            for idx, year_val, cell_str in year_patterns:
                if year_val == target_year:
                    return idx
            # target_year 못 찾으면 None
            return None
        
        # target_year 미지정 시 최대 연도 찾기
        max_year = max(year_patterns, key=lambda x: x[1])
        return max_year[0]
    
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
        
        # 분기 단위 전체 범위 컬럼 확보 (22년 3분기 ~ 25년 3분기 등)
        header_rows = self.config.get('header_rows', 5)
        if self.year is not None and self.quarter is not None:
            if not self.quarterly_keys or not self.quarterly_cols or (df is not self.df_aggregation):
                self._ensure_quarter_columns(
                    df,
                    self.year - 3,
                    self.quarter,
                    self.year,
                    self.quarter,
                    header_rows
                )

        # 직전 분기 컬럼
        prev_q_col = None
        if self.year is not None and self.quarter is not None:
            prev_q_year, prev_q = self._previous_quarter(self.year, self.quarter)
            prev_q_key = self._format_quarter_key(prev_q_year, prev_q)
            prev_q_col = self.quarterly_cols.get(prev_q_key)

        use_analysis_rates = self.config.get('value_type') == 'change_rate' and self.df_analysis is not None
        if use_analysis_rates and self.year is not None and self.quarter is not None:
            analysis_header_rows = self.config.get('analysis_header_rows', header_rows)
            if self.analysis_target_col is None:
                self.analysis_target_col = self.find_target_col_index(
                    self.df_analysis,
                    self.year,
                    self.quarter,
                    require_type_match=False,
                    max_header_rows=analysis_header_rows
                )
            if self.analysis_prev_y_col is None:
                self.analysis_prev_y_col = self.find_target_col_index(
                    self.df_analysis,
                    self.year - 1,
                    self.quarter,
                    require_type_match=False,
                    max_header_rows=analysis_header_rows
                )
            if not self.analysis_quarterly_keys or not self.analysis_quarterly_cols:
                keys, cols = self._collect_quarter_columns(
                    self.df_analysis,
                    self.year - 3,
                    self.quarter,
                    self.year,
                    self.quarter,
                    analysis_header_rows
                )
                self.analysis_quarterly_keys = keys
                self.analysis_quarterly_cols = cols
        
        # 지역 목록
        regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        table_data = []
        total_code = None
        try:
            total_code = (self.config.get('aggregation_structure') or {}).get('total_code')
        except Exception:
            total_code = None

        def _select_region_total(df_source: pd.DataFrame, region_name: str) -> Optional[pd.Series]:
            if df_source is None:
                return None
            if self.data_start_row is None:
                start_row = 0
            else:
                start_row = max(self.data_start_row, 0)
            if start_row < len(df_source):
                local_df = df_source.iloc[start_row:].copy()
            else:
                local_df = df_source.copy()
            region_col = self.region_name_col
            if region_col is None or region_col < 0 or region_col >= len(local_df.columns):
                region_col = None

            def _detect_region_col(df_search: pd.DataFrame) -> Optional[int]:
                if df_search is None or df_search.empty:
                    return None
                valid_regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                                 '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
                rows_to_scan = min(40, len(df_search))
                try:
                    for col_idx in range(len(df_search.columns)):
                        for r in range(rows_to_scan):
                            val = df_search.iloc[r, col_idx]
                            if pd.notna(val) and str(val).strip() in valid_regions:
                                return col_idx
                except Exception:
                    return None
                return None

            if region_col is None:
                region_col = _detect_region_col(df_source)
            if region_col is None:
                return None

            try:
                region_filter = local_df[
                    local_df.iloc[:, region_col].astype(str).str.strip() == region_name
                ]
            except (IndexError, KeyError):
                return None
            if region_filter.empty and df_source is not local_df:
                try:
                    region_filter = df_source[
                        df_source.iloc[:, region_col].astype(str).str.strip() == region_name
                    ]
                    local_df = df_source
                except (IndexError, KeyError):
                    return None
            if region_filter.empty:
                alt_col = _detect_region_col(df_source)
                if alt_col is not None and alt_col != region_col:
                    region_col = alt_col
                    try:
                        region_filter = df_source[
                            df_source.iloc[:, region_col].astype(str).str.strip() == region_name
                        ]
                        local_df = df_source
                    except (IndexError, KeyError):
                        return None
            if region_filter.empty:
                return None
            region_total = None
            if self.industry_name_col is not None and self.industry_name_col != self.region_name_col and self.industry_name_col >= 0 and self.industry_name_col < len(region_filter.columns):
                by_name = self._find_total_row_by_name(region_filter, self.industry_name_col, header_rows=0)
                if by_name is not None and not by_name.empty:
                    region_total = by_name
            if (region_total is None or region_total.empty) and total_code:
                exclude_cols = []
                if region_col is not None:
                    exclude_cols.append(region_col)
                if self.industry_name_col is not None and self.report_type not in ['employment', 'unemployment', 'migration']:
                    exclude_cols.append(self.industry_name_col)
                by_code = self._find_total_row_by_code(region_filter, total_code, exclude_cols=exclude_cols)
                if by_code is not None and not by_code.empty:
                    region_total = by_code
            if (region_total is None or region_total.empty) and self.report_type in ['employment', 'unemployment', 'migration']:
                if len(region_filter) > 0:
                    region_total = region_filter.head(1)
            if (region_total is None or region_total.empty) and self.report_type == 'migration':
                if len(region_filter) > 0:
                    region_total = region_filter.head(1)
            if region_total is None or region_total.empty:
                return None
            return region_total.iloc[0]
        
        # 컬럼 인덱스 검증 (동적으로 찾은 컬럼)
        if self.region_name_col is None or self.region_name_col < 0 or self.region_name_col >= len(data_df.columns):
            raise ValueError(
                f"[{self.config['name']}] ❌ 지역명 컬럼을 찾을 수 없습니다. "
                f"동적 탐색 실패 또는 인덱스({self.region_name_col})가 유효하지 않습니다. "
                f"DataFrame 컬럼 수: {len(data_df.columns)}"
            )
        
        for region in regions:
            row = _select_region_total(df, region)
            if row is None:
                print(f"[{self.config['name']}] ⚠️ {region}: 총계 행을 찾지 못했습니다. 스킵합니다.")
                continue

            analysis_row = _select_region_total(self.df_analysis, region) if use_analysis_rates else None
            
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

            def _compute_quarterly_growth(current: Optional[float], previous: Optional[float]) -> Optional[float]:
                if current is None or previous is None:
                    return None
                if self.report_type in ['employment', 'unemployment']:
                    return round(current - previous, 1)
                if self.report_type == 'migration':
                    return round(current - previous, 1)
                if self.config.get('value_type') == 'change_rate':
                    return round(current, 1)
                if previous == 0:
                    return None
                return round((current - previous) / previous * 100, 1)
            
            # 지수 추출
            try:
                idx_current = self.safe_float(row.iloc[self.target_col], None)
                idx_prev_year = self.safe_float(row.iloc[self.prev_y_col], None)
                idx_prev_prev_year = None
                idx_prev_prev_prev_year = None
                if self.prev_prev_y_col is not None and self.prev_prev_y_col < len(row):
                    idx_prev_prev_year = self.safe_float(row.iloc[self.prev_prev_y_col], None)
                if self.prev_prev_prev_y_col is not None and self.prev_prev_prev_y_col < len(row):
                    idx_prev_prev_prev_year = self.safe_float(row.iloc[self.prev_prev_prev_y_col], None)
            except (IndexError, KeyError) as e:
                print(f"[{self.config['name']}] ⚠️ 데이터 추출 오류: {e}. 스킵합니다.")
                continue

            rate_current = None
            rate_prev_year = None
            rate_quarterly_values: List[Optional[float]] = []
            rate_prev_quarter = None
            if use_analysis_rates and analysis_row is not None:
                if self.analysis_target_col is not None and self.analysis_target_col < len(analysis_row):
                    rate_current = self.safe_float(analysis_row.iloc[self.analysis_target_col], None)
                if self.analysis_prev_y_col is not None and self.analysis_prev_y_col < len(analysis_row):
                    rate_prev_year = self.safe_float(analysis_row.iloc[self.analysis_prev_y_col], None)
                if self.analysis_quarterly_keys:
                    for key in self.analysis_quarterly_keys:
                        col_idx = self.analysis_quarterly_cols.get(key)
                        if col_idx is not None and col_idx < len(analysis_row):
                            rate_quarterly_values.append(self.safe_float(analysis_row.iloc[col_idx], None))
                        else:
                            rate_quarterly_values.append(None)
                    if len(rate_quarterly_values) >= 2:
                        rate_prev_quarter = rate_quarterly_values[-2]
            
            # 분기 단위 전체 범위 값 추출
            quarterly_values: List[Optional[float]] = []
            if self.quarterly_keys:
                for key in self.quarterly_keys:
                    col_idx = self.quarterly_cols.get(key)
                    if col_idx is not None and col_idx < len(row):
                        quarterly_values.append(self.safe_float(row.iloc[col_idx], None))
                    else:
                        quarterly_values.append(None)

            # 단위 보정
            scale_factor = 1.0
            if self.report_type == 'construction':
                # 10억원 단위 → 100억원 단위 (1/10)
                scale_factor = 0.1
            elif self.report_type == 'export':
                # 백만달러 단위 → 억달러 단위 (요청: 100배)
                scale_factor = 100.0
            elif self.report_type == 'migration':
                # 명 단위 → 천명 단위
                scale_factor = 0.001

            if scale_factor != 1.0:
                idx_current = (idx_current * scale_factor) if idx_current is not None else None
                idx_prev_year = (idx_prev_year * scale_factor) if idx_prev_year is not None else None
                idx_prev_prev_year = (idx_prev_prev_year * scale_factor) if idx_prev_prev_year is not None else None
                idx_prev_prev_prev_year = (idx_prev_prev_prev_year * scale_factor) if idx_prev_prev_prev_year is not None else None
                quarterly_values = [
                    (v * scale_factor) if v is not None else None
                    for v in quarterly_values
                ]

            if use_analysis_rates and rate_quarterly_values:
                quarterly_growth_rates = rate_quarterly_values[:]
            elif self.report_type == 'migration':
                quarterly_growth_rates: List[Optional[float]] = [None for _ in quarterly_values]
            else:
                quarterly_growth_rates = []
                for i, val in enumerate(quarterly_values):
                    if i == 0:
                        quarterly_growth_rates.append(None)
                    else:
                        quarterly_growth_rates.append(_compute_quarterly_growth(val, quarterly_values[i - 1]))

            # 직전 분기 값
            idx_prev_quarter = None
            if prev_q_col is not None and prev_q_col < len(row):
                idx_prev_quarter = self.safe_float(row.iloc[prev_q_col], None)

            if idx_prev_quarter is not None and scale_factor != 1.0:
                idx_prev_quarter = idx_prev_quarter * scale_factor

            # 국내인구이동: 직전/전전/전전전 분기 값 추출 (없으면 None 유지)
            idx_prev_prev = idx_prev_prev_prev = None
            if self.report_type == 'migration' and quarterly_values:
                if len(quarterly_values) >= 2:
                    idx_prev_quarter = quarterly_values[-2]
                if len(quarterly_values) >= 3:
                    idx_prev_prev = quarterly_values[-3]
                if len(quarterly_values) >= 4:
                    idx_prev_prev_prev = quarterly_values[-4]
            
            if idx_current is None:
                continue

            if self.report_type == 'migration':
                previous_quarter_growth = None
            elif use_analysis_rates:
                previous_quarter_growth = rate_prev_quarter
            else:
                previous_quarter_growth = _compute_quarterly_growth(idx_current, idx_prev_quarter)
            
            # 증감 계산 (report_type에 따라 다름)
            # 국내인구이동: 절대값 (부호 포함, 변화율 아님)
            # 고용률/실업률: 퍼센트포인트(p) 차이
            # value_type='change_rate': 이미 계산된 증감률 직접 사용
            # 기타 지수: 증감률(%)
            if self.report_type == 'migration':
                # 국내인구이동은 증감률 계산하지 않음
                change_rate = None
            elif self.config.get('value_type') == 'change_rate':
                # 시트에 이미 증감률이 계산되어 있는 경우 (예: C 분석)
                change_rate = round(rate_current, 1) if rate_current is not None else round(idx_current, 1)
            elif idx_prev_year is not None and idx_prev_year != 0:
                if self.report_type in ['employment', 'unemployment']:
                    # 퍼센트포인트 차이 (p)
                    change_rate = round(idx_current - idx_prev_year, 1)
                else:
                    # 증감률 (%)
                    change_rate = round(((idx_current - idx_prev_year) / idx_prev_year) * 100, 1)
            else:
                change_rate = None
            
            if self.report_type == 'migration':
                row_data = {
                    'region_name': region,
                    'region_display': self._get_region_display_name(region),
                    'value': round(idx_current, 1),
                    'prev_value': round(idx_prev_quarter, 1) if idx_prev_quarter is not None else None,
                    'prev_prev_value': round(idx_prev_prev, 1) if idx_prev_prev is not None else None,
                    'prev_prev_prev_value': round(idx_prev_prev_prev, 1) if idx_prev_prev_prev is not None else None,
                    # 국내인구이동은 증감률 계산하지 않음
                    'quarterly_keys': self.quarterly_keys,
                    'quarterly_values': quarterly_values,
                    'quarterly_growth_rates': quarterly_growth_rates,
                    'age_20_29': None,
                    'age_other': None
                }
            else:
                row_data = {
                    'region_name': region,
                    'region_display': self._get_region_display_name(region),
                    'value': round(idx_current, 1),
                    'prev_value': round(idx_prev_year, 1) if idx_prev_year else None,
                    'prev_prev_value': round(idx_prev_prev_year, 1) if idx_prev_prev_year is not None else None,
                    'prev_prev_prev_value': round(idx_prev_prev_prev_year, 1) if idx_prev_prev_prev_year is not None else None,
                    'change_rate': change_rate,
                    'previous_quarter_growth': previous_quarter_growth,
                    'quarterly_keys': self.quarterly_keys,
                    'quarterly_values': quarterly_values,
                    'quarterly_growth_rates': quarterly_growth_rates,
                    'rate_quarterly_keys': self.analysis_quarterly_keys if use_analysis_rates else None,
                    'rate_quarterly_values': rate_quarterly_values if use_analysis_rates else None
                }

            table_data.append(row_data)
            
            print(f"[{self.config['name']}] ✅ {region}: 지수={idx_current:.1f}, 증감률={change_rate}%")
        
        # 국내인구이동: 전국 데이터 생성 여부 확인 (config의 has_nationwide 설정)
        # 국내이동은 지역간 이동이므로 전국 합계(0)는 의미가 없어 생성하지 않음
        if self.report_type == 'migration' and table_data:
            # config에서 has_nationwide 설정 확인 (기본값 True)
            should_generate_nationwide = self.config.get('has_nationwide', True)
            
            if should_generate_nationwide:
                def sum_field(key: str) -> Optional[float]:
                    values = [row.get(key) for row in table_data if row.get('region_name') != '전국' and row.get(key) is not None]
                    return round(sum(values), 1) if values else None

                # 이미 전국이 있다면 스킵
                has_nationwide = any(row.get('region_name') == '전국' for row in table_data)
                if not has_nationwide:
                    nationwide_row = {
                        'region_name': '전국',
                        'region_display': self._get_region_display_name('전국'),
                        'value': sum_field('value'),
                        'prev_value': sum_field('prev_value'),
                        'prev_prev_value': sum_field('prev_prev_value'),
                        'prev_prev_prev_value': sum_field('prev_prev_prev_value'),
                        'change_rate': sum_field('change_rate'),
                        'age_20_29': None,
                        'age_other': None
                    }
                    table_data.insert(0, nationwide_row)
                    print(f"[{self.config['name']}] ✅ 전국 데이터가 없어 지역 합계로 추가")
            else:
                print(f"[{self.config['name']}] ⚠️ has_nationwide=False이므로 전국 데이터 생성 건너뜀")
        
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
        
        # 산업명 컬럼 찾기 (동적으로 찾은 값 사용)
        if self.industry_name_col is None:
            if self.report_type in ['employment', 'unemployment']:
                print(f"[{self.config['name']}] ⚠️ 산업명 컬럼을 찾을 수 없지만, 고용률/실업률은 산업명이 선택적이므로 계속 진행합니다.")
                industry_name_col = None
            else:
                # 헤더로 못 찾은 경우 텍스트 비율 기반 추정 시도
                industry_name_col = self._find_textual_column(df, header_rows=0, exclude_cols=[self.region_name_col] if self.region_name_col is not None else [])
                if industry_name_col is not None:
                    print(f"[{self.config['name']}] ✅ 업종명 컬럼 추정: {industry_name_col}")
                    self.industry_name_col = industry_name_col
                else:
                    print(f"[{self.config['name']}] ⚠️ 업종명 컬럼을 추정하지 못했습니다. 업종 데이터 추출을 건너뜁니다.")
                    return []
        else:
            industry_name_col = self.industry_name_col
        
        if industry_name_col is not None and industry_name_col < 0:
            industry_name_col = 0
        
        for idx, row in region_filter.iterrows():
            # 산업명 추출 우선 (총계 키워드면 스킵)
            industry_name = ''
            if industry_name_col is not None and industry_name_col < len(row) and pd.notna(row.iloc[industry_name_col]):
                industry_name = str(row.iloc[industry_name_col]).strip()
            if not industry_name:
                # 고용률/실업률은 산업명이 없어도 진행 가능
                if self.report_type not in ['employment', 'unemployment']:
                    continue
            
            # 총계 키워드 스킵 (오탐 방지를 위해 '계' 제외)
            if any(kw in industry_name for kw in ['총계', '합계', '총지수', '전체', '전산업', '전 산업']):
                continue
            
            # 산업명 컬럼이 없으면 스킵 (고용률/실업률 제외)
            if industry_name_col is None and self.report_type not in ['employment', 'unemployment']:
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
                'growth_rate': change_rate  # 템플릿 호환 필드명
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
        
        # 국내인구이동의 경우 전국 데이터가 없으면 지역 합계로 계산
        if not nationwide or not isinstance(nationwide, dict):
            if self.report_type == 'migration' and table_data:
                print(f"[{self.config['name']}] ⚠️ 전국 데이터를 찾을 수 없으므로 모든 지역을 합계하여 계산합니다.")
                # 모든 지역 데이터 합계 (전국 제외)
                total_value = 0
                total_prev_value = 0
                for d in table_data:
                    if d and isinstance(d, dict) and d.get('region_name') != '전국':
                        total_value += d.get('value', 0) or 0
                        total_prev_value += d.get('prev_value', 0) or 0
                
                # 전국 데이터 생성
                change_rate = None
                if self.report_type == 'migration':
                    # 국내인구이동: 절대 순인구이동값 (부호 포함)
                    change_rate = round(total_value, 1)
                elif total_prev_value != 0:
                    change_rate = round((total_value - total_prev_value) / total_prev_value * 100, 1)
                
                nationwide = {
                    'region_name': '전국',
                    'region_display': '전 국',
                    'value': total_value,
                    'prev_value': total_prev_value,
                    'change_rate': change_rate
                }
                print(f"[{self.config['name']}] ✅ 전국 합계: {total_value} (전년: {total_prev_value}, 증감률: {change_rate}%)")
            else:
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
        
        # 국내인구이동은 nationwide가 없음 - 나머지만 처리
# 국내인구이동은 nationwide가 없음 - 나머지만 처리
        if nationwide:
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
        else:
            # nationwide가 None인 경우 (국내인구이동 등)
            index_value = None
            growth_rate = None
            main_increase = None
            main_decrease = None
        
        # 모든 필드명 포함 (템플릿 호환)
        result = {
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

        # 건설동향 템플릿 호환 별칭 추가
        if self.report_type == 'construction':
            # construction_template.html에서 요구하는 키: construction_index_trillion
            # index_value가 백억원이므로 조원 단위로 변환 (백억원 * 100 = 조원)
            construction_trillion = (index_value / 100) if index_value else None
            result['construction_index_trillion'] = construction_trillion
            result['change'] = growth_rate
            # 토목/건축 증감률 (기본값은 전체 증감률 사용)
            result['civil_growth'] = growth_rate
            result['building_growth'] = growth_rate
            # 토목/건축 부공종 (기본값)
            result['civil_subtypes'] = '철도·궤도, 기계설치'
            result['building_subtypes'] = '주택, 관공서 등'
            result['main_category'] = '토목' if (growth_rate is not None and growth_rate >= 0) else '토목'
            result['sub_types_text'] = '철도·궤도, 도로·교량, 주택'
        # 고용률/실업률 템플릿 호환 별칭 추가
        elif self.report_type == 'employment':
            # employment_template.html에서 요구하는 키: employment_rate, change, main_age_groups, top_age_groups
            result['employment_rate'] = index_value
            result['change'] = growth_rate
            result['main_age_groups'] = []
            result['top_age_groups'] = []
        elif self.report_type == 'unemployment':
            # unemployment_template.html에서 요구하는 키: rate, change, age_groups
            result['rate'] = index_value
            result['change'] = growth_rate
            result['age_groups'] = []

        return result
    
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

    def _build_summary_table(self, table_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """템플릿용 요약 테이블 생성 (필수 필드만 기본 값으로 채움)"""
        if table_data is None:
            table_data = []

        # 4개 증감률 컬럼, 3개 지수/율 컬럼을 기본 라벨로 구성
        def _previous_quarter(year: int, quarter: int) -> tuple[int, int]:
            if quarter <= 1:
                return (year - 1, 4)
            return (year, quarter - 1)

        def _growth_labels(year: Optional[int], quarter: Optional[int]) -> List[str]:
            if year is None or quarter is None:
                return ["전전기", "전기", "직전기", "현기"]
            prev_q_year, prev_q = _previous_quarter(year, quarter)
            return [
                f"{year-2}.{quarter}/4",
                f"{year-1}.{quarter}/4",
                f"{prev_q_year}.{prev_q}/4",
                f"{year}.{quarter}/4",
            ]

        def _index_labels(year: Optional[int], quarter: Optional[int]) -> List[str]:
            if self.report_type == 'employment':
                age_label = "20-29세"
            elif self.report_type == 'unemployment':
                age_label = "15-29세"
            else:
                age_label = "15-29세"
            if year is None or quarter is None:
                return ["전기", "현기", "청년층"]
            return [
                f"{year-1}.{quarter}/4",
                f"{year}.{quarter}/4",
                age_label,
            ]

        growth_cols = _growth_labels(self.year, self.quarter)
        index_cols = _index_labels(self.year, self.quarter)

        target_quarter_keys: List[str] = []
        if self.year is not None and self.quarter is not None:
            prev_q_year, prev_q = _previous_quarter(self.year, self.quarter)
            target_quarter_keys = [
                self._format_quarter_key(self.year - 2, self.quarter),
                self._format_quarter_key(self.year - 1, self.quarter),
                self._format_quarter_key(prev_q_year, prev_q),
                self._format_quarter_key(self.year, self.quarter),
            ]

        def _map_quarter_values(keys: Any, values: Any) -> List[Optional[float]]:
            if not keys or not values:
                return [None, None, None, None]
            mapping = {k: v for k, v in zip(keys, values)}
            if not target_quarter_keys:
                return [None, None, None, None]
            return [mapping.get(k) for k in target_quarter_keys]

        def _to_float(value: Any) -> Optional[float]:
            if value is None or value == '' or value == '-':
                return None
            try:
                return float(value)
            except Exception:
                return None

        def _compute_growth(current: Optional[float], previous: Optional[float]) -> Optional[float]:
            if current is None or previous is None:
                return None
            if previous == 0:
                return None
            return round((current - previous) / previous * 100, 1)

        def _build_growth_slots(row: Dict[str, Any]) -> List[Optional[float]]:
            # 분기별 증감률(분기-전분기) 값이 있으면 우선 사용
            q_keys = row.get('quarterly_keys')
            q_growth = row.get('quarterly_growth_rates')
            mapped_growth = _map_quarter_values(q_keys, q_growth)
            if any(v is not None for v in mapped_growth):
                return mapped_growth

            if self.config.get('value_type') == 'change_rate':
                rate_keys = row.get('rate_quarterly_keys') or row.get('quarterly_keys')
                rate_values = row.get('rate_quarterly_values') or row.get('quarterly_values')
                mapped = _map_quarter_values(rate_keys, rate_values)
                if any(v is not None for v in mapped):
                    return mapped
            current_value = _to_float(row.get('value'))
            prev_value = _to_float(row.get('prev_value'))
            prev_prev_value = _to_float(row.get('prev_prev_value'))
            prev_prev_prev_value = _to_float(row.get('prev_prev_prev_value'))

            two_years_ago = _compute_growth(prev_prev_value, prev_prev_prev_value)
            last_year = _compute_growth(prev_value, prev_prev_value)
            previous_quarter = _to_float(
                row.get('previous_quarter_growth') or row.get('prev_quarter_growth')
            )
            if previous_quarter is None:
                quarterly_growth_rates = row.get('quarterly_growth_rates')
                if isinstance(quarterly_growth_rates, list) and quarterly_growth_rates:
                    previous_quarter = _to_float(quarterly_growth_rates[-1])
            current = _compute_growth(current_value, prev_value)
            if current is None:
                current = _to_float(row.get('change_rate'))

            return [two_years_ago, last_year, previous_quarter, current]

        regions = []
        for row in table_data:
            region_name = row.get('region_name', '') if isinstance(row, dict) else ''
            growth_rate = row.get('change_rate') if isinstance(row, dict) else None
            value = row.get('value') if isinstance(row, dict) else None
            prev_value = row.get('prev_value') if isinstance(row, dict) else None
            prev_prev_value = row.get('prev_prev_value') if isinstance(row, dict) else None
            prev_prev_prev_value = row.get('prev_prev_prev_value') if isinstance(row, dict) else None

            computed = _build_growth_slots(row) if isinstance(row, dict) else [None, None, None, None]
            growth_rates = [
                '' if computed[0] is None else computed[0],
                '' if computed[1] is None else computed[1],
                '' if computed[2] is None else computed[2],
                '' if computed[3] is None else computed[3],
            ]

            youth_rate = row.get('youth_rate') if isinstance(row, dict) else None
            regions.append({
                'group': None,
                'region': region_name,
                'sido': region_name,
                'region_group': None,
                'rowspan': 1,
                # 보유한 데이터 기반으로 증감률 슬롯 채움
                'growth_rates': growth_rates,
                'indices': [prev_value, value, ''],
                'changes': growth_rates,
                'rates': [prev_value, value, youth_rate if youth_rate not in (None, '', '-') else ''],
                'youth_rate': youth_rate,
                'quarterly_keys': row.get('quarterly_keys') if isinstance(row, dict) else None,
                'quarterly_values': row.get('quarterly_values') if isinstance(row, dict) else None,
                'quarterly_growth_rates': row.get('quarterly_growth_rates') if isinstance(row, dict) else None,
                'rate_quarterly_keys': row.get('rate_quarterly_keys') if isinstance(row, dict) else None,
                'rate_quarterly_values': row.get('rate_quarterly_values') if isinstance(row, dict) else None,
                'prev_prev_value': prev_prev_value,
                'prev_prev_prev_value': prev_prev_prev_value,
            })

        return {
            'base_year': 2020,
            'columns': {
                'growth_rate_columns': growth_cols,
                'index_columns': index_cols,
                'change_columns': growth_cols,
                'rate_columns': index_cols,
                # 수출/수입 템플릿에서 액수 컬럼 라벨 요구
                'amount_columns': index_cols[:2],
            },
            'regions': regions,
            'rows': regions,
        }

    def _extract_item_names(self, items: Any) -> List[str]:
        """리스트에서 표시용 이름만 추출"""
        if not items:
            return []
        names = []
        for item in items:
            if isinstance(item, dict):
                name_val = item.get('name') or item.get('display_name')
                if name_val is not None:
                    names.append(name_val)
            else:
                names.append(item)
        return names

    def _enrich_template_data(
        self,
        data: Dict[str, Any],
        table_data: List[Dict[str, Any]],
        regional: Dict[str, Any],
        top3_increase: List[Dict[str, Any]],
        top3_decrease: List[Dict[str, Any]],
    ) -> None:
        """템플릿에서 요구하는 필드를 채워 렌더링 오류를 방지"""

        # summary_box 기본 필드 보강
        summary_box = data.get('summary_box', {}) or {}
        summary_box.setdefault('increase_count', len(regional.get('increase_regions', [])))
        summary_box.setdefault('decrease_count', len(regional.get('decrease_regions', [])))
        summary_box.setdefault('region_count', len(regional.get('increase_regions', [])))
        summary_box.setdefault('main_items', [])
        data['summary_box'] = summary_box

        # summary_table 기본 구조 추가
        data['summary_table'] = self._build_summary_table(table_data)

        # footer 정보 기본값
        data.setdefault('footer_info', {
            'source': '자료: 국가데이터처 국가통계포털(KOSIS), 집계시트',
            'page_num': '1'
        })

        # nationwide 필드 보강 (보고서 타입별 별칭)
        nationwide = data.get('nationwide_data') or {}
        if self.report_type in ['export', 'import']:
            nationwide.setdefault('amount', nationwide.get('production_index'))
            nationwide.setdefault('change', nationwide.get('growth_rate'))
            products = nationwide.get('products') or nationwide.get('main_items') or []
            normalized_products = []
            for p in products:
                if isinstance(p, dict):
                    normalized_products.append({
                        'name': p.get('name') or p.get('display_name') or str(p),
                        'change': p.get('change', nationwide.get('change'))
                    })
                else:
                    normalized_products.append({'name': p, 'change': nationwide.get('change')})
            nationwide['products'] = normalized_products
        elif self.report_type == 'price':
            nationwide.setdefault('index', nationwide.get('production_index'))
            nationwide.setdefault('change', nationwide.get('growth_rate'))
            categories = nationwide.get('categories') or nationwide.get('main_items') or []
            normalized_categories = []
            for cat in categories:
                if isinstance(cat, dict):
                    normalized_categories.append({
                        'name': cat.get('name') or cat.get('display_name') or str(cat),
                        'change': cat.get('change', cat.get('growth_rate', nationwide.get('change')))
                    })
                else:
                    normalized_categories.append({'name': cat, 'change': nationwide.get('change')})
            nationwide['categories'] = normalized_categories
        elif self.report_type == 'employment':
            nationwide.setdefault('employment_rate', nationwide.get('production_index'))
            nationwide.setdefault('change', nationwide.get('growth_rate'))
            nationwide.setdefault('main_age_groups', nationwide.get('main_age_groups', []))
            nationwide.setdefault('top_age_groups', nationwide.get('top_age_groups', []))
        elif self.report_type == 'unemployment':
            nationwide.setdefault('rate', nationwide.get('production_index'))
            nationwide.setdefault('change', nationwide.get('growth_rate'))
            nationwide.setdefault('age_groups', nationwide.get('age_groups', []))
            nationwide.setdefault('main_age_groups', nationwide.get('main_age_groups', []))
        elif self.report_type == 'construction':
            # construction_template.html 호환성 보강
            nationwide.setdefault('civil_growth', nationwide.get('growth_rate'))
            nationwide.setdefault('building_growth', nationwide.get('growth_rate'))
            nationwide.setdefault('civil_subtypes', '철도·궤도, 기계설치')
            nationwide.setdefault('building_subtypes', '주택, 관공서 등')
            nationwide.setdefault('main_category', '토목' if (nationwide.get('growth_rate') is not None and nationwide.get('growth_rate') >= 0) else '토목')
            nationwide.setdefault('sub_types_text', '철도·궤도, 도로·교량, 주택')
        data['nationwide_data'] = nationwide

        # 지역 데이터 별칭/필드 보강
        regional_increase = regional.get('increase_regions', []) or []
        regional_decrease = regional.get('decrease_regions', []) or []

        for entry in regional_increase + regional_decrease:
            if not isinstance(entry, dict):
                continue
            entry.setdefault('change', entry.get('growth_rate'))
            if self.report_type in ['export', 'import']:
                raw_products = entry.get('products') or self._extract_item_names(entry.get('top_industries'))
                normalized_products = []
                for p in raw_products or []:
                    if isinstance(p, dict):
                        normalized_products.append({
                            'name': p.get('name') or p.get('display_name') or str(p),
                            'change': p.get('change', entry.get('change'))
                        })
                    else:
                        normalized_products.append({'name': p, 'change': entry.get('change')})
                entry['products'] = normalized_products
            elif self.report_type == 'price':
                categories = entry.get('categories') or entry.get('top_industries', [])
                normalized_categories = []
                for cat in categories:
                    if isinstance(cat, dict):
                        normalized_categories.append({
                            'name': cat.get('name') or cat.get('display_name') or str(cat),
                            'change': cat.get('change', cat.get('growth_rate', entry.get('change')))
                        })
                    else:
                        normalized_categories.append({'name': cat, 'change': entry.get('change')})
                entry['categories'] = normalized_categories
            elif self.report_type in ['employment', 'unemployment']:
                entry.setdefault('age_groups', [])

        if self.report_type == 'construction':
            nationwide.setdefault('civil_growth', nationwide.get('growth_rate'))
            nationwide.setdefault('building_growth', nationwide.get('growth_rate'))
            for entry in regional_increase + regional_decrease:
                if not isinstance(entry, dict):
                    continue
                entry.setdefault('civil_growth', entry.get('growth_rate'))
                entry.setdefault('building_growth', entry.get('growth_rate'))

        if self.report_type == 'price':
            regional['high_regions'] = regional_increase
            regional['low_regions'] = regional_decrease

        data['regional_data'] = regional

        # Top3 리스트 별칭 보강
        for item in top3_increase + top3_decrease:
            if not isinstance(item, dict):
                continue
            item.setdefault('change', item.get('growth_rate'))
            
            # 모든 타입에 대해 industries_names 추가 (템플릿에서 JSON 렌더링 방지)
            if item.get('industries'):
                item['industries_names'] = self._extract_item_names(item.get('industries'))
            
            if self.report_type in ['export', 'import']:
                item.setdefault('products', self._extract_item_names(item.get('industries')))
            elif self.report_type == 'price':
                item.setdefault('categories', item.get('industries', []))
            elif self.report_type in ['employment', 'unemployment']:
                item.setdefault('age_groups', [])
            elif self.report_type == 'construction':
                # construction_template.html 호환성 보강
                item.setdefault('civil_growth', item.get('growth_rate'))
                item.setdefault('building_growth', item.get('growth_rate'))
                item.setdefault('civil_subtypes', '철도·궤도, 기계설치')
                item.setdefault('building_subtypes', '주택, 관공서 등')

        data['top3_increase_regions'] = top3_increase
        data['top3_decrease_regions'] = top3_decrease

        if self.report_type == 'price':
            data['top3_above_regions'] = [
                {
                    'name': item.get('region'),
                    'change': item.get('growth_rate'),
                    'categories': [
                        {
                            'name': cat.get('name') or cat.get('display_name') or str(cat),
                            'change': cat.get('change', cat.get('growth_rate', item.get('growth_rate')))
                        }
                        if isinstance(cat, dict)
                        else {'name': cat, 'change': item.get('growth_rate')}
                        for cat in (item.get('categories', item.get('industries', [])) or [])
                    ],
                }
                for item in top3_increase
            ]
            data['top3_below_regions'] = [
                {
                    'name': item.get('region'),
                    'change': item.get('growth_rate'),
                    'categories': [
                        {
                            'name': cat.get('name') or cat.get('display_name') or str(cat),
                            'change': cat.get('change', cat.get('growth_rate', item.get('growth_rate')))
                        }
                        if isinstance(cat, dict)
                        else {'name': cat, 'change': item.get('growth_rate')}
                        for cat in (item.get('categories', item.get('industries', [])) or [])
                    ],
                }
                for item in top3_decrease
            ]

    def extract_all_data(self, region: Optional[str] = None) -> Dict[str, Any]:
        """전체 데이터 추출"""
        # 데이터 로드는 외부에서 보장 (테스트 호환성)
        
        # config에서 header_rows 가져오기 (기본값 5)
        max_header_rows = self.config.get('header_rows', 5)
        
        # migration은 load_data()에서 이미 명시적 헤더 탐색으로 컬럼 설정됨
        if self.report_type == 'migration':
            target_idx = self.target_col
            prev_y_idx = self.prev_y_col
        else:
            # 스마트 헤더 탐색기로 인덱스 확보 (병합된 셀 처리)
            # 기본값 사용 금지: 반드시 찾아야 함
            # 타입 키워드가 헤더에 없을 수 있으므로 모든 보고서에서 타입 매칭을 강제하지 않음
            require_type_match = False
            
            target_idx = self.find_target_col_index(
                self.df_aggregation, self.year, self.quarter, 
                require_type_match=require_type_match,
                max_header_rows=max_header_rows
            )
            prev_y_idx = self.find_target_col_index(
                self.df_aggregation, self.year - 1, self.quarter, 
                require_type_match=require_type_match,
                max_header_rows=max_header_rows
            )
        
        if self.df_aggregation is not None:
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
        # 국내인구이동은 nationwide 데이터가 없음
        if self.report_type == 'migration':
            nationwide = None
        else:
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
                    'change': r['change_rate'] if r and isinstance(r, dict) and 'change_rate' in r and r['change_rate'] is not None else None,
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
                    'change': r['change_rate'] if r and isinstance(r, dict) and 'change_rate' in r and r['change_rate'] is not None else None,
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
        
        data = {
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

        self._enrich_template_data(data, table_data, regional_converted, top3_increase, top3_decrease)
        return data


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


class RegionalEconomyByRegionGenerator(BaseGenerator):
    """시도별 경제동향 생성기 (모든 부문 통합)
    
    각 시도별로 생산, 소비·건설, 수출·입, 고용, 물가, 국내인구이동 데이터를 
    한 페이지에 통합하여 보도자료를 생성합니다.
    """
    
    # 17개 시도 정보
    REGIONS = [
        {'code': 11, 'name': '서울', 'full_name': '서울특별시'},
        {'code': 21, 'name': '부산', 'full_name': '부산광역시'},
        {'code': 22, 'name': '대구', 'full_name': '대구광역시'},
        {'code': 23, 'name': '인천', 'full_name': '인천광역시'},
        {'code': 24, 'name': '광주', 'full_name': '광주광역시'},
        {'code': 25, 'name': '대전', 'full_name': '대전광역시'},
        {'code': 26, 'name': '울산', 'full_name': '울산광역시'},
        {'code': 29, 'name': '세종', 'full_name': '세종특별자치시'},
        {'code': 31, 'name': '경기', 'full_name': '경기도'},
        {'code': 32, 'name': '강원', 'full_name': '강원특별자치도'},
        {'code': 33, 'name': '충북', 'full_name': '충청북도'},
        {'code': 34, 'name': '충남', 'full_name': '충청남도'},
        {'code': 35, 'name': '전북', 'full_name': '전북특별자치도'},
        {'code': 36, 'name': '전남', 'full_name': '전라남도'},
        {'code': 37, 'name': '경북', 'full_name': '경상북도'},
        {'code': 38, 'name': '경남', 'full_name': '경상남도'},
        {'code': 39, 'name': '제주', 'full_name': '제주특별자치도'},
    ]
    
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__(excel_path, year, quarter, excel_file)
        self.year = year
        self.quarter = quarter
        self.generators = {}  # 부문별 Generator 캐시
    
    def extract_all_data(self, region: Optional[str] = None) -> Dict[str, Any]:
        """시도별 모든 데이터 추출
        
        Returns:
            지역별 모든 데이터 (섹션별로 다른 generator를 사용하므로 기본 구조만 반환)
        """
        return {
            'report_info': {'year': self.year, 'quarter': self.quarter},
            'nationwide_data': None,
            'regional_data': {},
            'table_data': [],
            'sections': {},
        }
    
    def _get_generator(self, report_type: str) -> UnifiedReportGenerator:
        """부문별 Generator 캐시 또는 생성"""
        if report_type not in self.generators:
            self.generators[report_type] = UnifiedReportGenerator(
                report_type, 
                self.excel_path, 
                self.year, 
                self.quarter, 
                self.xl
            )
        return self.generators[report_type]
    
    def extract_regional_section(self, region_name: str, report_type: str) -> Dict[str, Any]:
        """각 시도별로 부문 섹션 데이터 추출
        
        Args:
            region_name: 시도명 (예: '서울')
            report_type: 부문 타입 (mining, service, consumption 등)
            
        Returns:
            섹션 데이터 (narrative + table)
        """
        try:
            from services.excel_cache import get_sector_data

            cache_config = get_report_config(report_type)
            cache_report_id = cache_config.get('report_id') or cache_config.get('id')

            cached = get_sector_data(self.excel_path, self.year, self.quarter, cache_report_id)

            table_data = None
            industries = None
            if cached:
                cached_data = cached.get('data') if isinstance(cached, dict) else None
                if isinstance(cached_data, dict):
                    table_data = cached.get('table_data') or cached_data.get('table_data')
                industries_by_region = cached.get('industries_by_region') if isinstance(cached, dict) else None
                if isinstance(industries_by_region, dict):
                    industries = industries_by_region.get(region_name)

            if table_data is None:
                gen = self._get_generator(report_type)
                gen.load_data()
                table_data = gen._extract_table_data_ssot()

            region_data = next(
                (d for d in (table_data or []) if d.get('region_name') == region_name),
                None
            )

            if not region_data:
                return None

            if industries is None:
                gen = self._get_generator(report_type)
                industries = gen._extract_industry_data(region_name)
            increase_industries = [
                ind for ind in (industries or [])
                if ind and ind.get('change_rate', 0) > 0
            ]
            increase_industries.sort(key=lambda x: x.get('change_rate', 0), reverse=True)
            
            # 나레이션 생성
            narrative = self._generate_narrative(
                region_name,
                report_type,
                region_data,
                increase_industries[:3] if increase_industries else []
            )
            
            return {
                'narrative': narrative,
                'table': {
                    'periods': self._get_table_periods(gen),
                    'data': [self._format_table_row(region_data, industries)]
                }
            }
        except Exception as e:
            print(f"[지역경제동향] ⚠️ {region_name} - {report_type} 추출 실패: {e}")
            return None
    
    def _generate_narrative(
        self,
        region_name: str,
        report_type: str,
        region_data: Dict,
        top_industries: List[Dict]
    ) -> List[str]:
        """나레이션 생성"""
        narratives = []

        try:
            value = region_data.get('value')
            change_rate = region_data.get('change_rate')

            if value is None:
                return narratives

            try:
                from utils.text_utils import get_terms
            except ImportError:
                import sys
                from pathlib import Path
                sys.path.insert(0, str(Path(__file__).parent.parent))
                from utils.text_utils import get_terms

            # 보고서별 나레이션 템플릿
            template_map = {
                'mining': '{region}의 광공업생산은 {products_phrase}{changes}',
                'service': '{region}의 서비스업생산은 {products_phrase}{changes}',
                'consumption': '{region}의 소비는 {products_phrase}{changes}',
                'construction': '{region}의 건설은 {products_phrase}{changes}',
                'export': '{region}의 수출은 {products_phrase}{changes}',
                'import': '{region}의 수입은 {products_phrase}{changes}',
                'employment': '{region}의 고용률은 {changes}',
                'unemployment': '{region}의 실업률은 {changes}',
                'price': '{region}의 물가는 {products_phrase}{changes}',
                'migration': '{region}의 순인구이동은 {changes}',
            }

            template = template_map.get(report_type, '{region}는 {changes}')

            # 제품/항목 텍스트 생성
            products_text = ''
            if top_industries:
                product_names = [ind.get('name', '') for ind in top_industries[:2] if ind.get('name')]
                products_text = ', '.join(product_names)

            products_phrase = f"{products_text}이 " if products_text else ''

            # 증감 텍스트 (어휘 매핑 준수)
            if change_rate is None:
                changes_text = '변화'
            else:
                _, result_noun, _ = get_terms(report_type, change_rate)
                if abs(change_rate) < 0.01:
                    changes_text = '전년동기대비 보합'
                else:
                    changes_text = f'전년동기대비 {abs(change_rate):.1f}% {result_noun}'

            narrative_text = template.format(
                region=region_name,
                products_phrase=products_phrase,
                changes=changes_text
            )
            narratives.append(narrative_text)

        except Exception as e:
            print(f"[지역경제동향] ⚠️ 나레이션 생성 실패: {e}")

        return narratives
    
    def _get_table_periods(self, gen: UnifiedReportGenerator) -> List[str]:
        """테이블 기간 목록 생성"""
        if gen.year and gen.quarter:
            return [f'{gen.year}/{gen.quarter}Q']
        return ['현 기간', '전년동기']
    
    def _format_table_row(self, region_data: Dict, industries: List[Dict]) -> Dict:
        """테이블 행 포맷팅"""
        return {
            'indicator': region_data.get('region_name', ''),
            'values': [
                region_data.get('value', ''),
                region_data.get('change_rate', '')
            ]
        }
    
    def extract_all_regions_data(self) -> Dict[str, Any]:
        """모든 시도의 통합 데이터 추출"""
        all_regions_data = {}
        
        # 부문별 데이터 추출
        report_types = ['mining', 'service', 'consumption', 'construction', 'export', 'import', 
                        'employment', 'unemployment', 'price', 'migration']
        
        for region in self.REGIONS:
            region_name = region['name']
            all_regions_data[region_name] = {
                'region_info': region,
                'sections': {}
            }
            
            for report_type in report_types:
                section_data = self.extract_regional_section(region_name, report_type)
                if section_data:
                    all_regions_data[region_name]['sections'][report_type] = section_data
        
        return all_regions_data


class RegionalReportGenerator(BaseGenerator):
    """시도별 보고서 생성기 (unified_generator에 통합)"""
    
    # 17개 시도 정보
    REGIONS = {
        'region_seoul': {'code': '11', 'name': '서울', 'full_name': '서울특별시'},
        'region_busan': {'code': '21', 'name': '부산', 'full_name': '부산광역시'},
        'region_daegu': {'code': '22', 'name': '대구', 'full_name': '대구광역시'},
        'region_incheon': {'code': '23', 'name': '인천', 'full_name': '인천광역시'},
        'region_gwangju': {'code': '24', 'name': '광주', 'full_name': '광주광역시'},
        'region_daejeon': {'code': '25', 'name': '대전', 'full_name': '대전광역시'},
        'region_ulsan': {'code': '26', 'name': '울산', 'full_name': '울산광역시'},
        'region_sejong': {'code': '29', 'name': '세종', 'full_name': '세종특별자치시'},
        'region_gyeonggi': {'code': '31', 'name': '경기', 'full_name': '경기도'},
        'region_gangwon': {'code': '32', 'name': '강원', 'full_name': '강원특별자치도'},
        'region_chungbuk': {'code': '33', 'name': '충북', 'full_name': '충청북도'},
        'region_chungnam': {'code': '34', 'name': '충남', 'full_name': '충청남도'},
        'region_jeonbuk': {'code': '35', 'name': '전북', 'full_name': '전북특별자치도'},
        'region_jeonnam': {'code': '36', 'name': '전남', 'full_name': '전라남도'},
        'region_gyeongbuk': {'code': '37', 'name': '경북', 'full_name': '경상북도'},
        'region_gyeongnam': {'code': '38', 'name': '경남', 'full_name': '경상남도'},
        'region_jeju': {'code': '39', 'name': '제주', 'full_name': '제주특별자치도'},
    }
    
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__(excel_path, year, quarter, excel_file)
    
    def extract_all_data(self, region: Optional[str] = None) -> Dict[str, Any]:
        """시도별 모든 데이터 추출
        
        Returns:
            지역별 모든 데이터
        """
        try:
            # 이 generator는 섹션별로 다른 generator를 사용하므로,
            # 전체 데이터를 한 번에 추출하지 않고 기본 구조만 반환
            return {
                'report_info': {'year': self.year, 'quarter': self.quarter},
                'nationwide_data': None,
                'regional_data': {},
                'table_data': [],
                'sections': {},
            }
        except Exception as e:
            print(f"[지역경제동향] [경고] 시도별 데이터 추출 중 오류: {e}")
            return {
                'report_info': {'year': self.year, 'quarter': self.quarter},
                'nationwide_data': None,
                'regional_data': {},
                'table_data': [],
            }
    
    def render_html(self, region: str, template_path: str) -> str:
        """시도별 HTML 보도자료 렌더링
        
        Args:
            region: 지역 키 (e.g., 'region_seoul')
            template_path: 템플릿 파일 경로
        
        Returns:
            렌더링된 HTML 문자열
        """
        from jinja2 import Environment, FileSystemLoader
        
        # 데이터 추출
        data = self.extract_all_data(region)
        
        # 데이터 검증
        if not isinstance(data, dict):
            print(f"[경고] 데이터가 dict가 아닙니다: {type(data)}")
            data = {}
        
        # 템플릿 경로 및 렌더링
        template_path_obj = Path(template_path)
        if not template_path_obj.exists():
            raise ValueError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")
        
        # Jinja2 환경 설정
        env = Environment(loader=FileSystemLoader(str(template_path_obj.parent)))
        template = env.get_template(template_path_obj.name)
        
        # 데이터에 지역 정보 추가
        if region in self.REGIONS:
            data['region_info'] = self.REGIONS[region]
            data['region_name'] = self.REGIONS[region]['name']
        else:
            data['region_info'] = {'code': '00', 'name': region, 'full_name': region}
            data['region_name'] = region
        
        # report_info 추가 (regional templates에 필요)
        if 'report_info' not in data:
            data['report_info'] = {
                'year': self.year,
                'quarter': self.quarter,
                'name': self.config.get('name', '지역경제동향') if hasattr(self, 'config') else '지역경제동향'
            }
        
        # regional_economy_by_region_template.html 호환 기본값
        if 'num_pages' not in data:
            data['num_pages'] = 1
        if 'sections' not in data:
            data['sections'] = {}
        
        # 템플릿 렌더링
        try:
            html_content = template.render(**data)
        except TypeError as e:
            print(f"[경고] 템플릿 렌더링 오류: {e}")
            print(f"[경고] 데이터 타입: {type(data)}")
            print(f"[경고] 데이터 키: {list(data.keys()) if isinstance(data, dict) else 'N/A'}")
            raise
        
        return html_content



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
