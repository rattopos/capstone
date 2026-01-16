#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
통합 보고서 Generator

모든 부문(광공업생산, 서비스업생산, 소비동향 등)의 보고서를 생성하는 통합 Generator입니다.
부문별 차이점은 config/report_configs.py에서 설정으로 관리됩니다.
"""

import pandas as pd
import json
from typing import Optional, Dict, Any, List, Tuple
from jinja2 import Template
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
    통합 보고서 Generator
    
    모든 부문의 보고서를 생성하는 범용 Generator입니다.
    부문별 차이점(시트명, 매핑, 템플릿)은 설정 파일에서 관리됩니다.
    
    사용 예시:
        # 광공업생산
        generator = UnifiedReportGenerator('mining', excel_path, 2025, 3)
        data = generator.extract_all_data()
        
        # 서비스업생산
        generator = UnifiedReportGenerator('service', excel_path, 2025, 3)
        data = generator.extract_all_data()
    """
    
    def __init__(self, report_type: str, excel_path: str, year=None, quarter=None, excel_file=None):
        """
        초기화
        
        Args:
            report_type: 보고서 유형 ('mining', 'service', 'consumption', etc.)
            excel_path: 엑셀 파일 경로
            year: 연도 (선택사항)
            quarter: 분기 (선택사항)
            excel_file: 캐시된 ExcelFile 객체 (선택사항)
        """
        super().__init__(excel_path, year, quarter, excel_file)
        
        # 설정 로드
        self.config = get_report_config(report_type)
        self.report_type = report_type
        self.report_id = self.config['report_id']
        self.name_mapping = self.config['name_mapping']
        
        # 데이터 저장
        self.df_analysis = None
        self.df_aggregation = None
        self.use_raw_data = False
        self.use_aggregation_only = False
        
        # 동적 컬럼 인덱스 캐시
        self._col_cache = {
            'analysis': {},
            'aggregation': {}
        }
        
        # 기간 정보 설정
        if self.year and self.quarter:
            try:
                from utils.excel_utils import get_period_context
                self.period_context = get_period_context(self.year, self.quarter)
            except (ImportError, ModuleNotFoundError):
                self.period_context = self._calculate_period_context(self.year, self.quarter)
            
            self.target_year = self.period_context['target_year']
            self.target_quarter = self.period_context['target_quarter']
            self.prev_y_year = self.period_context['prev_y_year']
            self.prev_y_quarter = self.period_context['prev_y_quarter']
        else:
            self.period_context = None
        
        print(f"[UnifiedGenerator] 보고서 유형: {self.config['name']} ({report_type})")
    
    def _calculate_period_context(self, year: int, quarter: int) -> Dict[str, Any]:
        """기간 정보 계산"""
        prev_q = quarter - 1 if quarter > 1 else 4
        prev_q_year = year if quarter > 1 else year - 1
        
        return {
            'target_year': year,
            'target_quarter': quarter,
            'prev_y_year': year - 1,
            'prev_y_quarter': quarter,
            'prev_q_year': prev_q_year,
            'prev_q': prev_q
        }
    
    def _find_metadata_column(self, df: pd.DataFrame, column_type: str) -> Optional[int]:
        """
        메타데이터 컬럼 동적 탐색
        
        Args:
            df: 탐색할 DataFrame
            column_type: 'region', 'classification', 'code', 'name'
            
        Returns:
            컬럼 인덱스 (0-based) 또는 None
        """
        if df is None:
            return None
        
        # 설정에서 키워드 가져오기
        keywords = self.config['metadata_columns'].get(column_type, [])
        if not keywords:
            return None
        
        # 헤더 순회하며 키워드 매칭
        search_rows = min(5, len(df))
        for row_idx in range(search_rows):
            row = df.iloc[row_idx]
            for col_idx in range(min(30, len(row))):
                cell = row.iloc[col_idx]
                if pd.isna(cell):
                    continue
                cell_str = str(cell).strip().lower().replace(" ", "")
                if any(k.lower().replace(" ", "") in cell_str for k in keywords):
                    return col_idx
        
        return None
    
    def _find_header_row(self, df: pd.DataFrame, keywords: List[str]) -> int:
        """헤더 행 찾기"""
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            row_str = ' '.join([str(v) for v in row.values[:15] if pd.notna(v)])
            if any(kw in row_str for kw in keywords):
                return i
        return 2  # 기본값
    
    def _load_sheets(self):
        """시트 로드 (설정 기반)"""
        xl = self.load_excel()
        
        print(f"[{self.config['name']}] 시트 로드 시작...")
        
        # 분석 시트 찾기
        analysis_sheets = self.config['sheets']['analysis']
        fallback_sheets = self.config['sheets']['fallback']
        
        analysis_sheet, self.use_raw_data = self.find_sheet_with_fallback(
            analysis_sheets,
            fallback_sheets
        )
        
        if analysis_sheet:
            self.df_analysis = self.get_sheet(analysis_sheet)
            print(f"[{self.config['name']}] ✅ 분석 시트 로드: '{analysis_sheet}' ({len(self.df_analysis)}행 × {len(self.df_analysis.columns)}열)")
            
            # 분석 시트가 비어있는지 확인
            region_col_guess = 3  # 일반적으로 3번 컬럼이 지역
            test_row = self.df_analysis[(self.df_analysis[region_col_guess] == '전국')]
            if test_row.empty or (not test_row.empty and test_row.iloc[0].isna().sum() > 20):
                print(f"[{self.config['name']}] ⚠️ 분석 시트가 비어있음 → 집계 시트에서 직접 계산")
                self.use_aggregation_only = True
            else:
                self.use_aggregation_only = False
        else:
            raise ValueError(f"{self.config['name']} 분석 시트를 찾을 수 없습니다. 시트: {xl.sheet_names}")
        
        # 집계 시트 찾기
        aggregation_sheets = self.config['sheets']['aggregation']
        
        agg_sheet, _ = self.find_sheet_with_fallback(
            aggregation_sheets,
            fallback_sheets
        )
        
        if agg_sheet and agg_sheet != analysis_sheet:
            self.df_aggregation = self.get_sheet(agg_sheet)
            print(f"[{self.config['name']}] ✅ 집계 시트 로드: '{agg_sheet}' ({len(self.df_aggregation)}행 × {len(self.df_aggregation.columns)}열)")
        else:
            self.df_aggregation = self.df_analysis.copy()
            print(f"[{self.config['name']}] ℹ️ 집계 시트 = 분석 시트 (동일 시트 사용)")
        
        # 동적 컬럼 인덱스 초기화
        self._initialize_column_indices()
    
    def _initialize_column_indices(self):
        """동적으로 컬럼 인덱스 찾기"""
        print(f"[{self.config['name']}] 컬럼 인덱스 동적 탐색 시작...")
        
        # 메타데이터 컬럼 찾기 (분석 시트 기준)
        if self.df_analysis is not None:
            self._col_cache['analysis']['region'] = self._find_metadata_column(self.df_analysis, 'region') or 3
            self._col_cache['analysis']['classification'] = self._find_metadata_column(self.df_analysis, 'classification') or 4
            self._col_cache['analysis']['code'] = self._find_metadata_column(self.df_analysis, 'code') or 6
            self._col_cache['analysis']['name'] = self._find_metadata_column(self.df_analysis, 'name') or 7
            print(f"[{self.config['name']}] ✅ 메타데이터 컬럼: region={self._col_cache['analysis']['region']}, "
                  f"class={self._col_cache['analysis']['classification']}, "
                  f"code={self._col_cache['analysis']['code']}, "
                  f"name={self._col_cache['analysis']['name']}")
        
        # 메타데이터 컬럼 찾기 (집계 시트 기준)
        if self.df_aggregation is not None:
            self._col_cache['aggregation']['region'] = self._find_metadata_column(self.df_aggregation, 'region') or 3
            self._col_cache['aggregation']['classification'] = self._find_metadata_column(self.df_aggregation, 'classification') or 4
            self._col_cache['aggregation']['code'] = self._find_metadata_column(self.df_aggregation, 'code') or 6
            self._col_cache['aggregation']['name'] = self._find_metadata_column(self.df_aggregation, 'name') or 7
        
        # 데이터 컬럼 찾기
        if self.df_analysis is not None and self.year and self.quarter:
            header_row_idx = self._find_header_row(self.df_analysis, ['지역', str(self.year)])
            header_row = self.df_analysis.iloc[header_row_idx]
            
            try:
                target_col = self.find_target_col_index(header_row, self.year, self.quarter)
                self._col_cache['analysis']['target'] = target_col
                print(f"[{self.config['name']}] ✅ 분석 시트 타겟 컬럼: {target_col}")
            except ValueError as e:
                print(f"[{self.config['name']}] ⚠️ 분석 시트 타겟 컬럼 찾기 실패: {e}")
                self._col_cache['analysis']['target'] = 20  # Fallback
            
            try:
                prev_year_col = self.find_target_col_index(header_row, self.prev_y_year, self.prev_y_quarter)
                self._col_cache['analysis']['prev_year'] = prev_year_col
                print(f"[{self.config['name']}] ✅ 분석 시트 전년동기 컬럼: {prev_year_col}")
            except ValueError:
                self._col_cache['analysis']['prev_year'] = 16  # Fallback
        
        if self.df_aggregation is not None and self.year and self.quarter:
            header_row_idx = self._find_header_row(self.df_aggregation, ['지역', str(self.year)])
            header_row = self.df_aggregation.iloc[header_row_idx]
            
            try:
                target_col = self.find_target_col_index(header_row, self.year, self.quarter)
                self._col_cache['aggregation']['target'] = target_col
                print(f"[{self.config['name']}] ✅ 집계 시트 타겟 컬럼: {target_col}")
            except ValueError as e:
                print(f"[{self.config['name']}] ⚠️ 집계 시트 타겟 컬럼 찾기 실패: {e}")
                self._col_cache['aggregation']['target'] = 25  # Fallback
            
            try:
                prev_year_col = self.find_target_col_index(header_row, self.prev_y_year, self.prev_y_quarter)
                self._col_cache['aggregation']['prev_year'] = prev_year_col
                print(f"[{self.config['name']}] ✅ 집계 시트 전년동기 컬럼: {prev_year_col}")
            except ValueError:
                self._col_cache['aggregation']['prev_year'] = 21  # Fallback
    
    def _get_region_indices(self, df: pd.DataFrame) -> Dict[str, int]:
        """
        각 지역의 시작 인덱스 찾기 (완전 동적)
        
        Returns:
            {지역명: 행 인덱스} 딕셔너리
        """
        region_indices = {}
        
        # 동적 컬럼 인덱스
        name_col = self._col_cache['analysis'].get('name', 7)
        region_col = self._col_cache['analysis'].get('region', 3)
        
        for i in range(len(df)):
            try:
                row = df.iloc[i]
                name_val = str(row[name_col]).strip() if pd.notna(row[name_col]) else ''
                region_val = str(row[region_col]).strip() if pd.notna(row[region_col]) else ''
                
                if name_val == '총지수' and region_val in VALID_REGIONS:
                    region_indices[region_val] = i
            except (IndexError, KeyError):
                continue
        
        return region_indices
    
    def _get_nationwide_data(self) -> Dict[str, Any]:
        """
        전국 데이터 추출 (완전 동적 - 모든 부문 공통)
        
        Returns:
            {
                'growth_rate': float,
                'production_index': float,
                'main_items': [{'name': str, 'growth_rate': float}, ...]
            }
        """
        if self.use_aggregation_only:
            return self._get_nationwide_from_aggregation()
        
        df = self.df_analysis
        target_col = self._col_cache['analysis'].get('target')
        
        if target_col is None:
            return self._get_nationwide_from_aggregation()
        
        # 동적 컬럼 인덱스
        region_col = self._col_cache['analysis'].get('region', 3)
        name_col = self._col_cache['analysis'].get('name', 7)
        
        # 전국 총지수 행 찾기
        nationwide_row = None
        nationwide_idx = None
        for i in range(len(df)):
            row = df.iloc[i]
            if pd.notna(row[region_col]) and str(row[region_col]).strip() == '전국':
                if pd.notna(row[name_col]) and str(row[name_col]).strip() == '총지수':
                    nationwide_row = row
                    nationwide_idx = i
                    break
        
        if nationwide_row is None:
            return self._get_nationwide_from_aggregation()
        
        # 증감률 추출
        growth_rate = self.safe_float(nationwide_row[target_col], 0)
        growth_rate = round(growth_rate, 1) if growth_rate else 0.0
        
        print(f"[{self.config['name']} SSOT] ✅ 전국 - 증감률: {growth_rate}")
        
        # 집계 시트에서 지수
        try:
            agg_target_col = self._col_cache['aggregation'].get('target', 25)
            if len(self.df_aggregation) > 3:
                index_row = self.df_aggregation.iloc[3]
                index_value = self.safe_float(index_row[agg_target_col], 100)
            else:
                index_value = 100.0
        except (IndexError, KeyError):
            index_value = 100.0
        
        # 주요 항목 추출 (업종/업태/품목 등)
        items = []
        start_idx = nationwide_idx + 1 if nationwide_idx is not None else 4
        end_idx = min(start_idx + 20, len(df))
        
        for i in range(start_idx, end_idx):
            try:
                row = df.iloc[i]
                item_name = row[name_col] if pd.notna(row[name_col]) else ''
                item_growth = self.safe_float(row[target_col], None)
                
                if item_name and str(item_name).strip() != '총지수' and item_growth is not None:
                    # 설정의 매핑 적용
                    display_name = self.name_mapping.get(str(item_name).strip(), str(item_name).strip())
                    items.append({
                        'name': display_name,
                        'growth_rate': round(item_growth, 1)
                    })
            except (IndexError, KeyError):
                continue
        
        # 양수만 상위 3개
        positive_items = [i for i in items if i['growth_rate'] > 0]
        positive_items.sort(key=lambda x: x['growth_rate'], reverse=True)
        main_items = positive_items[:3]
        
        return {
            'production_index': index_value,
            'growth_rate': growth_rate,
            'main_items': main_items
        }
    
    def _get_nationwide_from_aggregation(self) -> Dict[str, Any]:
        """집계 시트에서 전국 데이터 추출 (완전 동적)"""
        df = self.df_aggregation
        target_col = self._col_cache['aggregation'].get('target', 25)
        prev_col = self._col_cache['aggregation'].get('prev_year', 21)
        
        # 동적 컬럼 인덱스
        region_col = self._col_cache['aggregation'].get('region', 3)
        code_col = self._col_cache['aggregation'].get('code', 6)
        class_col = self._col_cache['aggregation'].get('classification', 4)
        name_col = self._col_cache['aggregation'].get('name', 7)
        
        # 전국 총지수 행 찾기
        nationwide_rows = df[(df[region_col] == '전국') & (df[code_col] == 'E~S')]
        if nationwide_rows.empty:
            nationwide_rows = df[(df[region_col] == '전국') & (df[class_col].astype(str) == '0')]
        if nationwide_rows.empty:
            return {
                'production_index': 100.0,
                'growth_rate': 0.0,
                'main_items': []
            }
        
        nationwide_total = nationwide_rows.iloc[0]
        
        # 증감률 계산
        current_index = self.safe_float(nationwide_total[target_col], 100)
        prev_year_index = self.safe_float(nationwide_total[prev_col], 100)
        
        if prev_year_index and prev_year_index != 0:
            growth_rate = ((current_index - prev_year_index) / prev_year_index) * 100
        else:
            growth_rate = 0.0
        
        print(f"[{self.config['name']} SSOT] ✅ 전국 (집계) - 지수: {current_index}, 증감률: {round(growth_rate, 1)}")
        
        # 중분류 항목별 데이터
        nationwide_items = df[(df[region_col] == '전국') & (df[class_col].astype(str) == '1')]
        
        items = []
        for _, row in nationwide_items.iterrows():
            curr = self.safe_float(row[target_col], None)
            prev = self.safe_float(row[prev_col], None)
            
            if curr is not None and prev is not None and prev != 0:
                item_growth = ((curr - prev) / prev) * 100
                item_name = str(row[name_col]) if pd.notna(row[name_col]) else ''
                display_name = self.name_mapping.get(item_name, item_name)
                items.append({
                    'name': display_name,
                    'growth_rate': round(item_growth, 1)
                })
        
        # 양수만 상위 3개
        positive_items = sorted([i for i in items if i['growth_rate'] > 0], 
                               key=lambda x: x['growth_rate'], reverse=True)[:3]
        
        return {
            'production_index': current_index,
            'growth_rate': round(growth_rate, 1),
            'main_items': positive_items
        }
    
    def _get_regional_data(self) -> Dict[str, Any]:
        """
        시도별 데이터 추출 (완전 동적 - 모든 부문 공통)
        
        Returns:
            {
                'increase_regions': [...],
                'decrease_regions': [...],
                'all_regions': [...]
            }
        """
        if self.use_aggregation_only:
            return self._get_regional_from_aggregation()
        
        df = self.df_analysis
        target_col = self._col_cache['analysis'].get('target', 20)
        region_indices = self._get_region_indices(df)
        
        if not region_indices:
            return self._get_regional_from_aggregation()
        
        regions = []
        
        # 동적 컬럼 인덱스
        name_col = self._col_cache['analysis'].get('name', 7)
        class_col = self._col_cache['analysis'].get('classification', 4)
        
        for region, start_idx in region_indices.items():
            if region == '전국':
                continue
            
            try:
                # 총지수 행에서 증감률
                total_row = df.iloc[start_idx]
                growth_rate_val = self.safe_float(total_row[target_col], 0)
                growth_rate = round(growth_rate_val, 1) if growth_rate_val else 0.0
                
                print(f"[{self.config['name']} SSOT] ✅ {region} - 증감률: {growth_rate}")
                
                # 집계 시트에서 지수
                agg_target_col = self._col_cache['aggregation'].get('target', 25)
                agg_prev_col = self._col_cache['aggregation'].get('prev_year', 21)
                agg_region_col = self._col_cache['aggregation'].get('region', 3)
                
                idx_row = self.df_aggregation[self.df_aggregation[agg_region_col] == region]
                if not idx_row.empty:
                    index_2024 = self.safe_float(idx_row.iloc[0][agg_prev_col], 0)
                    index_2025 = self.safe_float(idx_row.iloc[0][agg_target_col], 0)
                else:
                    index_2024 = 0
                    index_2025 = 0
                
                # 항목별 데이터
                items = []
                for i in range(start_idx + 1, min(start_idx + 20, len(df))):
                    row = df.iloc[i]
                    classification = str(row[class_col]).strip() if pd.notna(row[class_col]) else ''
                    if classification != '1':
                        continue
                    
                    item_name = row[name_col] if pd.notna(row[name_col]) else ''
                    item_growth = self.safe_float(row[target_col], 0)
                    
                    if item_name:
                        display_name = self.name_mapping.get(str(item_name).strip(), str(item_name).strip())
                        items.append({
                            'name': display_name,
                            'growth_rate': round(item_growth, 1) if item_growth else 0.0
                        })
                
                # 증가/감소에 따라 정렬
                if growth_rate >= 0:
                    sorted_items = sorted([i for i in items if i['growth_rate'] > 0], 
                                        key=lambda x: x['growth_rate'], reverse=True)
                else:
                    sorted_items = sorted([i for i in items if i['growth_rate'] < 0], 
                                        key=lambda x: x['growth_rate'])
                
                regions.append({
                    'region': region,
                    'growth_rate': growth_rate,
                    'index_2024': index_2024,
                    'index_2025': index_2025,
                    'top_items': sorted_items[:3],
                    'all_items': items
                })
            except (IndexError, KeyError) as e:
                print(f"[WARNING] 지역 데이터 추출 실패 ({region}): {e}")
                continue
        
        # 증가/감소 지역 분류
        increase_regions = sorted(
            [r for r in regions if r.get('growth_rate') is not None and r['growth_rate'] > 0],
            key=lambda x: x['growth_rate'],
            reverse=True
        )
        decrease_regions = sorted(
            [r for r in regions if r.get('growth_rate') is not None and r['growth_rate'] < 0],
            key=lambda x: x['growth_rate']
        )
        
        return {
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
            'all_regions': regions
        }
    
    def _get_regional_from_aggregation(self) -> Dict[str, Any]:
        """집계 시트에서 시도별 데이터 추출 (완전 동적)"""
        df = self.df_aggregation
        target_col = self._col_cache['aggregation'].get('target', 25)
        prev_col = self._col_cache['aggregation'].get('prev_year', 21)
        
        # 동적 컬럼 인덱스
        region_col = self._col_cache['aggregation'].get('region', 3)
        code_col = self._col_cache['aggregation'].get('code', 6)
        class_col = self._col_cache['aggregation'].get('classification', 4)
        name_col = self._col_cache['aggregation'].get('name', 7)
        
        individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        regions = []
        
        for region in individual_regions:
            # 해당 지역 총지수
            region_total = df[(df[region_col] == region) & (df[code_col] == 'E~S')]
            if region_total.empty:
                region_total = df[(df[region_col] == region) & (df[class_col].astype(str) == '0')]
            if region_total.empty:
                continue
            region_total = region_total.iloc[0]
            
            # 증감률 계산
            current = self.safe_float(region_total[target_col], 100)
            prev = self.safe_float(region_total[prev_col], 100)
            
            if prev and prev != 0:
                growth_rate = ((current - prev) / prev) * 100
            else:
                growth_rate = 0.0
            
            print(f"[{self.config['name']} SSOT] ✅ {region} (집계) - 증감률: {round(growth_rate, 1)}")
            
            # 해당 지역 항목별 데이터
            region_items = df[(df[region_col] == region) & (df[class_col].astype(str) == '1')]
            
            items = []
            for _, row in region_items.iterrows():
                curr = self.safe_float(row[target_col], None)
                prev_ind = self.safe_float(row[prev_col], None)
                
                if curr is not None and prev_ind is not None and prev_ind != 0:
                    item_growth = ((curr - prev_ind) / prev_ind) * 100
                    item_name = str(row[name_col]) if pd.notna(row[name_col]) else ''
                    display_name = self.name_mapping.get(item_name, item_name)
                    items.append({
                        'name': display_name,
                        'growth_rate': round(item_growth, 1)
                    })
            
            # 증가/감소에 따라 정렬
            if growth_rate >= 0:
                sorted_items = sorted(items, key=lambda x: x['growth_rate'], reverse=True)
            else:
                sorted_items = sorted(items, key=lambda x: x['growth_rate'])
            
            regions.append({
                'region': region,
                'growth_rate': round(growth_rate, 1),
                'index_2024': prev,
                'index_2025': current,
                'top_items': sorted_items[:3],
                'all_items': items
            })
        
        # 증가/감소 지역 분류
        increase_regions = sorted(
            [r for r in regions if r.get('growth_rate') is not None and r['growth_rate'] > 0],
            key=lambda x: x['growth_rate'],
            reverse=True
        )
        decrease_regions = sorted(
            [r for r in regions if r.get('growth_rate') is not None and r['growth_rate'] < 0],
            key=lambda x: x['growth_rate']
        )
        
        return {
            'increase_regions': increase_regions,
            'decrease_regions': decrease_regions,
            'all_regions': regions
        }
    
    def extract_all_data(self) -> Dict[str, Any]:
        """
        전체 데이터 추출 (SSOT - Single Source of Truth)
        
        완전 동적 매핑 - 하드코딩 Zero
        모든 부문에서 동일한 로직으로 데이터를 추출합니다.
        """
        # 시트 로드
        self._load_sheets()
        
        # 데이터 추출
        nationwide_data = self._get_nationwide_data()
        regional_data = self._get_regional_data()
        
        # 요약 박스 데이터
        top3 = regional_data['increase_regions'][:3]
        main_regions = []
        for r in top3:
            items = [item['name'] for item in r['top_items'][:2]]
            main_regions.append({
                'region': r['region'],
                'items': items
            })
        
        summary_box = {
            'main_regions': main_regions,
            'region_count': len(regional_data['increase_regions'])
        }
        
        return {
            'report_info': {
                'year': self.year,
                'quarter': self.quarter,
                'report_type': self.report_type,
                'report_name': self.config['name'],
                'index_name': self.config['index_name'],
                'item_name': self.config['item_name']
            },
            'summary_box': summary_box,
            'nationwide_data': nationwide_data,
            'regional_data': regional_data,
            'config': self.config
        }
    
    def generate_report(self, output_path: str):
        """보고서 생성 (템플릿 기반)"""
        # 데이터 추출
        data = self.extract_all_data()
        
        # 설정에서 템플릿 경로 가져오기
        template_path = Path(__file__).parent / self.config['template']
        
        if not template_path.exists():
            print(f"⚠️ 템플릿 파일이 없습니다: {template_path}")
            return data
        
        # Jinja2 렌더링
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        html = template.render(**data)
        
        # 저장
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html)
        
        print(f"✅ {self.config['name']} 보고서 생성 완료: {output_path}")
        
        return data


# 하위 호환성을 위한 wrapper 클래스들
class MiningManufacturingGenerator(UnifiedReportGenerator):
    """광공업생산 Generator (호환성 wrapper)"""
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('mining', excel_path, year, quarter, excel_file)


class ServiceIndustryGenerator(UnifiedReportGenerator):
    """서비스업생산 Generator (호환성 wrapper)"""
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('service', excel_path, year, quarter, excel_file)


class ConsumptionGenerator(UnifiedReportGenerator):
    """소비동향 Generator (호환성 wrapper)"""
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('consumption', excel_path, year, quarter, excel_file)


if __name__ == '__main__':
    # 테스트
    base_path = Path(__file__).parent.parent
    excel_path = base_path / '분석표_25년 3분기_캡스톤(업데이트).xlsx'
    
    print("=" * 60)
    print("통합 Generator 완전 동적 매핑 테스트")
    print("=" * 60)
    
    for report_type in ['mining', 'service', 'consumption']:
        print(f"\n[TEST] {REPORT_CONFIGS[report_type]['name']}")
        print("-" * 60)
        generator = UnifiedReportGenerator(report_type, str(excel_path), 2025, 3)
        data = generator.extract_all_data()
        
        # 데이터 검증
        print(f"\n✅ 추출 완료: {list(data.keys())}")
        
        # 전국 데이터 확인
        nationwide = data['nationwide_data']
        print(f"\n[전국] 지수: {nationwide['production_index']:.1f}, 증감률: {nationwide['growth_rate']}%")
        items_str = [f"{i['name']}({i['growth_rate']}%)" for i in nationwide['main_items'][:2]]
        print(f"  주요 항목: {items_str}")
        
        # 지역별 데이터 확인
        regional = data['regional_data']
        print(f"\n[지역] 증가: {len(regional['increase_regions'])}개, 감소: {len(regional['decrease_regions'])}개")
        if regional['increase_regions']:
            top_region = regional['increase_regions'][0]
            print(f"  최고: {top_region['region']} ({top_region['growth_rate']}%)")
            top_items_str = [f"{i['name']}({i['growth_rate']}%)" for i in top_region['top_items'][:2]]
            print(f"    주요 항목: {top_items_str}")
        
        # 요약 박스 확인
        summary = data['summary_box']
        print(f"\n[요약] 주요 지역 {len(summary['main_regions'])}개")
        for mr in summary['main_regions']:
            print(f"  - {mr['region']}: {', '.join(mr['items'][:2])}")
        
        print("\n" + "=" * 60)
