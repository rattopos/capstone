#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
서비스업생산 보도자료 생성기

엑셀 데이터를 읽어 스키마에 맞게 데이터를 추출하고,
Jinja2 템플릿을 사용하여 HTML 보도자료를 생성합니다.
"""

import pandas as pd
import json
import re
from typing import Optional, Dict, Any, List, Tuple
from jinja2 import Environment, FileSystemLoader, Template
from pathlib import Path
try:
    from .base_generator import BaseGenerator
except ImportError:
    # 직접 실행 시 상대 import 실패 방지
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).parent))
    from base_generator import BaseGenerator


class ServiceIndustryGenerator(BaseGenerator):
    """서비스업생산 보도자료 생성 클래스"""
    
    # 업종명 매핑
    INDUSTRY_MAPPING = {
        '수도, 하수 및 폐기물 처리, 원료 재생업': '수도·하수',
        '도매 및 소매업': '도소매',
        '운수 및 창고업': '운수·창고',
        '숙박 및 음식점업': '숙박·음식점',
        '정보통신업': '정보통신',
        '금융 및 보험업': '금융·보험',
        '부동산업': '부동산',
        '전문, 과학 및 기술 서비스업': '전문·과학·기술',
        '사업시설관리, 사업지원 및 임대 서비스업': '사업시설관리·사업지원·임대',
        '교육 서비스업': '교육',
        '보건업 및 사회복지 서비스업': '보건·복지',
        '예술, 스포츠 및 여가관련 서비스업': '예술·스포츠·여가',
        '협회 및 단체, 수리  및 기타 개인 서비스업': '협회·수리·개인서비스'
    }
    
    # 지역명 매핑 (표 표시용)
    REGION_DISPLAY_MAPPING = {
        '전국': '전 국',
        '서울': '서 울',
        '부산': '부 산',
        '대구': '대 구',
        '인천': '인 천',
        '광주': '광 주',
        '대전': '대 전',
        '울산': '울 산',
        '세종': '세 종',
        '경기': '경 기',
        '강원': '강 원',
        '충북': '충 북',
        '충남': '충 남',
        '전북': '전 북',
        '전남': '전 남',
        '경북': '경 북',
        '경남': '경 남',
        '제주': '제 주'
    }
    
    # 지역 그룹
    REGION_GROUPS = {
        '경인': ['서울', '인천', '경기'],
        '충청': ['대전', '세종', '충북', '충남'],
        '호남': ['광주', '전북', '전남', '제주'],
        '동북': ['대구', '경북', '강원'],
        '동남': ['부산', '울산', '경남']
    }
    
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        """
        초기화
        
        Args:
            excel_path: 엑셀 파일 경로
            year: 연도 (선택사항)
            quarter: 분기 (선택사항)
            excel_file: 캐시된 ExcelFile 객체 (선택사항)
        """
        # 부모 생성자 호출 필수
        super().__init__(excel_path, year, quarter, excel_file)
        self.report_id = 'service'  # 보고서 ID
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
                # 모듈이 없으면 직접 계산
                self.period_context = self._calculate_period_context(self.year, self.quarter)
            
            self.target_year = self.period_context['target_year']
            self.target_quarter = self.period_context['target_quarter']
            self.prev_y_year = self.period_context['prev_y_year']
            self.prev_y_quarter = self.period_context['prev_y_quarter']
        else:
            self.period_context = None
    
    def _calculate_period_context(self, year: int, quarter: int) -> Dict[str, Any]:
        """기간 정보 계산 (utils 모듈이 없을 때 사용)"""
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
    
    def _find_metadata_column(self, column_type: str) -> Optional[int]:
        """
        메타데이터 컬럼 동적 탐색 (지역명, 분류단계, 업종명 등)
        
        Args:
            column_type: 'region', 'classification', 'industry_code', 'industry_name', 'weight'
            
        Returns:
            컬럼 인덱스 (0-based) 또는 None
        """
        # 분석 시트 또는 집계 시트 중 사용 가능한 것 선택
        df = self.df_analysis if self.df_analysis is not None else self.df_aggregation
        if df is None:
            return None
        
        # 헤더 행 추정 (보통 2~3행)
        search_rows = min(5, len(df))
        
        # 컬럼 타입별 키워드
        keywords_map = {
            'region': ['지역', 'region', '시도'],
            'classification': ['분류단계', 'classification', '단계'],
            'industry_code': ['산업코드', '업태코드', 'code', '코드'],
            'industry_name': ['산업명', '업태명', 'industry', '업종', '산업이름'],
            'weight': ['가중치', 'weight', '비중']
        }
        
        keywords = keywords_map.get(column_type, [])
        if not keywords:
            return None
        
        # 헤더 순회하며 키워드 매칭
        for row_idx in range(search_rows):
            row = df.iloc[row_idx]
            for col_idx in range(min(30, len(row))):  # 처음 30개 컬럼만 검색
                cell = row.iloc[col_idx]
                if pd.isna(cell):
                    continue
                cell_str = str(cell).strip().lower().replace(" ", "")
                if any(k.lower().replace(" ", "") in cell_str for k in keywords):
                    return col_idx
        
        return None
    
    def _load_sheets(self):
        """시트 로드 및 초기 검증"""
        xl = self.load_excel()
        
        print("[ServiceIndustry] 시트 로드 시작...")
        
        # 분석 시트 찾기
        analysis_sheet, self.use_raw_data = self.find_sheet_with_fallback(
            ['B 분석', 'B분석'],
            ['서비스업생산', '서비스업생산지수']
        )
        
        if analysis_sheet:
            self.df_analysis = self.get_sheet(analysis_sheet)
            print(f"[ServiceIndustry] ✅ 분석 시트 로드: '{analysis_sheet}' ({len(self.df_analysis)}행 × {len(self.df_analysis.columns)}열)")
            
            # 분석 시트가 비어있는지 확인
            test_row = self.df_analysis[(self.df_analysis[3] == '전국') | (self.df_analysis[4] == '전국')]
            if test_row.empty or test_row.iloc[0].isna().sum() > 20:
                print(f"[ServiceIndustry] ⚠️ 분석 시트가 비어있음 → 집계 시트에서 직접 계산")
                self.use_aggregation_only = True
            else:
                self.use_aggregation_only = False
        else:
            raise ValueError(f"서비스업생산 분석 시트를 찾을 수 없습니다. 시트: {xl.sheet_names}")
        
        # 집계 시트 찾기
        agg_sheet, _ = self.find_sheet_with_fallback(
            ['B(서비스업생산)집계', 'B 집계'],
            ['서비스업생산', '서비스업생산지수']
        )
        
        if agg_sheet and agg_sheet != analysis_sheet:
            self.df_aggregation = self.get_sheet(agg_sheet)
            print(f"[ServiceIndustry] ✅ 집계 시트 로드: '{agg_sheet}' ({len(self.df_aggregation)}행 × {len(self.df_aggregation.columns)}열)")
        else:
            self.df_aggregation = self.df_analysis.copy()
            print(f"[ServiceIndustry] ℹ️ 집계 시트 = 분석 시트 (동일 시트 사용)")
        
        # 동적 컬럼 인덱스 초기화
        self._initialize_column_indices()
    
    def _find_header_row(self, df: pd.DataFrame, keywords: List[str]) -> int:
        """헤더 행 찾기"""
        # 일반적으로 행 2에 헤더가 있음
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            row_str = ' '.join([str(v) for v in row.values[:15] if pd.notna(v)])
            # keywords 중 하나라도 있으면 헤더로 간주
            if any(kw in row_str for kw in keywords):
                return i
        # 못 찾으면 기본값 2 반환
        return 2
    
    def _initialize_column_indices(self):
        """동적으로 컬럼 인덱스 찾기"""
        print("[ServiceIndustry] 컬럼 인덱스 동적 탐색 시작...")
        
        # 메타데이터 컬럼 찾기 (분석 시트 기준)
        if self.df_analysis is not None:
            self._col_cache['analysis']['region'] = self._find_metadata_column('region') or 3
            self._col_cache['analysis']['classification'] = self._find_metadata_column('classification') or 4
            self._col_cache['analysis']['industry_code'] = self._find_metadata_column('industry_code') or 6
            self._col_cache['analysis']['industry_name'] = self._find_metadata_column('industry_name') or 7
            print(f"[ServiceIndustry] ✅ 메타데이터 컬럼: region={self._col_cache['analysis']['region']}, "
                  f"class={self._col_cache['analysis']['classification']}, "
                  f"code={self._col_cache['analysis']['industry_code']}, "
                  f"name={self._col_cache['analysis']['industry_name']}")
        
        # 메타데이터 컬럼 찾기 (집계 시트 기준)
        if self.df_aggregation is not None:
            self._col_cache['aggregation']['region'] = self._find_metadata_column('region') or 3
            self._col_cache['aggregation']['classification'] = self._find_metadata_column('classification') or 4
            self._col_cache['aggregation']['industry_code'] = self._find_metadata_column('industry_code') or 6
            self._col_cache['aggregation']['industry_name'] = self._find_metadata_column('industry_name') or 7
        
        # 분석 시트 헤더 찾기
        if self.df_analysis is not None and self.year and self.quarter:
            header_row_idx = self._find_header_row(self.df_analysis, ['지역', '산업', '2025', '2024'])
            header_row = self.df_analysis.iloc[header_row_idx]
            
            # 타겟 컬럼 찾기 (현재 분기)
            try:
                target_col = self.find_target_col_index(header_row, self.year, self.quarter)
                self._col_cache['analysis']['target'] = target_col
                print(f"[ServiceIndustry] ✅ 분석 시트 타겟 컬럼: {target_col}")
            except ValueError as e:
                print(f"[ServiceIndustry] ⚠️ 분석 시트 타겟 컬럼 찾기 실패: {e}")
                # Fallback: 하드코딩된 값 사용
                self._col_cache['analysis']['target'] = 20
            
            # 전년동기 컬럼 찾기
            try:
                prev_year_col = self.find_target_col_index(header_row, self.prev_y_year, self.prev_y_quarter)
                self._col_cache['analysis']['prev_year'] = prev_year_col
                print(f"[ServiceIndustry] ✅ 분석 시트 전년동기 컬럼: {prev_year_col}")
            except ValueError:
                self._col_cache['analysis']['prev_year'] = 16
        
        # 집계 시트 헤더 찾기
        if self.df_aggregation is not None and self.year and self.quarter:
            header_row_idx = self._find_header_row(self.df_aggregation, ['지역', '산업', '2025', '2024'])
            header_row = self.df_aggregation.iloc[header_row_idx]
            
            # 타겟 컬럼 찾기 (현재 분기)
            try:
                target_col = self.find_target_col_index(header_row, self.year, self.quarter)
                self._col_cache['aggregation']['target'] = target_col
                print(f"[ServiceIndustry] ✅ 집계 시트 타겟 컬럼: {target_col}")
            except ValueError as e:
                print(f"[ServiceIndustry] ⚠️ 집계 시트 타겟 컬럼 찾기 실패: {e}")
                self._col_cache['aggregation']['target'] = 25
            
            # 전년동기 컬럼 찾기
            try:
                prev_year_col = self.find_target_col_index(header_row, self.prev_y_year, self.prev_y_quarter)
                self._col_cache['aggregation']['prev_year'] = prev_year_col
                print(f"[ServiceIndustry] ✅ 집계 시트 전년동기 컬럼: {prev_year_col}")
            except ValueError:
                self._col_cache['aggregation']['prev_year'] = 21
    
    def _get_region_indices(self, df) -> Dict[str, int]:
        """각 지역의 시작 인덱스 찾기"""
        region_indices = {}
        
        # 시도 목록
        VALID_REGIONS = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                         '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 동적 컬럼 인덱스 가져오기
        name_col = self._col_cache['analysis'].get('industry_name', 7)
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
        """전국 데이터 추출"""
        if self.use_aggregation_only:
            return self._get_nationwide_from_aggregation()
        else:
            return self._get_nationwide_from_analysis()
    
    def _get_nationwide_from_analysis(self) -> Dict[str, Any]:
        """분석 시트에서 전국 데이터 추출"""
        df = self.df_analysis
        target_col = self._col_cache['analysis'].get('target')
        
        if target_col is None:
            # Fallback: 하드코딩된 컬럼 사용 (2025년 2분기 기준)
            print("[ServiceIndustry] ⚠️ 타겟 컬럼을 찾을 수 없어 Fallback 사용")
            target_col = 20
        
        # 전국 총지수 행 찾기 (동적 컬럼 사용)
        region_col = self._col_cache['analysis'].get('region', 3)
        name_col = self._col_cache['analysis'].get('industry_name', 7)
        
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
            # 집계 시트에서 계산
            return self._get_nationwide_from_aggregation()
        
        # 분석 시트에서 증감률 읽기
        growth_rate = self.safe_float(nationwide_row[target_col], 0)
        growth_rate = round(growth_rate, 1) if growth_rate else 0.0
        
        print(f"[ServiceIndustry SSOT] ✅ 전국 - 증감률: {growth_rate}")
        
        # 집계 시트에서 전국 지수
        try:
            agg_target_col = self._col_cache['aggregation'].get('target', 25)
            if len(self.df_aggregation) > 3:
                index_row = self.df_aggregation.iloc[3]
                production_index = self.safe_float(index_row[agg_target_col], 100)
            else:
                production_index = 100.0
        except (IndexError, KeyError):
            production_index = 100.0
        
        # 전국 주요 업종 (증가율 기준 상위 3개 - 양수만)
        industries = []
        start_idx = nationwide_idx + 1 if nationwide_idx is not None else 4
        end_idx = min(start_idx + 13, len(df))
        
        for i in range(start_idx, end_idx):
            try:
                row = df.iloc[i]
                industry_name = row[name_col] if pd.notna(row[name_col]) else ''
                industry_growth = self.safe_float(row[target_col], 0)
                if industry_growth is not None and str(industry_name).strip() != '총지수':
                    industries.append({
                        'name': self.INDUSTRY_MAPPING.get(str(industry_name).strip(), str(industry_name).strip()),
                        'growth_rate': round(industry_growth, 1)
                    })
            except (IndexError, KeyError):
                continue
        
        # 양수 증가율 중 상위 3개
        positive_industries = [i for i in industries if i['growth_rate'] > 0]
        positive_industries.sort(key=lambda x: x['growth_rate'], reverse=True)
        main_industries = positive_industries[:3]
        
        return {
            'production_index': production_index,
            'growth_rate': growth_rate,
            'main_industries': main_industries
        }
    
    def _get_nationwide_from_aggregation(self) -> Dict[str, Any]:
        """집계 시트에서 전국 데이터 추출 (증감률 직접 계산)"""
        df = self.df_aggregation
        target_col = self._col_cache['aggregation'].get('target', 25)
        prev_col = self._col_cache['aggregation'].get('prev_year', 21)
        
        # 동적 컬럼 인덱스
        region_col = self._col_cache['aggregation'].get('region', 3)
        code_col = self._col_cache['aggregation'].get('industry_code', 6)
        class_col = self._col_cache['aggregation'].get('classification', 4)
        name_col = self._col_cache['aggregation'].get('industry_name', 7)
        
        # 전국 총지수 행 찾기
        nationwide_rows = df[(df[region_col] == '전국') & (df[code_col] == 'E~S')]
        if nationwide_rows.empty:
            nationwide_rows = df[(df[region_col] == '전국') & (df[class_col].astype(str) == '0')]
        if nationwide_rows.empty:
            return {
                'production_index': 100.0,
                'growth_rate': 0.0,
                'main_industries': []
            }
        
        nationwide_total = nationwide_rows.iloc[0]
        
        # 증감률 계산
        current_index = self.safe_float(nationwide_total[target_col], 100)
        prev_year_index = self.safe_float(nationwide_total[prev_col], 100)
        
        if prev_year_index and prev_year_index != 0:
            growth_rate = ((current_index - prev_year_index) / prev_year_index) * 100
        else:
            growth_rate = 0.0
        
        print(f"[ServiceIndustry SSOT] ✅ 전국 (집계) - 지수: {current_index}, 전년: {prev_year_index}, 증감률: {round(growth_rate, 1)}")
        
        # 전국 중분류 업종별 데이터
        nationwide_industries = df[(df[region_col] == '전국') & (df[class_col].astype(str) == '1')]
        
        industries = []
        for _, row in nationwide_industries.iterrows():
            curr = self.safe_float(row[target_col], None)
            prev = self.safe_float(row[prev_col], None)
            
            if curr is not None and prev is not None and prev != 0:
                ind_growth = ((curr - prev) / prev) * 100
                industries.append({
                    'name': self.INDUSTRY_MAPPING.get(str(row[name_col]) if pd.notna(row[name_col]) else '', str(row[name_col]) if pd.notna(row[name_col]) else ''),
                    'growth_rate': round(ind_growth, 1)
                })
        
        # 양수 증가율 중 상위 3개
        positive_industries = sorted([i for i in industries if i['growth_rate'] > 0], 
                                    key=lambda x: x['growth_rate'], reverse=True)[:3]
        
        return {
            'production_index': current_index,
            'growth_rate': round(growth_rate, 1),
            'main_industries': positive_industries
        }
    
    def _get_regional_data(self) -> Dict[str, Any]:
        """시도별 데이터 추출"""
        if self.use_aggregation_only:
            return self._get_regional_from_aggregation()
        else:
            return self._get_regional_from_analysis()
    
    def _get_regional_from_analysis(self) -> Dict[str, Any]:
        """분석 시트에서 시도별 데이터 추출"""
        df = self.df_analysis
        target_col = self._col_cache['analysis'].get('target', 20)
        region_indices = self._get_region_indices(df)
        regions = []
        
        if not region_indices:
            return self._get_regional_from_aggregation()
        
        for region, start_idx in region_indices.items():
            if region == '전국':
                continue
            
            try:
                # 총지수 행에서 증감률
                total_row = df.iloc[start_idx]
                growth_rate_val = self.safe_float(total_row[target_col], 0)
                growth_rate = round(growth_rate_val, 1) if growth_rate_val else 0.0
                
                # 집계 시트에서 지수
                agg_target_col = self._col_cache['aggregation'].get('target', 25)
                agg_prev_col = self._col_cache['aggregation'].get('prev_year', 21)
                idx_row = self.df_aggregation[self.df_aggregation[3] == region]
                if not idx_row.empty:
                    index_2024 = self.safe_float(idx_row.iloc[0][agg_prev_col], 0)
                    index_2025 = self.safe_float(idx_row.iloc[0][agg_target_col], 0)
                else:
                    index_2024 = 0
                    index_2025 = 0
                
                print(f"[ServiceIndustry SSOT] ✅ {region} - 증감률: {growth_rate}")
                
                # 동적 컬럼 인덱스
                name_col = self._col_cache['analysis'].get('industry_name', 7)
                class_col = self._col_cache['analysis'].get('classification', 4)
                
                # 업종별 기여도
                industries = []
                contribution_col = target_col + 6  # 일반적으로 증감률 + 6 컬럼이 기여도
                for i in range(start_idx + 1, min(start_idx + 14, len(df))):
                    row = df.iloc[i]
                    classification = str(row[class_col]).strip() if pd.notna(row[class_col]) else ''
                    if classification != '1':
                        continue
                    industry_name = row[name_col] if pd.notna(row[name_col]) else ''
                    industry_growth = self.safe_float(row[target_col], 0)
                    contribution = self.safe_float(row[contribution_col], 0)
                    
                    if contribution is not None:
                        industries.append({
                            'name': self.INDUSTRY_MAPPING.get(str(industry_name).strip(), str(industry_name).strip()),
                            'growth_rate': round(industry_growth, 1) if industry_growth else 0.0,
                            'contribution': contribution
                        })
                
                # 기여도 순 정렬
                if growth_rate >= 0:
                    sorted_ind = sorted([i for i in industries if i['contribution'] > 0], 
                                      key=lambda x: x['contribution'], reverse=True)
                else:
                    sorted_ind = sorted([i for i in industries if i['contribution'] < 0], 
                                      key=lambda x: x['contribution'])
                
                regions.append({
                    'region': region,
                    'growth_rate': growth_rate,
                    'index_2024': index_2024,
                    'index_2025': index_2025,
                    'top_industries': sorted_ind[:3],
                    'all_industries': industries
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
        """집계 시트에서 시도별 데이터 추출"""
        df = self.df_aggregation
        target_col = self._col_cache['aggregation'].get('target', 25)
        prev_col = self._col_cache['aggregation'].get('prev_year', 21)
        
        # 동적 컬럼 인덱스
        region_col = self._col_cache['aggregation'].get('region', 3)
        code_col = self._col_cache['aggregation'].get('industry_code', 6)
        class_col = self._col_cache['aggregation'].get('classification', 4)
        name_col = self._col_cache['aggregation'].get('industry_name', 7)
        
        individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                              '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        # 전국 전년동분기 지수 (기여도 계산용)
        nationwide_rows = df[(df[region_col] == '전국') & (df[code_col] == 'E~S')]
        if nationwide_rows.empty:
            nationwide_rows = df[(df[region_col] == '전국') & (df[class_col].astype(str) == '0')]
        nationwide_prev = self.safe_float(nationwide_rows.iloc[0][prev_col], 100) if not nationwide_rows.empty else 100
        
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
            
            print(f"[ServiceIndustry SSOT] ✅ {region} (집계) - 지수: {current}, 전년: {prev}, 증감률: {round(growth_rate, 1)}")
            
            # 해당 지역 업종별 데이터
            region_industries = df[(df[region_col] == region) & (df[class_col].astype(str) == '1')]
            
            industries = []
            for _, row in region_industries.iterrows():
                curr = self.safe_float(row[target_col], None)
                prev_ind = self.safe_float(row[prev_col], None)
                
                if curr is not None and prev_ind is not None and prev_ind != 0:
                    ind_growth = ((curr - prev_ind) / prev_ind) * 100
                    contribution = (curr - prev_ind) / nationwide_prev * 100 if nationwide_prev else 0
                    industries.append({
                        'name': self.INDUSTRY_MAPPING.get(str(row[name_col]) if pd.notna(row[name_col]) else '', str(row[name_col]) if pd.notna(row[name_col]) else ''),
                        'growth_rate': round(ind_growth, 1),
                        'contribution': round(contribution, 6)
                    })
            
            # 기여도 순 정렬
            if growth_rate >= 0:
                sorted_ind = sorted(industries, key=lambda x: x['contribution'], reverse=True)
            else:
                sorted_ind = sorted(industries, key=lambda x: x['contribution'])
            
            regions.append({
                'region': region,
                'growth_rate': round(growth_rate, 1),
                'index_2024': prev,
                'index_2025': current,
                'top_industries': sorted_ind[:3],
                'all_industries': industries
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
    
    def _get_growth_rates_table(self) -> List[Dict[str, Any]]:
        """표에 들어갈 증감률 및 지수 데이터 생성"""
        if self.use_aggregation_only or len(self._get_region_indices(self.df_analysis)) <= 1:
            return self._get_table_data_from_aggregation()
        else:
            return self._get_table_data_from_analysis()
    
    def _get_table_data_from_analysis(self) -> List[Dict[str, Any]]:
        """분석 시트에서 테이블 데이터 추출"""
        df = self.df_analysis
        df_agg = self.df_aggregation
        region_indices = self._get_region_indices(df)
        table_data = []
        
        # 동적 컬럼 찾기
        target_col = self._col_cache['analysis'].get('target', 20)
        agg_target_col = self._col_cache['aggregation'].get('target', 25)
        agg_prev_col = self._col_cache['aggregation'].get('prev_year', 21)
        
        # [완전 동적] 표에 표시할 과거 분기들의 컬럼도 동적으로 찾기
        header_row = df.iloc[self._find_header_row(df, ['지역', '산업', str(self.year)])]
        
        # 2년 전 동분기 (예: 2025.2 → 2023.2)
        try:
            col_2years_ago = self.find_target_col_index(header_row, self.year - 2, self.quarter)
        except ValueError:
            col_2years_ago = target_col - 8  # Fallback
        
        # 1년 전 동분기 (예: 2025.2 → 2024.2)
        try:
            col_1year_ago = self.find_target_col_index(header_row, self.year - 1, self.quarter)
        except ValueError:
            col_1year_ago = target_col - 4  # Fallback
        
        # 전분기 (예: 2025.2 → 2025.1)
        prev_q = self.quarter - 1 if self.quarter > 1 else 4
        prev_q_year = self.year if self.quarter > 1 else self.year - 1
        try:
            col_prev_quarter = self.find_target_col_index(header_row, prev_q_year, prev_q)
        except ValueError:
            col_prev_quarter = target_col - 1  # Fallback
        
        # 전국 데이터
        nationwide_idx = region_indices.get('전국', 3)
        try:
            nationwide_row = df.iloc[nationwide_idx]
            nationwide_idx_row = df_agg.iloc[nationwide_idx] if len(df_agg) > nationwide_idx else None
            
            table_data.append({
                'group': None,
                'rowspan': None,
                'region': self.REGION_DISPLAY_MAPPING['전국'],
                'growth_rates': [
                    round(self.safe_float(nationwide_row[col_2years_ago], 0), 1),
                    round(self.safe_float(nationwide_row[col_1year_ago], 0), 1),
                    round(self.safe_float(nationwide_row[col_prev_quarter], 0), 1),
                    round(self.safe_float(nationwide_row[target_col], 0), 1),
                ],
                'indices': [
                    self.safe_float(nationwide_idx_row[agg_prev_col], 0) if nationwide_idx_row is not None else 0,
                    self.safe_float(nationwide_idx_row[agg_target_col], 0) if nationwide_idx_row is not None else 0,
                ]
            })
        except (IndexError, KeyError):
            table_data.append({
                'group': None,
                'rowspan': None,
                'region': self.REGION_DISPLAY_MAPPING['전국'],
                'growth_rates': [0.0, 0.0, 0.0, 0.0],
                'indices': [0.0, 0.0]
            })
        
        # 지역별 그룹
        for group_name, group_regions in self.REGION_GROUPS.items():
            for i, region in enumerate(group_regions):
                if region not in region_indices:
                    continue
                
                try:
                    start_idx = region_indices[region]
                    row = df.iloc[start_idx]
                    idx_row = df_agg[df_agg[3] == region]
                    
                    if idx_row.empty:
                        continue
                        
                    idx_row = idx_row.iloc[0]
                    
                    entry = {
                        'region': self.REGION_DISPLAY_MAPPING.get(region, region),
                        'growth_rates': [
                            round(self.safe_float(row[col_2years_ago], 0), 1),
                            round(self.safe_float(row[col_1year_ago], 0), 1),
                            round(self.safe_float(row[col_prev_quarter], 0), 1),
                            round(self.safe_float(row[target_col], 0), 1),
                        ],
                        'indices': [
                            self.safe_float(idx_row[agg_prev_col], 0),
                            self.safe_float(idx_row[agg_target_col], 0),
                        ]
                    }
                    
                    if i == 0:
                        entry['group'] = group_name
                        entry['rowspan'] = len(group_regions)
                    else:
                        entry['group'] = None
                        entry['rowspan'] = None
                        
                    table_data.append(entry)
                except (IndexError, KeyError) as e:
                    print(f"[WARNING] 지역 테이블 데이터 추출 실패 ({region}): {e}")
                    continue
        
        return table_data
    
    def _get_table_data_from_aggregation(self) -> List[Dict[str, Any]]:
        """집계 시트에서 테이블 데이터 추출"""
        df = self.df_aggregation
        target_col = self._col_cache['aggregation'].get('target', 25)
        prev_col = self._col_cache['aggregation'].get('prev_year', 21)
        
        # [완전 동적] 과거 분기 컬럼도 동적으로 찾기
        header_row = df.iloc[self._find_header_row(df, ['지역', '산업', str(self.year)])]
        
        # 2년 전 동분기
        try:
            col_2years_ago = self.find_target_col_index(header_row, self.year - 2, self.quarter)
        except ValueError:
            col_2years_ago = prev_col - 4  # Fallback
        
        # 전분기
        prev_q = self.quarter - 1 if self.quarter > 1 else 4
        prev_q_year = self.year if self.quarter > 1 else self.year - 1
        try:
            col_prev_quarter = self.find_target_col_index(header_row, prev_q_year, prev_q)
        except ValueError:
            col_prev_quarter = target_col - 1  # Fallback
        
        # 전년 전분기 (예: 2025.2 → 2024.1)
        try:
            col_prev_year_prev_q = self.find_target_col_index(header_row, self.year - 1, prev_q)
        except ValueError:
            col_prev_year_prev_q = prev_col - 1  # Fallback
        
        table_data = []
        all_regions = ['전국', '서울', '인천', '경기', '대전', '세종', '충북', '충남', 
                       '광주', '전북', '전남', '제주', '대구', '경북', '강원', '부산', '울산', '경남']
        
        # 동적 컬럼 인덱스
        region_col = self._col_cache['aggregation'].get('region', 3)
        code_col = self._col_cache['aggregation'].get('industry_code', 6)
        class_col = self._col_cache['aggregation'].get('classification', 4)
        
        for region in all_regions:
            # 해당 지역 총지수
            region_total = df[(df[region_col] == region) & (df[code_col] == 'E~S')]
            if region_total.empty:
                region_total = df[(df[region_col] == region) & (df[class_col].astype(str) == '0')]
            if region_total.empty:
                continue
            
            row = region_total.iloc[0]
            
            # 지수 추출 (동적 컬럼 사용)
            idx_current = self.safe_float(row[target_col], 0)
            idx_prev_year = self.safe_float(row[prev_col], 0)
            idx_2years_ago = self.safe_float(row[col_2years_ago], 0)
            idx_prev_quarter = self.safe_float(row[col_prev_quarter], 0)
            idx_prev_year_prev_q = self.safe_float(row[col_prev_year_prev_q], 0)
            
            # 증감률 계산
            growth_2years_ago = ((idx_2years_ago - self.safe_float(row[col_2years_ago - 4], 100)) / self.safe_float(row[col_2years_ago - 4], 100) * 100) if self.safe_float(row[col_2years_ago - 4], 0) != 0 else 0.0
            growth_prev_year = ((idx_prev_year - idx_2years_ago) / idx_2years_ago * 100) if idx_2years_ago and idx_2years_ago != 0 else 0.0
            growth_prev_quarter = ((idx_prev_quarter - idx_prev_year_prev_q) / idx_prev_year_prev_q * 100) if idx_prev_year_prev_q and idx_prev_year_prev_q != 0 else 0.0
            growth_current = ((idx_current - idx_prev_year) / idx_prev_year * 100) if idx_prev_year and idx_prev_year != 0 else 0.0
            
            table_data.append({
                'region': self.REGION_DISPLAY_MAPPING.get(region, region),
                'growth_rates': [
                    round(growth_2years_ago, 1),
                    round(growth_prev_year, 1),
                    round(growth_prev_quarter, 1),
                    round(growth_current, 1)
                ],
                'indices': [
                    round(idx_prev_year, 1),
                    round(idx_current, 1)
                ]
            })
        
        # 그룹 정보 추가
        result_data = []
        
        # 전국 먼저
        nationwide = next((r for r in table_data if r['region'] == '전 국'), None)
        if nationwide:
            nationwide['group'] = None
            nationwide['rowspan'] = None
            result_data.append(nationwide)
        
        # 지역 그룹별로 추가
        for group_name, group_regions in self.REGION_GROUPS.items():
            for i, region in enumerate(group_regions):
                region_data = next((r for r in table_data if r['region'] == self.REGION_DISPLAY_MAPPING.get(region, region)), None)
                if region_data:
                    if i == 0:
                        region_data['group'] = group_name
                        region_data['rowspan'] = len(group_regions)
                    else:
                        region_data['group'] = None
                        region_data['rowspan'] = None
                    result_data.append(region_data)
        
        return result_data
    
    def _get_summary_box_data(self, regional_data: Dict[str, Any]) -> Dict[str, Any]:
        """요약 박스 데이터 생성"""
        top3 = regional_data['increase_regions'][:3]
        
        main_regions = []
        for r in top3:
            industries = [ind['name'] for ind in r['top_industries'][:2]]
            main_regions.append({
                'region': r['region'],
                'industries': industries
            })
        
        return {
            'main_increase_regions': main_regions,
            'region_count': len(regional_data['increase_regions'])
        }
    
    def extract_all_data(self) -> Dict[str, Any]:
        """전체 데이터 추출 (SSOT)"""
        # 시트 로드
        self._load_sheets()
        
        # 데이터 추출
        nationwide_data = self._get_nationwide_data()
        regional_data = self._get_regional_data()
        summary_box = self._get_summary_box_data(regional_data)
        table_data = self._get_growth_rates_table()
        
        # Top 3 증가/감소 지역
        top3_increase = []
        for r in regional_data['increase_regions'][:3]:
            top3_increase.append({
                'region': r['region'],
                'growth_rate': r['growth_rate'],
                'industries': r['top_industries']
            })
        
        top3_decrease = []
        for r in regional_data['decrease_regions'][:3]:
            top3_decrease.append({
                'region': r['region'],
                'growth_rate': r['growth_rate'],
                'industries': r['top_industries']
            })
        
        # 감소/증가 업종 텍스트 생성
        decrease_industries = set()
        for r in regional_data['decrease_regions'][:3]:
            for ind in r['top_industries'][:2]:
                decrease_industries.add(ind['name'])
        decrease_industries_text = ', '.join(list(decrease_industries)[:4])
        
        increase_industries = set()
        for r in regional_data['increase_regions'][:3]:
            for ind in r['top_industries'][:2]:
                increase_industries.add(ind['name'])
        increase_industries_text = ', '.join(list(increase_industries)[:4])
        
        return {
            'report_info': {
                'year': self.year,
                'quarter': self.quarter
            },
            'summary_box': summary_box,
            'nationwide_data': nationwide_data,
            'regional_data': regional_data,
            'top3_increase_regions': top3_increase,
            'top3_decrease_regions': top3_decrease,
            'decrease_industries_text': decrease_industries_text,
            'increase_industries_text': increase_industries_text,
            'summary_table': {
                'base_year': 2020,
                'columns': {
                    'growth_rate_columns': [
                        f'{self.year - 2}.{self.quarter}/4',
                        f'{self.year - 1}.{self.quarter}/4',
                        f'{prev_q_year}.{prev_q}/4',
                        f'{self.year}.{self.quarter}/4p'
                    ],
                    'index_columns': [
                        f'{self.year - 1}.{self.quarter}/4',
                        f'{self.year}.{self.quarter}/4p'
                    ]
                },
                'regions': table_data
            }
        }
    
    def generate_report_data(self, raw_excel_path=None):
        """미리보기용 보도자료 데이터 생성 (하위 호환성)"""
        return self.extract_all_data()
    
    def generate_report(self, template_path: str, output_path: str, raw_excel_path=None):
        """보도자료 생성"""
        # 데이터 추출
        template_data = self.extract_all_data()
        
        # JSON 데이터 저장
        data_path = Path(output_path).parent / 'service_industry_data.json'
        with open(data_path, 'w', encoding='utf-8') as f:
            json.dump(template_data, f, ensure_ascii=False, indent=2, default=str)
        
        # 템플릿 렌더링
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        html_output = template.render(**template_data)
        
        # HTML 저장
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_output)
        
        print(f"보도자료 생성 완료: {output_path}")
        print(f"데이터 파일 저장: {data_path}")
        
        return template_data


# 하위 호환성을 위한 함수들
def load_data(excel_path):
    """하위 호환성: 엑셀 파일에서 데이터 로드"""
    generator = ServiceIndustryGenerator(excel_path)
    generator._load_sheets()
    return generator.df_analysis, generator.df_aggregation


def generate_report_data(excel_path, raw_excel_path=None, year=None, quarter=None):
    """하위 호환성: 미리보기용 보도자료 데이터 생성"""
    generator = ServiceIndustryGenerator(excel_path, year, quarter)
    return generator.extract_all_data()


def generate_report(excel_path, template_path, output_path, raw_excel_path=None, year=None, quarter=None):
    """하위 호환성: 보도자료 생성"""
    generator = ServiceIndustryGenerator(excel_path, year, quarter)
    return generator.generate_report(template_path, output_path, raw_excel_path)


if __name__ == '__main__':
    base_path = Path(__file__).parent.parent
    excel_path = base_path / '분석표_25년 3분기_캡스톤(업데이트).xlsx'
    template_path = Path(__file__).parent / 'service_industry_template.html'
    output_path = Path(__file__).parent / 'service_industry_output.html'
    
    generator = ServiceIndustryGenerator(excel_path, year=2025, quarter=3)
    data = generator.generate_report(template_path, output_path)
    
    # 검증용 출력
    print("\n=== 전국 데이터 ===")
    print(f"생산지수: {data['nationwide_data']['production_index']}")
    print(f"증감률: {data['nationwide_data']['growth_rate']}%")
    print(f"주요 업종: {data['nationwide_data']['main_industries']}")
    
    print("\n=== 증가 지역 Top 3 ===")
    for r in data['top3_increase_regions']:
        print(f"{r['region']}({r['growth_rate']}%): {r['industries']}")
    
    print("\n=== 감소 지역 Top 3 ===")
    for r in data['top3_decrease_regions']:
        print(f"{r['region']}({r['growth_rate']}%): {r['industries']}")
