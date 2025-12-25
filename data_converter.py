#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
기초자료 수집표 → 분석표 변환 모듈
기초자료에서 데이터를 추출하여 분석표 형식으로 변환합니다.
또한 GRDP 데이터도 추출하여 참고_GRDP 보고서용 데이터를 생성합니다.
"""

import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import json


class DataConverter:
    """기초자료 수집표 → 분석표 변환기"""
    
    # 기초자료 시트 → 분석표 시트 매핑
    SHEET_MAPPING = {
        '광공업생산': 'A 분석',
        '서비스업생산': 'B 분석',
        '소비(소매, 추가)': 'C 분석',
        '고용률': 'D(고용률)분석',
        '실업자 수': 'D(실업)분석',
        '품목성질별 물가': 'E(품목성질물가)분석',
        '건설 (공표자료)': "F'분석",
        '수출': 'G 분석',
        '수입': 'H 분석',
        '시도 간 이동': 'I(순인구이동)집계',
    }
    
    # 지역 코드 → 이름 매핑
    REGION_CODES = {
        '00': '전국',
        '11': '서울',
        '21': '부산',
        '22': '대구',
        '23': '인천',
        '24': '광주',
        '25': '대전',
        '26': '울산',
        '29': '세종',
        '31': '경기',
        '32': '강원',
        '33': '충북',
        '34': '충남',
        '35': '전북',
        '36': '전남',
        '37': '경북',
        '38': '경남',
        '39': '제주',
    }
    
    REGIONS_ORDER = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                     '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
    
    def __init__(self, raw_excel_path: str):
        """
        Args:
            raw_excel_path: 기초자료 수집표 엑셀 파일 경로
        """
        self.raw_excel_path = Path(raw_excel_path)
        self.xl = pd.ExcelFile(raw_excel_path)
        self.year = None
        self.quarter = None
        self._detect_year_quarter()
    
    def _detect_year_quarter(self):
        """파일명 또는 데이터에서 연도/분기 추출"""
        filename = self.raw_excel_path.stem
        
        # 파일명에서 추출 시도
        if '2025년' in filename and '2분기' in filename:
            self.year, self.quarter = 2025, 2
        elif '25년' in filename and '2분기' in filename:
            self.year, self.quarter = 2025, 2
        elif '2025년' in filename and '1분기' in filename:
            self.year, self.quarter = 2025, 1
        else:
            # 기본값
            self.year, self.quarter = 2025, 2
    
    def convert_all(self, output_path: str = None) -> str:
        """모든 시트를 분석표 형식으로 변환
        
        Args:
            output_path: 출력 파일 경로 (None이면 자동 생성)
            
        Returns:
            생성된 분석표 파일 경로
        """
        if output_path is None:
            output_path = str(self.raw_excel_path.parent / f"분석표_{self.year}년_{self.quarter}분기_자동생성.xlsx")
        
        # 각 시트 변환
        converted_sheets = {}
        
        for raw_sheet, analysis_sheet in self.SHEET_MAPPING.items():
            try:
                print(f"[변환] {raw_sheet} → {analysis_sheet}")
                df = self._convert_sheet(raw_sheet)
                if df is not None and len(df) > 0:
                    converted_sheets[analysis_sheet] = df
                    print(f"  → {len(df)}행 변환 완료")
            except Exception as e:
                import traceback
                print(f"[오류] {raw_sheet} 변환 실패: {e}")
                traceback.print_exc()
        
        # 엑셀 파일로 저장
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in converted_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"[완료] 분석표 생성: {output_path}")
        return output_path
    
    def _convert_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """개별 시트 변환"""
        if sheet_name not in self.xl.sheet_names:
            print(f"  [경고] 시트 '{sheet_name}'를 찾을 수 없습니다.")
            return None
        
        df = pd.read_excel(self.xl, sheet_name=sheet_name, header=None)
        
        # 시트별 변환 로직
        if sheet_name == '광공업생산':
            return self._convert_index_data(df, '광공업')
        elif sheet_name == '서비스업생산':
            return self._convert_index_data(df, '서비스업')
        elif sheet_name == '소비(소매, 추가)':
            return self._convert_consumption(df)
        elif sheet_name == '고용률':
            return self._convert_employment(df, is_rate=True)
        elif sheet_name == '실업자 수':
            return self._convert_employment(df, is_rate=False)
        elif sheet_name == '품목성질별 물가':
            return self._convert_price(df)
        elif sheet_name == '건설 (공표자료)':
            return self._convert_construction(df)
        elif sheet_name == '수출':
            return self._convert_trade(df, '수출')
        elif sheet_name == '수입':
            return self._convert_trade(df, '수입')
        elif sheet_name == '시도 간 이동':
            return self._convert_migration(df)
        
        return None
    
    def _safe_float(self, val):
        """안전하게 float으로 변환"""
        if pd.isna(val):
            return np.nan
        try:
            return float(val)
        except (ValueError, TypeError):
            return np.nan
    
    def _calculate_yoy_growth(self, current, prev_year):
        """전년동기비 증감률 계산"""
        current = self._safe_float(current)
        prev_year = self._safe_float(prev_year)
        
        if pd.isna(current) or pd.isna(prev_year) or prev_year == 0:
            return np.nan
        
        return ((current - prev_year) / abs(prev_year)) * 100
    
    def _calculate_yoy_diff(self, current, prev_year):
        """전년동기비 차이 계산 (고용률, 실업률 등)"""
        current = self._safe_float(current)
        prev_year = self._safe_float(prev_year)
        
        if pd.isna(current) or pd.isna(prev_year):
            return np.nan
        
        return current - prev_year
    
    def _find_quarter_columns(self, df, header_row=2):
        """분기별 컬럼 인덱스 찾기"""
        quarter_cols = {}
        
        for col_idx in range(len(df.columns)):
            cell = str(df.iloc[header_row, col_idx]).strip()
            
            # 연도 컬럼 (예: 2021, 2022, 2023, 2024)
            if cell.isdigit() and 2010 <= int(cell) <= 2030:
                quarter_cols[f'Y{cell}'] = col_idx
            # 분기 컬럼 (예: 2024  2/4, 2025  1/4p)
            elif '/4' in cell:
                # 정규화
                key = cell.replace(' ', '').replace('p', '').strip()
                quarter_cols[key] = col_idx
        
        return quarter_cols
    
    def _convert_index_data(self, df, data_type='광공업'):
        """지수 데이터 변환 (광공업, 서비스업)"""
        header_row = 2
        quarter_cols = self._find_quarter_columns(df, header_row)
        
        # 현재 분기와 전년동분기 컬럼 찾기
        current_key = f'{self.year}{self.quarter}/4'
        prev_year_key = f'{self.year - 1}{self.quarter}/4'
        
        current_col = quarter_cols.get(current_key)
        prev_year_col = quarter_cols.get(prev_year_key)
        
        if current_col is None:
            # 마지막 컬럼 사용
            current_col = len(df.columns) - 1
            prev_year_col = current_col - 4
        
        print(f"  [컬럼] 당분기: {current_key}(col {current_col}), 전년동기: {prev_year_key}(col {prev_year_col})")
        
        result_data = []
        
        for i in range(3, len(df)):
            region_code = str(df.iloc[i, 0]).strip() if not pd.isna(df.iloc[i, 0]) else ''
            region_name = str(df.iloc[i, 1]).strip() if not pd.isna(df.iloc[i, 1]) else ''
            level = str(df.iloc[i, 2]).strip() if not pd.isna(df.iloc[i, 2]) else ''
            weight = self._safe_float(df.iloc[i, 3])
            industry_code = str(df.iloc[i, 4]).strip() if not pd.isna(df.iloc[i, 4]) else ''
            industry_name = str(df.iloc[i, 5]).strip() if not pd.isna(df.iloc[i, 5]) else ''
            
            if not region_name or not level:
                continue
            
            # 연도별 데이터 및 증감률
            yearly_data = {}
            for year in range(2021, 2025):
                y_key = f'Y{year}'
                if y_key in quarter_cols:
                    current_y = self._safe_float(df.iloc[i, quarter_cols[y_key]])
                    prev_y_key = f'Y{year-1}'
                    if prev_y_key in quarter_cols:
                        prev_y = self._safe_float(df.iloc[i, quarter_cols[prev_y_key]])
                        yearly_data[str(year)] = self._calculate_yoy_growth(current_y, prev_y)
                    else:
                        yearly_data[str(year)] = np.nan
            
            # 분기별 증감률 계산
            quarterly_data = {}
            quarters_needed = [
                (f'{self.year-2}4/4', f'{self.year-3}4/4'),
                (f'{self.year-1}1/4', f'{self.year-2}1/4'),
                (f'{self.year-1}2/4', f'{self.year-2}2/4'),
                (f'{self.year-1}3/4', f'{self.year-2}3/4'),
                (f'{self.year-1}4/4', f'{self.year-2}4/4'),
                (f'{self.year}1/4', f'{self.year-1}1/4'),
                (f'{self.year}2/4', f'{self.year-1}2/4'),
            ]
            
            for current_q, prev_q in quarters_needed:
                if current_q in quarter_cols and prev_q in quarter_cols:
                    current_val = self._safe_float(df.iloc[i, quarter_cols[current_q]])
                    prev_val = self._safe_float(df.iloc[i, quarter_cols[prev_q]])
                    quarterly_data[current_q] = self._calculate_yoy_growth(current_val, prev_val)
            
            # 당분기 증감률
            current_val = self._safe_float(df.iloc[i, current_col])
            prev_val = self._safe_float(df.iloc[i, prev_year_col]) if prev_year_col else np.nan
            yoy_growth = self._calculate_yoy_growth(current_val, prev_val)
            
            # 기여도 계산
            contribution = yoy_growth * weight / 10000 if not pd.isna(yoy_growth) and not pd.isna(weight) else np.nan
            
            row_data = {
                '참고용': f'{region_name}{industry_code}',
                '조회용': f'{region_code}{industry_code}',
                '지역코드': region_code,
                '지역이름': region_name,
                '분류단계': level,
                '중분류순위': '',
                '산업코드': industry_code,
                '산업이름': industry_name,
                '가중치': weight,
            }
            
            # 연도별 증감률 추가
            for year in range(2021, 2025):
                row_data[str(year)] = yearly_data.get(str(year), np.nan)
            
            # 분기별 증감률 추가
            for q_key in quarterly_data.keys():
                col_name = f'{q_key[:4]} {q_key[4:]}'
                row_data[col_name] = quarterly_data[q_key]
            
            # 당분기
            row_data[f'{self.year} {self.quarter}/4'] = yoy_growth
            row_data['증감'] = yoy_growth
            row_data['증감X가중치'] = contribution * 10000 if not pd.isna(contribution) else np.nan
            row_data['대분류'] = ''
            
            result_data.append(row_data)
        
        return pd.DataFrame(result_data)
    
    def _convert_consumption(self, df):
        """소비동향 변환"""
        return self._convert_index_data(df, '소비')
    
    def _convert_employment(self, df, is_rate=True):
        """고용/실업 데이터 변환"""
        header_row = 2
        quarter_cols = self._find_quarter_columns(df, header_row)
        
        # 현재 분기와 전년동분기 컬럼 찾기
        current_key = f'{self.year}{self.quarter}/4'
        prev_year_key = f'{self.year - 1}{self.quarter}/4'
        
        current_col = quarter_cols.get(current_key)
        prev_year_col = quarter_cols.get(prev_year_key)
        
        if current_col is None:
            current_col = len(df.columns) - 1
            prev_year_col = current_col - 4
        
        result_data = []
        
        for i in range(3, len(df)):
            region_code = str(df.iloc[i, 0]).strip() if not pd.isna(df.iloc[i, 0]) else ''
            region_name = str(df.iloc[i, 1]).strip() if not pd.isna(df.iloc[i, 1]) else ''
            level = str(df.iloc[i, 2]).strip() if not pd.isna(df.iloc[i, 2]) else ''
            category = str(df.iloc[i, 3]).strip() if not pd.isna(df.iloc[i, 3]) else ''
            
            if not region_name or not level:
                continue
            
            # 당분기 값과 전년동기비 차이
            current_val = self._safe_float(df.iloc[i, current_col])
            prev_val = self._safe_float(df.iloc[i, prev_year_col])
            yoy_diff = self._calculate_yoy_diff(current_val, prev_val)
            
            row_data = {
                '지역코드': region_code,
                '지역이름': region_name,
                '분류단계': level,
                '구분': category,
            }
            
            # 분기별 데이터 추가
            for q_key, col_idx in quarter_cols.items():
                if '/4' in q_key:
                    val = self._safe_float(df.iloc[i, col_idx])
                    row_data[q_key] = val
            
            row_data['당분기'] = current_val
            row_data['전년동기비'] = yoy_diff
            
            result_data.append(row_data)
        
        return pd.DataFrame(result_data)
    
    def _convert_price(self, df):
        """물가동향 변환"""
        header_row = 2
        quarter_cols = self._find_quarter_columns(df, header_row)
        
        current_key = f'{self.year}{self.quarter}/4'
        prev_year_key = f'{self.year - 1}{self.quarter}/4'
        
        current_col = quarter_cols.get(current_key)
        prev_year_col = quarter_cols.get(prev_year_key)
        
        if current_col is None:
            current_col = len(df.columns) - 1
            prev_year_col = current_col - 4
        
        result_data = []
        
        for i in range(3, len(df)):
            region_name = str(df.iloc[i, 0]).strip() if not pd.isna(df.iloc[i, 0]) else ''
            level = str(df.iloc[i, 1]).strip() if not pd.isna(df.iloc[i, 1]) else ''
            weight = self._safe_float(df.iloc[i, 2])
            category = str(df.iloc[i, 3]).strip() if not pd.isna(df.iloc[i, 3]) else ''
            
            if not region_name or not level:
                continue
            
            current_val = self._safe_float(df.iloc[i, current_col])
            prev_val = self._safe_float(df.iloc[i, prev_year_col])
            yoy_diff = self._calculate_yoy_diff(current_val, prev_val)
            
            row_data = {
                '지역': region_name,
                '분류단계': level,
                '가중치': weight,
                '분류명': category,
                '당분기지수': current_val,
                '전년동기비': yoy_diff,
            }
            
            result_data.append(row_data)
        
        return pd.DataFrame(result_data)
    
    def _convert_construction(self, df):
        """건설동향 변환"""
        header_row = 2
        quarter_cols = self._find_quarter_columns(df, header_row)
        
        current_key = f'{self.year}{self.quarter}/4'
        prev_year_key = f'{self.year - 1}{self.quarter}/4'
        
        current_col = quarter_cols.get(current_key)
        prev_year_col = quarter_cols.get(prev_year_key)
        
        if current_col is None:
            current_col = len(df.columns) - 1
            prev_year_col = current_col - 4
        
        result_data = []
        
        for i in range(3, len(df)):
            region_code = str(df.iloc[i, 0]).strip() if not pd.isna(df.iloc[i, 0]) else ''
            region_name = str(df.iloc[i, 1]).strip() if not pd.isna(df.iloc[i, 1]) else ''
            level = str(df.iloc[i, 2]).strip() if not pd.isna(df.iloc[i, 2]) else ''
            category_code = str(df.iloc[i, 3]).strip() if not pd.isna(df.iloc[i, 3]) else ''
            category_name = str(df.iloc[i, 4]).strip() if not pd.isna(df.iloc[i, 4]) else ''
            
            if not region_name or not level:
                continue
            
            current_val = self._safe_float(df.iloc[i, current_col])
            prev_val = self._safe_float(df.iloc[i, prev_year_col])
            yoy_growth = self._calculate_yoy_growth(current_val, prev_val)
            
            row_data = {
                '지역코드': region_code,
                '지역이름': region_name,
                '분류단계': level,
                '분류코드': category_code,
                '공정이름': category_name,
                '당분기금액': current_val,
                '전년동기금액': prev_val,
                '전년동기비': yoy_growth,
            }
            
            result_data.append(row_data)
        
        return pd.DataFrame(result_data)
    
    def _convert_trade(self, df, trade_type='수출'):
        """수출입 데이터 변환"""
        header_row = 2
        quarter_cols = self._find_quarter_columns(df, header_row)
        
        current_key = f'{self.year}{self.quarter}/4'
        prev_year_key = f'{self.year - 1}{self.quarter}/4'
        
        current_col = quarter_cols.get(current_key)
        prev_year_col = quarter_cols.get(prev_year_key)
        
        if current_col is None:
            current_col = len(df.columns) - 1
            prev_year_col = current_col - 4
        
        result_data = []
        
        for i in range(3, len(df)):
            region_code = str(df.iloc[i, 0]).strip() if not pd.isna(df.iloc[i, 0]) else ''
            region_name = str(df.iloc[i, 1]).strip() if not pd.isna(df.iloc[i, 1]) else ''
            level = str(df.iloc[i, 2]).strip() if not pd.isna(df.iloc[i, 2]) else ''
            
            if not region_name or not level:
                continue
            
            # 수출/수입은 컬럼 구조가 다름
            try:
                item_code = str(df.iloc[i, 4]).strip() if not pd.isna(df.iloc[i, 4]) else ''
                item_name = str(df.iloc[i, 5]).strip() if not pd.isna(df.iloc[i, 5]) else ''
            except:
                item_code = ''
                item_name = ''
            
            current_val = self._safe_float(df.iloc[i, current_col])
            prev_val = self._safe_float(df.iloc[i, prev_year_col])
            yoy_growth = self._calculate_yoy_growth(current_val, prev_val)
            
            row_data = {
                '지역코드': region_code,
                '지역이름': region_name,
                '분류단계': level,
                '상품코드': item_code,
                '상품이름': item_name,
                '당분기금액': current_val,
                '전년동기금액': prev_val,
                '전년동기비': yoy_growth,
            }
            
            result_data.append(row_data)
        
        return pd.DataFrame(result_data)
    
    def _convert_migration(self, df):
        """인구이동 변환"""
        header_row = 2
        quarter_cols = self._find_quarter_columns(df, header_row)
        
        current_key = f'{self.year}{self.quarter}/4'
        prev_year_key = f'{self.year - 1}{self.quarter}/4'
        
        current_col = quarter_cols.get(current_key)
        prev_year_col = quarter_cols.get(prev_year_key)
        
        if current_col is None:
            current_col = len(df.columns) - 1
            prev_year_col = current_col - 4
        
        result_data = []
        region_data = {}  # 지역별 유입/유출 데이터 저장
        
        for i in range(3, len(df)):
            region_code = str(df.iloc[i, 0]).strip() if not pd.isna(df.iloc[i, 0]) else ''
            region_name = str(df.iloc[i, 1]).strip() if not pd.isna(df.iloc[i, 1]) else ''
            category = str(df.iloc[i, 2]).strip() if not pd.isna(df.iloc[i, 2]) else ''
            
            if not region_name or not category:
                continue
            
            current_val = self._safe_float(df.iloc[i, current_col])
            prev_val = self._safe_float(df.iloc[i, prev_year_col])
            
            if region_name not in region_data:
                region_data[region_name] = {'code': region_code, 'inflow': 0, 'outflow': 0, 'prev_inflow': 0, 'prev_outflow': 0}
            
            if '유입' in category:
                region_data[region_name]['inflow'] = current_val
                region_data[region_name]['prev_inflow'] = prev_val
            elif '유출' in category:
                region_data[region_name]['outflow'] = current_val
                region_data[region_name]['prev_outflow'] = prev_val
        
        # 순이동 계산
        for region_name, data in region_data.items():
            net_current = data['inflow'] - data['outflow'] if not pd.isna(data['inflow']) and not pd.isna(data['outflow']) else np.nan
            net_prev = data['prev_inflow'] - data['prev_outflow'] if not pd.isna(data['prev_inflow']) and not pd.isna(data['prev_outflow']) else np.nan
            net_diff = net_current - net_prev if not pd.isna(net_current) and not pd.isna(net_prev) else np.nan
            
            result_data.append({
                '지역코드': data['code'],
                '지역이름': region_name,
                '분류단계': '0',
                '유입인구': data['inflow'],
                '유출인구': data['outflow'],
                '순이동': net_current,
                '전년동기순이동': net_prev,
                '전년동기비': net_diff,
            })
        
        return pd.DataFrame(result_data)
    
    def extract_grdp_data(self) -> Dict:
        """GRDP 데이터 추출
        
        Returns:
            GRDP 보고서용 데이터 딕셔너리
        """
        if '분기 GRDP' not in self.xl.sheet_names:
            print("[경고] '분기 GRDP' 시트를 찾을 수 없습니다.")
            return self._get_placeholder_grdp()
        
        df = pd.read_excel(self.xl, sheet_name='분기 GRDP', header=None)
        
        # 헤더 행 (3행)
        header_row = 2
        
        # 최신 분기 컬럼 찾기 (2025 2/4p)
        current_quarter_col = -1
        prev_year_quarter_col = -1
        
        for col_idx in range(len(df.columns)):
            cell = str(df.iloc[header_row, col_idx]).strip()
            if '2025' in cell and '2/4' in cell:
                current_quarter_col = col_idx
            elif '2024' in cell and '2/4' in cell:
                prev_year_quarter_col = col_idx
        
        if current_quarter_col == -1:
            # 마지막 컬럼 사용
            current_quarter_col = len(df.columns) - 1
            prev_year_quarter_col = current_quarter_col - 4
        
        print(f"[GRDP] 당분기 컬럼: {current_quarter_col}, 전년동기 컬럼: {prev_year_quarter_col}")
        
        # 지역별 데이터 추출
        regional_data = []
        national_data = {}
        
        # 그룹 매핑
        region_groups = {
            '서울': '경인', '인천': '경인', '경기': '경인',
            '대전': '충청', '세종': '충청', '충북': '충청', '충남': '충청',
            '광주': '호남', '전북': '호남', '전남': '호남', '제주': '호남',
            '대구': '동북', '경북': '동북', '강원': '동북',
            '부산': '동남', '울산': '동남', '경남': '동남',
        }
        
        # 지역별 데이터 수집 (분류단계별로 저장)
        # 각 지역의 분류단계 1 항목들을 저장 (item 이름 기준)
        region_values = {}  # {지역: {'total': {...}, 'industries': [{item, current, prev}]}}
        
        for i in range(3, len(df)):
            region = str(df.iloc[i, 1]).strip() if not pd.isna(df.iloc[i, 1]) else ''
            level = str(df.iloc[i, 2]).strip() if not pd.isna(df.iloc[i, 2]) else ''
            item = str(df.iloc[i, 4]).strip() if not pd.isna(df.iloc[i, 4]) else ''
            
            if not region or region not in self.REGIONS_ORDER:
                continue
            
            current_val = df.iloc[i, current_quarter_col] if not pd.isna(df.iloc[i, current_quarter_col]) else 0.0
            prev_year_val = df.iloc[i, prev_year_quarter_col] if not pd.isna(df.iloc[i, prev_year_quarter_col]) else 0.0
            
            if region not in region_values:
                region_values[region] = {'total': None, 'industries': []}
            
            if level == '0':
                region_values[region]['total'] = {
                    'item': item,
                    'current': float(current_val),
                    'prev_year': float(prev_year_val),
                }
            elif level == '1':
                region_values[region]['industries'].append({
                    'item': item,
                    'current': float(current_val),
                    'prev_year': float(prev_year_val),
                })
        
        # 성장률 및 기여도 계산
        for region in self.REGIONS_ORDER:
            if region not in region_values:
                continue
            
            values = region_values[region]
            
            # 총 GRDP (분류단계 0)
            total = values.get('total', {'current': 0, 'prev_year': 0})
            if total is None:
                total = {'current': 0, 'prev_year': 0}
            total_current = total['current']
            total_prev = total['prev_year']
            
            # 성장률 계산
            if total_prev > 0:
                growth_rate = round(((total_current - total_prev) / total_prev) * 100, 1)
            else:
                growth_rate = 0.0
            
            # 산업별 기여도 계산 (분류단계 1)
            manufacturing_contrib = 0.0
            construction_contrib = 0.0
            service_contrib = 0.0
            other_contrib = 0.0
            
            for industry in values.get('industries', []):
                item_name = industry.get('item', '')
                contrib = self._calculate_contribution(
                    industry['current'], industry['prev_year'], total_prev
                )
                
                if '제조' in item_name or '광업' in item_name:
                    manufacturing_contrib = contrib
                elif '건설' in item_name:
                    construction_contrib = contrib
                elif '서비스' in item_name:
                    service_contrib = contrib
                elif '기타' in item_name or '순생산' in item_name:
                    other_contrib = contrib
            
            region_data = {
                'region': region,
                'region_group': region_groups.get(region, ''),
                'growth_rate': growth_rate,
                'manufacturing': manufacturing_contrib,
                'construction': construction_contrib,
                'service': service_contrib,
                'other': other_contrib,
                'placeholder': False
            }
            
            if region == '전국':
                national_data = region_data.copy()
            
            regional_data.append(region_data)
        
        # 성장률 1위 지역 찾기
        non_national = [r for r in regional_data if r['region'] != '전국']
        if non_national:
            top_region = max(non_national, key=lambda x: x['growth_rate'])
        else:
            top_region = {'name': '-', 'growth_rate': 0.0, 'manufacturing': 0.0, 
                         'construction': 0.0, 'service': 0.0, 'other': 0.0}
        
        return {
            'report_info': {
                'year': self.year,
                'quarter': self.quarter,
                'page_number': 20
            },
            'national_summary': {
                'growth_rate': national_data.get('growth_rate', 0.0),
                'direction': '증가' if national_data.get('growth_rate', 0) >= 0 else '감소',
                'contributions': {
                    'manufacturing': national_data.get('manufacturing', 0.0),
                    'construction': national_data.get('construction', 0.0),
                    'service': national_data.get('service', 0.0),
                    'other': national_data.get('other', 0.0),
                },
                'placeholder': False
            },
            'top_region': {
                'name': top_region.get('region', '-'),
                'growth_rate': top_region.get('growth_rate', 0.0),
                'contributions': {
                    'manufacturing': top_region.get('manufacturing', 0.0),
                    'construction': top_region.get('construction', 0.0),
                    'service': top_region.get('service', 0.0),
                    'other': top_region.get('other', 0.0),
                },
                'placeholder': False
            },
            'regional_data': regional_data,
            'chart_config': {
                'y_axis': {
                    'min': -6,
                    'max': 8,
                    'step': 2
                }
            }
        }
    
    def _calculate_contribution(self, current: float, prev: float, total_prev: float) -> float:
        """산업별 기여도 계산"""
        if total_prev == 0:
            return 0.0
        return round(((current - prev) / total_prev) * 100, 1)
    
    def _get_placeholder_grdp(self) -> Dict:
        """플레이스홀더 GRDP 데이터"""
        regional_data = []
        region_groups = {
            '서울': '경인', '인천': '경인', '경기': '경인',
            '대전': '충청', '세종': '충청', '충북': '충청', '충남': '충청',
            '광주': '호남', '전북': '호남', '전남': '호남', '제주': '호남',
            '대구': '동북', '경북': '동북', '강원': '동북',
            '부산': '동남', '울산': '동남', '경남': '동남',
        }
        
        for region in self.REGIONS_ORDER:
            regional_data.append({
                'region': region,
                'region_group': region_groups.get(region, ''),
                'growth_rate': 0.0,
                'manufacturing': 0.0,
                'construction': 0.0,
                'service': 0.0,
                'other': 0.0,
                'placeholder': True
            })
        
        return {
            'report_info': {
                'year': self.year,
                'quarter': self.quarter,
                'page_number': 20
            },
            'national_summary': {
                'growth_rate': 0.0,
                'direction': '증가',
                'contributions': {
                    'manufacturing': 0.0,
                    'construction': 0.0,
                    'service': 0.0,
                    'other': 0.0,
                },
                'placeholder': True
            },
            'top_region': {
                'name': '-',
                'growth_rate': 0.0,
                'contributions': {
                    'manufacturing': 0.0,
                    'construction': 0.0,
                    'service': 0.0,
                    'other': 0.0,
                },
                'placeholder': True
            },
            'regional_data': regional_data,
            'chart_config': {
                'y_axis': {
                    'min': -6,
                    'max': 8,
                    'step': 2
                }
            }
        }
    
    def save_grdp_json(self, output_path: str) -> Dict:
        """GRDP 데이터를 JSON으로 저장"""
        data = self.extract_grdp_data()
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return data


def convert_raw_to_analysis(raw_excel_path: str, output_path: str = None) -> Tuple[str, Dict]:
    """기초자료 수집표 → 분석표 변환 및 GRDP 추출
    
    Args:
        raw_excel_path: 기초자료 수집표 경로
        output_path: 분석표 출력 경로 (None이면 자동 생성)
        
    Returns:
        (분석표 경로, GRDP 데이터)
    """
    converter = DataConverter(raw_excel_path)
    analysis_path = converter.convert_all(output_path)
    grdp_data = converter.extract_grdp_data()
    
    return analysis_path, grdp_data


if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='기초자료 수집표 → 분석표 변환')
    parser.add_argument('input', type=str, help='기초자료 수집표 엑셀 파일')
    parser.add_argument('--output', type=str, default=None, help='출력 분석표 경로')
    parser.add_argument('--grdp-json', type=str, default='grdp_data.json', help='GRDP JSON 출력 경로')
    
    args = parser.parse_args()
    
    # 변환 실행
    analysis_path, grdp_data = convert_raw_to_analysis(args.input, args.output)
    
    # GRDP JSON 저장
    with open(args.grdp_json, 'w', encoding='utf-8') as f:
        json.dump(grdp_data, f, ensure_ascii=False, indent=2)
    
    print(f"\n=== 변환 완료 ===")
    print(f"분석표: {analysis_path}")
    print(f"GRDP JSON: {args.grdp_json}")
    print(f"전국 성장률: {grdp_data['national_summary']['growth_rate']}%")
    print(f"1위 지역: {grdp_data['top_region']['name']} ({grdp_data['top_region']['growth_rate']}%)")
