#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
기초자료 수집표 → 분석표 변환 모듈

기존 분석표를 템플릿으로 사용하고, 기초자료의 데이터를 집계 시트에 복사합니다.
분석 시트의 엑셀 수식은 그대로 유지되어 자동으로 계산됩니다.
"""

import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import json
import shutil


class DataConverter:
    """기초자료 수집표 → 분석표 변환기 (수식 보존 방식)"""
    
    # 기초자료 시트 → 분석표 집계 시트 매핑
    SHEET_MAPPING = {
        '광공업생산': 'A(광공업생산)집계',
        '서비스업생산': 'B(서비스업생산)집계',
        '소비(소매, 추가)': 'C(소비)집계',
        '고용률': 'D(고용률)집계',
        '실업자 수': 'D(실업)집계',
        '지출목적별 물가': 'E(지출목적물가)집계',
        '품목성질별 물가': 'E(품목성질물가)집계',
        '건설 (공표자료)': "F'(건설)집계",
        '수출': 'G(수출)집계',
        '수입': 'H(수입)집계',
        '시도 간 이동': 'I(순인구이동)집계',
    }
    
    # 지역 순서
    REGIONS_ORDER = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                     '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
    
    def __init__(self, raw_excel_path: str, template_path: str = None):
        """
        Args:
            raw_excel_path: 기초자료 수집표 엑셀 파일 경로
            template_path: 분석표 템플릿 경로 (None이면 기본 템플릿 사용)
        """
        self.raw_excel_path = Path(raw_excel_path)
        
        # 템플릿 경로 설정
        if template_path:
            self.template_path = Path(template_path)
        else:
            # 기본 템플릿 (프로젝트 내 분석표)
            self.template_path = Path(__file__).parent / '분석표_25년 2분기_캡스톤.xlsx'
            if not self.template_path.exists():
                # 대안 경로
                self.template_path = self.raw_excel_path.parent / '분석표_25년 2분기_캡스톤.xlsx'
        
        self.year = None
        self.quarter = None
        self._detect_year_quarter()
    
    def _detect_year_quarter(self):
        """기초자료에서 연도/분기 자동 추출"""
        import re
        
        # 1. 먼저 기초자료 헤더에서 추출 시도
        try:
            xl = pd.ExcelFile(self.raw_excel_path)
            
            # 광공업생산 시트 또는 첫 번째 데이터 시트에서 헤더 확인
            for sheet_name in ['광공업생산', '서비스업생산', '소비(소매, 추가)']:
                if sheet_name in xl.sheet_names:
                    df = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=4)
                    
                    # 헤더 행(3행)에서 가장 최신 분기 찾기
                    latest_year = 0
                    latest_quarter = 0
                    
                    for col_idx in range(len(df.columns)):
                        header_val = str(df.iloc[2, col_idx]) if not pd.isna(df.iloc[2, col_idx]) else ''
                        
                        # "2025  2/4p" 또는 "2025 2/4" 패턴 찾기
                        match = re.search(r'(\d{4})\s*(\d)/4', header_val)
                        if match:
                            year = int(match.group(1))
                            quarter = int(match.group(2))
                            
                            # 가장 최신 연도/분기 저장
                            if year > latest_year or (year == latest_year and quarter > latest_quarter):
                                latest_year = year
                                latest_quarter = quarter
                    
                    if latest_year > 0 and latest_quarter > 0:
                        self.year = latest_year
                        self.quarter = latest_quarter
                        print(f"[자동감지] 기초자료에서 연도/분기 추출: {self.year}년 {self.quarter}분기")
                        return
                    break
        except Exception as e:
            print(f"[경고] 기초자료에서 연도/분기 추출 실패: {e}")
        
        # 2. 파일명에서 추출 시도
        filename = self.raw_excel_path.stem
        
        # 다양한 패턴 매칭
        patterns = [
            r'(\d{4})년\s*(\d)분기',  # 2025년 2분기
            r'(\d{2})년\s*(\d)분기',   # 25년 2분기
            r'(\d{4})_(\d)',           # 2025_2
        ]
        
        for pattern in patterns:
            match = re.search(pattern, filename)
            if match:
                year_str = match.group(1)
                quarter = int(match.group(2))
                
                # 2자리 연도 처리
                if len(year_str) == 2:
                    year = 2000 + int(year_str)
                else:
                    year = int(year_str)
                
                self.year = year
                self.quarter = quarter
                print(f"[자동감지] 파일명에서 연도/분기 추출: {self.year}년 {self.quarter}분기")
                return
        
        # 3. 기본값
        self.year, self.quarter = 2025, 2
        print(f"[기본값] 연도/분기: {self.year}년 {self.quarter}분기")
    
    def convert_all(self, output_path: str = None) -> str:
        """분석표 생성 (템플릿 복사 + 집계 시트 데이터 교체)
        
        Args:
            output_path: 출력 파일 경로 (None이면 자동 생성)
            
        Returns:
            생성된 분석표 파일 경로
        """
        if output_path is None:
            output_path = str(self.raw_excel_path.parent / f"분석표_{self.year}년_{self.quarter}분기_자동생성.xlsx")
        
        # 1. 템플릿 파일 복사
        if not self.template_path.exists():
            raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {self.template_path}")
        
        print(f"[변환] 템플릿 복사: {self.template_path.name}")
        shutil.copy(self.template_path, output_path)
        
        # 2. 복사된 파일 열기 (수식 보존)
        wb = openpyxl.load_workbook(output_path)
        
        # 3. 이용관련 시트에서 연도/분기 설정 (핵심!)
        if '이용관련' in wb.sheetnames:
            ws_config = wb['이용관련']
            ws_config.cell(row=17, column=11).value = self.year    # K17: 연도
            ws_config.cell(row=17, column=13).value = self.quarter # M17: 분기
            print(f"[변환] 이용관련 시트 설정: {self.year}년 {self.quarter}분기")
        else:
            print("[경고] '이용관련' 시트를 찾을 수 없습니다.")
        
        # 4. 기초자료 열기
        raw_xl = pd.ExcelFile(self.raw_excel_path)
        
        # 5. 각 집계 시트에 데이터 복사
        for raw_sheet, target_sheet in self.SHEET_MAPPING.items():
            if raw_sheet in raw_xl.sheet_names and target_sheet in wb.sheetnames:
                print(f"[변환] {raw_sheet} → {target_sheet}")
                self._copy_sheet_data(raw_xl, raw_sheet, wb[target_sheet])
            else:
                if raw_sheet not in raw_xl.sheet_names:
                    print(f"  [경고] 기초자료에 '{raw_sheet}' 시트 없음")
                if target_sheet not in wb.sheetnames:
                    print(f"  [경고] 템플릿에 '{target_sheet}' 시트 없음")
        
        # 5. 저장
        wb.save(output_path)
        wb.close()
        
        print(f"[완료] 분석표 생성: {output_path}")
        return output_path
    
    def _copy_sheet_data(self, raw_xl: pd.ExcelFile, raw_sheet: str, target_ws):
        """기초자료 시트의 데이터를 분석표 집계 시트에 복사"""
        from openpyxl.cell.cell import MergedCell
        
        # 기초자료 읽기 (헤더 없이)
        raw_df = pd.read_excel(raw_xl, sheet_name=raw_sheet, header=None)
        
        copied_count = 0
        skipped_count = 0
        
        # 데이터 복사 (행/열 순회)
        for row_idx in range(len(raw_df)):
            for col_idx in range(len(raw_df.columns)):
                value = raw_df.iloc[row_idx, col_idx]
                
                # NaN 처리
                if pd.isna(value):
                    continue
                
                # openpyxl은 1-based 인덱스
                cell = target_ws.cell(row=row_idx + 1, column=col_idx + 1)
                
                # 병합된 셀은 건너뛰기
                if isinstance(cell, MergedCell):
                    skipped_count += 1
                    continue
                
                # 값 복사
                try:
                    cell.value = value
                    copied_count += 1
                except Exception:
                    skipped_count += 1
        
        print(f"  → {copied_count}개 셀 복사 ({skipped_count}개 건너뜀)")
    
    def _calculate_analysis_values(self, wb):
        """분석 시트의 수식을 계산하여 값으로 저장 (캐시용)
        
        엑셀에서 열지 않아도 Python에서 값을 읽을 수 있도록
        분석 시트의 수식 셀에 계산된 값을 추가로 저장
        """
        # 분석 시트 목록
        analysis_sheets = [
            ('A 분석', 'A(광공업생산)집계'),
            ('B 분석', 'B(서비스업생산)집계'),
            ('C 분석', 'C(소비)집계'),
            ("D(고용률)분석", 'D(고용률)집계'),
            ("D(실업)분석", 'D(실업)집계'),
            ('E(지출목적물가) 분석', 'E(지출목적물가)집계'),
            ('E(품목성질물가)분석', 'E(품목성질물가)집계'),
            ("F'분석", "F'(건설)집계"),
            ('G 분석', 'G(수출)집계'),
            ('H 분석', 'H(수입)집계'),
            ('I(순인구이동)분석', 'I(순인구이동)집계'),
        ]
        
        for analysis_sheet, source_sheet in analysis_sheets:
            if analysis_sheet not in wb.sheetnames:
                continue
            if source_sheet not in wb.sheetnames:
                continue
                
            ws = wb[analysis_sheet]
            source_ws = wb[source_sheet]
            
            # 수식이 있는 셀을 찾아서 계산
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                        # 수식 파싱 및 계산
                        calculated = self._evaluate_formula(cell.value, wb)
                        if calculated is not None:
                            # 수식을 유지하면서 캐시된 값도 저장
                            # openpyxl에서는 이게 안되므로 수식을 계산된 값으로 대체
                            pass  # 수식을 유지하는 것이 사용자 요청
    
    def _evaluate_formula(self, formula: str, wb) -> Optional[float]:
        """엑셀 수식 평가 (간단한 수식만 지원)"""
        import re
        
        # IFERROR 처리
        if formula.startswith('=IFERROR('):
            # 내부 수식 추출
            inner = formula[9:-1]  # =IFERROR( 와 ) 제거
            # 마지막 ,"없음") 부분 제거
            if ',"없음"' in inner:
                inner = inner.replace(',"없음"', '').replace(',"없음")', ')')
            return self._evaluate_formula('=' + inner.split(',')[0], wb)
        
        # 셀 참조 패턴: 'Sheet'!Cell 또는 Sheet!Cell
        cell_pattern = r"'?([^'!]+)'?!([A-Z]+)(\d+)"
        
        try:
            matches = re.findall(cell_pattern, formula)
            if not matches:
                return None
            
            values = {}
            for sheet_name, col, row in matches:
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    cell = ws[f"{col}{row}"]
                    val = cell.value
                    if val is None or (isinstance(val, str) and not val.replace('.','').replace('-','').isdigit()):
                        return None
                    values[f"'{sheet_name}'!{col}{row}"] = float(val) if val else 0
            
            # 간단한 증감률 계산: (A-B)/B*100
            if len(matches) == 2 and '-' in formula and '/' in formula and '*100' in formula:
                keys = list(values.keys())
                if len(keys) == 2:
                    a, b = values[keys[0]], values[keys[1]]
                    if b != 0:
                        return round((a - b) / b * 100, 2)
            
            return None
        except Exception:
            return None
    
    def extract_grdp_data(self) -> Dict:
        """GRDP 데이터 추출
        
        Returns:
            GRDP 보고서용 데이터 딕셔너리
        """
        raw_xl = pd.ExcelFile(self.raw_excel_path)
        
        if '분기 GRDP' not in raw_xl.sheet_names:
            print("[경고] '분기 GRDP' 시트를 찾을 수 없습니다.")
            return self._get_placeholder_grdp()
        
        df = pd.read_excel(raw_xl, sheet_name='분기 GRDP', header=None)
        
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
        
        # 지역별 데이터 수집
        region_values = {}
        
        for i in range(3, len(df)):
            region = str(df.iloc[i, 1]).strip() if not pd.isna(df.iloc[i, 1]) else ''
            level = str(df.iloc[i, 2]).strip() if not pd.isna(df.iloc[i, 2]) else ''
            item = str(df.iloc[i, 4]).strip() if not pd.isna(df.iloc[i, 4]) else ''
            
            if not region or region not in self.REGIONS_ORDER:
                continue
            
            current_val = self._safe_float(df.iloc[i, current_quarter_col])
            prev_year_val = self._safe_float(df.iloc[i, prev_year_quarter_col])
            
            if region not in region_values:
                region_values[region] = {'total': None, 'industries': []}
            
            if level == '0':
                region_values[region]['total'] = {
                    'item': item,
                    'current': current_val,
                    'prev_year': prev_year_val,
                }
            elif level == '1':
                region_values[region]['industries'].append({
                    'item': item,
                    'current': current_val,
                    'prev_year': prev_year_val,
                })
        
        # 성장률 및 기여도 계산
        for region in self.REGIONS_ORDER:
            if region not in region_values:
                continue
            
            values = region_values[region]
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
            
            # 산업별 기여도 계산
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
            
            region_entry = {
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
                national_data = region_entry.copy()
            
            regional_data.append(region_entry)
        
        # 성장률 1위 지역 찾기
        non_national = [r for r in regional_data if r['region'] != '전국']
        if non_national:
            top_region = max(non_national, key=lambda x: x['growth_rate'])
        else:
            top_region = {'region': '-', 'growth_rate': 0.0, 'manufacturing': 0.0, 
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
    
    def _safe_float(self, val) -> float:
        """안전하게 float으로 변환"""
        if pd.isna(val):
            return 0.0
        try:
            return float(val)
        except (ValueError, TypeError):
            return 0.0
    
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


def detect_file_type(excel_path: str) -> str:
    """엑셀 파일 유형 감지 (기초자료 vs 분석표)"""
    try:
        xl = pd.ExcelFile(excel_path)
        sheet_names = xl.sheet_names
        
        # '분석' 키워드가 포함된 시트가 있으면 분석표
        analysis_sheets = [s for s in sheet_names if '분석' in s]
        if analysis_sheets:
            return 'analysis'
        
        # 기초자료 특유의 시트가 있으면 기초자료
        raw_indicators = ['광공업생산', '서비스업생산', '분기 GRDP', '완료체크']
        if any(ind in sheet_names for ind in raw_indicators):
            return 'raw'
        
        return 'unknown'
    except Exception as e:
        print(f"파일 유형 감지 실패: {e}")
        return 'unknown'


def convert_raw_to_analysis(raw_excel_path: str, output_path: str = None, 
                           template_path: str = None) -> Tuple[str, Dict]:
    """기초자료 수집표 → 분석표 변환 및 GRDP 추출
    
    Args:
        raw_excel_path: 기초자료 수집표 경로
        output_path: 분석표 출력 경로 (None이면 자동 생성)
        template_path: 분석표 템플릿 경로 (None이면 기본 템플릿 사용)
        
    Returns:
        (분석표 경로, GRDP 데이터)
    """
    converter = DataConverter(raw_excel_path, template_path)
    analysis_path = converter.convert_all(output_path)
    grdp_data = converter.extract_grdp_data()
    
    return analysis_path, grdp_data


if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='기초자료 수집표 → 분석표 변환')
    parser.add_argument('input', type=str, help='기초자료 수집표 엑셀 파일')
    parser.add_argument('--output', type=str, default=None, help='출력 분석표 경로')
    parser.add_argument('--template', type=str, default=None, help='분석표 템플릿 경로')
    parser.add_argument('--grdp-json', type=str, default='grdp_data.json', help='GRDP JSON 출력 경로')
    
    args = parser.parse_args()
    
    # 변환 실행
    analysis_path, grdp_data = convert_raw_to_analysis(
        args.input, args.output, args.template
    )
    
    # GRDP JSON 저장
    with open(args.grdp_json, 'w', encoding='utf-8') as f:
        json.dump(grdp_data, f, ensure_ascii=False, indent=2)
    
    print(f"\n=== 변환 완료 ===")
    print(f"분석표: {analysis_path}")
    print(f"GRDP JSON: {args.grdp_json}")
    print(f"전국 성장률: {grdp_data['national_summary']['growth_rate']}%")
    print(f"1위 지역: {grdp_data['top_region']['name']} ({grdp_data['top_region']['growth_rate']}%)")
