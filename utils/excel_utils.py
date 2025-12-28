# -*- coding: utf-8 -*-
"""
엑셀 파일 관련 유틸리티 함수
"""

import importlib.util
from pathlib import Path
import pandas as pd

from config.settings import TEMPLATES_DIR


def load_generator_module(generator_name):
    """동적으로 generator 모듈 로드"""
    generator_path = TEMPLATES_DIR / generator_name
    if not generator_path.exists():
        return None
    
    spec = importlib.util.spec_from_file_location(
        generator_name.replace('.py', ''),
        str(generator_path)
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def extract_year_quarter_from_excel(filepath):
    """엑셀 파일에서 최신 연도와 분기 추출 (분석표용)"""
    try:
        xl = pd.ExcelFile(filepath)
        # A 분석 시트에서 연도/분기 정보 추출 시도
        df = pd.read_excel(xl, sheet_name='A 분석', header=None)
        
        # 일반적으로 컬럼 헤더에서 연도/분기 정보를 찾음
        for row_idx in range(min(5, len(df))):
            for col_idx in range(len(df.columns)):
                cell = str(df.iloc[row_idx, col_idx])
                if '2025.2/4' in cell or '25.2/4' in cell:
                    return 2025, 2
                elif '2025.1/4' in cell or '25.1/4' in cell:
                    return 2025, 1
                elif '2024.4/4' in cell or '24.4/4' in cell:
                    return 2024, 4
        
        # 파일명에서 추출 시도
        filename = Path(filepath).stem
        if '25년' in filename and '2분기' in filename:
            return 2025, 2
        elif '25년' in filename and '1분기' in filename:
            return 2025, 1
        
        return 2025, 2  # 기본값
    except Exception as e:
        print(f"연도/분기 추출 오류: {e}")
        return 2025, 2


def extract_year_quarter_from_raw(filepath):
    """기초자료 수집표에서 연도와 분기 추출"""
    try:
        # 먼저 파일명에서 추출 시도
        filename = Path(filepath).stem
        import re
        
        # 파일명 패턴: "기초자료 수집표_2025년 2분기" 또는 "25년_2분기" 등
        year_match = re.search(r'(20\d{2}|25|24)년?', filename)
        quarter_match = re.search(r'(\d)분기', filename)
        
        if year_match and quarter_match:
            year = int(year_match.group(1))
            if year < 100:
                year += 2000
            quarter = int(quarter_match.group(1))
            return year, quarter
        
        # 시트에서 연도/분기 정보 추출 시도
        xl = pd.ExcelFile(filepath)
        
        # 기초자료 수집표의 첫 번째 시트에서 헤더 확인
        for sheet_name in xl.sheet_names[:3]:
            try:
                df = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=10)
                for row_idx in range(min(10, len(df))):
                    for col_idx in range(min(20, len(df.columns))):
                        cell = str(df.iloc[row_idx, col_idx])
                        if '2025.2/4' in cell or '25.2/4' in cell or '2025년 2분기' in cell:
                            return 2025, 2
                        elif '2025.1/4' in cell or '25.1/4' in cell or '2025년 1분기' in cell:
                            return 2025, 1
                        elif '2024.4/4' in cell or '24.4/4' in cell or '2024년 4분기' in cell:
                            return 2024, 4
            except:
                continue
        
        return 2025, 2  # 기본값
    except Exception as e:
        print(f"기초자료 연도/분기 추출 오류: {e}")
        return 2025, 2


def detect_file_type(filepath: str) -> str:
    """엑셀 파일 유형 자동 감지 (기초자료 수집표 vs 분석표)"""
    try:
        xl = pd.ExcelFile(filepath)
        sheet_names = xl.sheet_names
        
        # 기초자료 수집표 특징: '광공업생산', '서비스업생산', '분기 GRDP' 등의 시트
        raw_indicators = ['광공업생산', '서비스업생산', '고용률', '분기 GRDP']
        raw_count = sum(1 for s in raw_indicators if s in sheet_names)
        
        # 분석표 특징: 'A 분석', 'B 분석' 등의 시트
        analysis_indicators = ['A 분석', 'B 분석', 'C 분석', 'D(고용률)분석']
        analysis_count = sum(1 for s in analysis_indicators if s in sheet_names)
        
        if raw_count >= 2:
            return 'raw'  # 기초자료 수집표
        elif analysis_count >= 2:
            return 'analysis'  # 분석표
        else:
            # 파일명으로 추정
            filename = Path(filepath).stem.lower()
            if '기초' in filename or '수집' in filename:
                return 'raw'
            elif '분석' in filename:
                return 'analysis'
            return 'analysis'  # 기본값
    except Exception as e:
        print(f"[경고] 파일 유형 감지 실패: {e}")
        return 'analysis'

