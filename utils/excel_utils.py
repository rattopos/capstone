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
    """엑셀 파일 유형 자동 감지 (기초자료 수집표 vs 분석표)
    
    기초자료 수집표와 분석표는 시트 구조가 완전히 다릅니다:
    
    [기초자료 수집표] 17개 시트
    - 시트명 예: '광공업생산', '서비스업생산', '소비(소매, 추가)', '고용률', 
                 '실업자 수', '수출', '수입', '분기 GRDP', '완료체크' 등
    - 특징: 원본 데이터 형태, 한글 시트명
    
    [분석표] 42개 시트
    - 시트명 예: 'A(광공업생산)집계', 'A 분석', 'A 참고', 'B(서비스업생산)집계' 등
    - 특징: 알파벳+키워드(집계/분석/참고) 패턴, '본청', '이용관련' 시트 존재
    
    두 파일은 공통 시트가 하나도 없습니다.
    
    Returns:
        'raw': 기초자료 수집표
        'analysis': 분석표
        'unknown': 알 수 없는 형식
    """
    try:
        xl = pd.ExcelFile(filepath)
        sheet_names = set(xl.sheet_names)
        sheet_count = len(sheet_names)
        
        # ========================================
        # 기초자료 수집표 전용 시트 (분석표에는 없음)
        # ========================================
        RAW_ONLY_SHEETS = {
            '광공업생산', '서비스업생산', '소비(소매, 추가)', 
            '고용', '고용(kosis)', '고용률', '실업자 수',
            '지출목적별 물가', '품목성질별 물가', '건설 (공표자료)',
            '수출', '수입', '연령별 인구이동', '시도 간 이동',
            '시군구인구이동', '분기 GRDP', '완료체크'
        }
        
        # ========================================
        # 분석표 전용 시트 (기초자료에는 없음)
        # ========================================
        ANALYSIS_ONLY_SHEETS = {
            '본청', '시도별 현황', '지방청 이용자료', '이용관련',
            'A(광공업생산)집계', 'A 분석', 'A 참고',
            'B(서비스업생산)집계', 'B 분석', 'B 참고',
            'C(소비)집계', 'C 분석', 'C 참고',
            'D(고용률)집계', 'D(고용률)분석',
            'D(실업)집계', 'D(실업)분석',
            'E(지출목적물가)집계', 'E(지출목적물가) 분석',
            'E(품목성질물가)집계', 'E(품목성질물가)분석',
            "F'(건설)집계", "F'분석",
            'G(수출)집계', 'G 분석',
            'H(수입)집계', 'H 분석',
            'I(순인구이동)집계',
        }
        
        # 매칭된 시트 수 계산
        raw_matches = len(sheet_names & RAW_ONLY_SHEETS)
        analysis_matches = len(sheet_names & ANALYSIS_ONLY_SHEETS)
        
        # ========================================
        # 판정 로직 (상호 배타적 시트 기반)
        # ========================================
        
        # 1. 분석표 전용 시트가 있으면 분석표
        if analysis_matches >= 3:
            return 'analysis'
        
        # 2. 기초자료 전용 시트가 있으면 기초자료
        if raw_matches >= 3:
            return 'raw'
        
        # 3. 시트명 패턴으로 추가 판정
        # 분석표: '분석', '집계', '참고' 키워드 포함
        analysis_pattern_count = sum(1 for s in sheet_names 
                                      if '분석' in s or '집계' in s or '참고' in s)
        if analysis_pattern_count >= 3:
            return 'analysis'
        
        # 4. 시트 개수로 추정 (기초자료: ~17개, 분석표: ~42개)
        if sheet_count <= 20:
            # 기초자료일 가능성
            if raw_matches >= 1:
                return 'raw'
        elif sheet_count >= 30:
            # 분석표일 가능성
            if analysis_matches >= 1:
                return 'analysis'
        
        # 5. 파일명으로 최종 추정
        filename = Path(filepath).stem.lower()
        if '기초' in filename or '수집' in filename or 'raw' in filename:
            return 'raw'
        elif '분석' in filename or 'analysis' in filename:
            return 'analysis'
        
        # 기본값: 분석표 (더 복잡한 처리가 필요하므로)
        print(f"[경고] 파일 유형을 명확히 판단할 수 없음: {filepath}")
        print(f"  - 시트 수: {sheet_count}")
        print(f"  - 기초자료 매칭: {raw_matches}, 분석표 매칭: {analysis_matches}")
        return 'unknown'
            
    except Exception as e:
        print(f"[오류] 파일 유형 감지 실패: {e}")
        return 'unknown'


def get_file_type_details(filepath: str) -> dict:
    """파일 유형과 상세 정보를 함께 반환
    
    Returns:
        {
            'type': 'raw' | 'analysis' | 'unknown',
            'sheet_count': int,
            'matched_raw_sheets': list,
            'matched_analysis_sheets': list,
            'confidence': 'high' | 'medium' | 'low',
            'reason': str
        }
    """
    try:
        xl = pd.ExcelFile(filepath)
        sheet_names = set(xl.sheet_names)
        sheet_count = len(sheet_names)
        
        # 기초자료 전용 시트
        RAW_ONLY_SHEETS = {
            '광공업생산', '서비스업생산', '소비(소매, 추가)', 
            '고용', '고용(kosis)', '고용률', '실업자 수',
            '지출목적별 물가', '품목성질별 물가', '건설 (공표자료)',
            '수출', '수입', '분기 GRDP', '완료체크'
        }
        
        # 분석표 전용 시트
        ANALYSIS_ONLY_SHEETS = {
            '본청', '시도별 현황', '이용관련',
            'A(광공업생산)집계', 'A 분석',
            'B(서비스업생산)집계', 'B 분석',
            "F'(건설)집계", "F'분석",
            'G(수출)집계', 'G 분석',
        }
        
        matched_raw = sorted(sheet_names & RAW_ONLY_SHEETS)
        matched_analysis = sorted(sheet_names & ANALYSIS_ONLY_SHEETS)
        
        result = {
            'sheet_count': sheet_count,
            'sheet_names': sorted(sheet_names),
            'matched_raw_sheets': matched_raw,
            'matched_analysis_sheets': matched_analysis,
        }
        
        # 판정
        if len(matched_analysis) >= 5:
            result['type'] = 'analysis'
            result['confidence'] = 'high'
            result['reason'] = f"분석표 전용 시트 {len(matched_analysis)}개 발견"
        elif len(matched_raw) >= 5:
            result['type'] = 'raw'
            result['confidence'] = 'high'
            result['reason'] = f"기초자료 전용 시트 {len(matched_raw)}개 발견"
        elif len(matched_analysis) >= 2:
            result['type'] = 'analysis'
            result['confidence'] = 'medium'
            result['reason'] = f"분석표 전용 시트 {len(matched_analysis)}개 발견"
        elif len(matched_raw) >= 2:
            result['type'] = 'raw'
            result['confidence'] = 'medium'
            result['reason'] = f"기초자료 전용 시트 {len(matched_raw)}개 발견"
        else:
            result['type'] = 'unknown'
            result['confidence'] = 'low'
            result['reason'] = "시트 구조를 명확히 판단할 수 없음"
        
        return result
        
    except Exception as e:
        return {
            'type': 'unknown',
            'confidence': 'low',
            'reason': f"파일 읽기 오류: {e}",
            'sheet_count': 0,
            'matched_raw_sheets': [],
            'matched_analysis_sheets': []
        }


def validate_sheet_structure(filepath: str, expected_type: str) -> dict:
    """파일의 시트 구조가 예상 유형과 일치하는지 검증
    
    Args:
        filepath: 엑셀 파일 경로
        expected_type: 'raw' 또는 'analysis'
        
    Returns:
        {
            'valid': bool,
            'missing_sheets': list,  # 필수 시트 중 누락된 것
            'extra_sheets': list,    # 예상치 못한 추가 시트
            'warnings': list
        }
    """
    try:
        xl = pd.ExcelFile(filepath)
        sheet_names = set(xl.sheet_names)
        
        result = {
            'valid': True,
            'missing_sheets': [],
            'extra_sheets': [],
            'warnings': []
        }
        
        if expected_type == 'raw':
            # 기초자료 필수 시트
            REQUIRED_RAW_SHEETS = {
                '광공업생산', '서비스업생산', '고용률', '실업자 수',
                '수출', '수입', '시도 간 이동'
            }
            OPTIONAL_RAW_SHEETS = {
                '소비(소매, 추가)', '지출목적별 물가', '품목성질별 물가',
                '건설 (공표자료)', '분기 GRDP', '완료체크',
                '고용', '고용(kosis)', '연령별 인구이동', '시군구인구이동'
            }
            
            missing = REQUIRED_RAW_SHEETS - sheet_names
            if missing:
                result['valid'] = False
                result['missing_sheets'] = sorted(missing)
            
            # 예상치 못한 시트 (분석표 시트가 섞여있는 경우)
            unexpected = [s for s in sheet_names if '분석' in s or '집계' in s]
            if unexpected:
                result['warnings'].append(f"기초자료에 분석표 시트가 포함됨: {unexpected}")
                
        elif expected_type == 'analysis':
            # 분석표 필수 시트
            REQUIRED_ANALYSIS_SHEETS = {
                'A(광공업생산)집계', 'A 분석',
                'B(서비스업생산)집계', 'B 분석',
                'C(소비)집계', 'C 분석',
                'D(고용률)집계', 'D(고용률)분석',
                '이용관련'
            }
            
            missing = REQUIRED_ANALYSIS_SHEETS - sheet_names
            if missing:
                result['valid'] = False
                result['missing_sheets'] = sorted(missing)
            
            # GRDP 시트 체크 (선택이지만 경고)
            grdp_sheets = {'GRDP', 'I GRDP', '분기 GRDP'}
            if not (sheet_names & grdp_sheets):
                result['warnings'].append("GRDP 시트가 없음 - 별도 업로드 필요")
        
        return result
        
    except Exception as e:
        return {
            'valid': False,
            'missing_sheets': [],
            'extra_sheets': [],
            'warnings': [f"파일 읽기 오류: {e}"]
        }

