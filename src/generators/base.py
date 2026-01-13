# -*- coding: utf-8 -*-
"""
Base Generator Abstract Class
모든 보도자료 생성기의 기본 클래스
"""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Dict, Any, Optional, Tuple
import pandas as pd


class BaseGenerator(ABC):
    """보도자료 생성기 기본 추상 클래스"""
    
    def __init__(
        self,
        excel_path: str,
        raw_excel_path: Optional[str] = None,
        year: Optional[int] = None,
        quarter: Optional[int] = None
    ):
        """
        생성기 초기화
        
        Args:
            excel_path: 분석 엑셀 파일 경로
            raw_excel_path: 기초자료 엑셀 파일 경로 (선택적)
            year: 연도 (선택적)
            quarter: 분기 (선택적)
        """
        self.excel_path = Path(excel_path)
        self.raw_excel_path = Path(raw_excel_path) if raw_excel_path else None
        self.year = year
        self.quarter = quarter
        self._data: Optional[Dict[str, Any]] = None
    
    @abstractmethod
    def generate(self) -> Dict[str, Any]:
        """
        보도자료 데이터 생성 (추상 메서드)
        
        Returns:
            생성된 데이터 딕셔너리
        """
        pass
    
    def render_html(
        self,
        template_path: str,
        output_path: Optional[str] = None
    ) -> str:
        """
        HTML 보도자료 렌더링
        
        Args:
            template_path: Jinja2 템플릿 파일 경로
            output_path: 출력 HTML 파일 경로 (선택적)
            
        Returns:
            렌더링된 HTML 문자열
        """
        from jinja2 import Environment, FileSystemLoader
        
        # 데이터 생성
        if self._data is None:
            self._data = self.generate()
        
        # Jinja2 환경 설정
        template_dir = Path(template_path).parent
        env = Environment(loader=FileSystemLoader(str(template_dir)))
        template = env.get_template(Path(template_path).name)
        
        # 렌더링
        html_content = template.render(**self._data)
        
        # 파일 저장
        if output_path:
            output_file = Path(output_path)
            output_file.parent.mkdir(parents=True, exist_ok=True)
            output_file.write_text(html_content, encoding='utf-8')
            print(f"보도자료가 생성되었습니다: {output_path}")
        
        return html_content
    
    def export_data_json(self, output_path: str) -> None:
        """
        추출된 데이터를 JSON으로 내보내기
        
        Args:
            output_path: 출력 JSON 파일 경로
        """
        import json
        
        if self._data is None:
            self._data = self.generate()
        
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        output_file.write_text(
            json.dumps(self._data, ensure_ascii=False, indent=2),
            encoding='utf-8'
        )
        print(f"데이터가 저장되었습니다: {output_path}")
    
    @staticmethod
    def safe_float(value: Any, default: Optional[float] = None) -> Optional[float]:
        """
        안전한 float 변환 함수 (NaN, '-' 체크 포함)
        
        Args:
            value: 변환할 값
            default: 기본값 (변환 실패 시)
            
        Returns:
            float 값 또는 기본값
        """
        if value is None:
            return default
        try:
            if pd.isna(value):
                return default
            if isinstance(value, str):
                value = value.strip()
                if value == '-' or value == '':
                    return default
            result = float(value)
            if pd.isna(result):
                return default
            return result
        except (ValueError, TypeError):
            return default
    
    @staticmethod
    def safe_round(value: Any, decimals: int = 1, default: Optional[float] = None) -> Optional[float]:
        """
        안전한 반올림 함수
        
        Args:
            value: 반올림할 값
            decimals: 소수점 자릿수
            default: 기본값 (변환 실패 시)
            
        Returns:
            반올림된 값 또는 기본값
        """
        result = BaseGenerator.safe_float(value, default)
        if result is None:
            return default
        return round(result, decimals)
    
    @staticmethod
    def find_sheet_with_fallback(
        xl: pd.ExcelFile,
        primary_sheets: list[str],
        fallback_sheets: list[str]
    ) -> Tuple[Optional[str], bool]:
        """
        시트 찾기 - 기본 시트가 없으면 대체 시트 사용
        
        Args:
            xl: pandas ExcelFile 객체
            primary_sheets: 우선 시트 이름 리스트
            fallback_sheets: 대체 시트 이름 리스트
            
        Returns:
            (시트 이름, 대체 사용 여부) 튜플
        """
        for sheet in primary_sheets:
            if sheet in xl.sheet_names:
                return sheet, False
        for sheet in fallback_sheets:
            if sheet in xl.sheet_names:
                print(f"[시트 대체] '{primary_sheets[0]}' → '{sheet}' (기초자료)")
                return sheet, True
        return None, False
