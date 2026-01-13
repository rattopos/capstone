# -*- coding: utf-8 -*-
"""
애플리케이션 설정 클래스
포트, 경로, 보도자료 순서 등 설정 관리
"""

from pathlib import Path
from typing import List, Dict, Any
from .reports import REPORT_ORDER, SUMMARY_REPORTS, SECTOR_REPORTS, STATISTICS_REPORTS, REGIONAL_REPORTS
from .settings import UPLOAD_FOLDER, EXPORT_FOLDER, BASE_DIR


class Config:
    """애플리케이션 설정 클래스"""
    
    # 포트 설정
    DEFAULT_PORT: int = 5050
    
    # 디렉토리 설정
    UPLOAD_FOLDER: Path = UPLOAD_FOLDER
    OUTPUT_FOLDER: Path = BASE_DIR / 'output'
    TEMPLATES_DIR: Path = BASE_DIR / 'templates'
    SCHEMAS_DIR: Path = BASE_DIR / 'src' / 'schemas'
    GENERATORS_DIR: Path = BASE_DIR / 'src' / 'generators'
    
    # 보도자료 순서 설정
    REPORT_ORDER: List[Dict[str, Any]] = REPORT_ORDER
    SUMMARY_REPORTS: List[Dict[str, Any]] = SUMMARY_REPORTS
    SECTOR_REPORTS: List[Dict[str, Any]] = SECTOR_REPORTS
    STATISTICS_REPORTS: List[Dict[str, Any]] = STATISTICS_REPORTS
    REGIONAL_REPORTS: List[Dict[str, Any]] = REGIONAL_REPORTS
    
    @classmethod
    def init_directories(cls) -> None:
        """필요한 디렉토리 생성"""
        cls.OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
        cls.UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
        cls.TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
        cls.SCHEMAS_DIR.mkdir(parents=True, exist_ok=True)
        cls.GENERATORS_DIR.mkdir(parents=True, exist_ok=True)
