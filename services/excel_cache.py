# -*- coding: utf-8 -*-
"""
엑셀 파일 캐싱 서비스

동일한 엑셀 파일을 여러 Generator가 읽을 때마다 다시 열지 않도록,
엑셀 객체(ExcelFile, Workbook)를 한 번만 로드하고 재사용합니다.
"""

import pandas as pd
from pathlib import Path
from typing import Optional, Dict, Any
import threading
from datetime import datetime


class ExcelCache:
    """엑셀 파일 캐싱 클래스 (Thread-safe)"""
    
    def __init__(self):
        self._cache: Dict[str, Dict[str, Any]] = {}
        self._lock = threading.Lock()
    
    def get_excel_file(self, excel_path: str, use_data_only: bool = True) -> Optional[pd.ExcelFile]:
        """
        엑셀 파일을 캐시에서 가져오거나 새로 로드
        
        Args:
            excel_path: 엑셀 파일 경로
            use_data_only: data_only=True로 읽을지 여부 (기본값: True, 계산된 값 읽기)
        
        Returns:
            ExcelFile 객체 또는 None
        """
        cache_key = f"{excel_path}:data_only={use_data_only}"
        
        with self._lock:
            # 캐시 확인
            if cache_key in self._cache:
                cache_entry = self._cache[cache_key]
                file_mtime = Path(excel_path).stat().st_mtime
                
                # 파일이 수정되지 않았으면 캐시된 객체 반환
                if cache_entry['mtime'] == file_mtime:
                    return cache_entry['xl']
                else:
                    # 파일이 수정되었으면 캐시 무효화
                    del self._cache[cache_key]
            
            # 캐시에 없거나 무효화된 경우 새로 로드
            try:
                xl = pd.ExcelFile(excel_path)
                
                # 캐시에 저장
                self._cache[cache_key] = {
                    'xl': xl,
                    'mtime': Path(excel_path).stat().st_mtime,
                    'timestamp': datetime.now()
                }
                
                return xl
            except Exception as e:
                print(f"[ExcelCache] 파일 로드 실패: {excel_path}, 오류: {e}")
                return None
    
    def get_openpyxl_workbook(self, excel_path: str, data_only: bool = True):
        """
        openpyxl Workbook을 캐시에서 가져오거나 새로 로드
        
        Args:
            excel_path: 엑셀 파일 경로
            data_only: data_only 옵션 (기본값: True)
        
        Returns:
            openpyxl Workbook 객체 또는 None
        """
        try:
            import openpyxl
        except ImportError:
            return None
        
        cache_key = f"{excel_path}:openpyxl:data_only={data_only}"
        
        with self._lock:
            # 캐시 확인
            if cache_key in self._cache:
                cache_entry = self._cache[cache_key]
                file_mtime = Path(excel_path).stat().st_mtime
                
                # 파일이 수정되지 않았으면 캐시된 객체 반환
                if cache_entry['mtime'] == file_mtime:
                    return cache_entry['wb']
                else:
                    # 파일이 수정되었으면 캐시 무효화
                    del self._cache[cache_key]
            
            # 캐시에 없거나 무효화된 경우 새로 로드
            try:
                wb = openpyxl.load_workbook(excel_path, data_only=data_only, read_only=True)
                
                # 캐시에 저장
                self._cache[cache_key] = {
                    'wb': wb,
                    'mtime': Path(excel_path).stat().st_mtime,
                    'timestamp': datetime.now()
                }
                
                return wb
            except Exception as e:
                print(f"[ExcelCache] openpyxl Workbook 로드 실패: {excel_path}, 오류: {e}")
                return None

    def get_calculated_path(self, excel_path: str) -> Optional[str]:
        """
        계산된 엑셀 파일 경로를 캐시에서 가져오기
        """
        cache_key = f"{excel_path}:calculated_path"

        with self._lock:
            cache_entry = self._cache.get(cache_key)
            if not cache_entry:
                return None

            try:
                file_mtime = Path(excel_path).stat().st_mtime
            except OSError:
                del self._cache[cache_key]
                return None

            calculated_path = cache_entry.get('calculated_path')
            if cache_entry.get('mtime') != file_mtime:
                del self._cache[cache_key]
                return None
            if not calculated_path or not Path(calculated_path).exists():
                del self._cache[cache_key]
                return None

            return calculated_path

    def set_calculated_path(self, excel_path: str, calculated_path: str):
        """
        계산된 엑셀 파일 경로를 캐시에 저장
        """
        cache_key = f"{excel_path}:calculated_path"

        try:
            file_mtime = Path(excel_path).stat().st_mtime
        except OSError:
            return

        with self._lock:
            self._cache[cache_key] = {
                'calculated_path': calculated_path,
                'mtime': file_mtime,
                'timestamp': datetime.now()
            }
    
    def clear_cache(self, excel_path: Optional[str] = None, preserve_calculated_path: bool = False):
        """
        캐시 정리
        
        Args:
            excel_path: 특정 파일의 캐시만 정리 (None이면 전체 정리)
            preserve_calculated_path: 계산된 파일 경로 캐시는 보존할지 여부
        """
        with self._lock:
            if excel_path:
                # 특정 파일의 모든 캐시 항목 제거
                keys_to_remove = [k for k in self._cache.keys() if k.startswith(excel_path)]
                if preserve_calculated_path:
                    keys_to_remove = [k for k in keys_to_remove if not k.endswith(':calculated_path')]
                for key in keys_to_remove:
                    try:
                        # openpyxl Workbook은 명시적으로 닫아야 함
                        entry = self._cache[key]
                        if 'wb' in entry:
                            entry['wb'].close()
                    except:
                        pass
                    del self._cache[key]
            else:
                # 전체 캐시 정리
                for entry in self._cache.values():
                    try:
                        if 'wb' in entry:
                            entry['wb'].close()
                    except:
                        pass
                self._cache.clear()
    
    def get_cache_info(self) -> Dict[str, Any]:
        """캐시 정보 반환"""
        with self._lock:
            return {
                'cache_size': len(self._cache),
                'cached_files': list(set(k.split(':')[0] for k in self._cache.keys()))
            }


# 전역 캐시 인스턴스
_excel_cache = ExcelCache()


def get_excel_file(excel_path: str, use_data_only: bool = True) -> Optional[pd.ExcelFile]:
    """전역 캐시에서 ExcelFile 가져오기"""
    return _excel_cache.get_excel_file(excel_path, use_data_only)


def get_openpyxl_workbook(excel_path: str, data_only: bool = True):
    """전역 캐시에서 openpyxl Workbook 가져오기"""
    return _excel_cache.get_openpyxl_workbook(excel_path, data_only)


def get_cached_calculated_path(excel_path: str) -> Optional[str]:
    """전역 캐시에서 계산된 엑셀 경로 가져오기"""
    return _excel_cache.get_calculated_path(excel_path)


def set_cached_calculated_path(excel_path: str, calculated_path: str):
    """전역 캐시에 계산된 엑셀 경로 저장"""
    _excel_cache.set_calculated_path(excel_path, calculated_path)


def clear_excel_cache(excel_path: Optional[str] = None, preserve_calculated_path: bool = False):
    """전역 캐시 정리"""
    _excel_cache.clear_cache(excel_path, preserve_calculated_path)


def get_cache_info() -> Dict[str, Any]:
    """캐시 정보 반환"""
    return _excel_cache.get_cache_info()
