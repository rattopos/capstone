# -*- coding: utf-8 -*-
"""
데이터 처리 유틸리티 함수
"""

import pandas as pd


def check_missing_data(data, report_id):
    """보고서 생성에 필수적인 결측치만 확인"""
    missing_fields = []
    
    # 보고서별 필수 필드 정의
    REQUIRED_FIELDS = {
        'manufacturing': [],
        'service': [],
        'consumption': [],
        'employment': [],
        'unemployment': [],
        'price': [],
        'export': [],
        'import': [],
        'population': [],
    }
    
    def get_nested_value(obj, path):
        """중첩된 경로에서 값 가져오기"""
        keys = path.replace('[', '.').replace(']', '').split('.')
        current = obj
        for key in keys:
            if current is None:
                return None
            if isinstance(current, dict):
                current = current.get(key)
            elif isinstance(current, list) and key.isdigit():
                idx = int(key)
                current = current[idx] if idx < len(current) else None
            else:
                return None
        return current
    
    def is_missing(value):
        """값이 결측치인지 확인"""
        if value is None:
            return True
        if value == '':
            return True
        if isinstance(value, float) and pd.isna(value):
            return True
        return False
    
    # 해당 보고서의 필수 필드만 확인
    required = REQUIRED_FIELDS.get(report_id, [])
    for field_path in required:
        value = get_nested_value(data, field_path)
        if is_missing(value):
            missing_fields.append(field_path)
    
    return missing_fields

