# -*- coding: utf-8 -*-
"""
Jinja2 커스텀 필터
"""

import pandas as pd


def is_missing(value):
    """결측치 여부 확인"""
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() in ['', '-', 'N/A', '없음', '미입력', '비공개']
    if isinstance(value, float) and pd.isna(value):
        return True
    return False


def format_value(value, format_str="%.1f", placeholder="[  ]"):
    """값 포맷팅 (결측치는 플레이스홀더로)"""
    if is_missing(value):
        return f'<span class="editable-placeholder" contenteditable="true">{placeholder}</span>'
    try:
        if format_str:
            return format_str % float(value)
        return str(value)
    except (ValueError, TypeError):
        return f'<span class="editable-placeholder" contenteditable="true">{placeholder}</span>'


def editable(value, format_str="%.1f"):
    """편집 가능한 값 표시 (결측치는 노란색 하이라이트)"""
    if is_missing(value):
        return f'<span class="editable-placeholder" contenteditable="true">[  ]</span>'
    try:
        formatted = format_str % float(value) if format_str else str(value)
        return formatted
    except (ValueError, TypeError):
        return f'<span class="editable-placeholder" contenteditable="true">[  ]</span>'


def safe_abs(value):
    """안전한 절대값 (None이면 N/A 반환)"""
    if is_missing(value):
        return None
    try:
        return abs(float(value))
    except (ValueError, TypeError):
        return None


def safe_format(value, format_str="%.1f", default="N/A"):
    """안전한 포맷팅 (None이면 N/A 반환)"""
    if is_missing(value):
        return default
    try:
        return format_str % float(value)
    except (ValueError, TypeError):
        return default


def val(value, format_str="%.1f", default="-"):
    """템플릿에서 안전하게 값 출력 (None이면 '-' 반환)
    
    템플릿에서 {{ val(value) }} 형태로 사용
    """
    if is_missing(value):
        return default
    try:
        return format_str % float(value)
    except (ValueError, TypeError):
        return str(value) if value else default


def register_filters(app):
    """Flask 앱에 Jinja2 필터 등록"""
    app.jinja_env.filters['is_missing'] = is_missing
    app.jinja_env.filters['format_value'] = format_value
    app.jinja_env.filters['editable'] = editable
    app.jinja_env.filters['safe_abs'] = safe_abs
    app.jinja_env.filters['safe_format'] = safe_format
    app.jinja_env.globals['is_missing'] = is_missing
    app.jinja_env.globals['safe_abs'] = safe_abs

