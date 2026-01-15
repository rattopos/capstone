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


def comma(value):
    """천 단위 구분 기호(콤마) 추가"""
    if value is None:
        return ""
    try:
        # 숫자로 변환 시도
        if isinstance(value, str):
            # 문자열에서 숫자 부분만 추출
            import re
            num_str = re.sub(r'[^\d.-]', '', value)
            if not num_str:
                return value
            num_value = float(num_str)
        else:
            num_value = float(value)
        
        # 정수인지 소수인지 확인
        if num_value == int(num_value):
            return f"{int(num_value):,}"
        else:
            # 소수점이 있으면 소수점 이하 자리수 유지
            formatted = f"{num_value:,.2f}".rstrip('0').rstrip('.')
            # 천 단위 콤마 추가 (소수점 부분은 콤마 없이)
            parts = formatted.split('.')
            if len(parts) == 2:
                return f"{int(float(parts[0])):,}.{parts[1]}"
            else:
                return f"{int(num_value):,}"
    except (ValueError, TypeError):
        # 숫자 변환 실패 시 원본 반환
        return str(value)


def josa_eun_neun(word):
    """은/는 조사 자동 선택"""
    if not word or not isinstance(word, str):
        return "는"
    
    # 마지막 글자의 받침 유무 확인
    last_char = word[-1]
    last_code = ord(last_char)
    
    # 한글 유니코드 범위: AC00-D7A3
    if 0xAC00 <= last_code <= 0xD7A3:
        # 받침이 있으면 (종성 != 0)
        if (last_code - 0xAC00) % 28 != 0:
            return "은"
        else:
            return "는"
    else:
        # 한글이 아니면 기본값
        return "는"


def josa_i_ga(word):
    """이/가 조사 자동 선택"""
    if not word or not isinstance(word, str):
        return "가"
    
    # 마지막 글자의 받침 유무 확인
    last_char = word[-1]
    last_code = ord(last_char)
    
    # 한글 유니코드 범위: AC00-D7A3
    if 0xAC00 <= last_code <= 0xD7A3:
        # 받침이 있으면 (종성 != 0)
        if (last_code - 0xAC00) % 28 != 0:
            return "이"
        else:
            return "가"
    else:
        # 한글이 아니면 기본값
        return "가"


def josa_eul_reul(word):
    """을/를 조사 자동 선택"""
    if not word or not isinstance(word, str):
        return "를"
    
    # 마지막 글자의 받침 유무 확인
    last_char = word[-1]
    last_code = ord(last_char)
    
    # 한글 유니코드 범위: AC00-D7A3
    if 0xAC00 <= last_code <= 0xD7A3:
        # 받침이 있으면 (종성 != 0)
        if (last_code - 0xAC00) % 28 != 0:
            return "을"
        else:
            return "를"
    else:
        # 한글이 아니면 기본값
        return "를"


def register_filters(app):
    """Flask 앱에 Jinja2 필터 등록"""
    app.jinja_env.filters['is_missing'] = is_missing
    app.jinja_env.filters['format_value'] = format_value
    app.jinja_env.filters['editable'] = editable
    app.jinja_env.filters['comma'] = comma
    app.jinja_env.filters['josa_eun_neun'] = josa_eun_neun
    app.jinja_env.filters['josa_i_ga'] = josa_i_ga
    app.jinja_env.filters['josa_eul_reul'] = josa_eul_reul
    app.jinja_env.globals['is_missing'] = is_missing

