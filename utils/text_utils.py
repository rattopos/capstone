# -*- coding: utf-8 -*-
"""
텍스트 유틸리티 모듈
나레이션 생성 및 텍스트 포맷팅 함수 제공
어휘 통제 센터: 보고서 종류별 단어 사용 통제
"""

from typing import List, Dict, Optional, Tuple

# ============================================================================
# 어휘 통제 센터: 보고서 유형별 어휘 상수 (Master Vocabulary)
# ============================================================================

# 보고서 유형별 어휘 세트
NARRATIVE_MAP = {
    # 1. 양적 데이터 (Quantity): 늘어/줄어 -> 증가/감소
    'quantity': {
        'cause_up': "늘어",
        'cause_down': "줄어",
        'result_up': "증가",
        'result_down': "감소",
        'conjunction_up': "하였으나",  # 증가했을 때 활용형
        'conjunction_down': "하였으나",  # 감소했을 때 활용형
        'comparative_rate_low': "증가율이 낮았으나",  # 전국 대비 낮음 (증가 시)
        'comparative_rate_high': "높음",  # 전국 대비 높음 (증가 시)
    },
    # 2. 지수/비율 데이터 (Index/Rate): 올라/내려 -> 상승/하락
    'price': {
        'cause_up': "올라",
        'cause_down': "내려",
        'result_up': "상승",
        'result_down': "하락",
        'conjunction_up': "하였으나",  # 상승했을 때 활용형
        'conjunction_down': "하였으나",  # 하락했을 때 활용형
        'comparative_rate_low': "상승률이 낮았으나",  # 전국 대비 낮음 (상승 시)
        'comparative_rate_high': "높음",  # 전국 대비 높음 (상승 시)
    }
}

# 리포트 ID별 타입 매핑
REPORT_TYPE_MAP = {
    # [Quantity 그룹] - 양적 데이터 (늘어/줄어 -> 증가/감소)
    'manufacturing': 'quantity',           # 광공업생산
    'mining': 'quantity',                  # 광공업 (별칭)
    'service': 'quantity',                 # 서비스업생산
    'consumption': 'quantity',             # 소비동향
    'construction': 'quantity',            # 건설
    'export': 'quantity',                 # 수출
    'import': 'quantity',                  # 수입
    'migration': 'quantity',                # 인구이동
    'population': 'quantity',               # 인구이동 (별칭, 하위 호환성)
    'domestic_migration': 'quantity',       # 인구이동 (별칭)
    'employment_count': 'quantity',        # 취업자수
    
    # [Price 그룹] - 지수/비율 데이터 (올라/내려 -> 상승/하락)
    'price': 'price',                      # 물가동향
    'employment': 'price',                 # 고용률
    'employment_rate': 'price',            # 고용률 (별칭)
    'unemployment': 'price',               # 실업률
    'unemployment_rate': 'price',          # 실업률 (별칭)
}

# ============================================================================
# 푸터 통제 센터: 보고서별 자료 출처 매핑 (PDF 기준)
# ============================================================================

# 보고서별 자료 출처 매핑 (PDF 기준)
SOURCE_MAP = {
    'manufacturing': "광업제조업동향조사",
    'mining': "광업제조업동향조사",           # 광공업 (별칭)
    'service': "서비스업동향조사",
    'consumption': "서비스업동향조사(소매판매)",
    'construction': "건설경기동향조사",
    'export': "무역통계",                    # 관세청 등 출처 확인 필요 시 수정
    'import': "무역통계",
    'price': "소비자물가조사",
    'employment': "경제활동인구조사",
    'employment_rate': "경제활동인구조사",    # 고용률 (별칭)
    'unemployment': "경제활동인구조사",
    'unemployment_rate': "경제활동인구조사",  # 실업률 (별칭)
    'migration': "국내인구이동통계",
    'population': "국내인구이동통계",        # 인구이동 (별칭, 하위 호환성)
    'domestic_migration': "국내인구이동통계", # 인구이동 (별칭)
    'regional': "각 통계 원천 참조"          # 시도별은 통합이므로 예외 처리
}


def get_footer_source(report_id: str) -> str:
    """
    report_id에 맞는 조사명을 포함한 전체 출처 문자열 반환
    
    Args:
        report_id: 보고서 ID (예: 'mining', 'price', 'employment')
    
    Returns:
        str: 전체 출처 문자열
        예: "자료: 국가데이터처 국가통계포털(KOSIS), 광업제조업동향조사"
        예: "자료: 국가데이터처 국가통계포털(KOSIS), 소비자물가조사"
    """
    survey_name = SOURCE_MAP.get(report_id, "지역경제동향")
    return f"자료: 국가데이터처 국가통계포털(KOSIS), {survey_name}"


def get_terms(report_id: str, value: float) -> Tuple[Optional[str], str, str]:
    """
    report_id와 값(부호)을 넣으면 알맞은 단어 세트를 반환
    
    [문서 1] 엄격한 어휘 매핑 규칙 준수:
    - Type A (물량): 증가/감소, 늘어/줄어
    - Type B (가격): 상승/하락, 올라/내려
    
    Args:
        report_id: 보고서 ID (예: 'manufacturing', 'price', 'employment')
        value: 증감률 값 (양수면 증가, 음수면 감소, 0.0이면 보합)
    
    Returns:
        Tuple[Optional[str], str, str]: (원인 서술어 or None, 결과 서술어, 활용형)
        예: get_terms('manufacturing', 1.5) -> ('늘어', '증가', '하였으나')
        예: get_terms('manufacturing', -0.5) -> ('줄어', '감소', '하였으나')
        예: get_terms('manufacturing', 0.0) -> (None, '보합', '')
        예: get_terms('price', 2.1) -> ('올라', '상승', '하였으나')
        예: get_terms('price', -1.5) -> ('내려', '하락', '하였으나')
        예: get_terms('price', 0.0) -> (None, '보합', '')
    
    Note:
        - 알 수 없는 report_id는 기본값 'quantity' 사용
        - 보합일 때 원인 서술어는 None (문장 구조가 다름)
        - 활용형은 결과 서술어 뒤에 붙는 형태 (예: "증가하였으나", "감소하였으나")
    """
    # 1. 타입 결정 (기본값은 quantity)
    n_type = REPORT_TYPE_MAP.get(report_id, 'quantity')
    vocabs = NARRATIVE_MAP[n_type]
    
    # 2. 보합 처리 (최우선)
    if abs(value) < 0.01:  # 0.0%로 간주
        return (None, "보합", "")
    
    # 3. 증감 처리
    if value > 0:
        return vocabs['cause_up'], vocabs['result_up'], vocabs['conjunction_up']
    else:
        return vocabs['cause_down'], vocabs['result_down'], vocabs['conjunction_down']


def get_comparative_terms(report_id: str, nationwide_direction: int) -> Tuple[str, str]:
    """
    전국 방향에 따른 지역별 비교 표현 반환 (물가동향 등에서 사용)
    
    Args:
        report_id: 보고서 ID (예: 'price')
        nationwide_direction: 전국의 방향 (양수: 상승, 음수: 하락)
    
    Returns:
        Tuple[str, str]: (낮은 지역 표현, 높은 지역 표현)
        예: get_comparative_terms('price', 1) -> ('상승률이 낮았으나', '높음')
        예: get_comparative_terms('price', -1) -> ('하락률이 낮았으나', '높음')
    
    Note:
        - 전국이 상승일 때: "상승률이 낮았으나" / "높음"
        - 전국이 하락일 때: "하락률이 낮았으나" / "높음"
    """
    n_type = REPORT_TYPE_MAP.get(report_id, 'price')
    vocabs = NARRATIVE_MAP[n_type]
    
    if nationwide_direction >= 0:
        # 상승/증가 시
        result_verb = vocabs['result_up']  # "상승" 또는 "증가"
        low_expression = f"{result_verb}률이 낮았으나"
        high_expression = "높음"
    else:
        # 하락/감소 시
        result_verb = vocabs['result_down']  # "하락" 또는 "감소"
        low_expression = f"{result_verb}률이 낮았으나"
        high_expression = "높음"
    
    return low_expression, high_expression


def get_cause_verb(value: float, report_id: str = 'quantity') -> str:
    """
    값이 양수면 '늘어' 또는 '올라', 음수면 '줄어' 또는 '내려'를 반환
    0일 경우 '보합세를 보여' 반환
    
    Args:
        value: 증감률 값 (양수면 증가, 음수면 감소)
        report_id: 보고서 ID (기본값: 'quantity')
    
    Returns:
        str: 서술어 ('늘어', '줄어', '올라', '내려', '보합세를 보여')
    
    Deprecated:
        이 함수는 하위 호환성을 위해 유지됩니다.
        새로운 코드는 get_terms() 함수를 사용하세요.
    """
    if value == 0:
        return "보합세를 보여"
    
    cause_verb, _ = get_terms(report_id, value)
    return cause_verb


def get_josa(word: str, josa_pair: str = "은/는") -> str:
    """
    [Robust Dynamic Parsing System - 4단계]
    word의 마지막 글자 받침 유무에 따라 조사 반환
    
    Args:
        word: 조사를 붙일 단어 (예: '서울', '부산', '경기')
        josa_pair: 조사 쌍 (기본값: "은/는")
            - "은/는": 주제 조사
            - "이/가": 주어 조사
            - "을/를": 목적어 조사
            - "와/과": 접속 조사
    
    Returns:
        str: 적절한 조사
        예: get_josa('서울', '은/는') -> '은'
        예: get_josa('경기', '은/는') -> '는'
        예: get_josa('서울', '이/가') -> '이'
        예: get_josa('경기', '이/가') -> '가'
    
    Note:
        - 한글 유니코드 범위: 0xAC00 ~ 0xD7A3 (44032 ~ 55203)
        - 받침 계산 공식: (코드 - 0xAC00) % 28 > 0 이면 받침 있음
        - 받침 있으면: 첫 번째 조사 (은, 이, 을, 와)
        - 받침 없으면: 두 번째 조사 (는, 가, 를, 과)
    """
    # 타입 체크 및 문자열 변환 (템플릿에서 None이나 다른 타입이 올 수 있음)
    if word is None:
        return ""
    
    # 문자열이 아니면 문자열로 변환 시도
    if not isinstance(word, str):
        try:
            word = str(word).strip()
        except:
            return ""
    
    # 빈 문자열 체크
    if not word:
        return ""
    
    # 호환성: 기존 type 파라미터 지원
    if josa_pair in ["Topic", "Subject"]:
        if josa_pair == "Topic":
            josa_pair = "은/는"
        elif josa_pair == "Subject":
            josa_pair = "이/가"
    
    # 마지막 글자 추출
    last_char = word[-1]
    
    # 한글 범위 체크 및 받침 계산
    try:
        last_code = ord(last_char)
        if 0xAC00 <= last_code <= 0xD7A3:
            # 한글 유니코드 공식: (코드 - 0xAC00) % 28 > 0 이면 받침 있음
            has_batchim = (last_code - 0xAC00) % 28 > 0
        else:
            # 한글 아니면(숫자, 영어 등) 기본값(받침 없음 가정)
            has_batchim = False
    except (TypeError, ValueError):
        # ord() 실패 시 기본값(받침 없음)
        has_batchim = False
    
    # 조사 쌍 분리
    try:
        first, second = josa_pair.split("/")
        return first if has_batchim else second
    except (ValueError, AttributeError):
        # 조사 쌍 형식이 잘못되었을 경우 기본값 반환
        return "는"


def get_contrast_narrative(nationwide_val: float, inc_regions: List[Dict], dec_regions: List[Dict], report_id: str = 'quantity') -> str:
    """
    지그재그 화법을 사용한 대조 나레이션 생성
    
    논리 구조:
    1. 전국 증가(+)인 경우: 전국 증가 -> 감소 지역(-, 역접) -> 증가 지역(+, 결론)
       => "전국은 ~ 증가. [감소지역]은 ~ 감소하였으나, [증가지역]은 ~ 증가"
       
    2. 전국 감소(-)인 경우: 전국 감소 -> 증가 지역(+, 역접) -> 감소 지역(-, 결론)
       => "전국은 ~ 감소. [증가지역]은 ~ 증가하였으나, [감소지역]은 ~ 감소"
    
    Args:
        nationwide_val: 전국 증감률 값 (양수면 증가, 음수면 감소)
        inc_regions: 증가 지역 리스트 (각 항목은 {'name': str, 'value': float} 형식)
        dec_regions: 감소 지역 리스트 (각 항목은 {'name': str, 'value': float} 형식)
        report_id: 보고서 ID (기본값: 'quantity')
    
    Returns:
        str: 생성된 나레이션 텍스트
    """
    
    # 어휘 세트 가져오기
    inc_cause, inc_result = get_terms(report_id, 1.0)  # 증가 지역용 (양수)
    dec_cause, dec_result = get_terms(report_id, -1.0)  # 감소 지역용 (음수)
    
    # 1. 지역 텍스트 포맷팅 (리스트가 비었을 때 방어 로직 포함)
    if inc_regions:
        if len(inc_regions) == 1:
            region_name = inc_regions[0]['name']
            josa = get_josa(region_name, "Topic")
            inc_text = f"{region_name}{josa}({inc_regions[0]['value']}%)"
        else:
            # 여러 지역인 경우 첫 번째 지역만 표시하거나 모두 표시
            inc_names = []
            for r in inc_regions[:3]:
                region_name = r['name']
                josa = get_josa(region_name, "Topic")
                inc_names.append(f"{region_name}{josa}({r['value']}%)")
            inc_text = ", ".join(inc_names)
    else:
        inc_text = "일부 지역"
    
    if dec_regions:
        if len(dec_regions) == 1:
            region_name = dec_regions[0]['name']
            josa = get_josa(region_name, "Topic")
            dec_text = f"{region_name}{josa}({dec_regions[0]['value']}%)"
        else:
            # 여러 지역인 경우 첫 번째 지역만 표시하거나 모두 표시
            dec_names = []
            for r in dec_regions[:3]:
                region_name = r['name']
                josa = get_josa(region_name, "Topic")
                dec_names.append(f"{region_name}{josa}({r['value']}%)")
            dec_text = ", ".join(dec_names)
    else:
        dec_text = "일부 지역"
    
    # 2. 전국 방향에 따른 지그재그 패턴 결정
    if nationwide_val > 0:
        # 패턴: 전국(+) -> 감소 지역(-) -> 증가 지역(+)
        # 예: "서울은 감소하였으나, 경기는 증가" (quantity)
        # 예: "서울은 하락하였으나, 경기는 상승" (price)
        return f"{dec_text} {dec_result}하였으나, {inc_text} {inc_cause} {inc_result}"
    else:
        # 패턴: 전국(-) -> 증가 지역(+) -> 감소 지역(-)
        # 예: "울산은 증가하였으나, 제주는 감소" (quantity)
        # 예: "울산은 상승하였으나, 제주는 하락" (price)
        return f"{inc_text} {inc_result}하였으나, {dec_text} {dec_cause} {dec_result}"
