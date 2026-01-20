# 자연어 나레이션 규칙화 및 구현 요약

## 1. 목적
- 지역경제 보도자료에서 **일관된 문장 패턴**과 **통제된 어휘**를 사용해 나레이션을 생성한다.
- 보고서 유형(물량/가격 등)에 따라 **증가·감소 표현**을 엄격히 구분한다.

---

## 2. 규칙화 개요
### 2.1 패턴 엔진 (4가지)
나레이션은 증감률과 전분기 방향, 업종 상반 여부에 따라 4가지 패턴 중 하나를 선택한다.

1) **패턴 C (보합)**
- 조건: $|growth\_rate| < 0.01$
- 예: “전년동분기대비 보합”

2) **패턴 D (방향 전환)**
- 조건: 전분기와 **부호가 반대**
- 예: “전분기 증가하였으나, 이번 분기 … 감소”

3) **패턴 B (역접)**
- 조건: 상반된 업종 혼재
- 예: “일부 업종은 줄었으나, 주요 업종이 늘어 … 증가”

4) **패턴 A (순접, 기본)**
- 위 조건에 해당하지 않으면 기본 패턴 사용

---

### 2.2 어휘 통제 규칙 (Master Vocabulary)
보고서 유형에 따라 **원인 서술어**와 **결과 서술어**를 엄격히 매핑한다.

- **물량(Quantity)**: 늘어/줄어 → 증가/감소
- **가격·비율(Price/Rate)**: 올라/내려 → 상승/하락

보합 시에는 원인 서술어를 제거하고 결과 서술어를 “보합”으로 고정한다.

---

### 2.3 조사/형태 처리
- 지역명 뒤 조사(은/는 등)는 **받침 유무**를 계산하여 자동 결정.
- 업종명이 비어있는 경우 문장 깨짐을 방지하기 위해 **제품 구문을 생략**한다.

---

## 3. 구현 구조
### 3.1 핵심 모듈
- **패턴 선택 및 기본 나레이션**: [templates/base_generator.py](../templates/base_generator.py#L1000)
- **어휘 통제/조사 처리**: [utils/text_utils.py](../utils/text_utils.py#L1)
- **시도별 나레이션**: [templates/unified_generator.py](../templates/unified_generator.py#L2620)

---

### 3.2 `get_terms()` 설명
`get_terms()`는 보고서 유형과 증감률 부호에 따라 **원인 서술어**와 **결과 서술어**를 반환하는 핵심 함수다.

- 위치: [utils/text_utils.py](../utils/text_utils.py#L54)
- 입력: `report_id`, `value`(증감률)
- 출력: `(원인 서술어 | None, 결과 서술어, 활용형)`

동작 규칙:
- `report_id`로 **물량/가격 타입**을 결정한다.
- 증감률이 양수면 증가/상승 계열, 음수면 감소/하락 계열을 반환한다.
- 보합($|value| < 0.01$)이면 원인 서술어를 비워 `(None, "보합", "")`를 반환한다.

---

### 3.2 주요 함수 흐름
1) `select_narrative_pattern()`
- 입력: 증감률, 전분기 증감률, 상반 업종 여부
- 출력: 패턴 A/B/C/D

2) `get_terms(report_id, value)`
- 입력: 보고서 ID, 증감률
- 출력: (원인 서술어, 결과 서술어, 활용형)
- 보합 시: `(None, "보합", "")`

3) `generate_narrative(...)`
- 선택된 패턴에 따라 **정형 문장** 생성
- 예: “{지역}{조사} {업종} 등이 {원인서술어} 전년동분기대비 {증감률}% {결과서술어}.”

4) 시도별 통합 나레이션 `_generate_narrative(...)`
- 동일 어휘 매핑을 적용하여 **지표별 템플릿**에 결합

---

## 4. 최근 개선 사항 (일관성 보강)
- 시도별 나레이션에도 `get_terms()`를 적용하여 **어휘 불일치 해소**
- 보합 구간에서 증가/감소 오판을 방지
- 제품명이 없을 때 “이/가”가 남는 문제를 제거

---

## 5. 예시
- 물량 데이터(증가):
  - “서울은 자동차, 반도체 등이 늘어 전년동분기대비 3.2% 증가.”

- 가격 데이터(하락):
  - “부산의 물가는 전년동기대비 0.8% 하락.”

- 보합:
  - “경기의 소비는 전년동기대비 보합.”

---

## 6. 보도자료 생성 프로세스 (순서별 설명 + 의사코드)

### 6.1 코드 비전공자를 위한 설명
아래는 실제 실행 흐름을 사람이 이해하기 쉬운 순서대로 정리한 것이다.

1) **진입점 선택**
- 웹 화면에서 실행하거나([app.py](../app.py), [dashboard.html](../dashboard.html)),
- 명령행에서 실행한다([report_generator.py](../report_generator.py)).

2) **보도자료 목록과 템플릿을 불러옴**
- 생성할 보도자료 목록은 [config/reports.py](../config/reports.py)에서 정의한다.
- 각 보도자료는 템플릿(HTML)과 생성기(클래스/함수)를 가진다.

3) **엑셀 데이터 로딩 및 캐시**
- 엑셀 파일을 읽고 필요한 시트/표를 추출한다.
- 반복 사용을 위해 캐시를 활용한다([services/excel_cache.py](../services/excel_cache.py)).

4) **데이터 가공 및 파생값 계산**
- 증감률, 전년동기 대비 변화, 상·하위 지역 등을 계산한다.
- 부문별 보고서는 통합 생성기([templates/unified_generator.py](../templates/unified_generator.py))를 통해 공통 로직을 공유한다.

5) **나레이션 생성**
- 공통 패턴과 어휘 규칙을 적용해 문장을 만든다.
- 규칙 엔진은 [templates/base_generator.py](../templates/base_generator.py), 어휘는 [utils/text_utils.py](../utils/text_utils.py)에서 관리된다.

6) **템플릿 렌더링**
- 정제된 데이터와 나레이션을 HTML 템플릿에 바인딩한다.
- 요약 보고서는 요약 데이터 빌더([services/summary_data.py](../services/summary_data.py))를 거쳐 스키마 기반으로 생성된다.

7) **파일 저장 및 다운로드**
- 생성된 HTML은 exports 폴더 등에 저장되고 웹/CLI에서 다운로드 가능하다.

---

### 6.2 파이썬 이해자를 위한 의사코드
```
START
  entry = WEB or CLI
  excel_path = 업로드/인자에서 받음

  # 1) 보고서 정의 로드
  report_list = config/reports.py 의 REPORT_ORDER

  # 2) 엑셀 로딩/캐싱
  workbook = services/excel_cache.py 로드

  FOR each report in report_list (부문별 → 시도별 → 요약 순)
    IF report.uses_functions:
       data = 함수형 generator에서 테이블/요약 데이터 추출
    ELSE:
       data = 클래스형 generator에서 데이터 추출

    data = 파생값 계산(증감률, 상하위 지역, etc)

    narrative = base_generator 패턴 + text_utils 어휘로 생성

    html = template 렌더링(Jinja2)

    저장(exports/*.html)
  END FOR

  RETURN 결과 목록(성공/실패)
END
```
