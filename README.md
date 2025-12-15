# 통계청 보도자료 자동화 시스템

11기 국가데이터처 캡스톤 프로젝트

## 개요

엑셀 파일의 데이터를 추출하여 계산을 수행하고, HTML 템플릿의 마커 위치에 정확히 삽입하여 보도자료를 자동 생성하는 시스템입니다.

## 기능

- HTML 템플릿 기반 보도자료 생성
- 엑셀 파일에서 데이터 자동 추출
- 다양한 계산 기능 지원 (합계, 평균, 증감률, 증감액 등)
- 자동 숫자 포맷팅 (천 단위 구분, 퍼센트 표시 등)

## 설치

```bash
pip install -r requirements.txt
```

## 사용 방법

### 기본 사용

기본 템플릿(`templates/mining_manufacturing_production.html`) 사용:

```bash
python -m src.main --excel data/data.xlsx --output output/result.html
```

다른 템플릿 파일 사용:

```bash
python -m src.main --template templates/template.html --excel data/data.xlsx --output output/result.html
```

### 옵션

- `--template, -t`: HTML 템플릿 파일 경로 (선택, 기본값: `templates/mining_manufacturing_production.html`)
- `--excel, -e`: 엑셀 데이터 파일 경로 (필수)
- `--output, -o`: 출력 파일 경로 (필수)
- `--verbose, -v`: 상세한 로그 출력

### 마커 형식

템플릿 HTML 파일에서 다음과 같은 형식으로 마커를 사용할 수 있습니다:

- `{시트명:셀주소}`: 단일 셀 값
  - 예: `{광공업생산:A1}`

- `{시트명:셀주소:계산식}`: 계산식 적용
  - 예: `{광공업생산:A1:A5:sum}` - 합계
  - 예: `{광공업생산:A1:A5:average}` - 평균
  - 예: `{광공업생산:A1:A2:증감률}` - 증감률 (퍼센트)
  - 예: `{광공업생산:A1:A2:증감액}` - 증감액

### 지원하는 계산식

- `sum`, `합계`: 합계
- `average`, `avg`, `평균`: 평균
- `max`, `최대값`, `최대`: 최대값
- `min`, `최소값`, `최소`: 최소값
- `growth_rate`, `증감률`, `증가율`: 증감률 (퍼센트)
- `growth_amount`, `증감액`, `증가액`: 증감액

## 프로젝트 구조

```
capstone/
├── templates/          # HTML 템플릿 파일
├── data/              # 엑셀 데이터 파일
├── output/            # 생성된 보도자료
├── src/               # 소스 코드
│   ├── main.py                 # 메인 실행 파일
│   ├── template_manager.py     # 템플릿 관리
│   ├── excel_extractor.py      # 엑셀 데이터 추출
│   ├── calculator.py           # 계산 엔진
│   └── template_filler.py      # 템플릿 채우기
├── requirements.txt   # Python 패키지 의존성
└── README.md
```

## 테스트

```bash
python src/test_basic.py
```

## 예시

템플릿 파일 예시 (`templates/sample_template.html` 참조):

```html
<h1>2025년 2분기 지역경제 보도자료</h1>
<p>총 인구: <span>{광공업생산:A1}</span>명</p>
<p>전년 대비 증감률: <span>{광공업생산:A1:A2:증감률}</span>%</p>
```

이 템플릿은 엑셀 파일의 '광공업생산' 시트에서 데이터를 추출하여 자동으로 채워집니다.
