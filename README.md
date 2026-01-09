# 지역경제동향 보도자료 자동생성

11기 국가데이터처 캡스톤 프로젝트

## 개요

엑셀 파일의 데이터를 추출하여 계산을 수행하고, HTML 템플릿의 마커 위치에 정확히 삽입하여 보도자료를 자동 생성하는 시스템입니다. Flask 기반 웹 인터페이스를 통해 손쉽게 사용할 수 있으며, PDF 및 DOCX 형식으로도 출력할 수 있습니다.

## 주요 기능

- **10개 부문별 보도자료 템플릿 지원**
  - 광공업생산, 서비스업생산, 소매판매, 고용률, 실업률
  - 물가동향, 건설수주, 수출, 수입, 국내인구이동
- **다양한 출력 형식**: HTML, PDF, DOCX
- **엑셀 데이터 자동 추출 및 매핑**
- **스키마 기반 유연한 데이터 처리**
- **결측치 감지 및 사용자 입력 지원**
- **자동 계산 기능** (합계, 평균, 증감률, 증감액 등)
- **숫자 포맷팅** (천 단위 구분, 퍼센트 표시 등)
- **지역별 순위 자동 계산**

## 설치

### 1. 가상환경 생성 (권장)

```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
```

### 2. 패키지 설치

```bash
pip install -r requirements.txt
```

### 3. 의존성 패키지

- `pandas`: 데이터 처리
- `openpyxl`: 엑셀 파일 처리
- `beautifulsoup4`, `lxml`: HTML 파싱
- `flask`: 웹 서버
- `python-docx`: DOCX 생성
- `easyocr`, `opencv-python`: OCR (템플릿 생성용)

## 사용 방법

### 웹 애플리케이션 (권장)

웹 브라우저를 통해 쉽게 사용할 수 있는 웹 인터페이스를 제공합니다.

1. **웹 서버 시작**:
```bash
python app.py
```

2. **브라우저에서 접속**:
   - http://localhost:8000 을 열어주세요

3. **사용 방법**:
   - 연도 및 분기 선택
   - 템플릿 선택 (광공업생산, 고용률 등)
   - 엑셀 파일 업로드 (선택사항 - 기본 파일 사용 가능)
   - "보도자료 생성" 버튼 클릭
   - 생성된 결과를 미리보기하거나 다운로드

4. **PDF/DOCX 일괄 생성**:
   - "PDF 생성" 또는 "DOCX 생성" 탭 선택
   - 연도/분기 입력 후 생성 버튼 클릭
   - 10개 템플릿이 순서대로 처리되어 하나의 파일로 생성

### CLI 사용 (명령줄)

기본 템플릿(`templates/광공업생산.html`) 사용:

```bash
python -m src.main --excel data/data.xlsx --output output/result.html
```

다른 템플릿 파일 사용:

```bash
python -m src.main --template templates/고용률.html --excel data/data.xlsx --output output/result.html
```

### CLI 옵션

- `--template, -t`: HTML 템플릿 파일 경로 (선택, 기본값: `templates/광공업생산.html`)
- `--excel, -e`: 엑셀 데이터 파일 경로 (필수)
- `--output, -o`: 출력 파일 경로 (필수)
- `--verbose, -v`: 상세한 로그 출력

## 지원 템플릿

| 템플릿명 | 파일명 | 필요 시트 |
|---------|--------|----------|
| 광공업생산 | `광공업생산.html` | 광공업생산 |
| 서비스업생산 | `서비스업생산.html` | 서비스업생산 |
| 소매판매 | `소매판매.html` | 소비(소매, 추가) |
| 고용률 | `고용률.html` | 고용, 고용률 |
| 실업률 | `실업률.html` | 실업자 수, 실업률 |
| 물가동향 | `물가동향.html` | 품목성질별 물가, 지출목적별 물가 |
| 건설수주 | `건설수주.html` | 건설 (공표자료) |
| 수출 | `수출.html` | 수출 |
| 수입 | `수입.html` | 수입 |
| 국내인구이동 | `국내인구이동.html` | 시도 간 이동, 연령별 인구이동 |

## 마커 형식

템플릿 HTML 파일에서 다음과 같은 형식으로 마커를 사용할 수 있습니다:

### 기본 마커

- `{시트명:셀주소}`: 단일 셀 값
  - 예: `{광공업생산:A1}`

### 계산 마커

- `{시트명:셀주소:셀주소:계산식}`: 두 셀 간 계산
  - 예: `{광공업생산:A1:A2:증감률}` - 증감률 계산
  - 예: `{광공업생산:A1:A2:증감액}` - 증감액 계산

### 범위 계산 마커

- `{시트명:셀주소:셀주소:sum}` - 합계
- `{시트명:셀주소:셀주소:average}` - 평균

### 지역 순위 마커

- `{시트명:지역:카테고리:high_1}` - 가장 높은 1위 지역
- `{시트명:지역:카테고리:low_1}` - 가장 낮은 1위 지역

### 지원하는 계산식

| 계산식 | 설명 |
|-------|------|
| `sum`, `합계` | 합계 |
| `average`, `avg`, `평균` | 평균 |
| `max`, `최대값`, `최대` | 최대값 |
| `min`, `최소값`, `최소` | 최소값 |
| `growth_rate`, `증감률`, `증가율` | 증감률 (퍼센트) |
| `growth_amount`, `증감액`, `증가액` | 증감액 |

## 프로젝트 구조

```
capstone/
├── app.py                      # Flask 웹 애플리케이션 메인
├── requirements.txt            # Python 패키지 의존성
├── README.md                   # 프로젝트 문서
│
├── templates/                  # HTML 보도자료 템플릿
│   ├── 광공업생산.html
│   ├── 서비스업생산.html
│   ├── 소매판매.html
│   ├── 고용률.html
│   ├── 실업률.html
│   ├── 물가동향.html
│   ├── 건설수주.html
│   ├── 수출.html
│   ├── 수입.html
│   └── 국내인구이동.html
│
├── flask_templates/            # Flask 웹 페이지 템플릿
│   └── index.html
│
├── static/                     # 정적 파일 (CSS, JS)
│   ├── css/
│   │   └── style.css
│   └── js/
│       └── main.js
│
├── schemas/                    # 데이터 스키마 정의
│   ├── template_mapping.json   # 시트-템플릿 매핑
│   ├── name_mappings.json      # 이름 매핑
│   ├── output_formats/         # 출력 형식 스키마
│   └── sheets/                 # 시트별 설정
│
├── routes/                     # Flask Blueprint 라우트
│   ├── __init__.py
│   ├── templates.py            # 템플릿 관련 API
│   ├── processing.py           # 데이터 처리 API
│   ├── export.py               # 내보내기 API
│   └── validation.py           # 검증 API
│
├── src/                        # 핵심 소스 코드
│   ├── main.py                 # CLI 메인 실행 파일
│   ├── template_manager.py     # 템플릿 관리
│   ├── template_filler.py      # 템플릿 데이터 채우기
│   ├── excel_extractor.py      # 엑셀 데이터 추출
│   ├── calculator.py           # 계산 엔진
│   ├── data_analyzer.py        # 데이터 분석
│   ├── period_detector.py      # 연도/분기 감지
│   ├── schema_loader.py        # 스키마 로드
│   ├── flexible_mapper.py      # 유연한 데이터 매핑
│   ├── dynamic_sheet_parser.py # 동적 시트 파싱
│   ├── pdf_generator.py        # PDF 생성
│   ├── docx_generator.py       # DOCX 생성
│   ├── template_generator.py   # 템플릿 생성 (OCR)
│   ├── excel_header_parser.py  # 엑셀 헤더 파싱
│   │
│   ├── markers/                # 마커 처리 모듈
│   │   ├── base.py             # 기본 마커 처리
│   │   ├── statistics.py       # 통계 마커
│   │   ├── region_ranking.py   # 지역 순위 마커
│   │   ├── national.py         # 전국 데이터 마커
│   │   ├── unemployment.py     # 실업률 마커
│   │   └── dynamic_processor.py# 동적 마커 처리
│   │
│   ├── analyzers/              # 분석 모듈
│   │   ├── data_analyzer.py
│   │   ├── dynamic_sheet_parser.py
│   │   ├── excel_header_parser.py
│   │   └── period_detector.py
│   │
│   ├── core/                   # 핵심 모듈
│   │   ├── excel_extractor.py
│   │   ├── schema_loader.py
│   │   ├── template_filler.py
│   │   └── template_manager.py
│   │
│   ├── generators/             # 생성기 모듈
│   │   ├── base_generator.py
│   │   ├── docx_generator.py
│   │   ├── pdf_generator.py
│   │   └── template_generator.py
│   │
│   └── utils/                  # 유틸리티
│       ├── formatters.py       # 포맷팅 유틸리티
│       ├── region_utils.py     # 지역 관련 유틸리티
│       └── sheet_utils.py      # 시트 관련 유틸리티
│
├── output/                     # 생성된 보도자료 출력
├── test_output/                # 테스트 출력
└── correct_answer/             # 정답 이미지 (검증용)
    ├── 부문별/
    └── 시도별/
```

## API 엔드포인트

### 템플릿 관련

| 엔드포인트 | 메서드 | 설명 |
|-----------|--------|------|
| `/api/templates` | GET | 사용 가능한 템플릿 목록 조회 |
| `/api/template-sheets` | POST | 템플릿이 필요로 하는 시트 목록 조회 |
| `/api/validate-template` | POST | 템플릿 마커 검증 |

### 처리 관련

| 엔드포인트 | 메서드 | 설명 |
|-----------|--------|------|
| `/api/process` | POST | 엑셀 + 템플릿 처리하여 HTML 생성 |
| `/api/generate-pdf` | POST | 10개 템플릿으로 PDF 일괄 생성 |
| `/api/generate-docx` | POST | 10개 템플릿으로 DOCX 일괄 생성 |

### 검증 관련

| 엔드포인트 | 메서드 | 설명 |
|-----------|--------|------|
| `/api/validate-excel` | POST | 엑셀 파일 유효성 검증 |
| `/api/check-structure` | POST | 엑셀 시트 구조 검증 |
| `/api/check-missing-values` | POST | 결측치 확인 |
| `/api/check-default-file` | GET | 기본 엑셀 파일 존재 확인 |

### 파일 관련

| 엔드포인트 | 메서드 | 설명 |
|-----------|--------|------|
| `/api/download/<filename>` | GET | 생성된 파일 다운로드 |
| `/api/preview/<filename>` | GET | 생성된 파일 미리보기 |

## 스키마 시스템

### template_mapping.json

시트명과 템플릿 파일 간의 매핑을 정의합니다:

```json
{
  "광공업생산": {
    "template": "광공업생산.html",
    "display_name": "광공업생산",
    "output_format": "광공업생산"
  }
}
```

### sheets/ 폴더

각 시트별 데이터 구조와 열 매핑을 정의합니다:

- 지역 열, 카테고리 열 위치
- 분기별 데이터 열 패턴
- 특수 처리 규칙

### output_formats/ 폴더

출력 형식별 설정을 정의합니다:

- 숫자 포맷 (소수점, 천 단위 등)
- 순위 표시 형식
- 특수 마커 처리 규칙

## 테스트

```bash
python src/test_basic.py
```

전체 템플릿 테스트:

```bash
python test_all_templates.py
```

## 예시

템플릿 파일 예시:

```html
<h1>2025년 2분기 지역경제 보도자료</h1>
<p>광공업생산지수: <span>{광공업생산:전국:전산업:current}</span></p>
<p>전년 동분기 대비 증감률: <span>{광공업생산:전국:전산업:yoy_rate}</span>%</p>
<p>가장 높은 지역: <span>{광공업생산:지역:전산업:high_1}</span></p>
```

이 템플릿은 엑셀 파일의 '광공업생산' 시트에서 데이터를 추출하여 자동으로 채워집니다.

## 주의사항

- 엑셀 파일은 읽기 전용으로 처리됩니다 (수정 불가)
- 최대 업로드 파일 크기: 100MB
- 지원 파일 형식: `.xlsx`, `.xls`

## 라이선스

이 프로젝트는 교육 목적으로 제작되었습니다.
