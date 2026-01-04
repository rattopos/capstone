# 지역경제동향 보도자료 생성 시스템 - 종합 이해 가이드

> 이 문서는 프로젝트의 모든 측면을 깊이 있게 이해하고, 어떤 질문에도 대비할 수 있도록 작성되었습니다.

---

## 목차

1. [프로젝트 개요](#1-프로젝트-개요)
2. [시스템 아키텍처](#2-시스템-아키텍처)
3. [기술 스택](#3-기술-스택)
4. [데이터 플로우](#4-데이터-플로우)
5. [프로젝트 구조](#5-프로젝트-구조)
6. [핵심 컴포넌트 상세](#6-핵심-컴포넌트-상세)
7. [보도자료 생성 프로세스](#7-보도자료-생성-프로세스)
8. [API 엔드포인트](#8-api-엔드포인트)
9. [데이터 구조 및 형식](#9-데이터-구조-및-형식)
10. [주요 기능 상세](#10-주요-기능-상세)
11. [설정 및 환경](#11-설정-및-환경)
12. [배포 및 운영](#12-배포-및-운영)
13. [문제 해결 가이드](#13-문제-해결-가이드)

---

## 1. 프로젝트 개요

### 1.1 프로젝트 목적

**지역경제동향 보도자료 생성 시스템**은 국가데이터처(구 통계청)에서 분기별로 발행하는 지역경제동향 보도자료를 자동으로 생성하는 웹 애플리케이션입니다.

### 1.2 핵심 가치

- **시간 절감**: 보도자료 초안 작성 시간을 5일에서 약 3분으로 단축 (99.9% 절감)
- **업무 효율화**: 반복 작업 자동화로 해석 작업 시간 3배 확보 (2일 → 6일)
- **정확성 향상**: 수동 입력 오류 방지 및 데이터 일관성 보장
- **실무자 만족도 향상**: 단순 반복 작업 해방으로 업무 피로도 감소

### 1.3 지원 보도자료 유형

#### 1.3.1 요약 보도자료 (Summary Reports)
- 표지 (Cover)
- 일러두기 (Guide)
- 목차 (Table of Contents)
- 인포그래픽 (Infographic)
- 요약-지역경제동향 (Summary Overview)
- 요약-생산 (Summary Production)
- 요약-소비건설 (Summary Consumption & Construction)
- 요약-수출물가 (Summary Export & Price)
- 요약-고용인구 (Summary Employment & Population)

#### 1.3.2 부문별 보도자료 (Sectoral Reports)
- 광공업생산 (Manufacturing)
- 서비스업생산 (Service Industry)
- 소비동향 (Consumption)
- 건설동향 (Construction)
- 수출 (Export)
- 수입 (Import)
- 물가동향 (Price Trend)
- 고용률 (Employment Rate)
- 실업률 (Unemployment Rate)
- 국내인구이동 (Domestic Migration)

#### 1.3.3 시도별 보도자료 (Regional Reports)
- 17개 시도별 개별 보도자료 (서울, 부산, 대구, 인천, 광주, 대전, 울산, 세종, 경기, 강원, 충북, 충남, 전북, 전남, 경북, 경남, 제주)
- 참고_GRDP (Reference GRDP)

#### 1.3.4 통계표 보도자료 (Statistics Reports)
- 통계표 목차
- 개별 통계표 (광공업생산지수, 서비스업생산지수, 소매판매액지수, 건설수주액, 고용률, 실업률, 국내인구이동, 수출액, 수입액, 소비자물가지수)
- 통계표-참고-GRDP
- 부록-주요용어정의

### 1.4 주요 사용자

- **실무팀**: 담당 사무관 1명, 주무관 1명 (총 2명)
- **작업 주기**: 분기별 (1년에 4회)
- **기존 작업 시간**: 약 7일
- **개선 후 작업 시간**: 약 3분 (생성) + 편집/검토 시간

---

## 2. 시스템 아키텍처

### 2.1 전체 구조

```
┌─────────────────────────────────────────────────────────────┐
│                    웹 브라우저 (클라이언트)                    │
│  - dashboard.html (단일 페이지 애플리케이션)                  │
│  - JavaScript (Fetch API, 상태 관리)                         │
└───────────────────┬─────────────────────────────────────────┘
                    │ HTTP/HTTPS
┌───────────────────▼─────────────────────────────────────────┐
│              Flask 웹 서버 (Python)                          │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  routes/        (라우팅 계층)                         │   │
│  │  - main.py      (메인 페이지, 파일 서빙)               │   │
│  │  - api.py       (REST API 엔드포인트)                 │   │
│  │  - preview.py   (미리보기 엔드포인트)                  │   │
│  │  - debug.py     (디버그 엔드포인트)                    │   │
│  └──────────────┬───────────────────────────────────────┘   │
│  ┌──────────────▼───────────────────────────────────────┐   │
│  │  services/     (비즈니스 로직 계층)                    │   │
│  │  - report_generator.py  (보도자료 생성)                │   │
│  │  - excel_processor.py   (엑셀 전처리)                  │   │
│  │  - grdp_service.py      (GRDP 데이터 처리)             │   │
│  │  - summary_data.py      (요약 데이터 추출)              │   │
│  └──────────────┬───────────────────────────────────────┘   │
│  ┌──────────────▼───────────────────────────────────────┐   │
│  │  templates/    (데이터 추출 및 템플릿)                  │   │
│  │  - *_generator.py  (데이터 추출기)                     │   │
│  │  - *_template.html (Jinja2 템플릿)                    │   │
│  └──────────────┬───────────────────────────────────────┘   │
│  ┌──────────────▼───────────────────────────────────────┐   │
│  │  utils/        (유틸리티 함수)                         │   │
│  │  - filters.py        (Jinja2 필터)                    │   │
│  │  - excel_utils.py    (엑셀 처리)                      │   │
│  │  - data_utils.py     (데이터 검증)                    │   │
│  └──────────────────────────────────────────────────────┘   │
└───────────────────┬─────────────────────────────────────────┘
                    │
┌───────────────────▼─────────────────────────────────────────┐
│              데이터 계층                                      │
│  - uploads/      (업로드된 파일)                             │
│  - templates/    (생성된 HTML)                              │
│  - exports/      (내보내기 파일)                             │
│  - Excel 파일    (기초자료 수집표 / 분석표)                   │
└─────────────────────────────────────────────────────────────┘
```

### 2.2 레이어 구조

1. **프레젠테이션 레이어** (dashboard.html)
   - 사용자 인터페이스
   - 상태 관리 (JavaScript)
   - API 호출 및 응답 처리

2. **라우팅 레이어** (routes/)
   - HTTP 요청 라우팅
   - 세션 관리
   - 파일 업로드/다운로드 처리

3. **비즈니스 로직 레이어** (services/)
   - 보도자료 생성 오케스트레이션
   - 데이터 전처리
   - GRDP 데이터 처리

4. **데이터 추출 레이어** (templates/*_generator.py)
   - 엑셀 데이터 추출
   - 데이터 변환 및 가공
   - 템플릿 데이터 준비

5. **템플릿 레이어** (templates/*_template.html)
   - Jinja2 템플릿
   - HTML 렌더링
   - 데이터 바인딩

6. **유틸리티 레이어** (utils/)
   - 공통 함수
   - 필터 및 헬퍼
   - 데이터 검증

### 2.3 데이터 흐름 패턴

#### 패턴 1: 파일 업로드 → 데이터 추출 → HTML 생성
```
엑셀 파일 업로드
    ↓
파일 유형 감지 (기초자료 vs 분석표)
    ↓
데이터 전처리 (수식 계산 등)
    ↓
Generator를 통한 데이터 추출
    ↓
Jinja2 템플릿 렌더링
    ↓
HTML 반환
```

#### 패턴 2: 기초자료 → 분석표 변환 → 보도자료 생성
```
기초자료 수집표 업로드
    ↓
DataConverter를 통한 분석표 생성
    ↓
분석표 저장
    ↓
보도자료 생성 프로세스 실행
```

---

## 3. 기술 스택

### 3.1 백엔드

- **Python 3.10+**: 메인 프로그래밍 언어
- **Flask 2.3.0+**: 웹 프레임워크
  - Blueprint 기반 모듈화된 라우팅
  - 세션 관리
  - 파일 업로드/다운로드 처리
- **Jinja2 3.1.0+**: 템플릿 엔진
  - 커스텀 필터 지원
  - 템플릿 상속 및 포함
- **Werkzeug 2.3.0+**: WSGI 유틸리티 (Flask 의존성)

### 3.2 데이터 처리

- **pandas 2.0.0+**: 데이터 분석 및 처리
  - 엑셀 파일 읽기/쓰기
  - DataFrame 조작
  - 데이터 변환
- **openpyxl 3.1.0+**: 엑셀 파일 직접 조작
  - 시트 읽기/쓰기
  - 셀 수정
  - 수식 처리
- **numpy 1.24.0+**: 수치 연산
- **xlwings 0.30.0+**: Excel 애플리케이션 연동 (수식 계산용, 선택사항)
- **formulas 1.2.0+**: 순수 Python 수식 계산 (선택사항)

### 3.3 프론트엔드

- **HTML5/CSS3**: 마크업 및 스타일링
- **Vanilla JavaScript (ES6+)**: 클라이언트 사이드 로직
  - Fetch API (비동기 통신)
  - DOM 조작
  - 상태 관리
- **반응형 디자인**: CSS Grid, Flexbox

### 3.4 기타 도구

- **BeautifulSoup4 4.12.0+**: HTML 파싱 및 조작
- **Pillow 10.0.0+**: 이미지 처리
- **python-dateutil 2.8.0+**: 날짜 처리
- **requests 2.31.0+**: HTTP 요청 (KOSIS API 등)
- **python-dotenv 1.0.0+**: 환경 변수 관리

### 3.5 개발 도구

- **Git**: 버전 관리
- **Playwright 1.40.0+**: 데모 녹화 (선택사항)

---

## 4. 데이터 플로우

### 4.1 전체 데이터 플로우도

```
[기초자료 수집표] 또는 [분석표]
        ↓
    파일 업로드
        ↓
    파일 유형 감지
        ├─→ 분석표 → 바로 보도자료 생성
        └─→ 기초자료 → 분석표 변환 → 보도자료 생성
        ↓
    GRDP 데이터 확인
        ├─→ 있음 → 보도자료 생성
        ├─→ 없음 → GRDP 업로드 요청
        └─→ 기본값 사용
        ↓
    보도자료별 데이터 추출
        ├─→ Generator 실행
        ├─→ 데이터 가공
        └─→ 템플릿 데이터 준비
        ↓
    Jinja2 템플릿 렌더링
        ↓
    HTML 생성
        ↓
    브라우저에서 미리보기 / 최종 생성
        ↓
    내보내기 (PDF, XLSX, 한글용 HTML)
```

### 4.2 기초자료 → 분석표 변환 플로우

```
기초자료 수집표 엑셀 파일
        ↓
DataConverter 초기화
        ↓
연도/분기 자동 감지
        ↓
분석표 템플릿 로드
        ↓
시트별 변환 처리:
    ├─→ 시트 매핑 확인 (예: '광공업생산' → 'A(광공업생산)집계')
    ├─→ 열 구조 분석
    ├─→ 데이터 추출 및 변환
    ├─→ 집계 시트에 데이터 복사
    └─→ 메타데이터 보존
        ↓
가중치 처리 (광공업생산, 서비스업생산)
        ↓
분석 시트 수식 계산 (선택사항)
        ↓
분석표 엑셀 파일 저장
```

### 4.3 보도자료 생성 플로우

```
보도자료 ID + 엑셀 파일 경로
        ↓
Generator 모듈 동적 로드
        ├─→ 클래스 기반 Generator (대부분)
        └─→ 함수 기반 Generator (고용률, 실업률)
        ↓
데이터 추출
    ├─→ 엑셀 파일 읽기
    ├─→ 특정 시트 접근
    ├─→ 데이터 파싱
    ├─→ 계산 및 가공
    └─→ 딕셔너리 형태로 구조화
        ↓
커스텀 데이터 병합 (결측치 대체 등)
        ↓
결측치 검증
        ↓
Jinja2 템플릿 로드
        ↓
템플릿 렌더링 (데이터 바인딩)
        ↓
HTML 문자열 반환
```

### 4.4 GRDP 데이터 처리 플로우

```
GRDP 데이터 필요 시
        ↓
데이터 소스 우선순위:
    1. 세션에 저장된 GRDP 데이터
    2. JSON 파일 (grdp_extracted.json)
    3. 기초자료에서 추출
    4. 분석표의 GRDP 시트
    5. 업로드된 KOSIS GRDP 파일
    6. 기본값 (플레이스홀더)
        ↓
데이터 파싱 및 검증
        ↓
세션 및 JSON 파일에 저장
        ↓
보도자료 생성 시 사용
```

---

## 5. 프로젝트 구조

### 5.1 디렉토리 구조

```
capstone/
├── app.py                          # Flask 애플리케이션 진입점
├── dashboard.html                  # 메인 대시보드 UI (단일 파일)
├── report_generator.py             # 통합 보도자료 생성기 (CLI용)
├── data_converter.py               # 기초자료 → 분석표 변환기
│
├── config/                         # 설정 파일
│   ├── __init__.py
│   ├── settings.py                 # 기본 설정 (경로, 상수)
│   └── reports.py                  # 보도자료 정의 및 순서
│
├── routes/                         # Flask 라우팅
│   ├── __init__.py
│   ├── main.py                     # 메인 페이지 라우트
│   ├── api.py                      # REST API 엔드포인트
│   ├── preview.py                  # 미리보기 라우트
│   └── debug.py                    # 디버그 라우트
│
├── services/                       # 비즈니스 로직
│   ├── __init__.py
│   ├── report_generator.py         # 보도자료 생성 서비스
│   ├── excel_processor.py          # 엑셀 전처리 서비스
│   ├── grdp_service.py             # GRDP 데이터 처리 서비스
│   └── summary_data.py             # 요약 데이터 추출 서비스
│
├── utils/                          # 유틸리티 함수
│   ├── __init__.py
│   ├── filters.py                  # Jinja2 커스텀 필터
│   ├── excel_utils.py              # 엑셀 관련 유틸리티
│   └── data_utils.py               # 데이터 검증 유틸리티
│
├── templates/                      # 템플릿 및 생성기
│   ├── *_generator.py              # 데이터 추출기 (각 보도자료별)
│   ├── *_template.html             # Jinja2 템플릿 (각 보도자료별)
│   ├── *_schema.json               # 데이터 스키마 (일부 보도자료)
│   └── grdp_extracted.json         # 추출된 GRDP 데이터 (런타임 생성)
│
├── uploads/                        # 업로드된 파일 저장소
├── debug/                          # 디버그용 HTML 파일
├── exports/                        # 내보내기 파일 (한글용 HTML 등)
│
├── docs/                           # 문서
│   ├── PROJECT_COMPREHENSIVE_GUIDE.md  # 이 문서
│   ├── DEBUG_LOG.md                # 디버그 로그
│   ├── DEPLOYMENT_GUIDE.md         # 배포 가이드
│   └── ...
│
├── correct_answer/                 # 참고용 정답 이미지
├── requirements.txt                # Python 의존성
└── README.md                       # 프로젝트 개요
```

### 5.2 주요 파일 역할

#### 5.2.1 애플리케이션 진입점

- **app.py**: Flask 애플리케이션 팩토리 및 설정
  - Flask 인스턴스 생성
  - Blueprint 등록
  - 필터 등록
  - 서버 실행

#### 5.2.2 설정 파일

- **config/settings.py**: 
  - BASE_DIR, TEMPLATES_DIR 등 경로 설정
  - UPLOAD_FOLDER, DEBUG_FOLDER, EXPORT_FOLDER
  - Flask 설정 (SECRET_KEY, MAX_CONTENT_LENGTH)
  - Vercel 환경 감지 및 디렉토리 설정

- **config/reports.py**:
  - REPORT_ORDER: 전체 보도자료 순서
  - SUMMARY_REPORTS: 요약 보도자료 목록
  - SECTOR_REPORTS: 부문별 보도자료 목록
  - REGIONAL_REPORTS: 시도별 보도자료 목록
  - STATISTICS_REPORTS: 통계표 보도자료 목록
  - PAGE_CONFIG: 페이지 번호 설정
  - TOC_SECTOR_ITEMS, TOC_REGION_ITEMS: 목차 항목

#### 5.2.3 라우팅 파일

- **routes/main.py**: 
  - `/`: 대시보드 페이지
  - `/uploads/<filename>`: 업로드 파일 다운로드
  - `/view/<filename>`: 파일 직접 보기
  - `/exports/<path:filepath>`: 내보내기 파일 보기
  - `/templates/<filename>`: 템플릿 파일 서빙

- **routes/api.py**:
  - `/api/upload`: 파일 업로드
  - `/api/generate-preview`: 미리보기 생성
  - `/api/generate-all`: 전체 보도자료 생성
  - `/api/upload-grdp`: GRDP 파일 업로드
  - `/api/export-final`: 최종 문서 내보내기
  - `/api/export-xlsx`: XLSX 형식 내보내기
  - `/api/export-hwp-ready`: 한글 복붙용 HTML 생성

- **routes/preview.py**: 미리보기 전용 라우트

- **routes/debug.py**: 디버그용 라우트

#### 5.2.4 서비스 파일

- **services/report_generator.py**:
  - `generate_report_html()`: 개별 보도자료 HTML 생성
  - `generate_regional_report_html()`: 시도별 보도자료 생성
  - `generate_statistics_report_html()`: 통계표 보도자료 생성
  - Generator 모듈 로드 및 실행 오케스트레이션

- **services/excel_processor.py**:
  - `preprocess_excel()`: 엑셀 수식 계산
  - `_try_openpyxl_calculation()`: openpyxl로 직접 계산
  - `_try_formulas()`: formulas 라이브러리 사용
  - `_try_xlwings()`: xlwings로 Excel 앱 사용

- **services/grdp_service.py**:
  - `parse_kosis_grdp_file()`: KOSIS GRDP 파일 파싱
  - `get_default_grdp_data()`: 기본 GRDP 데이터 생성
  - `save_extracted_contributions()`: 기여율 데이터 저장

- **services/summary_data.py**:
  - `get_summary_overview_data()`: 요약-지역경제동향 데이터
  - `get_summary_production_data()`: 요약-생산 데이터
  - 각종 요약 데이터 추출 함수

#### 5.2.5 유틸리티 파일

- **utils/filters.py**:
  - `is_missing()`: 결측치 확인 필터
  - `format_value()`: 값 포맷팅 필터
  - `register_filters()`: Flask에 필터 등록

- **utils/excel_utils.py**:
  - `load_generator_module()`: Generator 모듈 동적 로드
  - `extract_year_quarter_from_excel()`: 연도/분기 추출
  - `detect_file_type()`: 파일 유형 감지

- **utils/data_utils.py**:
  - `check_missing_data()`: 데이터 결측치 검증

#### 5.2.6 데이터 변환 파일

- **data_converter.py**:
  - `DataConverter` 클래스: 기초자료 → 분석표 변환
  - 시트 매핑 및 데이터 변환
  - 가중치 처리
  - GRDP 데이터 추출

### 5.3 템플릿 구조

각 보도자료는 다음 파일들로 구성됩니다:

1. **Generator 파일** (`*_generator.py`):
   - 데이터 추출 로직
   - 엑셀 파일에서 데이터 읽기
   - 데이터 가공 및 변환
   - 딕셔너리 형태로 반환

2. **Template 파일** (`*_template.html`):
   - Jinja2 템플릿
   - HTML 구조 정의
   - 데이터 바인딩
   - 스타일 포함

3. **Schema 파일** (`*_schema.json`, 선택사항):
   - 데이터 구조 정의
   - 기본값 제공 (일러두기 등)

## 6. 핵심 컴포넌트 상세

### 6.1 Flask 애플리케이션 (app.py)

Flask 애플리케이션의 진입점으로, 모든 설정과 Blueprint 등록을 담당합니다.

**주요 기능:**
- Flask 인스턴스 생성 및 설정
- Blueprint 등록 (main, api, preview, debug)
- Jinja2 커스텀 필터 등록
- 서버 실행 (포트 5050)

**설정 항목:**
- `SECRET_KEY`: 세션 암호화 키
- `UPLOAD_FOLDER`: 업로드 파일 저장 경로
- `MAX_CONTENT_LENGTH`: 최대 업로드 파일 크기 (50MB)
- `template_folder`: 템플릿 디렉토리 (BASE_DIR)
- `static_folder`: 정적 파일 디렉토리 (BASE_DIR)

### 6.2 데이터 변환기 (data_converter.py)

기초자료 수집표를 분석표 형식으로 변환하는 핵심 모듈입니다.

#### 6.2.1 DataConverter 클래스

**초기화:**
- `raw_excel_path`: 기초자료 수집표 파일 경로
- `template_path`: 분석표 템플릿 경로 (기본값: 프로젝트 내 템플릿)
- `year`, `quarter`: 연도/분기 (자동 감지)

**주요 메서드:**

1. **`_detect_year_quarter()`**: 
   - 기초자료 헤더 또는 파일명에서 연도/분기 추출
   - 여러 패턴 시도 후 기본값 사용

2. **`convert_all()`**:
   - 모든 시트를 분석표 형식으로 변환
   - 시트 매핑: `SHEET_MAPPING` 딕셔너리 사용
   - 가중치 설정 지원 (광공업생산, 서비스업생산)

3. **`convert_sheet()`**:
   - 개별 시트 변환
   - 열 구조 분석 (`SHEET_STRUCTURE`)
   - 데이터 추출 및 변환
   - 집계 시트에 데이터 복사

4. **`extract_grdp_data()`**:
   - 기초자료에서 GRDP 데이터 추출
   - 성장률 및 기여율 계산

**시트 매핑 구조:**
```python
SHEET_MAPPING = {
    '광공업생산': 'A(광공업생산)집계',
    '서비스업생산': 'B(서비스업생산)집계',
    '소비(소매, 추가)': 'C(소비)집계',
    # ... 기타 시트
}
```

**열 구조 정의:**
각 집계 시트별로 다음 정보를 정의:
- `meta_start`: 메타데이터 시작 열 (1-based)
- `raw_meta_cols`: 기초자료에서 메타데이터 열 개수
- `year_start`: 연도 데이터 시작 열
- `quarter_start`: 분기 데이터 시작 열
- `weight_col`: 가중치 열 위치 (선택사항)

### 6.3 보도자료 생성 서비스 (services/report_generator.py)

보도자료 HTML 생성의 핵심 오케스트레이션을 담당합니다.

#### 6.3.1 generate_report_html()

**파라미터:**
- `excel_path`: 엑셀 파일 경로
- `report_config`: 보도자료 설정 (id, name, generator, template 등)
- `year`, `quarter`: 연도/분기
- `custom_data`: 커스텀 데이터 (결측치 대체용)
- `raw_excel_path`: 기초자료 경로 (선택사항)

**처리 과정:**

1. **Generator 모듈 로드**:
   - `load_generator_module()` 호출
   - 동적 모듈 임포트

2. **데이터 추출 방식 결정**:
   - `generate_report_data()` 함수 우선
   - `generate_report()` 함수 (레거시)
   - Generator 클래스 사용 (최후 수단)

3. **데이터 후처리**:
   - Top3 지역 데이터 정규화
   - 커스텀 데이터 병합
   - 결측치 검증

4. **템플릿 렌더링**:
   - Jinja2 템플릿 로드
   - 필터 등록
   - HTML 생성

**반환값:**
- `(html_content, error, missing_fields)`
  - `html_content`: 생성된 HTML 문자열
  - `error`: 오류 메시지 (성공 시 None)
  - `missing_fields`: 결측치 필드 리스트

#### 6.3.2 generate_regional_report_html()

시도별 보도자료 생성 전용 함수입니다.

**특수 처리:**
- `RegionalGenerator` 클래스 사용
- 지역명을 파라미터로 전달
- 참고_GRDP의 경우 별도 처리

#### 6.3.3 generate_statistics_report_html()

통계표 보도자료 생성 함수입니다.

**특수 처리:**
- `통계표Generator` 클래스 사용
- 여러 통계표를 하나의 HTML로 통합
- 페이지 번호 자동 계산

### 6.4 엑셀 전처리 서비스 (services/excel_processor.py)

엑셀 파일의 수식을 계산하여 값을 채우는 전처리 서비스입니다.

#### 6.4.1 preprocess_excel()

**처리 우선순위:**

1. **openpyxl 직접 계산** (가장 빠름):
   - 분석 시트의 수식이 집계 시트를 참조하는 경우
   - `='시트이름'!셀주소` 패턴 파싱
   - 백엔드에서 직접 값 매핑
   - 수식 보존 (엑셀에서 열면 자동 계산)

2. **formulas 라이브러리** (순수 Python):
   - 복잡한 수식 지원
   - Excel 앱 불필요

3. **xlwings** (Fallback):
   - Excel 앱 실행 필요
   - 가장 느리지만 가장 정확

**반환값:**
- `(processed_path, success, message)`
  - `processed_path`: 처리된 파일 경로
  - `success`: 성공 여부
  - `message`: 처리 방법 설명

#### 6.4.2 check_available_methods()

사용 가능한 수식 계산 방법을 확인합니다.

**확인 항목:**
- `xlwings`: xlwings 설치 여부
- `formulas`: formulas 설치 여부
- `openpyxl`: openpyxl 설치 여부
- `excel_installed`: Excel 앱 설치 여부 (Mac/Windows)

### 6.5 GRDP 서비스 (services/grdp_service.py)

GRDP (분기 지역내총생산) 데이터 처리 전담 서비스입니다.

#### 6.5.1 parse_kosis_grdp_file()

KOSIS에서 다운로드한 GRDP 엑셀 파일을 파싱합니다.

**파싱 과정:**

1. **시트 확인**:
   - '성장률' 시트 존재 여부
   - '기여도' 시트 존재 여부

2. **데이터 추출**:
   - 시트별 데이터 추출 함수 호출
   - 또는 실질금액에서 계산

3. **데이터 구조화**:
   ```python
   {
       'report_info': {'year': ..., 'quarter': ...},
       'national_summary': {
           'growth_rate': ...,
           'direction': '증가' or '감소',
           'contributions': {...}
       },
       'top_region': {...},
       'regional_data': [...]
   }
   ```

#### 6.5.2 get_default_grdp_data()

기본 GRDP 데이터 (플레이스홀더)를 생성합니다.

**용도:**
- GRDP 데이터가 없을 때 임시 사용
- 사용자에게 기여율 수정 요청

### 6.6 Generator 패턴

각 보도자료는 Generator 파일을 통해 데이터를 추출합니다.

#### 6.6.1 클래스 기반 Generator

**구조:**
```python
class MiningManufacturingGenerator:
    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.df = None
    
    def load_data(self):
        # 엑셀 파일 로드
        pass
    
    def extract_all_data(self) -> dict:
        # 모든 데이터 추출
        return {
            'report_info': {...},
            'national_data': {...},
            'regional_data': {...},
            # ...
        }
```

**주요 메서드:**
- `load_data()`: 엑셀 파일 로드
- `extract_all_data()`: 전체 데이터 추출
- 개별 데이터 추출 메서드들

#### 6.6.2 함수 기반 Generator (고용률, 실업률)

**구조:**
```python
def load_data(excel_path: str):
    # 데이터 로드
    return df_analysis, df_index

def get_nationwide_data(df_analysis, df_index):
    # 전국 데이터 추출
    return {...}

def get_regional_data(df_analysis, df_index):
    # 지역 데이터 추출
    return {...}

def get_summary_box_data(regional_data):
    # 요약 박스 데이터
    return {...}

def get_table_data(df_analysis, df_index):
    # 표 데이터
    return [...]
```

#### 6.6.3 generate_report_data() 함수 패턴

일부 Generator는 `generate_report_data()` 함수를 제공합니다.

**시그니처:**
```python
def generate_report_data(excel_path: str, raw_excel_path: str = None, 
                        year: int = None, quarter: int = None) -> dict:
    # 데이터 추출 및 반환
    return {...}
```

### 6.7 Jinja2 템플릿

각 보도자료는 Jinja2 템플릿을 사용하여 HTML을 생성합니다.

**템플릿 구조:**
- HTML 구조 정의
- Jinja2 문법으로 데이터 바인딩
- 인라인 CSS 스타일
- Chart.js 등 JavaScript 라이브러리 사용 (일부)

**주요 Jinja2 기능 사용:**
- 변수: `{{ variable }}`
- 제어문: `{% if %}`, `{% for %}`, `{% macro %}`
- 필터: `{{ value | format_value }}`, `{{ value | is_missing }}`
- 포함: `{% include 'partial.html' %}`

**커스텀 필터:**
- `is_missing`: 결측치 확인
- `format_value`: 값 포맷팅 (결측치는 플레이스홀더)
- `editable`: 편집 가능한 값 표시

### 6.8 대시보드 (dashboard.html)

단일 HTML 파일로 구현된 클라이언트 사이드 애플리케이션입니다.

#### 6.8.1 상태 관리

JavaScript 객체로 상태를 관리합니다:

```javascript
const state = {
    fileUploaded: false,
    allGenerated: false,
    currentReportIndex: -1,
    allReports: [],
    reports: [],
    summaryReports: [],
    sectoralReports: [],
    regionalReports: [],
    statisticsReports: [],
    // ...
};
```

#### 6.8.2 주요 기능

1. **파일 업로드**:
   - 드래그 앤 드롭 지원
   - 파일 선택 다이얼로그
   - 업로드 진행 상태 표시

2. **보도자료 미리보기**:
   - Fetch API로 `/api/generate-preview` 호출
   - iframe에 HTML 렌더링
   - 이전/다음 네비게이션

3. **결측치 처리**:
   - 결측치 필드 리스트 표시
   - 모달에서 값 입력
   - 커스텀 데이터로 재생성

4. **전체 생성**:
   - 모든 보도자료 일괄 생성
   - 진행 상태 표시
   - 완료 후 알림

5. **내보내기**:
   - PDF용 HTML 생성
   - XLSX 형식 내보내기
   - 한글 복붙용 HTML 생성

---

## 7. 보도자료 생성 프로세스

### 7.1 전체 프로세스 개요

```
사용자 작업 흐름:
1. 파일 업로드
2. 파일 유형 감지 및 처리
3. GRDP 데이터 확인
4. 보도자료 미리보기 (개별 또는 전체)
5. 결측치 확인 및 입력
6. 검토 완료
7. 전체 생성
8. 내보내기
```

### 7.2 파일 업로드 프로세스

#### 7.2.1 분석표 업로드 (프로세스 2)

```
분석표 엑셀 파일 업로드
    ↓
파일 유형 감지: 'analysis'
    ↓
수식 계산 전처리 (preprocess_excel)
    ├─→ openpyxl 직접 계산 시도
    ├─→ formulas 라이브러리 시도
    └─→ xlwings 시도 (fallback)
    ↓
연도/분기 추출 (extract_year_quarter_from_excel)
    ↓
GRDP 시트 확인
    ├─→ 있음: GRDP 데이터 추출 및 세션 저장
    └─→ 없음: GRDP 업로드 요청 모달 표시
    ↓
세션에 저장:
    - excel_path
    - year, quarter
    - file_type: 'analysis'
    - grdp_data (있을 경우)
    ↓
대시보드 준비 완료
```

#### 7.2.2 기초자료 업로드 (프로세스 1)

```
기초자료 수집표 업로드
    ↓
파일 유형 감지: 'raw'
    ↓
DataConverter 초기화
    ↓
연도/분기 자동 감지
    ↓
분석표 자동 생성 (convert_all)
    ├─→ 시트별 변환
    ├─→ 가중치 처리
    └─→ 분석표 저장
    ↓
수식 계산 전처리
    ↓
GRDP 데이터 추출 시도
    ├─→ 기초자료에서 추출
    ├─→ 분석표 GRDP 시트 확인
    └─→ 없으면 기본값 사용
    ↓
세션에 저장:
    - raw_excel_path
    - excel_path (생성된 분석표)
    - year, quarter
    - file_type: 'raw_with_analysis'
    - grdp_data
    ↓
대시보드 준비 완료 (바로 보도자료 생성 가능)
```

### 7.3 개별 보도자료 생성 프로세스

```
보도자료 ID + 엑셀 경로
    ↓
Generator 모듈 동적 로드
    ├─→ templates/{generator_name} 로드
    └─→ importlib.util 사용
    ↓
데이터 추출 방식 결정
    ├─→ generate_report_data() 함수
    ├─→ generate_report() 함수 (레거시)
    └─→ Generator 클래스 (최후)
    ↓
데이터 추출 실행
    ├─→ 엑셀 파일 읽기
    ├─→ 특정 시트 접근
    ├─→ 데이터 파싱
    ├─→ 계산 및 가공
    └─→ 딕셔너리 형태 구조화
    ↓
데이터 후처리
    ├─→ Top3 지역 정규화
    ├─→ 커스텀 데이터 병합
    └─→ 결측치 검증
    ↓
Jinja2 템플릿 렌더링
    ├─→ 템플릿 파일 로드
    ├─→ 필터 등록
    ├─→ 데이터 바인딩
    └─→ HTML 생성
    ↓
HTML 반환
```

### 7.4 요약 보도자료 생성 프로세스

요약 보도자료는 특별한 처리가 필요합니다:

1. **표지, 일러두기**: 스키마 파일 또는 기본값 사용
2. **목차**: 동적 페이지 번호 계산 (`_get_toc_sections()`)
3. **인포그래픽**: 여러 시트 데이터 통합
4. **요약 보도자료**: `summary_data.py` 함수 사용

### 7.5 시도별 보도자료 생성 프로세스

```
지역명 + 엑셀 경로
    ↓
RegionalGenerator 클래스 초기화
    ↓
지역별 데이터 추출
    ├─→ 해당 지역 데이터 필터링
    ├─→ 전국 데이터와 비교
    └─→ 순위 및 변화율 계산
    ↓
템플릿 렌더링 (regional_template.html)
    ↓
HTML 반환
```

### 7.6 통계표 보도자료 생성 프로세스

```
통계표 ID + 엑셀 경로
    ↓
StatisticsTableGenerator 초기화
    ↓
TABLE_CONFIG에서 설정 확인
    ↓
데이터 추출
    ├─→ 연도별 데이터
    ├─→ 분기별 데이터
    └─→ 지역별 데이터
    ↓
페이지 분할 (1페이지: 전국~세종, 2페이지: 경기~제주)
    ↓
템플릿 렌더링 (statistics_table_index_template.html)
    ↓
HTML 반환
```

---

## 8. API 엔드포인트

### 8.1 메인 페이지 라우트 (routes/main.py)

| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET | `/` | 대시보드 페이지 |
| GET | `/preview/infographic` | 인포그래픽 미리보기 |
| GET | `/uploads/<filename>` | 업로드 파일 다운로드 |
| GET | `/view/<filename>` | 파일 직접 보기 |
| GET | `/debug/<filename>` | 디버그 파일 보기 |
| GET | `/exports/<path:filepath>` | 내보내기 파일 보기 |
| GET | `/templates/<filename>` | 템플릿 파일 서빙 |
| GET | `/download-export/<export_dir>` | 내보내기 ZIP 다운로드 |

### 8.2 API 라우트 (routes/api.py)

#### 8.2.1 파일 업로드 및 관리

**POST `/api/upload`**
- 파일 업로드
- 파일 유형 자동 감지
- 분석표 생성 또는 GRDP 추출
- 응답: `{success, filename, file_type, year, quarter, reports, ...}`

**POST `/api/upload-grdp`**
- GRDP 파일 업로드
- KOSIS GRDP 파일 파싱
- 분석표에 GRDP 시트 추가
- 응답: `{success, message, national_growth_rate, ...}`

**POST `/api/use-default-grdp`**
- 기본 GRDP 데이터 사용
- 플레이스홀더 데이터 생성
- 응답: `{success, message, is_placeholder, ...}`

**GET `/api/check-grdp`**
- GRDP 데이터 상태 확인
- 응답: `{success, has_grdp, source, ...}`

#### 8.2.2 보도자료 생성

**POST `/api/generate-preview`** (routes/preview.py)
- 개별 보도자료 미리보기 생성
- 요청: `{report_id, year, quarter, custom_data}`
- 응답: `{success, html, missing_fields, report_id, report_name}`

**POST `/api/generate-all`**
- 전체 보도자료 일괄 생성
- 요청: `{year, quarter, all_custom_data}`
- 응답: `{success, generated, errors}`

**POST `/api/generate-all-regional`**
- 시도별 보도자료 전체 생성
- 응답: `{success, generated, errors}`

#### 8.2.3 보도자료 순서 관리

**GET `/api/report-order`**
- 현재 보도자료 순서 반환
- 응답: `{reports, regional_reports}`

**POST `/api/report-order`**
- 보도자료 순서 업데이트
- 요청: `{order: [...]}`

#### 8.2.4 세션 관리

**GET `/api/session-info`**
- 현재 세션 정보 반환
- 응답: `{excel_path, year, quarter, has_file}`

#### 8.2.5 내보내기

**POST `/api/export-final`**
- 최종 문서 HTML 생성 (PDF용)
- 요청: `{pages, year, quarter, standalone}`
- 응답: `{success, html, filename, download_url, total_pages}`

**POST `/api/export-xlsx`**
- XLSX 형식 내보내기
- 요청: `{pages, year, quarter}`
- 응답: `{success, filename, download_url, xlsx_data, ...}`

**POST `/api/export-hwp-ready`**
- 한글 복붙용 HTML 생성
- 요청: `{pages, year, quarter}`
- 응답: `{success, html, filename, view_url, download_url}`

**POST `/api/save-html-to-project`**
- HTML을 프로젝트 메인 디렉토리에 저장
- 응답: `{success, filename, path, total_pages}`

#### 8.2.6 기타

**GET `/api/download-analysis`**
- 분석표 다운로드 (기초자료 업로드 시)
- 분석표 생성 및 수식 계산 후 다운로드

**POST `/api/generate-analysis-with-weights`**
- 가중치 설정 포함 분석표 생성

**POST `/api/cleanup-uploads`**
- 업로드 폴더 정리

**GET `/api/get-industry-weights`**
- 업종별 가중치 정보 추출

### 8.3 미리보기 라우트 (routes/preview.py)

**POST `/api/generate-summary-preview`**
- 요약 보도자료 미리보기
- 표지, 일러두기, 목차, 인포그래픽 등

**POST `/api/generate-regional-preview`**
- 시도별 보도자료 미리보기

**POST `/api/generate-statistics-preview`**
- 개별 통계표 미리보기

**POST `/api/generate-statistics-full-preview`**
- 통계표 전체 미리보기

## 9. 데이터 구조 및 형식

### 9.1 보도자료 데이터 구조

모든 보도자료는 공통된 기본 구조를 따릅니다:

```python
{
    'report_info': {
        'year': int,           # 연도 (예: 2025)
        'quarter': int,        # 분기 (1, 2, 3, 4)
        'page_number': str,    # 페이지 번호 (선택사항)
        'organization': str,   # 조직명 (예: '통계청')
        'department': str      # 부서명 (예: '경제통계심의관')
    },
    # 보도자료별 특정 데이터...
}
```

### 9.2 부문별 보도자료 데이터 구조

#### 9.2.1 광공업생산 / 서비스업생산

```python
{
    'report_info': {...},
    'national_data': {
        'current_value': float,      # 현재값 (지수)
        'previous_value': float,     # 전년동기값
        'change': float,             # 증감률 (%)
        'direction': str,            # '증가' or '감소'
        'top_industries': [          # 상위 업종
            {
                'name': str,
                'current_value': float,
                'change': float
            },
            ...
        ]
    },
    'regional_data': {
        'increase_regions': [        # 증가 지역 (정렬됨)
            {
                'region': str,
                'current_value': float,
                'change': float,
                'top_industries': [...]
            },
            ...
        ],
        'decrease_regions': [...],   # 감소 지역
        'all_regions': [...]         # 전체 지역 (원본 순서)
    },
    'summary_table': {
        'columns': {
            'change_columns': [str], # 증감률 컬럼명
            'rate_columns': [str]    # 지수 컬럼명
        },
        'regions': [
            {
                'region': str,
                'values': {
                    '2023.2/4': float,
                    '2024.2/4': float,
                    # ...
                },
                'changes': {
                    '2023.2/4': float,
                    '2024.2/4': float,
                    # ...
                }
            },
            ...
        ]
    },
    'top3_increase_regions': [...],  # Top3 증가 지역
    'top3_decrease_regions': [...]   # Top3 감소 지역
}
```

#### 9.2.2 고용률 / 실업률

```python
{
    'report_info': {...},
    'nationwide_data': {
        'current_rate': float,       # 현재 고용률/실업률
        'previous_rate': float,      # 전년동기
        'change': float,             # 증감 (%p)
        'direction': str,            # '증가' or '감소'
        'age_groups': [              # 연령대별
            {
                'age_group': str,    # 예: '20-29세'
                'current_rate': float,
                'change': float
            },
            ...
        ]
    },
    'regional_data': {
        'increase_regions': [
            {
                'region': str,
                'current_rate': float,
                'change': float,
                'top_age_groups': [...]
            },
            ...
        ],
        'decrease_regions': [...],
        'all_regions': [...]
    },
    'summary_box': {
        'total_regions': int,
        'increase_count': int,
        'decrease_count': int,
        'unchanged_count': int
    },
    'summary_table': {
        'columns': {...},
        'regions': [...]
    },
    'top3_increase_regions': [...],
    'top3_decrease_regions': [...]
}
```

### 9.3 GRDP 데이터 구조

```python
{
    'report_info': {
        'year': int,
        'quarter': int,
        'page_number': str
    },
    'national_summary': {
        'growth_rate': float,        # 전국 성장률 (%)
        'direction': str,            # '증가' or '감소'
        'contributions': {
            'manufacturing': float,  # 제조업 기여율
            'construction': float,   # 건설업 기여율
            'service': float,        # 서비스업 기여율
            'other': float           # 기타 기여율
        },
        'placeholder': bool          # 플레이스홀더 여부
    },
    'top_region': {
        'name': str,                 # 최고 성장 지역
        'growth_rate': float,
        'contributions': {...}
    },
    'regional_data': [
        {
            'region': str,
            'region_group': str,     # 권역 (경인, 충청, 호남 등)
            'growth_rate': float,
            'manufacturing': float,
            'construction': float,
            'service': float,
            'other': float,
            'is_group_start': bool,  # 권역 시작 여부
            'group_size': int,       # 권역 지역 수
            'placeholder': bool
        },
        ...
    ],
    'source': str,                   # 데이터 소스
    'kosis_info': {...}              # KOSIS 정보 (선택사항)
}
```

### 9.4 요약 보도자료 데이터 구조

#### 9.4.1 요약-지역경제동향

```python
{
    'report_info': {...},
    'summary': {
        'production': {
            'mining': {...},         # 광공업생산 요약
            'service': {...}         # 서비스업생산 요약
        },
        'consumption': {...},        # 소비 요약
        'exports': {...},            # 수출 요약
        'price': {...},              # 물가 요약
        'employment': {...}          # 고용 요약
    },
    'table_data': {
        # 표 데이터
    },
    'page_number': int
}
```

### 9.5 엑셀 파일 구조

#### 9.5.1 분석표 구조

분석표는 다음 시트들로 구성됩니다:

**분석 시트:**
- `A 분석`: 광공업생산 분석
- `B 분석`: 서비스업생산 분석
- `C 분석`: 소비 분석
- `D(고용률)분석`: 고용률 분석
- `D(실업)분석`: 실업률 분석
- `E(품목성질물가)분석`: 물가 분석
- `G 분석`: 수출 분석
- `H 분석`: 수입 분석
- `I(순인구이동)집계`: 인구이동 집계

**집계 시트:**
- `A(광공업생산)집계`: 광공업생산 집계 데이터
- `B(서비스업생산)집계`: 서비스업생산 집계 데이터
- `C(소비)집계`: 소비 집계 데이터
- `D(고용률)집계`: 고용률 집계 데이터
- `D(실업)집계`: 실업률 집계 데이터
- `E(품목성질물가)집계`: 물가 집계 데이터
- `G(수출)집계`: 수출 집계 데이터
- `H(수입)집계`: 수입 집계 데이터
- `I(순인구이동)집계`: 인구이동 집계 데이터

**기타 시트:**
- `I GRDP` 또는 `GRDP`: GRDP 데이터 (선택사항)
- `본청`: 메타데이터
- `이용관련`: 이용 안내

#### 9.5.2 기초자료 수집표 구조

기초자료 수집표는 다음과 같은 시트들로 구성됩니다:

- `광공업생산`
- `서비스업생산`
- `소비(소매, 추가)`
- `고용률`
- `실업자 수`
- `품목성질별 물가`
- `건설 (공표자료)`
- `수출`
- `수입`
- `시도 간 이동`
- `분기 GRDP` (선택사항)

### 9.6 데이터 형식 규칙

#### 9.6.1 숫자 형식

- **지수**: 소수점 첫째 자리 (예: 105.2)
- **증감률**: 소수점 첫째 자리, % 단위 (예: 2.5)
- **고용률/실업률**: 소수점 첫째 자리, % 단위 (예: 62.5)
- **증감 (%p)**: 소수점 첫째 자리, %p 단위 (예: 0.3)
- **금액**: 천 단위 구분 (예: 1,234,567)

#### 9.6.2 날짜 형식

- **연도/분기**: "2025.2/4" 형식
- **일자**: "2025. 8. 12.(화)" 형식
- **일시**: "2025. 8. 12.(화) 12:00" 형식

#### 9.6.3 지역명 형식

- **정식 명칭**: "서울특별시", "부산광역시" 등
- **약칭**: "서울", "부산", "경기" 등
- **보도자료 표기**: "서 울", "부 산", "경 기" (공백 포함)

---

## 10. 주요 기능 상세

### 10.1 파일 업로드 및 처리

#### 10.1.1 파일 업로드 기능

**드래그 앤 드롭 지원:**
- 파일을 업로드 영역에 드래그하여 업로드
- 여러 파일 선택 시 첫 번째 파일만 처리

**파일 유형 감지:**
- 파일명 패턴 분석
- 시트명 분석
- 자동으로 'raw' 또는 'analysis' 판정

**업로드 진행 상태:**
- 단계별 진행 표시
- 파일 저장 → 유형 분석 → 처리 완료

#### 10.1.2 파일 전처리

**수식 계산:**
- 분석 시트의 수식이 집계 시트를 참조하는 경우 처리
- 여러 방법 시도 (openpyxl → formulas → xlwings)

**데이터 검증:**
- 파일 무결성 확인
- 필수 시트 존재 여부 확인
- 데이터 형식 검증

### 10.2 보도자료 미리보기

#### 10.2.1 실시간 미리보기

- 보도자료 목록에서 클릭 시 즉시 미리보기 생성
- iframe에 HTML 렌더링
- 새로고침 버튼으로 재생성

#### 10.2.2 네비게이션

- 이전/다음 버튼으로 보도자료 이동
- 키보드 단축키 지원 (화살표 키)
- 현재 위치 표시

#### 10.2.3 결측치 처리

- 결측치 자동 감지
- 모달창에서 값 입력
- 커스텀 데이터로 재생성
- 건너뛰기 옵션

### 10.3 보도자료 생성

#### 10.3.1 개별 생성

- 미리보기 시 자동 생성
- 요청 시 재생성
- 커스텀 데이터 반영

#### 10.3.2 전체 생성

- 모든 보도자료 일괄 생성
- 진행 상태 표시
- 성공/실패 통계
- 완료 후 알림

#### 10.3.3 배치 생성

- 시도별 보도자료 전체 생성
- 통계표 전체 생성
- 병렬 처리 가능 (향후 개선)

### 10.4 내보내기 기능

#### 10.4.1 PDF용 HTML

- A4 크기 최적화
- 페이지 브레이크 설정
- 인쇄 스타일 적용
- 브라우저 인쇄 기능으로 PDF 저장

#### 10.4.2 XLSX 형식

- HTML을 엑셀 형식으로 변환
- 표 데이터 보존
- 이미지는 Base64 인코딩
- 여러 시트로 분리

#### 10.4.3 한글 복붙용 HTML

- 한글 프로그램에서 열 수 있는 형식
- 인라인 스타일 적용
- 표 서식 보존
- 이미지는 Base64 인코딩
- 복사-붙여넣기 최적화

### 10.5 검토 시스템

#### 10.5.1 검토 상태 추적

- 각 보도자료별 검토 완료 상태 저장
- 로컬 스토리지에 상태 저장
- 페이지 새로고침 시에도 유지

#### 10.5.2 검토 완료 표시

- 검토 완료 버튼 클릭 시 상태 변경
- 시각적 피드백 (체크 표시)
- 진행률 표시

### 10.6 GRDP 데이터 관리

#### 10.6.1 GRDP 데이터 추출

- 기초자료에서 자동 추출
- 분석표 GRDP 시트에서 추출
- KOSIS 파일 파싱

#### 10.6.2 GRDP 업로드

- KOSIS GRDP 파일 업로드
- 파일 파싱 및 검증
- 분석표에 GRDP 시트 추가

#### 10.6.3 기본값 사용

- GRDP 데이터가 없을 때 플레이스홀더 생성
- 사용자에게 수정 안내
- 나중에 실제 데이터로 업데이트 가능

---

## 11. 설정 및 환경

### 11.1 환경 변수

현재 프로젝트는 환경 변수를 직접 사용하지 않지만, 향후 확장 가능:

- `FLASK_ENV`: Flask 환경 (development, production)
- `SECRET_KEY`: 세션 암호화 키 (현재는 코드에 하드코딩)
- `MAX_CONTENT_LENGTH`: 최대 업로드 크기
- `VERCEL`: Vercel 환경 감지

### 11.2 경로 설정 (config/settings.py)

```python
BASE_DIR = Path(__file__).parent.parent
TEMPLATES_DIR = BASE_DIR / 'templates'
UPLOAD_FOLDER = BASE_DIR / 'uploads'  # 로컬
# 또는
UPLOAD_FOLDER = Path('/tmp/uploads')  # Vercel
```

### 11.3 Flask 설정

```python
app.secret_key = SECRET_KEY
app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH  # 50MB
```

### 11.4 보도자료 설정 (config/reports.py)

**보도자료 순서:**
- `REPORT_ORDER`: 전체 보도자료 순서
- `SUMMARY_REPORTS`: 요약 보도자료 목록
- `SECTOR_REPORTS`: 부문별 보도자료 목록

**페이지 번호 설정:**
- `PAGE_CONFIG`: 섹션별 페이지 수
- 목차 생성 시 동적 계산

### 11.5 개발 환경

**로컬 개발:**
- Python 3.10+
- 가상환경 권장
- 포트: 5050

**프로덕션 (Vercel):**
- 서버리스 함수
- `/tmp` 디렉토리만 쓰기 가능
- 환경 변수로 설정 관리

---

## 12. 배포 및 운영

### 12.1 로컬 실행

```bash
# 가상환경 활성화
source VENV/bin/activate

# 서버 실행
python app.py
```

**접속:**
- http://localhost:5050

### 12.2 Vercel 배포

**배포 설정:**
- `vercel.json`: 라우팅 설정
- `vercel-requirements.txt`: 의존성 (선택사항)
- 환경 변수 설정 (Vercel 대시보드)

**주의사항:**
- `/tmp` 디렉토리만 쓰기 가능
- 세션 데이터는 메모리 저장 (서버리스 특성상 제한적)
- 파일 업로드 크기 제한 확인

### 12.3 파일 관리

**업로드 파일:**
- `uploads/`: 업로드된 엑셀 파일
- 자동 정리 기능 (세션 종료 시)

**생성 파일:**
- `templates/*_output.html`: 생성된 보도자료
- `exports/`: 내보내기 파일
- `debug/`: 디버그 파일

**로그:**
- 콘솔 출력 (stdout/stderr)
- Vercel 로그 확인

---

## 13. 문제 해결 가이드

### 13.1 일반적인 오류

#### 파일 업로드 실패

**원인:**
- 파일 크기 초과 (50MB 제한)
- 파일 형식 오류
- 디스크 공간 부족

**해결:**
- 파일 크기 확인
- `.xlsx` 또는 `.xls` 형식 확인
- 디스크 공간 확인
- Vercel 환경: `/tmp` 디렉토리 확인

#### Generator 모듈 로드 실패

**원인:**
- Generator 파일 경로 오류
- Python 문법 오류
- 모듈 의존성 문제

**해결:**
- Generator 파일 존재 확인
- Python 문법 검사
- 필요한 라이브러리 설치 확인

#### 데이터 추출 실패

**원인:**
- 시트명 불일치
- 데이터 형식 오류
- 셀 위치 변경

**해결:**
- 엑셀 파일 시트명 확인
- 데이터 형식 검증
- Generator 코드 확인 및 수정

### 13.2 성능 문제

#### 보도자료 생성 느림

**원인:**
- 대용량 엑셀 파일
- 복잡한 데이터 처리
- 네트워크 지연

**해결:**
- 파일 크기 최적화
- 데이터 캐싱 고려
- 병렬 처리 (향후 개선)

#### 메모리 부족

**원인:**
- 대용량 파일 처리
- 여러 파일 동시 처리

**해결:**
- 파일 크기 제한
- 순차 처리
- 메모리 효율적인 데이터 구조 사용

### 13.3 데이터 문제

#### 결측치 발생

**원인:**
- 원본 데이터 누락
- 데이터 형식 불일치
- 셀 위치 오류

**해결:**
- 원본 엑셀 파일 확인
- 결측치 입력 모달 사용
- Generator 코드 검증

#### 수식 계산 오류

**원인:**
- 수식 계산 라이브러리 부재
- 복잡한 수식
- Excel 앱 미설치 (xlwings 사용 시)

**해결:**
- `formulas` 라이브러리 설치
- Excel 앱 설치 (xlwings 사용 시)
- 수식 단순화 고려

### 13.4 템플릿 문제

#### HTML 렌더링 오류

**원인:**
- 템플릿 문법 오류
- 데이터 형식 불일치
- 필터 함수 오류

**해결:**
- Jinja2 문법 확인
- 데이터 구조 검증
- 필터 함수 테스트

#### 스타일 적용 안 됨

**원인:**
- CSS 경로 오류
- 인라인 스타일 누락
- 브라우저 호환성

**해결:**
- CSS 파일 경로 확인
- 인라인 스타일 사용
- 브라우저 호환성 테스트

### 13.5 디버깅 방법

#### 로그 확인

- 콘솔 출력 확인
- Flask 디버그 모드 활성화
- Vercel 로그 확인

#### 단계별 테스트

1. 파일 업로드 테스트
2. 데이터 추출 테스트
3. 템플릿 렌더링 테스트
4. 전체 프로세스 테스트

#### 디버그 모드

- `app.py`에서 `debug=True` 설정
- 자동 재로드 활성화
- 상세 오류 메시지 표시

---

## 부록

### A. 용어 정리

- **기초자료 수집표**: 원본 데이터가 담긴 엑셀 파일
- **분석표**: 기초자료를 변환하여 만든 분석용 엑셀 파일
- **보도자료**: 최종 생성되는 HTML 문서
- **Generator**: 데이터 추출 로직을 담은 Python 모듈
- **Template**: HTML 생성에 사용되는 Jinja2 템플릿
- **GRDP**: 분기 지역내총생산 (Gross Regional Domestic Product)

### B. 참고 자료

- Flask 공식 문서: https://flask.palletsprojects.com/
- Jinja2 문서: https://jinja.palletsprojects.com/
- pandas 문서: https://pandas.pydata.org/
- openpyxl 문서: https://openpyxl.readthedocs.io/

### C. 변경 이력

- 2025년: 프로젝트 초기 버전
- 주요 기능: 기초자료 → 분석표 변환, 보도자료 자동 생성

---

**문서 버전**: 1.0  
**최종 업데이트**: 2025년  
**작성자**: 프로젝트 개발팀

