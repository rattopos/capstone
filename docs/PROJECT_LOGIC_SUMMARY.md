# 프로젝트 로직 요약

## 📋 개요
**지역경제동향 보도자료 자동 생성 시스템** - 분석표 엑셀 파일을 업로드하면 HTML 보도자료를 자동으로 생성하는 Flask 기반 웹 애플리케이션

---

## 🔄 주요 워크플로우

### 1. 파일 업로드 및 전처리
```
사용자 업로드 (분석표 엑셀 파일)
    ↓
파일 유형 감지 ('analysis')
    ↓
수식 계산 전처리 (preprocess_excel)
    ├─→ openpyxl 직접 계산 시도 (1순위)
    ├─→ formulas 라이브러리 시도 (2순위)
    └─→ xlwings 시도 (3순위, Excel 앱 필요)
    ↓
연도/분기 자동 추출 (extract_year_quarter_from_excel)
    ↓
GRDP 데이터 확인 및 추출
    ↓
세션에 저장 (excel_path, year, quarter, grdp_data)
```

### 2. 보도자료 생성 프로세스
```
보도자료 ID + 엑셀 경로
    ↓
Generator 모듈 동적 로드 (templates/{generator_name}.py)
    ├─→ 클래스 기반 Generator (대부분)
    └─→ 함수 기반 Generator (고용률, 실업률)
    ↓
데이터 추출
    ├─→ 엑셀 파일 읽기 (pandas/openpyxl)
    ├─→ 특정 시트 접근 (예: 'A 분석', 'B 분석')
    ├─→ 데이터 파싱 및 계산
    └─→ 딕셔너리 형태로 구조화
    ↓
커스텀 데이터 병합 (결측치 대체 등)
    ↓
결측치 검증 (check_missing_data)
    ↓
Jinja2 템플릿 렌더링
    ├─→ 템플릿 로드 (templates/{template_name}.html)
    ├─→ 데이터 바인딩
    └─→ HTML 문자열 생성
    ↓
HTML 반환 (미리보기 또는 최종 생성)
```

### 3. 보도자료 종류
- **개별 보도자료** (9개): 광공업생산, 서비스업생산, 소비동향, 고용률, 실업률, 물가동향, 수출, 수입, 국내인구이동
- **지역별 보도자료**: 서울, 부산, 대구, 인천 등 주요 도시별 경제동향
- **요약 보도자료**: 생산, 고용, 수출/물가, 소비/건설, 지역경제동향
- **통계표**: 목차, 참고-GRDP, 부록 등

---

## 🏗️ 아키텍처 구조

### 계층 구조
```
┌─────────────────────────────────────┐
│   Frontend (dashboard.html)         │  ← 사용자 인터페이스
│   - 파일 업로드                      │
│   - 미리보기                         │
│   - 검토 시스템                      │
└─────────────────────────────────────┘
              ↕ HTTP API
┌─────────────────────────────────────┐
│   Routes Layer (routes/api.py)      │  ← API 엔드포인트
│   - /api/upload                     │
│   - /api/generate-preview           │
│   - /api/generate-all               │
└─────────────────────────────────────┘
              ↕
┌─────────────────────────────────────┐
│   Services Layer                    │  ← 비즈니스 로직
│   - report_generator.py             │  ← 보도자료 생성
│   - excel_processor.py              │  ← 엑셀 전처리
│   - grdp_service.py                 │  ← GRDP 데이터 처리
│   - summary_data.py                 │  ← 요약 데이터 생성
└─────────────────────────────────────┘
              ↕
┌─────────────────────────────────────┐
│   Generator Layer (templates/)      │  ← 데이터 추출
│   - {report}_generator.py           │  ← 각 보도자료별 Generator
│   - {report}_template.html          │  ← Jinja2 템플릿
│   - {report}_schema.json            │  ← 데이터 스키마
└─────────────────────────────────────┘
              ↕
┌─────────────────────────────────────┐
│   Utils Layer                       │  ← 유틸리티 함수
│   - excel_utils.py                  │  ← 엑셀 처리 유틸
│   - data_utils.py                   │  ← 데이터 처리 유틸
│   - filters.py                      │  ← Jinja2 필터
└─────────────────────────────────────┘
```

---

## 🔧 핵심 컴포넌트

### 1. **ReportGenerator** (`services/report_generator.py`)
- 보도자료 생성의 핵심 엔진
- Generator 모듈을 동적으로 로드하여 데이터 추출
- Jinja2 템플릿으로 HTML 생성
- 결측치 검증 및 커스텀 데이터 병합

### 2. **Excel Processor** (`services/excel_processor.py`)
- 엑셀 파일의 수식 계산 전처리
- 여러 방법 시도: openpyxl → formulas → xlwings
- 연도/분기 자동 추출

### 3. **Generator 모듈들** (`templates/*_generator.py`)
- 각 보도자료별 데이터 추출 로직
- 엑셀 시트에서 데이터 읽기 및 가공
- 구조화된 딕셔너리 반환

### 4. **Data Converter** (`data_converter.py`)
- ⚠️ **현재 사용하지 않음** (기초자료 사용 금지 규칙)
- 기초자료 수집표 → 분석표 변환 기능 (레거시)

### 5. **GRDP Service** (`services/grdp_service.py`)
- GRDP(지역내총생산) 데이터 처리
- KOSIS API 연동 또는 파일 파싱
- 기여율 계산

---

## 📊 데이터 흐름

### 입력 데이터
- **분석표 엑셀 파일**: 각 보도자료별 분석 시트 (예: 'A 분석', 'B 분석')
- **GRDP 데이터**: 지역별 경제 기여율 정보 (선택사항)

### 처리 과정
1. 엑셀 파일 읽기 → pandas DataFrame 또는 openpyxl Workbook
2. 시트별 데이터 추출 → Generator가 특정 시트에서 데이터 파싱
3. 데이터 가공 → 계산, 필터링, 포맷팅
4. 구조화 → 딕셔너리 형태로 변환 (템플릿에 전달 가능한 형태)

### 출력 데이터
- **HTML 보도자료**: 한글 복붙용 HTML 형식
- **인라인 스타일**: 한글 프로그램에서 인식 가능한 형식
- **표(table)**: 인라인 스타일로 border, padding 등 명시

---

## 🎯 주요 기능

### 1. **자동 데이터 추출**
- 엑셀 파일에서 연도/분기 자동 감지
- 각 보도자료별 Generator가 해당 시트에서 데이터 추출
- 수식 계산 결과 활용

### 2. **결측치 처리**
- 자동 결측치 감지 (`check_missing_data`)
- 모달창을 통한 수동 입력
- 커스텀 데이터 병합

### 3. **미리보기 시스템**
- 실시간 HTML 미리보기
- 보도자료 간 네비게이션 (이전/다음)
- 검토 완료 상태 추적

### 4. **일괄 생성**
- 모든 보도자료 한 번에 생성
- 한글 복붙용 HTML 형식으로 출력
- 파일 다운로드 지원

### 5. **보도자료 순서 관리**
- 드래그 앤 드롭으로 순서 변경
- 세션에 저장하여 유지

---

## 🔐 세션 관리

### 저장되는 정보
- `excel_path`: 업로드된 분석표 경로
- `year`, `quarter`: 연도/분기 정보
- `file_type`: 파일 유형 ('analysis')
- `grdp_data`: GRDP 데이터 (있을 경우)
- `reviewed_reports`: 검토 완료된 보도자료 목록
- `report_order`: 보도자료 순서

---

## ⚙️ 설정 및 상수

### 보도자료 설정 (`config/reports.py`)
- `REPORT_ORDER`: 보도자료 목록 및 순서
- `REGIONAL_REPORTS`: 지역별 보도자료 설정
- `SUMMARY_REPORTS`: 요약 보도자료 설정
- `STATISTICS_REPORTS`: 통계표 설정

### 경로 설정 (`config/settings.py`)
- `BASE_DIR`: 프로젝트 루트 디렉토리
- `TEMPLATES_DIR`: 템플릿 디렉토리
- `UPLOAD_FOLDER`: 업로드 파일 저장소
- `EXPORT_FOLDER`: 생성된 파일 저장소

---

## 🚨 중요 규칙

### 1. **기초자료 사용 금지**
- ⚠️ **기초자료 수집표는 사용하지 않음**
- `raw_excel_path` 파라미터는 항상 `None`
- 분석표만 사용하여 보도자료 생성

### 2. **데이터 무결성**
- 결측치는 기본값으로 채우지 않음
- 'N/A', None, 또는 빈 값으로 표시
- 원본 데이터 의미 왜곡 금지

### 3. **한글 복붙 형식**
- 한글(HWP) 파일에 붙여넣었을 때 형식 유지
- 인라인 스타일 사용
- 복잡한 CSS/JavaScript 제거

---

## 📝 주요 API 엔드포인트

| 메서드 | 엔드포인트 | 설명 |
|--------|-----------|------|
| `POST` | `/api/upload` | 엑셀 파일 업로드 |
| `POST` | `/api/generate-preview` | 특정 보도자료 미리보기 생성 |
| `POST` | `/api/generate-all` | 전체 보도자료 생성 |
| `GET` | `/api/report-order` | 보도자료 순서 조회 |
| `POST` | `/api/report-order` | 보도자료 순서 변경 |
| `GET` | `/api/session-info` | 현재 세션 정보 조회 |
| `POST` | `/api/check-grdp-status` | GRDP 데이터 상태 확인 |
| `POST` | `/api/upload-grdp-file` | GRDP 파일 업로드 |

---

## 🔄 처리 우선순위

### 수식 계산 방법
1. **openpyxl 직접 계산** (가장 빠름, 백엔드 직접 계산)
2. **formulas 라이브러리** (순수 Python, 복잡한 수식 지원)
3. **xlwings** (Excel 앱 필요, 느림, 마지막 fallback)

### Generator 로드 방법
1. **generate_report_data 함수** (우선)
2. **Generator 클래스의 generate 메서드**
3. **스키마 기본값** (Generator 없는 경우)

---

## 📌 핵심 원칙

1. **최소한의 변경**: 정상 동작하는 코드는 건드리지 않음
2. **데이터 무결성**: 원본 데이터 왜곡 금지
3. **한글 호환성**: 한글 프로그램에서 붙여넣기 시 형식 유지
4. **순차적 처리**: 체크리스트 기반 작업 시 섹션별 순서대로 처리
5. **상태 검증**: 코드 수정 전 현재 상태 확인

---

## 🎓 기술 스택

- **Backend**: Flask (Python 웹 프레임워크)
- **템플릿 엔진**: Jinja2
- **데이터 처리**: pandas, openpyxl
- **수식 계산**: openpyxl, formulas, xlwings
- **프론트엔드**: HTML, JavaScript (Vanilla)
- **스타일링**: 인라인 CSS (한글 호환성)

---

*최종 업데이트: 2025년*
