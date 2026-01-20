---

## 5-2. 표 위치(시트/범위) 관리 및 예시

### 표 위치 관리 방식
- 모든 데이터 표의 위치(엑셀 파일명, 시트명, 셀 범위, 헤더 포함 여부 등)는 `data_table_locations.md` 파일에 마크다운 형식으로 기록합니다.
- `config/table_locations.py`에서 이 파일을 파싱하여, 각 표의 시트/범위 정보를 딕셔너리로 자동 로드합니다.
- 범위는 예) `E3:AA540`처럼 엑셀 표의 시작~끝 셀로 지정하며, 코드에서 start_col, start_row, end_col, end_row로 분리해 사용합니다.

### 실제 표 위치 예시 (2025년 3분기)

| 구분       | 시트명                | 범위         | 헤더 | 템플릿                  |
|------------|----------------------|--------------|------|-------------------------|
| 광공업생산 | A(광공업생산)집계    | E3:AA540     | 예   | mining_template.html    |
| 서비스업생산| B(서비스업생산)집계  | D3:Z255      | 예   | service_template.html   |
| 소매판매   | C(소비)집계          | C3:Y109      | 예   | consumption_template.html|
| 고용률     | D(고용률)집계        | B3:V111      | 예   | employment_template.html|
| 실업률     | D(실업)집계          | A80:T152     | 예   | unemployment_template.html|
| 물가       | E(품목성질물가)집계  | A3:V471      | 예   | price_template.html     |
| 건설       | F'(건설)집계         | B3:W363      | 예   | construction_template.html|
| 수출       | G(수출)집계          | D3:AA2811    | 예   | export_template.html    |
| 수입       | H(수입)집계          | D3:AA3333    | 예   | import_template.html    |
| 순인구이동 | I(순인구이동)집계    | E3:AD309     | 예   | migration_template.html |

각 표의 실제 위치/범위는 data_table_locations.md에서 관리하며, 코드에서는 table_locations.py의 load_table_locations() 함수로 일관되게 불러옵니다.

---
# 📚 프로젝트 통합 문서 (capstone)

---

## 1. 프로젝트 개요

- **이름:** 지역경제동향 보도자료 생성 시스템
- **설명:** 분석표(엑셀)를 업로드하면 전국/부문별/지역별 보도자료를 자동 생성하는 웹 애플리케이션 및 CLI 도구
- **주요 효과:** 수작업 7일 → 3분, 해석/분석 시간 3배 확보, 실수 감소, 업무 효율화

---

## 2. 전체 구조 및 주요 파일

- **app.py**: Flask 웹 서버 진입점, 대시보드/REST API 제공
- **dashboard.html**: 대시보드 UI 템플릿
- **report_generator.py**: CLI/웹 공용 보도자료 생성기
- **requirements.txt**: Python 의존성
- **README.md**: 프로젝트 설명서
- **templates/**: Jinja2 HTML 템플릿, 데이터 생성기, 스키마, 예시 데이터, 결과물
- **uploads/**: 업로드된 엑셀 파일 저장
- **exports/**: 생성된 HTML/한글 복붙용 결과물 저장
- **config/**: 보도자료 종류/순서/설정
- **services/**: 데이터 처리, 요약/캐시/보고서 생성 등 비즈니스 로직
- **routes/**: API/뷰 라우트
- **utils/**: 엑셀/텍스트/공통 유틸리티
- **schemas/**: 데이터 스키마(JSON)

---

## 3. 데이터 흐름 및 실행 순서

1. **엑셀 업로드**: `/api/upload` → uploads/ 저장
2. **전처리**: 엑셀 수식 계산/검증, 캐시 등록 (services/excel_processor.py, services/excel_cache.py)
3. **부문별 데이터 추출/캐시**: 부문별 → 시도별 → 요약 순으로 데이터 추출 및 캐시 (services/report_generator.py)
4. **전체통합**: 요약 → 부문별 → 시도별 순으로 HTML 결합 (routes/api.py)
5. **내보내기**: exports/에 HTML 저장, 다운로드/뷰 제공

---

## 4. 주요 코드/스크립트 역할

- **app.py**: Flask 앱 생성, 라우트 등록, 서버 실행
- **routes/api.py**: 업로드, 전체 생성, 내보내기 등 API
- **routes/main.py**: 대시보드/다운로드 뷰
- **config/reports.py**: 보고서 종류/순서/설정
- **services/report_generator.py**: 템플릿 선택, 데이터 준비, Jinja2 렌더링, 스키마 기반 생성
- **templates/unified_generator.py**: 엑셀 데이터 추출, 증감률/상하위 계산
- **services/summary_data.py**: 요약 데이터 빌더
- **utils/text_utils.py**: 자연어 어휘/문장 규칙

---

## 5. 데이터/스키마/입출력 파일 구조

- **엑셀 업로드:** uploads/에 저장
- **중간 데이터:** exports/_temp/output, exports/_temp/regional_output
- **최종 결과물:** exports/지역경제동향_YYYY년_Q분기.html, exports/지역경제동향_YYYY년_Q분기_한글복붙용.html
- **스키마:** schemas/ 및 templates/*_schema.json

---

## 6. 실행 방법 및 환경설정

### 웹 서버 실행
```bash
python -m venv VENV
source VENV/bin/activate
pip install -r requirements.txt
python app.py
```
- 브라우저: http://localhost:5050

### CLI 실행
```bash
python report_generator.py --list
python report_generator.py -e 분석표_25년_2분기_캡스톤.xlsx -r employment
python report_generator.py -e 분석표_25년_2분기_캡스톤.xlsx
```

### 주요 의존성
- Python 3.10 이상
- Flask ≥2.3.0, Jinja2 ≥3.1.0, pandas ≥2.0.0, openpyxl ≥3.1.0, numpy ≥1.24.0

---

## 7. 참고 문서/흐름도
- README.md, REPORT_FLOW.md, docs/PROCESS_FLOW.md, docs/narration_rules.md
- extract_project_structure.py: 구조 자동화 문서 생성

---

## 8. 기타/기여자
- 캡스톤 프로젝트 팀
- 라이선스: 캡스톤 프로젝트용

---

> 본 문서는 2026-01-20 기준, 전체 프로젝트 구조/흐름/실행/데이터/코드/설정/문서화를 통합 정리한 최신 버전입니다.
