# 전체 보고서 생성 흐름

아래는 **전체 보고서 생성**의 핵심 흐름과 파일 역할 요약입니다.

## 1) 진입점 (웹/CLI)

- **웹 UI 실행**: [app.py](app.py)
  - Flask 서버 시작
  - 대시보드/API 라우트 연결
- **대시보드 UI**: [dashboard.html](dashboard.html)
  - 업로드/연도·분기 선택/보도자료 목록/다운로드
- **CLI 생성기**: [report_generator.py](report_generator.py)
  - 커맨드라인으로 단일/전체 보도자료 생성

---

## 2) 라우팅 레이어 (API)

- **보도자료 생성 API**: [routes/api.py](routes/api.py)
  - 업로드 처리, 전체 생성
  - 결과 HTML 저장
  - 한글 복붙용 내보내기(`export-hwp-ready`) 생성
- **뷰 라우트**: [routes/main.py](routes/main.py)
  - 다운로드 엔드포인트

---

## 3) 설정/정의 (어떤 보고서를 만들지)

- **보고서 목록/순서/설정**: [config/reports.py](config/reports.py)
  - 부문별 보고서(`SECTOR_REPORTS`)
  - 요약 보고서(`SUMMARY_REPORTS`)
  - 시도 목록/권역 매핑
  - 전체 생성 순서: **부문별 → 시도별 → 요약**
- **레거시 설정**: [config/report_configs.py](config/report_configs.py)
  - 구버전 설정(현재는 `reports.py`가 기준)

---

## 4) 생성 서비스 (핵심 오케스트레이션)

- **보도자료 생성 서비스**: [services/report_generator.py](services/report_generator.py)
  - 템플릿 선택
  - 데이터 준비
  - Jinja2 렌더링
  - 스키마 기반 생성(요약 보고서 등)

---

## 5) 데이터 로딩/가공

- **통합 생성기**: [templates/unified_generator.py](templates/unified_generator.py)
  - 엑셀 집계시트 기반 데이터 추출
  - 증감률 계산, 지역별 상/하위 추출
- **공통 베이스**: [templates/base_generator.py](templates/base_generator.py)
  - 나레이션 패턴, 공통 계산 로직
- **요약 데이터 빌더**: [services/summary_data.py](services/summary_data.py)
  - 요약 보고서용 집계 데이터 구성
- **엑셀 캐시/유틸**: [services/excel_cache.py](services/excel_cache.py), [utils/excel_utils.py](utils/excel_utils.py)
  - 로드/캐시/컬럼 탐색 보조

---

## 6) 템플릿 (HTML 구조)

- **부문별 템플릿**: [templates](templates)
  - mining_template.html, service_template.html, consumption_template.html, construction_template.html,
    export_template.html, import_template.html, price_template.html, employment_template.html,
    unemployment_template.html, migration_template.html
- **요약 템플릿**: [templates/summary_*.html](templates)
  - 요약-지역경제동향, 요약-생산, 요약-소비건설 등
- **시도별 통합 템플릿**: [templates/regional_economy_by_region_template.html](templates/regional_economy_by_region_template.html)

---

## 7) 스키마 (기본 구조/예시)

- **요약/통합 스키마**: [templates/*_schema.json](templates)
  - 스키마 기반 렌더링 또는 데이터 구조 참고

---

## 8) 결과물 저장

- **생성 HTML**: [templates](templates)
  - `*_output.html` 형태로 저장
- **한글 복붙용 결과물**: [exports](exports)
  - `지역경제동향_YYYY년_Q분기_한글복붙용.html`

---

## 9) 자연어 문장 규칙

- **어휘 통제/문장 규칙**: [utils/text_utils.py](utils/text_utils.py)
  - 증가/감소 vs 상승/하락 매핑
  - 순유입/순유출 등 용어 통제

---

## 전체 흐름 요약 (단계별)

1. **업로드/연도·분기 선택** → [app.py](app.py), [dashboard.html](dashboard.html)
2. **API 호출** → [routes/api.py](routes/api.py)
3. **보고서 구성/템플릿 선택** → [config/reports.py](config/reports.py)
4. **데이터 추출/가공** → [templates/unified_generator.py](templates/unified_generator.py), [services/summary_data.py](services/summary_data.py)
5. **템플릿 렌더링** → [services/report_generator.py](services/report_generator.py)
6. **HTML 저장/내보내기** → [templates](templates), [exports](exports)
