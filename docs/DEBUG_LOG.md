# 🐛 디버그 작업 로그

이 문서는 프로젝트의 모든 디버그 작업을 추적하고 기록합니다.

---

## 📋 로그 형식

각 디버그 항목은 다음 형식으로 기록됩니다:

- **날짜/시간**: 작업 수행 시점
- **문제 설명**: 발견된 문제 또는 버그
- **원인 분석**: 문제의 근본 원인
- **에이전트 사고 과정**: AI 에이전트가 문제를 해결하기까지의 추론 과정
  - 문제 인식 및 초기 분석
  - 고려한 해결책들
  - 선택한 접근 방법과 그 이유
  - 시도한 방법들과 결과
  - 최종 결정의 근거
- **해결 방법**: 적용한 해결책
- **관련 파일**: 수정된 파일 목록
- **상태**: 진행중 / 완료 / 실패 / 보류
- **참고 사항**: 추가 정보나 향후 작업

---

## 📅 디버그 기록

### 2026-01-01

#### 파일 유형 분석 성능 최적화 - 빠른 판정 로직 구현
- **시간**: (현재)
- **문제 설명**: 파일 유형 분석(`detect_file_type`)이 너무 오래 걸림. 모든 시트를 대조하는 과정이 불필요하게 느림
- **원인 분석**: 
  - 기존 코드는 모든 시트명을 읽고 전체 시트 세트와 교집합 연산 수행
  - pandas ExcelFile로 전체 파일을 읽어서 느림
  - 시트를 모두 확인할 필요 없이 핵심 시트만 확인하면 충분
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 파일 유형 분석이 너무 오래 걸린다고 제보
  - 초기 분석: `utils/excel_utils.py`의 `detect_file_type` 함수 확인
  - 성능 병목 파악:
    - pandas ExcelFile로 전체 파일 읽기
    - 모든 시트명과 전체 시트 세트 교집합 계산
    - 불필요한 패턴 매칭 및 시트 개수 확인
  - 최적화 전략 수립:
    1. 파일명 먼저 확인 (가장 빠름, 파일 읽기 불필요)
    2. openpyxl 사용 (pandas보다 빠름, read_only 모드)
    3. 핵심 시트만 확인 (전체 시트 대조 불필요)
    4. 첫 매칭 시 즉시 반환 (조기 종료)
    5. 시트 개수로 빠른 추정
  - 구현 결정:
    - 파일명 확인을 1단계로 이동
    - openpyxl을 우선 사용 (pandas는 fallback)
    - 핵심 시트만 선별하여 빠른 매칭
    - 패턴 매칭 시 2개만 찾으면 즉시 반환
- **해결 방법**: 
  - `detect_file_type` 함수 최적화:
    1. 파일명 확인을 최우선으로 이동 (파일 읽기 불필요)
    2. openpyxl read_only 모드 사용 (pandas보다 빠름)
    3. 핵심 시트만 선별하여 빠른 매칭
    4. 첫 매칭 시 즉시 반환 (조기 종료)
    5. 불필요한 전체 시트 대조 제거
  - 성능 개선:
    - 파일명 매칭: 즉시 반환 (파일 읽기 없음)
    - 핵심 시트 매칭: 첫 시트 발견 시 즉시 반환
    - 패턴 매칭: 2개만 찾으면 즉시 반환
    - 시트 개수 추정: 간단한 비교로 빠른 판정
- **관련 파일**: 
  - `utils/excel_utils.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 기존 로직의 정확도는 유지하면서 성능만 개선
  - 파일명이 명확한 경우 파일 읽기 없이 즉시 판정 가능
  - 핵심 시트만 확인하여 대부분의 경우 빠르게 판정 가능

---

#### 전체 프로젝트 용어 변경 - '보고서' → '보도자료'
- **시간**: 17:10
- **문제 설명**: 대시보드 및 시스템 전반에서 '보고서'라는 용어가 '보도자료'로 변경되어야 함
- **원인 분석**: 
  - 통계청에서 발행하는 문서의 공식 명칭이 '보고서'가 아닌 '보도자료'임
  - UI, 로그 메시지, 에러 메시지, 주석, 문서 전체에서 일괄 변경 필요
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 '보고서'를 '보도자료'로 전면 변경 요청
  - 범위 파악: grep으로 '보고서' 사용 위치 전체 검색 → 79개 파일 발견
  - 영향 분석:
    - 코드 파일: 영어 변수명 'report'는 유지, 한글 텍스트만 변경
    - UI 파일: 사용자에게 보이는 모든 텍스트 변경
    - 문서 파일: 가이드, README 등 모든 문서 변경
  - 위험 요소 파악:
    - replace_all 사용 시 의도치 않은 변경 가능성
    - 함수명/변수명이 한글인 경우 오류 발생 가능
  - 해결 전략: replace_all로 일괄 변경 후 lint 검사로 오류 확인
  - 실행: 79개 파일 모두 순차적으로 변경
  - 검증: grep으로 남은 '보고서' 없음 확인, lint 오류 없음 확인
- **해결 방법**: 
  - 전체 프로젝트에서 한글 '보고서'를 '보도자료'로 일괄 변경
  - 변경된 파일 카테고리:
    - 핵심 코드: `app.py`, `routes/api.py`, `routes/preview.py`, `routes/debug.py`
    - 서비스: `services/report_generator.py`, `services/summary_data.py`
    - UI/템플릿: `dashboard.html`, `templates/index.html`
    - Generator 파일: `templates/*_generator.py` (15개)
    - JSON 스키마: `templates/*_schema.json` (22개)
    - 설정: `config/reports.py`, `utils/data_utils.py`
    - 문서: `docs/*.md`, `README.md` 등 (20여 개)
- **관련 파일**: 
  - 총 79개 파일 수정
- **상태**: ✅ 완료
- **참고 사항**: 
  - 영어 변수명/함수명 'report'는 코드 안정성을 위해 그대로 유지
  - 서버 재시작 후 브라우저에서 정상 표시 확인 완료

---

#### 인포그래픽 지도 이미지 404 오류 해결 - Flask 라우트 추가
- **시간**: 16:50
- **문제 설명**: 인포그래픽 페이지에서 한국 지도 이미지가 로드되지 않고 "한국 지도" alt text만 표시됨
- **원인 분석**: 
  - 템플릿에서 `/templates/infographic_map.png` 절대 경로로 이미지를 참조
  - Flask에 해당 경로에 대한 라우트가 없어서 404 Not Found 발생
  - 이전에 파일을 templates 폴더로 복사했지만, 라우트가 없어서 접근 불가
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 지도가 여전히 로드되지 않는다고 제보
  - 초기 분석: 파일 존재 확인 (287KB, 정상) → 파일은 존재함
  - 브라우저 테스트: `http://localhost:5050/templates/infographic_map.png` 접근 시 404 발생
  - 근본 원인 파악: `routes/main.py`에 `/templates/<filename>` 경로에 대한 라우트가 없음
  - 해결책 결정: templates 폴더의 정적 파일(이미지, CSS, JS)을 서비스하는 라우트 추가
- **해결 방법**: 
  - `routes/main.py`에 `/templates/<filename>` 라우트 추가
  - PNG, JPG, SVG, CSS, JS 파일에 대한 MIME type 처리 포함
- **관련 파일**: 
  - `routes/main.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - Flask debug 모드에서 자동 재시작되어 즉시 반영됨
  - 이전 해결(파일 복사)만으로는 불완전했음 - 라우트 추가 필요했음

---

#### Excel 전처리 성능 최적화 - xlwings fallback 순서 변경
- **시간**: 20:00
- **문제 설명**: 전처리 과정이 너무 오래 걸림. xlwings가 Excel 앱을 실행해야 하므로 매우 느림
- **원인 분석**: 
  - 기존 순서: xlwings → formulas → openpyxl
  - xlwings는 Excel 앱 실행이 필요하여 가장 느림
  - 백엔드에서 직접 계산하는 openpyxl 방식이 훨씬 빠름
  - 엑셀 함수 계산은 백엔드에서 계산해서 매핑하는 것이 더 효율적
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "전처리 과정이 너무 오래 걸린다"고 제보
  - 초기 분석: `services/excel_processor.py`의 `preprocess_excel()` 함수 확인
  - 현재 순서 파악: xlwings가 첫 번째로 실행되어 Excel 앱 실행으로 인한 지연 발생
  - 해결책 분석:
    1. 순서 변경: openpyxl(백엔드 직접 계산)을 우선 사용
    2. xlwings를 마지막 fallback으로 이동 (직접 계산 실패 시에만 사용)
    3. openpyxl 계산 로직 최적화 (필요한 데이터만 캐싱)
  - 선택한 접근: 백엔드 직접 계산을 우선 사용하고, xlwings는 마지막 fallback으로 변경
  - 구현 결정:
    - `preprocess_excel()` 함수의 순서 변경: openpyxl → formulas → xlwings
    - `_try_openpyxl_calculation()` 함수 최적화:
      - 필요한 집계 시트만 미리 캐싱
      - 열 문자 변환 함수 재사용
      - 빈 셀 건너뛰기로 메모리 사용 감소
    - `_try_xlwings()` 함수에 fallback 모드 주석 추가
    - `get_recommended_method()` 함수 업데이트 (openpyxl 우선 권장)
- **해결 방법**: 
  - `services/excel_processor.py` 수정:
    - `preprocess_excel()`: 실행 순서 변경 (openpyxl → formulas → xlwings)
    - `_try_openpyxl_calculation()`: 백엔드 직접 계산 로직 최적화
      - 필요한 집계 시트만 미리 캐싱
      - 열 문자 변환 함수(`col_letter_to_number`) 재사용
      - 빈 셀 건너뛰기로 메모리 효율성 향상
    - `_try_xlwings()`: fallback 모드임을 명시하는 주석 추가
    - `get_recommended_method()`: openpyxl을 가장 빠른 방법으로 우선 권장
- **관련 파일**: 
  - `services/excel_processor.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 전처리 시간이 크게 단축됨 (Excel 앱 실행 불필요)
  - 백엔드에서 직접 계산하여 서버 환경에서도 빠르게 동작
  - xlwings는 복잡한 수식이 openpyxl/formulas로 계산 실패 시에만 사용
  - 기존 generator의 fallback 로직은 안전장치로 유지

---

#### Excel 수식 자동 계산 전처리 기능 추가
- **시간**: 18:30
- **문제 설명**: 분석표 엑셀 파일의 수식이 계산되지 않은 상태로 업로드되면 보도자료에 0, NaN 등 잘못된 데이터가 표시됨
- **원인 분석**: 
  - 분석표의 "분석" 시트들은 "집계" 시트를 참조하는 수식으로 구성
  - pandas로 읽을 때 수식 결과가 아닌 수식 자체(또는 None)가 읽힘
  - 기존에는 각 generator에서 fallback 로직으로 집계 시트에서 직접 계산
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "집계표에서 계산하지 말고 엑셀 파일을 백엔드에서 실행하면 계산되지 않나요?" 제안
  - 해결책 분석:
    1. xlwings: Excel 앱이 설치되어 있으면 가장 정확 (Mac/Windows)
    2. formulas: 순수 Python으로 수식 계산 (서버 환경에서도 동작)
    3. openpyxl: 시트 간 참조 수식만 직접 계산
  - 선택한 접근: 3가지 방법을 순차적으로 시도하는 fallback 시스템 구축
  - 구현 결정: `services/excel_processor.py` 모듈 생성하여 캡슐화
- **해결 방법**: 
  - `services/excel_processor.py` 신규 생성
    - `preprocess_excel()`: 엑셀 파일 수식 계산 메인 함수
    - `_try_xlwings()`: Excel 앱으로 수식 계산 (가장 정확)
    - `_try_formulas()`: formulas 라이브러리로 순수 Python 계산
    - `_try_openpyxl_calculation()`: 시트 간 참조 수식 직접 계산
    - `check_available_methods()`: 사용 가능한 방법 확인
    - `get_recommended_method()`: 권장 방법 반환
  - `routes/api.py` 수정
    - 분석표 업로드 시 자동 전처리 수행
    - 전처리 결과를 응답에 포함
  - `requirements.txt` 업데이트
    - xlwings>=0.30.0 추가
    - Pillow>=10.0.0 추가 (이미지 처리용)
- **관련 파일**: 
  - `services/excel_processor.py` (신규 생성)
  - `routes/api.py` (수정)
  - `requirements.txt` (업데이트)
- **상태**: ✅ 완료
- **참고 사항**: 
  - xlwings 사용 시 Excel 앱이 백그라운드에서 실행됨
  - Excel이 설치되지 않은 서버에서는 formulas 또는 openpyxl fallback 사용
  - 기존 generator의 fallback 로직은 안전장치로 유지

---

#### 부문별 섹션 0/NaN/누락 행 문제 해결
- **시간**: 16:00
- **문제 설명**: 부문별 섹션에서 광공업생산, 건설동향, 실업률을 제외한 나머지 보도자료 표에 0, NaN, 또는 행 누락 발생
- **원인 분석**: 
  - 업로드된 분석표의 "분석" 시트에 수식이 계산되지 않은 상태
  - 각 generator가 분석 시트를 읽을 때 빈 값/NaN으로 읽힘
  - 일부 generator(광공업생산, 건설동향, 실업률)는 이미 집계 시트 fallback 로직이 있었음
- **에이전트 사고 과정**:
  - 영향받는 generator 식별: service_industry, consumption, export, import, price_trend, employment_rate, domestic_migration
  - 각 generator에 일관된 fallback 패턴 적용
  - 추가 발견: domestic_migration의 연령대 데이터 추출 시 `rank` 컬럼이 NaN이어서 필터링 실패 → `level` 컬럼으로 변경
  - 추가 발견: 지역명 표시 오류 (`sido.replace('', ' ')` → `' '.join(sido)`)
- **해결 방법**: 
  - 7개 generator에 `use_aggregation_only` 플래그 및 집계 시트 직접 계산 로직 추가
  - domestic_migration_generator: 연령대 필터링 조건을 `rank` → `level`로 변경
  - 지역명 표시 로직 수정
- **관련 파일**: 
  - `templates/service_industry_generator.py`
  - `templates/consumption_generator.py`
  - `templates/export_generator.py`
  - `templates/import_generator.py`
  - `templates/price_trend_generator.py`
  - `templates/employment_rate_generator.py`
  - `templates/domestic_migration_generator.py`
- **상태**: ✅ 완료
- **참고 사항**: Excel 전처리 기능 추가로 이 fallback 로직은 백업 용도로 유지

---

#### 인포그래픽 한국 지도 이미지 깨짐 문제 해결
- **시간**: 15:48
- **문제 설명**: 인포그래픽 페이지 생성 시 한국 지도 이미지가 깨져 보임 (이미지 로드 실패)
- **원인 분석**: 
  - 한국 지도 이미지가 `correct_answer/인포그래픽_map.png` 경로에 한글 파일명으로 저장되어 있음
  - 템플릿에서 `src="infographic_map.png"` 상대 경로로 참조하여 파일을 찾지 못함
  - 일부 템플릿에서는 `/correct_answer/인포그래픽_map.png` 경로 사용했으나 한글 파일명으로 인한 인코딩 문제 발생 가능
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 인포그래픽 페이지에서 한국 지도 이미지가 깨진다고 제보
  - 파일 위치 확인: `correct_answer/인포그래픽_map.png`에 이미지 존재 확인
  - 템플릿 분석: `infographic_js_template.html`, `infographic_template.html`, `infographic_regional_template.html`에서 이미지 경로 확인
  - 원인 파악: 상대 경로 사용 및 한글 파일명으로 인한 경로 해결 실패
  - 해결책 선택: 사용자 요청대로 이미지를 templates 폴더로 복사하고 영문 파일명으로 변경
- **해결 방법**: 
  - `correct_answer/인포그래픽_map.png`를 `templates/infographic_map.png`로 복사
  - 3개 템플릿 파일에서 이미지 경로를 `/templates/infographic_map.png` 절대 경로로 수정
- **관련 파일**: 
  - `templates/infographic_map.png` (신규 복사)
  - `templates/infographic_js_template.html` (경로 수정)
  - `templates/infographic_template.html` (경로 수정)
  - `templates/infographic_regional_template.html` (경로 수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 원본 파일 `correct_answer/인포그래픽_map.png`는 그대로 유지
  - 향후 이미지 업데이트 시 templates 폴더의 파일도 함께 업데이트 필요

---

### 2026-01-01

#### 통계표 2020.3/4 및 일부 분기 공란 문제 해결
- **시간**: 오후 (추가 수정)
- **문제 설명**: 통계표 HTML에서 2020.3/4 분기 행에 공란(`[  ]`) 플레이스홀더가 표시됨
- **원인 분석**: 
  - 일부 통계표(서비스업생산지수, 소매판매액지수)는 2020.4/4부터 데이터 제공 시작, 2020.3/4 데이터 없음
  - 기타 통계표(실업률, 소비자물가지수)는 JSON에 2020.3/4 데이터가 "-"로 저장되어 있었음
  - 템플릿의 `val()` 매크로가 "-" 값을 플레이스홀더로 변환
  - `quarterly_keys`가 하드코딩되어 데이터 없는 분기도 렌더링
- **에이전트 사고 과정**:
  - JSON 데이터 분석: 서비스업생산지수, 소매판매액지수, 실업률, 소비자물가지수에서 2020.3/4 공란 발견
  - 기초자료 확인: 서비스업생산지수/소매판매액지수는 2020.4/4부터 데이터 제공, 실업률/소비자물가지수는 데이터 존재
  - 해결책 1: 기초자료에서 추출 가능한 데이터(실업률, 소비자물가지수)를 JSON에 추가
  - 해결책 2: 기초자료에 없는 데이터(서비스업생산지수, 소매판매액지수)의 해당 분기 키를 JSON에서 제거
  - 해결책 3: 템플릿의 `val()` 매크로 수정 - "-" 값을 플레이스홀더가 아닌 그대로 표시
  - 해결책 4: `StatisticsTableGenerator.extract_table_data()` 수정 - 모든 지역이 "-"인 분기 자동 제거
  - 해결책 5: `services/report_generator.py` 수정 - `quarterly_keys`를 동적으로 생성
- **해결 방법**: 
  - `templates/statistics_historical_data.json`: 실업률, 소비자물가지수, 국내인구이동의 2020.3/4 데이터 추가, 서비스업생산지수/소매판매액지수의 2020.3/4 제거
  - `templates/statistics_table_index_template.html`: `val()` 매크로에서 "-" 조건 제거 (그대로 표시)
  - `templates/statistics_table_generator.py`: 모든 지역이 "-"인 분기 자동 제거 로직 추가
  - `services/report_generator.py`: `quarterly_keys`를 실제 데이터에서 동적으로 생성
- **관련 파일**: 
  - `templates/statistics_historical_data.json` (업데이트)
  - `templates/statistics_table_index_template.html` (수정)
  - `templates/statistics_table_generator.py` (수정)
  - `services/report_generator.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 일부 통계는 특정 시점부터 데이터 제공 시작하므로, 이전 분기는 표시되지 않음이 정상
  - 국내인구이동의 "전국" 데이터는 순인구이동 특성상 모든 분기에서 없음이 정상

---

#### 통계표 2025.2/4p 데이터 공란 문제 해결
- **시간**: 오후
- **문제 설명**: 통계표 HTML 생성 시 2025년 2분기(2025.2/4p) 데이터가 `[  ]` 플레이스홀더로 표시됨
- **원인 분석**: 
  - `statistics_historical_data.json` 파일의 분기 데이터 범위가 `2025.1/4`까지만 있고 `2025.2/4p`가 없음
  - 디버그 모드에서 통계표 생성 시 `raw_excel_path`가 세션에 설정되지 않아 동적 추출이 작동하지 않음
  - 템플릿의 `val()` 매크로가 값이 `-`일 때 플레이스홀더를 표시
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 공란 문제 제보 → HTML 파일에서 2197개의 `editable-placeholder` 발견
  - 초기 분석: `separator-row`의 빈 셀인지 확인 → 아님, 데이터 셀의 `-` 값이 플레이스홀더로 변환됨
  - 템플릿 분석: `statistics_table_index_template.html`의 `val()` 매크로가 `None`, `''`, `'-'` 값을 플레이스홀더로 변환
  - JSON 확인: `statistics_historical_data.json`의 `quarterly_range.end`가 `2025.1/4`로 확인
  - 동적 추출 테스트: 기초자료 파일이 있을 때 `StatisticsTableGenerator`가 올바르게 데이터 추출함 확인
  - 근본 원인 파악: 디버그 모드에서 세션에 `raw_excel_path`가 없어 동적 추출 불가
  - 해결책 결정: JSON 파일에 2025.2/4p 데이터를 직접 추가하여 영구적으로 해결
- **해결 방법**: 
  - `StatisticsTableGenerator`를 사용하여 기초자료에서 10개 통계표의 2025.2/4p 데이터 추출
  - `statistics_historical_data.json` 파일에 추출된 데이터 자동 저장
  - 메타데이터의 `quarterly_range.end`가 `2025.2/4`로 업데이트됨
- **관련 파일**: 
  - `templates/statistics_historical_data.json` (업데이트)
  - `debug/20260101_151713_statistics_2025Q2.html` (문제 파일)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 새 분기 데이터 추가 시 `StatisticsTableGenerator.extract_and_save_all()` 메서드 사용 권장
  - 또는 기초자료 파일 경로를 세션에 설정하면 동적 추출 가능

---

### 2025-01-01

#### 디버그 로그 시스템 구축
- **시간**: 초기 설정
- **문제 설명**: 디버그 작업 추적 시스템이 없음
- **원인 분석**: 디버그 작업이 체계적으로 기록되지 않음
- **해결 방법**: `DEBUG_LOG.md` 파일 생성 및 추적 시스템 구축
- **관련 파일**: 
  - `docs/DEBUG_LOG.md` (신규 생성)
- **상태**: ✅ 완료
- **참고 사항**: 앞으로 모든 디버그 작업은 이 파일에 기록됩니다.

---

## 📊 통계

### 전체 디버그 항목 수
- 총 항목: 9
- 완료: 9
- 진행중: 0
- 실패: 0
- 보류: 0

### 최근 활동
- 마지막 업데이트: 2026-01-01 (현재)

---

## 🔍 빠른 검색

### 카테고리별 분류
- [ ] 데이터 처리 오류
- [ ] UI/UX 문제
- [ ] 성능 이슈
- [ ] 의존성 문제
- [ ] 설정/환경 문제
- [ ] 기타

---

## 📝 사용 방법

새로운 디버그 항목을 추가할 때는 다음 템플릿을 사용하세요:

```markdown
### YYYY-MM-DD

#### [문제 제목]
- **시간**: HH:MM
- **문제 설명**: 
- **원인 분석**: 
- **에이전트 사고 과정**: 
  - 문제 인식: 
  - 고려한 해결책: 
  - 선택한 접근 방법: 
  - 시도한 방법들: 
  - 최종 결정 근거: 
- **해결 방법**: 
- **관련 파일**: 
  - `path/to/file1.py`
  - `path/to/file2.html`
- **상태**: 진행중 / 완료 / 실패 / 보류
- **참고 사항**: 
```

---

*이 문서는 프로젝트의 디버그 작업을 체계적으로 추적하기 위해 작성되었습니다.*

