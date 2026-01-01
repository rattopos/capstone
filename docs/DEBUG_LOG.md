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

#### Excel 수식 자동 계산 전처리 기능 추가
- **시간**: 18:30
- **문제 설명**: 분석표 엑셀 파일의 수식이 계산되지 않은 상태로 업로드되면 보고서에 0, NaN 등 잘못된 데이터가 표시됨
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
- **문제 설명**: 부문별 섹션에서 광공업생산, 건설동향, 실업률을 제외한 나머지 보고서 표에 0, NaN, 또는 행 누락 발생
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
- 총 항목: 6
- 완료: 6
- 진행중: 0
- 실패: 0
- 보류: 0

### 최근 활동
- 마지막 업데이트: 2026-01-01 18:30

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

