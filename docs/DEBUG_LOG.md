# 디버그 로그

## 2026-01-21: 고용률/실업률 연령별 데이터 미표시 및 테이블 빈칸 버그 수정

### 문제 설명
1. 고용률/실업률 챕터의 '주요 등락지역 및 연령별 고용률/실업률' 섹션에서 연령별 데이터가 표시되지 않음
2. 테이블에서 일부 셀(2025. 2/4 증감, 청년층 비율 등)이 빈칸으로 표시됨
3. 전국의 연령별 데이터가 비어있음

### 원인 분석
1. **컬럼 인덱스 하드코딩 문제**:
   - `_extract_age_groups_for_region` 함수에서 `region_col = 0`, `age_col = 1`로 하드코딩
   - 실제 고용률 엑셀 구조: 지역명=컬럼 1, 연령=컬럼 3 (산업 이름 컬럼)
   - 실업률 엑셀 구조: 지역명=컬럼 0, 연령=컬럼 1

2. **data_start_row 설정 문제**:
   - `aggregation_structure.data_start_row` 값이 0일 때 Falsy로 처리되어 무시됨
   - `or` 연산자 대신 명시적인 `is not None` 체크 필요

3. **youth_rate 추출 문제**:
   - `_select_youth_rate` 함수에서도 컬럼 인덱스가 하드코딩
   - 고용률에서 '20-29세'를 찾으려 했으나, 실제 데이터에는 '15-29세'만 존재

4. **nationwide_data.age_groups 미설정**:
   - 템플릿에서 `nationwide_data.age_groups` 사용하지만, 코드에서는 `top_age_groups`만 설정

### 에이전트 사고 과정
1. **문제 인식**: 사용자 제공 스크린샷에서 연령별 데이터와 테이블 빈칸 확인
2. **엑셀 구조 분석**:
   - D(고용률)집계: 컬럼 0=지역코드, 컬럼 1=지역명, 컬럼 3=연령
   - D(실업)집계: 컬럼 0=시도별, 컬럼 1=연령계층별
3. **코드 분석**:
   - `_extract_age_groups_for_region`에서 하드코딩된 인덱스 발견
   - `_select_youth_rate`에서도 동일한 문제 발견
   - `data_start_row` 로직에서 0값 처리 오류 발견
4. **단계별 수정**:
   - config/reports.py에서 올바른 컬럼 인덱스 설정
   - unified_generator.py에서 설정값 사용하도록 변경
   - 0값도 유효하게 처리하도록 로직 수정
   - age_groups 키 추가

### 해결 방법
**config/reports.py 수정**:
- 고용률: `region_name_col: 1`, `age_col: 3`, `data_start_row: 0`
- 실업률: `region_name_col: 0`, `age_col: 1`, `data_start_row: 1`

**templates/unified_generator.py 수정**:
1. `_extract_age_groups_for_region`: 설정에서 컬럼 인덱스 가져오기
2. `_select_youth_rate`: 설정에서 컬럼 인덱스 가져오기, 15-29세 패턴 사용
3. `data_start_row` 설정 시 `is not None` 체크 추가
4. `nationwide['age_groups']` 설정 추가

### 관련 파일 목록
- `config/reports.py` (수정)
- `templates/unified_generator.py` (수정)
- `templates/employment_template.html` (참조)
- `templates/unemployment_template.html` (참조)

### 작업 상태
완료

### 참고 사항
- 일부 지역의 2025. 2/4 분기 전년동분기대비 증감 데이터가 빈칸으로 남아있음
- 이는 직전 분기의 전년동분기대비 증감 계산 로직이 별도로 필요함
- 주요 문제(연령별 데이터 표시, 청년층 비율)는 해결됨

---

## 2026-01-21: 수출/수입 '주요 증감지역 및 품목'에서 품목 코드가 이름 대신 표시되는 버그 수정

### 문제 설명
- 수출/수입 동향 챕터의 '주요 증감지역 및 품목' 박스에서 품목 코드(숫자)가 이름 대신 표시됨
- 예시: `·전 국( 6.5%): 43301(151.0%), 44202(150.0%), 41301(149.0%)`
- 정상: `·전 국( 6.5%): 기타 일반기계(151.0%), 컴퓨터 주변기기(150.0%), ...`

### 원인 분석
- `_extract_trade_product_data` 함수에서 `analysis_structure.industry_name_col`이 7로 설정됨
- 집계 시트(`G(수출)집계`)와 분석 시트(`G 분석`)의 열 구조가 다름:
  - **집계 시트**: 열 6=상품코드, 열 7=상품이름
  - **분석 시트**: 열 6=순위, 열 7=상품코드, 열 8=상품이름
- 분석 시트에서 `industry_name_col=7`로 설정되어 있어 상품코드를 가져옴

### 에이전트 사고 과정
1. **문제 인식**: 사용자 제공 스크린샷에서 숫자 코드(43301 등)가 품목 이름 대신 표시됨
2. **데이터 흐름 분석**:
   - `_extract_trade_product_data`에서 `analysis_structure.industry_name_col`을 사용
   - 분석 시트(G 분석, H 분석)에서 품목 데이터를 추출
3. **엑셀 구조 확인**:
   - `G 분석` 시트: 열 7=`상품\n코드`, 열 8=`상품 이름`
   - `H 분석` 시트: 열 7=`상품\n코드`, 열 8=`상품 이름`
4. **설정 분석**:
   - 수출: `analysis_structure.industry_name_col: 7` → 상품코드를 가져옴 (오류)
   - 수입: `analysis_structure.industry_name_col: 7` → 상품코드를 가져옴 (오류)
5. **해결 방안**: `industry_name_col`을 8로 변경하여 실제 상품 이름을 가져옴

### 해결 방법
**config/reports.py 수정**:
- 수출 `analysis_structure.industry_name_col`: 7 → 8
- 수입 `analysis_structure.industry_name_col`: 7 → 8

### 관련 파일 목록
- `config/reports.py` (수정)
- `templates/unified_generator.py` (참조)

### 작업 상태
완료

### 참고 사항
- 집계 시트와 분석 시트의 열 구조가 다르므로 각각 별도의 설정 필요
- 수출/수입의 `aggregation_structure.industry_name_col`은 7로 유지 (집계 시트에서는 열 7이 상품 이름)
- 품목 코드와 이름의 매핑은 엑셀 파일 내에서 직접 참조하므로 별도의 매핑 파일 불필요

---

## 2026-01-21: 물가 동향 챕터 데이터 불일치 및 품목 미표시 버그 수정

### 문제 설명
1. **내레이션과 표 데이터 불일치**: 부산의 증감률이 내레이션에서는 2.1%, 표에서는 2.2%로 다르게 표시됨
2. **'주요 등락지역 및 품목'에서 지역별 품목 미표시**: 경남, 부산, 울산 등 개별 지역의 품목이 비어있음
3. **내레이션 미완성**: "등이 줄어" 앞에 지역명이 누락됨 (low_regions가 비어있음)

### 원인 분석
1. **증감률 분류 오류**: 물가 동향에서 `increase_regions`와 `decrease_regions`가 `change_rate > 0` / `change_rate < 0` 기준으로 분류됨. 물가는 모든 지역이 양수 증가율이므로 `decrease_regions`가 비어있고, 이것이 `low_regions`로 매핑됨
2. **지역명 매칭 실패**: `_extract_industry_data` 함수에서 "부산"으로 검색하지만, 엑셀에서는 "부산광역시"로 저장됨. 또한 "경남"과 "경상남도"는 부분 문자열로도 매칭되지 않음
3. **표/내레이션 계산 방식 차이**: 
   - 내레이션: 원본 지수 값 (116.78, 114.34) 기반 계산 → 2.1%
   - 표: 반올림된 지수 값 (116.8, 114.3) 기반 계산 → 2.2%

### 에이전트 사고 과정
1. **문제 인식**: 사용자 제공 이미지에서 세 가지 문제 확인
2. **데이터 흐름 분석**:
   - `extract_regional_data` → `_enrich_template_data`에서 `high_regions`/`low_regions` 매핑
   - `_extract_industry_data`에서 지역명 필터링 로직 확인
   - `_build_summary_table`에서 표 데이터 생성 로직 확인
3. **근본 원인 탐색**:
   - 물가는 "증가/감소"가 아닌 "전국보다 높음/낮음"으로 분류해야 함
   - 엑셀 지역명이 "부산광역시", "경상남도" 등 전체 이름으로 저장됨
   - 표와 내레이션이 다른 계산 방식 사용
4. **해결 방안 설계**:
   - 물가의 경우 전국 증감률과 비교하여 high/low 분류
   - 지역명 패턴 매핑 추가 (예: "경남" → "경상남|경남")
   - 표에서 `row.get('change_rate')`를 직접 사용하여 일관성 유지

### 해결 방법
1. **`_enrich_template_data` 수정 (물가 전국 대비 분류)**:
   - 물가의 경우 `nationwide_change` 값과 비교하여 `high_regions`와 `low_regions` 분류
   - 각 지역에 `change`와 `region` 필드 추가 (템플릿 호환)

2. **`_extract_industry_data` 수정 (지역명 패턴 매핑)**:
   - 정확한 매칭 실패 시 패턴 매칭 시도
   - `region_name_patterns` 딕셔너리로 짧은 지역명을 전체 지역명 패턴으로 변환
   - 예: `'경남': '경상남|경남'`, `'충북': '충청북|충북'`

3. **`_extract_industry_data` 수정 (품목 정렬)**:
   - 추출된 품목을 변화율 절대값 기준으로 정렬 (영향력 높은 품목 우선)

4. **`_build_summary_table._build_growth_slots` 수정 (표/내레이션 일치)**:
   - 현재 분기 증감률로 `row.get('change_rate')`를 우선 사용
   - 계산된 값이 아닌 원본 증감률 사용으로 일관성 유지

### 관련 파일 목록
- `templates/unified_generator.py` (수정)
- `templates/price_template.html` (참조)
- `config/reports.py` (참조)

### 작업 상태
완료

### 참고 사항
- 물가 집계 시트(`E(품목성질물가)집계`)에서 지역명은 "서울특별시", "부산광역시", "경상남도" 등 전체 이름으로 저장됨
- 품목 데이터는 변화율 절대값 기준으로 정렬되어 영향력이 큰 품목이 먼저 표시됨
- 내레이션과 표가 동일한 `change_rate` 값을 사용하여 데이터 일관성 유지

---

## 2026-01-21: 수출/수입 '주요 증감지역 및 품목' 품목 미표시 버그 수정

### 문제 설명
- 수출/수입 동향 챕터의 '주요 증감지역 및 품목' 표에서 지역 이름 옆에 품목이 표시되지 않음
- 예: `·서울(-2.8%):` 로만 나오고, 품목이 비어있음
- 정상: `·서울(-2.8%): 반도체(-5.2%), 자동차(-3.1%), ...`

### 원인 분석
- `_extract_industry_data` 함수가 수출/수입에서 빈 배열을 반환
- 수출/수입의 품목별 데이터가 집계 시트가 아닌 분석 시트에 존재
- 분석 시트가 로드되지 않아 품목 데이터를 추출할 수 없었음
- 또한 기여율(contribution_rate) 데이터가 템플릿에 전달되지 않음

### 에이전트 사고 과정
1. **문제 인식**: 사용자 제공 이미지에서 품목이 비어있음을 확인
2. **템플릿 분석**: `export_template.html`에서 `region.products`를 사용하여 품목 렌더링
3. **데이터 흐름 추적**:
   - `extract_all_data` → `_enrich_template_data` → `ensure_products`
   - `ensure_products`에서 `_extract_industry_data` 호출
   - `_extract_industry_data`가 빈 배열 반환
4. **설정 분석**:
   - 수출: `sheet: 'G 분석'`, `aggregation_structure.sheet: 'G(수출)집계'`
   - 분석 시트(`G 분석`)에 품목별 데이터가 있고, 집계 시트에는 지역별 합계만 있음
   - `analysis_sheet` 설정이 없어 분석 시트가 로드되지 않음
5. **해결 방안 설계**:
   - 수출/수입 설정에 `analysis_sheet` 추가
   - 수출/수입 전용 품목 추출 함수 `_extract_trade_product_data` 추가
   - 기여율 데이터를 템플릿에 전달하도록 수정

### 해결 방법
1. **config/reports.py 수정**:
   - 수출 설정에 `analysis_sheet: 'G 분석'` 추가
   - 수입 설정에 `analysis_sheet: 'H 분석'` 추가
   - 분석 시트 구조 정보 `analysis_structure` 추가 (region_name_col, industry_name_col, contribution_col)

2. **templates/unified_generator.py 수정**:
   - `_extract_trade_product_data` 함수 추가: 수출/수입 품목별 기여율 데이터 추출
   - `_extract_industry_data`에서 수출/수입인 경우 `_extract_trade_product_data` 호출
   - `_enrich_template_data`에서 품목 정규화 시 `contribution_rate` 포함
   - 지역별 데이터에서 품목이 없으면 `_extract_industry_data` 재호출

### 관련 파일 목록
- `config/reports.py` (수정)
- `templates/unified_generator.py` (수정)
- `templates/export_template.html` (참조)
- `templates/import_template.html` (참조)

### 작업 상태
완료

### 참고 사항
- 수출/수입의 품목별 데이터 구조가 다른 부문(광공업, 서비스업 등)과 다름
- 수출/수입: 분석 시트에 지역+품목별 상세 데이터, 기여율 포함
- 다른 부문: 집계 시트에 지역+산업별 데이터
- 템플릿에서 `contribution_rate`가 있으면 우선 표시, 없으면 `change_rate` 표시

---

## 2026-01-21: 수출/수입 템플릿 딕셔너리 속성 접근 오류 수정

### 문제 설명
- 수출(export) 및 수입(import) 보도자료 생성 시 템플릿 렌더링 오류 발생
- 오류 메시지: `'dict object' has no attribute 'contribution_rate'`
- `export_template.html` 338줄, `import_template.html` 334줄에서 발생

### 원인 분석
- Jinja2 템플릿에서 딕셔너리 객체에 점(`.`) 표기법으로 존재하지 않는 키에 접근 시도
- `prod.contribution_rate if prod.contribution_rate is not none` 형태의 조건 체크에서도 속성 접근이 먼저 시도되어 `UndefinedError` 발생
- 딕셔너리에 `contribution_rate` 키가 없을 때 조건 체크 자체가 실패함

### 에이전트 사고 과정
1. 문제 인식: 터미널 로그에서 `'dict object' has no attribute 'contribution_rate'` 오류 확인
2. 오류 위치 추적: 스택 트레이스에서 `export_template.html` 338줄, `import_template.html` 334줄 확인
3. 템플릿 코드 분석:
   - `{% set rate = prod.contribution_rate if prod.contribution_rate is not none else ... %}` 형태의 코드 확인
   - Jinja2에서 딕셔너리에 점 표기법으로 접근 시 키가 없으면 오류 발생
4. 해결 방안 검토:
   - 방안 A: `.get()` 메서드 사용으로 안전하게 접근 (선택)
   - 방안 B: `|default(none)` 필터 사용 (코드가 더 복잡해짐)
5. 구현: 모든 딕셔너리 속성 접근을 `.get()` 메서드로 변경

### 해결 방법
1. `export_template.html` 수정 (3개 위치):
   - 전국 데이터 품목 표시 부분 (334-342줄)
   - 증가 지역 Top 3 품목 표시 부분 (348-358줄)
   - 감소 지역 Top 3 품목 표시 부분 (363-375줄)
2. `import_template.html` 수정 (3개 위치):
   - 전국 데이터 품목 표시 부분 (330-338줄)
   - 증가 지역 Top 3 품목 표시 부분 (344-354줄)
   - 감소 지역 Top 3 품목 표시 부분 (359-371줄)
3. 변경 내용:
   - `prod.contribution_rate` → `prod.get('contribution_rate')`
   - `prod.contribution` → `prod.get('contribution')`
   - `prod.change` → `prod.get('change')`
   - `prod.growth_rate` → `prod.get('growth_rate')`
   - `prod.name` → `prod.get('name', prod)`

### 관련 파일 목록
- `templates/export_template.html` (수정)
- `templates/import_template.html` (수정)

### 작업 상태
완료

### 참고 사항
- Jinja2에서 딕셔너리 키 접근 시 `.get()` 메서드를 사용하면 키가 없어도 `None`을 반환하여 오류 방지
- 동일한 패턴이 다른 템플릿에도 있을 수 있으므로 향후 검토 필요

---

## 2026-01-21: 건설 동향 토목/건축 증감률 및 세부 공종 동일 표시 버그 수정

### 문제 설명
- 건설 동향 보고서의 '주요 증감지역 및 공종' 영역에서 두 가지 문제 발생:
  1. 각 지역별로 토목과 건축의 **증감률**이 동일하게 표시됨
     - 예: 충북(104.4%) - 토목(104.4%), 건축(104.4%) → 둘 다 전체 증감률과 동일
  2. 각 지역별로 토목과 건축의 **세부 공종**이 동일하게 표시됨
     - 예: 모든 지역에서 토목 ⇒ "철도·궤도, 기계설치", 건축 ⇒ "주택, 관공서 등"

### 원인 분석
- `templates/unified_generator.py`에서 `civil_growth`, `building_growth`, `civil_subtypes`, `building_subtypes`가 모두 하드코딩된 기본값으로 설정되고 있었음
- 관련 위치:
  1. 1491-1492줄: 전국 데이터에서 토목/건축 증감률 설정
  2. 1530-1531줄: 전국 데이터에서 세부 공종 하드코딩
  3. 1905-1906줄: nationwide 데이터 보강에서 세부 공종 하드코딩
  4. 1928-1934줄: 지역별 데이터에서 토목/건축 증감률 설정
  5. 2038-2041줄, 2114-2115줄: `_enrich_template_data`에서 세부 공종 하드코딩

### 에이전트 사고 과정
1. 문제 인식: 사용자 제공 이미지에서 모든 지역의 토목/건축 증감률과 세부 공종이 동일함을 확인
2. 코드 분석: `construction_template.html`에서 사용되는 필드 확인
   - `region.civil_growth`, `region.building_growth` (증감률)
   - `region.civil_subtypes`, `region.building_subtypes` (세부 공종)
3. 근본 원인 탐색: 
   - 증감률: 모두 `growth_rate`로 설정
   - 세부 공종: 하드코딩된 "철도·궤도, 기계설치", "주택, 관공서 등" 사용
4. 해결 방안 설계:
   - 방안 A: 엑셀에서 토목/건축 및 세부 공종 데이터를 별도로 추출하는 로직 추가 (선택)
   - 방안 B: 하드코딩된 값 사용 (기본값/폴백 사용 금지 규칙 위반)
5. 구현: `_get_civil_building_growth` 함수를 확장하여 세부 공종도 추출

### 해결 방법
1. `_get_civil_building_growth(region)` 헬퍼 함수 확장
   - 기존: 토목/건축 증감률만 추출
   - 확장: 세부 공종(`civil_subtypes`, `building_subtypes`)도 함께 추출
   - 세부 공종 키워드 매칭:
     - 토목: 철도, 궤도, 도로, 교량, 기계설치, 항만, 공항, 터널, 토지조성 등
     - 건축: 주택, 관공서, 사무실, 점포, 공장, 창고, 학교, 병원 등
   - 증감률 절대값 기준으로 상위 2개 세부 공종 선택
2. 전국 데이터 처리 수정
   - 토목/건축 증감률과 세부 공종 모두 실제 데이터에서 추출
3. 지역별 데이터 처리 수정
   - 각 지역의 토목/건축 증감률과 세부 공종 모두 실제 데이터에서 추출
4. `_enrich_template_data` 수정
   - 최종 보정 시에도 실제 데이터 사용

### 관련 파일 목록
- `templates/unified_generator.py` (수정)
- `config/reports.py` (참조)
- `templates/construction_template.html` (참조)

### 작업 상태
완료

### 참고 사항
- 엑셀 집계 시트에서 토목과 건축은 `industry_name_col`(컬럼 4)에 저장됨
- 세부 공종은 토목/건축 외의 다른 업종으로 저장되어 있음
- `name_mapping`에서 '토목' → '토목', '건축' → '건축'으로 그대로 매핑됨
- 세부 공종 키워드 매칭으로 토목/건축 하위 공종을 분류
- 실제 엑셀 데이터가 없으면 fallback으로 기본값 사용 (데이터 무결성 유지)
