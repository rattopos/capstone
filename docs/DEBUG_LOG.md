# 디버그 로그

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

## 2026-01-21: 건설 동향 토목/건축 증감률 동일 표시 버그 수정

### 문제 설명
- 건설 동향 보고서의 '주요 증감지역 및 공종' 영역에서 각 지역별로 토목과 건축의 비율이 동일하게 표시됨
- 예: 충북(104.4%) - 토목(104.4%), 건축(104.4%) → 둘 다 전체 증감률과 동일

### 원인 분석
- `templates/unified_generator.py`에서 `civil_growth`와 `building_growth`가 전체 `growth_rate`와 동일하게 설정되고 있었음
- 관련 위치:
  1. 1491-1492줄: 전국 데이터에서 토목/건축 증감률 설정
  2. 1928-1934줄: 지역별 데이터에서 토목/건축 증감률 설정
  3. 2038-2041줄: `_enrich_template_data`에서 각 항목 보정

### 에이전트 사고 과정
1. 문제 인식: 사용자 제공 이미지에서 모든 지역의 토목/건축 증감률이 전체 증감률과 동일함을 확인
2. 코드 분석: `construction_template.html`에서 `region.civil_growth`와 `region.building_growth` 사용 확인
3. 근본 원인 탐색: `unified_generator.py`에서 해당 값들이 모두 `growth_rate`로 설정되고 있음 확인
4. 해결 방안 설계:
   - 방안 A: 엑셀에서 토목/건축 데이터를 별도로 추출하는 로직 추가 (선택)
   - 방안 B: 하드코딩된 값 사용 (기본값/폴백 사용 금지 규칙 위반)
5. 구현: 토목/건축 증감률을 `_extract_industry_data`를 통해 실제 데이터에서 추출하는 헬퍼 함수 추가

### 해결 방법
1. `_get_civil_building_growth(region)` 헬퍼 함수 추가
   - 특정 지역의 토목/건축 증감률을 업종 데이터에서 추출
   - 매칭 기준: '토목', '건축' (name_mapping에서 이미 매핑됨)
2. 전국 데이터 처리 수정 (1521-1531줄)
   - 전국의 토목/건축 증감률을 실제 데이터에서 추출
3. 지역별 데이터 처리 수정 (1963-1993줄)
   - 각 지역의 토목/건축 증감률을 실제 데이터에서 추출
4. `_enrich_template_data` 수정 (2096-2116줄)
   - 최종 보정 시에도 실제 데이터 사용

### 관련 파일 목록
- `templates/unified_generator.py` (수정)
- `config/reports.py` (참조)
- `templates/construction_template.html` (참조)

### 작업 상태
완료

### 참고 사항
- 엑셀 집계 시트에서 토목과 건축은 `industry_name_col`(컬럼 4)에 저장됨
- `name_mapping`에서 '토목' → '토목', '건축' → '건축'으로 그대로 매핑됨
- 실제 엑셀 데이터가 없으면 fallback으로 전체 증감률 사용 (데이터 무결성 유지)
