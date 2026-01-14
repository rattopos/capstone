# 연도/분기 추출 문제 수정 체크리스트

## ✅ 완료된 수정 사항

### 1. `services/report_generator.py`
- ✅ `generate_report_html`에서 Generator에 `year`, `quarter` 전달 로직 추가
- ✅ `report_info` 강제 추가/업데이트 로직 추가

### 2. `utils/excel_utils.py`
- ✅ `extract_year_quarter_from_data`에 기본값 지원 추가

### 3. `templates/infographic_generator.py`
- ✅ `__init__`에 `year`, `quarter` 파라미터 추가
- ✅ 파일명에서 자동 추출 로직 추가
- ✅ `generate_report_data`에서 `year`, `quarter` 전달

---

## ⚠️ 예상되는 누락 항목

### 1. Generator 클래스들 - `report_info` 누락 가능성

#### 1.1 `templates/construction_generator.py`
- **현재 상태**: `generate_report_data`가 `year`, `quarter`를 받지만 사용하지 않음
- **문제**: 반환 데이터에 `report_info`가 포함되지 않을 수 있음
- **확인 필요**: `generate_report_data` 반환값에 `report_info` 포함 여부 확인
- **수정 방안**: 반환 데이터에 `report_info` 추가

#### 1.2 `templates/service_industry_generator.py`
- **현재 상태**: `generate_report_data`가 `year`, `quarter`를 받지만 사용하지 않음
- **문제**: 반환 데이터에 `report_info`가 포함되지 않을 수 있음
- **확인 필요**: `generate_report_data` 반환값에 `report_info` 포함 여부 확인
- **수정 방안**: 반환 데이터에 `report_info` 추가

#### 1.3 `templates/unemployment_generator.py`
- **현재 상태**: `report_info`를 포함하지만 `year`, `quarter`가 없음
- **문제**: `report_info`에 `year`, `quarter` 필드가 누락됨
- **확인 필요**: 반환 데이터의 `report_info` 구조 확인
- **수정 방안**: `report_info`에 `year`, `quarter` 추가

#### 1.4 `templates/export_generator.py`
- **현재 상태**: `generate_report_data`가 `year`, `quarter`를 받지만 사용하지 않음
- **확인 필요**: 반환 데이터에 `report_info` 포함 여부 확인
- **수정 방안**: 반환 데이터에 `report_info` 추가

#### 1.5 `templates/import_generator.py`
- **현재 상태**: `generate_report_data`가 `year`, `quarter`를 받지만 사용하지 않음
- **확인 필요**: 반환 데이터에 `report_info` 포함 여부 확인
- **수정 방안**: 반환 데이터에 `report_info` 추가

#### 1.6 `templates/price_trend_generator.py`
- **현재 상태**: `report_info`를 포함하지만 `year`, `quarter` 확인 필요
- **확인 필요**: 반환 데이터의 `report_info`에 `year`, `quarter` 포함 여부 확인
- **수정 방안**: 필요시 `year`, `quarter` 추가

#### 1.7 `templates/domestic_migration_generator.py`
- **현재 상태**: `generate_report_data`가 `year`, `quarter`를 받지만 `DomesticMigrationGenerator`에 전달하지 않음
- **문제**: `DomesticMigrationGenerator.__init__`이 `year`, `quarter`를 받지 않음
- **수정 방안**: 
  - `DomesticMigrationGenerator.__init__`에 `year`, `quarter` 파라미터 추가
  - `generate_report_data`에서 Generator에 전달
  - 반환 데이터에 `report_info` 추가

### 2. Generator 클래스들 - `__init__` 파라미터 누락

#### 2.1 `templates/mining_manufacturing_generator.py`
- **현재 상태**: `광공업생산Generator.__init__`이 `year`, `quarter`를 받지 않음
- **문제**: 클래스 기반 Generator 사용 시 `year`, `quarter` 전달 불가
- **수정 방안**: `__init__`에 `year`, `quarter` 파라미터 추가 (선택사항)

#### 2.2 `templates/regional_generator.py`
- **현재 상태**: `RegionalGenerator.__init__`이 `year`, `quarter`를 받지 않음
- **문제**: 클래스 기반 Generator 사용 시 `year`, `quarter` 전달 불가
- **수정 방안**: `__init__`에 `year`, `quarter` 파라미터 추가 (선택사항)

#### 2.3 `templates/statistics_table_generator.py`
- **현재 상태**: `StatisticsTableGenerator.__init__`이 `current_year`, `current_quarter`를 받음 (이미 구현됨)
- **상태**: ✅ 정상

#### 2.4 `templates/reference_grdp_generator.py`
- **현재 상태**: `참고_GRDP_Generator.__init__`이 `year`, `quarter`를 받지 않음
- **문제**: `generate_report_data`에서 `year`, `quarter`를 받지만 Generator에 전달하지 않음
- **수정 방안**: `__init__`에 `year`, `quarter` 파라미터 추가

### 3. API 엔드포인트 - 기본값 처리

#### 3.1 `routes/preview.py`
- **현재 상태**: `extract_year_quarter_from_data` 호출 시 기본값을 전달하지 않음
- **문제**: 추출 실패 시 예외 발생으로 미리보기 실패 가능
- **수정 방안**: 기본값 전달 (`default_year=2025, default_quarter=2`)

```python
# 현재 (77번 줄)
year, quarter = extract_year_quarter_from_data(excel_path)

# 수정 제안
year, quarter = extract_year_quarter_from_data(excel_path, default_year=2025, default_quarter=2)
```

#### 3.2 `routes/preview.py` - 다른 호출 지점
- **확인 필요**: `extract_year_quarter_from_data` 호출하는 다른 위치들도 기본값 전달 필요

### 4. `extract_year_quarter_from_excel` - 기본값 지원

#### 4.1 `utils/excel_utils.py`
- **현재 상태**: `extract_year_quarter_from_excel`이 기본값을 지원하지 않음
- **문제**: 추출 실패 시 예외 발생
- **수정 방안**: `extract_year_quarter_from_data`와 동일하게 기본값 지원 추가

### 5. 템플릿에서 `report_info` 사용 확인

#### 5.1 모든 템플릿 파일
- **확인 필요**: 템플릿에서 `{{ report_info.year }}`, `{{ report_info.quarter }}` 사용 여부
- **영향**: `report_info`가 없으면 템플릿 렌더링 오류 가능
- **해결**: `services/report_generator.py`에서 강제 추가하므로 대부분 해결됨

---

## 우선순위별 수정 계획

### 🔴 높은 우선순위 (즉시 수정 필요)

1. **`routes/preview.py`** - 기본값 전달 추가
   - 미리보기 기능이 실패하지 않도록 보장

2. **Generator들의 `report_info` 추가**
   - `construction_generator.py`
   - `service_industry_generator.py`
   - `export_generator.py`
   - `import_generator.py`
   - `domestic_migration_generator.py`

3. **`unemployment_generator.py`** - `report_info`에 `year`, `quarter` 추가

### 🟡 중간 우선순위 (단기 수정)

4. **Generator 클래스 `__init__` 파라미터 추가**
   - `mining_manufacturing_generator.py`
   - `regional_generator.py`
   - `reference_grdp_generator.py`
   - `domestic_migration_generator.py`

5. **`extract_year_quarter_from_excel` 기본값 지원**
   - `utils/excel_utils.py`

### 🟢 낮은 우선순위 (장기 개선)

6. **템플릿 검증**
   - 모든 템플릿에서 `report_info` 사용 확인
   - 누락된 경우 오류 처리 추가

---

## 검증 방법

### 1. 단위 테스트
```python
# 각 Generator의 generate_report_data 호출 시 report_info 확인
data = generator.generate_report_data(excel_path, year=2025, quarter=3)
assert 'report_info' in data
assert 'year' in data['report_info']
assert 'quarter' in data['report_info']
```

### 2. 통합 테스트
- 실제 엑셀 파일로 보도자료 생성 테스트
- 템플릿 렌더링 시 `report_info.year`, `report_info.quarter` 사용 확인

### 3. 로그 확인
- `[DEBUG] report_info 설정:` 로그 확인
- 모든 보도자료 생성 시 `report_info`가 포함되는지 확인

---

## 참고 사항

1. **데이터 무결성 원칙**: 기본값 사용 시에도 명시적으로 표시해야 함
2. **하위 호환성**: 기존 코드와의 호환성 유지 필요
3. **에러 처리**: 추출 실패 시 적절한 폴백 로직 필요
