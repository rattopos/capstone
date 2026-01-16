# Generator 동적 매핑 마이그레이션 계획

## 📊 현황 분석

### Generator별 하드코딩 패턴 통계

| Generator | 하드코딩 패턴 수 | 우선순위 | 상태 |
|-----------|------------------|----------|------|
| `mining_manufacturing_generator.py` | 69 | ✅ | **완료** (25년 3분기 검증 완료) |
| `regional_generator.py` | 62 | 🔴 높음 | 대기 |
| `export_generator.py` | 56 | 🔴 높음 | 대기 |
| `service_industry_generator.py` | 55 | 🔴 높음 | 대기 |
| `consumption_generator.py` | 54 | 🔴 높음 | 대기 |
| `construction_generator.py` | 52 | 🔴 높음 | 대기 |
| `employment_rate_generator.py` | 50 | 🟡 중간 | 대기 |
| `price_trend_generator.py` | 41 | 🟡 중간 | 대기 |
| `import_generator.py` | 39 | 🟡 중간 | 대기 |
| `domestic_migration_generator.py` | 12 | 🟢 낮음 | 대기 |
| `unemployment_generator.py` | 6 | 🟢 낮음 | 대기 |
| `base_generator.py` | 1 | - | 공통 기반 (이미 완료) |
| `infographic_generator.py` | 1 | - | 차트/인포그래픽 (자동화 미사용) |
| `statistics_table_generator.py` | 0 | - | 테이블 전용 |

---

## 🎯 마이그레이션 전략

### Phase 1: 주요 보고서 (5개) - 🔴 높은 우선순위

1. **service_industry_generator.py** (서비스업생산)
   - 시트: `B(서비스업생산)집계`, `B 분석`
   - 패턴: 55개
   - 중요도: ⭐⭐⭐⭐⭐
   - 예상 작업: 2-3시간

2. **consumption_generator.py** (소비동향)
   - 시트: `C(소비)집계`, `C 분석`
   - 패턴: 54개
   - 중요도: ⭐⭐⭐⭐⭐
   - 예상 작업: 2-3시간

3. **construction_generator.py** (건설수주)
   - 시트: `F'(건설)집계`, `F'분석`
   - 패턴: 52개
   - 중요도: ⭐⭐⭐⭐
   - 예상 작업: 2-3시간

4. **export_generator.py** (수출)
   - 시트: `G(수출)집계`, `G 분석`
   - 패턴: 56개
   - 중요도: ⭐⭐⭐⭐
   - 예상 작업: 2-3시간

5. **regional_generator.py** (지역별 종합)
   - 시트: 여러 시트 통합
   - 패턴: 62개 (가장 많음)
   - 중요도: ⭐⭐⭐⭐⭐
   - 예상 작업: 3-4시간

**Phase 1 총 예상 시간**: 11-16시간

---

### Phase 2: 보조 지표 (3개) - 🟡 중간 우선순위

6. **employment_rate_generator.py** (고용률)
   - 시트: `D(고용률)집계`, `D(고용률)분석`
   - 패턴: 50개
   - 중요도: ⭐⭐⭐
   - 예상 작업: 2시간

7. **price_trend_generator.py** (물가)
   - 시트: `E(품목성질물가)집계`, `E(품목성질물가)분석`
   - 패턴: 41개
   - 중요도: ⭐⭐⭐
   - 예상 작업: 2시간

8. **import_generator.py** (수입)
   - 시트: `H(수입)집계`, `H 분석`
   - 패턴: 39개
   - 중요도: ⭐⭐⭐
   - 예상 작업: 2시간

**Phase 2 총 예상 시간**: 6시간

---

### Phase 3: 소규모 지표 (2개) - 🟢 낮은 우선순위

9. **domestic_migration_generator.py** (인구이동)
   - 시트: `I(순인구이동)집계`
   - 패턴: 12개
   - 중요도: ⭐⭐
   - 예상 작업: 1시간

10. **unemployment_generator.py** (실업)
    - 시트: `D(실업)집계`, `D(실업)분석`
    - 패턴: 6개
    - 중요도: ⭐⭐
    - 예상 작업: 30분

**Phase 3 총 예상 시간**: 1.5시간

---

## 🔧 마이그레이션 작업 항목 (각 Generator 공통)

### 1. 시트 로드 개선
```python
# ❌ Before
df = pd.read_excel(excel_path, sheet_name='B 분석')

# ✅ After
sheet_name, use_raw = self.find_sheet_with_fallback(
    ['B 분석', 'B(서비스업생산)집계'],
    ['서비스업생산', '서비스업']
)
df = self.get_sheet(sheet_name)
```

### 2. 헤더 행 동적 탐색
```python
# ❌ Before
header_row = df.iloc[2]

# ✅ After
header_row_idx = self.find_header_row(df, keywords=['지역', '산업', '2025'])
header_row = df.iloc[header_row_idx]
```

### 3. 컬럼 인덱스 동적 탐색
```python
# ❌ Before
growth_rate_col = 21

# ✅ After
growth_rate_col = self.find_target_col_index(header_row, self.year, self.quarter)
```

### 4. 메타데이터 컬럼 동적 탐색
```python
# ❌ Before
region_name_col = 4
industry_name_col = 8

# ✅ After
region_name_col = self._find_metadata_column('region')
industry_name_col = self._find_metadata_column('industry_name')
```

---

## 📋 체크리스트 (각 Generator마다)

- [ ] `__init__` 메서드에서 하드코딩된 컬럼 상수 제거
- [ ] `_find_metadata_column` 메서드 구현 또는 활용
- [ ] `find_target_col_index` 사용하여 연도/분기 컬럼 동적 탐색
- [ ] `find_sheet_with_fallback` 사용하여 시트명 Fallback 구현
- [ ] 모든 `row[숫자]` 패턴을 `row[동적_변수]`로 교체
- [ ] 테스트: 25년 3분기 데이터 정확성 검증
- [ ] 로그 추가: SmartSearch 스타일 디버그 로그
- [ ] 문서화: 코드 주석 및 docstring 업데이트

---

## 🚀 실행 계획

### 옵션 A: 전체 마이그레이션 (권장)
**목표**: 모든 generator를 동적 매핑으로 전환하여 완전한 유연성 확보

**장점**:
- 엑셀 구조 변경에 완전히 대응 가능
- 유지보수 용이
- 2025년 4분기, 2026년 1분기 등 자동 적응

**단점**:
- 총 작업 시간: **18.5-23.5시간**
- 테스트 및 검증 시간 추가 필요

**진행 방식**:
1. Phase 1 완료 → 중간 검증
2. Phase 2 완료 → 중간 검증
3. Phase 3 완료 → 최종 통합 테스트

---

### 옵션 B: 선택적 마이그레이션 (빠른 적용)
**목표**: 가장 중요한 5개 generator만 우선 처리

**대상**:
- service_industry_generator.py
- consumption_generator.py
- construction_generator.py
- export_generator.py
- regional_generator.py

**장점**:
- 총 작업 시간: **11-16시간** (40% 단축)
- 핵심 기능 빠른 안정화

**단점**:
- 나머지 generator는 여전히 하드코딩 상태
- 일부 지표는 엑셀 구조 변경 시 수동 수정 필요

---

### 옵션 C: 점진적 마이그레이션 (위험 최소화)
**목표**: 한 번에 1-2개씩 마이그레이션하고 프로덕션 검증 후 다음 진행

**진행 방식**:
1. Week 1: service_industry_generator.py → 프로덕션 테스트
2. Week 2: consumption_generator.py → 프로덕션 테스트
3. Week 3-4: 나머지 generator 순차 진행

**장점**:
- 위험 최소화
- 각 단계에서 피드백 반영 가능
- 예기치 않은 문제 조기 발견

**단점**:
- 총 완료 기간: **4-6주**
- 관리 오버헤드 증가

---

## 🎯 권장 사항

### 즉시 실행 (현재 세션)
1. **service_industry_generator.py** 마이그레이션
   - 가장 높은 우선순위 보고서 중 하나
   - mining_manufacturing_generator와 유사한 구조
   - 예상 시간: 2-3시간

### 단기 계획 (이번 주)
2-3. **consumption_generator.py**, **construction_generator.py** 마이그레이션
   - Phase 1의 나머지 핵심 generator
   - 예상 시간: 4-6시간

### 중기 계획 (다음 주)
4-8. Phase 2 완료
   - 보조 지표 마이그레이션
   - 예상 시간: 6시간

### 장기 계획 (필요 시)
9-10. Phase 3 완료
   - 소규모 지표 마이그레이션
   - 예상 시간: 1.5시간

---

## 📊 성공 지표

각 generator 마이그레이션 후:

1. **정확성 검증**: 25년 3분기 데이터 추출 결과가 엑셀 원본과 일치
2. **로그 확인**: SmartSearch 로그에서 컬럼 자동 감지 확인
3. **에러 없음**: 전체 데이터 추출 과정에서 AttributeError 없음
4. **Fallback 작동**: 시트명이 다를 때 자동으로 대체 시트 찾기
5. **문서화**: 각 generator의 동적 매핑 전략 문서화

---

## 💡 다음 단계

**선택해주세요**:

A. **지금 바로 시작** - service_industry_generator.py 마이그레이션 (2-3시간)
B. **Phase 1 전체** - 5개 주요 generator 일괄 마이그레이션 (11-16시간)
C. **계획 수정** - 다른 우선순위나 전략 제안

어떤 옵션을 선택하시겠습니까?
