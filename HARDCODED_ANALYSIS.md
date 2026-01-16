# 하드코딩 현황 분석 보고서

## 🔍 분석 결과 요약

### 하드코딩 패턴 통계

| 파일 | `self.COL_*` 패턴 | `row[숫자]` 패턴 | 상태 |
|------|-------------------|------------------|------|
| `mining_manufacturing_generator.py` | 63개 | 55개 | ⚠️ 부분 개선 필요 |
| 기타 generator 파일들 | 미조사 | 미조사 | ❌ 조사 필요 |

---

## 📊 25년 3분기 데이터 매핑 현황

### ✅ 현재 작동하는 부분

1. **`find_target_col_index()` 메서드**: `base_generator.py`에 이미 구현됨
   - 동적으로 연도/분기 컬럼 찾기
   - 2025년 3분기 → `'2025'`, `'3/4'` 패턴 검색

2. **`SmartSearch` 로그 확인됨**:
   ```
   [SmartSearch] 2025년 3분기 데이터 열 탐색 시작...
   [SmartSearch] ✅ 발견! Index 26: '2025 3/4'
   ```

### ⚠️ 개선 필요한 부분

#### 1. `mining_manufacturing_generator.py`

**문제점**: 여전히 많은 `self.COL_*` 속성 참조가 남아있음 (63개)

**예시**:
```python
# 하드코딩된 속성 참조 (동적으로 설정되지만 여전히 하드코딩된 속성명)
self.COL_REGION_NAME
self.COL_CLASSIFICATION
self.COL_INDUSTRY_NAME
self.COL_INDUSTRY_CODE
self.COL_GROWTH_RATE
self.COL_CONTRIBUTION
self.COL_WEIGHT
```

**해결 방안**:
- 이 속성들은 `__init__()` 메서드에서 동적으로 설정되므로 **실제로는 문제 없음**
- 하지만 코드 가독성과 유지보수를 위해 다음과 같이 개선 가능:
  ```python
  # 현재
  industry_name = row[self.COL_INDUSTRY_NAME]
  
  # 개선안 (더 명확한 의도 표현)
  industry_name = self.get_cell(row, 'industry_name')
  ```

#### 2. 직접 인덱스 참조 (55개)

**문제점**: `row[15]`, `df.iloc[row_idx][21]` 같은 직접 숫자 인덱스

**위험도**:
- 🔴 **높음**: 엑셀 구조 변경 시 즉시 오류 발생
- 🔴 **높음**: 유지보수 어려움 (숫자만 보고 의미 파악 불가)

**발견 위치** (추정):
- fallback 로직 (기초자료 시트 처리)
- 레거시 코드
- 임시 하드코딩

**해결 방안**:
1. **우선순위 1**: 동적 헤더 탐색으로 교체
   ```python
   # ❌ Before
   value = row[21]
   
   # ✅ After
   col_idx = self.find_column_by_keyword(header_row, ['증감률', 'growth'])
   value = row[col_idx]
   ```

2. **우선순위 2**: 명명된 상수 사용
   ```python
   # ❌ Before
   value = row[21]
   
   # ✅ After (임시 완화책)
   GROWTH_RATE_COL = 21  # 문서화: 2025년 3분기 기준
   value = row[GROWTH_RATE_COL]
   ```

---

## 🎯 우선순위별 개선 계획

### Priority 1: 치명적 하드코딩 제거 (1-2시간)

**대상**: 연도/분기 데이터 컬럼 (가장 자주 변경됨)

**파일**:
1. `mining_manufacturing_generator.py`
2. `service_industry_generator.py`
3. `consumption_generator.py`
4. `construction_generator.py`
5. `export_generator.py`
6. `import_generator.py`

**작업**:
- 모든 `row[21]`, `row[26]` 같은 분기 데이터 접근을 `find_target_col_index()` 사용으로 교체

### Priority 2: 시트명 Fallback 강화 (30분)

**대상**: 시트 이름 하드코딩

**현재 상태**:
```python
# 일부 파일은 이미 구현됨
sheet_name, use_raw = find_sheet(['A 분석', 'A(광공업생산)집계', '광공업생산'])
```

**작업**:
- 모든 generator에 통일된 fallback 체계 적용
- `base_generator.py`에 공통 메서드 추가

### Priority 3: 헤더 행 자동 감지 (1시간)

**대상**: 헤더 행 위치 가정 (`row_idx = 2`)

**작업**:
- `find_header_row()` 메서드 구현
- 키워드 기반 헤더 탐색

### Priority 4: 리팩토링 (2-3시간)

**대상**: 코드 품질 개선

**작업**:
- `self.COL_*` → 헬퍼 메서드로 캡슐화
- 매직 넘버 제거
- 문서화 강화

---

## 📝 구체적 코드 예시

### 현재 문제 (mining_manufacturing_generator.py 385-396행 예시)

```python
def _extract_nationwide_industries_from_analysis(self) -> dict:
    df = self.df_analysis
    data_df = df.iloc[self.DATA_START_ROW:].copy()
    
    # 동적으로 컬럼 인덱스 찾기
    growth_rate_col = self._find_column_by_header(df, ['증감률', 'growth', 'rate'], search_rows=5)
    
    # fallback: 기존 하드코딩된 인덱스 사용 ← ⚠️ 문제
    if growth_rate_col is None:
        growth_rate_col = self.COL_GROWTH_RATE  # ← 이 값은 어디서 오는가?
        print(f"[광공업생산 분석시트] ⚠️ 증감률 컬럼 fallback: {growth_rate_col}")
```

**문제점**:
- `self.COL_GROWTH_RATE`가 `__init__()`에서 동적으로 설정되지만, 설정 로직이 복잡함
- 설정되지 않으면 `AttributeError` 발생 가능

### 개선안

```python
def _extract_nationwide_industries_from_analysis(self) -> dict:
    df = self.df_analysis
    header_row_idx = self.find_header_row(df, keywords=['지역', '산업', '증감률'])
    header_row = df.iloc[header_row_idx]
    
    # 동적으로 컬럼 찾기 (fallback 없이)
    growth_rate_col = self.find_column_by_keyword(
        header_row, 
        keywords=['증감률', 'growth', 'rate'],
        required=True  # 필수 컬럼 - 없으면 에러
    )
    
    # 명확한 에러 메시지
    if growth_rate_col is None:
        raise ValueError(
            f"증감률 컬럼을 찾을 수 없습니다. "
            f"헤더: {list(header_row)[:20]}"
        )
```

---

## ✅ 검증 방법

### 1. 25년 3분기 데이터 정확성 확인

**테스트 케이스**:
```python
# 광공업생산 - 전국 총지수
expected_growth_rate = "값 확인 필요"  # 엑셀에서 직접 확인

# 시스템에서 추출한 값
actual_growth_rate = generator.extract_nationwide_data()['growth_rate']

assert actual_growth_rate == expected_growth_rate, \
    f"불일치: {actual_growth_rate} != {expected_growth_rate}"
```

### 2. 25년 4분기 대비 검증 (가상)

**목적**: 분기 변경 시 동적 매핑 작동 확인

```python
# 25년 4분기로 변경 시 자동 적응 확인
generator_q4 = MiningManufacturingGenerator(
    excel_path, 
    year=2025, 
    quarter=4
)

# find_target_col_index가 자동으로 2025 4/4 컬럼 찾아야 함
data_q4 = generator_q4.extract_all_data()
```

---

## 🚀 즉시 적용 가능한 개선

### 1. DATA_START_ROW 동적 설정

**현재**: `DATA_START_ROW = 3` (고정)

**개선**:
```python
def find_data_start_row(self, df, header_row_idx):
    """헤더 다음 행을 데이터 시작 행으로 자동 설정"""
    return header_row_idx + 1

# 사용
header_row_idx = self.find_header_row(df)
data_start_row = self.find_data_start_row(df, header_row_idx)
data_df = df.iloc[data_start_row:].copy()
```

### 2. 컬럼 매핑 캐싱

**현재**: 매번 헤더 탐색

**개선**:
```python
def get_column_index(self, df, column_name, use_cache=True):
    """컬럼 인덱스 조회 (캐시 활용)"""
    if use_cache and column_name in self._col_cache:
        return self._col_cache[column_name]
    
    col_idx = self.find_column_by_keyword(
        df.iloc[self.header_row_idx], 
        keywords=self.COLUMN_KEYWORDS[column_name]
    )
    
    self._col_cache[column_name] = col_idx
    return col_idx
```

---

## 📌 권장 사항

### 즉시 실행
1. ✅ `find_target_col_index()` 활용도 확대
2. ✅ 터미널 로그에서 `SmartSearch` 작동 확인 → **이미 작동 중!**

### 단기 (1-2일)
3. ⏳ Priority 1 개선 (치명적 하드코딩 제거)
4. ⏳ Priority 2 개선 (시트명 fallback)

### 중기 (1주)
5. ⏳ Priority 3 개선 (헤더 행 자동 감지)
6. ⏳ 전체 generator 통합 테스트

### 장기 (지속적)
7. ⏳ Priority 4 개선 (리팩토링)
8. ⏳ 문서화 및 유지보수 가이드 작성
