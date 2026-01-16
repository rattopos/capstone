# ServiceIndustryGenerator 하드코딩 감사 보고서

**날짜**: 2026년 1월 16일  
**감사 대상**: `templates/service_industry_generator.py`  
**총 발견**: 28개 하드코딩된 컬럼/행 참조

---

## 🔴 심각도 높음 - 즉시 수정 필요

### 1. 메타데이터 컬럼 하드코딩 (영향: 전체)

**문제**: 지역, 분류단계, 산업코드, 산업명 컬럼이 하드코딩됨

| 컬럼 인덱스 | 의미 | 사용 위치 | 영향도 |
|------------|------|----------|--------|
| `row[3]` | 지역명 | 12개 위치 | ⚠️ 높음 |
| `row[4]` | 분류단계 | 4개 위치 | ⚠️ 높음 |
| `row[6]` | 산업코드 | 6개 위치 | ⚠️ 높음 |
| `row[7]` | 산업명 | 8개 위치 | ⚠️ 높음 |

**예시**:
```python
# ❌ 현재 (하드코딩)
nationwide_rows = df[(df[3] == '전국') & (df[6] == 'E~S')]
industry_name = row[7]
classification = str(row[4]).strip()

# ✅ 수정 필요
region_col = self._find_metadata_column('region')
code_col = self._find_metadata_column('industry_code')
name_col = self._find_metadata_column('industry_name')
class_col = self._find_metadata_column('classification')

nationwide_rows = df[(df[region_col] == '전국') & (df[code_col] == 'E~S')]
industry_name = row[name_col]
classification = str(row[class_col]).strip()
```

**위험성**:
- 엑셀 구조 변경 시 전체 시스템 실패
- 다른 포맷의 분석표 사용 불가
- 디버깅 어려움

---

### 2. 전국 데이터 행 하드코딩

**문제**: 전국 데이터가 항상 행 3에 있다고 가정

```python
# ❌ 현재 (하드코딩)
nationwide_idx = region_indices.get('전국', 3)  # 기본값 3
index_row = self.df_aggregation.iloc[3]  # 고정된 행 3

# ✅ 수정 필요
nationwide_idx = self._find_region_row('전국')
```

**위험성**:
- 엑셀 시트에 주석이나 헤더가 추가되면 실패
- 전국 데이터 위치가 바뀌면 잘못된 값 추출

---

## 🟡 심각도 중간 - 개선 권장

### 3. 업종 범위 하드코딩

```python
# ❌ 현재 (하드코딩)
for i in range(start_idx + 1, min(start_idx + 14, len(df))):  # 13개 업종 가정
    ...

# ✅ 수정 필요
# 다음 총지수 행까지 동적으로 찾기
next_total_idx = self._find_next_total_row(start_idx)
for i in range(start_idx + 1, next_total_idx):
    ...
```

---

## 🟢 심각도 낮음 - 선택적 개선

### 4. 기여도 컬럼 상대 위치

```python
# 🤔 현재 (상대 위치)
contribution_col = target_col + 6  # 증감률 + 6 = 기여도

# ⚠️ 위험: 엑셀 구조 변경 시 틀릴 수 있음
# ✅ 이상적: 헤더에서 "기여도" 키워드로 찾기
```

---

## 📊 비교: mining_manufacturing vs service_industry

| 기능 | mining_manufacturing | service_industry |
|------|---------------------|------------------|
| `_find_metadata_column()` | ✅ 구현됨 | ❌ 없음 |
| 메타데이터 컬럼 동적 탐색 | ✅ 사용 | ❌ 하드코딩 |
| `_col_cache` 체계 | ✅ 완전 | 🟡 부분적 |
| 전국 행 동적 탐색 | ✅ 검색 | ❌ 하드코딩 |

**결론**: service_industry는 mining_manufacturing보다 **동적 수준이 낮음**

---

## 🔧 수정 계획

### Phase A: 메타데이터 컬럼 동적화 (최우선)

**작업 시간**: 1-2시간  
**영향도**: 매우 높음

1. `_find_metadata_column()` 메서드 추가 (base_generator 또는 복사)
2. `_initialize_column_indices()`에서 메타데이터 컬럼 캐싱
3. 모든 `row[3]`, `row[4]`, `row[6]`, `row[7]` 교체

**예상 수정 위치**:
- `_get_region_indices()`: 3개
- `_get_nationwide_from_analysis()`: 2개
- `_get_nationwide_from_aggregation()`: 4개
- `_get_regional_from_analysis()`: 3개
- `_get_regional_from_aggregation()`: 6개
- `_get_table_data_from_analysis()`: 2개
- `_get_table_data_from_aggregation()`: 3개

**총**: ~23개 위치

### Phase B: 행 탐색 동적화

**작업 시간**: 30분  
**영향도**: 중간

1. `_find_region_row()` 메서드 추가
2. `iloc[3]` 하드코딩 제거

### Phase C: 업종 범위 동적화

**작업 시간**: 30분  
**영향도**: 낮음

1. `_find_next_total_row()` 메서드 추가
2. `range(start + 14)` 하드코딩 제거

---

## 🎯 권장 사항

### 즉시 조치:
1. ✅ **Phase A 먼저 진행** (메타데이터 컬럼 동적화)
   - 가장 영향이 크고 위험도가 높음
   - 다른 엑셀 구조에서 완전히 실패할 수 있음

2. 🔄 **mining_manufacturing 패턴 적용**
   - 이미 검증된 동적 시스템
   - 코드 재사용 가능

### 장기 조치:
3. 📋 **모든 generator 통일**
   - 동일한 동적 탐색 패턴 적용
   - base_generator에 공통 메서드 추가

---

## 📈 예상 효과

### Before (현재):
- 엑셀 컬럼 순서 변경 → ❌ 전체 실패
- 다른 포맷의 분석표 → ❌ 사용 불가
- 메타데이터 위치 변경 → ❌ 잘못된 값 추출

### After (수정 후):
- 엑셀 컬럼 순서 변경 → ✅ 자동 대응
- 다른 포맷의 분석표 → ✅ 자동 탐색
- 메타데이터 위치 변경 → ✅ 정확한 값 추출

---

## ⚠️ 주의사항

1. **테스트 필수**: 수정 후 다양한 분기로 검증
2. **Fallback 전략**: 찾지 못했을 때 기본값 제공
3. **로그 추가**: 어떤 컬럼을 찾았는지 명확히 출력

---

**작성자**: AI Assistant  
**검증 도구**: grep, manual code review  
**다음 단계**: Phase A 작업 시작 (사용자 승인 후)
