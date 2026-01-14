# 가중치 하드코딩 위치 보고서

가중치 정보가 없어서 하드코딩된 곳들을 정리한 문서입니다.

## 1. 건설 시트 가중치 없음

### 위치: `data_converter.py`

```python
"F'(건설)집계": {
    'meta_start': 2, 
    'raw_meta_cols': 4, 
    'year_start': 6, 
    'quarter_start': 11, 
    'weight_col': None,  # ⚠️ 가중치 열이 None으로 설정됨
    'raw_weight_col': None  # ⚠️ 기초자료 가중치 열도 None
}
```

**영향**: 건설 시트는 가중치 정보가 없어서 기여율 계산 시 다른 방식을 사용해야 합니다.

---

## 2. 건설 시트 가중치 처리 로직

### 위치: `fill_analysis_blanks.py` (333-345줄)

```python
# 가중치 가져오기 (건설은 가중치가 없을 수 있으므로 다른 방식 사용)
# 건설 시트는 가중치 대신 다른 방식으로 기여도 계산할 수 있음
# 건설은 가중치 없이도 증감률을 기여도로 사용할 수 있음
```

**영향**: 건설 시트는 가중치 없이 증감률을 기여도로 사용하는 로직이 있습니다.

---

## 3. GRDP 기여율 기본값 하드코딩

### 위치: `templates/default_contributions.json`

이 파일에는 GRDP 산업별 기여율 기본값이 하드코딩되어 있습니다:

```json
{
  "national": {
    "growth_rate": 1.2,
    "contributions": {
      "manufacturing": 0.5,  // ⚠️ 하드코딩된 값
      "construction": 0.1,   // ⚠️ 하드코딩된 값
      "service": 0.5,        // ⚠️ 하드코딩된 값
      "other": 0.1           // ⚠️ 하드코딩된 값
    },
    "is_placeholder": true
  },
  "regional": {
    "서울": {
      "growth_rate": 1.0,
      "manufacturing": 0.2,  // ⚠️ 하드코딩된 값
      "construction": 0.1,   // ⚠️ 하드코딩된 값
      "service": 0.6,        // ⚠️ 하드코딩된 값
      "other": 0.1           // ⚠️ 하드코딩된 값
    },
    // ... 다른 지역들도 모두 하드코딩됨
  }
}
```

**사용 위치**: `services/grdp_service.py`의 `get_default_grdp_data()` 함수

**영향**: GRDP 데이터가 없을 때 이 기본값들이 사용됩니다. 파일 주석에 "추정치"라고 명시되어 있으며 실제 데이터로 업데이트가 필요합니다.

---

## 4. 기본값 0.0 하드코딩

### 위치: `services/grdp_service.py`

여러 곳에서 기본값 `0.0`이 하드코딩되어 있습니다:

```python
# 516-527줄: 지역별 기본값
regional_data.append({
    'region': region,
    'region_group': region_groups.get(region, ''),
    'growth_rate': 0.0,        # ⚠️ 하드코딩
    'manufacturing': 0.0,      # ⚠️ 하드코딩
    'construction': 0.0,       # ⚠️ 하드코딩
    'service': 0.0,            # ⚠️ 하드코딩
    'other': 0.0,              # ⚠️ 하드코딩
    'placeholder': True,
    'needs_review': True
})

# 530-531줄: 전국 기본값
national_growth = 0.0
national_contributions = {
    'manufacturing': 0.0,      # ⚠️ 하드코딩
    'construction': 0.0,      # ⚠️ 하드코딩
    'service': 0.0,           # ⚠️ 하드코딩
    'other': 0.0              # ⚠️ 하드코딩
}

# 545줄: 1위 지역 기본값
top_region = {
    'region': '-', 
    'growth_rate': 0.0,      # ⚠️ 하드코딩
    'manufacturing': 0.0,    # ⚠️ 하드코딩
    'construction': 0.0,     # ⚠️ 하드코딩
    'service': 0.0,          # ⚠️ 하드코딩
    'other': 0.0,            # ⚠️ 하드코딩
    'placeholder': True
}
```

**영향**: 데이터가 없을 때 모든 값이 0.0으로 설정됩니다.

---

## 5. 기타 기본값 하드코딩

### 위치: `services/summary_data.py`

여러 함수에서 기본값들이 하드코딩되어 있습니다:

- `_get_default_chart_data()`: `index: 100.0`, `change: 0.0`
- `_get_default_employment_data()`: `rate: 60.0`, `change: 0.0`
- `_get_default_trade_data()`: `amount: 0`, `change: 0.0`

**영향**: 데이터 추출 실패 시 기본값이 사용됩니다.

---

## 권장 사항

1. **건설 시트 가중치**: 건설 시트에 가중치 정보가 필요한지 확인하고, 필요하다면 기초자료 수집표에서 가중치를 추출하도록 수정

2. **GRDP 기본값**: `templates/default_contributions.json`의 값들은 추정치이므로, 실제 데이터가 있으면 업데이트 필요

3. **기본값 처리**: 데이터가 없을 때 0.0 대신 명시적으로 "데이터 없음" 또는 플레이스홀더 표시를 고려

4. **가중치 추출**: 기초자료 수집표에서 가중치 정보를 자동으로 추출하는 로직 추가 검토

---

## 관련 파일 목록

- `data_converter.py`: 시트 구조 정의 (건설 가중치 None)
- `fill_analysis_blanks.py`: 건설 가중치 처리 로직
- `templates/default_contributions.json`: GRDP 기여율 기본값
- `services/grdp_service.py`: 기본값 사용 로직
- `services/summary_data.py`: 기타 기본값 정의
