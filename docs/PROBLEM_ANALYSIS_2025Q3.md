# 🔍 문제 분석: 2025년 3분기 데이터 처리 오류

**분석일자**: 2026-01-11  
**대상 파일**: 
- 엑셀: `기초자료 수집표_2025년 3분기(캡스톤).xlsx`
- HTML: `~/Downloads/지역경제동향_2025년_3분기 (5).html`

---

## 📊 문제 현황

### HTML 파일 분석 결과
- **N/A 또는 없음**: 275개 발견
- **0.0%**: 49개 발견  
- **공란 [ ]**: 332개 발견
- **공란 포함 행**: 75개

### 엑셀 파일 분석 결과
- ✅ **2025년 3분기 데이터 정상 존재**
  - 광공업생산: col 65 = '2025  3/4p', 전국 데이터: 115.2
  - 서비스업생산: col 65 = '2025  3/4p', 전국 데이터: 119.2
  - 소비(소매, 추가): col 64 = '2025  3/4p', 전국 데이터: 102.8
  - 수출: col 69 = '2025  3/4p', 전국 데이터: 184966.939414

### HTML에 표시된 분기 정보
- **2025년 3/4**: 50개 발견 ✅
- **2025년 2/4**: 36개 발견 ⚠️ (잘못된 분기)
- **2024년 3/4**: 58개 발견 ✅

---

## 🔴 핵심 문제

### 문제 1: 하드코딩된 분기 값

**위치**: `data_converter.py` 라인 1230

```python
# 최신 분기 컬럼 찾기 (2025 2/4p)  ← 하드코딩!
current_quarter_col = -1
prev_year_quarter_col = -1

for col_idx in range(len(df.columns)):
    cell = str(df.iloc[header_row, col_idx]).strip()
    if '2025' in cell and '2/4' in cell:  # ← 2/4 하드코딩!
        current_quarter_col = col_idx
    elif '2024' in cell and '2/4' in cell:  # ← 2/4 하드코딩!
        prev_year_quarter_col = col_idx
```

**문제점**:
- 2025년 3분기 엑셀 파일을 처리할 때도 **2025년 2/4**를 찾으려고 시도
- 실제 엑셀에는 **2025년 3/4** 데이터만 존재
- 결과: `current_quarter_col = -1` → 잘못된 컬럼 또는 데이터 누락

### 문제 2: 연도/분기 동적 처리 부재

**영향받는 함수들**:
1. `data_converter.py::extract_grdp_data()` - GRDP 데이터 추출
2. 기타 하드코딩된 분기 값을 사용하는 함수들

**현재 상태**:
- 엑셀 파일명에서 연도/분기를 추출하는 로직은 있음 (`utils/excel_utils.py`)
- 하지만 `data_converter.py`는 이를 활용하지 않고 하드코딩 사용

### 문제 3: Smart Scout 시스템 미적용

**새로 구현한 앵커 기반 컬럼 찾기가 적용되지 않은 부분**:
- `data_converter.py`는 여전히 하드코딩된 패턴 사용
- `extractors/base.py`의 `find_column_by_anchor()` 활용 안 함

---

## 💡 해결 방안

### 즉시 수정 필요

#### 1. `data_converter.py::extract_grdp_data()` 수정

**현재 코드**:
```python
if '2025' in cell and '2/4' in cell:  # 하드코딩
```

**수정 방안 A: 동적 분기 찾기**
```python
# 엑셀 파일명에서 연도/분기 추출
from utils.excel_utils import extract_year_quarter_from_raw
year, quarter = extract_year_quarter_from_raw(self.raw_excel_path)

# 동적으로 분기 패턴 찾기
target_pattern = f"{year}.{quarter}/4"
for col_idx in range(len(df.columns)):
    cell = str(df.iloc[header_row, col_idx]).strip()
    if target_pattern in cell or f"{year} {quarter}/4" in cell:
        current_quarter_col = col_idx
        break
```

**수정 방안 B: Smart Scout 활용**
```python
from extractors.base import BaseExtractor

extractor = BaseExtractor(self.raw_excel_path, year, quarter)
current_quarter_col = extractor.find_column_by_anchor(
    '분기 GRDP', year, quarter
)
```

#### 2. 다른 하드코딩된 분기 값 찾기

**검색 필요**:
```bash
grep -r "2/4\|2025.*2" --include="*.py" | grep -v "2025_2Q"
```

---

## 🔍 추가 확인 사항

### 1. HTML 생성 시 연도/분기 전달 확인

**확인 위치**:
- `routes/preview.py` - HTML 생성 엔드포인트
- `report_generator.py` - 보고서 생성 로직

**확인 사항**:
- 엑셀 파일명에서 추출한 연도/분기가 올바르게 전달되는가?
- 모든 데이터 추출 함수에 연도/분기가 전달되는가?

### 2. Fallback 메커니즘 확인

**확인 사항**:
- 앵커 기반 탐색 실패 시 기존 동적 계산이 제대로 작동하는가?
- 2025년 3분기 컬럼을 찾지 못했을 때 적절한 오류 처리가 있는가?

---

## 📝 수정 우선순위

### 🔴 긴급 (데이터 누락 발생)
1. `data_converter.py::extract_grdp_data()` - 하드코딩 제거
2. 다른 하드코딩된 분기 값 찾아서 수정

### 🟡 중요 (데이터 정확성)
3. HTML 생성 시 연도/분기 전달 확인
4. Smart Scout 시스템을 모든 데이터 추출 함수에 적용

### 🟢 개선 (장기)
5. 모든 하드코딩된 분기 값을 동적 처리로 전환
6. 통합 테스트: 다양한 연도/분기 조합 테스트

---

## 🧪 테스트 시나리오

수정 후 다음 시나리오로 테스트 필요:

1. **2025년 3분기 엑셀 파일 처리**
   - ✅ 2025년 3/4 데이터 정상 추출
   - ✅ N/A, 공란 최소화

2. **2025년 2분기 엑셀 파일 처리**
   - ✅ 2025년 2/4 데이터 정상 추출
   - ✅ 기존 기능 유지

3. **2024년 4분기 엑셀 파일 처리**
   - ✅ 2024년 4/4 데이터 정상 추출
   - ✅ 동적 처리 검증

---

**분석 완료일**: 2026-01-11
