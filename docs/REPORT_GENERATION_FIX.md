# 광공업생산 보고서 생성 오류 분석 및 해결

## 📋 문제 요약

**증상**: 분석표 파일을 업로드했을 때 광공업생산 보고서가 생성되지 않고 다음 오류 발생:
```
[ERROR] 보고서 생성 오류: single positional indexer is out-of-bounds
```

**영향 범위**: 광공업생산 보고서 (`mining_manufacturing_generator.py`)

---

## 🔍 원인 분석

### 1. 엑셀 수식 미계산 문제

분석표 엑셀 파일의 구조:
- **집계 시트** (`A(광공업생산)집계`): 원본 데이터 (값으로 저장됨) ✅
- **분석 시트** (`A 분석`): 집계 시트를 참조하는 **수식**으로 구성 ❌

```
┌─────────────────────────────────────────────────────────────┐
│  분석표.xlsx                                                 │
├─────────────────────────────────────────────────────────────┤
│  A(광공업생산)집계  →  실제 데이터 값 저장                    │
│  A 분석            →  =집계시트!A1 형태의 수식                │
│                       (Python에서 읽으면 수식 미계산 → 빈 값)  │
└─────────────────────────────────────────────────────────────┘
```

### 2. Python에서 엑셀 수식 처리

```python
# openpyxl/pandas는 기본적으로 수식을 계산하지 않음
df = pd.read_excel('분석표.xlsx', sheet_name='A 분석')

# 결과: 수식이 계산되지 않아 대부분 NaN 또는 빈 값
```

### 3. 기존 코드의 문제점

```python
# mining_manufacturing_generator.py (수정 전)
def extract_nationwide_data(self):
    df = self.df_analysis  # 분석 시트 사용
    
    # 분석 시트에서 '전국' + 'BCD' 조건으로 검색
    nationwide_total = df[(df[3] == '전국') & (df[6] == 'BCD')].iloc[0]
    #                                                           ^^^^
    # 문제: 분석 시트가 비어있어서 조건에 맞는 행이 없음 → IndexError 발생
```

---

## ✅ 해결 방법

### 1. 분석 시트 데이터 유무 확인 로직 추가

```python
def load_data(self):
    # 분석 시트 로드 후 실제 데이터가 있는지 확인
    if analysis_sheet:
        self.df_analysis = pd.read_excel(xl, sheet_name=analysis_sheet, header=None)
        
        # 분석 시트에 실제 데이터가 있는지 확인 (수식 미계산 체크)
        test_row = self.df_analysis[(self.df_analysis[3] == '전국') | (self.df_analysis[4] == '전국')]
        if test_row.empty or test_row.iloc[0].isna().sum() > 20:
            print(f"[광공업생산] 분석 시트가 비어있음 → 집계 시트에서 직접 계산")
            self.use_aggregation_only = True
```

### 2. 집계 시트에서 직접 증감률 계산

```python
def _extract_nationwide_from_aggregation(self):
    """집계 시트에서 전국 데이터 추출 (증감률 직접 계산)"""
    df = self.df_aggregation
    
    # 집계 시트 컬럼 구조:
    # 4=지역이름, 5=분류단계, 6=가중치, 7=산업코드, 8=산업이름
    # 22=2024.2/4(전년동분기), 26=2025.2/4p(당분기)
    
    # 전국 총지수 행 (BCD)
    nationwide_total = df[(df[4] == '전국') & (df[7] == 'BCD')].iloc[0]
    
    # 당분기와 전년동분기 지수로 증감률 계산
    current_index = nationwide_total[26]   # 2025.2/4p = 114.4
    prev_year_index = nationwide_total[22] # 2024.2/4  = 112.0
    
    growth_rate = ((current_index - prev_year_index) / prev_year_index) * 100
    # = ((114.4 - 112.0) / 112.0) * 100 = 2.1%
```

---

## 📊 수정 결과

### 수정 전
```
[ERROR] 보고서 생성 오류: single positional indexer is out-of-bounds
IndexError: single positional indexer is out-of-bounds
```

### 수정 후
```
[광공업생산] 분석 시트가 비어있음 → 집계 시트에서 직접 계산
[DEBUG] 추출된 데이터 키: ['report_info', 'nationwide_data', 'regional_data', ...]
[DEBUG] 보고서 생성 성공!

=== 추출된 데이터 ===
전국 증감률: 2.1%
전국 지수: 114.4
증가 지역 수: 6개
  - 충북: 14.1%
  - 경기: 12.3%
  - 광주: 11.3%
감소 지역:
  - 서울: -10.1%
  - 충남: -6.4%
  - 부산: -4.0%
```

---

## 📁 수정된 파일

| 파일 | 수정 내용 |
|------|----------|
| `templates/mining_manufacturing_generator.py` | 집계 시트 우선 사용, 증감률 직접 계산 로직 추가 |

---

## 🔧 기술적 세부사항

### 집계 시트 컬럼 매핑

| 컬럼 인덱스 | 내용 | 예시 값 |
|------------|------|---------|
| 4 | 지역이름 | 전국, 서울, 부산... |
| 5 | 분류단계 | 0(총지수), 1(대분류), 2(중분류) |
| 6 | 가중치 | 10000(전국), 623(식료품)... |
| 7 | 산업코드 | BCD(총지수), C(제조업)... |
| 8 | 산업이름 | 총지수, 제조업, 식료품 제조업... |
| 14 | 2022.2/4 | 지수 값 |
| 18 | 2023.2/4 | 지수 값 |
| 21 | 2024.1/4 | 지수 값 |
| 22 | 2024.2/4 | 지수 값 (전년동분기) |
| 25 | 2025.1/4 | 지수 값 |
| 26 | 2025.2/4p | 지수 값 (당분기) |

### 증감률 계산 공식

```
증감률(%) = (당분기지수 - 전년동분기지수) / 전년동분기지수 × 100
```

### 기여도 계산 공식

```
기여도 = (당분기지수 - 전년동분기지수) / 전국전년동분기지수 × (가중치/10000) × 100
```

---

## 💡 교훈 및 권장사항

### 1. 엑셀 수식 의존 문제
- **문제**: Python pandas/openpyxl은 엑셀 수식을 계산하지 않음
- **해결**: 원본 데이터가 있는 집계 시트에서 직접 계산

### 2. 방어적 프로그래밍
- **문제**: 데이터가 비어있을 때 `.iloc[0]`이 IndexError 발생
- **해결**: 데이터 유무 먼저 확인 후 처리

### 3. 폴백(Fallback) 전략
- **원칙**: 주 데이터 소스 실패 시 대체 소스 사용
- **적용**: 분석 시트 → 집계 시트 → 기초자료 시트

---

## 📅 수정 일자

- **2025년 12월 30일**
- **수정자**: AI Assistant

---

## ✔️ 테스트 결과

```bash
# 테스트 명령
python3 -c "
from templates.mining_manufacturing_generator import 광공업생산Generator
generator = 광공업생산Generator('uploads/분석표_25년_2분기_캡스톤.xlsx')
data = generator.extract_all_data()
print(f'전국 증감률: {data[\"nationwide_data\"][\"growth_rate\"]}%')
"

# 출력
[광공업생산] 분석 시트가 비어있음 → 집계 시트에서 직접 계산
전국 증감률: 2.1%
```

✅ **테스트 통과**

