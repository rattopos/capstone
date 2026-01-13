# 데이터 추출 파이프라인 디버깅 로그

## 날짜: 2025-01-XX

## 문제 발견

### 1. 컬럼 인덱스 불일치 (Schema Mismatch)

**문제:**
- Generator가 잘못된 컬럼 인덱스를 사용하여 데이터를 찾지 못함
- 실제 Excel 파일 구조와 generator 코드의 컬럼 인덱스가 일치하지 않음

**실제 Excel 구조 (기초자료 수집표):**
```
행 2 (헤더):
  열 0: 지역 코드
  열 1: 지역 이름
  열 2: 분류 단계
  열 3: 가중치
  열 4: 산업 코드
  열 5: 산업 이름
  열 6~17: 연도별 데이터 (2013~2024)
  열 19~: 분기별 데이터 (2014 1/4 ~ 2025 3/4p)

분기 데이터:
  - 2024 2/4: 열 60
  - 2025 2/4: 열 64
  - 2025 3/4p: 열 65
```

**Generator가 사용하는 잘못된 인덱스:**
- 지역 이름: 열 3 또는 4 (실제: 열 1)
- 산업 코드: 열 6 또는 7 (실제: 열 4)
- 산업 이름: 열 7 또는 8 (실제: 열 5)
- 분류 단계: 열 4 또는 5 (실제: 열 2)
- 가중치: 열 6 (실제: 열 3)
- 2024.2/4: 열 22 (실제: 열 60)
- 2025.2/4: 열 26 (실제: 열 64)

### 2. 코드 버그

**발견된 버그:**
- `self.self.safe_float` → `self.safe_float`로 수정 필요 (7곳)

## 해결 방법

### 1. 동적 컬럼 인덱스 찾기

기초자료 형식의 Excel 파일에서 올바른 컬럼을 동적으로 찾도록 수정:

1. 헤더 행에서 컬럼 이름으로 인덱스 찾기
2. 분기 데이터는 "YYYY Q/4" 형식으로 찾기
3. 지역/산업 코드는 실제 값으로 찾기

### 2. 분석표 vs 기초자료 구분

- 분석표: 다른 컬럼 구조 사용
- 기초자료: 현재 발견된 구조 사용
- 두 형식을 모두 지원하도록 수정

## 수정 계획

1. `_extract_nationwide_from_aggregation()` 메서드 수정
   - 컬럼 인덱스를 동적으로 찾기
   - 분기 데이터 컬럼을 헤더에서 찾기

2. `_extract_regional_from_aggregation()` 메서드 수정
   - 동일한 방식으로 컬럼 인덱스 수정

3. 모든 `self.self.` 버그 수정

4. 테스트 및 검증

## 수정 완료 ✅

### 1. 버그 수정
- ✅ `self.self.safe_float` → `self.safe_float` (7곳 수정)

### 2. 컬럼 인덱스 동적 감지 구현
- ✅ 기초자료 형식 자동 감지 (헤더 행에서 "지역" 확인)
- ✅ 분기 데이터 컬럼을 헤더에서 동적으로 찾기
- ✅ 2024 2/4, 2025 2/4, 2025 3/4p 컬럼 자동 매핑
- ✅ 분석표 형식과 기초자료 형식 모두 지원

### 3. 테스트 결과
- ✅ 전국 데이터 추출 성공
  - production_index: 115.2
  - growth_rate: 2.9%
  - 증가 업종: 5개
  - 감소 업종: 5개
- ✅ 지역별 데이터 추출 성공
  - 증가 지역: 8개
  - 감소 지역: 9개
  - 전체 지역: 17개 (all 필드 추가 완료)

### 4. 구현 세부사항

**동적 컬럼 감지 로직:**
```python
# 기초자료 형식 감지
is_raw_format = False
if len(df.columns) > 1 and pd.notna(df.iloc[2, 1]) and '지역' in str(df.iloc[2, 1]):
    is_raw_format = True
    region_col = 1
    class_col = 2
    weight_col = 3
    code_col = 4
    name_col = 5
else:
    # 분석표 형식
    region_col = 4
    class_col = 5
    weight_col = 6
    code_col = 7
    name_col = 8

# 분기 컬럼 동적 찾기
for col_idx in range(len(header_row)):
    val = str(header_row[col_idx]) if pd.notna(header_row[col_idx]) else ''
    val_clean = val.strip().replace('.', ' ').replace('p', '').replace('P', '')
    
    if '2024' in val_clean and '2/4' in val_clean:
        col_2024_2q = col_idx
    if '2025' in val_clean and '2/4' in val_clean:
        col_2025_2q = col_idx
    if '2025' in val_clean and '3/4' in val_clean:
        col_2025_3q = col_idx
```

### 5. 남은 작업
- ⚠️ `price_trend_generator.py`의 오류 확인 필요 (다른 generator들은 정상 작동)
- ⚠️ 다른 generator들도 동일한 방식으로 수정 필요 (선택 사항)

## 추가 작업 완료 ✅

### 1. 공통 유틸리티 함수 생성
- ✅ `utils/column_detector.py` 생성
- ✅ `detect_column_structure()`: 컬럼 구조 자동 감지
- ✅ `find_quarter_columns()`: 분기 컬럼 동적 찾기
- ✅ `get_column_mapping()`: 전체 매핑 반환

### 2. Generator 수정 완료
- ✅ `service_industry_generator.py`: 동적 컬럼 감지 적용 완료
- ✅ `export_generator.py`: 정상 작동 확인
- ✅ `import_generator.py`: 정상 작동 확인

### 3. 통합 테스트 결과
- ✅ 광공업생산: 성공 (전국 데이터 2.9%, 지역 17개)
- ✅ 서비스업생산: 성공
- ✅ 수출: 성공
- ✅ 수입: 성공
- **성공률: 4/4 (100%)**

### 4. 개선 사항
- ✅ 동적 컬럼 감지로 기초자료/분석표 형식 모두 지원
- ✅ 분기 데이터 컬럼 자동 매핑
- ✅ `regional_data['all']` 필드 추가 (광공업생산)

## 휴리스틱 파서 구현 완료 ✅

### 1. ExcelHeuristicParser 클래스 생성
- ✅ `src/utils/excel_heuristic_parser.py` 생성
- ✅ Fuzzy Sheet Discovery: 키워드 기반 시트 찾기
- ✅ Dynamic Header Anchoring: 동적 헤더 행 찾기
- ✅ Content-Based Validation: 내용 기반 검증

### 2. Generator 적용
- ✅ `mining_manufacturing_generator.py`: 휴리스틱 파서 적용
- ✅ `service_industry_generator.py`: 휴리스틱 파서 적용
- ✅ `export_generator.py`: 휴리스틱 파서 적용
- ✅ `import_generator.py`: 휴리스틱 파서 적용

### 3. 테스트 결과
- ✅ ExcelHeuristicParser import 성공
- ✅ 광공업생산 Generator 정상 작동 (전국 데이터 2.9%)
- ✅ 시트 발견 테스트 성공
- ✅ 헤더 행 찾기 테스트 성공

### 4. 주요 개선사항
- ✅ 하드코딩된 시트 이름 제거
- ✅ 하드코딩된 헤더 행 인덱스 제거
- ✅ 키워드 기반 휴리스틱 매칭
- ✅ 내용 기반 시트 검증
