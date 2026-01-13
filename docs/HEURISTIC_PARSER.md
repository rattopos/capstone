# Excel Heuristic Parser 문서

## 개요

`ExcelHeuristicParser`는 하드코딩된 시트 이름이나 행 인덱스 없이 Excel 파일을 동적으로 파싱하는 방어적이고 휴리스틱한 파서입니다.

## 주요 기능

### 1. Fuzzy Sheet Discovery (키워드 기반 시트 찾기)

하드코딩된 시트 이름 대신 키워드 리스트를 사용하여 시트를 찾습니다.

**예시:**
```python
parser = ExcelHeuristicParser(excel_path)

# 키워드 기반 시트 찾기
result = parser.find_target_sheet(
    keywords=['광공업생산', '집계'],
    required_columns=['지역', '분류', '산업'],
    required_row_labels=['전국', 'BCD']
)

if result:
    sheet_name, df = result
    print(f"시트 발견: {sheet_name}")
```

**특징:**
- 키워드 점수 기반 매칭 (정확한 매칭 > 부분 매칭 > 단어 매칭)
- 여러 시트가 매칭되면 컬럼/행 레이블로 검증
- 분할된 데이터셋 지원 (`find_multiple_sheets()`)

### 2. Dynamic Header Anchoring (동적 헤더 행 찾기)

하드코딩된 `header=2` 대신 앵커 키워드로 헤더 행을 동적으로 찾습니다.

**예시:**
```python
header_row = parser.locate_table_start(
    df,
    anchor_keywords=['지역', '분류', '산업', '2024', '2025']
)

if header_row is not None:
    print(f"헤더 행: {header_row}")
    # 헤더 행 이후부터 데이터 읽기
```

**특징:**
- 앵커 키워드 매칭 점수 기반
- 최소 점수 이상이면 헤더 행으로 인식
- 타이틀 행이 추가되어도 자동 대응

### 3. Content-Based Validation (내용 기반 검증)

시트가 실제로 필요한 데이터를 포함하는지 검증합니다.

**예시:**
```python
result = parser.find_target_sheet(
    keywords=['서비스업생산'],
    required_columns=['지역', '분류', '산업'],  # 70% 이상 매칭 필요
    required_row_labels=['전국', 'E~S', '총지수']  # 70% 이상 매칭 필요
)
```

**특징:**
- 컬럼 이름 검증 (헤더 행에서 찾기)
- 행 레이블 검증 (데이터 행에서 찾기)
- 70% 이상 매칭 시 통과 (유연한 검증)

## 사용 예시

### 기본 사용법

```python
from src.utils.excel_heuristic_parser import ExcelHeuristicParser

parser = ExcelHeuristicParser('data.xlsx')

try:
    # 시트 찾기
    result = parser.get_sheet_by_fallback(
        primary_keywords=['광공업생산', '집계', 'A'],
        fallback_keywords=['광공업생산', '광공업생산지수'],
        required_columns=['지역', '분류', '산업'],
        required_row_labels=['전국', 'BCD']
    )
    
    if result:
        sheet_name, df, is_fallback = result
        print(f"시트: {sheet_name}, Fallback: {is_fallback}")
        
        # 헤더 행 찾기
        header_row = parser.locate_table_start(
            df,
            anchor_keywords=['지역', '분류', '산업', '2024', '2025']
        )
        
        if header_row:
            print(f"헤더 행: {header_row}")
            # 데이터 처리...

finally:
    parser.close()
```

### Generator에서 사용

```python
def load_data(excel_path):
    parser = ExcelHeuristicParser(excel_path)
    
    try:
        # 분석 시트 찾기
        analysis_result = parser.get_sheet_by_fallback(
            primary_keywords=['서비스업생산', '분석', 'B'],
            fallback_keywords=['서비스업생산', '서비스업생산지수'],
            required_columns=['지역', '분류', '산업'],
            required_row_labels=['전국', 'E~S']
        )
        
        if analysis_result:
            sheet_name, df_analysis, use_raw = analysis_result
            # 데이터 처리...
    
    finally:
        parser.close()
```

## 점수 계산 로직

### 시트 이름 점수

1. **정확한 매칭**: 10.0 × (키워드 우선순위)
2. **시작 부분 매칭**: 5.0 × (키워드 우선순위)
3. **부분 매칭**: 2.0 × (키워드 우선순위)
4. **단어 매칭**: (매칭 단어 수 / 전체 단어 수) × 1.0 × (키워드 우선순위)

### 헤더 행 점수

- **정확한 컬럼 이름 매칭**: 10점
- **부분 매칭**: 5점
- 최소 점수: `len(anchor_keywords) × 3`

## 적용된 Generator

- ✅ `mining_manufacturing_generator.py`
- ✅ `service_industry_generator.py`
- ✅ `export_generator.py`
- ✅ `import_generator.py`

## 장점

1. **유연성**: 시트 이름이 변경되어도 자동 대응
2. **견고성**: 타이틀 행 추가, 시트 분할 등에 자동 대응
3. **검증**: 실제 데이터 내용으로 시트 검증
4. **재사용성**: 모든 generator에서 공통 유틸리티 사용

## 제약사항

- 시트 이름에 키워드가 포함되어야 함
- 필요한 컬럼/행 레이블이 70% 이상 있어야 함
- 헤더 행은 앵커 키워드로 찾을 수 있어야 함
