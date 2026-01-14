# 엑셀 처리 최적화 요약

## 개요

`preprocess_excel` 함수와 데이터 추출 로직을 최적화하여 성능을 향상시키고 오류를 줄였습니다.

## 주요 변경사항

### 1. preprocess_excel 함수 최적화

#### 변경 전
- openpyxl 직접 계산 → formulas → xlwings 순서로 시도
- xlwings가 기본 로직에 포함되어 느림

#### 변경 후
- **openpyxl data_only=True 우선 사용** (가장 빠름)
- openpyxl 직접 계산 (시트 간 참조 매핑)
- formulas 선택적 사용 (필요시에만)
- **xlwings는 기본 로직에서 제외** (명시적 요청 시에만)

#### 처리 우선순위
1. `_try_openpyxl_data_only()`: data_only=True로 계산된 값 직접 읽기
2. `_try_openpyxl_calculation()`: 시트 간 참조 수식 직접 계산
3. `_try_formulas()`: 복잡한 수식 지원 (선택적)
4. `_try_xlwings()`: Excel 앱 사용 (명시적 요청 시에만, `use_xlwings=True`)

### 2. 엑셀 파일 캐싱 구조

#### ExcelCache 클래스 (`services/excel_cache.py`)
- Thread-safe 캐싱 구현
- 파일 수정 시간 기반 캐시 무효화
- pandas ExcelFile 및 openpyxl Workbook 캐싱 지원

#### 주요 기능
- `get_excel_file()`: pandas ExcelFile 캐싱
- `get_openpyxl_workbook()`: openpyxl Workbook 캐싱
- `clear_cache()`: 캐시 정리

### 3. ReportGenerator 레벨 캐싱

#### 변경 전
- 각 Generator가 엑셀 파일을 개별적으로 로드
- 전체 보도자료 생성 시 동일 파일을 여러 번 읽음

#### 변경 후
- `generate_report_html()`에서 엑셀 파일을 한 번만 로드
- 캐시된 ExcelFile 객체를 모든 Generator에 전달
- `generate_all_reports()`에서 캐싱 활용

### 4. BaseGenerator 통합

- `BaseGenerator`가 캐시된 ExcelFile을 받을 수 있도록 수정
- `__init__`에 `excel_file` 파라미터 추가
- 캐시된 객체가 있으면 재사용, 없으면 캐시에서 가져오기

## 성능 개선 효과

1. **엑셀 파일 로드 시간 단축**: 동일 파일을 여러 번 읽지 않음
2. **메모리 사용량 감소**: 캐시된 객체 재사용
3. **처리 속도 향상**: data_only=True 우선 사용으로 불필요한 계산 제거
4. **오류 감소**: 과도한 fallback 로직 제거

## 사용 방법

### 기본 사용 (자동 최적화)
```python
# 자동으로 최적화된 방법 사용
processed_path, success, message = preprocess_excel(excel_path)
```

### xlwings 명시적 사용 (필요시에만)
```python
# Excel 앱이 필요한 경우에만 명시적으로 요청
processed_path, success, message = preprocess_excel(
    excel_path, 
    use_xlwings=True  # Excel 앱 사용
)
```

### 캐싱 활용
```python
from services.excel_cache import get_excel_file

# 엑셀 파일을 한 번만 로드하고 재사용
excel_file = get_excel_file(excel_path, use_data_only=True)

# 여러 Generator에 전달
for generator in generators:
    generator.load_data(excel_file=excel_file)
```

## 주의사항

1. **xlwings 사용**: 기본적으로 사용하지 않으며, 명시적으로 요청할 때만 실행됩니다.
2. **캐시 관리**: 작업 완료 후 `clear_excel_cache()`로 메모리 정리 권장
3. **파일 수정**: 파일이 수정되면 캐시가 자동으로 무효화됩니다.

## 관련 파일

- `services/excel_processor.py`: 전처리 로직 최적화
- `services/excel_cache.py`: 캐싱 구조 (신규)
- `services/report_generator.py`: 캐싱 통합
- `templates/base_generator.py`: 캐시된 ExcelFile 지원
- `routes/api.py`: 전체 보도자료 생성 시 캐싱 활용
