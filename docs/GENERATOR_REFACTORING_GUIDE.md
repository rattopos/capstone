# Generator 리팩토링 가이드

## 개요

모든 Generator 클래스들이 공통 기능을 `BaseGenerator`에서 상속받도록 리팩토링했습니다.

## BaseGenerator 제공 기능

### 1. 공통 유틸리티 메서드
- `safe_float(value, default=None)`: 안전한 float 변환
- `safe_round(value, decimals=1, default=None)`: 안전한 반올림
- `safe_int(value, default=None)`: 안전한 int 변환
- `safe_str(value, default='')`: 안전한 문자열 변환

### 2. 엑셀 파일 처리
- `load_excel()`: 엑셀 파일 로드 (캐싱)
- `get_sheet(sheet_name, use_cache=True)`: 시트 읽기 (캐싱)
- `find_sheet_with_fallback(primary_sheets, fallback_sheets)`: 시트 찾기 (대체 시트 지원)

### 3. 데이터 추출 헬퍼
- `get_cell_value(df, row, col, default=None)`: 셀 값 안전하게 추출
- `find_row_by_value(df, col, value, start_row=0)`: 특정 값으로 행 찾기
- `find_rows_by_condition(df, conditions, start_row=0)`: 여러 조건으로 행 찾기
- `check_sheet_has_data(df, test_conditions, max_empty_cells=20)`: 시트에 데이터 있는지 확인

### 4. 기본 정보
- `get_report_info()`: report_info 기본 구조 반환

## 리팩토링 패턴

### 1. 클래스 기반 Generator (예: mining_manufacturing_generator.py)

**Before:**
```python
def safe_float(value, default=None):
    # ... 중복 코드 ...

def find_sheet_with_fallback(xl, primary_sheets, fallback_sheets):
    # ... 중복 코드 ...

class 광공업생산Generator:
    def __init__(self, excel_path: str, year=None, quarter=None):
        self.excel_path = excel_path
        # ...
    
    def load_data(self):
        xl = pd.ExcelFile(self.excel_path)
        # ...
```

**After:**
```python
from .base_generator import BaseGenerator

class 광공업생산Generator(BaseGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None):
        super().__init__(excel_path, year, quarter)
        # ...
    
    def load_data(self):
        xl = self.load_excel()  # BaseGenerator 메서드 사용
        sheet, use_raw = self.find_sheet_with_fallback(...)  # BaseGenerator 메서드 사용
        # ...
    
    # safe_float, safe_round 등은 self.safe_float, self.safe_round로 변경
```

### 2. 함수 기반 Generator (예: construction_generator.py)

함수 기반 Generator는 두 가지 방법으로 리팩토링할 수 있습니다:

#### 방법 1: 클래스로 변환 (권장)
```python
from .base_generator import BaseGenerator

class ConstructionGenerator(BaseGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None):
        super().__init__(excel_path, year, quarter)
        self.df_analysis = None
        self.df_index = None
    
    def load_data(self):
        # 기존 load_data() 함수 내용을 여기로 이동
        xl = self.load_excel()
        # ...
    
    def extract_all_data(self):
        # 기존 generate_report_data() 함수 내용을 여기로 이동
        # ...
```

#### 방법 2: 함수에서 BaseGenerator 유틸리티 사용
```python
from .base_generator import BaseGenerator

def load_data(excel_path):
    # BaseGenerator 인스턴스 생성 (유틸리티만 사용)
    base = BaseGenerator(excel_path)
    xl = base.load_excel()
    sheet, use_raw = base.find_sheet_with_fallback(...)
    # ...
```

## 리팩토링 체크리스트

각 Generator 파일을 리팩토링할 때 다음 사항을 확인하세요:

- [ ] `BaseGenerator` import 추가
- [ ] 클래스가 `BaseGenerator`를 상속받도록 변경
- [ ] `__init__`에서 `super().__init__()` 호출
- [ ] 중복된 `safe_float`, `safe_round` 함수 제거
- [ ] 모든 `safe_float(...)` 호출을 `self.safe_float(...)`로 변경
- [ ] 모든 `safe_round(...)` 호출을 `self.safe_round(...)`로 변경
- [ ] `find_sheet_with_fallback` 호출을 `self.find_sheet_with_fallback`로 변경
- [ ] `pd.ExcelFile(...)` 호출을 `self.load_excel()`로 변경
- [ ] `pd.read_excel(...)` 호출을 `self.get_sheet(...)`로 변경 (가능한 경우)
- [ ] `extract_all_data()` 메서드 구현 (abstract method)

## 리팩토링 대상 파일

다음 파일들을 순서대로 리팩토링하세요:

1. ✅ `mining_manufacturing_generator.py` (완료)
2. `construction_generator.py` (함수 기반 → 클래스 변환 필요)
3. `service_industry_generator.py` (함수 기반 → 클래스 변환 필요)
4. `consumption_generator.py` (함수 기반 → 클래스 변환 필요)
5. `export_generator.py` (함수 기반 → 클래스 변환 필요)
6. `import_generator.py`
7. `employment_rate_generator.py`
8. `unemployment_generator.py`
9. `price_trend_generator.py`
10. `domestic_migration_generator.py`
11. `regional_generator.py`
12. `reference_grdp_generator.py`
13. `infographic_generator.py`
14. `statistics_table_generator.py`

## 주의사항

1. **하위 호환성**: 기존 함수 시그니처를 유지하거나, 호출하는 코드도 함께 수정해야 합니다.
2. **에러 처리**: BaseGenerator의 메서드들은 안전하게 에러를 처리하지만, 추가적인 검증이 필요할 수 있습니다.
3. **캐싱**: `get_sheet()`는 자동으로 캐싱하므로, 같은 시트를 여러 번 읽어도 한 번만 로드됩니다.
4. **테스트**: 리팩토링 후 반드시 실제 데이터로 테스트하세요.

## 예제: 완전한 리팩토링 예시

`mining_manufacturing_generator.py`를 참고하세요. 이 파일은 완전히 리팩토링되어 BaseGenerator를 상속받고 있습니다.
