# Unified Generator 비교 문서

## ⚠️ 중요 업데이트 (2025-01-17)

**파일 스와핑 완료**: `unified_generator_v2.py`가 메인 엔진으로 승격되었습니다.

- ✅ `unified_generator.py` → `unified_generator_legacy.py.bak` (백업)
- ✅ `unified_generator_v2.py` → `unified_generator.py` (승격)
- ✅ 현재 `unified_generator.py`는 **V2 (간소화 버전)**입니다

## 개요

프로젝트에는 두 가지 버전의 `UnifiedReportGenerator`가 있었습니다:
- `templates/unified_generator_legacy.py.bak` (524줄) - **완전 버전** (백업됨)
- `templates/unified_generator.py` (345줄) - **간소화 버전** (현재 활성)

## 파일 위치

```
templates/
├── unified_generator.py              # ✅ 현재 활성 (V2 간소화 버전)
├── unified_generator_legacy.py.bak   # 백업된 완전 버전
└── unified_generator_old.py          # 구버전 (레거시)
```

## 주요 차이점

### 1. `unified_generator.py` (완전 버전)

**특징:**
- 템플릿 호환 구조를 완전히 지원
- 모든 필수 필드와 메타데이터 생성
- 상세한 데이터 변환 및 필드명 매핑

**주요 메서드:**
- `_generate_summary_table()`: 요약 테이블 생성
- `extract_all_data()`: 완전한 데이터 구조 반환

**반환 데이터 구조:**
```python
{
    'report_info': {...},
    'footer_info': {...},           # ✅ 있음
    'summary_box': {...},
    'nationwide_data': {...},
    'regional_data': {...},          # 필드명 변환됨
    'table_data': [...],             # 필드명 변환됨
    'summary_table': {...},          # ✅ 있음
    'top3_increase_regions': [...],  # ✅ 있음
    'top3_decrease_regions': [...],  # ✅ 있음
    **extra_fields                    # 템플릿별 추가 필드
}
```

**추가 기능:**
- `footer_info` 생성 (자료 출처 정보)
- `summary_table` 생성 (권역별 그룹핑된 요약 테이블)
- `top3_increase_regions`, `top3_decrease_regions` 생성
- 템플릿별 추가 필드 처리:
  - 소비동향: `increase_businesses_text`, `decrease_businesses_text`
  - 서비스업: `increase_industries_text`, `decrease_industries_text`
- 지역 데이터 필드명 변환 (`region_name` → `region` 등)

### 2. `unified_generator_v2.py` (간소화 버전)

**특징:**
- 최소한의 데이터만 추출
- 기본적인 구조만 제공
- 더 가볍고 빠름

**주요 메서드:**
- `extract_all_data()`: 기본 데이터 구조만 반환

**반환 데이터 구조:**
```python
{
    'report_info': {...},
    # 'footer_info': 없음 ❌
    'summary_box': {...},
    'nationwide_data': {...},
    'regional_data': {...},          # 원본 필드명 유지
    'table_data': [...]              # 원본 필드명 유지
    # 'summary_table': 없음 ❌
    # 'top3_increase_regions': 없음 ❌
    # 'top3_decrease_regions': 없음 ❌
}
```

**제외된 기능:**
- `footer_info` 생성 없음
- `summary_table` 생성 없음
- `top3_increase_regions`, `top3_decrease_regions` 없음
- 템플릿별 추가 필드 없음
- 필드명 변환 없음 (원본 필드명 유지)

## 공통 기능

두 버전 모두 공통으로 제공하는 기능:

1. **스마트 헤더 탐색기**
   - `find_target_col_index()`: 병합된 셀 처리
   - DataFrame 기반 헤더 탐색
   - `ffill`을 사용한 병합 셀 채우기

2. **데이터 추출**
   - `_extract_table_data_ssot()`: 집계 시트에서 데이터 추출
   - `extract_nationwide_data()`: 전국 데이터 추출
   - `extract_regional_data()`: 시도별 데이터 추출

3. **기본 구조**
   - `load_data()`: 집계 시트 로드
   - `_find_data_columns()`: 동적 컬럼 탐색
   - `BaseGenerator` 상속

## 사용 가이드

### 언제 `unified_generator.py`를 사용해야 하나?

✅ **완전 버전을 사용해야 하는 경우:**
- 템플릿에서 `footer_info`, `summary_table`이 필요한 경우
- `top3_increase_regions`, `top3_decrease_regions`가 필요한 경우
- 템플릿별 추가 필드가 필요한 경우
- 필드명 변환이 필요한 경우
- **현재 프로덕션에서 사용 중인 버전**

### 언제 `unified_generator_v2.py`를 사용해야 하나?

✅ **간소화 버전을 사용해야 하는 경우:**
- 최소한의 데이터만 필요한 경우
- 빠른 프로토타이핑
- 간단한 데이터 추출만 필요한 경우
- 메모리 사용량을 줄이고 싶은 경우

## 현재 사용 현황

### `MiningManufacturingGenerator` 클래스

두 파일 모두 동일한 래퍼 클래스를 제공:

```python
# unified_generator.py
class MiningManufacturingGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('mining', excel_path, year, quarter, excel_file)

# unified_generator_v2.py
class MiningManufacturingGenerator(UnifiedReportGenerator):
    def __init__(self, excel_path: str, year=None, quarter=None, excel_file=None):
        super().__init__('mining', excel_path, year, quarter, excel_file)
```

**주의:** 두 파일이 같은 클래스 이름을 사용하므로, import 시 어떤 파일을 사용하는지 명확히 해야 합니다.

## 최근 업데이트 (2025년)

### 스마트 헤더 탐색기 적용

두 버전 모두 최근에 다음 기능이 추가되었습니다:

1. **병합된 셀 처리**
   - `find_target_col_index()` 메서드가 DataFrame을 받아서 처리
   - `ffill(axis=1)`, `ffill(axis=0)`로 병합 셀 채우기
   - 연도 + 분기 + "증감률"/"등락" 조건으로 정확한 열 찾기

2. **None 처리 강화**
   - 인덱스를 찾지 못해도 중단하지 않고 계속 진행
   - `extract_all_data()` 내부에서 명시적으로 인덱스 확보
   - 모든 데이터 추출 단계에서 None 체크

## 권장 사항

1. **프로덕션 환경**: `unified_generator.py` 사용 (완전 버전)
2. **개발/테스트**: 필요에 따라 선택
3. **향후 계획**: 두 버전을 통합하거나 명확히 분리하는 것을 고려

## 참고

- 두 파일 모두 `BaseGenerator`를 상속받음
- 두 파일 모두 `config/report_configs.py`의 설정을 사용
- 두 파일 모두 동일한 데이터 추출 로직 사용 (차이는 반환 구조)
