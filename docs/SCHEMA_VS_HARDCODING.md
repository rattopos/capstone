# 스키마 방식 vs 하드코딩 방식 분석

## 개요

본 프로젝트는 **스키마 기반**과 **하드코딩** 두 가지 방식이 혼재되어 있습니다. 초기에는 하드코딩으로 시작했으나, 2025년 4분기 보도자료 형식 변경을 계기로 **스키마 기반 아키텍처로 전환**했습니다.

---

## 📊 방식별 분류

### ✅ 스키마 방식 (유연함)

| 구분 | 파일 | 설명 | 수정 시 |
|------|------|------|---------|
| **데이터 구조 정의** | `*_schema.json` (21개) | 보도자료별 데이터 필드 정의 | JSON 파일만 수정 |
| **업종명 매핑** | 스키마 내 `industry_name_mapping` | 엑셀 업종명 → 보도자료 표기명 | JSON 수정 |
| **지역명 매핑** | 스키마 내 `region_mapping` | 지역명 변환 규칙 | JSON 수정 |
| **텍스트 생성 규칙** | 스키마 내 `rules` | 증가/감소 문장 패턴 | JSON 수정 |
| **컬럼 매핑 메타데이터** | 스키마 내 `excel_column_mapping` | 엑셀 컬럼 인덱스 정보 | JSON 수정 |

#### 스키마 파일 예시 (`mining_manufacturing_schema.json`)
```json
{
  "industry_name_mapping": {
    "반도체·전자부품": ["전자 부품, 컴퓨터, 영상, 음향 및 통신장비 제조업"],
    "의료·정밀": ["의료, 정밀, 광학 기기 및 시계 제조업"]
  },
  "excel_column_mapping": {
    "지역이름": 3,
    "산업이름": 7,
    "증감": 22,
    "기여도": 28
  },
  "rules": {
    "text_generation": {
      "increase_pattern": "전국 광공업생산({production_index})은 {industries} 등의 생산이 늘어..."
    }
  }
}
```

---

### ⚠️ 하드코딩 방식 (수정 시 Python 코드 변경 필요)

| 구분 | 위치 | 내용 | 문제점 |
|------|------|------|--------|
| **시트명** | Generator 클래스 | `'A 분석'`, `'A(광공업생산)집계'` | 시트명 변경 시 코드 수정 필요 |
| **컬럼 인덱스** | Generator 메서드 | `region_agg[18]`, `region_agg[22]` | 엑셀 구조 변경 시 전체 수정 |
| **업종명 매핑** | Generator 클래스 상수 | `INDUSTRY_NAME_MAP = {...}` | 스키마와 중복 |
| **지역 그룹** | Generator 클래스 상수 | `REGION_GROUPS = {...}` | 지역 추가 시 코드 수정 |
| **지역 순서** | Generator 메서드 | `individual_regions = ['서울', '부산', ...]` | 순서 변경 시 코드 수정 |

#### 하드코딩 예시 (`mining_manufacturing_generator.py`)
```python
class 광공업생산Generator:
    # ⚠️ 하드코딩: 업종명 매핑 (스키마와 중복)
    INDUSTRY_NAME_MAP = {
        "전자 부품, 컴퓨터, 영상, 음향 및 통신장비 제조업": "반도체·전자부품",
        "의료, 정밀, 광학 기기 및 시계 제조업": "의료·정밀",
        ...
    }
    
    # ⚠️ 하드코딩: 지역 그룹
    REGION_GROUPS = {
        "경인": {"regions": ["서 울", "인 천", "경 기"], "group": "경인"},
        ...
    }
    
    def extract_summary_table(self):
        # ⚠️ 하드코딩: 컬럼 인덱스
        idx_2023_2q = safe_float(region_agg[18], 100)  # 2023.2/4
        idx_2024_2q = safe_float(region_agg[22], 100)  # 2024.2/4
        idx_2025_1q = safe_float(region_agg[25], 100)  # 2025.1/4
        idx_2025_2q = safe_float(region_agg[26], 100)  # 2025.2/4p
```

---

## 📁 파일별 분류

### Generator 파일 (14개) - 혼재

| 파일 | 스키마 활용 | 하드코딩 정도 |
|------|-------------|---------------|
| `mining_manufacturing_generator.py` | 부분적 | 🔴 높음 (컬럼 인덱스, 시트명) |
| `service_industry_generator.py` | 부분적 | 🔴 높음 |
| `consumption_generator.py` | 부분적 | 🟡 중간 |
| `construction_generator.py` | 부분적 | 🟡 중간 |
| `export_generator.py` | 부분적 | 🟡 중간 |
| `import_generator.py` | 부분적 | 🟡 중간 |
| `price_trend_generator.py` | 부분적 | 🟡 중간 |
| `employment_rate_generator.py` | 부분적 | 🟡 중간 |
| `unemployment_generator.py` | 부분적 | 🟡 중간 |
| `domestic_migration_generator.py` | 부분적 | 🟡 중간 |
| `regional_generator.py` | 부분적 | 🔴 높음 (2595줄) |
| `statistics_table_generator.py` | 부분적 | 🟡 중간 |
| `infographic_generator.py` | 부분적 | 🟡 중간 |
| `reference_grdp_generator.py` | 부분적 | 🟡 중간 |

### 스키마 파일 (21개) - 완전 스키마 방식

모든 `*_schema.json` 파일은 스키마 방식입니다.

### 설정 파일 - 혼재

| 파일 | 방식 |
|------|------|
| `config/reports.py` | 🔴 하드코딩 (보도자료 목록, 페이지 설정) |
| `config/settings.py` | 🟡 설정 파일 (환경변수) |

---

## 🔍 구체적인 하드코딩 위치

### 1. 엑셀 컬럼 인덱스 하드코딩

**위치**: 모든 `*_generator.py` 파일

```python
# mining_manufacturing_generator.py:777-780
idx_2023_2q = safe_float(region_agg[18], 100)  # ⚠️ 18번 컬럼
idx_2024_2q = safe_float(region_agg[22], 100)  # ⚠️ 22번 컬럼
idx_2025_1q = safe_float(region_agg[25], 100)  # ⚠️ 25번 컬럼
idx_2025_2q = safe_float(region_agg[26], 100)  # ⚠️ 26번 컬럼
```

**영향**: 엑셀 형식 변경 시 모든 generator 파일 수정 필요

### 2. 시트명 하드코딩

**위치**: 모든 `*_generator.py` 파일

```python
# mining_manufacturing_generator.py:141-144
agg_sheet, _ = find_sheet_with_fallback(
    xl,
    ['A(광공업생산)집계', 'A 집계'],  # ⚠️ 시트명 하드코딩
    ['광공업생산', '광공업생산지수']
)
```

**영향**: 시트명 변경 시 코드 수정 필요

### 3. 지역 목록 하드코딩

**위치**: `mining_manufacturing_generator.py`, `regional_generator.py` 등

```python
# mining_manufacturing_generator.py:426-427
individual_regions = ['서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종', 
                      '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
```

**영향**: 지역 추가/변경 시 여러 파일 수정 필요

### 4. 보도자료 설정 하드코딩

**위치**: `config/reports.py`

```python
# config/reports.py:7-89
SUMMARY_REPORTS = [
    {
        'id': 'cover',
        'name': '표지',
        'sheet': None,
        'generator': None,
        'template': 'cover_template.html',
        ...
    },
    ...
]
```

**영향**: 보도자료 추가/변경 시 코드 수정 필요

---

## 📈 이상적인 구조 vs 현재 구조

### 이상적인 구조 (100% 스키마 기반)

```
┌─────────────────────────────────────────────────────────────┐
│                    설정 파일 (JSON/YAML)                      │
│  - 보도자료 목록                                              │
│  - 시트명 매핑                                                │
│  - 컬럼 인덱스 매핑                                           │
│  - 업종/지역명 매핑                                           │
│  - 텍스트 생성 규칙                                           │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                    범용 Generator 엔진                        │
│  - 스키마를 읽어서 동적으로 데이터 추출                         │
│  - 형식 변경 시 코드 수정 불필요                               │
└─────────────────────────────────────────────────────────────┘
```

### 현재 구조 (스키마 + 하드코딩 혼재)

```
┌─────────────────────────────────────────────────────────────┐
│                    스키마 파일 (JSON)                         │
│  - 데이터 구조 정의 ✅                                        │
│  - 업종/지역명 매핑 ✅                                        │
│  - 텍스트 생성 규칙 ✅                                        │
│  - 컬럼 인덱스 (메타데이터용, 실제 미사용) ⚠️                  │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                    개별 Generator 클래스                      │
│  - 시트명 하드코딩 ⚠️                                        │
│  - 컬럼 인덱스 하드코딩 ⚠️                                   │
│  - 업종/지역 매핑 중복 정의 ⚠️                               │
│  - 지역 목록 하드코딩 ⚠️                                     │
└─────────────────────────────────────────────────────────────┘
```

---

## 🔧 개선 방안 (향후 과제)

### 1단계: 컬럼 인덱스 스키마화
```json
{
  "excel_mapping": {
    "analysis_sheet": {
      "name": "A 분석",
      "columns": {
        "region": 3,
        "industry": 7,
        "current_quarter": 21,
        "contribution": 28
      }
    }
  }
}
```

### 2단계: 범용 데이터 추출기 구현
```python
class SchemaBasedExtractor:
    def __init__(self, schema_path):
        self.schema = load_json(schema_path)
    
    def extract(self, excel_path):
        sheet_name = self.schema['excel_mapping']['analysis_sheet']['name']
        columns = self.schema['excel_mapping']['analysis_sheet']['columns']
        # 스키마 기반으로 동적 추출
```

### 3단계: 설정 파일 외부화
- `config/reports.py` → `config/reports.json`으로 변환
- Python 코드 없이 설정만으로 보도자료 추가 가능

---

## 📊 현재 상태 요약

| 구분 | 스키마 방식 | 하드코딩 | 비율 |
|------|-------------|----------|------|
| 데이터 구조 정의 | ✅ | - | 100% 스키마 |
| 업종/지역 매핑 | ✅ (스키마) | ⚠️ (Generator 중복) | 50% 스키마 |
| 텍스트 생성 규칙 | ✅ | - | 100% 스키마 |
| 엑셀 컬럼 인덱스 | ⚠️ (메타데이터만) | ⚠️ (실제 사용) | 10% 스키마 |
| 시트명 | - | ⚠️ | 0% 스키마 |
| 보도자료 목록 | - | ⚠️ | 0% 스키마 |

**종합**: 약 **40% 스키마 방식**, **60% 하드코딩**

---

## ✅ 결론

1. **스키마가 완전히 활용되지 않음**: 스키마 파일에 컬럼 매핑 정보가 있지만, Generator에서 하드코딩으로 덮어씀
2. **중복 정의 문제**: 업종/지역 매핑이 스키마와 Generator에 모두 정의됨
3. **개선 필요**: 형식 변경에 유연하게 대응하려면 하드코딩 부분을 스키마로 이전 필요
4. **현실적 한계**: 개발 기간 제약으로 완전한 스키마 기반 전환은 향후 과제로 남김

---

*작성일: 2026년 1월 4일*

