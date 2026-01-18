# 시도별 경제동향 통합 보고서 생성 가이드

## 개요

시도별 경제동향 보도자료는 17개 시도별로 생산, 소비·건설, 수출·입, 고용, 물가, 국내인구이동 등 모든 부문의 경제지표를 한 페이지에 통합하여 제시하는 보고서입니다.

## 변경사항 (2025-01-18)

### 1. 파일 이동 및 정리

#### 레거시 이동
- `templates/regional_template.html` → `templates/legacy/regional_template.html.legacy`
- `templates/regional_generator.py` → `templates/legacy/regional_generator.py.legacy`

**이유**: 부문별 통합 방식으로 변경하여 기존 개별 시도 Generator는 더 이상 필요 없음

#### 새 생성 파일
- `templates/regional_economy_by_region_schema.json` - 시도별 경제동향 데이터 스키마
- `templates/regional_economy_by_region_template.html` - 시도별 경제동향 HTML 템플릿

### 2. Generator 구조 변경

#### 기존 방식 (레거시)
```
individual regional_generator.py
  ├─ 각 시도별로 데이터 추출
  └─ 제한된 부문 지원
```

#### 새로운 방식 (통합)
```
unified_generator.py
  ├─ UnifiedReportGenerator (기존 부문별 Generator 기반)
  ├─ RegionalEconomyByRegionGenerator (새 시도별 통합 Generator)
  │   └─ 각 시도별로:
  │       ├─ 생산 (광공업, 서비스업)
  │       ├─ 소비·건설
  │       ├─ 수출·입
  │       ├─ 고용 (고용률, 실업률)
  │       ├─ 물가
  │       └─ 국내인구이동
  └─ RegionalReportGenerator (레거시 호환)
```

### 3. 설정 추가 (config/report_configs.py)

새로운 보고서 유형 추가:

```python
'regional_economy_by_region': {
    'name': '시도별 경제동향',
    'report_id': 'regional_economy_by_region',
    'sheets': {
        'analysis': None,  # 분석시트 없음
        'aggregation': [모든 부문의 집계시트],
    },
    'is_regional_by_region': True,  # 플래그
    'require_analysis_sheet': False,  # 집계시트만 사용
    'has_nationwide': False,  # 시도별이므로 전국 제외
}
```

## 사용 방법

### 1. 기본 사용법

```python
from templates.unified_generator import RegionalEconomyByRegionGenerator

# Generator 초기화
gen = RegionalEconomyByRegionGenerator(
    excel_path='path/to/분석표.xlsx',
    year=2025,
    quarter=3
)

# 시도별 섹션 데이터 추출
section = gen.extract_regional_section('서울', 'mining')
# 반환값: {'narrative': [...], 'table': {...}}

# 모든 시도의 통합 데이터 추출
all_data = gen.extract_all_regions_data()
# 반환값: {'서울': {...}, '부산': {...}, ...}
```

### 2. 템플릿 렌더링

```python
from jinja2 import Environment, FileSystemLoader

env = Environment(loader=FileSystemLoader('templates'))
template = env.get_template('regional_economy_by_region_template.html')

# 시도별 데이터와 함께 렌더링
for region_name, region_data in all_data.items():
    context = {
        'region_name': region_name,
        'report_info': {'year': 2025, 'quarter': 3},
        'sections': region_data['sections'],  # 부문별 섹션
    }
    html = template.render(context)
    # HTML 저장 등의 작업
```

## 데이터 구조

### RegionalEconomyByRegionGenerator.REGIONS

17개 시도 목록:

```python
[
    {'code': 11, 'name': '서울', 'full_name': '서울특별시'},
    {'code': 21, 'name': '부산', 'full_name': '부산광역시'},
    ...
    {'code': 39, 'name': '제주', 'full_name': '제주특별자치도'},
]
```

### 섹션 데이터 구조

```python
{
    'narrative': [
        '서울의 광공업생산은 ...',
        '전년동기대비 증가 ...'
    ],
    'table': {
        'periods': ['2025/3Q'],
        'data': [
            {
                'indicator': '광공업생산지수',
                'values': [100.5, 2.5]
            }
        ]
    }
}
```

## 템플릿 구조

### 주요 섹션

1. **생산** (광공업생산, 서비스업생산)
2. **소비·건설** (소비, 건설)
3. **수출·입** (수출, 수입)
4. **고용** (고용률, 취업자)
5. **물가** (소비자물가)
6. **국내 인구이동** (순인구이동)

각 섹션은 다음을 포함:
- ○ 불릿 나레이션
- 데이터 표 (시계열 + 증감률)

## 테스트

```bash
python test_regional_economy_generator.py
```

## 주요 특징

### 통합 구조의 장점

1. **일관성**: 모든 부문이 동일한 Generator 기반 사용
2. **확장성**: 새 부문 추가 시 기존 로직 활용 가능
3. **유지보수성**: 부문별 Generator 변경 시 자동 반영
4. **성능**: 부문별 Generator 인스턴스 캐싱

### 그래프 제외

요청사항에 따라 나레이션과 표만 생성하고 그래프는 제외합니다.

### 시도별 데이터 합산

각 시도는 부문별로 독립적인 데이터를 가지며, 필요시 시도 간 비교가 가능합니다.

## 향후 개선

1. **더 정교한 나레이션**: 템플릿 기반 나레이션 생성 개선
2. **동적 주목도**: 증감률에 따른 주목도 표시
3. **비교 기능**: 시도 간 비교 분석 추가
4. **다중 분기**: 여러 분기 데이터 동시 제공

## 참고

- **레거시 코드**: `templates/legacy/` 에 보존됨
- **기존 시도별 보고서**: `templates/regional_output/` 에 생성된 파일 유지
- **호환성**: `RegionalReportGenerator` 클래스로 레거시 기능 유지 가능
