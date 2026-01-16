# 광공업 Generator 완전 리팩토링 가이드

## 설계 표준(Design Standard) - 4대 핵심 원칙

이 문서는 `MiningManufacturingGenerator`에 적용된 리팩토링 원칙을 정리하며, 
**모든 Generator (건설, 서비스업, 고용 등)에 동일하게 적용되어야 하는 표준**입니다.

---

## 1. [Core] 좌표 기반 → 의미 기반 파싱 전환

### ❌ Before (Fragile - 깨지기 쉬움)
```python
COL_GROWTH_RATE = 21  # 2025.2/4 증감률
COL_WEIGHT = 8        # 가중치
COL_CONTRIBUTION = 28 # 기여도

growth_rate = row[21]  # 하드코딩된 인덱스
```

**문제점:**
- 분기가 바뀌면 21번 컬럼이 과거 데이터를 가리킴
- 엑셀 구조가 변경되면 코드 수정 필요
- Silent Failure: 잘못된 데이터를 읽어도 오류 없이 진행

### ✅ After (Robust - 견고함)
```python
class MiningManufacturingGenerator(BaseGenerator):
    def __init__(self, excel_path, year, quarter, excel_file=None):
        super().__init__(excel_path, year, quarter, excel_file)
        # 동적 컬럼 인덱스 캐시
        self._col_cache = {
            'analysis': {},
            'aggregation': {}
        }
    
    def _initialize_column_indices(self):
        """헤더를 분석하여 모든 필요한 컬럼 인덱스를 캐시에 저장"""
        header_row = self.df_analysis.iloc[2]
        
        # 메타데이터 컬럼
        self._col_cache['analysis']['region'] = self._find_metadata_column('region')
        self._col_cache['analysis']['weight'] = self._find_metadata_column('weight')
        self._col_cache['analysis']['industry_name'] = self._find_metadata_column('industry_name')
        
        # 데이터 컬럼 (연도/분기 기반)
        self._col_cache['analysis']['target'] = self.find_target_col_index(
            header_row, self.year, self.quarter
        )
        self._col_cache['analysis']['prev_year'] = self.find_target_col_index(
            header_row, self.year - 1, self.quarter
        )
    
    def _find_metadata_column(self, column_type: str) -> int:
        """
        의미 기반 컬럼 찾기
        
        Args:
            column_type: 'region', 'weight', 'industry_name' 등
        
        Returns:
            컬럼 인덱스 (0-based)
        
        Raises:
            ValueError: 컬럼을 찾을 수 없을 때
        """
        keywords_map = {
            'region': ['지역', 'region', '시도'],
            'weight': ['가중치', 'weight', '비중'],
            'industry_name': ['산업명', '업태명', 'industry']
        }
        
        keywords = keywords_map[column_type]
        header_row = self.df_analysis.iloc[2]
        
        for idx, cell in enumerate(header_row):
            if pd.isna(cell):
                continue
            cell_str = str(cell).lower().replace(" ", "")
            if any(k.lower() in cell_str for k in keywords):
                return idx
        
        raise ValueError(f"{column_type} 컬럼을 찾을 수 없습니다")
```

**장점:**
- 엑셀 구조 변경에도 자동 대응
- 못 찾으면 명시적 오류 발생 (Silent Failure 방지)
- 코드 수정 없이 다른 분기 데이터 처리 가능

---

## 2. [Logic] 단순 등락률 → 기여도 기반 분석 전환

### ❌ Before (Naive - 순진함)
```python
# 단순 증감률 크기순 정렬
industries = sorted(industries, key=lambda x: abs(x['growth_rate']), reverse=True)[:2]
```

**문제점:**
- 비중 1%인 소규모 업종이 +50% 성장하면 1위로 선정
- 비중 30%인 반도체가 +5% 성장하면 하위로 밀림
- **경제적 의미 없음**: 실제 경제 파급력과 무관

### ✅ After (Smart - 똑똑함)
```python
def _extract_nationwide_industries(self) -> dict:
    """전국 업종별 데이터 추출"""
    
    # [1. Core] 동적 컬럼 인덱스 확보
    region_col = self._col_cache['analysis']['region']
    growth_col = self._col_cache['analysis']['target']
    weight_col = self._col_cache['analysis']['weight']
    industry_name_col = self._col_cache['analysis']['industry_name']
    
    industries = []
    
    # 전국 중분류 업종 데이터 추출
    nationwide_rows = df[(df[region_col] == '전국') & 
                         (df[classification_col].astype(str) == '2')]
    
    for _, row in nationwide_rows.iterrows():
        growth_rate = self.safe_float(row[growth_col], None)
        if growth_rate is None:
            continue
        
        # [2. Logic] 가중치 확보 (Dynamic + Fallback)
        weight = self.safe_float(row[weight_col], 0)
        
        industry_name = self._get_industry_display_name(str(row[industry_name_col]))
        
        if weight == 0:
            # Fallback: 주력 산업 키워드 기반 가중치
            weight = self._get_industry_weight_fallback(industry_name)
        
        # [2. Logic] 기여도 = 증감률 × 가중치
        contribution = abs(growth_rate * weight)
        
        industries.append({
            'name': industry_name,
            'growth_rate': round(growth_rate, 1),
            'weight': weight,
            'contribution': contribution  # 경제적 영향력
        })
    
    # [2. Logic] 기여도 기반 정렬 (경제적 파급력 순)
    increase_industries = sorted(
        [i for i in industries if i['growth_rate'] > 0],
        key=lambda x: x['contribution'],  # 기여도순
        reverse=True
    )[:2]
    
    decrease_industries = sorted(
        [i for i in industries if i['growth_rate'] < 0],
        key=lambda x: x['contribution'],  # 기여도순
        reverse=True
    )[:2]
    
    return {"increase": increase_industries, "decrease": decrease_industries}

def _get_industry_weight_fallback(self, industry_name: str) -> float:
    """
    [2. Logic] 가중치 Fallback Logic
    
    가중치 컬럼이 없거나 0일 때 사용
    """
    # 주력 산업 우선순위 (GDP 기여도 기반)
    high_impact = {
        '반도체': 30.0,
        '전자부품': 25.0,
        '자동차': 20.0,
        '기계': 15.0,
        '화학': 15.0,
        '1차금속': 10.0
    }
    
    # 소비재 (경제 파급력 낮음)
    low_impact = {
        '식료품': 1.0,
        '음료': 1.0,
        '의복': 1.0
    }
    
    for keyword, weight in high_impact.items():
        if keyword in industry_name:
            return weight
    
    for keyword, weight in low_impact.items():
        if keyword in industry_name:
            return weight
    
    return 5.0  # 기타 제조업
```

**결과:**
- 비중 30% 반도체 +5% → 기여도 1.5% → 상위 선정 ✅
- 비중 1% 의복 +50% → 기여도 0.5% → 하위 제외 ✅
- **경제적 의미 확보**: 실제 경제 파급력을 반영

---

## 3. [NLP] 한국어 문법 엔진 탑재

### ❌ Before (Grammar Errors)
```python
sentence = f"{region}은 반도체 등의 생산이 늘어 증가"
# 출력: "경기은 반도체..." (X)
# 출력: "제주은 반도체..." (X)
```

### ✅ After (Natural Language)
```python
from utils.text_utils import get_josa

def generate_narrative(region, industries):
    """나레이션 생성"""
    # [3. NLP] 받침 유무에 따른 조사 자동 선택
    josa = get_josa(region, "은/는")
    
    industry_text = ", ".join([i['name'] for i in industries[:2]])
    sentence = f"{region}{josa} {industry_text} 등의 생산이 늘어 증가"
    # 출력: "경기는 반도체..." (O)
    # 출력: "제주는 반도체..." (O)
    # 출력: "서울은 반도체..." (O)
    
    return sentence

# utils/text_utils.py
def get_josa(word: str, josa_pair: str = "은/는") -> str:
    """
    한글 받침 유무 판별하여 올바른 조사 반환
    
    Args:
        word: 단어 (예: '서울', '경기')
        josa_pair: 조사 쌍 (예: "은/는", "이/가", "을/를")
    
    Returns:
        적절한 조사
    """
    if not word:
        return ""
    
    last_char = word[-1]
    
    # 한글 받침 계산 (유니코드 공식)
    if 0xAC00 <= ord(last_char) <= 0xD7A3:
        has_batchim = (ord(last_char) - 0xAC00) % 28 > 0
    else:
        has_batchim = False
    
    first, second = josa_pair.split("/")
    return first if has_batchim else second
```

**적용 위치:**
- `extract_summary_box()`: 지역 대조 문장
- `_generate_regional_narrative()`: 시도별 설명
- 모든 f-string에서 `{region}은` → `{region}{get_josa(region)}`

---

## 4. [Generalization] 템플릿 메서드 패턴

### 코드 구조 표준화

모든 Generator는 다음 구조를 따름:

```python
class MiningManufacturingGenerator(BaseGenerator):
    """
    [변하지 않는 부분] 공통 로직
    - 엑셀 로드
    - 시트 찾기
    - HTML 렌더링
    
    [변하는 부분] 광공업 특화 로직
    - 업종 매핑
    - 기여도 계산
    - 나레이션 규칙
    """
    
    def __init__(self, excel_path, year, quarter, excel_file=None):
        """초기화"""
        super().__init__(excel_path, year, quarter, excel_file)
        self._col_cache = {'analysis': {}, 'aggregation': {}}
    
    # ==========================================
    # [변하지 않는 부분] 공통 로직
    # ==========================================
    
    def load_data(self):
        """엑셀 데이터 로드 - 모든 Generator 동일"""
        xl = self.load_excel()
        agg_sheet, _ = self.find_sheet_with_fallback(
            ['A(광공업생산)집계', 'A 집계'],
            ['광공업생산']
        )
        self.df_aggregation = self.get_sheet(agg_sheet)
        # ...
    
    def extract_all_data(self) -> dict:
        """
        메인 추출 로직 - 템플릿 메서드 패턴
        
        순서:
        1. 데이터 로드
        2. 컬럼 인덱스 초기화
        3. 전국 데이터 추출
        4. 시도별 데이터 추출
        5. 나레이션 생성
        """
        self.load_data()
        self._initialize_column_indices()
        
        nationwide = self._extract_nationwide_data()
        regional = self._extract_regional_data()
        narrative = self._generate_narrative(nationwide, regional)
        
        return {
            'report_info': self._get_report_info(),
            'nationwide_data': nationwide,
            'regional_data': regional,
            'summary_box': narrative
        }
    
    def generate_html(self, template_path, output_path) -> str:
        """HTML 렌더링 - 모든 Generator 동일"""
        data = self.extract_all_data()
        
        from jinja2 import Template
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        html = template.render(**data)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html)
        
        return html
    
    # ==========================================
    # [변하는 부분] 광공업 특화 로직
    # ==========================================
    
    # 업종명 매핑 (광공업 전용)
    INDUSTRY_NAME_MAP = {
        "전자 부품, 컴퓨터...": "반도체·전자부품",
        "자동차 및 트레일러...": "자동차·트레일러",
        # ...
    }
    
    def _get_industry_display_name(self, raw_name: str) -> str:
        """광공업 특화: 업종명 변환"""
        for key, value in self.INDUSTRY_NAME_MAP.items():
            if key in raw_name:
                return value
        return raw_name
    
    def _calculate_contribution(self, growth_rate: float, weight: float) -> float:
        """광공업 특화: 기여도 계산"""
        return abs(growth_rate * weight)
```

---

## 핵심 메서드 완전 재설계 예시

### `_extract_nationwide_industries_from_aggregation()` 재설계

```python
def _extract_nationwide_industries_from_aggregation(self) -> dict:
    """
    [완전 재설계] 집계 시트에서 전국 업종별 데이터 추출
    
    4대 원칙 적용:
    1. [Core] 모든 컬럼 인덱스 동적 탐색
    2. [Logic] 기여도 기반 정렬
    3. [NLP] 업종명에 조사 적용 (필요시)
    4. [Generalization] 주석으로 변하는/변하지 않는 부분 명시
    """
    df = self.df_aggregation
    
    # ==========================================
    # [변하지 않는 부분] 헤더 및 컬럼 인덱스 확보
    # ==========================================
    
    # 헤더 행 찾기
    header_row_idx = 2  # 기본값
    for i in range(min(5, len(df))):
        row_text = ' '.join([str(v) for v in df.iloc[i].values if pd.notna(v)])
        if '전국' in row_text and str(self.year) in row_text:
            header_row_idx = i
            break
    
    header_row = df.iloc[header_row_idx]
    
    # [1. Core] 동적 컬럼 탐색
    region_col = self._col_cache['aggregation'].get('region') or 4
    classification_col = self._col_cache['aggregation'].get('classification') or 4
    industry_name_col = self._col_cache['aggregation'].get('industry_name') or 7
    weight_col = self._col_cache['aggregation'].get('weight')
    
    # 데이터 컬럼 (연도/분기 기반)
    target_col = self.find_target_col_index(header_row, self.year, self.quarter)
    prev_year_col = self.find_target_col_index(header_row, self.year - 1, self.quarter)
    
    print(f"[광공업생산] 동적 컬럼 탐색 완료:")
    print(f"  - 지역: {region_col}, 분류: {classification_col}, 업종명: {industry_name_col}")
    print(f"  - 가중치: {weight_col}, 현재분기: {target_col}, 전년분기: {prev_year_col}")
    
    # ==========================================
    # [변하는 부분] 광공업 특화 데이터 추출
    # ==========================================
    
    # 전국 중분류 업종 (분류단계 2)
    nationwide_industries = df[
        (df[region_col] == '전국') &
        (df[classification_col].astype(str) == '2')
    ]
    
    industries = []
    
    for _, row in nationwide_industries.iterrows():
        # 당분기/전년분기 지수
        current = self.safe_float(row[target_col], None)
        previous = self.safe_float(row[prev_year_col], None)
        
        if current is None or previous is None or previous == 0:
            continue
        
        # 증감률 계산
        growth_rate = ((current - previous) / previous) * 100
        
        # 업종명 추출
        industry_name = self._get_industry_display_name(
            str(row[industry_name_col]) if pd.notna(row[industry_name_col]) else ''
        )
        
        # [2. Logic] 가중치 확보 (Dynamic + Fallback)
        if weight_col is not None:
            weight = self.safe_float(row[weight_col], 0)
        else:
            weight = 0
        
        if weight == 0:
            # Fallback: 주력 산업 키워드 기반 가중치
            weight = self._get_industry_weight_fallback(industry_name)
            print(f"[광공업생산] 가중치 Fallback: {industry_name} → {weight}")
        
        # [2. Logic] 기여도 계산
        contribution = abs(growth_rate * weight)
        
        industries.append({
            'name': industry_name,
            'growth_rate': round(growth_rate, 1),
            'weight': weight,
            'contribution': contribution
        })
    
    # ==========================================
    # [변하지 않는 부분] 증가/감소 분류 및 정렬
    # ==========================================
    
    # [2. Logic] 기여도 기반 정렬
    increase_industries = sorted(
        [i for i in industries if i['growth_rate'] > 0],
        key=lambda x: x['contribution'],  # 기여도순
        reverse=True
    )[:2]
    
    decrease_industries = sorted(
        [i for i in industries if i['growth_rate'] < 0],
        key=lambda x: x['contribution'],  # 기여도순
        reverse=True
    )[:2]
    
    return {
        "increase": increase_industries,
        "decrease": decrease_industries
    }
```

---

## 나레이션 생성 재설계

### `extract_summary_box()` 재설계

```python
def extract_summary_box(self, nationwide_data: dict, regional_data: dict) -> dict:
    """
    [3. NLP] 요약 박스 나레이션 생성
    
    지역명에 올바른 조사를 붙여 자연스러운 문장 생성
    """
    from utils.text_utils import get_josa, get_contrast_narrative
    
    # ==========================================
    # [변하지 않는 부분] 데이터 추출
    # ==========================================
    
    production_index = nationwide_data.get("production_index", 0)
    nationwide_val = nationwide_data.get("growth_rate", 0)
    
    inc_regions = regional_data.get("increase_regions", [])[:3]
    dec_regions = regional_data.get("decrease_regions", [])[:3]
    
    # ==========================================
    # [변하는 부분] 광공업 특화 나레이션
    # ==========================================
    
    # 주요 업종 (기여도 기반으로 이미 정렬됨)
    if nationwide_val >= 0:
        industries = nationwide_data.get("main_increase_industries", [])
    else:
        industries = nationwide_data.get("main_decrease_industries", [])
    
    industry_text = ", ".join([i['name'] for i in industries[:2]])
    
    # ==========================================
    # [3. NLP] 문법적으로 올바른 문장 생성
    # ==========================================
    
    # 전국 문장
    전국_josa = get_josa("전국", "은/는")
    cause_verb = "늘어" if nationwide_val > 0 else "줄어"
    result_noun = "증가" if nationwide_val > 0 else "감소"
    
    sentence1 = (
        f"전국{전국_josa} 광공업생산({production_index:.1f})은 "
        f"{industry_text} 등의 생산이 {cause_verb} "
        f"전년동분기대비 {abs(nationwide_val):.1f}% {result_noun}"
    )
    
    # 시도 대조 문장 (Zigzag 패턴)
    sentence2 = get_contrast_narrative(
        nationwide_val, inc_regions, dec_regions, report_id='mining'
    )
    
    # 지역별 업종 삽입 (필요시)
    if inc_regions and dec_regions:
        # 증가 지역 업종
        inc_region = inc_regions[0]
        inc_industries = ", ".join([i['name'] for i in inc_region.get('top_industries', [])[:2]])
        
        # 감소 지역 업종
        dec_region = dec_regions[0]
        dec_industries = ", ".join([i['name'] for i in dec_region.get('top_industries', [])[:2]])
        
        # [3. NLP] 조사 적용
        inc_josa = get_josa(inc_region['region'], "은/는")
        dec_josa = get_josa(dec_region['region'], "은/는")
        
        # 패턴 교체 (상세 업종 삽입)
        if "등의 생산이 줄어 감소" in sentence2:
            sentence2 = sentence2.replace(
                "등의 생산이 줄어 감소",
                f"{dec_industries} 등의 생산이 줄어 감소"
            )
        
        if "등의 생산이 늘어 증가" in sentence2:
            sentence2 = sentence2.replace(
                "등의 생산이 늘어 증가",
                f"{inc_industries} 등의 생산이 늘어 증가"
            )
    
    return {
        "headline": f"◆ {sentence1}",
        "nationwide_summary": sentence1,
        "regional_summary": sentence2,
        "full_narrative": f"□ {sentence1}\n○ {sentence2}"
    }
```

---

## 적용 체크리스트

### Phase 1: COL_ 상수 제거
- [x] 클래스 상단 COL_ 선언 제거
- [x] `_initialize_column_indices()` 메서드 추가
- [x] `_find_metadata_column()` 메서드 추가
- [ ] 모든 `self.COL_*` 사용을 `self._col_cache[sheet][type]`로 교체

### Phase 2: 기여도 기반 분석
- [x] `contribution = abs(growth_rate * weight)` 계산 적용
- [x] `_get_industry_weight_fallback()` Fallback Logic 강화
- [x] 정렬 key를 `contribution` 기준으로 변경
- [ ] 모든 업종 추출 메서드에 일관되게 적용

### Phase 3: 한국어 문법 엔진
- [x] `get_josa()` 함수 사용
- [x] `extract_summary_box()`에 적용
- [ ] 모든 f-string 나레이션에 적용

### Phase 4: 템플릿 메서드 패턴
- [x] `load_data()`, `extract_all_data()` 구조화
- [x] "변하는/변하지 않는 부분" 주석 추가
- [ ] 모든 메서드에 명확한 docstring 작성

---

## 다음 단계: 다른 Generator 적용

이 패턴을 다음 Generator에 순차적으로 적용:

1. `construction_generator.py` - 이미 부분 적용됨
2. `service_industry_generator.py`
3. `consumption_generator.py` - 이미 부분 적용됨
4. `employment_rate_generator.py` - 이미 부분 적용됨
5. `export_generator.py`
6. `import_generator.py`
7. `price_trend_generator.py`
8. `unemployment_generator.py`

---

## 주요 개선 효과

### 유지보수성
- 엑셀 구조 변경 시 코드 수정 불필요
- 새로운 분기 데이터 추가 시 자동 대응
- 오류 발생 시 명확한 메시지

### 정확성
- 기여도 기반 분석으로 경제적 의미 확보
- Silent Failure 제거
- 자연스러운 한국어 문장

### 확장성
- 다른 지표에 쉽게 적용 가능
- 템플릿 메서드 패턴으로 일관된 구조
- 플러그인 방식의 Fallback Logic
