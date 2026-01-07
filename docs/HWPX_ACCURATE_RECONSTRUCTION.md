# HWPX 파일 정확한 재현 가이드

## 목차
1. [HWPX 파일 구조 이해](#hwpx-파일-구조-이해)
2. [원본 파일 분석 방법](#원본-파일-분석-방법)
3. [정확한 재현을 위한 체크리스트](#정확한-재현을-위한-체크리스트)
4. [현재 구현의 개선점](#현재-구현의-개선점)
5. [실전 예제](#실전-예제)

## HWPX 파일 구조 이해

### HWPX 파일은 ZIP 압축 형식
```
sample.hwpx (ZIP 파일)
├── [Content_Types].xml        # 파일 타입 정의
├── _rels/
│   └── .rels                  # 파일 간 관계 정의
├── Contents/
│   ├── section0.xml           # 메인 문서 내용
│   └── ...                    # 추가 섹션들
├── BinData/                   # 이미지, 바이너리 데이터
│   ├── image1.png
│   └── ...
├── Styles/                    # 스타일 정의 (선택)
│   └── styles.xml
└── manifest.xml               # 파일 목록 및 메타데이터
```

### 주요 XML 네임스페이스
- `http://www.hancom.co.kr/hwpml/2011/hwpp` - HWPML 메인
- `http://www.hancom.co.kr/hwpml/2011/vml` - 벡터 그래픽
- `http://schemas.openxmlformats.org/package/2006/relationships` - 관계 파일
- `http://schemas.openxmlformats.org/package/2006/content-types` - 콘텐츠 타입

## 원본 파일 분석 방법

### 1단계: 분석 도구 사용
```bash
# HWPX 파일 분석
python utils/hwpx_analyzer.py 원본.hwpx 분석결과.json
```

### 2단계: 수동 분석
```bash
# ZIP으로 압축 해제
unzip 원본.hwpx -d extracted/

# 주요 파일 확인
cat extracted/Contents/section0.xml | xmllint --format -
cat extracted/[Content_Types].xml
cat extracted/manifest.xml
```

### 3단계: 구조 파악
- **문단 구조**: `<Paragraph>` → `<Run>` → `<Text>` 체계
- **스타일 구조**: `<CharShape>`, `<ParaShape>` 속성 확인
- **표 구조**: `<Table>` → `<Row>` → `<Cell>` 구조
- **이미지 바인딩**: `<Picture>` 요소의 `binID`와 `BinData/` 파일 매핑

## 정확한 재현을 위한 체크리스트

### ✅ 필수 요소
- [ ] **용지 설정**: 크기, 여백, 방향 정확히 매칭
- [ ] **텍스트 서식**: 폰트명, 크기, 색상, 굵기, 기울임
- [ ] **문단 서식**: 정렬, 들여쓰기, 줄간격, 문단 여백
- [ ] **표 구조**: 행/열 수, 셀 병합, 테두리, 배경색
- [ ] **이미지**: 정확한 크기, 위치, 투명도, 바인딩
- [ ] **페이지 나누기**: 강제 페이지 나누기 위치

### ✅ 고급 요소
- [ ] **스타일 시트**: 반복되는 스타일을 스타일 시트로 정의
- [ ] **텍스트 상자**: 특정 위치의 텍스트 상자
- [ ] **각주/미주**: 각주 및 미주 정보
- [ ] **머리말/꼬리말**: 페이지 머리말/꼬리말
- [ ] **책갈피/하이퍼링크**: 문서 내 연결
- [ ] **차트/도형**: 벡터 그래픽 요소

## 현재 구현의 개선점

### 1. 스타일 정보 보존 강화

**현재 문제:**
```python
# 현재: 하드코딩된 기본값만 사용
char_prop.set('size', '1000')  # 10pt 고정
```

**개선 방안:**
```python
def _extract_css_styles(self, element):
    """HTML 요소에서 CSS 스타일 추출"""
    style_attr = element.get('style', '')
    styles = {}
    for prop in style_attr.split(';'):
        if ':' in prop:
            key, value = prop.split(':', 1)
            styles[key.strip()] = value.strip()
    return styles

def _css_to_hwp_units(self, css_value):
    """CSS 값을 HWP 단위로 변환"""
    # px → HWP 단위 (1px ≈ 1270 HWP units)
    # pt → HWP 단위 (1pt = 200 HWP units)
    # ...
```

### 2. 표 정확도 향상

**현재 문제:**
- 열 너비가 자동 계산됨
- 테두리 스타일 미지원
- 셀 배경색 미지원

**개선 방안:**
```python
def _convert_table_with_styles(self, table_element):
    """스타일을 보존한 표 변환"""
    # 1. HTML 표에서 스타일 추출
    table_styles = self._extract_table_styles(table_element)
    
    # 2. 열 너비 계산 (CSS width 속성 반영)
    col_widths = self._calculate_column_widths(table_element)
    
    # 3. 테두리 및 배경색 처리
    borders = self._extract_border_styles(table_element)
    backgrounds = self._extract_background_styles(table_element)
```

### 3. 이미지 위치 정확도

**현재 문제:**
- 이미지 위치가 문단 내에만 배치됨
- 절대 위치 지정 불가

**개선 방안:**
```python
def _convert_image_with_position(self, img_element):
    """위치 정보를 포함한 이미지 변환"""
    # CSS position 추출
    position = self._extract_position(img_element)
    
    if position['type'] == 'absolute':
        # 절대 위치 지정
        pic.set('x', str(position['x']))
        pic.set('y', str(position['y']))
```

### 4. 원본 HWPX와의 비교 검증

**추가 기능:**
```python
def compare_with_original(self, generated_hwpx, original_hwpx):
    """생성된 HWPX와 원본 비교"""
    gen_analyzer = HWPXAnalyzer(generated_hwpx)
    orig_analyzer = HWPXAnalyzer(original_hwpx)
    
    gen_structure = gen_analyzer.analyze()
    orig_structure = orig_analyzer.analyze()
    
    # 비교 로직
    differences = self._find_differences(gen_structure, orig_structure)
    return differences
```

## 실전 예제

### 예제 1: 원본 HWPX 분석

```python
from utils.hwpx_analyzer import HWPXAnalyzer

# 원본 파일 분석
analyzer = HWPXAnalyzer('원본보도자료.hwpx')
structure = analyzer.analyze()

# 분석 결과 확인
analyzer.print_summary()

# JSON으로 저장
analyzer.export_analysis('원본_구조.json')
```

### 예제 2: 스타일 정보 추출

```python
# 원본에서 문단 스타일 확인
section_data = structure['sections']['Contents/section0.xml']
for para in section_data['paragraphs']:
    para_shape = para.get('para_shape', {})
    print(f"정렬: {para_shape.get('alignment')}")
    print(f"들여쓰기: {para_shape.get('indent')}")
    
    for run in para['runs']:
        char_shape = run.get('char_shape', {})
        print(f"폰트 크기: {char_shape.get('size')}")
        print(f"굵기: {char_shape.get('bold')}")
```

### 예제 3: 개선된 변환기 사용

```python
from utils.hwpx_converter import HWPXConverter

converter = HWPXConverter()

# 원본 구조를 참조하여 변환
pages = [
    {'html': '<div>...</div>', 'title': '페이지 1'},
    # ...
]

# 변환 옵션 (원본 분석 결과 기반)
options = {
    'page_def': {
        'width': '210000',  # 원본과 동일
        'height': '297000',
        'margins': {...}    # 원본 여백과 동일
    },
    'default_styles': {
        'font_size': '1000',
        'font_name': '맑은 고딕',
        # ...
    }
}

hwpx_data = converter.convert_html_to_hwpx(pages, 2025, 2, options)
```

## 권장 워크플로우

### 1단계: 원본 분석
```
원본.hwpx → hwpx_analyzer.py → 구조.json
```

### 2단계: 구조 기반 변환기 설정
```
구조.json → 변환 옵션 설정 → hwpx_converter.py
```

### 3단계: 변환 및 검증
```
HTML → hwpx_converter.py → 생성.hwpx
생성.hwpx ↔ 원본.hwpx → 비교 검증
```

### 4단계: 반복 개선
- 차이점 발견 → 변환 로직 수정 → 재검증

## 주의사항

1. **HWP 버전 차이**: 한글 2020, 2022, 2024 등 버전별로 XML 구조가 약간 다를 수 있음
2. **비표준 요소**: 한글 전용 기능(책갈피, 필드 등)은 완벽한 재현이 어려울 수 있음
3. **이미지 품질**: 이미지 압축 방식에 따라 품질 차이가 있을 수 있음
4. **테스트 환경**: 한글 프로그램이 설치된 환경에서 직접 확인 필수

## 참고 자료

- 한컴오피스 개발자 문서 (공개 여부 확인 필요)
- OOXML 표준 (HWPX는 OOXML 기반)
- 현재 프로젝트: `utils/hwpx_converter.py` (기본 구현)
- 분석 도구: `utils/hwpx_analyzer.py` (새로 추가)

## 결론

HWPX 파일을 **정확히 똑같이** 재현하려면:

1. ✅ 원본 파일을 상세히 분석
2. ✅ 모든 스타일 정보를 추출하고 보존
3. ✅ 표, 이미지, 레이아웃을 정확히 매칭
4. ✅ 변환 결과를 원본과 비교 검증
5. ✅ 반복적으로 개선

현재 구현(`hwpx_converter.py`)은 기본적인 변환은 가능하지만, **100% 동일한 재현**을 위해서는 위의 개선사항들을 적용해야 합니다.

