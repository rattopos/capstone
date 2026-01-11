# 📋 excel_mapping.json 검수 보고서

**검수일자**: 2026-01-11  
**검수자**: 개발팀  
**대상 파일**: `config/excel_mapping.json`

---

## 📄 현재 파일 내용

```json
{
  "manufacturing": {
    "keywords": ["광공업", "생산"],
    "header_anchors": ["지수", "증감률", "전년동기비"],
    "region_col_keywords": ["지역", "시도"],
    "code_col_keywords": ["분류", "코드", "단계"],
    "total_code_values": ["BCD", "0", "총지수", "계"]
  },
  "service": {
    "keywords": ["서비스", "생산"],
    "header_anchors": ["지수", "증감률", "전년동기비"],
    "region_col_keywords": ["지역", "시도"],
    "code_col_keywords": ["분류", "코드", "단계"],
    "total_code_values": ["E~S", "0", "총지수", "계"]
  },
  "consumption": {
    "keywords": ["소비", "소매", "판매"],
    "header_anchors": ["지수", "증감률", "전년동기비"],
    "region_col_keywords": ["지역", "시도"],
    "code_col_keywords": ["분류", "코드", "단계"],
    "total_code_values": ["총지수", "0", "계"]
  },
  "construction": {
    "keywords": ["건설", "공표"],
    "header_anchors": ["지수", "증감률", "전년동기비"],
    "region_col_keywords": ["지역", "시도"],
    "code_col_keywords": ["분류", "코드", "단계"],
    "total_code_values": ["0", "계", "총지수"]
  },
  "export": {
    "keywords": ["수출"],
    "header_anchors": ["지수", "증감률", "전년동기비"],
    "region_col_keywords": ["지역", "시도"],
    "code_col_keywords": ["분류", "코드", "단계"],
    "total_code_values": ["합계", "0", "계"]
  },
  "price": {
    "keywords": ["물가", "품목"],
    "header_anchors": ["지수", "증감률", "전년동기비"],
    "region_col_keywords": ["지역", "시도"],
    "code_col_keywords": ["분류", "코드", "단계"],
    "total_code_values": ["총지수", "0", "계"]
  },
  "employment": {
    "keywords": ["고용률", "연령별"],
    "header_anchors": ["고용률", "증감률", "전년동기비"],
    "region_col_keywords": ["지역", "시도"],
    "code_col_keywords": ["분류", "코드", "단계"],
    "total_code_values": ["계", "0", "총지수"]
  },
  "common": {
    "quarter_patterns": [
      "\\d{4}[.\\s]*\\d/4",
      "\\d{4}\\.\\d/4",
      "\\d{4} \\d/4"
    ],
    "year_patterns": [
      "\\d{4}",
      "\\d{4}\\.0"
    ],
    "header_scan_rows": 10,
    "default_header_row": 2
  }
}
```

---

## ✅ 검수 포인트별 분석

### 1. Keywords (시트명 키워드)

#### ✅ 검증 완료 항목
- **manufacturing**: `["광공업", "생산"]` ✅
  - 실제 시트명: "광공업생산"
  - `extractors/config.py`의 `SHEET_KEYWORDS`와 일치
  
- **service**: `["서비스", "생산"]` ✅
  - 실제 시트명: "서비스업생산"
  - `extractors/config.py`의 `SHEET_KEYWORDS`와 일치

- **consumption**: `["소비", "소매", "판매"]` ✅
  - 실제 시트명: "소비(소매, 추가)"
  - `extractors/config.py`의 `SHEET_KEYWORDS`와 일치

- **construction**: `["건설", "공표"]` ✅
  - 실제 시트명: "건설 (공표자료)"
  - `extractors/config.py`의 `SHEET_KEYWORDS`와 일치

- **export**: `["수출"]` ✅
  - 실제 시트명: "수출"
  - `extractors/config.py`의 `SHEET_KEYWORDS`와 일치

- **price**: `["물가", "품목"]` ✅
  - 실제 시트명: "품목성질별 물가"
  - `extractors/config.py`의 `SHEET_KEYWORDS`와 일치

- **employment**: `["고용률", "연령별"]` ✅
  - 실제 시트명: "연령별고용률"
  - `extractors/config.py`의 `SHEET_KEYWORDS`와 일치

#### ⚠️ 누락된 항목 (선택적 추가 고려)
- **import** (수입): 현재 없음 → 필요 시 추가
- **unemployment** (실업자 수): 현재 없음 → 필요 시 추가
- **population** (시도 간 이동): 현재 없음 → 필요 시 추가

---

### 2. Header Anchors (헤더 앵커)

#### ✅ 공통 헤더 앵커
모든 시트에서 공통적으로 사용되는 헤더 키워드:
- `"지수"`: ✅ 대부분의 시트에서 사용 (광공업생산지수, 서비스업생산지수 등)
- `"증감률"`: ✅ 전년동기비 증감률 컬럼 헤더
- `"전년동기비"`: ✅ 전년동기 대비 비교 컬럼 헤더

#### ⚠️ 시트별 특수 헤더 앵커
- **employment**: `["고용률", "증감률", "전년동기비"]` ✅
  - "고용률"이 추가로 포함되어 있어 더 정확함

#### 💡 개선 제안
일부 시트에서는 다음과 같은 추가 앵커가 유용할 수 있습니다:
- **export/import**: "수출액", "수입액", "억달러" 등
- **construction**: "건설수주액", "백억원" 등
- **price**: "소비자물가지수", "전년동기" 등

---

### 3. Total Code Values (총지수 값)

#### ✅ 실제 사용 값과 비교

| 시트 타입 | excel_mapping.json | 실제 사용 값 (services/summary_data.py) | 상태 |
|---------|-------------------|--------------------------------------|------|
| manufacturing | `["BCD", "0", "총지수", "계"]` | `'BCD'` | ✅ 정확 |
| service | `["E~S", "0", "총지수", "계"]` | `'E~S'` | ✅ 정확 |
| consumption | `["총지수", "0", "계"]` | `'총지수'` | ✅ 정확 |
| construction | `["0", "계", "총지수"]` | `0` (숫자) | ⚠️ 숫자 0 추가 필요 |
| export | `["합계", "0", "계"]` | `'합계'` | ✅ 정확 |
| price | `["총지수", "0", "계"]` | `'총지수'` | ✅ 정확 |
| employment | `["계", "0", "총지수"]` | `'계'` | ✅ 정확 |

#### ⚠️ 수정 필요 사항
- **construction**: `total_code_values`에 숫자 `0` 추가 필요
  - 현재: `["0", "계", "총지수"]` (문자열 "0")
  - 실제: `0` (숫자) 또는 `"0"` (문자열) 둘 다 사용 가능하도록

---

### 4. Region/Code Column Keywords

#### ✅ 검증 완료
- **region_col_keywords**: `["지역", "시도"]` ✅
  - 대부분의 시트에서 지역 컬럼 헤더에 "지역" 또는 "시도" 사용
  
- **code_col_keywords**: `["분류", "코드", "단계"]` ✅
  - 분류단계, 코드, 단계 등 다양한 명칭 사용

#### 💡 참고사항
실제 컬럼 위치는 시트마다 다를 수 있지만, 키워드 기반 탐색이므로 유연하게 대응 가능합니다.

---

### 5. Common Settings

#### ✅ 검증 완료
- **quarter_patterns**: ✅ 분기 패턴 정규식 정확
  - `"\\d{4}[.\\s]*\\d/4"`: "2024.1/4", "2025 2/4" 등
  - `"\\d{4}\\.\\d/4"`: "2024.1/4" (점 포함)
  - `"\\d{4} \\d/4"`: "2024 2/4" (공백 포함)

- **year_patterns**: ✅ 연도 패턴 정규식 정확
  - `"\\d{4}"`: "2024", "2025" 등
  - `"\\d{4}\\.0"`: "2024.0" 등

- **header_scan_rows**: `10` ✅ 적절함
- **default_header_row**: `2` ✅ 대부분의 시트에서 헤더가 2행

---

## 🔍 실제 엑셀 파일과의 비교

### 실제 사용되는 시트명 (extractors/config.py 기준)
```
"광공업생산" → keywords: ["광공업", "생산"] ✅
"서비스업생산" → keywords: ["서비스", "생산"] ✅
"소비(소매, 추가)" → keywords: ["소비", "소매", "판매"] ✅
"건설 (공표자료)" → keywords: ["건설", "공표"] ✅
"수출" → keywords: ["수출"] ✅
"수입" → keywords: ["수입"] (누락) ⚠️
"품목성질별 물가" → keywords: ["물가", "품목"] ✅
"연령별고용률" → keywords: ["고용률", "연령별"] ✅
"실업자 수" → keywords: ["실업", "실업자"] (누락) ⚠️
"시도 간 이동" → keywords: ["인구", "이동", "시도"] (누락) ⚠️
```

### 실제 사용되는 total_code 값 (services/summary_data.py 기준)
```
광공업생산: 'BCD' ✅
서비스업생산: 'E~S' ✅
소비(소매, 추가): '총지수' ✅
수출: '합계' ✅
수입: '합계' (누락) ⚠️
품목성질별 물가: '총지수' ✅
연령별고용률: '계' ✅
실업자 수: '계' (누락) ⚠️
건설 (공표자료): 0 (숫자) ⚠️
```

---

## 📝 개선 제안

### 1. 누락된 시트 타입 추가 (선택적)
```json
"import": {
  "keywords": ["수입"],
  "header_anchors": ["지수", "증감률", "전년동기비"],
  "region_col_keywords": ["지역", "시도"],
  "code_col_keywords": ["분류", "코드", "단계"],
  "total_code_values": ["합계", "0", "계"]
},
"unemployment": {
  "keywords": ["실업", "실업자"],
  "header_anchors": ["지수", "증감률", "전년동기비"],
  "region_col_keywords": ["지역", "시도"],
  "code_col_keywords": ["분류", "코드", "단계"],
  "total_code_values": ["계", "0", "총지수"]
},
"population": {
  "keywords": ["인구", "이동", "시도"],
  "header_anchors": ["인구", "이동", "증감률"],
  "region_col_keywords": ["지역", "시도"],
  "code_col_keywords": ["분류", "코드", "단계"],
  "total_code_values": ["0", "계", "총지수"]
}
```

### 2. Construction의 total_code_values 수정
```json
"construction": {
  ...
  "total_code_values": [0, "0", "계", "총지수"]  // 숫자 0 추가
}
```

---

## ✅ 최종 검수 결과

### 통과 항목
- ✅ Keywords: 모든 주요 시트 타입의 키워드 정확
- ✅ Header Anchors: 공통 헤더 앵커 적절
- ✅ Total Code Values: 주요 시트의 총지수 값 정확
- ✅ Common Settings: 패턴 및 기본값 적절

### 개선 권장 사항
1. ⚠️ **Construction의 total_code_values**: 숫자 `0` 추가 고려
2. 💡 **누락된 시트 타입**: import, unemployment, population 추가 고려 (현재 사용 여부 확인 필요)

### 전체 평가
**✅ 검수 통과** - 현재 정의된 시트 타입에 대해서는 키워드와 앵커가 실제 엑셀 파일의 고유 특징을 잘 반영하고 있습니다. 다만, 일부 시트 타입이 누락되어 있으므로 필요 시 추가를 권장합니다.

---

**검수 완료일**: 2026-01-11
