# 지역경제동향 보도자료 자동생성

11기 국가데이터처 캡스톤 프로젝트

## 개요

PDF 파일을 입력으로 받아 Word 템플릿으로 변환하고, 엑셀 데이터를 자동으로 채워서 PDF 보도자료를 생성하는 시스템입니다.

**주요 워크플로우**: PDF 입력 → Word 템플릿 생성 → 데이터 채우기 → PDF 출력

## 주요 기능

### 새로운 Word 템플릿 워크플로우 (메인)
- **PDF → Word 템플릿 변환**: PDF 파일을 OCR을 통해 Word 템플릿으로 변환
- **디자인 보존**: PDF의 모든 페이지와 디자인을 그대로 재현
- **자동 데이터 채우기**: 엑셀 데이터를 파싱하여 객체화하고 Word 템플릿의 마커에 자동 매핑
- **PDF 출력**: 완성된 Word 문서를 PDF로 변환

### 기존 HTML 템플릿 워크플로우 (하위 호환성)
- HTML 템플릿 기반 보도자료 생성
- 엑셀 파일에서 데이터 자동 추출
- 다양한 계산 기능 지원 (합계, 평균, 증감률, 증감액 등)
- 자동 숫자 포맷팅 (천 단위 구분, 퍼센트 표시 등)

## 설치

### 필수 패키지 설치

```bash
pip install -r requirements.txt
```

### PDF 변환 도구 설치 (필수)

**Mac/Linux:**
```bash
brew install --cask libreoffice
```

**Windows:**
```bash
# docx2pdf가 자동으로 설치됩니다 (pip install -r requirements.txt)
```

## 사용 방법

### 웹 애플리케이션 (권장)

웹 브라우저를 통해 쉽게 사용할 수 있는 웹 인터페이스를 제공합니다.

1. **웹 서버 시작**:
```bash
python app.py
```

2. **브라우저에서 접속**:
   - http://localhost:8000 을 열어주세요

3. **새로운 Word 템플릿 워크플로우 사용**:
   - PDF 파일과 Excel 파일을 함께 업로드
   - 시트명, 연도, 분기 선택
   - "보도자료 생성" 버튼 클릭
   - 생성된 PDF 파일 다운로드

### API 엔드포인트

#### `/api/process-word-template` (메인)
PDF → Word 템플릿 → 데이터 채우기 → PDF 출력 전체 워크플로우

**요청 형식:**
- `pdf_file`: PDF 템플릿 파일
- `excel_file`: 엑셀 데이터 파일
- `sheet_name`: 시트명
- `year`: 연도 (선택, 자동 감지)
- `quarter`: 분기 (선택, 자동 감지)

#### `/api/pdf-to-word-template`
PDF 파일을 Word 템플릿으로 변환

#### `/api/process` (하위 호환성)
기존 HTML 템플릿 기반 워크플로우

#### `/api/pdf-to-html` (하위 호환성)
PDF를 HTML로 변환

## 마커 형식

Word 템플릿 또는 HTML 템플릿에서 다음과 같은 형식으로 마커를 사용할 수 있습니다:

- `{시트명:셀주소}`: 단일 셀 값
  - 예: `{광공업생산:A1}`

- `{시트명:셀주소:계산식}`: 계산식 적용
  - 예: `{광공업생산:A1:A5:sum}` - 합계
  - 예: `{광공업생산:A1:A5:average}` - 평균
  - 예: `{광공업생산:A1:A2:증감률}` - 증감률 (퍼센트)

- `{시트명:동적키}`: 동적 데이터
  - 예: `{광공업생산:전국_증감률}` - 전국 증감률
  - 예: `{광공업생산:상위시도1_이름}` - 상위 시도 이름
  - 예: `{광공업생산:상위시도1_증감률}` - 상위 시도 증감률

### 지원하는 계산식

- `sum`, `합계`: 합계
- `average`, `avg`, `평균`: 평균
- `max`, `최대값`, `최대`: 최대값
- `min`, `최소값`, `최소`: 최소값
- `growth_rate`, `증감률`, `증가율`: 증감률 (퍼센트)
- `growth_amount`, `증감액`, `증가액`: 증감액

## 프로젝트 구조

```
capstone/
├── templates/              # HTML 템플릿 파일 (하위 호환성)
├── flask_templates/        # Flask 웹 템플릿
├── static/                 # 정적 파일 (CSS, JS)
│   ├── css/
│   └── js/
├── image_cache/            # 참조용 이미지 캐시
├── src/                    # 소스 코드
│   ├── pdf_to_word.py          # PDF → Word 변환
│   ├── word_template_manager.py # Word 템플릿 관리
│   ├── word_template_filler.py  # Word 템플릿 채우기
│   ├── word_to_pdf.py          # Word → PDF 변환
│   ├── excel_extractor.py      # 엑셀 데이터 추출
│   ├── template_filler.py      # HTML 템플릿 채우기 (하위 호환성)
│   ├── template_manager.py    # HTML 템플릿 관리 (하위 호환성)
│   ├── calculator.py           # 계산 엔진
│   ├── data_analyzer.py        # 데이터 분석
│   ├── period_detector.py     # 연도/분기 감지
│   └── ...
├── app.py                  # Flask 웹 애플리케이션
├── requirements.txt        # Python 패키지 의존성
└── README.md
```

## 워크플로우 상세 설명

### 새로운 Word 템플릿 워크플로우

1. **PDF 입력**: 사용자가 PDF 템플릿 파일 업로드
2. **PDF → Word 변환**: 
   - PDF를 이미지로 변환 (PyMuPDF)
   - OCR을 통해 텍스트 및 레이아웃 추출 (EasyOCR/pytesseract)
   - Word 문서로 재구성 (python-docx)
3. **데이터 파싱**: 엑셀 파일에서 데이터 추출 및 객체화
4. **매핑 및 채우기**: Word 템플릿의 마커를 데이터로 치환
5. **PDF 출력**: 완성된 Word 문서를 PDF로 변환 (LibreOffice/docx2pdf)

### 기존 HTML 템플릿 워크플로우 (하위 호환성)

1. HTML 템플릿 파일 로드
2. 엑셀 데이터 추출
3. 마커 치환 및 계산
4. HTML 파일 출력

## OCR 엔진

### 지원하는 OCR 엔진

- **EasyOCR** (기본값): 한글과 영어를 지원하며 정확도가 높습니다
- **pytesseract**: Tesseract OCR 엔진 사용

### OCR 설정

- `use_easyocr`: true/false (기본값: true)
- `dpi`: PDF를 이미지로 변환할 때의 DPI (기본값: 300)

## 주의사항

- PDF 파일이 크거나 페이지가 많을 경우 처리 시간이 오래 걸릴 수 있습니다
- OCR 정확도는 PDF의 품질에 따라 달라질 수 있습니다
- 한글 PDF의 경우 EasyOCR을 사용하는 것을 권장합니다
- Word → PDF 변환을 위해서는 LibreOffice (Mac/Linux) 또는 docx2pdf (Windows)가 필요합니다
- 최대 파일 크기: 100MB

## 테스트

```bash
python src/test_basic.py
```

## 라이선스

11기 국가데이터처 캡스톤 프로젝트
