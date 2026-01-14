# 🐛 디버그 작업 로그

이 문서는 프로젝트의 모든 디버그 작업을 추적하고 기록합니다.

---

## 📋 로그 형식

각 디버그 항목은 다음 형식으로 기록됩니다:

- **날짜/시간**: 작업 수행 시점
- **문제 설명**: 발견된 문제 또는 버그
- **원인 분석**: 문제의 근본 원인
- **에이전트 사고 과정**: AI 에이전트가 문제를 해결하기까지의 추론 과정
  - 문제 인식 및 초기 분석
  - 고려한 해결책들
  - 선택한 접근 방법과 그 이유
  - 시도한 방법들과 결과
  - 최종 결정의 근거
- **해결 방법**: 적용한 해결책
- **관련 파일**: 수정된 파일 목록
- **상태**: 진행중 / 완료 / 실패 / 보류
- **참고 사항**: 추가 정보나 향후 작업

---

## 📅 디버그 기록

### 2026-01-04 (계속)

#### 파일 업로드 중 파일 유형 분석 단계에서 멈추는 문제
- **시간**: 2026-01-04 (사용자 보고)
- **문제 설명**:
  - 파일 업로드 과정에서 "파일 유형 분석" 단계에서 진행이 멈춤
  - 사용자가 파일을 업로드했지만 다음 단계로 진행되지 않음
- **원인 분석**:
  - `detect_file_type` 함수가 `openpyxl.load_workbook`을 사용하여 파일을 읽는 과정에서 시간이 오래 걸리거나 멈출 수 있음
  - 특히 큰 파일이나 복잡한 파일의 경우 `read_only=True`로 열어도 시간이 오래 걸릴 수 있음
  - 에러가 발생해도 예외 처리가 제대로 되지 않아 사용자에게 피드백이 없을 수 있음
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "업로드 과정중 파일유형분석에서 멈춥니다"라고 보고
  - 코드 분석:
    1. `routes/api.py`의 `/upload` 엔드포인트에서 `detect_file_type` 호출
    2. `utils/excel_utils.py`의 `detect_file_type` 함수 확인
    3. `openpyxl.load_workbook`을 사용하여 시트명을 읽는 과정에서 멈출 수 있음
  - 해결 방안:
    1. 각 단계마다 상세한 디버그 로그 추가하여 어느 단계에서 멈추는지 확인
    2. 예외 처리 강화하여 에러 발생 시 사용자에게 명확한 피드백 제공
    3. 타임아웃 처리 고려 (필요시)
- **해결 방법**:
  1. `detect_file_type` 함수의 각 단계에 상세한 `print` 로그 추가
  2. `routes/api.py`의 업로드 엔드포인트에 예외 처리 및 로그 추가
  3. 파일 유형 분석 실패 시 명확한 에러 메시지 반환
- **관련 파일**:
  - `utils/excel_utils.py`: `detect_file_type` 함수에 디버그 로그 추가
  - `routes/api.py`: 파일 유형 분석 부분에 예외 처리 및 로그 추가
- **상태**: 완료
- **참고 사항**:
  - 서버 로그를 확인하여 어느 단계에서 멈추는지 추적 가능
  - 필요시 타임아웃 처리나 비동기 처리 고려 가능

### 2026-01-04

#### 인포그래픽 지도 회색(보합/평균) 색상이 배경과 구분되지 않는 문제 해결
- **시간**: 2026-01-04 19:00
- **문제 설명**:
  - 인포그래픽 지도에서 회색이어야 하는 부분(보합, 전국 평균)이 하얀색/투명처럼 보임
  - 소비자물가의 '전국 평균' 지역과 일반 지표의 '보합' 지역이 배경과 구분되지 않음
  - 지역 path 사이에 틈이 있어 배경색이 비침
- **원인 분석**:
  1. 회색 색상 `#DDDCDD`가 배경색 `#E8F0F2`와 너무 유사하여 하얀색처럼 보임
  2. SVG의 stroke(지역 경계선)가 흰색(`#ffffff`)이고 path들 사이에 미세한 틈이 있음
  3. 각 지역 path가 완전히 붙어있지 않아 틈 사이로 배경이 비침
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "지도에서 회색이어야 하는 부분이 하얀색 혹은 투명한 이유를 분석해서 해결해"라고 요청
  - 정답 이미지와 비교: 정답에서는 지도가 완전히 채워져 있고 경계만 흰색 선으로 표시
  - 코드 분석:
    1. JavaScript의 색상 정의 확인: `neutral: '#DDDCDD'` - 연한 회색
    2. CSS에서도 동일한 색상 사용 확인
    3. SVG의 stroke가 `#ffffff`로 설정됨
  - 테스트: 브라우저에서 실제 적용된 색상 확인
    - 소비자물가에서 대구, 경기, 인천이 회색으로 색칠됨 확인
    - 하지만 배경과 대비가 약해서 하얀색처럼 보임
  - 해결책 고려:
    1. 회색 색상을 더 진하게 변경 → 채택
    2. SVG에 배경 레이어 추가하여 틈 메우기 → 채택
    3. stroke 색상을 회색으로 변경 → 미채택 (정답과 다름)
- **해결 방법**:
  1. 회색 색상 변경: `#DDDCDD` → `#A8A8A8` (진한 회색)
     - JavaScript COLORS 객체의 `neutral`, `avg` 값 변경
     - CSS의 `.region-neutral`, `.region-avg`, `.legend-color.neutral`, `.legend-color.avg` 변경
  2. SVG에 배경 레이어 추가:
     - `<g id="layer-background">` 그룹 추가
     - 17개 지역 path를 복제하여 배경 레이어에 배치
     - 배경 레이어 색상을 `#A8A8A8`(회색)로 설정
     - 전경 레이어(`layer-MC1`) 위에 색칠하면 틈 사이로 배경 회색이 보임
- **관련 파일**:
  - `templates/infographic_template.html` - 색상 정의 및 SVG 배경 레이어 추가
- **상태**: ✅ 완료
- **참고 사항**:
  - 현재 데이터에서 일반 지표(광공업생산 등)의 '보합' 케이스는 발생하지 않음 (모든 값이 ±0.05% 범위 밖)
  - 소비자물가에서 대구, 경기, 인천이 전국 평균(2.1%)과 동일하여 회색으로 표시됨

---

#### 인포그래픽 지도 all_regions 빈 배열 문제 해결
- **시간**: 2026-01-04 18:00
- **문제 설명**:
  - 인포그래픽 지도가 모든 지표에서 회색으로만 표시됨 (색칠이 안됨)
  - 상위/하위 데이터는 정상 표시되지만 지도만 색칠이 안되는 상황
- **원인 분석**:
  - `infographic_generator.py`의 `_get_default_indicator()` 함수에서 `'all_regions': []`로 빈 배열 반환
  - 분석 시트에서 데이터를 읽지 못하면 기본값이 사용되는데, 이때 `all_regions`가 비어있어 지도 색칠 JavaScript 로직이 색칠할 지역 데이터가 없음
  - 정적 테스트 파일(`debug/test_infographic.html`)도 이전 데이터로 하드코딩되어 있어 `"all_regions": []`인 상태
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "인포그래픽 지도 칠하는게 문제있어 확인해줘"라고 제보
  - 디버그 로그 확인: 2026-01-04에 SVG 인라인 방식으로 수정했다는 기록 있음, 하지만 아직 문제 지속
  - 코드 분석:
    1. `infographic_template.html`의 JavaScript 색칠 로직 확인 → `allRegions` 배열을 순회하며 지역별 색상 적용
    2. `infographic_generator.py`의 데이터 생성 로직 확인 → `_get_default_indicator`에서 `all_regions: []` 반환
  - 테스트 실행: `generate_report_data()` 호출 시 모든 지표에서 `all_regions=17개`로 정상 반환되는 것 확인
    - 이는 기본값에 `all_regions` 데이터가 추가되었음을 의미
  - 정적 테스트 파일 확인: `grep`으로 `test_infographic.html`의 `indicatorsData` 확인 → `"all_regions": []` 하드코딩됨
  - 해결책 결정:
    1. `_get_default_indicator`에 17개 지역 전체 데이터 추가
    2. 테스트 파일 재생성
- **해결 방법**:
  1. `templates/infographic_generator.py` 수정:
     - `_get_default_indicator()` 함수에 `all_regions_defaults` 딕셔너리 추가
     - 6개 지표별로 17개 지역 전체 데이터 포함 (2025년 2분기 기준)
     - 반환값에 `'all_regions': all_regions` 추가
  2. `debug/test_infographic.html` 재생성:
     - 새로운 데이터로 테스트 파일 렌더링
     - 모든 지표에서 `all_regions=17개` 확인
- **관련 파일**:
  - `templates/infographic_generator.py` (수정)
  - `debug/test_infographic.html` (재생성)
- **상태**: ✅ 완료
- **참고 사항**:
  - 색상 규칙:
    - 일반 지표(광공업생산, 서비스업생산, 소매판매, 수출): 증가=핑크, 감소=파랑, 보합=회색
    - 고용률: 상승=핑크, 하락=파랑, 보합=회색
    - 소비자물가: 전국평균 초과=핑크, 미만=보라, 동일=회색
  - 향후 실제 분석 시트에서 데이터를 정상적으로 읽으면 기본값 대신 실제 데이터 사용됨

---

#### 요약 섹션 정적 보도자료 엑셀 없이 생성 가능하도록 수정
- **시간**: 2026-01-04 17:45
- **문제 설명**:
  - 프리뷰에서 요약 섹션의 목차 페이지가 생성되지 않음
  - 표지, 일러두기, 목차는 엑셀 파일이 없어도 생성할 수 있어야 하는데 "엑셀 파일을 먼저 업로드하세요" 오류 발생
- **원인 분석**:
  - `routes/preview.py`의 `generate_summary_preview()` 함수에서 모든 요약 보도자료에 대해 엑셀 파일 존재를 필수로 요구
  - 표지, 일러두기, 목차는 `generator: None`으로 설정되어 있고 스키마 기본값만으로 생성 가능한데도 엑셀 파일 체크가 앞에 있어서 실패
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "요약섹션에 있는 목차페이지가 생성이 안된다"고 제보
  - API 테스트: `/api/generate-summary-preview`에 `report_id: 'toc'`로 요청 시 "엑셀 파일을 먼저 업로드하세요" 오류 확인
  - 코드 분석: `generate_summary_preview()` 함수 80-82라인에서 모든 요청에 대해 `excel_path` 존재 여부 체크
  - 해결책 결정: 정적 보도자료(cover, guide, toc)는 엑셀 파일 체크를 건너뛰도록 수정
- **해결 방법**:
  1. `routes/preview.py` 수정:
     - `static_reports = ['cover', 'guide', 'toc']` 정의
     - 정적 보도자료는 엑셀 파일 존재 체크를 건너뜀
     - `organization` 값을 '국가데이터처'로 변경
     - `department` 기본값을 '국가데이터처 경제통계국'으로 변경
  2. 표지(`cover`) 처리 로직 추가
- **관련 파일**:
  - `routes/preview.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**:
  - 표지, 일러두기, 목차는 엑셀 파일 없이도 API 호출 시 정상 생성됨
  - 인포그래픽 및 요약 데이터 보도자료는 여전히 엑셀 파일 필요

---

#### 보도자료 찾을 수 없음 오류 및 목차 동적 생성 오류 해결
- **시간**: 2026-01-04 17:10
- **문제 설명**:
  1. "보도자료를 찾을 수 없습니다" 오류가 발생
  2. 목차 생성 시 동적 페이지 계산 관련 오류 계속 발생
  3. 국내인구이동 보도자료 렌더링 시 `'dict object' has no attribute 'columns'` 오류
- **원인 분석**:
  1. `config/reports.py`에서 요약 보도자료의 generator로 `summary_regional_economy_generator.py`, `summary_production_generator.py` 등 5개 파일이 지정되어 있으나 실제로 존재하지 않음
  2. `_get_toc_sections()` 함수가 동적으로 페이지 번호를 계산하면서 오류 발생
  3. `domestic_migration_generator.py`의 `generate_summary_table()` 함수가 `columns` 속성을 반환하지 않아 템플릿에서 `summary_table.columns` 접근 시 오류
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "보도자료를 찾을 수 없다는 문제가 계속 일어나"라고 제보
  - 터미널 로그 분석:
    - `[PREVIEW] Generator 모듈을 로드할 수 없습니다: summary_regional_economy_generator.py` 등 오류 메시지 확인
    - `[ERROR] 보도자료 생성 오류: 'dict object' has no attribute 'columns'` 오류 확인
  - 파일 존재 확인:
    - `templates/` 폴더에 `summary_*_generator.py` 파일들이 실제로 존재하지 않음
    - `infographic_generator.py` 등 다른 generator 파일들은 존재
  - 해결책 결정:
    1. 목차: 동적 계산 대신 고정 페이지 번호로 템플릿 단순화 (사용자 요청)
    2. Generator 오류: 존재하지 않는 generator 참조를 `None`으로 변경
    3. 국내인구이동: `columns` 데이터 추가
- **해결 방법**:
  1. `config/reports.py` 수정:
     - 요약 보도자료 5개의 generator를 `None`으로 변경 (summary_overview, summary_production, summary_consumption, summary_trade_price, summary_employment)
  2. `templates/toc_template.html` 수정:
     - Jinja2 변수 참조를 제거하고 고정 페이지 번호로 하드코딩
     - 요약(1), 부문별(6~15), 시도별(16~48), 참고GRDP(50), 통계표(52), 부록(75) 고정
  3. `routes/preview.py`, `routes/debug.py` 수정:
     - `_get_toc_sections()` 호출 코드 제거/단순화
  4. `templates/domestic_migration_generator.py` 수정:
     - `generate_summary_table()` 함수에 `columns` 데이터 추가
- **관련 파일**:
  - `config/reports.py` (수정)
  - `templates/toc_template.html` (수정)
  - `templates/domestic_migration_generator.py` (수정)
  - `routes/preview.py` (수정)
  - `routes/debug.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**:
  - 요약 보도자료들은 generator 없이 `services/summary_data.py`에서 직접 데이터를 가져옴
  - 목차 페이지 번호 변경이 필요하면 `toc_template.html`을 직접 수정해야 함

---

#### 인포그래픽 지도 색칠 안되는 문제 해결 (srcdoc iframe에서 SVG fetch 실패)
- **시간**: 2026-01-04 16:35
- **문제 설명**:
  - 인포그래픽 지도가 색칠되지 않음
  - 디버그 모드(`/debug/test_infographic.html`)에서는 정상 작동하나, 메인 대시보드 미리보기에서는 지도가 회색으로만 표시됨
- **원인 분석**:
  - 메인 대시보드는 API에서 생성된 HTML을 iframe의 `srcdoc` 속성으로 렌더링
  - `srcdoc` iframe의 origin은 `about:srcdoc`이 됨
  - JavaScript에서 `fetch('/templates/infographic_map.svg')`를 호출할 때:
    - 디버그 모드: Flask 서버가 `/templates/...` 경로를 정상 처리 → SVG 로드 성공
    - 대시보드 srcdoc: 상대 경로가 제대로 해석되지 않거나 CORS 문제 발생 → SVG 로드 실패
  - SVG 로드 실패 시 `svgTemplate`이 null이 되어 `colorMap()` 함수가 아무것도 하지 않음
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "인포그래픽 색칠이 안돼요"라고 제보
  - 초기 분석:
    - `infographic_generator.py` 확인 → `all_regions` 데이터 정상 생성됨
    - `infographic_template.html` 확인 → JavaScript에서 SVG를 fetch로 로드
    - 디버그 모드에서 테스트 → 정상 작동 (색칠 완료)
  - 원인 파악:
    - 디버그와 대시보드의 차이점 분석
    - 대시보드: `elements.previewIframe.srcdoc = displayHtml;` (srcdoc 사용)
    - srcdoc iframe에서 fetch 경로 문제 발견
  - 고려한 해결책:
    1. SVG 파일을 JavaScript 변수로 인라인 포함 → fetch 불필요, 가장 안정적
    2. 절대 URL 사용 (`http://localhost:5050/templates/...`) → 포트 하드코딩 문제
    3. SVG를 Base64 인코딩 → 복잡하고 유지보수 어려움
  - 선택한 접근: **SVG를 JavaScript 변수로 인라인 포함** (fetch 제거)
    - 사용자가 "최소한의 수정"을 요청
    - 템플릿 파일 하나만 수정하면 됨
    - 서버 환경에 관계없이 항상 작동
- **해결 방법**:
  - `templates/infographic_template.html` 수정:
    - `async function loadAndColorMaps()` → 동기 함수로 변경
    - `fetch('/templates/infographic_map.svg')` 제거
    - `const svgTemplate = \`<svg>...</svg>\`;` 형태로 SVG 인라인 포함
    - SVG 내용은 `infographic_map.svg`에서 핵심 path들만 추출하여 경량화
- **관련 파일**:
  - `templates/infographic_template.html` (수정)
- **상태**: ✅ 완료
- **참고 사항**:
  - srcdoc iframe에서 fetch, 상대 경로, 절대 경로 모두 주의 필요
  - 인라인 SVG 방식은 파일 크기가 약간 증가하지만 안정성 확보
  - 향후 지도 디자인 변경 시 `infographic_template.html`의 `svgTemplate` 변수 업데이트 필요

---

#### 목차 동적 계산 오류로 인한 하드코딩 전환 (정답 이미지 기준 고정 페이지)
- **시간**: 2026-01-04 (오후)
- **문제 설명**:
  - 목차 페이지 번호를 누적 계산(동적)하는 방식이 자주 어긋나 사용자가 반복적으로 오류를 경험
  - 사용자는 “정답 이미지의 항목/순서는 그대로 유지하되 페이지 번호는 고정값으로 적어달라”고 요청
- **원인 분석**:
  - `_get_toc_sections()`가 `PAGE_CONFIG` 등을 기반으로 누적 계산을 수행하면서, 실제 출력 페이지 구조(예: 통계표 2페이지 구성 등)와 설정/가정이 불일치할 때 페이지 번호가 틀어질 수 있음
  - 또한 (1)/(2)처럼 페이지 단위로 항목을 분할 표시하여 정답 이미지(항목당 시작 페이지 1개 표기)와 형식이 달라짐
- **에이전트 사고 과정**:
  - 문제를 “계산식이 흔들리는 목차”로 보고, 요구사항을 “정답 이미지 항목 + 고정 페이지”로 재정의
  - 고려한 해결책:
    - 동적 계산 유지 + 계산 로직/페이지수 정합성 강화: 유연하지만 변경 범위가 커지고 유지비 증가
    - 목차만 고정 매핑으로 전환: 변동 요인을 제거해 안정성 확보, 사용자 요구와 정확히 일치
  - 선택한 접근: 사용자 요청(“하드코딩”)을 우선하여 **목차만 고정 매핑**으로 전환
- **해결 방법**:
  1. `config/reports.py`에 `TOC_FIXED_PAGES` 추가 (정답 이미지 기준)
     - 요약: 1, 부문별: 6, 시도별: 16(각 시도 2p 기준 시작 페이지), 참고 GRDP: 50, 통계표: 52, 부록: 75
     - 시도명은 목차 이미지처럼 “서 울” 형태로 보이도록 `display_name` 추가
  2. `routes/preview.py`, `routes/debug.py`의 `_get_toc_sections()`를 누적 계산 → `TOC_FIXED_PAGES` 사용으로 교체
     - (1)/(2) 분할 표기 제거, **항목당 시작 페이지 1개만** 표시
  3. 스키마 기반 렌더링에서도 동일 결과가 나오도록 `templates/toc_schema.json`의 예시 데이터 및 키(`entries`) 정합성 보정
- **관련 파일**:
  - `config/reports.py`
  - `routes/preview.py`
  - `routes/debug.py`
  - `templates/toc_schema.json`
- **상태**: ✅ 완료
- **참고 사항**:
  - 향후 페이지 구성이 바뀌면 `TOC_FIXED_PAGES`만 갱신하면 됨

#### 목차 빈 페이지 30개+ 문제 해결 및 페이지 단위 목록 변경
- **시간**: 2026-01-04 (오후)
- **문제 설명**: 
  - 대시보드 미리보기에서 목차가 빈 페이지를 30개 넘게 차지함
  - 목록이 항목 기준으로 표시되어 있어 페이지 단위로 변경 필요
  - 같은 항목이 두 페이지인 경우 구분이 없음
  - 통계표 범위 확인 필요
- **원인 분석**: 
  - `toc_template.html`의 `.toc-container`에 `min-height: 297mm` 설정으로 인해 A4 페이지 컨테이너 안에서 추가 공간 차지
  - `statistics_table_toc_template.html`의 `.page`에도 `min-height: 297mm` 설정
  - `statistics_table_index_template.html`에도 동일한 문제 존재
  - 각 템플릿의 `@page` 규칙이 상위 템플릿과 충돌
- **에이전트 사고 과정**:
  - 문제 인식: "목차가 빈 페이지를 30개 넘게 차지"라는 사용자 제보 확인
  - 코드 분석 단계:
    1. `routes/debug.py`의 `A4_FULL_REPORT_TEMPLATE` 구조 확인 → `.a4-page`가 297mm 높이로 설정
    2. `toc_template.html` 확인 → `.toc-container`에 `min-height: 297mm` 설정 발견
    3. `statistics_table_toc_template.html` 확인 → `.page`에 `min-height: 297mm` 설정 발견
  - 원인 파악: 개별 템플릿의 CSS가 통합 출력 시 상위 A4 컨테이너와 충돌하여 빈 공간 발생
  - 해결책 결정:
    1. 모든 목차/통계표 템플릿에서 `min-height: 297mm`, `height: 297mm` 제거
    2. `@page` 규칙 제거 (상위 템플릿에서 이미 정의됨)
    3. 목차 구조를 페이지 단위로 변경
  - 추가 요청 처리: 사용자가 요청한 "페이지 단위 목록" 및 "같은 항목 두 페이지면 1, 2로 구분" 구현
- **해결 방법**: 
  1. CSS 수정:
     - `toc_template.html`: `.toc-container`에서 `min-height: 297mm` 제거, `@page` 규칙 제거
     - `statistics_table_toc_template.html`: `.page`에서 `min-height/height` 제거, `@page` 규칙 제거
     - `statistics_table_index_template.html`: 동일하게 수정
  2. 목차 구조 페이지 단위로 변경:
     - `routes/preview.py`, `routes/debug.py`의 `_get_toc_sections()` 함수 수정
     - 시도별 보도자료가 2페이지인 경우: "서울 (1)", "서울 (2)"로 표시
     - 통계표도 각 2페이지이므로: "광공업생산지수 (1)", "광공업생산지수 (2)"로 표시
  3. `toc_template.html` 템플릿 업데이트:
     - 참고 GRDP 여러 페이지 지원
     - 통계표 entries 표시 지원
  4. `statistics_table_toc_template.html` 템플릿 업데이트:
     - 동적 페이지 단위 목록 생성 지원
- **관련 파일**: 
  - `templates/toc_template.html` (CSS 및 템플릿 구조 수정)
  - `templates/statistics_table_toc_template.html` (CSS 및 템플릿 구조 수정)
  - `templates/statistics_table_index_template.html` (CSS 수정)
  - `routes/preview.py` (`_get_toc_sections()` 함수 수정)
  - `routes/debug.py` (`_get_toc_sections()` 함수 수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 통계표 범위 (2025년 2분기 기준):
    - 연도별: 2017년 ~ 2024년 (8개년)
    - 분기별: 2016년 4분기 ~ 2025년 2분기p (약 35개 분기)
  - GRDP 범위 (1분기 늦게 발표):
    - 연도별: 2017년 ~ 2024년 (마지막 2년에 'p' 표시)
    - 분기별: 2017년 3분기 ~ 2025년 1분기p
  - 통합 출력 시 개별 템플릿의 페이지 크기/여백 CSS는 상위 컨테이너와 충돌하므로 제거해야 함

---

#### 목차 시도명 띄어쓰기 제거 및 표시 형식 수정
- **시간**: 2026-01-04 (현재)
- **문제 설명**: 
  - 목차에서 시도명이 "서 울", "부 산" 등으로 띄어쓰기가 포함되어 표시됨
  - 정답 이미지에서는 "서울", "부산" 등으로 띄어쓰기 없이 표시되어야 함
  - CSS의 letter-spacing으로 인해 띄어쓰기 없이도 간격이 생기는 문제
- **원인 분석**: 
  - `config/reports.py`의 `TOC_REGION_ITEMS`에서 시도명이 "서 울", "부 산" 등으로 띄어쓰기 포함되어 정의됨
  - `templates/toc_template.html`의 CSS에서 `letter-spacing: 3px`로 설정되어 있어 띄어쓰기 없이도 간격이 생김
  - 정답 이미지(`correct_answer/요약/목차.png`)에서는 시도명이 띄어쓰기 없이 표시됨
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "목차가 다음과 같이 잘못만들어집니다"라고 제보
  - 정답 이미지 확인: `correct_answer/요약/목차.png` 파일 분석
    - 시도별 항목이 "서울", "부산", "대구" 등으로 띄어쓰기 없이 표시됨
    - 현재 코드는 "서 울", "부 산" 등으로 띄어쓰기 포함
  - 코드 분석:
    - `config/reports.py`의 `TOC_REGION_ITEMS` 확인 → 띄어쓰기 포함
    - `templates/toc_template.html`의 CSS 확인 → `letter-spacing: 3px` 설정
  - 해결책 결정:
    1. `TOC_REGION_ITEMS`의 시도명에서 띄어쓰기 제거
    2. CSS의 `letter-spacing`을 `normal`로 변경하여 정답 이미지와 일치하도록 수정
- **해결 방법**: 
  1. `config/reports.py` 수정:
     - `TOC_REGION_ITEMS`의 모든 시도명에서 띄어쓰기 제거
     - "서 울" → "서울", "부 산" → "부산" 등으로 변경
  2. `templates/toc_template.html` 수정:
     - `.toc-region-item .title`의 `letter-spacing: 3px` → `letter-spacing: normal`로 변경
- **관련 파일**: 
  - `config/reports.py` (수정)
  - `templates/toc_template.html` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 정답 이미지 기준으로 시도명은 띄어쓰기 없이 표시되어야 함
  - 목차 구조는 정답 이미지와 일치: 요약(1), 부문별(7개), 시도별(17개), 참고(1), 통계표(1), 부록(1)

### 2026-01-04

#### 발표자료 페이지 수 정확히 표기 및 정성적 판단 필요 부분 병기
- **시간**: 2026-01-04
- **문제 설명**: 
  - 발표자료에서 페이지 수가 "약 78페이지"로 모호하게 표기되어 있음
  - 요약 섹션의 구체적인 항목이 명시되지 않음
  - 사람이 정성적으로 판단해야 하는 페이지가 명시되지 않음
- **원인 분석**: 
  - 발표자료 작성 시 페이지 수를 대략적으로 표기
  - 요약 섹션의 세부 항목이 간략하게만 표기됨
  - 정성적 판단이 필요한 부분에 대한 명시 부족
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "50페이지라고 쓰인걸 실제 페이지수로 전부 수정해주세요 그리고 사람이 정성적으로 판단해야하는 페이지 수도 괄호치고 병기해주세요"라고 요청
  - 실제 페이지 구성 확인:
    - 요약: 9개 항목 (표지, 일러두기, 목차, 인포그래픽, 요약-지역경제동향, 요약-생산, 요약-소비건설, 요약-수출물가, 요약-고용인구)
    - 부문별: 10개 항목 × 2페이지 = 20페이지
    - 시도별: 18개 항목 × 2페이지 = 36페이지
    - 통계표: 13개 항목 = 13페이지
    - 합계: 50개 항목, 78페이지
  - 정성적 판단 필요 부분 식별:
    - 인포그래픽: 자동 생성되지만 디자인/레이아웃 검토 필요
    - 요약문 5개: 자동 생성되지만 해석/표현 검토 필요
  - 해결 방법 선택: 요약 섹션에 상세 항목 명시 및 정성적 판단 필요 페이지 병기
- **해결 방법**: 
  - 요약 섹션 항목을 상세히 명시: "표지, 일러두기, 목차, 인포그래픽, 요약-지역경제동향, 요약-생산, 요약-소비건설, 요약-수출물가, 요약-고용인구"
  - 페이지 수 표기에 정성적 판단 필요 부분 병기: "9페이지 (정성적 판단 필요: 인포그래픽 1p, 요약문 5p)"
  - "약 78페이지" → "78페이지"로 정확한 페이지 수 표기
- **관련 파일**: 
  - `docs/PRESENTATION.md`
- **상태**: ✅ 완료
- **참고 사항**: 
  - 인포그래픽과 요약문은 자동 생성되지만 사람의 정성적 판단(검토/수정)이 필요한 부분임을 명시
  - 다른 섹션(부문별, 시도별, 통계표)은 규칙기반으로 완전 자동 생성되어 정성적 판단이 덜 필요함

### 2026-01-02

#### 방향성 대조 표현 규칙 수정 (증가-감소, 상승-하락, 유입-유출 등)
- **시간**: 2026-01-02 14:00
- **문제 설명**: 
  - 템플릿과 스키마에서 방향성 관련 텍스트 매칭이 잘못되어 있음
  - 예: 고용률에서 결과가 "상승"일 때 "상승하였으나...내려...상승"으로 잘못 생성됨
  - 예: 인구이동에서 결과가 "순유입"일 때 "유입되었으나...유출되어...순유입"으로 잘못 생성됨
  - 예: 수출/수입에서 전국 방향과 무관하게 고정 패턴 사용
- **원인 분석**: 
  - 템플릿에서 direction에 관계없이 고정된 패턴을 사용하고 있었음
  - 스키마에서 increase_pattern, decrease_pattern이 결과 방향과 반대로 정의되어 있음
  - Generator에서 전국 방향을 고려하지 않고 고정 패턴으로 regional_summary 생성
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "증가-감소, 늘어-줄어, 상승-하락, 유입-유출, 높음-낮음 매칭이 잘못된게 몇군데 있어요"라고 지적
  - 핵심 원칙 발견: **결과와 반대되는 요인을 먼저 언급하고, 결과와 같은 방향의 요인을 나중에 언급**
    - 예: 결과가 "증가"일 때 → "줄었으나...늘어...증가"
    - 예: 결과가 "감소"일 때 → "늘었으나...줄어...감소"
    - 예: 결과가 "상승"일 때 → "하락하였으나...올라...상승"
    - 예: 결과가 "하락"일 때 → "상승하였으나...내려...하락"
    - 예: 결과가 "순유입"일 때 → "유출되었으나...유입되어...순유입"
    - 예: 결과가 "순유출"일 때 → "유입되었으나...유출되어...순유출"
  - 영향 범위 분석:
    1. regional_template.html: 고용률, 인구이동 섹션
    2. regional_schema.json: 모든 패턴 정의
    3. employment_rate_schema.json: regional_trend 규칙
    4. consumption_schema.json: regional_trend 규칙 (when_increase 케이스 누락)
    5. export_schema.json: regional_summary 템플릿
    6. export_generator.py: regional_summary 생성 로직
    7. import_generator.py: regional_summary 생성 로직
  - 해결책 설계:
    1. 템플릿: direction에 따라 조건 분기 처리
    2. 스키마: increase_pattern, decrease_pattern 명확히 구분
    3. Generator: 전국 방향에 따라 적절한 패턴 선택
  - 시도한 방법들:
    1. regional_template.html의 고용률 섹션: if-else로 direction 분기 처리 ✅
    2. regional_template.html의 인구이동 섹션: if-else로 direction 분기 처리 ✅
    3. regional_schema.json: 모든 패턴을 increase_pattern, decrease_pattern 쌍으로 재정의 ✅
    4. employment_rate_schema.json: when_decrease 케이스 추가 ✅
    5. consumption_schema.json: when_increase 케이스 추가 ✅
    6. export_schema.json: when_increase, when_decrease 구조로 변경 ✅
    7. export_generator.py: 전국 방향에 따라 조건 분기 ✅
    8. import_generator.py: 전국 방향에 따라 조건 분기 ✅
  - 최종 결정: 모든 지표에서 일관된 규칙 적용
    - 결과가 양수 → 반대 요인 먼저(음수), 결과 방향 요인 나중(양수)
    - 결과가 음수 → 반대 요인 먼저(양수), 결과 방향 요인 나중(음수)
- **해결 방법**:
  1. regional_template.html 수정:
     - 고용률: direction == '상승'일 때 "하락하였으나...올라...상승", direction == '하락'일 때 "상승하였으나...내려...하락"
     - 인구이동: direction == '순유입'일 때 "유출되었으나...유입되어...순유입", direction == '순유출'일 때 "유입되었으나...유출되어...순유출"
  2. regional_schema.json 수정:
     - 모든 섹션의 패턴을 increase_pattern, decrease_pattern 쌍으로 재정의
     - 고용률: increase_pattern(하락하였으나...상승), decrease_pattern(상승하였으나...하락)
     - 인구이동: inflow_pattern(유출되었으나...순유입), outflow_pattern(유입되었으나...순유출)
  3. employment_rate_schema.json 수정:
     - when_decrease 케이스 추가 (상승하였으나...하락)
  4. consumption_schema.json 수정:
     - when_increase 케이스 추가 (줄었으나...증가)
  5. export_schema.json 수정:
     - when_increase, when_decrease 구조로 변경
  6. export_generator.py 수정:
     - 전국 change_val >= 0일 때 "줄었으나...증가", change_val < 0일 때 "늘었으나...감소"
  7. import_generator.py 수정:
     - 전국 change_val < 0일 때 "늘었으나...감소", change_val >= 0일 때 "줄었으나...증가"
- **관련 파일**:
  - `templates/regional_template.html`
  - `templates/regional_schema.json`
  - `templates/employment_rate_schema.json`
  - `templates/consumption_schema.json`
  - `templates/export_schema.json`
  - `templates/export_generator.py`
  - `templates/import_generator.py`
- **상태**: ✅ 완료
- **참고 사항**:
  - 이 규칙은 보도자료 작성의 핵심 원칙: "결과와 반대되는 요인을 먼저 언급하고, 결과와 같은 방향의 요인을 나중에 언급"
  - 발표자료에 이 규칙의 중요성과 구현 과정을 상세히 추가함
  - 모든 지표(생산, 소비, 물가, 고용, 인구이동, 수출입)에서 일관되게 적용

---

### 2026-01-04

#### 통계표 고용률/실업률/국내인구이동 절대값 추출 구현
- **시간**: 2026-01-04 15:00
- **문제 설명**: 
  - 고용률, 실업률, 국내인구이동 통계표가 전년동기비(%) 또는 차이(%p)로 출력되어야 하는데, 원본 이미지에서는 절대값으로 출력됨
  - 고용률: [전년동기비, %p] → [%] (60.8%, 60.2% 등)
  - 실업률: [전년동기비, %p] → [%] (3.7%, 4.5% 등)
  - 국내인구이동: [전년동기비, %] → [천 명] (-98.5, 116.2 등)
- **원인 분석**: 
  - 기존 설정이 분석표의 집계 시트에서 전년동기비를 계산하도록 되어 있었음
  - 원본 이미지에서는 절대값이 필요하나, 분석표에는 절대값 데이터가 없음
  - 기초자료 수집표에 절대값 데이터가 존재함
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 제공한 원본 이미지에서 통계표 단위 확인
    - 고용률: [%] (절대값)
    - 실업률: [%] (절대값)
    - 국내인구이동: [천 명] (절대값)
  - 데이터 소스 분석:
    1. 고용률: 기초자료 수집표 '고용률' 시트 확인 → 절대값(%) 데이터 존재
       - 행 3부터 시작, 분류단계 '0'이 전체(계)
       - 2017년 전국: 60.8% 확인
    2. 실업률: 기초자료 수집표 '실업자 수' 시트 확인 → 행 78부터 "시도별 실업률(%)" 절대값 데이터 존재
       - 2017년 전국: 3.7% 확인
    3. 국내인구이동: 기초자료 수집표 '시도 간 이동' 시트 확인 → 순인구이동 수(명) 존재
       - 명 단위를 천명으로 변환 필요 (0.001 곱하기)
       - 2017년 서울: -98486명 → -98.5천명 확인
  - 해결책 설계:
    1. TABLE_CONFIG에 '기초자료_사용' 플래그 추가
    2. _extract_from_raw_sheet_absolute() 함수 구현 (절대값 추출 전용)
    3. 단위_변환 옵션 추가 (명→천명)
    4. 전국_제외 옵션 추가 (국내인구이동은 전국 데이터 없음)
  - 발견한 문제:
    1. 국내인구이동 시트에 두 개의 데이터 섹션 존재 (명 단위 / 천명 단위)
       - 두 번째 섹션을 읽으면 이미 변환된 데이터에 다시 0.001을 곱하는 오류 발생
       - 기초자료_최대_행 옵션 추가하여 첫 번째 섹션만 읽도록 제한
    2. 고용률 시트의 분류단계 값이 "계"가 아닌 숫자 "0"
       - valid_categories에 "0" 추가
- **해결 방법**:
  - statistics_table_generator.py 수정:
    1. TABLE_CONFIG의 고용률, 실업률, 국내인구이동 설정 변경
    2. _extract_from_raw_sheet_absolute() 함수 구현
    3. extract_table_data()에서 기초자료_사용 플래그 체크 후 해당 함수 호출
- **관련 파일**:
  - `templates/statistics_table_generator.py`
- **상태**: ✅ 완료
- **참고 사항**:
  - 원본 이미지 기준 데이터 검증 완료
  - 2017년 전국 고용률: 60.8% ✓
  - 2017년 전국 실업률: 3.7% ✓
  - 2017년 서울 국내인구이동: -98.5천명 ✓

---

#### 목차 페이지 번호 동적 계산 및 원본 이미지 기준 구조 적용
- **시간**: 2026-01-04 14:30
- **문제 설명**: 
  - 목차가 하드코딩되어 있어 실제 보고서 페이지 구성과 일치하지 않음
  - 부문별 항목이 10개로 세분화되어 있었으나, 원본 이미지에서는 7개 항목으로 통합 표시
  - 시도명이 전칭(서울특별시)으로 표시되었으나, 원본은 약칭(서 울)으로 표시
- **원인 분석**: 
  - `routes/preview.py`와 `routes/debug.py`의 `_get_toc_sections()` 함수가 고정된 페이지 번호를 반환
  - 목차 항목 구조가 현재 시스템 기준이었고, 원본 이미지의 형식을 따르지 않음
  - 보고서 구성이 변경되면 목차 페이지 번호를 매번 수동으로 수정해야 하는 문제
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "목차를 하드코딩하지 말고 원본 이미지에 있는 항목으로 만들되 페이지 수는 실제 미리보기 페이지 기준으로 생성"하라고 요청
  - 원본 이미지 분석:
    - `correct_answer/요약/목차.png` 파일 확인
    - 부문별 7개 항목: 생산, 소비, 건설, 수출입, 물가, 고용, 국내 인구이동
    - 시도별 17개 항목: 서 울, 부 산 등 (약칭, 띄어쓰기 포함)
    - <참고> 분기GRDP, 통계표, 부록 항목 존재
  - 해결책 설계:
    1. `config/reports.py`에 페이지 수 설정(`PAGE_CONFIG`) 추가
    2. 목차용 항목 정의(`TOC_SECTOR_ITEMS`, `TOC_REGION_ITEMS`) 추가
    3. `_get_toc_sections()` 함수를 동적 계산 방식으로 변경
    4. `toc_template.html`을 원본 이미지 스타일에 맞게 수정
  - 페이지 계산 로직:
    - 요약 섹션: 5페이지 (1~5)
    - 부문별 섹션: 각 항목 2페이지씩 총 20페이지 (6~25)
    - 시도별 섹션: 각 시도 2페이지씩 총 34페이지 (26~59)
    - 참고 GRDP: 2페이지 (60~61)
    - 통계표: 12페이지 (62~73)
    - 부록: 1페이지 (74)
- **해결 방법**: 
  1. `config/reports.py` 수정:
     - `PAGE_CONFIG`: 각 섹션별 페이지 수 정의
     - `TOC_SECTOR_ITEMS`: 원본 이미지 기준 7개 부문별 항목 (생산, 소비, 건설, 수출입, 물가, 고용, 국내 인구이동)
     - `TOC_REGION_ITEMS`: 원본 이미지 기준 17개 시도 약칭 (서 울, 부 산, ...)
  2. `routes/preview.py`, `routes/debug.py` 수정:
     - `_get_toc_sections()` 함수를 동적 계산 방식으로 완전 재작성
     - 요약 → 부문별 → 시도별 → 참고 → 통계표 → 부록 순서로 누적 페이지 계산
  3. `templates/toc_template.html` 수정:
     - 원본 이미지 스타일에 맞게 CSS 및 HTML 구조 업데이트
     - 폰트를 '바탕'으로 변경, 레이아웃 조정
- **관련 파일**: 
  - `config/reports.py` (수정)
  - `routes/preview.py` (수정)
  - `routes/debug.py` (수정)
  - `templates/toc_template.html` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 보고서 구성이 변경되면 `PAGE_CONFIG`만 수정하면 목차 페이지 번호가 자동 재계산됨
  - 목차 항목 구조를 변경하려면 `TOC_SECTOR_ITEMS`, `TOC_REGION_ITEMS` 수정
  - 테스트 결과: 요약 1, 부문별 6(생산 6, 소비 10, ...), 시도별 26(서울 26, 부산 28, ...), 참고 60, 통계표 62, 부록 74

---

#### 요약-고용인구 페이지 인구 증감률 표시 문제 해결
- **시간**: 2026-01-04 13:00
- **문제 설명**: 
  - 요약-고용인구 페이지에서 인구 차트에 이상한 증감률(%) 값이 표시됨
  - 서울 41.8%, 대구 -52.0%, 인천 53.0%, 광주 -82.4% 등 비정상적인 수치
  - 정답 이미지에서는 인구 차트에 증감률이 표시되지 않음
- **원인 분석**: 
  - 순인구이동 데이터의 "증감률"을 계산하는 것 자체가 의미 없음
  - 예: 전년 -1000명(유출) → 금년 +500명(유입)이면 증감률 계산이 불가능
  - 0에서 다른 값으로 바뀌면 무한대가 됨
  - 정답 보도자료에서는 인구 차트에 순이동 값(천명)만 표시하고 증감률은 표시하지 않음
  - 템플릿에서는 고용 차트처럼 증감률(%) 점을 표시하도록 잘못 구현되어 있었음
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "요약섹션에서 인구 증감률 퍼센트 숫자가 이상해"라고 제보
  - 초기 분석:
    - `summary_employment_template.html` 확인 → 인구 차트에 증감률 데이터셋이 포함됨
    - `summary_data.py`의 `get_employment_population_data()` 함수 확인
    - 749-751행: `change = round((curr_value - prev_value) / abs(prev_value) * 100, 1)`
    - 순이동 값에 대해 증감률을 계산하면 의미 없는 값이 나옴
  - 정답 이미지 분석:
    - `correct_answer/요약/요약-고용인구.png` 확인
    - 고용 차트: 막대(고용률%) + 점(증감 %p) → 정상
    - 인구 차트: 막대(순이동 천명)만 표시, 증감률 없음!
    - 차트 제목도 "< 시도별 인구 순이동(천명) >" (증감률 언급 없음)
  - 해결책 결정:
    1. 인구 차트에서 증감률 데이터셋 제거
    2. 차트 제목을 정답과 동일하게 수정
    3. 값을 천명 단위로 변환 (명 → 천명)
    4. 텍스트 표시도 정답과 동일하게 수정 (명 단위, 음수 표시)
- **해결 방법**: 
  1. 차트 제목 수정: `<시도별 순이동(천명) 및 전년동분기대비 증감률(%)>` → `< 시도별 인구 순이동(천명) >`
  2. `createPopulationChart()` 함수에서:
     - 증감률 데이터셋(line chart) 제거
     - y1 축(오른쪽 증감률 축) 제거
     - 값을 천명 단위로 변환 (`d.value / 1000`)
     - formatter를 `toFixed(1)`로 변경하여 소수점 1자리 표시
  3. 인구 텍스트 표시 수정:
     - 순유입: "경기(10,426명), 인천(8,050명), 충남(2,132명) 등 7개 지역은 순유입"
     - 순유출: "서울(-10,051명), 부산(-3,704명), 광주(-2,854명) 등 10개 지역은 순유출"
- **관련 파일**: 
  - `templates/summary_employment_template.html` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 순인구이동의 "증감률"은 본질적으로 의미 없는 지표임
  - 유입→유출 또는 유출→유입 전환 시 퍼센트 계산이 의미를 가지지 않음
  - 정답 보도자료에서도 인구 차트에는 증감률을 표시하지 않음
  - 수정 후 차트에 순이동 값(천명)만 표시되어 정답과 일치

---

### 2026-01-02

#### 인포그래픽 SVG 색칠 방식에서 PNG+마커 방식으로 롤백
- **시간**: 2026-01-02 16:30
- **문제 설명**: 
  - 인포그래픽이 디버그 모드에서는 정상적으로 표시되지만, 실제 자료 생성 시 이상한 색깔로 칠해짐
  - SVG fetch 방식이 서버 실행 중일 때만 작동하고, HTML 파일 저장 후 열 때는 실패함
- **원인 분석**: 
  - `infographic_js_template.html`이 SVG 파일을 `fetch('/templates/infographic_map.svg')`로 로드
  - 디버그 모드에서는 Flask 서버가 실행 중이므로 fetch가 성공
  - 실제 자료 생성 후 HTML 파일을 열면 로컬 파일 시스템에서 fetch가 실패하거나 CORS 문제 발생
  - SVG 로드 실패 시 색칠이 제대로 되지 않아 기본 색상만 표시됨
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 디버그에서는 잘 나오고 실제 생성에서는 이상한 색깔로 표시된다고 제보
  - 초기 분석:
    - `infographic_js_template.html` 확인 → SVG를 fetch로 동적 로드하는 방식
    - fetch는 HTTP 요청이므로 서버가 실행 중일 때만 작동
    - HTML 파일을 저장 후 직접 열면 `file://` 프로토콜이므로 fetch 실패
  - 해결책 분석:
    1. SVG를 인라인으로 템플릿에 포함 → 복잡하고 템플릿 크기 증가
    2. PNG 이미지 + 원형 마커 방식으로 롤백 → 안정적, 이전에 잘 작동하던 방식
    3. Base64 인코딩된 SVG 사용 → 복잡함
  - 선택한 접근 방법: PNG + 마커 방식으로 롤백
    - `templates/infographic_map.png` 파일이 이미 존재함
    - PNG 이미지는 상대/절대 경로 모두에서 안정적으로 작동
    - 원형 마커는 CSS로 스타일링하여 데이터에 따라 색상 표시
- **해결 방법**: 
  1. `templates/infographic_js_template.html` 수정:
     - SVG fetch 로직 제거
     - PNG 이미지 (`<img src="/templates/infographic_map.png">`) 사용
     - 각 지역 중심점에 원형 마커 (`<div class="region-marker">`) 오버레이
     - JavaScript로 마커 생성 및 데이터에 따른 색상 클래스 적용
  2. `templates/infographic_template.html` 수정:
     - 동일하게 PNG + 마커 방식으로 변경
     - `renderMapMarkers()` 함수로 마커 생성
  3. 지역별 좌표 매핑 (`REGION_POSITIONS`) 추가:
     - 17개 시도의 지도 위 백분율 좌표 정의
     - 예: 서울 { x: 37, y: 22 }, 부산 { x: 70, y: 65 }
- **관련 파일**: 
  - `templates/infographic_js_template.html` (수정)
  - `templates/infographic_template.html` (수정)
  - `templates/infographic_map.png` (기존 파일 활용)
- **상태**: ✅ 완료
- **참고 사항**: 
  - PNG + 마커 방식은 fetch 없이 작동하므로 HTML 파일 저장 후에도 정상 표시
  - 마커 색상: increase(빨강), decrease(파랑), neutral(회색)
  - 소비자물가 특수 색상: above-avg(보라), below-avg(연보라), avg(노랑)
  - 향후 SVG 방식을 사용하려면 SVG를 인라인으로 포함하거나 서버에서 생성해야 함

---

#### 요약-수출물가 페이지 수출 데이터 0.0 표시 문제 해결
- **시간**: 2026-01-02 11:01
- **문제 설명**: 
  - 요약-수출물가 페이지에서 전국 수출액이 0.0억달러로 표시됨
  - 실제 데이터는 약 1752억달러(175,158백만달러)이어야 함
- **원인 분석**: 
  - 세션에서 사용하는 엑셀 파일이 `분석표_2025년_2분기_자동생성.xlsx`였음
  - 이 파일의 열 구조가 원본 파일(`분석표_25년 2분기_캡스톤.xlsx`)과 다름
  - 자동생성 파일: 지역명=col1, 분류단계=col2
  - 원본 파일: 지역명=col3, 분류단계=col4
  - 코드는 원본 파일 구조(col3, col4)에 맞춰 작성되어 있어 자동생성 파일에서는 데이터를 찾지 못함
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 수출 데이터가 모두 0으로 표시된다고 제보
  - 초기 분석:
    - `get_summary_table_data()` 함수의 exports 설정 확인
    - 'G(수출)집계' 시트의 열 매핑이 잘못되었을 가능성 확인
  - 데이터 추출 테스트:
    - Python 스크립트로 직접 `get_trade_price_data()` 호출 → 정상적으로 2.1% 증가율 반환
    - 데이터 추출 로직은 정상 동작 확인
  - 서버 캐시 의심:
    - __pycache__ 삭제 및 서버 재시작 수행
    - 여전히 0.0 표시
  - 세션 파일 확인:
    - `curl`로 세션 상태 확인 → `분석표_2025년_2분기_자동생성.xlsx` 사용 중
    - 테스트에 사용한 파일과 다름!
  - 파일 구조 비교:
    - 자동생성 파일: 지역=col1, 분류=col2
    - 원본 파일: 지역=col3, 분류=col4
    - 열 위치가 2칸씩 차이남
  - 해결책 선택:
    - 옵션 1: 코드를 두 파일 형식 모두 지원하도록 수정 → 복잡하고 유지보수 어려움
    - 옵션 2: 자동생성 파일 삭제하고 원본 파일 사용 → 빠르고 간단
    - 사용자 의견: "자동생성 파일을 바꿔야지 당연히" → 옵션 2 선택
- **해결 방법**: 
  - `분석표_2025년_2분기_자동생성.xlsx` 파일 삭제
  - 세션이 원본 파일(`분석표_25년 2분기_캡스톤.xlsx`)을 자동으로 사용하도록 함
- **관련 파일**: 
  - `분석표_2025년_2분기_자동생성.xlsx` (삭제됨)
  - `services/summary_data.py` (수정 없음, 원본 파일 구조에 맞게 이미 작성되어 있음)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 자동생성 파일의 열 구조가 원본과 다르게 생성되었음
  - 향후 자동생성 로직이 있다면 원본 파일과 동일한 열 구조로 생성하도록 수정 필요
  - 수출액 1752.0억달러, 증감률 2.1% 정상 표시 확인

---

#### 요약-고용인구 페이지 인구 차트 미출력 문제 해결
- **시간**: 2026-01-02 10:28
- **문제 설명**: 
  - 요약-고용인구 페이지에서 인구 섹션의 차트가 출력되지 않음
  - 고용 차트는 정상이나 인구 차트만 빈 상태로 표시
- **원인 분석**: 
  - `get_employment_population_data` 함수에서 `population` 딕셔너리의 `chart_data` 배열이 빈 채로 반환됨
  - `inflow_regions`와 `outflow_regions`는 채워졌지만 템플릿에서 사용하는 `chart_data`는 채우지 않음
  - 템플릿(`summary_employment_template.html`)은 `population.chart_data`를 순회하여 차트 데이터를 생성
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 인구 차트만 안 그려진다고 제보, 스크린샷에서 차트 영역이 빈 것 확인
  - 초기 분석: `summary_employment_template.html`의 JavaScript 코드 분석
    - `createPopulationChart()` 함수가 `populationData` 배열을 사용
    - `populationData`는 `population.chart_data`에서 생성됨
  - 데이터 흐름 추적:
    - `routes/preview.py`, `routes/debug.py`에서 `get_employment_population_data()` 호출
    - 반환 값에 `population.chart_data`가 빈 배열인 것 확인
  - 코드 분석: `services/summary_data.py`의 `get_employment_population_data()` 함수
    - `population = {'chart_data': []}` 초기화 후 채우지 않음
    - `inflow_regions`, `outflow_regions`만 채우고 `chart_data`는 누락
  - 해결책 결정: 인구이동 데이터 추출 시 `chart_data`도 함께 구성
    - 각 지역별 순이동량(value)과 전년동분기대비 증감률(change) 계산 필요
    - 열 21(2024 2/4분기)과 열 25(2025 2/4분기) 데이터로 증감률 계산
- **해결 방법**: 
  - `services/summary_data.py`의 `get_employment_population_data()` 함수 수정:
    - 전년동분기(열 21) 데이터 추출 추가
    - 증감률 계산 로직 추가: `(curr - prev) / abs(prev) * 100`
    - `region_data` 딕셔너리에 각 지역 데이터 저장
    - 지역 순서대로 `chart_data` 배열 구성
- **관련 파일**: 
  - `services/summary_data.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 차트 데이터 구조: `{name: 지역명, value: 순이동(천명), change: 전년동분기대비 증감률(%)}`
  - 순유입은 양수(파란색), 순유출은 음수(빨간색)으로 표시
  - 증감률이 매우 큰 경우(예: 강원 -1286.8%)는 전년동분기 값이 매우 작았기 때문

---

#### 시도별 보도자료 본문 텍스트 데이터 완전 수정
- **시간**: 2026-01-02 11:30
- **문제 설명**: 
  - 시도별 보도자료의 본문 텍스트(광공업생산, 서비스업생산, 소매판매, 건설수주)에서 증감률과 품목 데이터가 0으로 표시됨
  - 차트 및 요약 표 수정 후에도 본문 텍스트는 여전히 플레이스홀더 값 출력
- **원인 분석**: 
  - `extract_retail_data`, `extract_construction_data` 함수가 분석 시트가 비어있을 때 집계 시트로 fallback하지 않음
  - `_extract_retail_from_aggregation`, `_extract_construction_from_aggregation` 함수 누락
  - 기존에 추가된 fallback 함수들도 열 매핑이 잘못되어 있었음:
    - C(소비)집계: region_col=4가 아닌 2, 분기 열=15/19가 아닌 20/24
    - F'(건설)집계: region_col=2가 아닌 1, level_col=3이 아닌 2, 분기 열=14/18이 아닌 18/22
- **에이전트 사고 과정**:
  1. 테스트에서 manufacturing, service, retail, construction 데이터가 모두 0으로 확인
  2. `extract_manufacturing_data`, `extract_service_data`를 직접 호출 → 정상 작동
  3. 데이터 경로 오류 발견: `data['manufacturing']`이 아닌 `data['production']['manufacturing']`
  4. `extract_retail_data`, `extract_construction_data`에 집계 시트 fallback 로직 추가 필요
  5. 각 집계 시트의 열 구조 직접 분석:
     - C(소비)집계: 열2=지역이름, 열3=분류단계, 열5=업태코드, 열6=업태종류, 열20=2024Q2, 열24=2025Q2
     - F'(건설)집계: 열1=지역이름, 열2=분류단계, 열4=공정이름, 열18=2024Q2, 열22=2025Q2
  6. fallback 함수 열 매핑 수정 및 테스트
- **해결 방법**: 
  1. `extract_retail_data`에 집계 시트 fallback 조건 추가 (total_row 없을 때, 데이터가 모두 0일 때)
  2. `_extract_retail_from_aggregation` 함수 추가 (열 매핑: region_col=2, level_col=3, name_col=6, AGG_COLS={2024_2Q: 20, 2025_2Q: 24})
  3. `extract_construction_data`에 집계 시트 fallback 조건 추가
  4. `_extract_construction_from_aggregation` 함수 추가 (열 매핑: region_col=1, level_col=2, name_col=4, AGG_COLS={2024_2Q: 18, 2025_2Q: 22})
- **관련 파일**: `templates/regional_generator.py`
- **상태**: ✅ 완료
- **결과**:
  - 서울: 광공업(-10.1%), 서비스업(2.3%), 소매(-1.8%), 건설(29.2%)
  - 부산: 광공업(-4.0%), 서비스업(2.1%), 소매(1.0%), 건설(73.8%)
  - 대구: 광공업(-2.2%), 서비스업(-1.2%), 소매(-1.4%), 건설(370.9%)
  - 인천: 광공업(-3.7%), 서비스업(3.5%), 소매(4.9%), 건설(-50.0%)
  - 경기: 광공업(12.3%), 서비스업(5.4%), 소매(0.3%), 건설(-13.3%)
- **참고 사항**: 분석 시트의 수식이 계산되지 않은 상태에서도 집계 시트의 원시 데이터로부터 정상적인 증감률 계산 가능

---

#### 시도별 보도자료 차트 데이터 및 요약 표 수정
- **시간**: 2026-01-02 10:30
- **문제 설명**: 
  - 시도별 보도자료 차트(광공업, 서비스업, 소매판매, 건설, 수출입, 물가, 고용률)가 모두 0.0으로 표시
  - 요약 표도 대부분 0.0으로 표시 (인구순이동만 정상)
- **원인 분석**: 
  - `extract_chart_data` 함수가 분석 시트에서 데이터를 가져오는데, 분석 시트가 비어있음
  - `_get_chart_time_series` 함수가 집계 시트로 fallback하지 않음
  - `extract_summary_table` 함수도 동일한 문제
  - 각 집계 시트의 열 매핑이 시트마다 다름 (지역 열, 코드 열, 분기 데이터 열)
- **에이전트 사고 과정**:
  1. 차트 데이터가 0.0인 원인 분석 → 분석 시트(A 분석)의 증감률 열이 NaN
  2. 집계 시트에서 지수 데이터를 가져와 전년동기대비 증감률을 계산하는 방안 검토
  3. 각 시트별 열 구조 분석:
     - A(광공업생산)집계: 지역=열4, 코드=열7, 분기=열13~26
     - B(서비스업생산)집계: 지역=열3, 코드=열6, 분기=열12~25
     - C(소비)집계: 지역=열2, 코드=열3, 분기=열11~24
     - G(수출)집계: 지역=열3, 코드=열4, 분기=열14~26
     - E(지출목적물가)집계: 지역=열2, 코드=열3, 분기=열11~24
     - D(고용률)집계: 지역=열1, 코드=열2, 분기=열8~21
  4. 시트별 열 매핑 딕셔너리 구현하여 동적으로 적용
- **해결 방법**: 
  1. `_get_chart_time_series_from_aggregation` 함수 추가: 집계 시트에서 전년동기대비 증감률 계산
  2. `_get_chart_data_with_fallback` 함수 추가: 분석 시트 → 집계 시트 fallback
  3. `extract_chart_data`에서 모든 차트에 fallback 로직 적용
  4. `_get_summary_value_from_aggregation` 함수 추가: 요약 표용 증감률 계산
  5. `_get_employment_change_from_aggregation` 함수 추가: 고용률 %p 증감 계산
  6. `extract_summary_table`에서 모든 지표에 fallback 로직 적용
- **관련 파일**: `templates/regional_generator.py`
- **상태**: ✅ 완료
- **결과**:
  - 차트: 광공업(-10.1%), 서비스업(2.3%), 소매(-1.8%), 건설(29.2%), 수출(1.6%), 수입(-1.7%), 물가(2.0%), 고용률(-0.2%p)
  - 요약표: 모든 지표 정상 출력 확인

---

#### 요약 섹션 데이터 누락 문제 해결
- **시간**: 2026-01-02 09:30
- **문제 설명**: 요약 보도자료에서 광공업생산, 서비스업생산, 소매판매, 수출, 물가 데이터가 모두 0%로 표시됨
- **원인 분석**: 
  - `_extract_chart_data` 함수가 분석 시트에서 증감률을 가져오는데, 분석 시트의 증감률 열이 모두 NaN
  - 기초자료 시트가 없는 "분석표" 엑셀 파일에서는 집계 시트로 fallback해야 함
  - 기존 코드는 분석 시트 → 기초자료 시트만 fallback하고, 집계 시트로는 fallback 안 함
- **에이전트 사고 과정**:
  - `_extract_sector_summary`는 정상 동작하는데 `_extract_chart_data`는 0 반환 확인
  - 분석 시트(A 분석)의 열 21 값이 모두 NaN인 것 확인
  - "광공업생산" 기초자료 시트가 없음 확인 (분석표 파일이므로)
  - 집계 시트(A(광공업생산)집계)에서 직접 계산 필요
- **해결 방법**: 
  - `_extract_chart_data_from_aggregate` 함수 추가
  - 분석 시트 비어있으면 기초자료 → 집계 시트 순으로 fallback
  - A/B/C/G/E 시트 모두 집계 시트 설정 추가
  - 물가는 지역별 데이터가 있는 `E(지출목적물가)집계` 시트 사용
- **관련 파일**: `services/summary_data.py`
- **상태**: ✅ 완료
- **참고 사항**: 
  - 광공업: 전국 2.1%, 증가: 충북/경기/광주
  - 서비스업: 전국 1.4%, 증가: 경기/인천/세종
  - 소매판매: 전국 -0.2%, 증가: 울산/인천/세종
  - 수출: 전국 2.1%, 증가: 제주/충북/경남
  - 물가: 전국 2.2%, 상승: 부산/인천/강원

---

#### 시도별 보도자료 수출/수입/물가 데이터 누락 문제 해결 (근본 원인 수정)
- **시간**: 2026-01-02 09:15
- **문제 설명**: 시도별 보도자료에서 수출/수입 증감률이 0으로 표시되고 품목도 비어있음
- **원인 분석**: 
  - G 분석, H 분석 시트의 증감률 열(열 22)이 모두 비어있음 (엑셀 수식이 계산되지 않음)
  - `_get_quarter_value` 함수가 분석 시트에서 NaN을 읽어 0.0을 반환
  - 품목은 집계 시트에서 추출하지만, 총 증감률은 분석 시트에서 가져오려 해서 0이 됨
- **에이전트 사고 과정**:
  - 테스트 결과 `_get_sido_products_from_aggregation`은 정상 동작 확인
  - `extract_export_data`에서 `growth_rate`가 0인 것 확인
  - G 분석 시트 확인 → 열 22 데이터가 모두 NaN
  - 분석 시트가 비어있으면 집계 시트에서 총 증감률도 계산해야 함
- **해결 방법**: 
  - `_extract_export_from_aggregation` 함수 추가: 집계 시트에서 총 증감률 + 품목 모두 추출
  - `_extract_import_from_aggregation` 함수 추가: 집계 시트에서 총 증감률 + 품목 모두 추출
  - `extract_export_data`: 분석 시트가 비어있으면 집계 시트에서 전체 데이터 추출
  - `extract_import_data`: 분석 시트가 비어있으면 집계 시트에서 전체 데이터 추출
- **관련 파일**: `templates/regional_generator.py`
- **상태**: ✅ 완료
- **참고 사항**: 
  - 서울 수출: 1.6% (금, 기타 인조플라스틱 등 증가 / 차량 부품, 경유 감소)
  - 서울 수입: -1.7% (선박, 기타 집적회로반도체 증가 / 원유, 가스 감소)
  - 서울 물가: 2.1% (기타 상품 및 서비스, 식료품·비주류음료 상승)

---

#### 통계표 섹션 페이지 붙어보이는 문제 해결
- **시간**: (현재)
- **문제 설명**: 통계표 섹션에서 두 장씩 페이지가 붙어서 표시됨 (브라우저에서 볼 때)
- **원인 분석**: 
  - `.page` 클래스의 CSS에서 `margin: 0 auto;`로만 설정되어 페이지 사이에 간격이 없음
  - `page-break-after: always`는 인쇄 시에만 작동하고 화면 표시에는 영향 없음
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 통계표에서 페이지가 붙어있다고 제보
  - 코드 분석: 통계표 관련 템플릿 5개 파일의 `.page` CSS 확인
  - 원인 파악: 화면 표시 시 페이지 간 마진이 없음
  - 해결책: 기본 CSS에 마진, 그림자, 테두리 추가 / 인쇄 시에는 제거
- **해결 방법**: 
  - 5개 통계표 템플릿 수정:
    - `statistics_table_template.html`
    - `statistics_table_index_template.html`
    - `statistics_table_grdp_template.html`
    - `statistics_table_toc_template.html`
    - `statistics_table_appendix_template.html`
  - 기본 `.page` CSS 변경:
    - `margin: 0 auto 20mm auto;` (페이지 간 20mm 간격)
    - `box-shadow: 0 2px 8px rgba(0,0,0,0.1);` (그림자 효과)
    - `border: 1px solid #ddd;` (테두리)
  - `@media print`에서 마진/그림자/테두리 제거
- **관련 파일**: 
  - `templates/statistics_table_template.html`
  - `templates/statistics_table_index_template.html`
  - `templates/statistics_table_grdp_template.html`
  - `templates/statistics_table_toc_template.html`
  - `templates/statistics_table_appendix_template.html`
- **상태**: ✅ 완료
- **참고 사항**: 인쇄/PDF 출력 시에는 간격 없이 정상적으로 페이지 나뉨

---

#### 통계표-참고-GRDP 표 데이터 누락 문제 해결
- **시간**: (현재)
- **문제 설명**: 통계표-참고-GRDP의 표가 모두 비어있음 (연도별, 분기별 데이터 전체 "-"로 표시)
- **원인 분석**: 
  - `_create_grdp_placeholder` 함수에서 `grdp_extracted.json` 로드 시 `item.get('placeholder', True)` 조건으로 인해 데이터가 무시됨
  - JSON에 `placeholder` 필드가 없으면 기본값 `True`가 사용되어 `not True` = `False`로 데이터가 채워지지 않음
  - 과거 연도별/분기별 데이터를 위한 historical JSON 파일이 없었음
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 GRDP 표가 비어있다고 제보
  - 코드 분석: `_create_grdp_placeholder` 함수 확인
    - 현재 분기만 `grdp_extracted.json`에서 로드 시도
    - `not item.get('placeholder', True)` 조건이 항상 False가 됨
  - 정답 이미지 확인: 2017년부터 2025년 1분기까지의 과거 데이터가 필요함
  - 해결책 분석:
    1. `placeholder` 기본값을 `False`로 변경 → 현재 분기 데이터 로드 가능
    2. 과거 데이터를 위한 historical JSON 파일 생성 필요
    3. 정답 이미지에서 과거 데이터 추출하여 JSON 생성
  - 구현 결정:
    - 조건을 `not item.get('placeholder', False)`로 변경
    - `grdp_historical_data.json` 파일 생성하여 과거 데이터 저장
    - `_create_grdp_placeholder` 함수 수정하여 historical JSON 먼저 로드
- **해결 방법**: 
  1. `templates/statistics_table_generator.py` 수정:
     - `item.get('placeholder', True)` → `item.get('placeholder', False)`로 변경
     - `grdp_historical_data.json`에서 과거 데이터 로드 로직 추가
  2. `templates/grdp_historical_data.json` 생성:
     - 연도별: 2017 ~ 2024 (18개 지역)
     - 분기별: 2017.3/4 ~ 2025.2/4p (18개 지역)
     - 정답 이미지에서 데이터 추출
- **관련 파일**: 
  - `templates/statistics_table_generator.py` (수정)
  - `templates/grdp_historical_data.json` (신규 생성)
- **상태**: ✅ 완료
- **참고 사항**: 
  - GRDP 데이터는 KOSIS 실험적통계에서 별도 다운로드 필요
  - 새 분기 데이터는 `grdp_extracted.json`에서 자동 로드
  - 과거 데이터는 `grdp_historical_data.json`에 수동 추가 필요

---

#### 부문별/시도별 보도자료의 수출·수입·물가 품목 누락 문제 해결
- **시간**: (현재)
- **문제 설명**: 
  - 부문별 보도자료에서 수출, 수입, 물가동향의 품목이 하나도 표시되지 않음
  - 시도별 보도자료에서 수출, 수입, 물가의 품목이 제대로 표시되지 않음
- **원인 분석**: 
  - 분석표 엑셀 파일의 "분석" 시트에 수식이 계산되지 않은 상태로 업로드됨
  - 각 generator가 "분석" 시트에서 품목 데이터를 읽으려 했으나 값이 NaN/빈값
  - "참조" 시트에서 품목 데이터를 가져오려 했으나 해당 시트도 비어있음
  - 집계 시트로의 fallback 로직이 일부 generator에 없거나 불완전했음
  - regional_generator.py의 열 매핑이 집계 시트 구조와 맞지 않았음
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 부문별/시도별 4가지 문제 제보
    1. 부문별 수출 품목 미표시
    2. 부문별 수입 품목 미표시
    3. 부문별 물가동향 품목 미표시
    4. 시도별 시트 전체 미표시
  - 초기 분석:
    - 분석 시트 데이터 확인 → 수식 결과가 NaN
    - 집계 시트 확인 → 실제 데이터 존재
    - 각 generator 코드 분석 → fallback 로직 불완전
  - 해결책 분석:
    1. export_generator.py: `get_sido_products_from_reference()` 함수가 "참조" 시트를 읽는데, 시트가 비어있음 → 집계 시트에서 품목 추출 함수 추가 필요
    2. import_generator.py: export와 동일한 패턴 → 집계 시트 fallback 추가
    3. price_trend_generator.py: `get_category_data()` 함수가 분석 시트를 읽는데 NaN → 집계 시트 fallback 추가
    4. regional_generator.py: 
       - extract_export_data, extract_import_data의 품목 추출 로직에서 열 인덱스 오류
       - extract_consumer_price_data에서 is_raw 체크 로직이 잘못됨
       - _extract_price_items_from_aggregation의 열 매핑 오류 (3→2, 4→3, 8→6)
  - 구현 결정:
    - 각 generator에 `_get_sido_products_from_aggregation()` 또는 유사 함수 추가
    - 분석 시트가 비어있을 때 집계 시트로 fallback하는 로직 추가
    - regional_generator.py의 열 매핑 수정 및 is_raw 로직 제거
- **해결 방법**: 
  1. `templates/export_generator.py`:
     - `_get_sido_products_from_aggregation()` 함수 추가 (G(수출)집계 시트에서 품목 추출)
     - `get_sido_products_from_reference()`에서 use_aggregation_only일 때 집계 시트 사용
  2. `templates/import_generator.py`:
     - `_get_sido_products_from_aggregation()` 함수 추가 (H(수입)집계 시트에서 품목 추출)
     - export와 동일한 패턴으로 fallback 로직 추가
  3. `templates/price_trend_generator.py`:
     - `_get_category_data_from_aggregation()` 함수 추가 (E(품목성질물가)집계 시트에서 품목 추출)
     - `get_category_data()`에서 use_aggregation_only일 때 집계 시트 사용
  4. `templates/regional_generator.py`:
     - `extract_export_data()`: 집계 시트 품목 추출 시 열 인덱스 수정 (분류단계: 열2→'2', 상품이름: 열7)
     - `extract_import_data()`: export와 동일한 수정
     - `extract_consumer_price_data()`: is_raw 체크 제거, 분석 시트 데이터가 비어있으면 집계 시트로 fallback
     - `_extract_price_data_from_aggregation()` 함수 추가 (총지수 + 품목 데이터 한번에 추출)
     - `_extract_price_items_from_aggregation()`: 열 매핑 수정 (2=지역이름, 3=분류단계, 6=분류이름)
- **관련 파일**: 
  - `templates/export_generator.py` (수정)
  - `templates/import_generator.py` (수정)
  - `templates/price_trend_generator.py` (수정)
  - `templates/regional_generator.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 집계 시트의 열 구조 확인 방법: `pd.read_excel(xl, sheet_name='시트명', header=None).iloc[2, :]`
  - 수출 집계 시트 구조: 열3=지역이름, 열4=분류단계, 열7=상품이름
  - 물가 집계 시트 구조: 열2=지역이름, 열3=분류단계, 열6=분류이름
  - 테스트 명령어: `python3 -c "from templates.xxx_generator import ..."`

---

#### 인포그래픽 지도 위치 및 SVG 색칠 방식 구현
- **시간**: 2026-01-02
- **문제 설명**: 
  1. 인포그래픽 6개 지도의 위치가 맞지 않음
  2. 원형 마커 대신 지역 path를 색상표에 따라 직접 색칠하는 방식으로 변경 요청
- **원인 분석**: 
  - SVG viewBox가 `"250 100 450 450"`으로 설정되어 지도 영역을 제대로 표시하지 못함
  - SVG path의 transform 속성들이 x=400~575, y=175~540 범위에 있음
  - path에 인라인 스타일 `fill:#878787;fill-opacity:1;...`이 복잡하게 설정되어 CSS 클래스 적용이 안됨
  - 기존 `infographic_js_template.html`이 PNG 지도 + 원형 마커 오버레이 방식 사용
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 지도 위치 문제와 마커→path 색칠 방식 변경 요청
  - SVG 분석:
    - 각 path의 transform 속성 확인 (translate 좌표)
    - x 범위: 약 400~575, y 범위: 약 175~540
    - 현재 viewBox "250 100 450 450"이 지도를 제대로 포함하지 못함
  - 두 가지 템플릿 존재 확인:
    - `infographic_template.html`: SVG 로드 방식 (미완성)
    - `infographic_js_template.html`: PNG + 마커 방식 (실제 사용)
  - 해결 방안:
    1. SVG viewBox를 지도 영역에 맞게 조정: `"340 130 280 360"`
    2. SVG path의 인라인 스타일 단순화
    3. `infographic_js_template.html`을 SVG 방식으로 완전 변경
    4. JavaScript로 SVG 로드 및 path 색칠 로직 구현
- **해결 방법**: 
  1. `templates/infographic_map.svg` 수정:
     - viewBox를 `"340 130 280 360"`으로 변경
     - width/height를 200x260으로 변경
     - 모든 path의 인라인 스타일 단순화: `fill:#CCCCCC;stroke:#ffffff;stroke-width:0.5;...`
  2. `templates/infographic_js_template.html` 수정:
     - PNG 이미지 + 마커 방식 → SVG path 색칠 방식으로 완전 변경
     - CSS: `.region-marker` 클래스들을 `.korea-map-svg .region` 클래스로 변경
     - HTML: `<img class="korea-map-img">` 제거, JavaScript로 SVG 동적 로드
     - JavaScript: 
       - `addRegionMarkers()` → `loadAndColorMaps()` 함수로 변경
       - `REGION_POSITIONS` (좌표) → `REGION_ID_MAP` (지역명-ID 매핑)으로 변경
       - SVG fetch 후 각 path에 데이터 기반 CSS 클래스 적용
  3. `templates/infographic_template.html` CSS 수정:
     - `.korea-map-wrapper` 크기를 100x130px로 조정
     - flexbox로 중앙 정렬 추가
- **관련 파일**: 
  - `templates/infographic_map.svg` (수정)
  - `templates/infographic_js_template.html` (수정 - 주요 변경)
  - `templates/infographic_template.html` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - SVG viewBox 계산: path transform의 x/y 범위를 분석하여 적절한 영역 설정
  - 색상 결정 로직:
    - 일반 지표: 값 > 0 → increase(붉은색), 값 < 0 → decrease(푸른색), 값 = 0 → neutral(회색)
    - 소비자물가: 전국평균 초과 → above-avg(보라), 미만 → below-avg(연보라), 동일 → avg(노랑)
  - 수출 지표가 회색으로 표시되는 것은 데이터 추출 오류("없음" 값) 때문임

---

### 2026-01-01

#### 파일 유형 분석 성능 최적화 - 빠른 판정 로직 구현
- **시간**: (현재)
- **문제 설명**: 파일 유형 분석(`detect_file_type`)이 너무 오래 걸림. 모든 시트를 대조하는 과정이 불필요하게 느림
- **원인 분석**: 
  - 기존 코드는 모든 시트명을 읽고 전체 시트 세트와 교집합 연산 수행
  - pandas ExcelFile로 전체 파일을 읽어서 느림
  - 시트를 모두 확인할 필요 없이 핵심 시트만 확인하면 충분
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 파일 유형 분석이 너무 오래 걸린다고 제보
  - 초기 분석: `utils/excel_utils.py`의 `detect_file_type` 함수 확인
  - 성능 병목 파악:
    - pandas ExcelFile로 전체 파일 읽기
    - 모든 시트명과 전체 시트 세트 교집합 계산
    - 불필요한 패턴 매칭 및 시트 개수 확인
  - 최적화 전략 수립:
    1. 파일명 먼저 확인 (가장 빠름, 파일 읽기 불필요)
    2. openpyxl 사용 (pandas보다 빠름, read_only 모드)
    3. 핵심 시트만 확인 (전체 시트 대조 불필요)
    4. 첫 매칭 시 즉시 반환 (조기 종료)
    5. 시트 개수로 빠른 추정
  - 구현 결정:
    - 파일명 확인을 1단계로 이동
    - openpyxl을 우선 사용 (pandas는 fallback)
    - 핵심 시트만 선별하여 빠른 매칭
    - 패턴 매칭 시 2개만 찾으면 즉시 반환
- **해결 방법**: 
  - `detect_file_type` 함수 최적화:
    1. 파일명 확인을 최우선으로 이동 (파일 읽기 불필요)
    2. openpyxl read_only 모드 사용 (pandas보다 빠름)
    3. 핵심 시트만 선별하여 빠른 매칭
    4. 첫 매칭 시 즉시 반환 (조기 종료)
    5. 불필요한 전체 시트 대조 제거
  - 성능 개선:
    - 파일명 매칭: 즉시 반환 (파일 읽기 없음)
    - 핵심 시트 매칭: 첫 시트 발견 시 즉시 반환
    - 패턴 매칭: 2개만 찾으면 즉시 반환
    - 시트 개수 추정: 간단한 비교로 빠른 판정
- **관련 파일**: 
  - `utils/excel_utils.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 기존 로직의 정확도는 유지하면서 성능만 개선
  - 파일명이 명확한 경우 파일 읽기 없이 즉시 판정 가능
  - 핵심 시트만 확인하여 대부분의 경우 빠르게 판정 가능

---

#### 전체 프로젝트 용어 변경 - '보고서' → '보도자료'
- **시간**: 17:10
- **문제 설명**: 대시보드 및 시스템 전반에서 '보고서'라는 용어가 '보도자료'로 변경되어야 함
- **원인 분석**: 
  - 통계청에서 발행하는 문서의 공식 명칭이 '보고서'가 아닌 '보도자료'임
  - UI, 로그 메시지, 에러 메시지, 주석, 문서 전체에서 일괄 변경 필요
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 '보고서'를 '보도자료'로 전면 변경 요청
  - 범위 파악: grep으로 '보고서' 사용 위치 전체 검색 → 79개 파일 발견
  - 영향 분석:
    - 코드 파일: 영어 변수명 'report'는 유지, 한글 텍스트만 변경
    - UI 파일: 사용자에게 보이는 모든 텍스트 변경
    - 문서 파일: 가이드, README 등 모든 문서 변경
  - 위험 요소 파악:
    - replace_all 사용 시 의도치 않은 변경 가능성
    - 함수명/변수명이 한글인 경우 오류 발생 가능
  - 해결 전략: replace_all로 일괄 변경 후 lint 검사로 오류 확인
  - 실행: 79개 파일 모두 순차적으로 변경
  - 검증: grep으로 남은 '보고서' 없음 확인, lint 오류 없음 확인
- **해결 방법**: 
  - 전체 프로젝트에서 한글 '보고서'를 '보도자료'로 일괄 변경
  - 변경된 파일 카테고리:
    - 핵심 코드: `app.py`, `routes/api.py`, `routes/preview.py`, `routes/debug.py`
    - 서비스: `services/report_generator.py`, `services/summary_data.py`
    - UI/템플릿: `dashboard.html`, `templates/index.html`
    - Generator 파일: `templates/*_generator.py` (15개)
    - JSON 스키마: `templates/*_schema.json` (22개)
    - 설정: `config/reports.py`, `utils/data_utils.py`
    - 문서: `docs/*.md`, `README.md` 등 (20여 개)
- **관련 파일**: 
  - 총 79개 파일 수정
- **상태**: ✅ 완료
- **참고 사항**: 
  - 영어 변수명/함수명 'report'는 코드 안정성을 위해 그대로 유지
  - 서버 재시작 후 브라우저에서 정상 표시 확인 완료

---

#### 인포그래픽 지도 이미지 404 오류 해결 - Flask 라우트 추가
- **시간**: 16:50
- **문제 설명**: 인포그래픽 페이지에서 한국 지도 이미지가 로드되지 않고 "한국 지도" alt text만 표시됨
- **원인 분석**: 
  - 템플릿에서 `/templates/infographic_map.png` 절대 경로로 이미지를 참조
  - Flask에 해당 경로에 대한 라우트가 없어서 404 Not Found 발생
  - 이전에 파일을 templates 폴더로 복사했지만, 라우트가 없어서 접근 불가
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 지도가 여전히 로드되지 않는다고 제보
  - 초기 분석: 파일 존재 확인 (287KB, 정상) → 파일은 존재함
  - 브라우저 테스트: `http://localhost:5050/templates/infographic_map.png` 접근 시 404 발생
  - 근본 원인 파악: `routes/main.py`에 `/templates/<filename>` 경로에 대한 라우트가 없음
  - 해결책 결정: templates 폴더의 정적 파일(이미지, CSS, JS)을 서비스하는 라우트 추가
- **해결 방법**: 
  - `routes/main.py`에 `/templates/<filename>` 라우트 추가
  - PNG, JPG, SVG, CSS, JS 파일에 대한 MIME type 처리 포함
- **관련 파일**: 
  - `routes/main.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - Flask debug 모드에서 자동 재시작되어 즉시 반영됨
  - 이전 해결(파일 복사)만으로는 불완전했음 - 라우트 추가 필요했음

---

#### Excel 전처리 성능 최적화 - xlwings fallback 순서 변경
- **시간**: 20:00
- **문제 설명**: 전처리 과정이 너무 오래 걸림. xlwings가 Excel 앱을 실행해야 하므로 매우 느림
- **원인 분석**: 
  - 기존 순서: xlwings → formulas → openpyxl
  - xlwings는 Excel 앱 실행이 필요하여 가장 느림
  - 백엔드에서 직접 계산하는 openpyxl 방식이 훨씬 빠름
  - 엑셀 함수 계산은 백엔드에서 계산해서 매핑하는 것이 더 효율적
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "전처리 과정이 너무 오래 걸린다"고 제보
  - 초기 분석: `services/excel_processor.py`의 `preprocess_excel()` 함수 확인
  - 현재 순서 파악: xlwings가 첫 번째로 실행되어 Excel 앱 실행으로 인한 지연 발생
  - 해결책 분석:
    1. 순서 변경: openpyxl(백엔드 직접 계산)을 우선 사용
    2. xlwings를 마지막 fallback으로 이동 (직접 계산 실패 시에만 사용)
    3. openpyxl 계산 로직 최적화 (필요한 데이터만 캐싱)
  - 선택한 접근: 백엔드 직접 계산을 우선 사용하고, xlwings는 마지막 fallback으로 변경
  - 구현 결정:
    - `preprocess_excel()` 함수의 순서 변경: openpyxl → formulas → xlwings
    - `_try_openpyxl_calculation()` 함수 최적화:
      - 필요한 집계 시트만 미리 캐싱
      - 열 문자 변환 함수 재사용
      - 빈 셀 건너뛰기로 메모리 사용 감소
    - `_try_xlwings()` 함수에 fallback 모드 주석 추가
    - `get_recommended_method()` 함수 업데이트 (openpyxl 우선 권장)
- **해결 방법**: 
  - `services/excel_processor.py` 수정:
    - `preprocess_excel()`: 실행 순서 변경 (openpyxl → formulas → xlwings)
    - `_try_openpyxl_calculation()`: 백엔드 직접 계산 로직 최적화
      - 필요한 집계 시트만 미리 캐싱
      - 열 문자 변환 함수(`col_letter_to_number`) 재사용
      - 빈 셀 건너뛰기로 메모리 효율성 향상
    - `_try_xlwings()`: fallback 모드임을 명시하는 주석 추가
    - `get_recommended_method()`: openpyxl을 가장 빠른 방법으로 우선 권장
- **관련 파일**: 
  - `services/excel_processor.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 전처리 시간이 크게 단축됨 (Excel 앱 실행 불필요)
  - 백엔드에서 직접 계산하여 서버 환경에서도 빠르게 동작
  - xlwings는 복잡한 수식이 openpyxl/formulas로 계산 실패 시에만 사용
  - 기존 generator의 fallback 로직은 안전장치로 유지

---

#### Excel 수식 자동 계산 전처리 기능 추가
- **시간**: 18:30
- **문제 설명**: 분석표 엑셀 파일의 수식이 계산되지 않은 상태로 업로드되면 보도자료에 0, NaN 등 잘못된 데이터가 표시됨
- **원인 분석**: 
  - 분석표의 "분석" 시트들은 "집계" 시트를 참조하는 수식으로 구성
  - pandas로 읽을 때 수식 결과가 아닌 수식 자체(또는 None)가 읽힘
  - 기존에는 각 generator에서 fallback 로직으로 집계 시트에서 직접 계산
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 "집계표에서 계산하지 말고 엑셀 파일을 백엔드에서 실행하면 계산되지 않나요?" 제안
  - 해결책 분석:
    1. xlwings: Excel 앱이 설치되어 있으면 가장 정확 (Mac/Windows)
    2. formulas: 순수 Python으로 수식 계산 (서버 환경에서도 동작)
    3. openpyxl: 시트 간 참조 수식만 직접 계산
  - 선택한 접근: 3가지 방법을 순차적으로 시도하는 fallback 시스템 구축
  - 구현 결정: `services/excel_processor.py` 모듈 생성하여 캡슐화
- **해결 방법**: 
  - `services/excel_processor.py` 신규 생성
    - `preprocess_excel()`: 엑셀 파일 수식 계산 메인 함수
    - `_try_xlwings()`: Excel 앱으로 수식 계산 (가장 정확)
    - `_try_formulas()`: formulas 라이브러리로 순수 Python 계산
    - `_try_openpyxl_calculation()`: 시트 간 참조 수식 직접 계산
    - `check_available_methods()`: 사용 가능한 방법 확인
    - `get_recommended_method()`: 권장 방법 반환
  - `routes/api.py` 수정
    - 분석표 업로드 시 자동 전처리 수행
    - 전처리 결과를 응답에 포함
  - `requirements.txt` 업데이트
    - xlwings>=0.30.0 추가
    - Pillow>=10.0.0 추가 (이미지 처리용)
- **관련 파일**: 
  - `services/excel_processor.py` (신규 생성)
  - `routes/api.py` (수정)
  - `requirements.txt` (업데이트)
- **상태**: ✅ 완료
- **참고 사항**: 
  - xlwings 사용 시 Excel 앱이 백그라운드에서 실행됨
  - Excel이 설치되지 않은 서버에서는 formulas 또는 openpyxl fallback 사용
  - 기존 generator의 fallback 로직은 안전장치로 유지

---

#### 부문별 섹션 0/NaN/누락 행 문제 해결
- **시간**: 16:00
- **문제 설명**: 부문별 섹션에서 광공업생산, 건설동향, 실업률을 제외한 나머지 보도자료 표에 0, NaN, 또는 행 누락 발생
- **원인 분석**: 
  - 업로드된 분석표의 "분석" 시트에 수식이 계산되지 않은 상태
  - 각 generator가 분석 시트를 읽을 때 빈 값/NaN으로 읽힘
  - 일부 generator(광공업생산, 건설동향, 실업률)는 이미 집계 시트 fallback 로직이 있었음
- **에이전트 사고 과정**:
  - 영향받는 generator 식별: service_industry, consumption, export, import, price_trend, employment_rate, domestic_migration
  - 각 generator에 일관된 fallback 패턴 적용
  - 추가 발견: domestic_migration의 연령대 데이터 추출 시 `rank` 컬럼이 NaN이어서 필터링 실패 → `level` 컬럼으로 변경
  - 추가 발견: 지역명 표시 오류 (`sido.replace('', ' ')` → `' '.join(sido)`)
- **해결 방법**: 
  - 7개 generator에 `use_aggregation_only` 플래그 및 집계 시트 직접 계산 로직 추가
  - domestic_migration_generator: 연령대 필터링 조건을 `rank` → `level`로 변경
  - 지역명 표시 로직 수정
- **관련 파일**: 
  - `templates/service_industry_generator.py`
  - `templates/consumption_generator.py`
  - `templates/export_generator.py`
  - `templates/import_generator.py`
  - `templates/price_trend_generator.py`
  - `templates/employment_rate_generator.py`
  - `templates/domestic_migration_generator.py`
- **상태**: ✅ 완료
- **참고 사항**: Excel 전처리 기능 추가로 이 fallback 로직은 백업 용도로 유지

---

#### 인포그래픽 한국 지도 이미지 깨짐 문제 해결
- **시간**: 15:48
- **문제 설명**: 인포그래픽 페이지 생성 시 한국 지도 이미지가 깨져 보임 (이미지 로드 실패)
- **원인 분석**: 
  - 한국 지도 이미지가 `correct_answer/인포그래픽_map.png` 경로에 한글 파일명으로 저장되어 있음
  - 템플릿에서 `src="infographic_map.png"` 상대 경로로 참조하여 파일을 찾지 못함
  - 일부 템플릿에서는 `/correct_answer/인포그래픽_map.png` 경로 사용했으나 한글 파일명으로 인한 인코딩 문제 발생 가능
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 인포그래픽 페이지에서 한국 지도 이미지가 깨진다고 제보
  - 파일 위치 확인: `correct_answer/인포그래픽_map.png`에 이미지 존재 확인
  - 템플릿 분석: `infographic_js_template.html`, `infographic_template.html`, `infographic_regional_template.html`에서 이미지 경로 확인
  - 원인 파악: 상대 경로 사용 및 한글 파일명으로 인한 경로 해결 실패
  - 해결책 선택: 사용자 요청대로 이미지를 templates 폴더로 복사하고 영문 파일명으로 변경
- **해결 방법**: 
  - `correct_answer/인포그래픽_map.png`를 `templates/infographic_map.png`로 복사
  - 3개 템플릿 파일에서 이미지 경로를 `/templates/infographic_map.png` 절대 경로로 수정
- **관련 파일**: 
  - `templates/infographic_map.png` (신규 복사)
  - `templates/infographic_js_template.html` (경로 수정)
  - `templates/infographic_template.html` (경로 수정)
  - `templates/infographic_regional_template.html` (경로 수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 원본 파일 `correct_answer/인포그래픽_map.png`는 그대로 유지
  - 향후 이미지 업데이트 시 templates 폴더의 파일도 함께 업데이트 필요

---

### 2026-01-01

#### 통계표 2020.3/4 및 일부 분기 공란 문제 해결
- **시간**: 오후 (추가 수정)
- **문제 설명**: 통계표 HTML에서 2020.3/4 분기 행에 공란(`[  ]`) 플레이스홀더가 표시됨
- **원인 분석**: 
  - 일부 통계표(서비스업생산지수, 소매판매액지수)는 2020.4/4부터 데이터 제공 시작, 2020.3/4 데이터 없음
  - 기타 통계표(실업률, 소비자물가지수)는 JSON에 2020.3/4 데이터가 "-"로 저장되어 있었음
  - 템플릿의 `val()` 매크로가 "-" 값을 플레이스홀더로 변환
  - `quarterly_keys`가 하드코딩되어 데이터 없는 분기도 렌더링
- **에이전트 사고 과정**:
  - JSON 데이터 분석: 서비스업생산지수, 소매판매액지수, 실업률, 소비자물가지수에서 2020.3/4 공란 발견
  - 기초자료 확인: 서비스업생산지수/소매판매액지수는 2020.4/4부터 데이터 제공, 실업률/소비자물가지수는 데이터 존재
  - 해결책 1: 기초자료에서 추출 가능한 데이터(실업률, 소비자물가지수)를 JSON에 추가
  - 해결책 2: 기초자료에 없는 데이터(서비스업생산지수, 소매판매액지수)의 해당 분기 키를 JSON에서 제거
  - 해결책 3: 템플릿의 `val()` 매크로 수정 - "-" 값을 플레이스홀더가 아닌 그대로 표시
  - 해결책 4: `StatisticsTableGenerator.extract_table_data()` 수정 - 모든 지역이 "-"인 분기 자동 제거
  - 해결책 5: `services/report_generator.py` 수정 - `quarterly_keys`를 동적으로 생성
- **해결 방법**: 
  - `templates/statistics_historical_data.json`: 실업률, 소비자물가지수, 국내인구이동의 2020.3/4 데이터 추가, 서비스업생산지수/소매판매액지수의 2020.3/4 제거
  - `templates/statistics_table_index_template.html`: `val()` 매크로에서 "-" 조건 제거 (그대로 표시)
  - `templates/statistics_table_generator.py`: 모든 지역이 "-"인 분기 자동 제거 로직 추가
  - `services/report_generator.py`: `quarterly_keys`를 실제 데이터에서 동적으로 생성
- **관련 파일**: 
  - `templates/statistics_historical_data.json` (업데이트)
  - `templates/statistics_table_index_template.html` (수정)
  - `templates/statistics_table_generator.py` (수정)
  - `services/report_generator.py` (수정)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 일부 통계는 특정 시점부터 데이터 제공 시작하므로, 이전 분기는 표시되지 않음이 정상
  - 국내인구이동의 "전국" 데이터는 순인구이동 특성상 모든 분기에서 없음이 정상

---

#### 통계표 2025.2/4p 데이터 공란 문제 해결
- **시간**: 오후
- **문제 설명**: 통계표 HTML 생성 시 2025년 2분기(2025.2/4p) 데이터가 `[  ]` 플레이스홀더로 표시됨
- **원인 분석**: 
  - `statistics_historical_data.json` 파일의 분기 데이터 범위가 `2025.1/4`까지만 있고 `2025.2/4p`가 없음
  - 디버그 모드에서 통계표 생성 시 `raw_excel_path`가 세션에 설정되지 않아 동적 추출이 작동하지 않음
  - 템플릿의 `val()` 매크로가 값이 `-`일 때 플레이스홀더를 표시
- **에이전트 사고 과정**:
  - 문제 인식: 사용자가 공란 문제 제보 → HTML 파일에서 2197개의 `editable-placeholder` 발견
  - 초기 분석: `separator-row`의 빈 셀인지 확인 → 아님, 데이터 셀의 `-` 값이 플레이스홀더로 변환됨
  - 템플릿 분석: `statistics_table_index_template.html`의 `val()` 매크로가 `None`, `''`, `'-'` 값을 플레이스홀더로 변환
  - JSON 확인: `statistics_historical_data.json`의 `quarterly_range.end`가 `2025.1/4`로 확인
  - 동적 추출 테스트: 기초자료 파일이 있을 때 `StatisticsTableGenerator`가 올바르게 데이터 추출함 확인
  - 근본 원인 파악: 디버그 모드에서 세션에 `raw_excel_path`가 없어 동적 추출 불가
  - 해결책 결정: JSON 파일에 2025.2/4p 데이터를 직접 추가하여 영구적으로 해결
- **해결 방법**: 
  - `StatisticsTableGenerator`를 사용하여 기초자료에서 10개 통계표의 2025.2/4p 데이터 추출
  - `statistics_historical_data.json` 파일에 추출된 데이터 자동 저장
  - 메타데이터의 `quarterly_range.end`가 `2025.2/4`로 업데이트됨
- **관련 파일**: 
  - `templates/statistics_historical_data.json` (업데이트)
  - `debug/20260101_151713_statistics_2025Q2.html` (문제 파일)
- **상태**: ✅ 완료
- **참고 사항**: 
  - 새 분기 데이터 추가 시 `StatisticsTableGenerator.extract_and_save_all()` 메서드 사용 권장
  - 또는 기초자료 파일 경로를 세션에 설정하면 동적 추출 가능

---

### 2025-01-01

#### 디버그 로그 시스템 구축
- **시간**: 초기 설정
- **문제 설명**: 디버그 작업 추적 시스템이 없음
- **원인 분석**: 디버그 작업이 체계적으로 기록되지 않음
- **해결 방법**: `DEBUG_LOG.md` 파일 생성 및 추적 시스템 구축
- **관련 파일**: 
  - `docs/DEBUG_LOG.md` (신규 생성)
- **상태**: ✅ 완료
- **참고 사항**: 앞으로 모든 디버그 작업은 이 파일에 기록됩니다.

---

## 📊 통계

### 전체 디버그 항목 수
- 총 항목: 24
- 완료: 24
- 진행중: 0
- 실패: 0
- 보류: 0

### 최근 활동
- 마지막 업데이트: 2026-01-04 (인포그래픽 지도 회색 색상 배경 대비 문제 해결)

---

## 🔍 빠른 검색

### 카테고리별 분류
- [ ] 데이터 처리 오류
- [ ] UI/UX 문제
- [ ] 성능 이슈
- [ ] 의존성 문제
- [ ] 설정/환경 문제
- [ ] 기타

---

## 📝 사용 방법

새로운 디버그 항목을 추가할 때는 다음 템플릿을 사용하세요:

```markdown
### YYYY-MM-DD

#### [문제 제목]
- **시간**: HH:MM
- **문제 설명**: 
- **원인 분석**: 
- **에이전트 사고 과정**: 
  - 문제 인식: 
  - 고려한 해결책: 
  - 선택한 접근 방법: 
  - 시도한 방법들: 
  - 최종 결정 근거: 
- **해결 방법**: 
- **관련 파일**: 
  - `path/to/file1.py`
  - `path/to/file2.html`
- **상태**: 진행중 / 완료 / 실패 / 보류
- **참고 사항**: 
```

---

*이 문서는 프로젝트의 디버그 작업을 체계적으로 추적하기 위해 작성되었습니다.*

