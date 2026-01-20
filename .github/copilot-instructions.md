# 🧑‍💻 capstone AI Agent Instructions

> **항상 작업을 시작하기 전에 이 문서를 읽으세요.**

## 프로젝트 개요
- **목적:** 엑셀 분석표 업로드 → 전국/부문별/지역별 보도자료 자동 생성 (웹/CLI)
- **주요 파일:**
  - app.py: Flask 진입점, REST API/대시보드
  - report_generator.py: CLI/웹 공용 보도자료 생성
  - templates/unified_generator.py: 모든 부문별 보도자료 생성의 핵심
  - config/reports.py, config/table_locations.py: 보고서/표 위치/설정
  - services/: 데이터 처리, 캐시, 요약, 보고서 생성 등
  - uploads/, exports/: 입출력 파일 저장

## 아키텍처/데이터 흐름
1. 엑셀 업로드(uploads/) → 전처리(services/excel_processor.py)
2. 부문별/시도별/요약 데이터 추출 및 캐시
3. Jinja2 템플릿 기반 HTML 생성(templates/)
4. 결과물은 exports/에 저장, 웹/CLI로 다운로드

## 개발/운영 규칙 (필수)
- **기초자료(원본표) 사용 금지:** 분석표만 사용, raw_excel_path/기초자료 관련 코드는 모두 제거/None 처리
- **시각화 요소(차트, 지도, 표지 등) 생성 금지:** 표와 텍스트만 생성, 관련 버그 리포트는 무시
- **외부 의존성 금지:** CDN, 외부 API, 외부 폰트/이미지 모두 금지, 로컬 자원만 사용
- **경로/포트:** 절대경로/하드코딩된 포트 금지, 상대경로/설정파일 사용
- **레거시 generator 금지:** 모든 보도자료 생성은 templates/unified_generator.py의 UnifiedReportGenerator만 사용
- **코드 무결성 감사관 역할:** 체크리스트 기반, 누락 없는 구현, 단계별 상태 검증/보고 필수

## 주요 워크플로우
- **웹 실행:**
  ```bash
  python -m venv VENV
  source VENV/bin/activate
  pip install -r requirements.txt
  python app.py
  ```
- **CLI 실행:**
  ```bash
  python report_generator.py --list
  python report_generator.py -e 분석표_25년_2분기_캡스톤.xlsx -r employment
  python report_generator.py -e 분석표_25년_2분기_캡스톤.xlsx
  ```
- **표 위치 관리:** data_table_locations.md → config/table_locations.py에서 자동 로드

## 프로젝트별 관례/패턴
- **ValueError 우선:** 데이터 누락시 fallback/default 사용 금지, 반드시 예외 발생
- **DEBUG_LOG.md 활용:** 버그 수정 전/후 반드시 로그 확인 및 기록
- **docs/PROCESS_FLOW.md, REPORT_FLOW.md:** 전체 흐름/작업 순서 참고
- **모든 설정/순서/종류는 config/에서 관리, 하드코딩 금지**

## 예외/특수 규칙
- **독립망 환경:** 외부 네트워크 차단, 모든 자원/패키지/이미지는 로컬만 사용
- **한글 파일명/경로 지원:** 파일명 인코딩, 상대경로 사용
- **이모지/외부 아이콘 금지:** 정부문서 규정, 텍스트/로컬 폰트만 사용

---
> 이 문서는 2026-01-20 기준, capstone 프로젝트의 AI 작업 가이드 최신 버전입니다.
