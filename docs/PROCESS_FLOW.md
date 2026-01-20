# 지역경제동향 프로젝트 전체통합 프로세스

본 문서는 **요청된 기준 흐름**에 맞춰 실제 코드 경로를 정리한다.
목표: 업로드 → 전처리 → 부문별 우선 처리 → 요약·부문별·시도별 순으로 전체통합(페이지 구분 없음) → 내보내기.

---

## 1. 업로드
- 엔드포인트: /api/upload
- 파일 저장: uploads/
- 관련: `upload_excel()` in routes/api.py

---

## 2. 전처리 (openpyxl 기반 수식 계산/검증 포함)
- 전처리 실행: `preprocess_excel()` in services/excel_processor.py
- 분석 시트 계산 결과 캐시 등록: `set_cached_calculated_path()` in services/excel_cache.py
- 목적: 이후 데이터 추출 시 계산 결과 재사용

---

## 3. 부문별 우선 처리 (캐시 생성)
- 생성 순서: 부문별 → 시도별 → 요약
- 실제 생성 루틴: `_generate_all_reports_core()` in routes/api.py
- 부문별 결과는 시도별/요약 데이터에 재사용됨
  - 캐시 저장: `get_sector_data()` in services/report_generator.py

---

## 4. 전체통합(concatenation, 페이지 구분 없음)
- 출력 순서: **요약 → 부문별 → 시도별**
- HTML 결합: `_export_hwp_ready_core()` in routes/api.py
- 입력 소스:
  - 부문별/요약: exports/_temp/output
  - 시도별: exports/_temp/regional_output

---

## 5. 내보내기
- 최종 HTML 저장:
  - exports/지역경제동향_{year}년_{quarter}분기.html
- 다운로드/뷰 URL 제공:
  - /exports/...

---

## 참고: CLI 전체통합
- 스크립트: generate_full_report.py
- 생성 순서(캐시 목적): 부문별 → 시도별 → 요약
- 출력 순서: 요약 → 부문별 → 시도별
