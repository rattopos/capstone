# 업로드 → 다운로드 전체 흐름 (모듈 · 로직 · 의사코드)

아래는 사용자가 **업로드**부터 **다운로드**까지 수행하는 과정을 시간순으로 정리한 문서다. 각 단계는 **모듈**, **로직**, **의사코드** 세트로 구성된다.

---

## 1) 업로드 및 요청 접수
**모듈**
- [app.py](../app.py)
- [routes/main.py](../routes/main.py)
- [routes/api.py](../routes/api.py)
- [dashboard.html](../dashboard.html)

**로직**
- 사용자가 엑셀 파일을 업로드한다.
- 서버는 업로드 파일을 저장하고, 연도/분기/보고서 선택 정보를 함께 받는다.
- “전체 생성” 또는 “특정 보고서 생성” 요청을 분기한다.

**의사코드**
```
on UploadRequest:
  save uploaded excel file
  read year, quarter, report selection
  if generate_all:
    call generate_all_reports()
  else:
    call generate_single_report(report_id)
```

---

## 2) 보도자료 정의 로딩
**모듈**
- [config/reports.py](../config/reports.py)

**로직**
- 생성 대상 보도자료 목록과 템플릿/생성기 정보를 불러온다.
- 보고서 종류별로 “부문별 → 시도별 → 요약” 순서를 유지한다.

**의사코드**
```
report_order = REPORT_ORDER
sector_reports = SECTOR_REPORTS
regional_reports = REGIONAL_REPORTS
summary_reports = SUMMARY_REPORTS
```

---

## 3) 엑셀 데이터 로딩/캐시
**모듈**
- [services/excel_cache.py](../services/excel_cache.py)
- [services/excel_processor.py](../services/excel_processor.py)
- [utils/excel_utils.py](../utils/excel_utils.py)

**로직**
- 엑셀 파일을 읽고 필요한 시트/범위를 확보한다.
- 반복 호출을 줄이기 위해 캐시를 적용한다.

**의사코드**
```
workbook = load_excel_with_cache(excel_path)
```

---

## 4) 부문별 보도자료 생성
**모듈**
- [report_generator.py](../report_generator.py)
- [templates/unified_generator.py](../templates/unified_generator.py)
- [templates/base_generator.py](../templates/base_generator.py)
- [utils/text_utils.py](../utils/text_utils.py)

**로직**
- 각 부문(광공업, 서비스업, 소비 등)에 대해 데이터를 추출한다.
- 증감률 계산, 상·하위 지역/업종 산출 후 나레이션을 생성한다.
- 템플릿에 데이터/나레이션을 결합해 HTML을 만든다.

**의사코드**
```
for report in SECTOR_REPORTS:
  data = extract_data(report)
  narrative = build_narrative(data)
  html = render_template(report.template, data, narrative)
  save_html(html)
```

---

## 5) 시도별(지역별) 보도자료 생성
**모듈**
- [services/report_generator.py](../services/report_generator.py)
- [templates/unified_generator.py](../templates/unified_generator.py)

**로직**
- 시도별 섹션 데이터를 수집해 지표별 요약 문장을 생성한다.
- 지역별 HTML을 별도 파일로 저장한다.

**의사코드**
```
for region in REGIONAL_REPORTS:
  data = extract_regional_data(region)
  narrative = build_regional_narrative(data)
  html = render_template(region.template, data, narrative)
  save_html(html)
```

---

## 6) 요약 보도자료 생성
**모듈**
- [services/summary_data.py](../services/summary_data.py)
- [services/report_generator.py](../services/report_generator.py)

**로직**
- 여러 부문 데이터를 묶어 요약 테이블/문장을 생성한다.
- 요약 템플릿에 반영해 최종 요약 HTML을 만든다.

**의사코드**
```
for report in SUMMARY_REPORTS:
  summary_data = build_summary_data(excel_path)
  html = render_template(report.template, summary_data)
  save_html(html)
```

---

## 7) 다운로드 제공
**모듈**
- [routes/api.py](../routes/api.py)
- [routes/main.py](../routes/main.py)

**로직**
- 생성된 HTML 파일을 다운로드 경로로 제공한다.
- 성공/실패 목록을 사용자에게 반환한다.

**의사코드**
```
return download_links(success_files), error_list
```
