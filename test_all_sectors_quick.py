#!/usr/bin/env python3
"""모든 부문별 보고서 재생성 스크립트"""

from templates.unified_generator import UnifiedReportGenerator
import sys
import os

excel_path = 'uploads/분석표_25년_3분기_캡스톤업데이트_f1da33c3.xlsx'
year, quarter = 2025, 3

# 생성할 보고서 목록 (unified_generator 사용하는 것들)
reports = [
    ('manufacturing', '광공업생산'),
    ('service', '서비스업생산'),
    ('consumption', '소비동향'),
    ('construction', '건설동향'),
    ('export', '수출'),
    ('import', '수입'),
    ('price', '물가동향'),
    ('employment', '고용률'),
    ('unemployment', '실업률'),
    ('migration', '국내인구이동')
]

print(f'=== {year}년 {quarter}분기 부문별 보고서 생성 ===\n')

results = []

for report_id, report_name in reports:
    try:
        print(f'\n[{report_name}] 생성 시작...')
        gen = UnifiedReportGenerator(report_id, excel_path, year, quarter)
        result = gen.extract_all_data()
        
        if not result or 'table_data' not in result:
            print(f'  ❌ 데이터 추출 실패')
            results.append((report_name, False, '데이터 추출 실패'))
            continue
        
        # 전국 데이터 확인
        nationwide = next((r for r in result['table_data'] if r.get('region_name') == '전국'), None)
        if nationwide:
            value = nationwide.get('value')
            prev_value = nationwide.get('prev_value')
            change_rate = nationwide.get('change_rate')
            print(f'  ✅ 전국: {value} (전년 {prev_value}, 증감 {change_rate})')
        else:
            print(f'  ⚠️ 전국 데이터 없음 (지역 수: {len(result["table_data"])})')
        
        results.append((report_name, True, f'{len(result["table_data"])}개 지역'))
        
    except Exception as e:
        print(f'  ❌ 오류: {e}')
        results.append((report_name, False, str(e)[:50]))

print('\n\n=== 생성 결과 요약 ===')
print(f'{"보고서":<15} {"상태":<8} {"상세":<40}')
print('-' * 65)

for name, success, detail in results:
    status = '✅ 성공' if success else '❌ 실패'
    print(f'{name:<15} {status:<8} {detail:<40}')

success_count = sum(1 for _, success, _ in results if success)
print(f'\n총 {len(reports)}개 중 {success_count}개 성공')

if success_count == len(reports):
    print('\n🎉 모든 보고서 생성 완료!')
else:
    print(f'\n⚠️ {len(reports) - success_count}개 보고서 생성 실패')
    sys.exit(1)
