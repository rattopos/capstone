#!/usr/bin/env python3
"""부문별 전국 데이터 상세 확인"""

from templates.unified_generator import UnifiedReportGenerator
import sys

excel_path = 'uploads/분석표_25년_3분기_캡스톤업데이트_f1da33c3.xlsx'
year, quarter = 2025, 3

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

print(f'=== {year}년 {quarter}분기 전국 데이터 요약 ===\n')
print(f'{"부문":<12} {"현재값":<10} {"전년값":<10} {"증감/p":<10} {"단위":<10}')
print('-' * 60)

for report_id, report_name in reports:
    try:
        gen = UnifiedReportGenerator(report_id, excel_path, year, quarter)
        result = gen.extract_all_data()
        
        if not result or 'table_data' not in result:
            print(f'{report_name:<12} {"N/A":<10} {"N/A":<10} {"N/A":<10}')
            continue
        
        nationwide = next((r for r in result['table_data'] if r.get('region_name') == '전국'), None)
        
        if not nationwide:
            print(f'{report_name:<12} {"(전국없음)":<10}')
            continue
        
        value = nationwide.get('value', 0)
        prev_value = nationwide.get('prev_value', 0)
        change_rate = nationwide.get('change_rate', 0)
        
        # 단위 결정
        if report_id in ['export', 'import']:
            unit = '백만불'
        elif report_id in ['employment', 'unemployment']:
            unit = '천명'
        elif report_id == 'migration':
            unit = '명'
        elif report_id == 'price':
            unit = '지수'
        else:
            unit = '지수'
        
        # 증감 표시 (p는 퍼센트포인트, %는 증감률, 명은 절대값)
        if report_id in ['employment', 'unemployment']:
            change_str = f'{change_rate}p'
        elif report_id == 'migration':
            change_str = f'{change_rate}명'
        else:
            change_str = f'{change_rate}%'
        
        print(f'{report_name:<12} {value:<10.1f} {prev_value:<10.1f} {change_str:<10} {unit:<10}')
        
    except Exception as e:
        print(f'{report_name:<12} {"오류":<10} {str(e)[:20]:<10}')

print('\n✅ 전국 데이터 요약 완료')
