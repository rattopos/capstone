#!/usr/bin/env python3
"""국내인구이동 이전 분기 데이터 추출 테스트"""

from templates.unified_generator import UnifiedReportGenerator
import sys

excel_path = 'uploads/분석표_25년_3분기_캡스톤업데이트_f1da33c3.xlsx'
gen = UnifiedReportGenerator('migration', excel_path, 2025, 3)

print('\n=== 데이터 추출 시작 ===')
result = gen.extract_all_data()

if result and 'table_data' in result:
    print(f'\n✅ 추출된 행 수: {len(result["table_data"])}')
    
    # 전국 데이터 확인
    nationwide = next((r for r in result['table_data'] if r.get('region_name') == '전국'), None)
    if nationwide:
        print(f'\n📊 전국 데이터:')
        print(f'  - 현재 (2025 3/4): {nationwide.get("value")}')
        print(f'  - 직전 분기 (2025 2/4): {nationwide.get("prev_value")}')
        print(f'  - 2분기 전 (2025 1/4): {nationwide.get("prev_prev_value")}')
        print(f'  - 3분기 전 (2024 4/4): {nationwide.get("prev_prev_prev_value")}')
        print(f'  - change_rate: {nationwide.get("change_rate")}')
        print(f'  - age_20_29: {nationwide.get("age_20_29")}')
        print(f'  - age_other: {nationwide.get("age_other")}')
    else:
        print('\n⚠️ 전국 데이터 없음')
    
    # 서울 데이터 확인
    seoul = next((r for r in result['table_data'] if r.get('region_name') == '서울'), None)
    if seoul:
        print(f'\n📊 서울 데이터:')
        print(f'  - 현재 (2025 3/4): {seoul.get("value")}')
        print(f'  - 직전 분기 (2025 2/4): {seoul.get("prev_value")}')
        print(f'  - 2분기 전 (2025 1/4): {seoul.get("prev_prev_value")}')
        print(f'  - 3분기 전 (2024 4/4): {seoul.get("prev_prev_prev_value")}')
    
    # 부산 데이터 확인
    busan = next((r for r in result['table_data'] if r.get('region_name') == '부산'), None)
    if busan:
        print(f'\n📊 부산 데이터:')
        print(f'  - 현재 (2025 3/4): {busan.get("value")}')
        print(f'  - 직전 분기: {busan.get("prev_value")}')
        print(f'  - 2분기 전: {busan.get("prev_prev_value")}')
        print(f'  - 3분기 전: {busan.get("prev_prev_prev_value")}')
    
    # 값의 합계 확인 (전국 제외)
    regional_sum = sum(r.get('value', 0) for r in result['table_data'] if r.get('region_name') != '전국')
    print(f'\n📊 지역별 합계 (전국 제외): {regional_sum:.1f}')
    
else:
    print('❌ 데이터 추출 실패')
    sys.exit(1)

print('\n✅ 테스트 완료')
