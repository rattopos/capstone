#!/usr/bin/env python3
"""국내인구이동 보고서 생성 테스트"""

from templates.unified_generator import UnifiedReportGenerator
import sys

excel_path = 'uploads/분석표_25년_3분기_캡스톤업데이트_f1da33c3.xlsx'

print('=== 국내인구이동 보고서 생성 시작 ===\n')

try:
    gen = UnifiedReportGenerator('migration', excel_path, 2025, 3)
    
    # 데이터 추출
    result = gen.extract_all_data()
    
    if not result or 'table_data' not in result:
        print('❌ 데이터 추출 실패')
        sys.exit(1)
    
    print(f'✅ 데이터 추출 완료: {len(result["table_data"])}개 지역')
    
    # HTML 확인
    if 'html' in result:
        html = result['html']
    else:
        print('⚠️ HTML이 result에 없음. extract_all_data()가 HTML을 포함하지 않는 것 같습니다.')
        # HTML 없이도 데이터 확인
        html = None
    
    if html:
        # 파일 저장
        output_path = 'exports/국내인구이동_2025년_3분기_테스트.html'
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html)
        
        print(f'\n✅ 보고서 생성 완료: {output_path}')
        print(f'   파일 크기: {len(html):,} bytes')
    
    # 전국 데이터 확인
    nationwide = next((r for r in result['table_data'] if r.get('region_name') == '전국'), None)
    if nationwide:
        print(f'\n📊 전국 데이터:')
        print(f'   2025 3/4: {nationwide.get("value")}명')
        print(f'   2025 2/4: {nationwide.get("prev_value")}명')
        print(f'   2025 1/4: {nationwide.get("prev_prev_value")}명')
        print(f'   2024 4/4: {nationwide.get("prev_prev_prev_value")}명')
    
except Exception as e:
    print(f'❌ 오류 발생: {e}')
    import traceback
    traceback.print_exc()
    sys.exit(1)

print('\n✅ 테스트 완료')
