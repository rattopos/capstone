#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, '.')

from templates.unified_generator import UnifiedReportGenerator
from pathlib import Path

excel_file = '분석표_25년 3분기_캡스톤(업데이트).xlsx'

print(f'[테스트] 파일: {excel_file}\n')
try:
    # 국내인구이동 보고서 생성 테스트 (2025년 3분기)
    generator = UnifiedReportGenerator('migration', excel_file, 2025, 3)
    generator.load_data()
    data = generator.extract_all_data()
    
    if data:
        print('\n✅ [성공] 국내인구이동 데이터 추출 완료!')
        print(f'   추출 데이터 수: {len(data)}개')
        if len(data) > 0:
            print(f'   첫 번째 항목 키: {list(data[0].keys())}')
            print(f'\n   샘플 데이터 (첫 항목):')
            for key, val in data[0].items():
                print(f'     {key}: {val}')
    else:
        print('\n❌ [실패] 데이터가 비어있음')
except Exception as e:
    print(f'\n❌ [오류] {e}')
    import traceback
    traceback.print_exc()
