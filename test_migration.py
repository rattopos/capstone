#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, '.')

from templates.unified_generator import UnifiedReportGenerator
from pathlib import Path
import os

files = sorted([f for f in list(Path('.').glob('*.xlsx')) + list(Path('uploads').glob('*.xlsx')) + list(Path('exports').glob('*.xlsx')) if '2025' in f.name and '3' in f.name], key=os.path.getmtime, reverse=True)

if files:
    excel_file = str(files[0])
    print(f'[테스트] 파일: {excel_file}\n')
    try:
        # 국내인구이동 보고서 생성 테스트 (UnifiedReportGenerator 직접 사용)
        generator = UnifiedReportGenerator('migration', excel_file, 2025, 3)
        generator.load_data()
        data = generator.extract_all_data()
        
        if data:
            print('\n✅ [성공] 국내인구이동 데이터 추출 완료!')
            print(f'   추출 데이터 수: {len(data)}개 지역/항목')
            if len(data) > 0:
                print(f'   샘플 데이터: {list(data[0].keys())}')
        else:
            print('\n❌ [실패] 데이터가 비어있음')
    except Exception as e:
        print(f'\n❌ [오류] {e}')
        import traceback
        traceback.print_exc()
else:
    print('[오류] 테스트 파일을 찾을 수 없습니다')
