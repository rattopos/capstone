"""단순 테스트 - 광공업생산만"""

import sys
sys.path.insert(0, '/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone')

excel_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/분석표_25년 3분기_캡스톤(업데이트).xlsx"

try:
    from templates.unified_generator import UnifiedReportGenerator
    
    print("광공업생산 테스트:")
    gen = UnifiedReportGenerator('manufacturing', excel_path, 2025, 3)
    
    print(f"  데이터 행 수: {len(gen.data)}")
    print(f"  컬럼들: {list(gen.data[0].keys()) if gen.data else '없음'}")
    
    # 전국 찾기
    for row in gen.data:
        region = str(row.get('region', ''))
        if region in ['전국', '전체', '합계']:
            print(f"\n  전국 데이터:")
            print(f"    지역: {row.get('region')}")
            print(f"    현재값: {row.get('current_value')}")
            print(f"    증감률: {row.get('change_rate')}")
            break
    else:
        print("\n  ⚠️ 전국 데이터 없음")
        print("\n  처음 3행:")
        for i, row in enumerate(gen.data[:3], 1):
            print(f"    {i}. region={row.get('region')}, current={row.get('current_value')}, rate={row.get('change_rate')}")
    
except Exception as e:
    print(f"❌ 오류: {e}")
    import traceback
    traceback.print_exc()
