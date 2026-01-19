"""국내인구이동 전국 데이터 생성 여부 테스트"""

import sys
sys.path.insert(0, '/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone')

from templates.unified_generator import UnifiedReportGenerator

excel_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/uploads/지역경제동향_2025년_3분기.xlsx"
year, quarter = 2025, 3

print("=" * 60)
print("국내인구이동 전국 데이터 생성 테스트")
print("=" * 60)

# migration 생성
gen = UnifiedReportGenerator('migration', excel_path, year, quarter)

print(f"\n✅ 총 생성된 데이터 수: {len(gen.data)}개")

# 지역 목록 확인
regions = [row.get('region') for row in gen.data]
print(f"✅ 지역 목록: {regions}")

# 전국 데이터 확인
nationwide = [row for row in gen.data if row.get('region') in ['전국', '전체', '합계']]

print("\n" + "=" * 60)
if nationwide:
    print("⚠️ 전국 데이터가 생성되었습니다 (잘못됨)")
    print(f"   데이터: {nationwide[0]}")
else:
    print("✅ 전국 데이터가 생성되지 않았습니다 (정상)")
    print("   국내이동은 지역간 이동이므로 전국 합계는 의미 없음")
print("=" * 60)

# 서울 데이터 확인
seoul = [row for row in gen.data if row.get('region') == '서울']
if seoul:
    print(f"\n서울 데이터 샘플:")
    print(f"  현재값: {seoul[0].get('current_value')}")
    print(f"  전년값: {seoul[0].get('prev_year_value')}")
    print(f"  증감: {seoul[0].get('change_rate')}")
