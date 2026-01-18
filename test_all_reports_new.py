import sys
import os

sys.path.insert(0, '/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone')
os.chdir('/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone')

from templates.unified_generator import (
    EmploymentRateGenerator, UnemploymentGenerator, DomesticMigrationGenerator
)

test_cases = [
    ('employment', EmploymentRateGenerator, '고용률'),
    ('unemployment', UnemploymentGenerator, '실업률'),
    ('migration', DomesticMigrationGenerator, '국내인구이동'),
]

print("="*70)
print("전체 리포트 생성 테스트")
print("="*70)

for report_id, generator_class, report_name in test_cases:
    try:
        print(f"\n테스트: {report_name} ({report_id})")
        print("-"*50)
        
        gen = generator_class(
            excel_path='분석표_25년 3분기_캡스톤(업데이트).xlsx',
            year=2025,
            quarter=3
        )
        
        gen.load_data()
        data = gen.extract_all_data()
        
        if 'table_data' in data:
            table_data = data['table_data']
            print(f"✅ 성공!")
            print(f"   - 데이터 행 수: {len(table_data)}")
            if table_data:
                nationwide = [d for d in table_data if d.get('region_name') == '전국']
                if nationwide:
                    print(f"   - 전국 지수: {nationwide[0].get('index_current')}")
                    print(f"   - 증감: {nationwide[0].get('change_rate')}%")
        else:
            print(f"❌ 예상치 못한 데이터 형식")
    
    except Exception as e:
        print(f"❌ 오류: {e}")

print("\n" + "="*70)
print("모든 리포트 테스트 완료!")
print("="*70)
