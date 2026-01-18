import sys
import os

# 경로 설정
sys.path.insert(0, '/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone')
os.chdir('/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone')

# 리포트 생성 테스트
from templates.unified_generator import EmploymentRateGenerator

try:
    print("="*60)
    print("테스트: 고용률 리포트 생성")
    print("="*60)
    
    # 제너레이터 생성 및 테스트
    gen = EmploymentRateGenerator(
        excel_path='분석표_25년 3분기_캡스톤(업데이트).xlsx',
        year=2025,
        quarter=3
    )
    
    print("\n✅ 제너레이터 초기화 완료")
    
    # 데이터 로드
    gen.load_data()
    print(f"\n데이터 로드 완료")
    
    # 데이터 추출
    print(f"데이터 추출 시작...")
    data = gen.extract_all_data()
    print(f"✅ 데이터 추출 완료!")
    
    # 데이터 구조 확인
    if 'table_data' in data:
        table_data = data['table_data']
        print(f"\n추출된 데이터:")
        print(f"  - 전국 데이터: {len([d for d in table_data if d.get('region_name') == '전국'])} 행")
        print(f"  - 총 데이터 행 수: {len(table_data)}")
        nationwide = [d for d in table_data if d.get('region_name') == '전국']
        if nationwide:
            print(f"  - 전국 지수: {nationwide[0].get('index_current')}")
            print(f"  - 전년 대비 증감: {nationwide[0].get('change_rate')}%")
    
    print(f"\n✅ 고용률 리포트 데이터 추출 성공!")
    
except Exception as e:
    print(f"\n❌ 오류: {e}")
    import traceback
    traceback.print_exc()
