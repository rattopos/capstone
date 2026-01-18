import pandas as pd
import sys
sys.path.insert(0, '/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone')

from templates.unified_generator import EmploymentRateGenerator

# 설정
config = {
    'id': 'employment',
    'name': '고용률',
    'sheet': 'D(고용률)분석',
    'category': 'employment',
    'metadata_columns': {
        'region': ['지역', 'region', '시도'],
        'code': ['코드', 'code', '산업코드', '업태코드', '품목코드', '분류코드'],
        'name': ['이름', 'name', '산업명', '산업 이름', '업태명', '품목명', '품목 이름', '공정이름', '공정명', '연령']
    }
}

file = '분석표_25년 3분기_캡스톤(업데이트).xlsx'

try:
    # 제너레이터 생성
    gen = EmploymentRateGenerator(config, file, 2025, 3)
    
    # load_data 호출 (지역명 컬럼 탐색)
    print("="*60)
    print("고용률 제너레이터 데이터 로드:")
    print("="*60)
    gen.load_data()
    
    print(f"\n제너레이터 상태 확인:")
    print(f"  - df_aggregation 로드됨: {gen.df_aggregation is not None}")
    print(f"  - df_analysis 로드됨: {gen.df_analysis is not None}")
    print(f"  - region_name_col: {gen.region_name_col}")
    print(f"  - industry_code_col: {gen.industry_code_col}")
    print(f"  - industry_name_col: {gen.industry_name_col}")
    print(f"  - data_start_row: {gen.data_start_row}")
    print(f"  - target_col: {gen.target_col}")
    print(f"  - prev_y_col: {gen.prev_y_col}")
    
    # 실제로 사용되는 시트 확인
    print(f"\n사용 중인 시트:")
    if gen.df_aggregation is not None:
        print(f"  - 집계 시트: {gen.df_aggregation.shape}")
    if gen.df_analysis is not None:
        print(f"  - 분석 시트: {gen.df_analysis.shape}")
    
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
