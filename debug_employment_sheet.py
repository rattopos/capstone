import pandas as pd
import sys

file = '분석표_25년 3분기_캡스톤(업데이트).xlsx'

try:
    # 고용률 분석 시트 읽기
    df_employment = pd.read_excel(file, sheet_name='D(고용률)분석', header=None)
    print("="*60)
    print("고용률 분석 시트 구조:")
    print("="*60)
    print(f"Shape: {df_employment.shape}")
    print(f"\n처음 5행 (10열까지):")
    print(df_employment.iloc[:5, :10])
    
    print(f"\n\n첫 번째 컬럼의 내용 (모든 행):")
    print(df_employment.iloc[:, 0].tolist()[:20])
    
    print(f"\n\n전국 데이터 찾기:")
    nationwide_found = False
    for idx in range(len(df_employment)):
        for cidx in range(len(df_employment.columns)):
            val = df_employment.iloc[idx, cidx]
            if pd.notna(val) and str(val).strip() == '전국':
                print(f"  ✅ Row {idx}, Col {cidx}: 전국 발견")
                print(f"     Row 데이터 (처음 15개): {df_employment.iloc[idx, :15].tolist()}")
                nationwide_found = True
                break
        if nationwide_found:
            break
    
    if not nationwide_found:
        print("  ❌ 전국 데이터 찾지 못함")
    
    print(f"\n\n실업률 분석 시트와 비교:")
    print("="*60)
    df_unemployment = pd.read_excel(file, sheet_name='D(실업)분석', header=None)
    print(f"Shape: {df_unemployment.shape}")
    print(f"\n처음 5행 (10열까지):")
    print(df_unemployment.iloc[:5, :10])
    
    print(f"\n\n전국 데이터 찾기:")
    for idx in range(len(df_unemployment)):
        for cidx in range(len(df_unemployment.columns)):
            val = df_unemployment.iloc[idx, cidx]
            if pd.notna(val) and str(val).strip() == '전국':
                print(f"  ✅ Row {idx}, Col {cidx}: 전국 발견")
                print(f"     Row 데이터 (처음 15개): {df_unemployment.iloc[idx, :15].tolist()}")
                break
        else:
            continue
        break
    
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
