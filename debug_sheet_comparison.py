import pandas as pd

file = '분석표_25년 3분기_캡스톤(업데이트).xlsx'

# 고용률 분석 시트 읽기
df_emp = pd.read_excel(file, sheet_name='D(고용률)분석', header=None)
df_unemp = pd.read_excel(file, sheet_name='D(실업)분석', header=None)

print("="*70)
print("고용률 시트 - 헤더 행 (0-2)")
print("="*70)
for row_idx in range(3):
    print(f"\nRow {row_idx}:")
    for col_idx in range(min(10, len(df_emp.columns))):
        val = df_emp.iloc[row_idx, col_idx]
        val_str = str(val)[:30] if pd.notna(val) else 'NaN'
        print(f"  Col {col_idx}: {val_str}")

print("\n\n" + "="*70)
print("실업률 시트 - 헤더 행 (0-2)")
print("="*70)
for row_idx in range(3):
    print(f"\nRow {row_idx}:")
    for col_idx in range(min(10, len(df_unemp.columns))):
        val = df_unemp.iloc[row_idx, col_idx]
        val_str = str(val)[:30] if pd.notna(val) else 'NaN'
        print(f"  Col {col_idx}: {val_str}")

print("\n\n" + "="*70)
print("데이터 행 비교 (Row 3)")
print("="*70)
print("\n고용률 Row 3:")
print(df_emp.iloc[3, :].to_list())

print("\n\n실업률 Row 3:")
print(df_unemp.iloc[3, :].to_list())

# 지역명 컬럼 찾기 로직 테스트
print("\n\n" + "="*70)
print("지역명 컬럼 탐색 테스트")
print("="*70)

region_keywords = ['지역', 'region', '시도']

print("\n고용률 시트에서 찾기:")
for row_idx in range(3):  # 헤더 행
    row = df_emp.iloc[row_idx]
    for col_idx, cell_value in enumerate(row):
        if pd.isna(cell_value):
            continue
        cell_str = str(cell_value).strip().lower()
        print(f"  Row {row_idx}, Col {col_idx}: '{cell_str}'")
        for keyword in region_keywords:
            if keyword.lower() in cell_str:
                print(f"    ✅ 매치: '{keyword}' 포함됨!")
                break

print("\n\n실업률 시트에서 찾기:")
for row_idx in range(3):  # 헤더 행
    row = df_unemp.iloc[row_idx]
    for col_idx, cell_value in enumerate(row):
        if pd.isna(cell_value):
            continue
        cell_str = str(cell_value).strip().lower()
        print(f"  Row {row_idx}, Col {col_idx}: '{cell_str}'")
        for keyword in region_keywords:
            if keyword.lower() in cell_str:
                print(f"    ✅ 매치: '{keyword}' 포함됨!")
                break
