import os
import pandas as pd

INPUT_DIR = "."
OUTPUT_FILE = "exports/분석표_데이터_매핑결과.csv"

# '분석표'로 시작하는 엑셀 파일 찾기
def find_analysis_excels(input_dir):
    # 항상 '분석표.xlsx'만 처리
    target = os.path.join(input_dir, "분석표.xlsx")
    return [target] if os.path.exists(target) else []

# 엑셀 파일에서 모든 시트의 데이터프레임을 합침
def read_all_sheets(excel_path):
    xls = pd.ExcelFile(excel_path)
    df_list = []
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        df["__sheet__"] = sheet
        df_list.append(df)
    return pd.concat(df_list, ignore_index=True)

# 데이터 매핑 예시: 컬럼명 표준화 및 주요 필드 추출
def map_analysis_data(df):
    # 예시: 컬럼명 소문자, 공백제거
    df.columns = [str(c).strip().lower().replace(" ", "") for c in df.columns]
    # 주요 필드만 추출(예시)
    main_cols = [c for c in df.columns if any(k in c for k in ["지역", "시도", "지표", "값", "value", "amount", "date", "기간"])]
    if main_cols:
        return df[main_cols + ["__sheet__"]]
    else:
        return df

def main():
    excels = find_analysis_excels(INPUT_DIR)
    if not excels:
        print("'분석표'로 시작하는 엑셀 파일이 없습니다.")
        return
    path = excels[0]
    print(f"처리 중: {path}")
    df = read_all_sheets(path)
    mapped = map_analysis_data(df)
    mapped["__file__"] = os.path.basename(path)
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    mapped.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")
    print(f"매핑 결과 저장: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
