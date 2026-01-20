
import pandas as pd
from pathlib import Path

excel_path = Path('분석표_25년 3분기_캡스톤(업데이트).xlsx')
sheet_name = 'A(광공업생산)집계'

try:
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, nrows=10)
    print(f"Sheet: {sheet_name}")
    print(df.to_string())
except Exception as e:
    print(f"Error reading sheet {sheet_name}: {e}")
