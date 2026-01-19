#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 파일에서 모든 데이터테이블을 추출하여 HTML 테이블로 출력
"""
import pandas as pd
from typing import Any, cast

from config.settings import BASE_DIR

# Excel 파일 경로
EXCEL_PATH = str((BASE_DIR / "분석표_25년 3분기_캡스톤(업데이트).xlsx").resolve())
pd_any: Any = cast(Any, pd)

def extract_all_sheets() -> list[dict[str, Any]]:
    """모든 시트와 테이블 추출"""
    excel_file = pd_any.ExcelFile(EXCEL_PATH)
    all_tables: list[dict[str, Any]] = []
    
    print(f"📊 Excel 파일 읽기: {EXCEL_PATH}")
    print(f"📋 총 시트 수: {len(excel_file.sheet_names)}")
    print("-" * 80)
    
    for sheet_name in excel_file.sheet_names:
        print(f"\n📄 시트: {sheet_name}")
        df = pd_any.read_excel(EXCEL_PATH, sheet_name=sheet_name)
        print(f"   크기: {df.shape[0]} 행 × {df.shape[1]} 열")
        
        all_tables.append({
            'sheet_name': sheet_name,
            'dataframe': df,
            'shape': df.shape
        })
    
    return all_tables

def dataframe_to_html_table(df: pd.DataFrame, title: str = "") -> str:
    """DataFrame을 HTML 테이블로 변환"""
    html = f"""<div style="margin-bottom: 40px; page-break-inside: avoid;">
    <h3 style="color: #333; border-bottom: 3px solid #0066cc; padding-bottom: 10px;">{title}</h3>
    <p style="font-size: 12px; color: #666;">크기: {df.shape[0]} 행 × {df.shape[1]} 열</p>
    {df.to_html(index=False, border=1, justify='left')}
</div>"""
    return html

def main():
    # 모든 시트 추출
    tables = extract_all_sheets()
    
    # HTML 생성
    sheet_count = len(tables)
    html_content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel 데이터 테이블 추출</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #0066cc;
            text-align: center;
            border-bottom: 4px solid #0066cc;
            padding-bottom: 15px;
            margin-bottom: 30px;
        }}
        h2 {{
            color: #333;
            margin-top: 40px;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #ddd;
        }}
        h3 {{
            color: #333;
            border-bottom: 3px solid #0066cc;
            padding-bottom: 10px;
            margin-top: 25px;
            margin-bottom: 15px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            margin-bottom: 20px;
        }}
        th {{
            background-color: #0066cc;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: bold;
            border: 1px solid #004499;
        }}
        td {{
            padding: 10px;
            border: 1px solid #ddd;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        tr:hover {{
            background-color: #f0f0f0;
        }}
        .sheet-info {{
            background-color: #e8f4f8;
            padding: 10px;
            border-left: 4px solid #0066cc;
            margin-bottom: 15px;
            border-radius: 4px;
        }}
        .table-count {{
            text-align: center;
            color: #666;
            font-size: 14px;
            margin: 20px 0;
            padding: 10px;
            background-color: #f0f0f0;
            border-radius: 4px;
        }}
        @media print {{
            body {{
                margin: 0;
                padding: 0;
                background-color: white;
            }}
            .container {{
                box-shadow: none;
                padding: 0;
                max-width: 100%;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Excel 데이터 테이블 목록</h1>
        <div class="table-count">
            총 <strong>{sheet_count}</strong>개 시트에서 데이터 추출 완료
        </div>
"""
    
    # 각 시트별 테이블 추가
    for i, table_info in enumerate(tables, 1):
        sheet_name = table_info['sheet_name']
        df = table_info['dataframe']
        
        html_content += f"""
        <h2>📋 {i}. {sheet_name}</h2>
        <div class="sheet-info">
            <strong>크기:</strong> {df.shape[0]} 행 × {df.shape[1]} 열 | 
            <strong>컬럼:</strong> {', '.join(df.columns.tolist()[:5])}{"..." if df.shape[1] > 5 else ""}
        </div>
        """
        
        # 처음 50개 행만 표시
        df_display = df.head(50)
        html_content += df_display.to_html(index=False, border=1, justify='left', classes='data-table')
        
        if df.shape[0] > 50:
            html_content += f"""
            <p style="text-align: center; color: #999; font-style: italic;">
                ✂️ 처음 50개 행 표시 (전체 {df.shape[0]}개 행 중)
            </p>
            """
    
    html_content += """
    </div>
</body>
</html>
"""
    
    # 출력 파일
    output_path = BASE_DIR / "exports" / "extracted_tables.html"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("\n" + "="*80)
    print(f"✅ HTML 파일 생성 완료: {output_path}")
    print("="*80)

if __name__ == "__main__":
    main()
