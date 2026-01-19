#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
실제 템플릿에서 추출되는 데이터를 테이블로 출력
"""
from typing import Any, cast

import pandas as pd
from config.settings import BASE_DIR
from config.reports import SUMMARY_REPORTS, SECTOR_REPORTS
from report_generator import ReportGenerator

EXCEL_PATH = str((BASE_DIR / "분석표_25년 3분기_캡스톤(업데이트).xlsx").resolve())

def format_data_for_display(data: Any, max_depth: int = 2, current_depth: int = 0) -> Any:
    """복잡한 데이터 구조를 보기 좋게 변환"""
    if current_depth > max_depth:
        return f"[중첩 데이터 {type(data).__name__}]"
    
    if isinstance(data, dict):
        data_dict = cast(dict[str, Any], data)
        result: dict[str, Any] = {}
        for k, v in list(data_dict.items())[:10]:  # 처음 10개 키만
            if isinstance(v, (dict, list)):
                if isinstance(v, dict):
                    v_dict = cast(dict[str, Any], v)
                    result[k] = f"[Dict] {len(v_dict)} items"
                else:
                    v_list = cast(list[Any], v)
                    result[k] = f"[List] {len(v_list)} items"
            elif isinstance(v, (int, float, str, bool)):
                result[k] = v
            else:
                result[k] = str(type(v).__name__)
        return result
    elif isinstance(data, list):
        data_list = cast(list[Any], data)
        if len(data_list) == 0:
            return "[]"
        first = data_list[0]
        if isinstance(first, dict):
            first_dict = cast(dict[str, Any], first)
            return f"[List of {len(data_list)} dicts] Keys: {list(first_dict.keys())[:5]}"
        else:
            return f"[List of {len(data_list)} items] Sample: {first}"
    else:
        return str(data)

def extract_table_data(data_dict: dict[str, Any], key_path: str = "") -> list[dict[str, Any]]:
    """데이터 구조에서 테이블 데이터 추출"""
    tables: list[dict[str, Any]] = []
    
    def traverse(obj: Any, path: str = "") -> None:
        if isinstance(obj, list):
            obj_list = cast(list[Any], obj)
        else:
            obj_list = []

        if obj_list and isinstance(obj_list[0], dict):
            # DataFrame 형태의 리스트 발견
            df = pd.DataFrame(obj_list)
            tables.append({
                'path': path or 'root',
                'shape': df.shape,
                'columns': list(df.columns),
                'data': df.head(10)  # 처음 10행만
            })
        elif isinstance(obj, dict):
            obj_dict = cast(dict[str, Any], obj)
            for k, v in obj_dict.items():
                new_path = f"{path}.{k}" if path else k
                traverse(v, new_path)
    
    traverse(data_dict)
    return tables

def generate_extracted_data_html() -> str:
    """각 보도자료에서 추출된 데이터를 HTML로 생성"""
    
    print("🔄 ReportGenerator 초기화 중...")
    generator = ReportGenerator(EXCEL_PATH)
    
    html_parts: list[str] = []
    
    # 요약 보도자료
    print("\n📊 요약 보도자료 처리 중...")
    for i, report in enumerate(SUMMARY_REPORTS, 1):
        report_id = report['id']
        report_name = report['name']
        print(f"  {i}/{len(SUMMARY_REPORTS)}: {report_name}...")
        
        try:
            data: dict[str, Any] = generator.extract_data(report_id)
            
            # 데이터 구조 분석
            html_parts.append(f"""
            <div style="page-break-inside: avoid; margin-bottom: 40px;">
                <h3 style="color: #0066cc; border-bottom: 2px solid #0066cc; padding-bottom: 10px;">
                    {report['icon']} {report_name}
                </h3>
                <div style="background-color: #f9f9f9; padding: 10px; border-radius: 4px; margin-bottom: 15px;">
                    <strong>ID:</strong> {report_id}<br>
                    <strong>시트:</strong> {report['sheet']}<br>
                    <strong>템플릿:</strong> {report['template']}<br>
                    <strong>데이터 키:</strong> {', '.join(list(data.keys())[:8])}
                </div>
            """)
            
            # 테이블 데이터 추출
            tables = extract_table_data(data)
            if tables:
                html_parts.append("<strong>📋 추출된 테이블:</strong><ul>")
                for table_info in tables[:3]:  # 최대 3개 테이블
                    path = table_info['path']
                    shape = table_info['shape']
                    cols = table_info['columns']
                    df_sample = table_info['data']
                    
                    html_parts.append(f"""
                    <li>
                        <strong>{path}</strong>: {shape[0]}행 × {shape[1]}열<br>
                        <em style="color: #666;">컬럼: {', '.join(cols[:10])}</em><br>
                    """)
                    
                    # 샘플 데이터 테이블
                    html_parts.append(df_sample.to_html(index=False, border=1, justify='left'))
                    html_parts.append("</li>")
                
                html_parts.append("</ul>")
            else:
                # 테이블 데이터가 없는 경우 데이터 구조 표시
                html_parts.append("<strong>📋 데이터 구조:</strong><pre style='background: #f5f5f5; padding: 10px; overflow-x: auto;'>")
                for key, value in list(data.items())[:8]:
                    if isinstance(value, dict):
                        value_dict = cast(dict[str, Any], value)
                        html_parts.append(f"{key}: Dict({len(value_dict)} items)\n")
                    elif isinstance(value, list):
                        value_list = cast(list[Any], value)
                        html_parts.append(f"{key}: List({len(value_list)} items)\n")
                    else:
                        html_parts.append(f"{key}: {type(value).__name__}\n")
                html_parts.append("</pre>")
            
            html_parts.append("</div>")
        
        except Exception as e:
            html_parts.append(f"""
            <div style="background-color: #ffebee; padding: 15px; border-left: 4px solid #f44336; margin-bottom: 20px;">
                <strong style="color: #f44336;">❌ 오류: {report_name}</strong><br>
                {str(e)[:200]}
            </div>
            """)
    
    # 부문별 보도자료
    print("\n🏭 부문별 보도자료 처리 중...")
    for i, report in enumerate(SECTOR_REPORTS, 1):
        report_id = report['id']
        report_name = report['name']
        print(f"  {i}/{len(SECTOR_REPORTS)}: {report_name}...")
        
        try:
            data: dict[str, Any] = generator.extract_data(report_id)
            
            html_parts.append(f"""
            <div style="page-break-inside: avoid; margin-bottom: 40px;">
                <h3 style="color: #333; border-bottom: 2px solid #666; padding-bottom: 10px;">
                    {report['icon']} {report_name} <span style="color: #999; font-size: 12px;">({report['category']})</span>
                </h3>
                <div style="background-color: #f9f9f9; padding: 10px; border-radius: 4px; margin-bottom: 15px;">
                    <strong>ID:</strong> {report_id}<br>
                    <strong>시트:</strong> {report['sheet']}<br>
                    <strong>카테고리:</strong> {report['category']}<br>
                    <strong>데이터 키:</strong> {', '.join(list(data.keys())[:8])}
                </div>
            """)
            
            # 테이블 데이터 추출
            tables = extract_table_data(data)
            if tables:
                html_parts.append("<strong>📋 추출된 테이블:</strong><ul>")
                for table_info in tables[:3]:
                    path = table_info['path']
                    shape = table_info['shape']
                    cols = table_info['columns']
                    df_sample = table_info['data']
                    
                    html_parts.append(f"""
                    <li>
                        <strong>{path}</strong>: {shape[0]}행 × {shape[1]}열<br>
                        <em style="color: #666;">컬럼: {', '.join(str(c) for c in cols[:10])}</em><br>
                    """)
                    
                    html_parts.append(df_sample.to_html(index=False, border=1, justify='left'))
                    html_parts.append("</li>")
                
                html_parts.append("</ul>")
            else:
                html_parts.append("<strong>📋 데이터 구조:</strong><pre style='background: #f5f5f5; padding: 10px; overflow-x: auto;'>")
                for key, value in list(data.items())[:10]:
                    if isinstance(value, dict):
                        value_dict = cast(dict[str, Any], value)
                        html_parts.append(f"{key}: Dict({len(value_dict)} items)\n")
                    elif isinstance(value, list):
                        value_list = cast(list[Any], value)
                        html_parts.append(f"{key}: List({len(value_list)} items)\n")
                    else:
                        html_parts.append(f"{key}: {type(value).__name__}\n")
                html_parts.append("</pre>")
            
            html_parts.append("</div>")
        
        except Exception as e:
            html_parts.append(f"""
            <div style="background-color: #ffebee; padding: 15px; border-left: 4px solid #f44336; margin-bottom: 20px;">
                <strong style="color: #f44336;">❌ 오류: {report_name}</strong><br>
                {str(e)[:200]}
            </div>
            """)
    
    return "".join(html_parts)

def main():
    print("=" * 80)
    print("🎯 실제 추출된 데이터 테이블 생성 중...")
    print("=" * 80)
    
    extracted_data_html = generate_extracted_data_html()
    
    html_content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>실제 추출된 데이터</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            padding: 20px;
            background-color: #f5f5f5;
            line-height: 1.6;
        }}
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #0066cc;
            text-align: center;
            border-bottom: 4px solid #0066cc;
            padding-bottom: 20px;
            margin-bottom: 30px;
            font-size: 28px;
        }}
        h2 {{
            color: #333;
            margin-top: 40px;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #ddd;
        }}
        h3 {{
            color: #0066cc;
            border-bottom: 2px solid #0066cc;
            padding-bottom: 10px;
            margin-top: 20px;
            margin-bottom: 15px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            margin: 15px 0;
            background-color: white;
        }}
        th {{
            background-color: #0066cc;
            color: white;
            padding: 10px;
            text-align: left;
            font-weight: bold;
            border: 1px solid #004499;
        }}
        td {{
            padding: 8px;
            border: 1px solid #ddd;
        }}
        tr:nth-child(even) {{
            background-color: #f9f9f9;
        }}
        tr:hover {{
            background-color: #f0f0f0;
        }}
        pre {{
            background-color: #f5f5f5;
            padding: 10px;
            border-radius: 4px;
            overflow-x: auto;
            font-size: 12px;
        }}
        ul {{
            margin: 15px 0;
            padding-left: 20px;
        }}
        li {{
            margin-bottom: 15px;
            padding: 10px;
            background-color: #fafafa;
            border-radius: 4px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>🎯 실제 추출되고 계산된 데이터</h1>
        
        <div style="background-color: #e8f4f8; padding: 15px; border-left: 4px solid #0066cc; margin-bottom: 30px; border-radius: 4px;">
            <strong>✅ 설명:</strong> 각 보도자료에서 실제로 추출되고 템플릿에 매핑되는 데이터입니다.
            <br>• <strong>요약 보도자료</strong>: 전국 단위 핵심 데이터
            <br>• <strong>부문별 보도자료</strong>: 경제 부문별 상세 데이터 및 통계
        </div>

        <h2>📊 요약 보도자료 데이터</h2>
        {extracted_data_html.split('🏭')[0]}

        <h2>🏭 부문별 보도자료 데이터</h2>
        {extracted_data_html.split('🏭')[1] if '🏭' in extracted_data_html else ''}
        
    </div>
</body>
</html>
"""
    
    output_path = BASE_DIR / "exports" / "extracted_data_tables.html"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("\n" + "=" * 80)
    print(f"✅ 추출된 데이터 문서 생성 완료:")
    print(f"   📄 {output_path}")
    print("=" * 80)

if __name__ == "__main__":
    main()
