#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
실제 데이터 구조 분석 및 정렬된 테이블 생성
현재+직전 4분기+작년 동분기 컬럼으로 구성
"""
from typing import Any, cast

import pandas as pd

from config.settings import BASE_DIR
from config.reports import SECTOR_REPORTS
from report_generator import ReportGenerator

EXCEL_PATH = str((BASE_DIR / "분석표_25년 3분기_캡스톤(업데이트).xlsx").resolve())

def generate_comprehensive_table() -> list[dict[str, Any]]:
    """
    실제 데이터를 분석하고 정렬된 테이블 생성
    광공업생산을 예로 시연
    """
    print("🔄 데이터 분석 중...")
    generator = ReportGenerator(EXCEL_PATH)
    
    # 광공업생산 데이터 추출
    data: dict[str, Any] = generator.extract_data('manufacturing')
    
    print("\n📊 광공업생산 데이터 구조 분석:")
    print(f"  데이터 키: {list(data.keys())}")
    
    # 주요 구조 분석
    for key, value in data.items():
        if isinstance(value, dict):
            value_dict = cast(dict[str, Any], value)
            print(f"  - {key}: Dict ({len(value_dict)} items)")
            if value_dict:
                first_key = list(value_dict.keys())[0]
                print(f"    샘플: {first_key} = {str(value_dict[first_key])[:100]}")
        elif isinstance(value, list):
            value_list = cast(list[Any], value)
            print(f"  - {key}: List ({len(value_list)} items)")
            if value_list and isinstance(value_list[0], dict):
                first_row = cast(dict[str, Any], value_list[0])
                print(f"    구조: {list(first_row.keys())}")
                print(f"    샘플: {first_row}")
        else:
            print(f"  - {key}: {type(value).__name__}")
    
    # 모든 부문별 보도자료 처리
    all_tables: list[dict[str, Any]] = []
    
    for report in SECTOR_REPORTS:
        report_id = report['id']
        report_name = report['name']
        print(f"\n🏭 {report_name} 데이터 추출 중...")
        
        try:
            data = generator.extract_data(report_id)
            
            # 지역 데이터 추출
            if 'regional_data' in data and isinstance(data['regional_data'], dict):
                rows: list[dict[str, Any]] = []
                regional_dict = cast(dict[str, Any], data['regional_data'])
                for region_id, region_data in regional_dict.items():
                    if isinstance(region_data, dict):
                        region_dict = cast(dict[str, Any], region_data)
                        row: dict[str, Any] = {
                            '보도자료': report_name,
                            '지역': region_dict.get('region', region_id),
                            '2025년 3분기': region_dict.get('current_value', region_dict.get('value')),
                            '2025년 2분기': region_dict.get('q2_value'),
                            '2024년 3분기': region_dict.get('yoy_value'),
                            '증감률': region_dict.get('change_rate', region_dict.get('growth_rate')),
                        }
                        rows.append(row)
                
                if rows:
                    df = pd.DataFrame(rows)
                    all_tables.append({
                        'name': report_name,
                        'data': df,
                        'type': '지역별'
                    })
                    print(f"   ✅ 지역 데이터: {len(rows)}개")
        
        except Exception as e:
            print(f"   ❌ 오류: {str(e)[:100]}")
    
    return all_tables

def main():
    print("=" * 80)
    print("🎯 실제 데이터 분석 및 테이블 생성")
    print("=" * 80)
    
    tables = generate_comprehensive_table()
    
    # HTML 생성
    html_content = """<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>실제 추출된 데이터 테이블 (현재+직전+작년)</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        h1 {
            color: #0066cc;
            text-align: center;
            border-bottom: 4px solid #0066cc;
            padding-bottom: 20px;
            margin-bottom: 30px;
        }
        h2 {
            color: #333;
            margin-top: 30px;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #ddd;
        }
        .info-box {
            background-color: #e8f4f8;
            padding: 15px;
            border-left: 4px solid #0066cc;
            margin-bottom: 30px;
            border-radius: 4px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
            font-size: 13px;
        }
        th {
            background-color: #0066cc;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: bold;
            border: 1px solid #004499;
        }
        td {
            padding: 10px;
            border: 1px solid #ddd;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        tr:hover {
            background-color: #f0f0f0;
        }
        .positive {
            color: #d32f2f;
        }
        .negative {
            color: #388e3c;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 실제 추출된 데이터 테이블</h1>
        
        <div class="info-box">
            <strong>✅ 구성:</strong>
            <br>• 컬럼: 보도자료 | 지역 | 2025년 3분기 (현재) | 2025년 2분기 (직전 4분기) | 2024년 3분기 (작년 동분기) | 증감률
            <br>• 중복 제거: growth_rate와 change_rate 통합
            <br>• 전체 데이터: 샘플이 아닌 모든 실제 데이터 포함
        </div>
"""
    
    if tables:
        for table_info in tables:
            html_content += f"""
        <h2>{table_info['name']} ({table_info['type']})</h2>
        {table_info['data'].to_html(index=False, border=1, justify='left', classes='data-table')}
"""
    else:
        html_content += "<p style='color: #f44336;'>⚠️ 추출된 테이블이 없습니다.</p>"
    
    html_content += """
    </div>
</body>
</html>
"""
    
    output_path = BASE_DIR / "exports" / "extracted_data_final.html"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("\n" + "=" * 80)
    print(f"✅ 데이터 테이블 생성 완료:")
    print(f"   📄 {output_path}")
    print(f"   📊 테이블 수: {len(tables)}")
    print("=" * 80)

if __name__ == "__main__":
    main()
