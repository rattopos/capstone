#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
템플릿 형식의 데이터 테이블 생성기
- 실제 추출된 데이터를 템플릿과 동일한 구조로 정렬
- 중복 컬럼 제거 (growth_rate vs change_rate)
- 현재분기 | 직전분기 | 작년동분기 구조
"""

import sys
import json
from pathlib import Path
from typing import Any, cast

# 프로젝트 경로 설정
sys.path.insert(0, str(Path(__file__).parent))

from report_generator import ReportGenerator


class DataTableBuilder:
    """템플릿 형식의 데이터 테이블 생성"""
    
    def __init__(self, excel_path: str) -> None:
        self.generator = ReportGenerator(excel_path)
        self.excel_path = excel_path
        
    def format_change_value(self, value: Any) -> str:
        """변화율 포맷팅 (소수점 1자리)"""
        if value is None or value == '' or value == '-':
            return '-'
        try:
            v = float(value)
            return f"{v:.1f}"
        except:
            return str(value)
    
    def format_index_value(self, value: Any) -> str:
        """지수 포맷팅 (소수점 1자리)"""
        if value is None or value == '' or value == '-':
            return '-'
        try:
            v = float(value)
            return f"{v:.1f}"
        except:
            return str(value)
    
    def extract_sector_report(self, report_id: str) -> dict[str, Any] | None:
        """부문별 보도자료 데이터 추출"""
        print(f"\n{'='*60}")
        print(f"보도자료: {report_id}")
        print(f"{'='*60}")
        
        try:
            data: dict[str, Any] = self.generator.extract_data(report_id)
            
            # 데이터 구조 파악
            print(f"\n📊 추출된 데이터 구조:")
            print(json.dumps({k: type(v).__name__ for k, v in data.items()}, indent=2, ensure_ascii=False))
            
            # regional_data 구조 확인
            if 'regional_data' in data:
                rd = data['regional_data']
                print(f"\n📍 Regional Data 타입: {type(rd).__name__}")
                if isinstance(rd, dict):
                    rd_dict = cast(dict[str, Any], rd)
                    print(f"   Keys: {list(rd_dict.keys())[:5]}...")  # 처음 5개만
                    if rd_dict:
                        first_key = list(rd_dict.keys())[0]
                        print(f"   First region '{first_key}': {type(rd_dict[first_key]).__name__}")
                        if isinstance(rd_dict[first_key], dict):
                            first_region = cast(dict[str, Any], rd_dict[first_key])
                            print(f"   - Fields: {list(first_region.keys())}")
                            print(f"   - Sample: {first_region}")
                elif isinstance(rd, list):
                    rd_list = cast(list[Any], rd)
                    print(f"   Total regions: {len(rd_list)}")
                    if rd_list:
                        print(f"   First region: {rd_list[0]}")
            
            # summary_box 구조 확인
            if 'summary_box' in data:
                print(f"\n📦 Summary Box: {data['summary_box']}")
            
            # table_data 구조 확인
            if 'table_data' in data:
                print(f"\n📋 Table Data 구조: {type(data['table_data']).__name__}")
                if isinstance(data['table_data'], dict):
                    table_dict = cast(dict[str, Any], data['table_data'])
                    print(f"   Keys: {list(table_dict.keys())}")
                    for key in list(table_dict.keys())[:3]:
                        val = table_dict[key]
                        print(f"   - {key}: {type(val).__name__}")
                        if isinstance(val, list) and val:
                            print(f"     First item: {val[0]}")
            
            # 테이블 생성을 위한 데이터 정렬
            summary_table = self._build_summary_table(data, report_id)
            
            return {
                'report_id': report_id,
                'raw_data': data,
                'summary_table': summary_table
            }
            
        except Exception as e:
            print(f"❌ 오류: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
    
    def _build_summary_table(self, data: dict[str, Any], report_id: str) -> dict[str, Any]:
        """요약 테이블 데이터 구성"""
        
        table: dict[str, Any] = {
            'columns': {},
            'regions': []
        }
        
        # regional_data 구조는 report마다 다름
        # 예: manufacturing는 dict, consumption도 dict
        # 각 regional_data에서 'all_regions' 확인
        
        regional_data: Any = data.get('regional_data', {})
        
        # 모든 지역 데이터 추출
        all_regions_list: list[dict[str, Any]] = []
        
        if isinstance(regional_data, dict):
            regional_dict = cast(dict[str, Any], regional_data)
            # dict 구조: all_regions, increase_regions, decrease_regions 등
            if 'all_regions' in regional_dict:
                all_regions_list = cast(list[dict[str, Any]], regional_dict['all_regions'])
            elif 'regions' in regional_dict:
                all_regions_list = cast(list[dict[str, Any]], regional_dict['regions'])
            else:
                # increase_regions + decrease_regions 합치기
                increase = cast(list[dict[str, Any]], regional_dict.get('increase_regions', []))
                decrease = cast(list[dict[str, Any]], regional_dict.get('decrease_regions', []))
                all_regions_list = increase + decrease
        elif isinstance(regional_data, list):
            all_regions_list = cast(list[dict[str, Any]], regional_data)
        
        # 지역별 행 생성
        for region_data in all_regions_list:
            region_name = region_data.get('region', region_data.get('name', ''))
            if not region_name:
                continue
            
            row: dict[str, Any] = {
                'region': region_name,
                'group': region_data.get('group'),
                'rowspan': region_data.get('rowspan')
            }
            
            # 증감률 또는 성장률 추출 (중복 제거)
            growth_rates: list[str] = []
            if 'growth_rate' in region_data or 'change_rate' in region_data:
                # 생산 지수 시리즈 (광공업, 서비스 등)
                val = region_data.get('growth_rate') or region_data.get('change_rate')
                growth_rates.append(self.format_change_value(val))
                growth_rates.append(self.format_change_value(region_data.get('previous_quarter_growth')))
                growth_rates.append(self.format_change_value(region_data.get('previous_year_growth')))
                growth_rates.append(self.format_change_value(region_data.get('previous_year_same_quarter_growth')))
                
                # 값이 있으면 추가
                if any(g != '-' for g in growth_rates):
                    row['growth_rates'] = growth_rates
            
            # 지수 추출
            indices: list[str] = []
            if 'index' in region_data or 'current_value' in region_data:
                indices.append(self.format_index_value(region_data.get('index') or region_data.get('current_value')))
                indices.append(self.format_index_value(region_data.get('previous_year_index') or region_data.get('previous_year_value')))
                
                # 값이 있으면 추가
                if any(i != '-' for i in indices):
                    row['indices'] = indices
            
            # 고용률 추출
            rates: list[str] = []
            if 'rate' in region_data or 'employment_rate' in region_data:
                rates.append(self.format_change_value(region_data.get('rate') or region_data.get('employment_rate')))
                rates.append(self.format_change_value(region_data.get('previous_quarter_rate')))
                rates.append(self.format_change_value(region_data.get('previous_year_rate')))
                
                # 값이 있으면 추가
                if any(r != '-' for r in rates):
                    row['rates'] = rates
            
            table['regions'].append(row)
        
        # 컬럼 헤더 설정
        if any('growth_rates' in r for r in table['regions']):
            table['columns']['growth_rate_columns'] = [
                '2025년 3분기',
                '2025년 2분기',
                '2024년 3분기',
                '전년동분기대비'
            ]
            table['columns']['index_columns'] = [
                '2025년 3분기',
                '2024년 3분기'
            ]
            table['base_year'] = '2020'
        
        if any('rates' in r for r in table['regions']):
            table['columns']['change_columns'] = [
                '2025년 3분기',
                '2025년 2분기',
                '2024년 3분기',
                '전년동분기대비'
            ]
            table['columns']['rate_columns'] = [
                '2025년 3분기',
                '2025년 2분기',
                '2024년 3분기'
            ]
        
        return table
    
    def generate_html_preview(self, sector_result: dict[str, Any]) -> str:
        """HTML 미리보기 생성"""
        
        report_id = sector_result['report_id']
        summary_table = sector_result['summary_table']
        
        html = f"""
<html>
<head>
    <meta charset="utf-8">
    <title>{report_id} - 데이터 테이블</title>
    <style>
        body {{ font-family: '맑은 고딕', sans-serif; margin: 20px; }}
        .report-title {{ font-size: 18pt; font-weight: bold; margin-bottom: 10px; }}
        .table-section {{ margin-top: 20px; }}
        .section-title {{ font-size: 12pt; font-weight: bold; margin: 10px 0 5px 0; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
        th, td {{ border: 1px solid #333; padding: 8px; text-align: center; }}
        th {{ background-color: #e6e0ec; }}
        .region-group {{ background-color: #f5f5f5; }}
        .index-section {{ background-color: #fef9e7; }}
        .rate-section {{ background-color: #fef9e7; }}
    </style>
</head>
<body>
    <div class="report-title">{report_id} - 데이터 테이블</div>
    
    <div class="table-section">
        <div class="section-title">추출된 컬럼 구조</div>
        <pre>{json.dumps(summary_table['columns'], indent=2, ensure_ascii=False)}</pre>
    </div>
    
    <div class="table-section">
        <div class="section-title">데이터 테이블</div>
        <table>
            <thead>
                <tr>
                    <th>지역</th>
        """
        
        # 컬럼 헤더 추가
        if 'growth_rate_columns' in summary_table['columns']:
            for col in summary_table['columns']['growth_rate_columns']:
                html += f"<th>{col}</th>"
            for col in summary_table['columns']['index_columns']:
                html += f'<th class="index-section">{col}</th>'
        
        if 'change_columns' in summary_table['columns']:
            for col in summary_table['columns']['change_columns']:
                html += f"<th>{col}</th>"
            for col in summary_table['columns']['rate_columns']:
                html += f'<th class="rate-section">{col}</th>'
        
        html += """
                </tr>
            </thead>
            <tbody>
        """
        
        # 데이터 행 추가
        for row in summary_table['regions']:
            html += f"<tr><td>{row['region']}</td>"
            
            if 'growth_rates' in row:
                for val in row['growth_rates']:
                    html += f"<td>{val}</td>"
                for val in row['indices']:
                    html += f'<td class="index-section">{val}</td>'
            
            if 'changes' in row:
                for val in row['changes']:
                    html += f"<td>{val}</td>"
                for val in row['rates']:
                    html += f'<td class="rate-section">{val}</td>'
            
            html += "</tr>"
        
        html += """
            </tbody>
        </table>
    </div>
    
</body>
</html>
        """
        
        return html


def main():
    excel_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/분석표_25년 3분기_캡스톤(업데이트).xlsx"
    
    builder = DataTableBuilder(excel_path)
    
    # 몇 가지 부문별 보도자료만 추출 테스트
    test_reports = ['manufacturing', 'service', 'consumption', 'construction']
    
    results: dict[str, dict[str, Any]] = {}
    for report_id in test_reports:
        result = builder.extract_sector_report(report_id)
        if result:
            results[report_id] = result
            
            # 각 보도자료별 HTML 미리보기 생성
            html_preview = builder.generate_html_preview(result)
            output_file = f"/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/exports/table_preview_{report_id}.html"
            Path(output_file).parent.mkdir(parents=True, exist_ok=True)
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_preview)
            print(f"✅ 미리보기 저장: {output_file}")
    
    # 요약 통계 출력
    print(f"\n{'='*60}")
    print(f"📊 추출 요약")
    print(f"{'='*60}")
    print(f"성공한 보도자료: {len(results)}/{len(test_reports)}")
    
    # 각 보도자료의 테이블 통계
    for report_id, result in results.items():
        summary_table = cast(dict[str, Any], result['summary_table'])
        region_count = len(cast(list[Any], summary_table.get('regions', [])))
        print(f"\n{report_id}:")
        print(f"  - 지역 개수: {region_count}")
        columns = cast(dict[str, Any], summary_table.get('columns', {}))
        print(f"  - 컬럼: {list(columns.keys())}")


if __name__ == '__main__':
    main()
