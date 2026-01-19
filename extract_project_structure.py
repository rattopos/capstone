#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
프로젝트에서 실제 추출되는 보도자료 데이터 정리
"""
from config.settings import BASE_DIR
from config.reports import SUMMARY_REPORTS, SECTOR_REPORTS, REGIONAL_REPORTS

def generate_report_table_html() -> tuple[str, str, str, str]:
    """보도자료 구조를 HTML 테이블로 생성"""
    
    # 요약 보도자료
    summary_html = '<h2>📊 요약 보도자료 (SUMMARY_REPORTS)</h2>'
    summary_html += '<table border="1"><tr><th>번호</th><th>ID</th><th>이름</th><th>시트</th><th>템플릿</th><th>아이콘</th></tr>'
    for i, report in enumerate(SUMMARY_REPORTS, 1):
        summary_html += f"""<tr>
        <td>{i}</td>
        <td>{report['id']}</td>
        <td>{report['name']}</td>
        <td>{report['sheet']}</td>
        <td>{report['template']}</td>
        <td>{report['icon']}</td>
        </tr>"""
    summary_html += '</table>'
    
    # 부문별 보도자료
    sector_html = '<h2>🏭 부문별 보도자료 (SECTOR_REPORTS)</h2>'
    sector_html += '<table border="1"><tr><th>번호</th><th>ID</th><th>이름</th><th>카테고리</th><th>시트</th><th>집계시트</th><th>아이콘</th></tr>'
    for i, report in enumerate(SECTOR_REPORTS, 1):
        agg_sheet = report.get('aggregation_structure', {}).get('sheet', 'N/A')
        sector_html += f"""<tr>
        <td>{i}</td>
        <td>{report['id']}</td>
        <td>{report['name']}</td>
        <td>{report['category']}</td>
        <td>{report['sheet']}</td>
        <td>{agg_sheet}</td>
        <td>{report['icon']}</td>
        </tr>"""
    sector_html += '</table>'
    
    # 지역별 보도자료
    regional_html = '<h2>🗺️ 지역별 보도자료 (17개 지역)</h2>'
    regional_html += '<table border="1"><tr><th>번호</th><th>ID</th><th>지역명</th><th>전체명</th><th>아이콘</th></tr>'
    for report in REGIONAL_REPORTS:
        regional_html += f"""<tr>
        <td>{report['index']}</td>
        <td>{report['id']}</td>
        <td>{report['name']}</td>
        <td>{report['full_name']}</td>
        <td>{report['icon']}</td>
        </tr>"""
    regional_html += '</table>'
    
    # 전체 수량 정리
    summary_stats = f"""
    <h2>📈 통계</h2>
    <table border="1">
    <tr><td><strong>요약 보도자료</strong></td><td>{len(SUMMARY_REPORTS)}</td></tr>
    <tr><td><strong>부문별 보도자료</strong></td><td>{len(SECTOR_REPORTS)}</td></tr>
    <tr><td><strong>지역별 보도자료</strong></td><td>{len(REGIONAL_REPORTS)}</td></tr>
    <tr><td><strong>합계</strong></td><td>{len(SUMMARY_REPORTS) + len(SECTOR_REPORTS) + len(REGIONAL_REPORTS)}</td></tr>
    </table>
    """
    
    return summary_html, sector_html, regional_html, summary_stats

def main():
    summary_html, sector_html, regional_html, summary_stats = generate_report_table_html()
    
    html_content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>프로젝트 보도자료 구조</title>
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
            margin-top: 30px;
            margin-bottom: 15px;
            padding-bottom: 10px;
            border-bottom: 2px solid #ddd;
            font-size: 20px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 30px;
            font-size: 14px;
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
        .info-box {{
            background-color: #e8f4f8;
            padding: 15px;
            border-left: 4px solid #0066cc;
            margin-bottom: 20px;
            border-radius: 4px;
            font-size: 14px;
        }}
        .info-box strong {{
            color: #0066cc;
        }}
        .category {{
            display: inline-block;
            padding: 4px 8px;
            background-color: #e0e0e0;
            border-radius: 3px;
            font-size: 12px;
            margin-right: 5px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>🎯 지역경제동향 보도자료 생성 프로젝트 구조</h1>
        
        <div class="info-box">
            <strong>✅ 설명:</strong> 이 프로젝트에서는 Excel 파일로부터 <strong>3가지 유형</strong>의 보도자료를 생성합니다.
            <br>1️⃣ <strong>요약 보도자료</strong> - 전국 기준의 핵심 지표 요약
            <br>2️⃣ <strong>부문별 보도자료</strong> - 경제 부문별(생산, 소비, 고용 등) 상세 분석
            <br>3️⃣ <strong>지역별 보도자료</strong> - 17개 시도별 경제동향
        </div>

        {summary_html}
        
        {sector_html}
        
        {regional_html}
        
        {summary_stats}
        
        <h2>📝 설명</h2>
        <div class="info-box">
            <p><strong>• 요약 보도자료 (5개):</strong> 생산, 소비·건설, 수출·물가, 고용·인구, 지역경제동향 등 전국 단위의 요약본</p>
            <p><strong>• 부문별 보도자료 (9개):</strong></p>
            <ul>
                <li>생산: 광공업생산, 서비스업생산</li>
                <li>소비/건설: 소비동향, 건설동향</li>
                <li>무역: 수출, 수입</li>
                <li>물가: 물가동향</li>
                <li>고용: 고용률, 실업률</li>
                <li>인구: 국내인구이동</li>
            </ul>
            <p><strong>• 지역별 보도자료 (17개):</strong> 각 시/도별 경제 현황 분석</p>
        </div>
        
    </div>
</body>
</html>
"""
    
    output_path = BASE_DIR / "exports" / "project_structure.html"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"✅ 프로젝트 구조 문서 생성 완료:")
    print(f"   📄 {output_path}")
    print()
    print(f"📊 요약:")
    print(f"   • 요약 보도자료: {len(SUMMARY_REPORTS)}개")
    print(f"   • 부문별 보도자료: {len(SECTOR_REPORTS)}개")
    print(f"   • 지역별 보도자료: {len(REGIONAL_REPORTS)}개")
    print(f"   • 전체: {len(SUMMARY_REPORTS) + len(SECTOR_REPORTS) + len(REGIONAL_REPORTS)}개 보도자료")

if __name__ == "__main__":
    main()
