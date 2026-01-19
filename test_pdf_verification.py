"""PDF 대비 전국 데이터 최종 검증"""

import sys
sys.path.insert(0, '/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone')

from templates.unified_generator import UnifiedReportGenerator
from config.reports import SECTOR_REPORTS

excel_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/분석표_25년 3분기_캡스톤(업데이트).xlsx"
year, quarter = 2025, 3

print("=" * 80)
print("2025년 3분기 PDF 대비 전국 데이터 검증")
print("=" * 80)

# PDF 기준값
pdf_values = {
    'manufacturing': {'name': '광공업생산', 'pdf_rate': 5.8, 'pdf_current': 115.2, 'pdf_prev': 108.9},
    'service': {'name': '서비스업생산', 'pdf_rate': 3.1, 'pdf_current': 119.2, 'pdf_prev': 115.6},
    'consumption': {'name': '소비동향', 'pdf_rate': 3.5, 'pdf_current': 105.5, 'pdf_prev': None},  # Excel 기준
    'construction': {'name': '건설수주', 'pdf_rate': 26.5, 'pdf_current': None, 'pdf_prev': None},
    'export': {'name': '수출', 'pdf_rate': 6.5, 'pdf_current': None, 'pdf_prev': None},
    'import': {'name': '수입', 'pdf_rate': 1.5, 'pdf_current': None, 'pdf_prev': None},
    'price': {'name': '물가동향', 'pdf_rate': 2.0, 'pdf_current': 116.7, 'pdf_prev': 114.4},
    'employment': {'name': '고용률', 'pdf_rate': 0.2, 'pdf_current': 63.5, 'pdf_prev': 63.3},
    'unemployment': {'name': '실업자수', 'pdf_rate': None, 'pdf_current': 650.7, 'pdf_prev': 641.0},
}

print(f"\n{'부문':<15} {'PDF 증감률':>10} {'추출 증감률':>12} {'PDF 현재':>10} {'추출 현재':>12} {'결과':<8}")
print("=" * 80)

total = 0
matched = 0

for report in SECTOR_REPORTS:
    sector_id = report.get('report_id')
    name = report.get('name', sector_id)
    
    if sector_id not in pdf_values:
        continue
    
    total += 1
    pdf_data = pdf_values[sector_id]
    
    try:
        gen = UnifiedReportGenerator(sector_id, excel_path, year, quarter)
        
        # 전국 데이터 찾기
        nationwide = None
        for row in gen.data:
            region = str(row.get('region', ''))
            if region in ['전국', '전체', '합계']:
                nationwide = row
                break
        
        if nationwide:
            extracted_current = nationwide.get('current_value', '')
            extracted_rate = nationwide.get('change_rate', '')
            
            # 비교
            pdf_rate = pdf_data.get('pdf_rate')
            pdf_current = pdf_data.get('pdf_current')
            
            rate_match = ''
            current_match = ''
            overall_match = True
            
            if pdf_rate is not None and extracted_rate:
                try:
                    ext_rate = float(str(extracted_rate).replace('%', ''))
                    if abs(ext_rate - pdf_rate) < 0.2:  # 0.2% 이내 허용
                        rate_match = '✓'
                    else:
                        rate_match = f'✗({abs(ext_rate - pdf_rate):.1f}차이)'
                        overall_match = False
                except:
                    rate_match = '?'
                    overall_match = False
            
            if pdf_current is not None and extracted_current:
                try:
                    ext_current = float(str(extracted_current))
                    if abs(ext_current - pdf_current) < 0.2:  # 0.2 이내 허용
                        current_match = '✓'
                    else:
                        current_match = f'✗({abs(ext_current - pdf_current):.1f}차이)'
                        overall_match = False
                except:
                    current_match = '?'
                    overall_match = False
            
            result = "✅ 일치" if overall_match else "⚠️ 불일치"
            if overall_match:
                matched += 1
            
            print(f"{name:<15} {pdf_rate if pdf_rate else '-':>10} {str(extracted_rate):>12} "
                  f"{pdf_current if pdf_current else '-':>10} {str(extracted_current):>12} "
                  f"{result:<8} {rate_match} {current_match}")
        else:
            print(f"{name:<15} {'-':>10} {'전국없음':>12} {'-':>10} {'-':>12} ⚠️")
            
    except Exception as e:
        print(f"{name:<15} {'-':>10} {'ERROR':>12} {'-':>10} {'-':>12} ❌ {str(e)[:20]}")

print("=" * 80)
print(f"\n✅ 검증 완료: {matched}/{total}개 부문 일치")
print(f"📊 일치율: {(matched/total*100):.1f}%\n")
