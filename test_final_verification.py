"""PDF 대비 전국 데이터 최종 검증 - 간단 버전"""

import sys
import os
import logging
import io
from contextlib import redirect_stdout
sys.path.insert(0, '/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone')

# 로깅 완전 비활성화
logging.disable(logging.CRITICAL)
os.environ['PYTHONWARNINGS'] = 'ignore'

from templates.unified_generator import UnifiedReportGenerator
from config.reports import SECTOR_REPORTS

excel_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/분석표_25년 3분기_캡스톤(업데이트).xlsx"
year, quarter = 2025, 3

# PDF 기준값 (보도자료 기준)
pdf_values = {
    'manufacturing': {'name': '광공업생산', 'rate': 5.8},
    'service': {'name': '서비스업생산', 'rate': 3.1},
    'consumption': {'name': '소비동향', 'rate': 1.5},  # C 분석 시트 공식 전년동기대비 증감률
    'construction': {'name': '건설동향', 'rate': 26.5},
    'export': {'name': '수출', 'rate': 6.5},
    'import': {'name': '수입', 'rate': 1.5},
    'price': {'name': '물가동향', 'rate': 2.0},
    'employment': {'name': '고용률', 'rate': 0.2},
}

print("\n" + "="*70)
print(" 📊 2025년 3분기 PDF 대비 전국 데이터 검증")
print("="*70)
print(f"\n{'부문':<15} {'PDF 증감률':>12} {'추출 증감률':>14} {'차이':>10} {'결과':<10}")
print("-"*70)

matched = 0
total = 0

for report in SECTOR_REPORTS:
    sector_id = report.get('report_id')
    if sector_id not in pdf_values:
        continue
    
    total += 1
    pdf_data = pdf_values[sector_id]
    name = pdf_data['name']
    pdf_rate = pdf_data['rate']
    
    try:
        buf = io.StringIO()
        with redirect_stdout(buf):
            gen = UnifiedReportGenerator(sector_id, excel_path, year, quarter)
            data = gen.extract_all_data()
        table = data.get('table_data', []) if isinstance(data, dict) else []

        nationwide = next((row for row in table if str(row.get('region_name', '')).strip() in ['전국', '전체', '합계']), None)
        if not nationwide:
            print(f"{name:<15} {pdf_rate:>12.1f}% {'전국없음':>14} {'-':>10} ⚠️")
            continue

        extracted_rate = nationwide.get('change_rate')
        if extracted_rate is None or extracted_rate == '-':
            print(f"{name:<15} {pdf_rate:>12.1f}% {'없음':>14} {'-':>10} ⚠️")
            continue

        try:
            extracted_rate_val = float(str(extracted_rate).replace('%', '').strip())
        except Exception:
            print(f"{name:<15} {pdf_rate:>12.1f}% {'변환실패':>14} {'-':>10} ⚠️")
            continue

        diff = abs(extracted_rate_val - pdf_rate)
        if diff < 0.2:
            result = "✅ 일치"
            matched += 1
        else:
            result = "⚠️ 불일치"

        print(f"{name:<15} {pdf_rate:>12.1f}% {extracted_rate_val:>13.1f}% {diff:>9.1f}% {result:<10}")

    except Exception:
        print(f"{name:<15} {pdf_rate:>12.1f}% {'ERROR':>14} {'-':>10} ❌")

print("-"*70)
print(f"\n ✅ 검증 결과: {matched}/{total}개 부문 일치 ({matched/total*100:.1f}%)")
print("="*70)
print()

# migration은 별도 확인 (전국 데이터 없어야 함)
print("\n📍 국내인구이동 전국 데이터 제외 확인:")
try:
    buf = io.StringIO()
    with redirect_stdout(buf):
        gen = UnifiedReportGenerator('migration', excel_path, year, quarter)
        data = gen.extract_all_data()
    
    table = data.get('table_data', []) if isinstance(data, dict) else []
    has_nationwide = any(
        str(row.get('region_name', '')).strip() in ['전국', '전체', '합계']
        for row in table
    )
    
    if not has_nationwide:
        print("  ✅ 전국 데이터 없음 (정상)")
    else:
        print("  ⚠️ 전국 데이터 존재 (비정상)")
        
except Exception as e:
    print(f"  ❌ 확인 실패: {str(e)[:50]}")

print()
