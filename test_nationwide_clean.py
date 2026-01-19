#!/usr/bin/env python3
"""전국 데이터 요약 (로그 억제)"""
import sys
import os
import logging

# 로그 완전 억제
logging.disable(logging.CRITICAL)
os.environ['PYTHONWARNINGS'] = 'ignore'

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from templates.unified_generator import UnifiedReportGenerator
from config.reports import SECTOR_REPORTS

# SECTOR_REPORTS를 딕셔너리로 변환
reports_dict = {r['report_id']: r for r in SECTOR_REPORTS}

# 10개 부문
sectors = [
    'manufacturing',  # 광공업생산
    'service',        # 서비스업생산
    'consumption',    # 소비동향
    'construction',   # 건설동향
    'export',         # 수출
    'import',         # 수입
    'price',          # 물가동향
    'employment',     # 고용률
    'unemployment',   # 실업률
    'migration'       # 국내인구이동
]

# 전국 데이터 추출
excel_path = "uploads/분석표_25년_3분기_캡스톤업데이트_f1da33c3.xlsx"
year = 2025
quarter = 3

print(f"\n=== {year}년 {quarter}분기 전국 데이터 요약 ===\n")
print(f"{'부문':<12} {'현재값':<10} {'전년값':<10} {'증감/p':<10} {'단위':<10}")
print("-" * 60)

for sector_id in sectors:
    report_info = reports_dict.get(sector_id)
    if not report_info:
        continue
    
    try:
        gen = UnifiedReportGenerator(sector_id, excel_path, year, quarter)
        result = gen.extract_all_data()
        
        if not result or 'regions' not in result:
            continue
        
        # 전국 데이터 찾기
        nationwide = None
        for region in result['regions']:
            if region.get('region_name') == '전국':
                nationwide = region
                break
        
        if not nationwide:
            continue
        
        # 값 포맷팅
        value = nationwide.get('value', 0)
        prev_value = nationwide.get('prev_value', 0)
        
        # report_type 추출
        report_type = report_info.get('category', '')
        
        # 증감률/증감 포인트 계산
        if report_type in ['trade', 'construction']:
            # 무역: 증감률 (%)
            if prev_value != 0:
                change = ((value - prev_value) / abs(prev_value)) * 100
                change_str = f"{change:+.1f}%"
            else:
                change_str = "N/A"
        elif report_type in ['employment', 'unemployment']:
            # 고용/실업: 증감 포인트 (p)
            change = value - prev_value
            change_str = f"{change:+.1f}p"
        elif report_type == 'price':
            # 물가: 증감률 (%)
            if prev_value != 0:
                change = ((value - prev_value) / abs(prev_value)) * 100
                change_str = f"{change:+.1f}%"
            else:
                change_str = "N/A"
        elif report_type == 'migration':
            # 인구이동: 증감 (명)
            change = value - prev_value
            change_str = f"{change:+.0f}명"
        else:
            # 생산/소비: 증감률 (%)
            if prev_value != 0:
                change = ((value - prev_value) / abs(prev_value)) * 100
                change_str = f"{change:+.1f}%"
            else:
                change_str = "N/A"
        
        # 단위
        if report_type in ['production', 'service', 'price']:
            unit = "지수"
        elif report_type == 'trade':
            unit = "백만불"
        elif report_type in ['employment', 'unemployment']:
            unit = "천명"
        elif report_type == 'migration':
            unit = "명"
        elif report_type == 'construction':
            unit = "%"
        else:
            unit = "지수"
        
        # 출력
        display_name = report_info.get('name', sector_id)
        print(f"{display_name:<12} {value:<10.1f} {prev_value:<10.1f} {change_str:<10} {unit:<10}")
    except Exception:
        pass  # 에러 무시

print("\n✅ 전국 데이터 요약 완료")
