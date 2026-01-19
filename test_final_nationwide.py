"""전국 데이터 최종 확인 - 깔끔한 테이블만 출력"""
import logging
import os
import warnings

# 모든 로깅 및 경고 완전 차단
logging.disable(logging.CRITICAL)
os.environ['PYTHONWARNINGS'] = 'ignore'
warnings.filterwarnings('ignore')

import sys
from config.reports import SECTOR_REPORTS
from templates.unified_generator import UnifiedReportGenerator

excel_path = 'uploads/분석표_25년_3분기_캡스톤업데이트_f1da33c3.xlsx'
year = 2025
quarter = 3

print("\n=== 2025년 3분기 전국 데이터 요약 ===")
print(f"{'부문':<15}{'현재값':>10}{'전년값':>10}{'증감/p':>10}{'단위':<10}")
print("="*60)

# SECTOR_REPORTS를 딕셔너리로 변환
reports_dict = {r['report_id']: r for r in SECTOR_REPORTS}

for sector_id in ['manufacturing', 'service', 'consumption', 'construction', 
                  'export', 'import', 'price', 'employment', 'unemployment', 'migration']:
    try:
        # UnifiedReportGenerator는 report_id 문자열을 받음
        gen = UnifiedReportGenerator(sector_id, excel_path, year, quarter)
        
        # 전국 데이터 조회
        nationwide = None
        for row in gen.data:
            if row.get('region') in ['전국', '전체', '합계']:
                nationwide = row
                break
        
        if nationwide:
            # 설정 정보 가져오기
            report_info = reports_dict.get(sector_id, {})
            display_name = report_info.get('name', sector_id)
            report_type = report_info.get('category', '')
            
            # 단위 결정
            if report_type == 'migration':
                unit = '명'
            elif 'rate' in sector_id or 'employment' in sector_id or 'unemployment' in sector_id:
                unit = '%' if 'rate' in sector_id else '천명'
            else:
                unit = '지수'
            
            # 증감 표시
            change_value = nationwide.get('rate_change', 0) or nationwide.get('abs_change', 0)
            if report_type == 'migration':
                change_str = f"{change_value:,.1f}명"
            elif change_value > 0:
                change_str = f"+{change_value:.1f}{'p' if 'rate' in sector_id else '%'}"
            elif change_value < 0:
                change_str = f"{change_value:.1f}{'p' if 'rate' in sector_id else '%'}"
            else:
                change_str = "0.0"
            
            current = nationwide.get('current_value', 0)
            previous = nationwide.get('prev_value', 0)
            
            print(f"{display_name:<15}{current:>10.1f}{previous:>10.1f}{change_str:>10}{unit:<10}")
    
    except Exception as e:
        pass  # 오류 무시

print("="*60)
print("✅ 전국 데이터 요약 완료\n")
