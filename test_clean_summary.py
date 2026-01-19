"""전국 데이터 요약 - 완전 깔끔 버전"""
import sys
import os
import io
import logging

# 먼저 로깅 완전 비활성화
logging.disable(logging.CRITICAL)
os.environ['PYTHONWARNINGS'] = 'ignore'

# 표준 출력을 버퍼로 임시 리다이렉트
original_stdout = sys.stdout
original_stderr = sys.stderr
sys.stdout = io.StringIO()
sys.stderr = io.StringIO()

from config.reports import SECTOR_REPORTS
from services.report_generator import UnifiedReportGenerator

# Excel 파일 경로
excel_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/uploads/지역경제동향_2025년_3분기.xlsx"
year = 2025
quarter = 3

# 데이터 수집
results = []
for report in SECTOR_REPORTS:
    sector_id = report.get('report_id')
    name = report.get('name', sector_id)
    
    try:
        gen = UnifiedReportGenerator(sector_id, excel_path, year, quarter)
        
        # 전국 데이터 찾기
        for row in gen.data:
            region = str(row.get('region', ''))
            if region in ['전국', '전체', '합계']:
                current = row.get('current_value', '')
                prev = row.get('prev_year_value', '')
                change = row.get('change_rate', '')
                unit = row.get('unit', '')
                
                results.append({
                    'name': name,
                    'current': current,
                    'prev': prev,
                    'change': change,
                    'unit': unit
                })
                break
    except:
        pass

# 원래 stdout 복구
sys.stdout = original_stdout
sys.stderr = original_stderr

# 깔끔한 표 출력
print("\n=== 2025년 3분기 전국 데이터 요약 ===\n")
print(f"{'부문':<15} {'현재값':>10} {'전년값':>10} {'증감/p':>10} {'단위':<8}")
print("=" * 60)

for r in results:
    print(f"{r['name']:<15} {str(r['current']):>10} {str(r['prev']):>10} {str(r['change']):>10} {r['unit']:<8}")

print("\n✅ 데이터 추출 완료\n")
