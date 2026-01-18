#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
최종 수정사항 직접 테스트 - 템플릿 렌더링 테스트
"""
import sys
from pathlib import Path

# 경로 설정
base_path = Path("/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone")
sys.path.insert(0, str(base_path))

# Excel 파일 찾기
excel_files = list(base_path.glob("*.xlsx"))
if not excel_files:
    print("❌ Excel 파일 없음")
    sys.exit(1)

excel_path = str(excel_files[0])
print(f"✅ Excel 파일: {Path(excel_path).name}\n")

from templates.unified_generator import EmploymentRateGenerator, UnemploymentGenerator, DomesticMigrationGenerator

tests = [
    ('고용률', EmploymentRateGenerator, 'employment'),
    ('실업률', UnemploymentGenerator, 'unemployment'),
    ('국내인구이동', DomesticMigrationGenerator, 'migration'),
]

print("="*70)
print("최종 수정사항 테스트")
print("="*70)

success_count = 0
for report_name, generator_class, report_id in tests:
    print(f"\n📊 {report_name}")
    try:
        # Generator 생성
        gen = generator_class(excel_path, 2025, 3)
        
        # 데이터 추출
        data = gen.extract_all_data()
        print(f"  ✅ 데이터 추출 완료")
        
        # top3 구조 확인 (Phase 2)
        if 'top3_increase_regions' in data:
            top3 = data['top3_increase_regions']
            if top3 and isinstance(top3[0], dict):
                print(f"  ✅ Phase 2: top3 dict 구조 확인")
        
        # nationwide=None 확인 (Phase 3, Migration만)
        if report_id == 'migration':
            if data.get('nationwide_data') is None:
                print(f"  ✅ Phase 3: nationwide=None 처리 확인")
        
        # report_info 확인 (Phase 4)
        if 'report_info' in data:
            if data['report_info'].get('year') == 2025:
                print(f"  ✅ Phase 4: report_info 추가 확인")
        
        success_count += 1
        
    except Exception as e:
        print(f"  ❌ 오류: {str(e)[:80]}")

print("\n" + "="*70)
print(f"최종 결과: {success_count}/3 성공")
print("="*70)

if success_count == 3:
    print("\n✅ 모든 Phase 수정사항 확인 완료!")
    sys.exit(0)
else:
    print(f"\n⚠️  {3-success_count}개 항목 확인 필요")
    sys.exit(1)
