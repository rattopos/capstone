#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""모든 부문의 통합 Generator 테스트"""

from pathlib import Path
from templates.unified_generator import UnifiedReportGenerator
from config.report_configs import REPORT_CONFIGS

excel_path = Path(__file__).parent / '분석표_25년 3분기_캡스톤(업데이트).xlsx'

print("=" * 80)
print("통합 Generator - 전체 부문 테스트")
print("=" * 80)

# 테스트할 부문들
report_types = [
    'mining', 'service', 'consumption',  # 이미 검증된 3개
    'construction', 'export', 'import', 'price', 'employment', 'unemployment', 'migration'  # 나머지 7개
]

results = []

for report_type in report_types:
    print(f"\n{'='*80}")
    print(f"[TEST] {REPORT_CONFIGS[report_type]['name']}")
    print(f"{'='*80}\n")
    
    try:
        generator = UnifiedReportGenerator(report_type, str(excel_path), 2025, 3)
        data = generator.extract_all_data()
        
        # 결과 요약
        nationwide = data['nationwide_data']
        regional = data['regional_data']
        
        result = {
            'type': report_type,
            'name': REPORT_CONFIGS[report_type]['name'],
            'success': True,
            'nationwide_value': nationwide['production_index'],
            'nationwide_rate': nationwide['growth_rate'],
            'increase_count': len(regional['increase_regions']),
            'decrease_count': len(regional['decrease_regions'])
        }
        
        print(f"\n[결과] ✅ 성공")
        print(f"  전국: 지수={nationwide['production_index']:.1f}, 증감률={nationwide['growth_rate']}%")
        print(f"  지역: 증가={result['increase_count']}개, 감소={result['decrease_count']}개")
        
        if regional['increase_regions']:
            top = regional['increase_regions'][0]
            print(f"  최고: {top['region_name']} ({top['change_rate']}%)")
            result['top_region'] = top['region_name']
            result['top_rate'] = top['change_rate']
        
        results.append(result)
        
    except Exception as e:
        print(f"\n[결과] ❌ 실패: {e}")
        results.append({
            'type': report_type,
            'name': REPORT_CONFIGS[report_type]['name'],
            'success': False,
            'error': str(e)
        })

# 최종 요약
print("\n" + "=" * 80)
print("최종 요약")
print("=" * 80)

success_count = sum(1 for r in results if r['success'])
fail_count = len(results) - success_count

print(f"\n총 {len(results)}개 부문: ✅ 성공 {success_count}개, ❌ 실패 {fail_count}개\n")

if success_count > 0:
    print("성공한 부문:")
    for r in results:
        if r['success']:
            rate = r.get('nationwide_rate', 0)
            sign = "+" if rate >= 0 else ""
            print(f"  ✅ {r['name']:12s}: 전국 {sign}{rate}% (증가 {r['increase_count']:2d}, 감소 {r['decrease_count']:2d})")

if fail_count > 0:
    print("\n실패한 부문:")
    for r in results:
        if not r['success']:
            print(f"  ❌ {r['name']:12s}: {r['error'][:60]}...")
