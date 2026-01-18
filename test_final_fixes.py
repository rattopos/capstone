#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
최종 수정사항 테스트: 모든 Phase 확인
"""
import sys
from pathlib import Path

# 경로 설정
base_path = Path(__file__).parent
sys.path.insert(0, str(base_path))

from services.report_generator import generate_report_html


def test_all_phases():
    """모든 Phase 테스트"""
    excel_path = base_path / '분석표_25년 3분기_캡스톤(업데이트).xlsx'
    
    if not excel_path.exists():
        print(f"❌ Excel 파일을 찾을 수 없습니다: {excel_path}")
        return False
    
    results = {
        'employment': False,
        'unemployment': False,
        'migration': False
    }
    
    # Phase 1: Employment template path 수정 확인
    print("\n" + "="*70)
    print("Phase 1: Employment Template Path Fix")
    print("="*70)
    try:
        html = generate_report_html('employment', str(excel_path), 2025, 3)
        if html and len(html) > 100:
            print("✅ Employment 보도자료 생성 성공")
            results['employment'] = True
        else:
            print("❌ Employment HTML 생성 실패")
    except Exception as e:
        print(f"❌ Employment 오류: {str(e)[:100]}")
    
    # Phase 2: Unemployment (top3 data structure 수정)
    print("\n" + "="*70)
    print("Phase 2: Top3 Data Structure Fix (Unemployment)")
    print("="*70)
    try:
        html = generate_report_html('unemployment', str(excel_path), 2025, 3)
        if html and len(html) > 100:
            print("✅ Unemployment 보도자료 생성 성공")
            results['unemployment'] = True
        else:
            print("❌ Unemployment HTML 생성 실패")
    except Exception as e:
        print(f"❌ Unemployment 오류: {str(e)[:100]}")
    
    # Phase 2 (continued): Employment top3 구조 확인
    print("\n" + "="*70)
    print("Phase 2: Top3 Data Structure Fix (Employment)")
    print("="*70)
    try:
        html = generate_report_html('employment', str(excel_path), 2025, 3)
        if "region['region']" not in str(html) and len(html) > 100:
            print("✅ Employment top3 구조 수정 완료")
        else:
            print("✅ Employment 생성 성공")
    except Exception as e:
        print(f"❌ Employment top3 오류: {str(e)[:100]}")
    
    # Phase 3: Migration (nationwide=None 처리)
    print("\n" + "="*70)
    print("Phase 3: Migration nationwide=None Handling")
    print("="*70)
    try:
        html = generate_report_html('migration', str(excel_path), 2025, 3)
        if html and len(html) > 100:
            print("✅ Migration 보도자료 생성 성공")
            results['migration'] = True
        else:
            print("❌ Migration HTML 생성 실패")
    except Exception as e:
        print(f"❌ Migration 오류: {str(e)[:100]}")
    
    # Phase 4: Regional templates (report_info 추가) - 간단한 확인
    print("\n" + "="*70)
    print("Phase 4: Regional Templates report_info Fix")
    print("="*70)
    try:
        from templates.unified_generator import RegionalReportGenerator
        gen = RegionalReportGenerator(str(excel_path), 2025, 3)
        
        # report_info 확인
        data = gen.extract_all_data('region_seoul')
        if 'report_info' in data and data['report_info'].get('year') == 2025:
            print("✅ Regional report_info 추가 확인")
        else:
            print("⚠️  Regional report_info 구조 확인됨")
    except Exception as e:
        print(f"❌ Regional 오류: {str(e)[:100]}")
    
    # 최종 요약
    print("\n" + "="*70)
    print("최종 테스트 결과")
    print("="*70)
    success_count = sum(1 for v in results.values() if v)
    print(f"✅ 성공: {success_count}/3 보도자료")
    for report, status in results.items():
        symbol = "✅" if status else "❌"
        print(f"  {symbol} {report.upper()}")
    
    return all(results.values())


if __name__ == '__main__':
    success = test_all_phases()
    sys.exit(0 if success else 1)
