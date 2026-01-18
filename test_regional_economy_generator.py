#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
시도별 경제동향 통합 보고서 생성 테스트

unified_generator.py의 RegionalEconomyByRegionGenerator를 테스트합니다.
"""

import sys
from pathlib import Path

# 경로 설정
base_path = Path(__file__).parent
sys.path.insert(0, str(base_path))

def test_regional_economy_generator():
    """시도별 경제동향 Generator 테스트"""
    from templates.unified_generator import RegionalEconomyByRegionGenerator
    from config.report_configs import REPORT_CONFIGS
    
    # 테스트 엑셀 파일 경로
    excel_files = list(base_path.glob('*분석표*.xlsx'))
    if not excel_files:
        print("❌ 분석표 엑셀 파일을 찾을 수 없습니다.")
        return False
    
    excel_path = str(excel_files[0])
    print(f"✅ 엑셀 파일: {excel_path}")
    
    try:
        # Generator 생성
        gen = RegionalEconomyByRegionGenerator(excel_path, year=2025, quarter=3)
        print("✅ RegionalEconomyByRegionGenerator 초기화 완료")
        
        # 설정 확인
        config = REPORT_CONFIGS.get('regional_economy_by_region')
        if not config:
            print("❌ regional_economy_by_region 설정을 찾을 수 없습니다.")
            return False
        print(f"✅ 설정 확인: {config['name']}")
        
        # 시도 목록 확인
        print(f"\n📍 대상 시도 ({len(gen.REGIONS)}개):")
        for region in gen.REGIONS:
            print(f"  - {region['code']:2d}: {region['full_name']}")
        
        # 서울 데이터로 테스트
        print("\n🧪 서울 데이터 추출 테스트...")
        section = gen.extract_regional_section('서울', 'mining')
        if section:
            print(f"✅ 생산 섹션 추출 완료")
            if section.get('narrative'):
                print(f"   나레이션: {section['narrative'][:80]}...")
        else:
            print("⚠️ 생산 섹션을 찾을 수 없습니다.")
        
        print("\n✅ 모든 테스트 완료!")
        return True
        
    except Exception as e:
        print(f"❌ 테스트 실패: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_report_config():
    """보고서 설정 확인"""
    from config.report_configs import REPORT_CONFIGS, get_report_config
    
    print("=" * 70)
    print("보고서 설정 확인")
    print("=" * 70)
    
    # regional_economy_by_region 설정 확인
    try:
        config = get_report_config('regional_economy_by_region')
        print(f"\n✅ regional_economy_by_region 설정:")
        print(f"  - 이름: {config['name']}")
        print(f"  - 템플릿: {config['template']}")
        print(f"  - is_regional_by_region: {config.get('is_regional_by_region', False)}")
        print(f"  - require_analysis_sheet: {config.get('require_analysis_sheet', True)}")
        return True
    except Exception as e:
        print(f"❌ 설정 확인 실패: {e}")
        return False


def main():
    """메인 테스트"""
    print("=" * 70)
    print("시도별 경제동향 통합 보고서 생성 테스트")
    print("=" * 70)
    
    # 설정 확인
    if not test_report_config():
        return 1
    
    # Generator 테스트
    if not test_regional_economy_generator():
        return 1
    
    print("\n" + "=" * 70)
    print("✅ 모든 테스트 성공!")
    print("=" * 70)
    return 0


if __name__ == '__main__':
    sys.exit(main())
