#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
모든 리포트 타입에서 이름 기반 동적 탐색 테스트
"""
import sys
import os
from pathlib import Path

# 프로젝트 루트 경로 추가
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from templates.unified_generator import UnifiedReportGenerator

def test_report_type(report_type: str, excel_file: str):
    """리포트 타입 테스트"""
    print(f"\n{'='*80}")
    print(f"[테스트 시작] {report_type}")
    print(f"{'='*80}")
    
    try:
        # Generator 생성 및 데이터 로드
        gen = UnifiedReportGenerator(
            report_type=report_type,
            excel_path=excel_file,
            year=2025,
            quarter=3
        )
        
        # 데이터 로드
        gen.load_data()
        
        # 탐지 결과 출력
        print(f"\n[컬럼 탐지 결과]")
        print(f"  - Target 컬럼: {gen.target_col}")
        print(f"  - 전년 컬럼: {gen.prev_y_col}")
        print(f"  - 지역명 컬럼: {gen.region_name_col}")
        print(f"  - 업종명 컬럼: {gen.industry_name_col}")
        print(f"  - 데이터 시작 행: {gen.data_start_row}")
        
        # 테이블 데이터 추출
        table_data = gen._extract_table_data_ssot()
        print(f"\n[테이블 데이터 추출]")
        print(f"  - 추출된 지역 수: {len(table_data)}")
        if table_data:
            print(f"  - 첫 번째 지역 (예시): {table_data[0]}")
        
        # 전국 데이터
        nationwide = gen.extract_nationwide_data(table_data)
        print(f"\n[전국 데이터]")
        if nationwide:
            print(f"  - 지역명: {nationwide.get('region_name', 'N/A')}")
            print(f"  - 지수: {nationwide.get('production_index', nationwide.get('index', nationwide.get('value', 'N/A')))}")
            print(f"  - 증감률: {nationwide.get('growth_rate', nationwide.get('change_rate', 'N/A'))}")
        
        # 업종 데이터 (전국)
        if gen.industry_name_col is not None:
            industries = gen._extract_industry_data('전국')
            print(f"\n[업종 데이터 (전국)]")
            print(f"  - 추출된 업종 수: {len(industries)}")
            if industries and len(industries) > 0:
                print(f"  - 첫 번째 업종 (예시): {industries[0]}")
        
        print(f"\n✅ [{report_type}] 테스트 성공!")
        return True
        
    except Exception as e:
        print(f"\n❌ [{report_type}] 테스트 실패: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """메인 테스트"""
    excel_file = "분석표_25년 3분기_캡스톤(업데이트).xlsx"
    
    # 테스트할 리포트 타입 목록
    report_types = [
        'manufacturing',    # 광공업생산
        'service',          # 서비스업생산
        'consumption',      # 소비동향
        'construction',     # 건설동향
        'export',           # 수출
        'import',           # 수입
        'price',            # 물가동향
        'employment',       # 고용률
        'unemployment',     # 실업률
        'migration',        # 국내인구이동
    ]
    
    results = {}
    
    for report_type in report_types:
        results[report_type] = test_report_type(report_type, excel_file)
    
    # 결과 요약
    print(f"\n\n{'='*80}")
    print(f"[테스트 결과 요약]")
    print(f"{'='*80}")
    
    success_count = sum(1 for v in results.values() if v)
    total_count = len(results)
    
    for report_type, success in results.items():
        status = "✅ 성공" if success else "❌ 실패"
        print(f"  {status} - {report_type}")
    
    print(f"\n[전체 결과] {success_count}/{total_count} 성공")
    
    if success_count == total_count:
        print(f"✅ 모든 리포트 타입에서 이름 기반 동적 탐색 성공!")
        return 0
    else:
        print(f"⚠️ 일부 리포트 타입 실패: {total_count - success_count}개")
        return 1


if __name__ == '__main__':
    sys.exit(main())
