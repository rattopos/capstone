#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
간단한 Flask 앱 동작 테스트
"""
import sys
import os

# 프로젝트 루트 경로 추가
sys.path.insert(0, os.path.dirname(__file__))

def test_imports():
    """모든 필수 모듈 import 테스트"""
    print("\n[테스트 1] 필수 모듈 Import")
    try:
        from flask import Flask
        print("  ✅ Flask import 성공")
        
        from config.settings import BASE_DIR, SECRET_KEY
        print("  ✅ config.settings import 성공")
        
        from utils.filters import register_filters
        print("  ✅ utils.filters import 성공")
        
        from routes import main_bp, api_bp
        print("  ✅ routes import 성공")
        
        from templates.unified_generator import UnifiedReportGenerator
        print("  ✅ UnifiedReportGenerator import 성공")
        
        return True
    except Exception as e:
        print(f"  ❌ Import 실패: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_app_creation():
    """Flask 앱 생성 테스트"""
    print("\n[테스트 2] Flask 앱 생성")
    try:
        from app import create_app
        app = create_app()
        print(f"  ✅ Flask 앱 생성 성공")
        print(f"  - 등록된 블루프린트: {[bp.name for bp in app.blueprints.values()]}")
        return True
    except Exception as e:
        print(f"  ❌ 앱 생성 실패: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_unified_generator():
    """UnifiedReportGenerator 기본 기능 테스트"""
    print("\n[테스트 3] UnifiedReportGenerator 기본 기능")
    try:
        from templates.unified_generator import UnifiedReportGenerator
        
        excel_file = "분석표_25년 3분기_캡스톤(업데이트).xlsx"
        if not os.path.exists(excel_file):
            print(f"  ⚠️ 테스트 파일이 없습니다: {excel_file}")
            return None
        
        # Generator 생성
        gen = UnifiedReportGenerator(
            report_type='service',
            excel_path=excel_file,
            year=2025,
            quarter=3
        )
        print(f"  ✅ Generator 생성 성공")
        
        # 데이터 로드
        gen.load_data()
        print(f"  ✅ 데이터 로드 성공")
        print(f"    - Target 컬럼: {gen.target_col}")
        print(f"    - 지역명 컬럼: {gen.region_name_col}")
        print(f"    - 업종명 컬럼: {gen.industry_name_col}")
        
        # 테이블 데이터 추출
        table_data = gen._extract_table_data_ssot()
        print(f"  ✅ 테이블 데이터 추출 성공: {len(table_data)}개 지역")
        
        return True
    except Exception as e:
        print(f"  ❌ Generator 테스트 실패: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_routes():
    """라우트 엔드포인트 테스트"""
    print("\n[테스트 4] 라우트 엔드포인트")
    try:
        from app import create_app
        app = create_app()
        
        # 등록된 라우트 확인
        routes = []
        for rule in app.url_map.iter_rules():
            if rule.endpoint != 'static':
                routes.append(f"{rule.rule} [{', '.join(rule.methods)}]")
        
        print(f"  ✅ 등록된 라우트: {len(routes)}개")
        for route in sorted(routes)[:10]:  # 처음 10개만 표시
            print(f"    - {route}")
        
        if len(routes) > 10:
            print(f"    ... 외 {len(routes) - 10}개")
        
        return True
    except Exception as e:
        print(f"  ❌ 라우트 테스트 실패: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """메인 테스트"""
    print("="*60)
    print("Flask 앱 기본 동작 테스트")
    print("="*60)
    
    results = {}
    
    # 테스트 실행
    results['imports'] = test_imports()
    results['app_creation'] = test_app_creation()
    results['unified_generator'] = test_unified_generator()
    results['routes'] = test_routes()
    
    # 결과 요약
    print("\n" + "="*60)
    print("테스트 결과 요약")
    print("="*60)
    
    success_count = sum(1 for v in results.values() if v is True)
    skip_count = sum(1 for v in results.values() if v is None)
    total_count = len(results)
    
    for test_name, result in results.items():
        if result is True:
            status = "✅ 성공"
        elif result is None:
            status = "⚠️ 스킵"
        else:
            status = "❌ 실패"
        print(f"  {status} - {test_name}")
    
    print(f"\n[전체 결과] {success_count}/{total_count - skip_count} 성공 ({skip_count}개 스킵)")
    
    if success_count == total_count - skip_count:
        print("✅ 모든 테스트 통과! 앱이 정상 작동합니다.")
        return 0
    else:
        print("❌ 일부 테스트 실패")
        return 1

if __name__ == '__main__':
    sys.exit(main())
