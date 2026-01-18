#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Flask 앱 기능 테스트
"""
import requests
import json
import time

BASE_URL = "http://localhost:5050"

def test_main_page():
    """메인 페이지 접근 테스트"""
    print("\n[테스트 1] 메인 페이지 접근")
    try:
        response = requests.get(BASE_URL, timeout=5)
        if response.status_code == 200:
            print(f"  ✅ 성공: 상태 코드 {response.status_code}")
            print(f"  - 응답 크기: {len(response.content)} bytes")
            return True
        else:
            print(f"  ❌ 실패: 상태 코드 {response.status_code}")
            return False
    except Exception as e:
        print(f"  ❌ 실패: {e}")
        return False

def test_session_info():
    """세션 정보 API 테스트"""
    print("\n[테스트 2] 세션 정보 API")
    try:
        response = requests.get(f"{BASE_URL}/api/session-info", timeout=5)
        if response.status_code == 200:
            data = response.json()
            print(f"  ✅ 성공: {data}")
            return True
        else:
            print(f"  ❌ 실패: 상태 코드 {response.status_code}")
            return False
    except Exception as e:
        print(f"  ❌ 실패: {e}")
        return False

def test_file_upload():
    """파일 업로드 테스트"""
    print("\n[테스트 3] 파일 업로드 기능")
    excel_file = "분석표_25년 3분기_캡스톤(업데이트).xlsx"
    try:
        # 파일 존재 확인
        with open(excel_file, 'rb') as f:
            files = {'file': (excel_file, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            data = {
                'year': '2025',
                'quarter': '3'
            }
            response = requests.post(f"{BASE_URL}/api/upload", files=files, data=data, timeout=30)
            
        if response.status_code == 200:
            result = response.json()
            print(f"  ✅ 성공: {result.get('message', 'OK')}")
            return True
        else:
            print(f"  ❌ 실패: 상태 코드 {response.status_code}")
            if response.text:
                print(f"  - 오류 메시지: {response.text[:200]}")
            return False
    except FileNotFoundError:
        print(f"  ⚠️ 테스트 파일이 없습니다: {excel_file}")
        return None
    except Exception as e:
        print(f"  ❌ 실패: {e}")
        return False

def test_report_generation():
    """보고서 생성 테스트"""
    print("\n[테스트 4] 보고서 생성 기능")
    try:
        # 먼저 파일 업로드
        excel_file = "분석표_25년 3분기_캡스톤(업데이트).xlsx"
        with open(excel_file, 'rb') as f:
            files = {'file': (excel_file, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            data = {'year': '2025', 'quarter': '3'}
            upload_response = requests.post(f"{BASE_URL}/api/upload", files=files, data=data, timeout=30)
        
        if upload_response.status_code != 200:
            print(f"  ❌ 업로드 실패")
            return False
        
        # 보고서 생성 요청
        print("  - 서비스업생산 보고서 생성 중...")
        report_data = {
            'report_type': 'service',
            'year': '2025',
            'quarter': '3'
        }
        response = requests.post(f"{BASE_URL}/api/generate", json=report_data, timeout=60)
        
        if response.status_code == 200:
            result = response.json()
            if result.get('success'):
                print(f"  ✅ 성공: 보고서 생성 완료")
                print(f"  - HTML 크기: {len(result.get('html', ''))} bytes")
                return True
            else:
                print(f"  ❌ 실패: {result.get('error', 'Unknown error')}")
                return False
        else:
            print(f"  ❌ 실패: 상태 코드 {response.status_code}")
            return False
    except FileNotFoundError:
        print(f"  ⚠️ 테스트 파일이 없습니다")
        return None
    except Exception as e:
        print(f"  ❌ 실패: {e}")
        return False

def main():
    """메인 테스트"""
    print("="*60)
    print("Flask 앱 기능 테스트")
    print("="*60)
    
    # 서버가 실행 중인지 확인
    try:
        requests.get(BASE_URL, timeout=2)
    except:
        print("\n❌ 서버가 실행되지 않았습니다. python app.py 를 먼저 실행하세요.")
        return 1
    
    results = {}
    
    # 테스트 실행
    results['main_page'] = test_main_page()
    results['session_info'] = test_session_info()
    results['file_upload'] = test_file_upload()
    results['report_generation'] = test_report_generation()
    
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
    
    if success_count == total_count - skip_count and skip_count == 0:
        print("✅ 모든 테스트 통과!")
        return 0
    elif success_count > 0:
        print("⚠️ 일부 테스트 통과")
        return 0
    else:
        print("❌ 테스트 실패")
        return 1

if __name__ == '__main__':
    import sys
    sys.exit(main())
