#!/usr/bin/env python
"""
웹 애플리케이션 실행 스크립트
"""

from app import app

if __name__ == '__main__':
    print("=" * 60)
    print("통계청 보도자료 자동 생성 시스템 - 웹 서버")
    print("=" * 60)
    print("\n서버가 시작되었습니다!")
    print("브라우저에서 http://localhost:8000 으로 접속하세요.")
    print("\n종료하려면 Ctrl+C를 누르세요.")
    print("=" * 60)
    print()
    
    app.run(debug=True, host='0.0.0.0', port=8000)

