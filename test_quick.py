#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
최종 수정사항 간단 테스트
"""
import sys
from pathlib import Path

# 경로 설정
base_path = Path("/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone")
sys.path.insert(0, str(base_path))

# Excel 파일 찾기
excel_files = list(base_path.glob("*.xlsx"))
if excel_files:
    excel_path = excel_files[0]
    print(f"✅ Excel 파일 찾음: {excel_path.name}")
else:
    print("❌ Excel 파일 없음")
    sys.exit(1)

from services.report_generator import generate_report_html

tests = [
    ('employment', '고용률'),
    ('unemployment', '실업률'),
    ('migration', '국내인구이동'),
]

print("\n" + "="*70)
print("최종 수정사항 테스트")
print("="*70)

success_count = 0
for report_id, report_name in tests:
    print(f"\n📊 {report_name} (ID: {report_id})")
    try:
        html = generate_report_html(report_id, str(excel_path), 2025, 3)
        if html and len(html) > 500:
            print(f"  ✅ HTML 생성 성공 ({len(html)} bytes)")
            success_count += 1
        else:
            print(f"  ❌ HTML 생성 실패 (크기: {len(html) if html else 0})")
    except Exception as e:
        error_msg = str(e)
        if 'UndefinedError' in error_msg:
            print(f"  ❌ 템플릿 오류: {error_msg[:80]}")
        else:
            print(f"  ❌ 오류: {error_msg[:80]}")

print("\n" + "="*70)
print(f"최종 결과: {success_count}/3 성공")
print("="*70)
