#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""데이터 구조 상세 분석"""

import sys
import json
from pathlib import Path
from typing import Any, cast

sys.path.insert(0, str(Path(__file__).parent))

from report_generator import ReportGenerator

excel_path = "/Users/topos/Library/CloudStorage/GoogleDrive-ckdwo0605@gmail.com/내 드라이브/capstone/분석표_25년 3분기_캡스톤(업데이트).xlsx"

generator = ReportGenerator(excel_path)

# 생산동향 (광공업생산)
print("\n" + "="*70)
print("📊 광공업생산 (manufacturing) 데이터 구조")
print("="*70)

data: dict[str, Any] = generator.extract_data('manufacturing')

print("\n1️⃣ Top-level 키:")
for key in data.keys():
    print(f"  - {key}: {type(data[key]).__name__}")

print("\n2️⃣ regional_data 상세:")
rd: Any = data.get('regional_data', {})
print(f"  타입: {type(rd).__name__}")
if isinstance(rd, dict):
    rd_dict = cast(dict[str, Any], rd)
    print(f"  1차 키: {list(rd_dict.keys())}")
    
    # all_regions 확인
    if 'all_regions' in rd_dict:
        all_regions = cast(list[Any], rd_dict['all_regions'])
        print(f"\n  all_regions (타입: {type(all_regions).__name__}, 길이: {len(all_regions)}):")
        if all_regions:
            first_region = all_regions[0]
            print(f"    - 첫 번째 지역 타입: {type(first_region).__name__}")
            if isinstance(first_region, dict):
                first_region_dict = cast(dict[str, Any], first_region)
                print(f"    - 첫 번째 지역 필드: {list(first_region_dict.keys())}")
            else:
                print("    - 첫 번째 지역 필드: N/A")
            print(f"    - 샘플: {first_region}")
    
    # increase_regions 확인
    if 'increase_regions' in rd_dict:
        increase_regions = cast(list[Any], rd_dict['increase_regions'])
        print(f"\n  increase_regions (타입: {type(increase_regions).__name__}, 길이: {len(increase_regions)}):")
        if increase_regions:
            first = increase_regions[0]
            print(f"    - 첫 번째 항목 타입: {type(first).__name__}")
            print(f"    - 첫 번째 항목: {first}")

print("\n3️⃣ nationwide_data 샘플:")
nd: Any = data.get('nationwide_data', {})
print(f"  타입: {type(nd).__name__}")
print(f"  필드: {list(cast(dict[str, Any], nd).keys()) if isinstance(nd, dict) else 'N/A'}")
if isinstance(nd, dict):
    nd_dict = cast(dict[str, Any], nd)
    print(f"  샘플 값들:")
    for k, v in list(nd_dict.items())[:3]:
        print(f"    - {k}: {v} ({type(v).__name__})")

print("\n4️⃣ summary_table 구조:")
st: Any = data.get('summary_table', {})
if isinstance(st, dict):
    st_dict = cast(dict[str, Any], st)
    print(f"  1차 키: {list(st_dict.keys())}")
    
    if 'columns' in st_dict:
        print(f"\n  columns:")
        cols = st_dict['columns']
        if isinstance(cols, dict):
            cols_dict = cast(dict[str, Any], cols)
            for col_key, col_val in cols_dict.items():
                print(f"    - {col_key}: {col_val}")
    
    if 'regions' in st_dict:
        regions = cast(list[Any], st_dict['regions'])
        print(f"\n  regions (길이: {len(regions)}):")
        if regions:
            first_row = regions[0]
            print(f"    - 첫 번째 행: {first_row}")

# 한 가지만 깊게 체크
print("\n" + "="*70)
print("🔍 첫 번째 지역 complete 데이터:")
print("="*70)

if isinstance(rd, dict):
    rd_dict = cast(dict[str, Any], rd)
    all_regions = cast(list[Any], rd_dict.get('all_regions', []))
    if all_regions:
        region = all_regions[0]
        print(json.dumps(region, indent=2, ensure_ascii=False, default=str))
