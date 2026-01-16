#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generator import 테스트"""

from pathlib import Path

print("=" * 70)
print("Generator Import 테스트")
print("=" * 70)

excel_path = Path(__file__).parent / '분석표_25년 3분기_캡스톤(업데이트).xlsx'

# 1. mining_manufacturing_generator
print("\n[1] mining_manufacturing_generator.py import...")
try:
    from templates.mining_manufacturing_generator import MiningManufacturingGenerator
    generator = MiningManufacturingGenerator(str(excel_path), 2025, 3)
    data = generator.extract_all_data()
    print(f"✅ 성공: 전국 {data['nationwide_data']['growth_rate']}%")
except Exception as e:
    print(f"❌ 실패: {e}")
    import traceback
    traceback.print_exc()

# 2. service_industry_generator
print("\n[2] service_industry_generator.py import...")
try:
    from templates.service_industry_generator import ServiceIndustryGenerator
    generator = ServiceIndustryGenerator(str(excel_path), 2025, 3)
    data = generator.extract_all_data()
    print(f"✅ 성공: 전국 {data['nationwide_data']['growth_rate']}%")
except Exception as e:
    print(f"❌ 실패: {e}")
    import traceback
    traceback.print_exc()

# 3. consumption_generator
print("\n[3] consumption_generator.py import...")
try:
    from templates.consumption_generator import ConsumptionGenerator
    generator = ConsumptionGenerator(str(excel_path), 2025, 3)
    data = generator.extract_all_data()
    print(f"✅ 성공: 전국 {data['nationwide_data']['growth_rate']}%")
except Exception as e:
    print(f"❌ 실패: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 70)
print("테스트 완료")
print("=" * 70)
