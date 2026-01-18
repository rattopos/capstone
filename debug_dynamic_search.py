#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
from pathlib import Path

from templates.unified_generator import UnifiedReportGenerator


def main():
    base = Path(__file__).parent
    excel_path = base / "분석표_25년 3분기_캡스톤(업데이트).xlsx"
    if not excel_path.exists():
        print(f"엑셀 파일을 찾을 수 없습니다: {excel_path}")
        sys.exit(1)

    # 서비스업생산을 예시로 테스트 (필요시 다른 부문으로 변경)
    gen = UnifiedReportGenerator(
        report_type="service",
        excel_path=str(excel_path),
        year=2025,
        quarter=3,
    )

    table_data = gen._extract_table_data_ssot()
    print("\n[전국 및 시도 지수] (상위 5개 출력)")
    for row in table_data[:5]:
        print(row)

    nationwide = gen.extract_nationwide_data(table_data)
    print("\n[전국 요약]")
    print(nationwide)

    industries = gen._extract_industry_data('전국')
    print("\n[전국 업종 데이터] (상위 5개 출력)")
    for ind in (industries or [])[:5]:
        print(ind)


if __name__ == "__main__":
    main()
