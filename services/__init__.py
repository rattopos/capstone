# -*- coding: utf-8 -*-
"""
서비스 모듈 초기화
"""

from .report_generator import (
    generate_report_html,
    generate_regional_report_html,
    generate_statistics_report_html,
    generate_individual_statistics_html
)
from .summary_data import (
    get_summary_overview_data,
    get_summary_table_data,
    get_production_summary_data,
    get_consumption_construction_data,
    get_trade_price_data,
    get_employment_population_data
)

__all__ = [
    'generate_report_html',
    'generate_regional_report_html',
    'generate_statistics_report_html',
    'generate_individual_statistics_html',
    'get_summary_overview_data',
    'get_summary_table_data',
    'get_production_summary_data',
    'get_consumption_construction_data',
    'get_trade_price_data',
    'get_employment_population_data'
]

