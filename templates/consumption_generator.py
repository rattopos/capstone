#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
소비동향 보도자료 생성기 (통합 Generator 기반)
"""

# 통합 Generator의 Wrapper 사용
try:
    from .unified_generator import ConsumptionGenerator
except ImportError:
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).parent))
    from unified_generator import ConsumptionGenerator

__all__ = ['ConsumptionGenerator']
