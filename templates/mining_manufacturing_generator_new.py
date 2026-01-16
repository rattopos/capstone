#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
광공업생산 보도자료 생성기 (통합 Generator 기반)
"""

# 통합 Generator의 Wrapper 사용
try:
    from .unified_generator import MiningManufacturingGenerator
except ImportError:
    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).parent))
    from unified_generator import MiningManufacturingGenerator

# 하위 호환성을 위한 별칭
광공업생산Generator = MiningManufacturingGenerator

__all__ = ['MiningManufacturingGenerator', '광공업생산Generator']
