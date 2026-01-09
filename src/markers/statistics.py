"""
통계 관련 마커 처리 핸들러
증가/감소/상승/하락 시도수 등 처리
"""

import re
from typing import Optional, Dict, Any

from .base import MarkerHandler, MarkerContext


class StatisticsHandler(MarkerHandler):
    """통계 관련 마커를 처리하는 핸들러"""
    
    STATISTICS_KEYS = {
        '증가시도수', '증가_시도수', '감소시도수', '감소_시도수',
        '상승시도수', '상승_시도수', '하락시도수', '하락_시도수'
    }
    
    def can_handle(self, ctx: MarkerContext) -> bool:
        """통계 관련 마커인지 확인합니다."""
        return ctx.key in self.STATISTICS_KEYS
    
    def handle(self, ctx: MarkerContext) -> Optional[str]:
        """통계 관련 마커를 처리합니다."""
        key = ctx.key
        
        if key in ('증가시도수', '증가_시도수'):
            return self._count_increase_regions(ctx)
        elif key in ('감소시도수', '감소_시도수'):
            return self._count_decrease_regions(ctx)
        elif key in ('상승시도수', '상승_시도수'):
            return self._count_increase_regions(ctx)  # 상승 = 증가
        elif key in ('하락시도수', '하락_시도수'):
            return self._count_decrease_regions(ctx)  # 하락 = 감소
        
        return None
    
    def _count_regions_by_condition(self, ctx: MarkerContext, condition: str) -> int:
        """조건에 맞는 지역 수를 카운트합니다."""
        from ..utils.sheet_utils import get_actual_sheet_name
        
        sheet_name = get_actual_sheet_name(ctx.sheet_name)
        
        try:
            sheet = ctx.excel_extractor.get_sheet(sheet_name)
            if not sheet:
                return 0
            
            # 분기 열 가져오기
            current_col = ctx.dynamic_parser.get_quarter_column(sheet_name, ctx.year, ctx.quarter)
            prev_col = ctx.dynamic_parser.get_quarter_column(sheet_name, ctx.year - 1, ctx.quarter)
            
            if not current_col or not prev_col:
                return 0
            
            # 시트 설정 가져오기
            sheet_config = ctx.schema_loader.load_sheet_config(sheet_name)
            category_col = sheet_config.get('category_column', 6)
            
            count = 0
            seen_regions = set()
            
            for row in range(4, min(5000, sheet.max_row + 1)):
                cell_a = sheet.cell(row=row, column=1)  # 지역 코드
                cell_b = sheet.cell(row=row, column=2)  # 지역 이름
                cell_category = sheet.cell(row=row, column=category_col)
                
                # 총지수/계 확인
                is_total = False
                if cell_category.value:
                    category_str = str(cell_category.value).strip()
                    if category_str in ['총지수', '계', '   계', '합계']:
                        is_total = True
                
                if not cell_b.value or not is_total:
                    continue
                
                # 시도 코드 확인
                code_str = str(cell_a.value).strip() if cell_a.value else ''
                is_sido = (len(code_str) == 2 and code_str.isdigit() and code_str != '00')
                
                region_name = str(cell_b.value).strip()
                if is_sido and region_name not in seen_regions:
                    seen_regions.add(region_name)
                    
                    current = sheet.cell(row=row, column=current_col).value
                    prev = sheet.cell(row=row, column=prev_col).value
                    
                    if current is not None and prev is not None:
                        try:
                            prev_num = float(prev)
                            if prev_num == 0:
                                continue
                            current_num = float(current)
                            growth_rate = ((current_num / prev_num) - 1) * 100
                            
                            if condition == 'increase' and growth_rate > 0:
                                count += 1
                            elif condition == 'decrease' and growth_rate < 0:
                                count += 1
                        except (ValueError, TypeError):
                            continue
            
            return count
        except Exception:
            return 0
    
    def _count_increase_regions(self, ctx: MarkerContext) -> str:
        """증가한 지역 수를 반환합니다."""
        count = self._count_regions_by_condition(ctx, 'increase')
        return str(count)
    
    def _count_decrease_regions(self, ctx: MarkerContext) -> str:
        """감소한 지역 수를 반환합니다."""
        count = self._count_regions_by_condition(ctx, 'decrease')
        return str(count)

