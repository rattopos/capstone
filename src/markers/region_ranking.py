"""
지역 순위 관련 마커 처리 핸들러
상위시도, 하위시도, 강조시도 등 처리
"""

import re
from typing import Optional, List, Dict, Any

from .base import MarkerHandler, MarkerContext


class RegionRankingHandler(MarkerHandler):
    """지역 순위 관련 마커를 처리하는 핸들러"""
    
    def can_handle(self, ctx: MarkerContext) -> bool:
        """지역 순위 관련 마커인지 확인합니다."""
        key = ctx.key
        
        # 상위/하위/강조 시도 패턴
        if re.match(r'상위시도\d+_', key):
            return True
        if re.match(r'하위시도\d+_', key):
            return True
        if key.startswith('강조시도') or key == '강조_시도수' or key == '강조_방향':
            return True
        
        return False
    
    def handle(self, ctx: MarkerContext) -> Optional[str]:
        """지역 순위 관련 마커를 처리합니다."""
        key = ctx.key
        data = ctx.get_data()
        
        # 상위시도 패턴
        top_match = re.match(r'상위시도(\d+)_(.+)', key)
        if top_match:
            return self._handle_top_region(ctx, data, top_match)
        
        # 하위시도 패턴
        bottom_match = re.match(r'하위시도(\d+)_(.+)', key)
        if bottom_match:
            return self._handle_bottom_region(ctx, data, bottom_match)
        
        # 강조시도 패턴
        if key.startswith('강조시도'):
            return self._handle_emphasis_region(ctx, data)
        
        if key == '강조_시도수':
            return self._handle_emphasis_count(ctx, data)
        
        if key == '강조_방향':
            return self._handle_emphasis_direction(ctx, data)
        
        return None
    
    def _get_regions_sorted(self, ctx: MarkerContext, data: Dict[str, Any], 
                            ascending: bool = False) -> List[Dict[str, Any]]:
        """정렬된 지역 리스트를 반환합니다."""
        # 캐시된 데이터 사용
        if ascending:
            regions = data.get('bottom_regions', [])
        else:
            regions = data.get('top_regions', [])
        
        if regions:
            return regions
        
        # 직접 계산
        try:
            from ..utils.sheet_utils import get_actual_sheet_name
            actual_sheet_name = get_actual_sheet_name(ctx.sheet_name)
            
            # 분기 열 가져오기 - dynamic_parser 사용
            current_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year, ctx.quarter)
            prev_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year - 1, ctx.quarter)
            
            if current_col and prev_col:
                regions_data = ctx.data_analyzer.get_regions_with_growth_rate(
                    actual_sheet_name, current_col, prev_col
                )
                if regions_data:
                    return sorted(regions_data, key=lambda x: x.get('growth_rate', 0), reverse=not ascending)
        except Exception:
            pass
        
        return []
    
    def _handle_top_region(self, ctx: MarkerContext, data: Dict[str, Any], match) -> Optional[str]:
        """상위시도 마커를 처리합니다."""
        idx = int(match.group(1)) - 1
        field = match.group(2)
        
        # 실업률/물가 시트 특별 처리는 상위 클래스에서
        is_unemployment_sheet = (ctx.sheet_name == '실업률' or ctx.sheet_name == '실업자 수')
        is_price_sheet = ('물가' in ctx.sheet_name)
        
        if is_unemployment_sheet:
            return None  # UnemploymentHandler에서 처리
        
        regions = self._get_regions_sorted(ctx, data, ascending=False)
        
        if regions and idx < len(regions):
            return self._extract_region_field(ctx, regions[idx], field)
        
        return None
    
    def _handle_bottom_region(self, ctx: MarkerContext, data: Dict[str, Any], match) -> Optional[str]:
        """하위시도 마커를 처리합니다."""
        idx = int(match.group(1)) - 1
        field = match.group(2)
        
        is_unemployment_sheet = (ctx.sheet_name == '실업률' or ctx.sheet_name == '실업자 수')
        
        if is_unemployment_sheet:
            return None  # UnemploymentHandler에서 처리
        
        regions = self._get_regions_sorted(ctx, data, ascending=True)
        
        if regions and idx < len(regions):
            return self._extract_region_field(ctx, regions[idx], field)
        
        return None
    
    def _handle_emphasis_region(self, ctx: MarkerContext, data: Dict[str, Any]) -> Optional[str]:
        """강조시도 마커를 처리합니다."""
        emphasis_match = re.match(r'강조시도(\d+)_(.+)', ctx.key)
        if not emphasis_match:
            return None
        
        idx = int(emphasis_match.group(1)) - 1
        field = emphasis_match.group(2)
        
        # 전국 증감률에 따라 상위/하위 결정
        national_growth = self._get_national_growth(ctx, data)
        
        if national_growth is None:
            return None
        
        if national_growth >= 0:
            regions = self._get_regions_sorted(ctx, data, ascending=False)
        else:
            regions = self._get_regions_sorted(ctx, data, ascending=True)
        
        if regions and idx < len(regions):
            return self._extract_region_field(ctx, regions[idx], field)
        
        return None
    
    def _handle_emphasis_count(self, ctx: MarkerContext, data: Dict[str, Any]) -> Optional[str]:
        """강조_시도수 마커를 처리합니다."""
        national_growth = self._get_national_growth(ctx, data)
        
        if national_growth is None:
            return None
        
        # 전국이 증가면 증가_시도수, 감소면 감소_시도수
        # 실제 계산은 StatisticsHandler에서 처리
        return None
    
    def _handle_emphasis_direction(self, ctx: MarkerContext, data: Dict[str, Any]) -> Optional[str]:
        """강조_방향 마커를 처리합니다."""
        national_growth = self._get_national_growth(ctx, data)
        
        if national_growth is not None:
            return ctx.formatter.get_growth_direction(national_growth)
        
        return None
    
    def _get_national_growth(self, ctx: MarkerContext, data: Dict[str, Any]) -> Optional[float]:
        """전국 증감률을 가져옵니다."""
        if 'national_region' in data and data['national_region']:
            return data['national_region'].get('growth_rate')
        
        return ctx.dynamic_parser.calculate_growth_rate(ctx.sheet_name, '전국', ctx.year, ctx.quarter)
    
    def _extract_region_field(self, ctx: MarkerContext, region: Dict[str, Any], field: str) -> Optional[str]:
        """지역 데이터에서 필드 값을 추출합니다."""
        if field == '이름':
            return region.get('name', '')
        elif field in ('증감률', '증감pp'):
            return ctx.formatter.format_percentage(region.get('growth_rate', 0), decimal_places=1, include_percent=False)
        elif field == '방향':
            return ctx.formatter.get_growth_direction(region['growth_rate'], direction_type="rise_fall", expression_key="rate")
        elif field == '변화표현':
            return ctx.formatter.get_production_change_expression(region['growth_rate'], direction_type="rise_fall")
        elif field.startswith('업태') or field.startswith('산업'):
            return self._extract_industry_field(ctx, region, field)
        elif field.startswith('연령'):
            return None  # UnemploymentHandler에서 처리
        
        return None
    
    def _extract_industry_field(self, ctx: MarkerContext, region: Dict[str, Any], field: str) -> Optional[str]:
        """산업/업태 필드를 추출합니다."""
        industry_match = re.match(r'(업태|산업)(\d+)_(.+)', field)
        if not industry_match:
            return None
        
        industry_idx = int(industry_match.group(2)) - 1
        industry_field = industry_match.group(3)
        
        categories = region.get('top_industries', [])
        
        if categories and industry_idx < len(categories):
            category = categories[industry_idx]
            
            if industry_field == '이름':
                category_name = category['name']
                sheet_config = ctx.schema_loader.load_sheet_config(ctx.sheet_name)
                name_mapping = sheet_config.get('name_mapping', {})
                mapped_name = name_mapping.get(category_name)
                if not mapped_name:
                    for key_map, value_map in name_mapping.items():
                        if key_map in category_name or category_name in key_map:
                            mapped_name = value_map
                            break
                if not mapped_name:
                    item_display_mapping = ctx.schema_loader.get_name_mapping('item_display_mapping')
                    if item_display_mapping:
                        mapped_name = item_display_mapping.get(category_name)
                return mapped_name if mapped_name else category_name
            elif industry_field == '증감률':
                return ctx.formatter.format_percentage(category['growth_rate'], decimal_places=1, include_percent=False)
        
        return None

