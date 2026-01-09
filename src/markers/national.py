"""
전국 관련 마커 처리 핸들러
전국 증감률, 방향, 산업/업태 등 처리
"""

import re
from typing import Optional, List, Dict, Any

from .base import MarkerHandler, MarkerContext


class NationalMarkerHandler(MarkerHandler):
    """전국 관련 마커를 처리하는 핸들러"""
    
    # 전국 관련 키 패턴들
    NATIONAL_KEYS = {
        '전국_이름', '전국_증감률', '전국_증감방향', '전국_변화표현', '전국_방향',
        '전국_증감pp', '연도', '분기', '기준연도', '기준분기'
    }
    
    def can_handle(self, ctx: MarkerContext) -> bool:
        """전국 관련 마커인지 확인합니다."""
        key = ctx.key
        
        # 기본 전국 키
        if key in self.NATIONAL_KEYS:
            return True
        
        # 전국 업태/산업 패턴
        if key.startswith('전국_업태') or key.startswith('전국_산업'):
            return True
        
        # 전국 연령별 패턴
        if re.match(r'^전국_(.+)_증감(?:률|pp)$', key):
            # 전국_품목명_증감률 형태는 별도 핸들러에서 처리
            # 연령 관련 키만 처리
            age_patterns = ['60세이상', '30_59세', '15_29세', '계']
            for pattern in age_patterns:
                if pattern in key:
                    return True
            return False
        
        return False
    
    def handle(self, ctx: MarkerContext) -> Optional[str]:
        """전국 관련 마커를 처리합니다."""
        key = ctx.key
        data = ctx.get_data()
        formatter = ctx.formatter
        
        # 기본 전국 키 처리
        if key == '전국_이름':
            return self._handle_national_name(ctx, data)
        elif key == '전국_증감률':
            return self._handle_national_growth_rate(ctx, data)
        elif key == '전국_증감pp':
            return self._handle_national_growth_pp(ctx, data)
        elif key == '전국_증감방향':
            return self._handle_national_direction(ctx, data, "increase_decrease")
        elif key == '전국_변화표현':
            return self._handle_national_change_expression(ctx, data)
        elif key == '전국_방향':
            return self._handle_national_direction(ctx, data, "rise_fall")
        elif key == '연도':
            return str(ctx.year)
        elif key == '분기':
            return str(ctx.quarter)
        elif key == '기준연도':
            return str(ctx.year)
        elif key == '기준분기':
            return f'{ctx.quarter}/4'
        
        # 전국 업태/산업 패턴
        if key.startswith('전국_업태') or key.startswith('전국_산업'):
            return self._handle_national_industry(ctx, data)
        
        # 전국 연령별 패턴
        national_age_match = re.match(r'^전국_(.+)_증감(?:률|pp)$', key)
        if national_age_match:
            return self._handle_national_age_group(ctx, national_age_match.group(1))
        
        return None
    
    def _handle_national_name(self, ctx: MarkerContext, data: Dict[str, Any]) -> str:
        """전국_이름 마커를 처리합니다."""
        if 'national_region' in data and data['national_region']:
            return data['national_region']['name']
        return "전국"
    
    def _handle_national_growth_rate(self, ctx: MarkerContext, data: Dict[str, Any]) -> Optional[str]:
        """전국_증감률 마커를 처리합니다."""
        # 실업률/고용률 시트는 별도 처리
        is_unemployment_sheet = ('실업' in ctx.sheet_name or '고용' in ctx.sheet_name)
        if is_unemployment_sheet:
            return None  # 다음 핸들러에 위임
        
        if 'national_region' in data and data['national_region']:
            growth_rate = data['national_region'].get('growth_rate')
            if growth_rate is not None:
                return ctx.formatter.format_percentage(growth_rate, decimal_places=1, include_percent=False)
        
        # 직접 계산 시도
        growth_rate = ctx.dynamic_parser.calculate_growth_rate(ctx.sheet_name, '전국', ctx.year, ctx.quarter)
        if growth_rate is not None:
            return ctx.formatter.format_percentage(growth_rate, decimal_places=1, include_percent=False)
        
        return None
    
    def _handle_national_growth_pp(self, ctx: MarkerContext, data: Dict[str, Any]) -> Optional[str]:
        """전국_증감pp 마커를 처리합니다 (퍼센트포인트)."""
        return self._handle_national_growth_rate(ctx, data)
    
    def _handle_national_direction(self, ctx: MarkerContext, data: Dict[str, Any], direction_type: str) -> Optional[str]:
        """전국 증감 방향을 처리합니다."""
        growth_rate = None
        
        if 'national_region' in data and data['national_region']:
            growth_rate = data['national_region'].get('growth_rate')
        
        if growth_rate is None:
            growth_rate = ctx.dynamic_parser.calculate_growth_rate(ctx.sheet_name, '전국', ctx.year, ctx.quarter)
        
        if growth_rate is not None:
            if direction_type == "rise_fall":
                return ctx.formatter.get_growth_direction(growth_rate, direction_type="rise_fall", expression_key="rate")
            else:
                return ctx.formatter.get_growth_direction(growth_rate)
        
        return None
    
    def _handle_national_change_expression(self, ctx: MarkerContext, data: Dict[str, Any]) -> Optional[str]:
        """전국_변화표현 마커를 처리합니다."""
        growth_rate = None
        
        if 'national_region' in data and data['national_region']:
            growth_rate = data['national_region'].get('growth_rate')
        
        if growth_rate is None:
            growth_rate = ctx.dynamic_parser.calculate_growth_rate(ctx.sheet_name, '전국', ctx.year, ctx.quarter)
        
        if growth_rate is not None:
            return ctx.formatter.get_production_change_expression(growth_rate, direction_type="rise_fall")
        
        return None
    
    def _handle_national_industry(self, ctx: MarkerContext, data: Dict[str, Any]) -> Optional[str]:
        """전국 산업/업태 마커를 처리합니다."""
        industry_match = re.match(r'전국_(업태|산업)(\d+)_(.+)', ctx.key)
        if not industry_match:
            return None
        
        industry_idx = int(industry_match.group(2)) - 1
        industry_field = industry_match.group(3)
        
        # 물가 시트 여부 확인
        is_price_sheet = ('물가' in ctx.sheet_name)
        top_n = 4 if is_price_sheet else 3
        
        # 분석된 데이터에서 카테고리 가져오기
        categories = []
        if 'national_region' in data and data['national_region']:
            categories = data['national_region'].get('top_industries', [])
        
        if not categories:
            # 직접 계산 필요 - 상위 클래스에서 처리
            return None
        
        if categories and industry_idx < len(categories):
            category = categories[industry_idx]
            
            if industry_field == '이름':
                category_name = category['name']
                # 이름 매핑 적용
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
    
    def _handle_national_age_group(self, ctx: MarkerContext, age_group_key: str) -> Optional[str]:
        """전국 연령별 증감률/증감pp 마커를 처리합니다."""
        # 실업률 시트 여부 확인
        is_unemployment_sheet = (ctx.sheet_name == '실업률' or ctx.sheet_name == '실업자 수')
        
        if not is_unemployment_sheet:
            return None
        
        # 연령대 이름 매핑
        age_mapping = {
            '60세이상': ['60세이상', '60 세이상', '60세 이상'],
            '30_59세': ['30~59세', '30 - 59세', '30-59세', '30 ~ 59세'],
            '15_29세': ['15~29세', '15 - 29세', '15-29세', '15 ~ 29세']
        }
        target_ages = age_mapping.get(age_group_key, [age_group_key])
        
        # 이 부분은 UnemploymentHandler에서 처리해야 함
        return None

