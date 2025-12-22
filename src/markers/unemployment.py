"""
실업률/고용률 관련 마커 처리 핸들러
실업률, 고용률, 연령별 데이터 등 처리
"""

import re
from typing import Optional, List, Dict, Any

from .base import MarkerHandler, MarkerContext


class UnemploymentHandler(MarkerHandler):
    """실업률/고용률 관련 마커를 처리하는 핸들러"""
    
    def can_handle(self, ctx: MarkerContext) -> bool:
        """실업률/고용률 관련 마커인지 확인합니다."""
        # 실업률/고용률 시트인 경우
        is_unemployment_sheet = ('실업' in ctx.sheet_name or ctx.sheet_name == '실업자 수')
        is_employment_sheet = ('고용' in ctx.sheet_name or '고용률' in ctx.sheet_name)
        
        if not (is_unemployment_sheet or is_employment_sheet):
            return False
        
        key = ctx.key
        
        # 전국 증감률/증감pp (실업률/고용률 전용)
        if key in ('전국_증감률', '전국_증감pp'):
            return True
        
        # 지역별 실업률/고용률
        if re.match(r'^[가-힣]+_(실업률|고용률)$', key):
            return True
        
        # 지역별 증감pp
        if re.match(r'^[가-힣]+_증감pp$', key):
            return True
        
        # 상위/하위 시도 (실업률 시트)
        if is_unemployment_sheet and (re.match(r'상위시도\d+_', key) or re.match(r'하위시도\d+_', key)):
            return True
        
        # 연령별 데이터
        if '연령' in key:
            return True
        
        return False
    
    def handle(self, ctx: MarkerContext) -> Optional[str]:
        """실업률/고용률 관련 마커를 처리합니다."""
        key = ctx.key
        
        # 전국 증감률/증감pp
        if key in ('전국_증감률', '전국_증감pp'):
            return self._handle_national_unemployment_change(ctx)
        
        # 지역별 실업률/고용률
        region_value_match = re.match(r'^([가-힣]+)_(실업률|고용률)$', key)
        if region_value_match:
            region_name = region_value_match.group(1)
            value_type = region_value_match.group(2)
            return self._handle_region_rate(ctx, region_name, value_type)
        
        # 지역별 증감pp
        region_change_match = re.match(r'^([가-힣]+)_증감pp$', key)
        if region_change_match:
            region_name = region_change_match.group(1)
            return self._handle_region_change_pp(ctx, region_name)
        
        # 상위/하위 시도
        top_match = re.match(r'상위시도(\d+)_(.+)', key)
        if top_match:
            return self._handle_top_unemployment_region(ctx, top_match)
        
        bottom_match = re.match(r'하위시도(\d+)_(.+)', key)
        if bottom_match:
            return self._handle_bottom_unemployment_region(ctx, bottom_match)
        
        return None
    
    def _get_unemployment_rate_regions_data(self, ctx: MarkerContext) -> List[Dict[str, Any]]:
        """실업률 테이블에서 모든 지역의 증감률(백분율포인트)을 계산합니다."""
        from ..utils.region_utils import SIDO_LIST
        
        actual_sheet_name = '실업자 수'
        sheet = ctx.excel_extractor.get_sheet(actual_sheet_name)
        if not sheet:
            return []
        
        sheet_config = ctx.schema_loader.load_sheet_config(actual_sheet_name)
        unemployment_table_config = sheet_config.get('unemployment_rate_table', {})
        
        if not unemployment_table_config.get('enabled', True):
            return []
        
        start_row = unemployment_table_config.get('start_row', 81)
        region_mapping = unemployment_table_config.get('region_mapping', {})
        
        results = []
        
        # 분기 열 찾기
        current_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year, ctx.quarter)
        prev_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year - 1, ctx.quarter)
        
        if not current_col or not prev_col:
            return []
        
        current_region = None
        
        for row in range(start_row, min(5000, sheet.max_row + 1)):
            cell_a = sheet.cell(row=row, column=1)
            cell_b = sheet.cell(row=row, column=2)
            
            if cell_a.value:
                current_region = str(cell_a.value).strip()
            
            if cell_b.value and current_region:
                age_str = str(cell_b.value).strip()
                
                if age_str == '계':
                    # 짧은 지역명으로 변환
                    reverse_mapping = {v: k for k, v in region_mapping.items()}
                    short_name = reverse_mapping.get(current_region, current_region)
                    
                    if short_name in SIDO_LIST:
                        current_value = sheet.cell(row=row, column=current_col).value
                        prev_value = sheet.cell(row=row, column=prev_col).value
                        
                        if current_value is not None and prev_value is not None:
                            try:
                                current_num = float(current_value)
                                prev_num = float(prev_value)
                                growth_pp = current_num - prev_num  # 퍼센트포인트 차이
                                
                                results.append({
                                    'name': short_name,
                                    'growth_rate': growth_pp
                                })
                            except (ValueError, TypeError):
                                pass
        
        return results
    
    def _handle_national_unemployment_change(self, ctx: MarkerContext) -> Optional[str]:
        """전국 실업률/고용률 증감을 처리합니다."""
        is_unemployment_sheet = ('실업' in ctx.sheet_name or ctx.sheet_name == '실업자 수')
        
        if is_unemployment_sheet:
            # 전국 실업률 증감 계산
            actual_sheet_name = '실업자 수'
            sheet = ctx.excel_extractor.get_sheet(actual_sheet_name)
            if not sheet:
                return None
            
            # 분기 열 찾기
            current_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year, ctx.quarter)
            prev_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year - 1, ctx.quarter)
            
            if not current_col or not prev_col:
                return None
            
            sheet_config = ctx.schema_loader.load_sheet_config(actual_sheet_name)
            unemployment_table_config = sheet_config.get('unemployment_rate_table', {})
            start_row = unemployment_table_config.get('start_row', 81)
            
            current_region = None
            
            for row in range(start_row, min(5000, sheet.max_row + 1)):
                cell_a = sheet.cell(row=row, column=1)
                cell_b = sheet.cell(row=row, column=2)
                
                if cell_a.value:
                    current_region = str(cell_a.value).strip()
                
                if cell_b.value and current_region == '전국':
                    age_str = str(cell_b.value).strip()
                    
                    if age_str == '계':
                        current_value = sheet.cell(row=row, column=current_col).value
                        prev_value = sheet.cell(row=row, column=prev_col).value
                        
                        if current_value is not None and prev_value is not None:
                            try:
                                current_num = float(current_value)
                                prev_num = float(prev_value)
                                growth_pp = current_num - prev_num
                                return ctx.formatter.format_percentage(growth_pp, decimal_places=1, include_percent=False)
                            except (ValueError, TypeError):
                                pass
        
        return None
    
    def _handle_region_rate(self, ctx: MarkerContext, region_name: str, value_type: str) -> Optional[str]:
        """지역별 실업률/고용률 값을 반환합니다."""
        is_unemployment = (value_type == '실업률')
        
        actual_sheet_name = '실업자 수' if is_unemployment else ctx.sheet_name
        sheet = ctx.excel_extractor.get_sheet(actual_sheet_name)
        if not sheet:
            return None
        
        # 분기 열 찾기
        current_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year, ctx.quarter)
        if not current_col:
            return None
        
        sheet_config = ctx.schema_loader.load_sheet_config(actual_sheet_name)
        unemployment_table_config = sheet_config.get('unemployment_rate_table', {})
        start_row = unemployment_table_config.get('start_row', 81)
        region_mapping = unemployment_table_config.get('region_mapping', {})
        
        actual_region_name = region_mapping.get(region_name, region_name)
        current_region = None
        
        for row in range(start_row, min(5000, sheet.max_row + 1)):
            cell_a = sheet.cell(row=row, column=1)
            cell_b = sheet.cell(row=row, column=2)
            
            if cell_a.value:
                current_region = str(cell_a.value).strip()
            
            if cell_b.value and (current_region == actual_region_name or current_region == region_name):
                age_str = str(cell_b.value).strip()
                
                if age_str == '계':
                    value = sheet.cell(row=row, column=current_col).value
                    if value is not None:
                        return ctx.formatter.format_percentage(value, decimal_places=1, include_percent=False)
        
        return None
    
    def _handle_region_change_pp(self, ctx: MarkerContext, region_name: str) -> Optional[str]:
        """지역별 증감pp (퍼센트포인트)를 반환합니다."""
        regions_data = self._get_unemployment_rate_regions_data(ctx)
        
        for region in regions_data:
            if region['name'] == region_name:
                return ctx.formatter.format_percentage(region['growth_rate'], decimal_places=1, include_percent=False)
        
        return None
    
    def _handle_top_unemployment_region(self, ctx: MarkerContext, match) -> Optional[str]:
        """상위 실업률 지역을 처리합니다."""
        idx = int(match.group(1)) - 1
        field = match.group(2)
        
        regions_data = self._get_unemployment_rate_regions_data(ctx)
        if not regions_data:
            return None
        
        # 상위시도 = 실업률이 가장 많이 상승한 지역
        regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0), reverse=True)
        
        if regions and idx < len(regions):
            return self._extract_region_field(ctx, regions[idx], field)
        
        return None
    
    def _handle_bottom_unemployment_region(self, ctx: MarkerContext, match) -> Optional[str]:
        """하위 실업률 지역을 처리합니다."""
        idx = int(match.group(1)) - 1
        field = match.group(2)
        
        regions_data = self._get_unemployment_rate_regions_data(ctx)
        if not regions_data:
            return None
        
        # 하위시도 = 실업률이 가장 많이 하락한 지역
        regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0))
        
        if regions and idx < len(regions):
            return self._extract_region_field(ctx, regions[idx], field)
        
        return None
    
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
        
        return None

