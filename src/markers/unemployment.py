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
        
        # 전국 고용률
        if key == '전국_고용률':
            return True
        
        # 지역별 실업률/고용률
        if re.match(r'^[가-힣]+_(실업률|고용률)$', key):
            return True
        
        # 지역별 증감pp
        if re.match(r'^[가-힣]+_증감pp$', key):
            return True
        
        # 상위/하위 시도 (실업률 및 고용률 시트 모두)
        if re.match(r'상위시도\d+_', key) or re.match(r'하위시도\d+_', key):
            return True
        
        # 연령별 데이터 (전국_30대_증감pp 등)
        if re.match(r'^전국_\d+대.*_증감pp$', key):
            return True
        
        # 연령별 데이터 확장 패턴
        if '연령' in key:
            return True
        
        # 상승/하락 시도수
        if key in ('상승_시도수', '하락_시도수'):
            return True
        
        # 테이블 헤더 마커 (YYYY_N분기 형식)
        if re.match(r'^\d{4}_\d분기$', key):
            return True
        
        # 테이블 데이터 마커 (지역_고용률_YYYY_N분기, 지역_증감pp_YYYY_N분기)
        if re.match(r'^[가-힣]+_(고용률|증감pp)_\d{4}_\d분기$', key):
            return True
        
        # 연도/분기 마커
        if key in ('연도', '분기'):
            return True
        
        return False
    
    def handle(self, ctx: MarkerContext) -> Optional[str]:
        """실업률/고용률 관련 마커를 처리합니다."""
        key = ctx.key
        is_unemployment_sheet = ('실업' in ctx.sheet_name or ctx.sheet_name == '실업자 수')
        is_employment_sheet = ('고용' in ctx.sheet_name or '고용률' in ctx.sheet_name)
        
        # 연도/분기 마커
        if key == '연도':
            return str(ctx.year)
        if key == '분기':
            return str(ctx.quarter)
        
        # 전국 증감률/증감pp
        if key in ('전국_증감률', '전국_증감pp'):
            if is_employment_sheet:
                return self._handle_national_employment_change(ctx)
            return self._handle_national_unemployment_change(ctx)
        
        # 전국 고용률
        if key == '전국_고용률':
            return self._handle_national_employment_rate(ctx)
        
        # 상승/하락 시도수
        if key == '상승_시도수':
            return self._handle_region_count(ctx, is_increase=True, is_employment=is_employment_sheet)
        if key == '하락_시도수':
            return self._handle_region_count(ctx, is_increase=False, is_employment=is_employment_sheet)
        
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
            if is_employment_sheet:
                return self._handle_employment_region_change_pp(ctx, region_name)
            return self._handle_region_change_pp(ctx, region_name)
        
        # 전국 연령별 증감pp (전국_30대_증감pp 등)
        national_age_match = re.match(r'^전국_(\d+대(?:이상)?)_증감pp$', key)
        if national_age_match:
            age_group = national_age_match.group(1)
            return self._handle_national_age_group_change(ctx, age_group)
        
        # 상위/하위 시도
        top_match = re.match(r'상위시도(\d+)_(.+)', key)
        if top_match:
            if is_employment_sheet:
                return self._handle_top_employment_region(ctx, top_match)
            return self._handle_top_unemployment_region(ctx, top_match)
        
        bottom_match = re.match(r'하위시도(\d+)_(.+)', key)
        if bottom_match:
            if is_employment_sheet:
                return self._handle_bottom_employment_region(ctx, bottom_match)
            return self._handle_bottom_unemployment_region(ctx, bottom_match)
        
        # 테이블 헤더 마커
        header_match = re.match(r'^(\d{4})_(\d)분기$', key)
        if header_match:
            year = header_match.group(1)
            quarter = header_match.group(2)
            return f"{year}.{quarter}/4"
        
        # 테이블 데이터 마커
        table_data_match = re.match(r'^([가-힣]+)_(고용률|증감pp)_(\d{4})_(\d)분기$', key)
        if table_data_match:
            return self._handle_table_data_marker(ctx, table_data_match)
        
        return None
    
    # ==================== 고용률 관련 메서드 ====================
    
    def _get_employment_rate_regions_data(self, ctx: MarkerContext) -> List[Dict[str, Any]]:
        """고용률 테이블에서 모든 지역의 증감률(퍼센트포인트)을 계산합니다."""
        from ..utils.region_utils import SIDO_LIST
        from ..utils.sheet_utils import get_actual_sheet_name
        
        actual_sheet_name = get_actual_sheet_name('고용률')
        sheet = ctx.excel_extractor.get_sheet(actual_sheet_name)
        if not sheet:
            return []
        
        sheet_config = ctx.schema_loader.load_sheet_config('고용률')
        employment_table_config = sheet_config.get('employment_rate_table', {})
        
        if not employment_table_config.get('enabled', True):
            return []
        
        start_row = employment_table_config.get('start_row', 4)
        region_column = employment_table_config.get('region_column', 2)
        age_group_column = employment_table_config.get('age_group_column', 4)
        age_group_filter = employment_table_config.get('age_group_filter', '계')
        region_mapping = employment_table_config.get('region_mapping', {})
        
        results = []
        
        # 분기 열 찾기
        current_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year, ctx.quarter)
        prev_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year - 1, ctx.quarter)
        
        if not current_col or not prev_col:
            return []
        
        current_region = None
        seen_regions = set()
        
        for row in range(start_row, min(5000, sheet.max_row + 1)):
            cell_region = sheet.cell(row=row, column=region_column)
            cell_age = sheet.cell(row=row, column=age_group_column)
            
            if cell_region.value:
                current_region = str(cell_region.value).strip()
            
            if cell_age.value and current_region:
                age_str = str(cell_age.value).strip()
                
                if age_str == age_group_filter:
                    # 짧은 지역명으로 변환
                    reverse_mapping = {v: k for k, v in region_mapping.items()}
                    short_name = reverse_mapping.get(current_region, current_region)
                    
                    # 이미 처리한 지역은 스킵
                    if short_name in seen_regions:
                        continue
                    
                    # 전국 제외, 시도만 포함
                    if short_name != '전국' and short_name in SIDO_LIST:
                        current_value = sheet.cell(row=row, column=current_col).value
                        prev_value = sheet.cell(row=row, column=prev_col).value
                        
                        if current_value is not None and prev_value is not None:
                            try:
                                current_num = float(current_value)
                                prev_num = float(prev_value)
                                growth_pp = current_num - prev_num  # 퍼센트포인트 차이
                                
                                results.append({
                                    'name': short_name,
                                    'growth_rate': growth_pp,
                                    'current': current_num,
                                    'prev': prev_num,
                                    'row': row
                                })
                                seen_regions.add(short_name)
                            except (ValueError, TypeError):
                                pass
        
        return results
    
    def _get_employment_age_data_for_region(self, ctx: MarkerContext, region_name: str) -> List[Dict[str, Any]]:
        """특정 지역의 연령별 증감률 데이터를 반환합니다."""
        from ..utils.sheet_utils import get_actual_sheet_name
        
        actual_sheet_name = get_actual_sheet_name('고용률')
        sheet = ctx.excel_extractor.get_sheet(actual_sheet_name)
        if not sheet:
            return []
        
        sheet_config = ctx.schema_loader.load_sheet_config('고용률')
        employment_table_config = sheet_config.get('employment_rate_table', {})
        
        start_row = employment_table_config.get('start_row', 4)
        region_column = employment_table_config.get('region_column', 2)
        age_group_column = employment_table_config.get('age_group_column', 4)
        age_group_filter = employment_table_config.get('age_group_filter', '계')
        region_mapping = employment_table_config.get('region_mapping', {})
        age_groups = employment_table_config.get('age_groups', {})
        age_display_names = employment_table_config.get('age_display_names', {})
        
        # 분기 열 찾기
        current_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year, ctx.quarter)
        prev_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year - 1, ctx.quarter)
        
        if not current_col or not prev_col:
            return []
        
        # 긴 지역명으로 변환
        actual_region_name = region_mapping.get(region_name, region_name)
        
        results = []
        current_region = None
        found_region = False
        
        for row in range(start_row, min(5000, sheet.max_row + 1)):
            cell_region = sheet.cell(row=row, column=region_column)
            cell_age = sheet.cell(row=row, column=age_group_column)
            
            if cell_region.value:
                region_str = str(cell_region.value).strip()
                # 지역명 일치 확인 (짧은 이름 또는 긴 이름)
                if region_str == region_name or region_str == actual_region_name:
                    current_region = region_str
                    found_region = True
                elif found_region:
                    # 다른 지역이 나타나면 종료
                    break
            
            if found_region and cell_age.value:
                age_str = str(cell_age.value).strip()
                
                # '계'는 제외 (총합)
                if age_str == age_group_filter:
                    continue
                
                # 연령대 매핑 확인
                display_name = None
                for age_key, aliases in age_groups.items():
                    if age_str in aliases or age_str == age_key:
                        display_name = age_display_names.get(age_key, age_key)
                        break
                
                if not display_name:
                    display_name = age_str
                
                current_value = sheet.cell(row=row, column=current_col).value
                prev_value = sheet.cell(row=row, column=prev_col).value
                
                if current_value is not None and prev_value is not None:
                    try:
                        current_num = float(current_value)
                        prev_num = float(prev_value)
                        growth_pp = current_num - prev_num
                        
                        results.append({
                            'name': display_name,
                            'original_name': age_str,
                            'growth_rate': growth_pp,
                            'current': current_num,
                            'prev': prev_num
                        })
                    except (ValueError, TypeError):
                        pass
        
        # 증감률 절대값 기준 정렬 (가장 영향력 큰 연령대 순)
        results.sort(key=lambda x: abs(x['growth_rate']), reverse=True)
        
        return results
    
    def _handle_national_employment_rate(self, ctx: MarkerContext) -> Optional[str]:
        """전국 고용률 값을 반환합니다."""
        from ..utils.sheet_utils import get_actual_sheet_name
        
        actual_sheet_name = get_actual_sheet_name('고용률')
        sheet = ctx.excel_extractor.get_sheet(actual_sheet_name)
        if not sheet:
            return None
        
        sheet_config = ctx.schema_loader.load_sheet_config('고용률')
        employment_table_config = sheet_config.get('employment_rate_table', {})
        
        start_row = employment_table_config.get('start_row', 4)
        region_column = employment_table_config.get('region_column', 2)
        age_group_column = employment_table_config.get('age_group_column', 4)
        age_group_filter = employment_table_config.get('age_group_filter', '계')
        
        current_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year, ctx.quarter)
        if not current_col:
            return None
        
        current_region = None
        
        for row in range(start_row, min(5000, sheet.max_row + 1)):
            cell_region = sheet.cell(row=row, column=region_column)
            cell_age = sheet.cell(row=row, column=age_group_column)
            
            if cell_region.value:
                current_region = str(cell_region.value).strip()
            
            if cell_age.value and current_region == '전국':
                age_str = str(cell_age.value).strip()
                
                if age_str == age_group_filter:
                    value = sheet.cell(row=row, column=current_col).value
                    if value is not None:
                        return ctx.formatter.format_percentage(value, decimal_places=1, include_percent=False)
        
        return None
    
    def _handle_national_employment_change(self, ctx: MarkerContext) -> Optional[str]:
        """전국 고용률 증감을 처리합니다."""
        from ..utils.sheet_utils import get_actual_sheet_name
        
        actual_sheet_name = get_actual_sheet_name('고용률')
        sheet = ctx.excel_extractor.get_sheet(actual_sheet_name)
        if not sheet:
            return None
        
        sheet_config = ctx.schema_loader.load_sheet_config('고용률')
        employment_table_config = sheet_config.get('employment_rate_table', {})
        
        start_row = employment_table_config.get('start_row', 4)
        region_column = employment_table_config.get('region_column', 2)
        age_group_column = employment_table_config.get('age_group_column', 4)
        age_group_filter = employment_table_config.get('age_group_filter', '계')
        
        current_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year, ctx.quarter)
        prev_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year - 1, ctx.quarter)
        
        if not current_col or not prev_col:
            return None
        
        current_region = None
        
        for row in range(start_row, min(5000, sheet.max_row + 1)):
            cell_region = sheet.cell(row=row, column=region_column)
            cell_age = sheet.cell(row=row, column=age_group_column)
            
            if cell_region.value:
                current_region = str(cell_region.value).strip()
            
            if cell_age.value and current_region == '전국':
                age_str = str(cell_age.value).strip()
                
                if age_str == age_group_filter:
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
    
    def _handle_national_age_group_change(self, ctx: MarkerContext, age_group: str) -> Optional[str]:
        """전국 특정 연령대의 증감pp를 반환합니다."""
        age_data = self._get_employment_age_data_for_region(ctx, '전국')
        
        # 연령대 이름 매핑 (마커 이름 -> 시트 이름)
        age_mapping = {
            '30대': ['30~39세', '30-39세', '30 ~ 39세', '30~39'],
            '40대': ['40~49세', '40-49세', '40 ~ 49세', '40~49'],
            '60대이상': ['60세이상', '60세 이상', '60 세이상']
        }
        
        target_ages = age_mapping.get(age_group, [age_group])
        
        for data in age_data:
            if data['original_name'] in target_ages or data['name'] in target_ages:
                return ctx.formatter.format_percentage(data['growth_rate'], decimal_places=1, include_percent=False)
        
        return None
    
    def _handle_employment_region_change_pp(self, ctx: MarkerContext, region_name: str) -> Optional[str]:
        """고용률 지역별 증감pp를 반환합니다."""
        regions_data = self._get_employment_rate_regions_data(ctx)
        
        for region in regions_data:
            if region['name'] == region_name:
                return ctx.formatter.format_percentage(region['growth_rate'], decimal_places=1, include_percent=False)
        
        return None
    
    def _handle_region_count(self, ctx: MarkerContext, is_increase: bool, is_employment: bool) -> Optional[str]:
        """상승/하락 시도 수를 반환합니다."""
        if is_employment:
            regions_data = self._get_employment_rate_regions_data(ctx)
        else:
            regions_data = self._get_unemployment_rate_regions_data(ctx)
        
        if not regions_data:
            return None
        
        if is_increase:
            count = sum(1 for r in regions_data if r.get('growth_rate', 0) > 0)
        else:
            count = sum(1 for r in regions_data if r.get('growth_rate', 0) < 0)
        
        return str(count)
    
    def _handle_top_employment_region(self, ctx: MarkerContext, match) -> Optional[str]:
        """상위 고용률 지역을 처리합니다."""
        idx = int(match.group(1)) - 1
        field = match.group(2)
        
        regions_data = self._get_employment_rate_regions_data(ctx)
        if not regions_data:
            return None
        
        # 상위시도 = 고용률이 가장 많이 상승한 지역
        regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0), reverse=True)
        
        if regions and idx < len(regions):
            return self._extract_employment_region_field(ctx, regions[idx], field)
        
        return None
    
    def _handle_bottom_employment_region(self, ctx: MarkerContext, match) -> Optional[str]:
        """하위 고용률 지역을 처리합니다."""
        idx = int(match.group(1)) - 1
        field = match.group(2)
        
        regions_data = self._get_employment_rate_regions_data(ctx)
        if not regions_data:
            return None
        
        # 하위시도 = 고용률이 가장 많이 하락한 지역
        regions = sorted(regions_data, key=lambda x: x.get('growth_rate', 0))
        
        if regions and idx < len(regions):
            return self._extract_employment_region_field(ctx, regions[idx], field)
        
        return None
    
    def _extract_employment_region_field(self, ctx: MarkerContext, region: Dict[str, Any], field: str) -> Optional[str]:
        """고용률 지역 데이터에서 필드 값을 추출합니다."""
        if field == '이름':
            return region.get('name', '')
        elif field in ('증감률', '증감pp'):
            return ctx.formatter.format_percentage(region.get('growth_rate', 0), decimal_places=1, include_percent=False)
        elif field == '방향':
            return ctx.formatter.get_growth_direction(region['growth_rate'], direction_type="rise_fall", expression_key="rate")
        elif field == '변화표현':
            return ctx.formatter.get_production_change_expression(region['growth_rate'], direction_type="rise_fall")
        
        # 업태(연령별) 필드 처리 - 고용률에서 업태는 연령대
        age_match = re.match(r'업태(\d+)_(.+)', field)
        if age_match:
            age_idx = int(age_match.group(1)) - 1
            age_field = age_match.group(2)
            
            # 해당 지역의 연령별 데이터 가져오기
            age_data = self._get_employment_age_data_for_region(ctx, region['name'])
            
            if age_data and age_idx < len(age_data):
                age_item = age_data[age_idx]
                
                if age_field == '이름':
                    return age_item.get('name', '')
                elif age_field == '증감pp':
                    return ctx.formatter.format_percentage(age_item.get('growth_rate', 0), decimal_places=1, include_percent=False)
        
        return None
    
    def _handle_table_data_marker(self, ctx: MarkerContext, match) -> Optional[str]:
        """테이블 데이터 마커를 처리합니다 (지역_고용률_YYYY_N분기 형식)."""
        from ..utils.sheet_utils import get_actual_sheet_name
        
        region_name = match.group(1)
        value_type = match.group(2)  # '고용률' 또는 '증감pp'
        target_year = int(match.group(3))
        target_quarter = int(match.group(4))
        
        actual_sheet_name = get_actual_sheet_name('고용률')
        sheet = ctx.excel_extractor.get_sheet(actual_sheet_name)
        if not sheet:
            return None
        
        sheet_config = ctx.schema_loader.load_sheet_config('고용률')
        employment_table_config = sheet_config.get('employment_rate_table', {})
        
        start_row = employment_table_config.get('start_row', 4)
        region_column = employment_table_config.get('region_column', 2)
        age_group_column = employment_table_config.get('age_group_column', 4)
        age_group_filter = employment_table_config.get('age_group_filter', '계')
        region_mapping = employment_table_config.get('region_mapping', {})
        
        # 대상 분기 열 찾기
        target_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, target_year, target_quarter)
        if not target_col:
            return None
        
        # 긴 지역명으로 변환
        actual_region_name = region_mapping.get(region_name, region_name)
        
        current_region = None
        
        for row in range(start_row, min(5000, sheet.max_row + 1)):
            cell_region = sheet.cell(row=row, column=region_column)
            cell_age = sheet.cell(row=row, column=age_group_column)
            
            if cell_region.value:
                current_region = str(cell_region.value).strip()
            
            # 지역명 일치 확인
            if cell_age.value and (current_region == region_name or current_region == actual_region_name):
                age_str = str(cell_age.value).strip()
                
                if age_str == age_group_filter:
                    if value_type == '고용률':
                        value = sheet.cell(row=row, column=target_col).value
                        if value is not None:
                            return ctx.formatter.format_percentage(value, decimal_places=1, include_percent=False)
                    elif value_type == '증감pp':
                        # 전년 동분기 열 찾기
                        prev_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, target_year - 1, target_quarter)
                        if prev_col:
                            current_value = sheet.cell(row=row, column=target_col).value
                            prev_value = sheet.cell(row=row, column=prev_col).value
                            
                            if current_value is not None and prev_value is not None:
                                try:
                                    growth_pp = float(current_value) - float(prev_value)
                                    return ctx.formatter.format_percentage(growth_pp, decimal_places=1, include_percent=False)
                                except (ValueError, TypeError):
                                    pass
                    break
        
        return None
    
    # ==================== 실업률 관련 메서드 ====================
    
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
        """전국 실업률 증감을 처리합니다."""
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
        
        if is_unemployment:
            actual_sheet_name = '실업자 수'
            table_config_key = 'unemployment_rate_table'
        else:
            from ..utils.sheet_utils import get_actual_sheet_name
            actual_sheet_name = get_actual_sheet_name('고용률')
            table_config_key = 'employment_rate_table'
        
        sheet = ctx.excel_extractor.get_sheet(actual_sheet_name)
        if not sheet:
            return None
        
        # 분기 열 찾기
        current_col = ctx.dynamic_parser.get_quarter_column(actual_sheet_name, ctx.year, ctx.quarter)
        if not current_col:
            return None
        
        sheet_config = ctx.schema_loader.load_sheet_config(actual_sheet_name if is_unemployment else '고용률')
        table_config = sheet_config.get(table_config_key, {})
        start_row = table_config.get('start_row', 81 if is_unemployment else 4)
        region_column = table_config.get('region_column', 1 if is_unemployment else 2)
        age_group_column = table_config.get('age_group_column', 2 if is_unemployment else 4)
        region_mapping = table_config.get('region_mapping', {})
        
        actual_region_name = region_mapping.get(region_name, region_name)
        current_region = None
        
        for row in range(start_row, min(5000, sheet.max_row + 1)):
            cell_region = sheet.cell(row=row, column=region_column)
            cell_age = sheet.cell(row=row, column=age_group_column)
            
            if cell_region.value:
                current_region = str(cell_region.value).strip()
            
            if cell_age.value and (current_region == actual_region_name or current_region == region_name):
                age_str = str(cell_age.value).strip()
                
                if age_str == '계':
                    value = sheet.cell(row=row, column=current_col).value
                    if value is not None:
                        return ctx.formatter.format_percentage(value, decimal_places=1, include_percent=False)
        
        return None
    
    def _handle_region_change_pp(self, ctx: MarkerContext, region_name: str) -> Optional[str]:
        """실업률 지역별 증감pp (퍼센트포인트)를 반환합니다."""
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
        """실업률 지역 데이터에서 필드 값을 추출합니다."""
        if field == '이름':
            return region.get('name', '')
        elif field in ('증감률', '증감pp'):
            return ctx.formatter.format_percentage(region.get('growth_rate', 0), decimal_places=1, include_percent=False)
        elif field == '방향':
            return ctx.formatter.get_growth_direction(region['growth_rate'], direction_type="rise_fall", expression_key="rate")
        elif field == '변화표현':
            return ctx.formatter.get_production_change_expression(region['growth_rate'], direction_type="rise_fall")
        
        return None
