"""
차트 데이터 관련 마커 처리 핸들러
서울주요지표 등 차트가 포함된 템플릿용 데이터 생성
"""

import re
from typing import Optional, List, Tuple

from .base import MarkerHandler, MarkerContext


class ChartDataHandler(MarkerHandler):
    """차트용 시계열 데이터 및 테이블용 분기별 데이터를 생성하는 핸들러"""
    
    def can_handle(self, ctx: MarkerContext) -> bool:
        """차트 데이터 또는 분기별 데이터 마커인지 확인합니다."""
        key = ctx.key
        
        # 패턴 1: 지역_chart_data 또는 지역_증감_chart_data
        if key.endswith('_chart_data'):
            return True
        
        # 패턴 2: 지역_연도_분기 (예: 서울_2023_3분기)
        if re.match(r'^[가-힣]+_\d{4}_\d분기$', key):
            return True
        
        # 패턴 3: 지역_증감_연도_분기 (예: 서울_증감_2023_3분기)
        if re.match(r'^[가-힣]+_증감_\d{4}_\d분기$', key):
            return True
        
        return False
    
    def handle(self, ctx: MarkerContext) -> Optional[str]:
        """차트 데이터 또는 분기별 데이터 마커를 처리합니다."""
        key = ctx.key
        
        # 패턴 1: 차트 데이터 (전국_chart_data, 서울_chart_data, 전국_증감_chart_data)
        chart_match = re.match(r'^([가-힣]+)(?:_(증감))?_chart_data$', key)
        if chart_match:
            region_name = chart_match.group(1)
            use_growth_change = chart_match.group(2) == '증감'
            return self._get_chart_data(
                ctx, region_name, ctx.year, ctx.quarter, use_growth_change
            )
        
        # 패턴 2: 분기별 증감률 (서울_2023_3분기)
        quarterly_match = re.match(r'^([가-힣]+)_(\d{4})_(\d)분기$', key)
        if quarterly_match:
            region_name = quarterly_match.group(1)
            target_year = int(quarterly_match.group(2))
            target_quarter = int(quarterly_match.group(3))
            
            value = self._get_quarterly_growth_rate(ctx, region_name, target_year, target_quarter)
            if value is not None:
                return ctx.formatter.format_percentage(value, decimal_places=1, include_percent=False)
            return None
        
        # 패턴 3: 분기별 증감pp (서울_증감_2023_3분기)
        growth_match = re.match(r'^([가-힣]+)_증감_(\d{4})_(\d)분기$', key)
        if growth_match:
            region_name = growth_match.group(1)
            target_year = int(growth_match.group(2))
            target_quarter = int(growth_match.group(3))
            
            value = self._get_quarter_growth_change(ctx, region_name, target_year, target_quarter)
            if value is not None:
                return ctx.formatter.format_percentage(value, decimal_places=1, include_percent=False)
            return None
        
        return None
    
    def _get_chart_data_quarters(self, year: int, quarter: int) -> List[Tuple[int, int]]:
        """
        차트에 표시할 9개 분기 목록을 반환합니다.
        현재 분기부터 8분기 전까지 (총 9개 분기)
        
        Args:
            year: 현재 연도
            quarter: 현재 분기
            
        Returns:
            [(연도, 분기), ...] 형태의 리스트 (오래된 것부터)
        """
        quarters = []
        y, q = year, quarter
        
        # 9개 분기를 역순으로 수집 (현재 분기 포함)
        for _ in range(9):
            quarters.append((y, q))
            q -= 1
            if q == 0:
                q = 4
                y -= 1
        
        # 오래된 것부터 정렬
        quarters.reverse()
        return quarters
    
    def _get_chart_data(self, ctx: MarkerContext, region_name: str, 
                        year: int, quarter: int, use_growth_change: bool = False) -> str:
        """
        차트용 9분기 데이터를 JavaScript 배열 문자열로 반환합니다.
        
        Args:
            ctx: 마커 컨텍스트
            region_name: 지역 이름 (전국, 서울 등)
            year: 현재 연도
            quarter: 현재 분기
            use_growth_change: True면 전년동분기대비 증감(pp), False면 증감률(%)
            
        Returns:
            JavaScript 배열 형태의 문자열 (예: "1.2, 3.4, -0.5, ...")
        """
        quarters = self._get_chart_data_quarters(year, quarter)
        values = []
        
        for target_year, target_quarter in quarters:
            try:
                if use_growth_change:
                    # 고용률/실업률 등 %p 증감 계산
                    value = self._get_quarter_growth_change(
                        ctx, region_name, target_year, target_quarter
                    )
                else:
                    # 일반 증감률 계산
                    value = self._get_quarterly_growth_rate(
                        ctx, region_name, target_year, target_quarter
                    )
                
                if value is not None:
                    values.append(f"{value:.1f}")
                else:
                    values.append("null")
            except Exception:
                values.append("null")
        
        return ", ".join(values)
    
    def _get_quarterly_growth_rate(self, ctx: MarkerContext, region_name: str,
                                    target_year: int, target_quarter: int) -> Optional[float]:
        """
        특정 분기의 전년동분기대비 증감률을 계산합니다.
        지역별 데이터의 경우 분류 단계 1의 '총지수'도 허용합니다.
        """
        try:
            # 먼저 dynamic_parser 시도
            growth_rate = ctx.dynamic_parser.calculate_growth_rate(
                ctx.sheet_name, region_name, target_year, target_quarter
            )
            if growth_rate is not None:
                return growth_rate
            
            # 실패하면 직접 계산 (분류 단계 1의 총지수도 허용)
            sheet = ctx.excel_extractor.get_sheet(ctx.sheet_name)
            if not sheet:
                return None
            
            # 시트 설정 가져오기
            sheet_config = ctx.schema_loader.load_sheet_config(ctx.sheet_name)
            region_column = sheet_config.get('region_column', 2)
            classification_column = sheet_config.get('classification_column', 3)
            category_column = sheet_config.get('category_column', 6)
            
            # 열 번호 계산
            current_col = ctx.dynamic_parser.get_quarter_column(ctx.sheet_name, target_year, target_quarter)
            prev_col = ctx.dynamic_parser.get_quarter_column(ctx.sheet_name, target_year - 1, target_quarter)
            
            if current_col is None or prev_col is None:
                return None
            
            # 지역 행 찾기 (분류 단계 0 또는 1의 '총지수' 허용)
            target_row = None
            from ..utils.region_utils import normalize_region_name
            normalized_region = normalize_region_name(region_name)
            
            for row in range(4, min(2000, sheet.max_row + 1)):
                cell_region = sheet.cell(row=row, column=region_column).value
                if not cell_region:
                    continue
                
                cell_normalized = normalize_region_name(str(cell_region).strip())
                if cell_normalized != normalized_region:
                    continue
                
                # 카테고리 확인 (총지수 또는 계)
                cell_category = sheet.cell(row=row, column=category_column).value
                if cell_category:
                    cat_str = str(cell_category).strip()
                    if cat_str in ['총지수', '계', '   계']:
                        # 분류 단계 확인 (0 또는 1 허용)
                        cell_classification = sheet.cell(row=row, column=classification_column).value
                        if cell_classification in [0, 1, '0', '1', 0.0, 1.0]:
                            target_row = row
                            break
            
            if target_row is None:
                return None
            
            # 값 가져와서 증감률 계산
            current_value = sheet.cell(row=target_row, column=current_col).value
            prev_value = sheet.cell(row=target_row, column=prev_col).value
            
            if current_value is None or prev_value is None:
                return None
            
            try:
                current_float = float(current_value)
                prev_float = float(prev_value)
                if prev_float == 0:
                    return None
                return ((current_float / prev_float) - 1) * 100
            except (ValueError, TypeError):
                return None
                
        except Exception:
            return None
    
    def _get_quarter_growth_change(self, ctx: MarkerContext, region_name: str,
                                    target_year: int, target_quarter: int) -> Optional[float]:
        """
        특정 분기의 전년동분기대비 증감(pp)을 계산합니다.
        고용률, 실업률 등 %p 단위 데이터에 사용
        """
        try:
            from ..utils.sheet_utils import get_actual_sheet_name
            actual_sheet_name = get_actual_sheet_name(ctx.sheet_name)
            
            sheet = ctx.excel_extractor.get_sheet(actual_sheet_name)
            if not sheet:
                return None
            
            # 시트 설정 가져오기
            sheet_config = ctx.schema_loader.load_sheet_config(actual_sheet_name)
            region_column = sheet_config.get('region_column', 2)
            classification_column = sheet_config.get('classification_column', 3)
            category_column = sheet_config.get('category_column', 6)
            
            # 열 번호 계산 (DynamicSheetParser 사용)
            current_col = ctx.dynamic_parser.get_quarter_column(
                actual_sheet_name, target_year, target_quarter
            )
            prev_col = ctx.dynamic_parser.get_quarter_column(
                actual_sheet_name, target_year - 1, target_quarter
            )
            
            if current_col is None or prev_col is None:
                return None
            
            # 지역 행 찾기 (분류 단계 0 또는 1의 '계' 또는 지역명만 있는 행 허용)
            target_row = None
            from ..utils.region_utils import normalize_region_name
            normalized_region = normalize_region_name(region_name)
            
            for row in range(4, min(500, sheet.max_row + 1)):
                cell_region = sheet.cell(row=row, column=region_column).value
                if not cell_region:
                    continue
                
                cell_normalized = normalize_region_name(str(cell_region).strip())
                if cell_normalized != normalized_region:
                    continue
                
                # 분류 단계 확인 (0 또는 None 허용)
                classification = sheet.cell(row=row, column=classification_column).value
                
                # 카테고리 확인 (계, 합계, 또는 없음 허용)
                cell_category = sheet.cell(row=row, column=category_column).value
                cat_str = str(cell_category).strip() if cell_category else ''
                
                # 고용률/실업률 시트는 분류 0의 '계' 행 찾기
                if classification in [0, '0', 0.0, None]:
                    if cat_str in ['계', '', '   계', '합계', '총지수']:
                        target_row = row
                        break
            
            if target_row is None:
                return None
            
            # 현재 값과 전년 동분기 값 가져오기
            current_value = sheet.cell(row=target_row, column=current_col).value
            prev_value = sheet.cell(row=target_row, column=prev_col).value
            
            if current_value is None or prev_value is None:
                return None
            
            try:
                current_float = float(current_value)
                prev_float = float(prev_value)
                return current_float - prev_float
            except (ValueError, TypeError):
                return None
            
        except Exception:
            return None

