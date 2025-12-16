"""
템플릿 생성 모듈
이미지 분석 결과를 기반으로 HTML 템플릿 생성
"""

import re
from pathlib import Path
from typing import Dict, List, Any, Optional
from .image_analyzer import ImageAnalyzer


class TemplateGenerator:
    """이미지 분석 결과로부터 HTML 템플릿을 생성하는 클래스"""
    
    def __init__(self, image_analyzer: Optional[ImageAnalyzer] = None, device: Optional[str] = None):
        """
        템플릿 생성기 초기화
        
        Args:
            image_analyzer: 이미지 분석기 인스턴스 (None이면 새로 생성)
            device: 사용할 디바이스 ('cuda', 'mps', 'cpu', None이면 자동 감지)
        """
        self.image_analyzer = image_analyzer or ImageAnalyzer(use_easyocr=True, device=device)
    
    def generate_template_from_image(
        self,
        image_path: str,
        template_name: str,
        sheet_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        이미지로부터 HTML 템플릿을 생성합니다.
        
        Args:
            image_path: 이미지 파일 경로
            template_name: 템플릿 이름
            sheet_name: 엑셀 시트명 (선택적, 자동 추론 가능)
            
        Returns:
            템플릿 정보 딕셔너리:
            - 'template_html': 생성된 HTML 템플릿
            - 'template_name': 템플릿 이름
            - 'sheet_name': 시트명
            - 'markers': 템플릿에 포함된 마커 리스트
            - 'structure': 템플릿 구조 정보
        """
        # 이미지 분석
        analysis = self.image_analyzer.analyze_image(image_path)
        template_structure = self.image_analyzer.extract_template_structure(image_path)
        
        # 시트명 추론 (이미지 이름이나 텍스트에서)
        if not sheet_name:
            sheet_name = self._infer_sheet_name(image_path, analysis['full_text'])
        
        # HTML 템플릿 생성
        html_template = self._generate_html(
            analysis,
            template_structure,
            template_name,
            sheet_name
        )
        
        # 마커 추출
        markers = self._extract_markers_from_html(html_template)
        
        return {
            'template_html': html_template,
            'template_name': template_name,
            'sheet_name': sheet_name,
            'markers': markers,
            'structure': template_structure,
            'image_path': image_path
        }
    
    def _infer_sheet_name(self, image_path: str, full_text: str) -> str:
        """이미지 경로나 텍스트에서 시트명 추론"""
        image_name = Path(image_path).stem
        
        # 이미지 이름에서 시트명 추론
        sheet_name_mapping = {
            '건설수주': '건설 (공표자료)',
            '고용률': '고용률',
            '광공업생산': '광공업생산',
            '국내인구이동': '시도 간 이동',
            '물가동향': '지출목적별 물가',
            '서비스업생산': '서비스업생산',
            '소매판매': '소비(소매, 추가)',
            '수입': '수입',
            '수출': '수출',
            '실업률': '실업자 수'
        }
        
        if image_name in sheet_name_mapping:
            return sheet_name_mapping[image_name]
        
        # 텍스트에서 키워드로 추론
        keywords = {
            '건설': '건설 (공표자료)',
            '고용률': '고용률',
            '광공업': '광공업생산',
            '서비스업': '서비스업생산',
            '소매': '소비(소매, 추가)',
            '수입': '수입',
            '수출': '수출',
            '실업': '실업자 수',
            '물가': '지출목적별 물가',
            '인구이동': '시도 간 이동'
        }
        
        for keyword, sheet in keywords.items():
            if keyword in full_text or keyword in image_name:
                return sheet
        
        # 기본값
        return '광공업생산'
    
    def _generate_html(
        self,
        analysis: Dict[str, Any],
        template_structure: Dict[str, Any],
        template_name: str,
        sheet_name: str
    ) -> str:
        """이미지 분석 결과로부터 HTML 템플릿 생성"""
        html_parts = [
            '<!DOCTYPE html>',
            '<html lang="ko">',
            '<head>',
            '<meta charset="UTF-8">',
            '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
            f'<title>{template_name}</title>',
            '<style>',
            self._generate_css(),
            '</style>',
            '</head>',
            '<body>'
        ]
        
        # 레이아웃 분석 결과를 기반으로 HTML 구조 생성
        layout = analysis.get('layout', {})
        rows = layout.get('rows', [])
        
        # 텍스트 영역을 기반으로 HTML 생성
        for row_idx, row in enumerate(rows):
            row_html = self._generate_row_html(row, sheet_name, row_idx)
            html_parts.append(row_html)
        
        # 테이블이 있으면 테이블 HTML 생성
        tables = analysis.get('tables', [])
        for table in tables:
            table_html = self._generate_table_html(table, sheet_name)
            html_parts.append(table_html)
        
        html_parts.extend(['</body>', '</html>'])
        
        return '\n'.join(html_parts)
    
    def _generate_css(self) -> str:
        """기본 CSS 스타일 생성"""
        return """
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Malgun Gothic', '맑은 고딕', sans-serif;
            line-height: 1.6;
            color: #000;
            background-color: #fff;
            padding: 30px 40px;
            max-width: 1000px;
            margin: 0 auto;
            font-size: 14px;
        }
        
        .document-title {
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 20px;
        }
        
        .section-title {
            font-size: 16px;
            font-weight: bold;
            margin: 25px 0 12px 0;
        }
        
        .subsection-title {
            font-size: 14px;
            font-weight: bold;
            margin: 20px 0 10px 0;
        }
        
        ul {
            margin: 12px 0;
            padding-left: 0;
        }
        
        li {
            margin-bottom: 8px;
            line-height: 1.7;
        }
        
        .percentage {
            font-weight: bold;
        }
        
        .percentage.positive {
            color: #e74c3c;
        }
        
        .percentage.negative {
            color: #3498db;
        }
        """
    
    def _generate_row_html(self, row: List[Dict], sheet_name: str, row_idx: int) -> str:
        """텍스트 행을 HTML로 변환"""
        html_parts = []
        
        for region in row:
            text = region.get('text', '')
            if not text or len(text.strip()) < 2:
                continue
            
            # 텍스트에서 숫자나 키워드를 마커로 변환
            processed_text = self._convert_text_to_markers(text, sheet_name)
            
            # 문단 또는 리스트 아이템으로 생성
            if any(keyword in text for keyword in ['증가', '감소', '늘어', '줄어']):
                html_parts.append(f'<li>{processed_text}</li>')
            else:
                html_parts.append(f'<p>{processed_text}</p>')
        
        if html_parts:
            return '<ul class="bullet-list">' + ''.join(html_parts) + '</ul>'
        return ''
    
    def _convert_text_to_markers(self, text: str, sheet_name: str) -> str:
        """텍스트를 마커가 포함된 HTML로 변환"""
        # 지역명을 마커로 변환
        regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산',
                   '세종', '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        for region in regions:
            if region in text:
                # 지역명 + 증감률 패턴
                if '증감률' in text or '%' in text:
                    text = text.replace(region, f'<strong>{{{sheet_name}:{region}_증감률}}</strong>')
                else:
                    text = text.replace(region, f'<strong>{{{sheet_name}:{region}_이름}}</strong>')
        
        # 숫자 패턴을 마커로 변환
        # 예: "5.2%" -> "{sheet_name:전국_증감률}"
        number_pattern = re.compile(r'[-+]?\d+\.?\d*%?')
        matches = number_pattern.findall(text)
        
        for match in matches:
            if '%' in match or '증감률' in text or '증가' in text or '감소' in text:
                # 증감률 마커로 변환
                if '전국' in text:
                    marker = f'{{{sheet_name}:전국_증감률}}'
                elif any(region in text for region in regions):
                    for region in regions:
                        if region in text:
                            marker = f'{{{sheet_name}:{region}_증감률}}'
                            break
                    else:
                        marker = f'{{{sheet_name}:전국_증감률}}'
                else:
                    marker = f'{{{sheet_name}:전국_증감률}}'
                
                text = text.replace(match, marker, 1)
        
        # 키워드 기반 마커 생성
        if '기준연도' in text or '년' in text:
            text = re.sub(r'\d{4}년', f'<strong>{{{sheet_name}:기준연도}}년</strong>', text)
        if '기준분기' in text or '분기' in text:
            text = re.sub(r'\d/4', f'<strong>{{{sheet_name}:기준분기}}</strong>', text)
        
        return text
    
    def _generate_table_html(self, table: Dict, sheet_name: str) -> str:
        """테이블을 HTML로 변환"""
        html_parts = ['<table style="width: 100%; border-collapse: collapse; margin: 20px 0;">']
        
        # 헤더
        if 'header' in table:
            html_parts.append('<thead><tr>')
            for header in table['header']:
                html_parts.append(f'<th style="border: 1px solid #ddd; padding: 8px;">{header}</th>')
            html_parts.append('</tr></thead>')
        
        # 데이터 행
        if 'rows' in table:
            html_parts.append('<tbody>')
            for row in table['rows']:
                html_parts.append('<tr>')
                for cell in row:
                    # 셀 내용을 마커로 변환
                    processed_cell = self._convert_text_to_markers(str(cell), sheet_name)
                    html_parts.append(f'<td style="border: 1px solid #ddd; padding: 8px;">{processed_cell}</td>')
                html_parts.append('</tr>')
            html_parts.append('</tbody>')
        
        html_parts.append('</table>')
        return '\n'.join(html_parts)
    
    def _extract_markers_from_html(self, html: str) -> List[str]:
        """HTML 템플릿에서 마커 추출"""
        marker_pattern = re.compile(r'\{([^:{}]+):([^}]+)\}')
        markers = []
        
        for match in marker_pattern.finditer(html):
            markers.append(match.group(0))
        
        return list(set(markers))  # 중복 제거

