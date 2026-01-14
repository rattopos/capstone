# -*- coding: utf-8 -*-
"""
Qt6 애플리케이션 컨트롤러
비즈니스 로직 처리
"""

import re
from pathlib import Path
from typing import Optional, Tuple, List, Dict, Any

from data_converter import DataConverter
from utils.excel_utils import extract_year_quarter_from_raw
from services.report_generator import (
    generate_report_html,
    generate_regional_report_html,
    generate_statistics_report_html,
    generate_individual_statistics_html
)
from services.summary_data import (
    get_summary_overview_data,
    get_summary_table_data,
    get_production_summary_data,
    get_consumption_construction_data,
    get_trade_price_data,
    get_employment_population_data
)
from config.reports import (
    REPORT_ORDER, SUMMARY_REPORTS, SECTOR_REPORTS, REGIONAL_REPORTS, STATISTICS_REPORTS
)
from config.settings import TEMPLATES_DIR


class AppController:
    """애플리케이션 컨트롤러"""
    
    def __init__(self):
        self.raw_excel_path: Optional[str] = None
        self.year: Optional[int] = None
        self.quarter: Optional[int] = None
        self.file_type: Optional[str] = None
    
    def handle_file_upload(self, filepath: str) -> Tuple[bool, str]:
        """파일 업로드 처리
        
        Args:
            filepath: 업로드된 엑셀 파일 경로
            
        Returns:
            (성공 여부, 메시지)
        """
        try:
            filepath = Path(filepath)
            if not filepath.exists():
                return False, "파일을 찾을 수 없습니다."
            
            if not filepath.suffix.lower() in ['.xlsx', '.xls']:
                return False, "엑셀 파일만 업로드 가능합니다."
            
            # 연도/분기 추출
            year, quarter = extract_year_quarter_from_raw(str(filepath))
            
            # DataConverter로 검증
            try:
                converter = DataConverter(str(filepath))
                self.raw_excel_path = str(filepath)
                self.year = year
                self.quarter = quarter
                self.file_type = 'raw_direct'
                
                return True, f"파일 업로드 완료: {year}년 {quarter}분기"
            except Exception as e:
                return False, f"파일 처리 오류: {str(e)}"
                
        except Exception as e:
            return False, f"파일 업로드 실패: {str(e)}"
    
    def generate_all_reports(self) -> List[Dict[str, Any]]:
        """모든 보도자료 생성
        
        Returns:
            생성된 보도자료 페이지 리스트
        """
        if not self.raw_excel_path:
            return []
        
        pages = []
        
        # 요약 보도자료 생성
        for report_config in SUMMARY_REPORTS:
            try:
                html_content, error, missing_fields = self._generate_report(report_config)
                if html_content:
                    pages.append({
                        'id': report_config['id'],
                        'title': report_config['name'],
                        'category': report_config.get('category', 'summary'),
                        'html': html_content
                    })
            except Exception as e:
                print(f"[오류] {report_config['name']} 생성 실패: {e}")
        
        # 부문별 보도자료 생성
        for report_config in REPORT_ORDER:
            if report_config.get('category') not in ['summary']:
                try:
                    html_content, error, missing_fields = self._generate_report(report_config)
                    if html_content:
                        pages.append({
                            'id': report_config['id'],
                            'title': report_config['name'],
                            'category': report_config.get('category', 'sectoral'),
                            'html': html_content
                        })
                except Exception as e:
                    print(f"[오류] {report_config['name']} 생성 실패: {e}")
        
        return pages
    
    def _generate_report(self, report_config: Dict[str, Any]) -> Tuple[Optional[str], Optional[str], List[str]]:
        """개별 보도자료 생성"""
        try:
            report_id = report_config.get('id', '')
            
            if report_id in ['cover', 'guide', 'toc']:
                # 정적 보도자료는 별도 처리
                return self._generate_static_report(report_config)
            elif report_id.startswith('region_'):
                # 시도별 보도자료
                region_name = report_config.get('name', '')
                html_content, error = generate_regional_report_html(
                    self.raw_excel_path,
                    region_name,
                    is_reference=False
                )
                return html_content, error, []
            elif report_id.startswith('stat_'):
                # 통계표 보도자료
                result = generate_individual_statistics_html(
                    self.raw_excel_path,
                    report_config,
                    self.year,
                    self.quarter,
                    raw_excel_path=self.raw_excel_path
                )
                # 반환값이 (html_content, error) 또는 (html_content, error, missing_fields)일 수 있음
                if len(result) == 2:
                    html_content, error = result
                    return html_content, error, []
                else:
                    html_content, error, missing_fields = result
                    return html_content, error, missing_fields
            else:
                # 일반 데이터 기반 보도자료
                return generate_report_html(
                    self.raw_excel_path,
                    report_config,
                    self.year,
                    self.quarter,
                    custom_data=None,
                    raw_excel_path=self.raw_excel_path
                )
        except Exception as e:
            import traceback
            traceback.print_exc()
            return None, str(e), []
    
    def _generate_static_report(self, report_config: Dict[str, Any]) -> Tuple[Optional[str], Optional[str], List[str]]:
        """정적 보도자료 생성 (표지, 일러두기, 목차)"""
        from jinja2 import Template
        
        template_path = TEMPLATES_DIR / report_config['template']
        if not template_path.exists():
            return None, f"템플릿을 찾을 수 없습니다: {report_config['template']}", []
        
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        report_data = {
            'report_info': {
                'year': self.year,
                'quarter': self.quarter,
                'organization': '국가데이터처',
                'department': '경제통계심의관'
            }
        }
        
        if report_config['id'] == 'guide':
            report_data.update(self._get_guide_data())
        
        html_content = template.render(**report_data)
        return html_content, None, []
    
    def _get_guide_data(self) -> Dict[str, Any]:
        """일러두기 데이터"""
        return {
            'intro': {
                'background': '지역경제동향은 시·도별 경제 현황을 생산, 소비, 건설, 수출입, 물가, 고용, 인구 등의 주요 경제지표를 통하여 분석한 자료입니다.',
                'purpose': '지역경제의 동향 파악과 지역개발정책 수립 및 평가의 기초자료로 활용하고자 작성합니다.'
            },
            'content': {
                'description': f'본 보도자료는 {self.year}년 {self.quarter}/4분기 시·도별 지역경제동향을 수록하였습니다.',
                'indicator_note': '수록 지표는 총 7개 부문으로 다음과 같습니다.',
                'indicators': [
                    {'type': '생산', 'stat_items': ['광공업생산지수', '서비스업생산지수']},
                    {'type': '소비', 'stat_items': ['소매판매액지수']},
                    {'type': '건설', 'stat_items': ['건설수주액']},
                    {'type': '수출입', 'stat_items': ['수출액', '수입액']},
                    {'type': '물가', 'stat_items': ['소비자물가지수']},
                    {'type': '고용', 'stat_items': ['고용률', '실업률']},
                    {'type': '인구', 'stat_items': ['국내인구이동']}
                ]
            },
            'contacts': [
                {'category': '생산', 'statistics_name': '광공업생산지수', 'department': '산업동향과', 'phone': '042-481-2161'},
                {'category': '생산', 'statistics_name': '서비스업생산지수', 'department': '서비스업동향과', 'phone': '042-481-2190'},
                {'category': '소비', 'statistics_name': '소매판매액지수', 'department': '서비스업동향과', 'phone': '042-481-2197'},
                {'category': '건설', 'statistics_name': '건설수주액', 'department': '산업동향과', 'phone': '042-481-2158'},
                {'category': '수출입', 'statistics_name': '수출·수입', 'department': '관세청 정보데이터기획담당관', 'phone': '042-481-7845'},
                {'category': '물가', 'statistics_name': '소비자물가지수', 'department': '물가동향과', 'phone': '042-481-2531'},
                {'category': '고용', 'statistics_name': '고용률, 실업률', 'department': '고용통계과', 'phone': '042-481-2265'},
                {'category': '인구', 'statistics_name': '국내인구이동', 'department': '인구추계팀', 'phone': '042-481-2514'}
            ],
            'references': [
                {'content': '본문에 수록된 자료는 국가데이터처 홈페이지(http://mods.go.kr) 및 국가통계포털(http://kosis.kr)을 통해 이용할 수 있습니다.'}
            ],
            'notes': [
                '자료에 수록된 값은 잠정치이므로 추후 수정될 수 있습니다.'
            ]
        }
    
    def generate_hwp_html(self, pages: List[Dict[str, Any]]) -> str:
        """한글 호환 HTML 생성
        
        Args:
            pages: 보도자료 페이지 리스트
            
        Returns:
            HTML 문자열
        """
        # routes/api.py의 export_hwp_import 로직 재사용
        final_html = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="ko" xml:lang="ko">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <meta name="generator" content="지역경제동향 보도자료 시스템" />
    <title>{self.year}년 {self.quarter}/4분기 지역경제동향</title>
    <style type="text/css">
        /* 한글 호환 기본 스타일 */
        @page {{
            size: A4 portrait;
            margin: 20mm 15mm 20mm 15mm;
        }}
        
        body {{
            font-family: '맑은 고딕', 'Malgun Gothic', '바탕', 'Batang', serif;
            font-size: 10pt;
            line-height: 160%;
            color: #000000;
            background-color: #ffffff;
            margin: 0;
            padding: 0;
        }}
        
        /* 페이지 컨테이너 */
        .page-container {{
            width: 180mm;
            margin: 0 auto;
            padding: 10mm 0;
            page-break-after: always;
        }}
        
        .page-container:last-child {{
            page-break-after: auto;
        }}
        
        /* 제목 스타일 */
        h1 {{
            font-family: '맑은 고딕', 'Malgun Gothic', sans-serif;
            font-size: 18pt;
            font-weight: bold;
            color: #000000;
            margin: 0 0 15px 0;
            padding: 0;
            line-height: 140%;
        }}
        
        h2 {{
            font-family: '맑은 고딕', 'Malgun Gothic', sans-serif;
            font-size: 14pt;
            font-weight: bold;
            color: #000000;
            margin: 20px 0 10px 0;
            padding: 8px 10px;
            background-color: #f0f0f0;
            border-left: 4px solid #0066cc;
        }}
        
        h3 {{
            font-family: '맑은 고딕', 'Malgun Gothic', sans-serif;
            font-size: 12pt;
            font-weight: bold;
            color: #000000;
            margin: 15px 0 8px 0;
        }}
        
        h4 {{
            font-family: '맑은 고딕', 'Malgun Gothic', sans-serif;
            font-size: 11pt;
            font-weight: bold;
            color: #000000;
            margin: 10px 0 5px 0;
        }}
        
        /* 문단 스타일 */
        p {{
            font-family: '맑은 고딕', 'Malgun Gothic', sans-serif;
            font-size: 10pt;
            margin: 5px 0;
            line-height: 160%;
            text-align: justify;
        }}
        
        /* 표 스타일 - 한글 완벽 호환 */
        table {{
            border-collapse: collapse;
            width: 100%;
            margin: 10px 0;
            font-size: 9pt;
            border: 1px solid #000000;
            table-layout: fixed;
        }}
        
        th {{
            border: 1px solid #000000;
            padding: 6px 4px;
            text-align: center;
            vertical-align: middle;
            background-color: #d9d9d9;
            font-weight: bold;
            font-family: '맑은 고딕', 'Malgun Gothic', sans-serif;
        }}
        
        td {{
            border: 1px solid #000000;
            padding: 5px 4px;
            text-align: center;
            vertical-align: middle;
            font-family: '맑은 고딕', 'Malgun Gothic', sans-serif;
        }}
        
        /* 리스트 스타일 */
        ul, ol {{
            margin: 10px 0 10px 20px;
            padding: 0;
        }}
        
        li {{
            margin: 3px 0;
            line-height: 160%;
        }}
        
        /* 페이지 번호 */
        .page-number {{
            text-align: center;
            font-size: 9pt;
            color: #666666;
            margin-top: 20px;
            padding-top: 10px;
            border-top: 1px solid #cccccc;
        }}
        
        /* 인쇄 스타일 */
        @media print {{
            body {{
                background-color: #ffffff;
            }}
            .page-container {{
                page-break-after: always;
            }}
        }}
    </style>
</head>
<body>
'''
        
        # 각 페이지 처리
        for idx, page in enumerate(pages, 1):
            page_html = page.get('html', '')
            page_title = page.get('title', f'페이지 {idx}')
            category = page.get('category', '')
            
            # body 내용 추출
            body_content = page_html
            if '<body' in page_html.lower():
                body_match = re.search(r'<body[^>]*>(.*?)</body>', page_html, re.DOTALL | re.IGNORECASE)
                if body_match:
                    body_content = body_match.group(1)
            
            # 불필요한 태그 제거
            body_content = re.sub(r'<style[^>]*>.*?</style>', '', body_content, flags=re.DOTALL)
            body_content = re.sub(r'<script[^>]*>.*?</script>', '', body_content, flags=re.DOTALL)
            body_content = re.sub(r'<link[^>]*/?>', '', body_content)
            body_content = re.sub(r'<meta[^>]*/?>', '', body_content)
            body_content = re.sub(r'<!DOCTYPE[^>]*>', '', body_content)
            body_content = re.sub(r'<html[^>]*>', '', body_content)
            body_content = re.sub(r'</html>', '', body_content)
            body_content = re.sub(r'<head[^>]*>.*?</head>', '', body_content, flags=re.DOTALL)
            
            # 그림 제거: canvas, img, svg 태그 제거
            body_content = self._remove_images(body_content)
            
            # 한글 호환 스타일 적용
            body_content = self._add_hwp_compatible_styles(body_content)
            
            # 카테고리 한글명
            category_names = {
                'summary': '요약',
                'sectoral': '부문별',
                'regional': '시도별',
                'statistics': '통계표'
            }
            category_name = category_names.get(category, '')
            
            # 페이지 래퍼 추가
            final_html += f'''
    <!-- 페이지 {idx}: {page_title} -->
    <div class="page-container">
        <h2>[{category_name}] {page_title}</h2>
        {body_content}
        <p class="page-number">- {idx} / {len(pages)} -</p>
    </div>
'''
        
        final_html += '''
</body>
</html>
'''
        
        return final_html
    
    def _remove_images(self, html_content: str) -> str:
        """HTML에서 그림(차트, 이미지) 제거"""
        # canvas 태그 제거
        html_content = re.sub(r'<canvas[^>]*>.*?</canvas>', '', html_content, flags=re.DOTALL)
        html_content = re.sub(r'<canvas[^>]*/?>', '', html_content)
        
        # img 태그 제거 (base64 포함)
        html_content = re.sub(r'<img[^>]*>', '', html_content)
        
        # svg 태그 제거
        html_content = re.sub(r'<svg[^>]*>.*?</svg>', '', html_content, flags=re.DOTALL)
        
        return html_content
    
    def _add_hwp_compatible_styles(self, html_content: str) -> str:
        """한글 프로그램 완벽 호환을 위한 인라인 스타일 추가"""
        # table 태그에 인라인 스타일 추가
        html_content = re.sub(
            r'<table([^>]*)>',
            r'<table\1 style="border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 9pt; border: 1px solid #000000; table-layout: fixed;">',
            html_content
        )
        
        # th 태그에 인라인 스타일 추가
        html_content = re.sub(
            r'<th([^>]*)>',
            r'<th\1 style="border: 1px solid #000000; padding: 6px 4px; text-align: center; vertical-align: middle; background-color: #d9d9d9; font-weight: bold; font-family: 맑은 고딕, Malgun Gothic, sans-serif;">',
            html_content
        )
        
        # td 태그에 인라인 스타일 추가
        html_content = re.sub(
            r'<td([^>]*)>',
            r'<td\1 style="border: 1px solid #000000; padding: 5px 4px; text-align: center; vertical-align: middle; font-family: 맑은 고딕, Malgun Gothic, sans-serif;">',
            html_content
        )
        
        # 제목 태그들에 인라인 스타일 추가
        html_content = re.sub(
            r'<h1([^>]*)>',
            r'<h1\1 style="font-family: 맑은 고딕, Malgun Gothic, sans-serif; font-size: 18pt; font-weight: bold; margin: 0 0 15px 0;">',
            html_content
        )
        html_content = re.sub(
            r'<h2([^>]*)>',
            r'<h2\1 style="font-family: 맑은 고딕, Malgun Gothic, sans-serif; font-size: 14pt; font-weight: bold; margin: 20px 0 10px 0;">',
            html_content
        )
        html_content = re.sub(
            r'<h3([^>]*)>',
            r'<h3\1 style="font-family: 맑은 고딕, Malgun Gothic, sans-serif; font-size: 12pt; font-weight: bold; margin: 15px 0 8px 0;">',
            html_content
        )
        html_content = re.sub(
            r'<h4([^>]*)>',
            r'<h4\1 style="font-family: 맑은 고딕, Malgun Gothic, sans-serif; font-size: 11pt; font-weight: bold; margin: 10px 0 5px 0;">',
            html_content
        )
        
        # p 태그에 스타일 추가
        html_content = re.sub(
            r'<p([^>]*)>',
            r'<p\1 style="font-family: 맑은 고딕, Malgun Gothic, sans-serif; font-size: 10pt; margin: 5px 0; line-height: 160%;">',
            html_content
        )
        
        # ul, ol 태그에 스타일 추가
        html_content = re.sub(
            r'<ul([^>]*)>',
            r'<ul\1 style="margin: 10px 0 10px 20px; font-family: 맑은 고딕, Malgun Gothic, sans-serif;">',
            html_content
        )
        html_content = re.sub(
            r'<ol([^>]*)>',
            r'<ol\1 style="margin: 10px 0 10px 20px; font-family: 맑은 고딕, Malgun Gothic, sans-serif;">',
            html_content
        )
        html_content = re.sub(
            r'<li([^>]*)>',
            r'<li\1 style="margin: 3px 0; line-height: 160%;">',
            html_content
        )
        
        return html_content
