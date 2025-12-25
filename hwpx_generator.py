#!/usr/bin/env python3
"""
HWPX 파일 생성 모듈
- HTML 보고서를 HWPX (한글 2014+ 형식) 파일로 변환
- 차트/인포그래픽은 이미지로 렌더링하여 포함
"""

import os
import re
import zipfile
import base64
import hashlib
from pathlib import Path
from datetime import datetime
from bs4 import BeautifulSoup
from io import BytesIO
import json

# HWPX 파일의 기본 구조 상수
HWPX_MIMETYPE = "application/hwp+zip"

class HWPXGenerator:
    """HWPX 파일 생성기"""
    
    def __init__(self, year: int, quarter: int):
        self.year = year
        self.quarter = quarter
        self.images = {}  # 이미지 ID -> 이미지 데이터
        self.image_counter = 0
        
    def generate(self, pages: list, output_path: str) -> dict:
        """
        여러 페이지를 HWPX 파일로 생성
        
        Args:
            pages: [{"html": str, "title": str, "category": str}, ...]
            output_path: 출력 파일 경로
            
        Returns:
            {"success": bool, "filepath": str, "images_count": int}
        """
        try:
            # HWPX는 ZIP 파일
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as hwpx:
                # 1. mimetype (압축하지 않음)
                hwpx.writestr('mimetype', HWPX_MIMETYPE, compress_type=zipfile.ZIP_STORED)
                
                # 2. version.xml
                hwpx.writestr('version.xml', self._create_version_xml())
                
                # 3. settings.xml
                hwpx.writestr('settings.xml', self._create_settings_xml())
                
                # 4. META-INF/manifest.xml
                manifest_entries = []
                
                # 5. Contents 폴더 - 각 섹션별 XML 생성
                all_content = []
                for idx, page in enumerate(pages):
                    html = page.get('html', '')
                    title = page.get('title', f'페이지 {idx + 1}')
                    
                    # HTML을 HWPX XML로 변환
                    section_xml, images = self._convert_html_to_hwpx_section(html, title, idx)
                    all_content.append(section_xml)
                    
                    # 이미지 추가
                    for img_id, img_data in images.items():
                        self.images[img_id] = img_data
                
                # 전체 섹션을 하나의 section0.xml로 저장
                combined_section = self._create_section_xml(all_content)
                hwpx.writestr('Contents/section0.xml', combined_section)
                manifest_entries.append(('Contents/section0.xml', 'application/xml'))
                
                # 6. 이미지 파일 저장
                for img_id, img_data in self.images.items():
                    img_path = f'BinData/{img_id}'
                    hwpx.writestr(img_path, img_data['data'])
                    manifest_entries.append((img_path, img_data['mimetype']))
                
                # 7. content.hpf (문서 패키지 정보)
                hwpx.writestr('content.hpf', self._create_content_hpf(manifest_entries))
                
                # 8. header.xml
                hwpx.writestr('Contents/header.xml', self._create_header_xml())
                
            return {
                "success": True,
                "filepath": output_path,
                "images_count": len(self.images),
                "pages_count": len(pages)
            }
            
        except Exception as e:
            return {
                "success": False,
                "error": str(e)
            }
    
    def _create_version_xml(self) -> str:
        """버전 정보 XML 생성"""
        return '''<?xml version="1.0" encoding="UTF-8"?>
<hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2011/version">
    <hh:application version="1.0" program="capstone-report-generator"/>
</hh:version>'''

    def _create_settings_xml(self) -> str:
        """설정 XML 생성"""
        return '''<?xml version="1.0" encoding="UTF-8"?>
<hh:settings xmlns:hh="http://www.hancom.co.kr/hwpml/2011/settings">
    <hh:paperSize width="59528" height="84188"/>
    <hh:margins left="4252" right="4252" top="2835" bottom="2835" header="1417" footer="1417"/>
</hh:settings>'''

    def _create_header_xml(self) -> str:
        """헤더 정보 XML 생성"""
        return f'''<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
    <hh:docInfo>
        <hh:title>{self.year}년 {self.quarter}/4분기 지역경제동향</hh:title>
        <hh:creator>지역경제동향 보고서 시스템</hh:creator>
        <hh:date>{datetime.now().strftime("%Y-%m-%d")}</hh:date>
    </hh:docInfo>
</hh:head>'''

    def _create_content_hpf(self, entries: list) -> str:
        """문서 패키지 파일 생성"""
        items = '\n'.join([
            f'    <item href="{path}" media-type="{mtype}"/>'
            for path, mtype in entries
        ])
        return f'''<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.hancom.co.kr/hwpml/2011/package">
    <manifest>
{items}
    </manifest>
</package>'''

    def _create_section_xml(self, sections: list) -> str:
        """섹션 XML 생성"""
        content = '\n'.join(sections)
        return f'''<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
{content}
</hs:sec>'''

    def _convert_html_to_hwpx_section(self, html: str, title: str, page_idx: int) -> tuple:
        """
        HTML을 HWPX 섹션 XML로 변환
        
        Returns:
            (section_xml, images_dict)
        """
        soup = BeautifulSoup(html, 'html.parser')
        paragraphs = []
        images = {}
        
        # 페이지 구분선 추가
        if page_idx > 0:
            paragraphs.append(self._create_page_break())
        
        # 제목 추가
        paragraphs.append(self._create_title_paragraph(title, page_idx + 1))
        
        # body 내용 처리
        body = soup.find('body')
        if body:
            content_paragraphs, content_images = self._process_element(body, images)
            paragraphs.extend(content_paragraphs)
            images.update(content_images)
        
        return '\n'.join(paragraphs), images

    def _process_element(self, element, images: dict) -> tuple:
        """HTML 요소를 HWPX 단락으로 변환"""
        paragraphs = []
        
        for child in element.children:
            if isinstance(child, str):
                text = child.strip()
                if text:
                    paragraphs.append(self._create_text_paragraph(text))
            elif child.name:
                if child.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                    level = int(child.name[1])
                    text = child.get_text(strip=True)
                    if text:
                        paragraphs.append(self._create_heading_paragraph(text, level))
                        
                elif child.name == 'p':
                    text = child.get_text(strip=True)
                    if text:
                        paragraphs.append(self._create_text_paragraph(text))
                        
                elif child.name == 'table':
                    table_xml = self._convert_table(child)
                    paragraphs.append(table_xml)
                    
                elif child.name == 'img':
                    img_xml, img_data = self._process_image(child)
                    if img_xml:
                        paragraphs.append(img_xml)
                        if img_data:
                            images[img_data['id']] = img_data
                            
                elif child.name == 'svg':
                    # SVG를 이미지로 변환하여 처리
                    img_xml, img_data = self._process_svg(child)
                    if img_xml:
                        paragraphs.append(img_xml)
                        if img_data:
                            images[img_data['id']] = img_data
                            
                elif child.name == 'canvas':
                    # Canvas는 data-url로 처리
                    img_xml, img_data = self._process_canvas(child)
                    if img_xml:
                        paragraphs.append(img_xml)
                        if img_data:
                            images[img_data['id']] = img_data
                            
                elif child.name == 'div':
                    # div 내부 재귀 처리
                    div_paras, div_imgs = self._process_element(child, images)
                    paragraphs.extend(div_paras)
                    images.update(div_imgs)
                    
                elif child.name == 'ul' or child.name == 'ol':
                    list_paras = self._process_list(child)
                    paragraphs.extend(list_paras)
                    
                elif child.name in ['span', 'strong', 'b', 'em', 'i']:
                    text = child.get_text(strip=True)
                    if text:
                        paragraphs.append(self._create_text_paragraph(text))
                else:
                    # 기타 요소는 재귀 처리
                    nested_paras, nested_imgs = self._process_element(child, images)
                    paragraphs.extend(nested_paras)
                    images.update(nested_imgs)
        
        return paragraphs, images

    def _create_page_break(self) -> str:
        """페이지 구분선"""
        return '''<hp:p>
    <hp:ctrl>
        <hp:pageBreak/>
    </hp:ctrl>
</hp:p>'''

    def _create_title_paragraph(self, title: str, page_num: int) -> str:
        """제목 단락 생성"""
        return f'''<hp:p>
    <hp:run>
        <hp:t>{self._escape_xml(title)}</hp:t>
    </hp:run>
    <hp:lineseg>
        <hp:lineSegItem textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="160" horzpos="0" horzsize="48000" flags="0"/>
    </hp:lineseg>
</hp:p>'''

    def _create_heading_paragraph(self, text: str, level: int) -> str:
        """제목 단락 생성"""
        font_size = {1: 22, 2: 18, 3: 14, 4: 12, 5: 11, 6: 10}.get(level, 12)
        return f'''<hp:p>
    <hp:run>
        <hp:charPr>
            <hp:sz val="{font_size * 100}"/>
            <hp:bold val="true"/>
        </hp:charPr>
        <hp:t>{self._escape_xml(text)}</hp:t>
    </hp:run>
</hp:p>'''

    def _create_text_paragraph(self, text: str) -> str:
        """일반 텍스트 단락 생성"""
        return f'''<hp:p>
    <hp:run>
        <hp:t>{self._escape_xml(text)}</hp:t>
    </hp:run>
</hp:p>'''

    def _convert_table(self, table_element) -> str:
        """HTML 테이블을 HWPX 테이블 XML로 변환"""
        rows = []
        for tr in table_element.find_all('tr'):
            cells = []
            for cell in tr.find_all(['td', 'th']):
                text = cell.get_text(strip=True)
                colspan = int(cell.get('colspan', 1))
                rowspan = int(cell.get('rowspan', 1))
                is_header = cell.name == 'th'
                cells.append({
                    'text': text,
                    'colspan': colspan,
                    'rowspan': rowspan,
                    'is_header': is_header
                })
            rows.append(cells)
        
        if not rows:
            return ''
        
        # 테이블 XML 생성
        col_count = max(sum(c['colspan'] for c in row) for row in rows) if rows else 1
        row_count = len(rows)
        
        table_xml = f'''<hp:tbl cols="{col_count}" rows="{row_count}">'''
        
        for row_idx, row in enumerate(rows):
            table_xml += '\n    <hp:tr>'
            for cell in row:
                bold = 'true' if cell['is_header'] else 'false'
                table_xml += f'''
        <hp:tc colspan="{cell['colspan']}" rowspan="{cell['rowspan']}">
            <hp:p>
                <hp:run>
                    <hp:charPr><hp:bold val="{bold}"/></hp:charPr>
                    <hp:t>{self._escape_xml(cell['text'])}</hp:t>
                </hp:run>
            </hp:p>
        </hp:tc>'''
            table_xml += '\n    </hp:tr>'
        
        table_xml += '\n</hp:tbl>'
        return table_xml

    def _process_list(self, list_element) -> list:
        """리스트 처리"""
        paragraphs = []
        is_ordered = list_element.name == 'ol'
        
        for idx, li in enumerate(list_element.find_all('li', recursive=False)):
            prefix = f"{idx + 1}. " if is_ordered else "• "
            text = li.get_text(strip=True)
            paragraphs.append(f'''<hp:p>
    <hp:run>
        <hp:t>{self._escape_xml(prefix + text)}</hp:t>
    </hp:run>
</hp:p>''')
        
        return paragraphs

    def _process_image(self, img_element) -> tuple:
        """이미지 처리"""
        src = img_element.get('src', '')
        
        if src.startswith('data:'):
            # Data URL 처리
            match = re.match(r'data:([^;]+);base64,(.+)', src)
            if match:
                mimetype = match.group(1)
                data = base64.b64decode(match.group(2))
                
                self.image_counter += 1
                img_id = f"image{self.image_counter}.{mimetype.split('/')[-1]}"
                
                img_data = {
                    'id': img_id,
                    'data': data,
                    'mimetype': mimetype
                }
                
                img_xml = self._create_image_xml(img_id)
                return img_xml, img_data
        
        return '', None

    def _process_svg(self, svg_element) -> tuple:
        """SVG를 이미지로 변환"""
        # SVG를 문자열로 변환
        svg_str = str(svg_element)
        
        self.image_counter += 1
        img_id = f"image{self.image_counter}.svg"
        
        img_data = {
            'id': img_id,
            'data': svg_str.encode('utf-8'),
            'mimetype': 'image/svg+xml'
        }
        
        img_xml = self._create_image_xml(img_id)
        return img_xml, img_data

    def _process_canvas(self, canvas_element) -> tuple:
        """Canvas 처리 (data-url 속성 사용)"""
        data_url = canvas_element.get('data-image-url', '')
        if data_url:
            # data URL 처리
            match = re.match(r'data:([^;]+);base64,(.+)', data_url)
            if match:
                mimetype = match.group(1)
                data = base64.b64decode(match.group(2))
                
                self.image_counter += 1
                img_id = f"image{self.image_counter}.png"
                
                img_data = {
                    'id': img_id,
                    'data': data,
                    'mimetype': mimetype
                }
                
                img_xml = self._create_image_xml(img_id)
                return img_xml, img_data
        
        return '', None

    def _create_image_xml(self, img_id: str) -> str:
        """이미지 삽입 XML 생성"""
        return f'''<hp:p>
    <hp:run>
        <hp:pic>
            <hp:binItemRef binItem="{img_id}"/>
            <hp:imgRect>
                <hp:pt x="0" y="0"/>
                <hp:sz cx="48000" cy="36000"/>
            </hp:imgRect>
        </hp:pic>
    </hp:run>
</hp:p>'''

    def _escape_xml(self, text: str) -> str:
        """XML 특수문자 이스케이프"""
        if not text:
            return ''
        return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;')
            .replace("'", '&apos;'))


def create_hwpx_from_html(pages: list, year: int, quarter: int, output_path: str) -> dict:
    """
    HTML 페이지들을 HWPX 파일로 변환하는 헬퍼 함수
    
    Args:
        pages: [{"html": str, "title": str, "category": str}, ...]
        year: 연도
        quarter: 분기
        output_path: 출력 파일 경로
        
    Returns:
        {"success": bool, "filepath": str, ...}
    """
    generator = HWPXGenerator(year, quarter)
    return generator.generate(pages, output_path)

