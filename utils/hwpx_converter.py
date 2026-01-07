# -*- coding: utf-8 -*-
"""
HWPX (한글 Open XML) 변환 모듈
HTML 형식의 보도자료를 한글 2020 이상에서 지원하는 HWPX XML 형식으로 변환합니다.
"""

import xml.etree.ElementTree as ET
from xml.dom import minidom
from bs4 import BeautifulSoup
import base64
import io
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from typing import List, Dict, Optional, Tuple
from PIL import Image


class HWPXConverter:
    """HTML을 HWPX 형식으로 변환하는 클래스"""
    
    # HWPX 네임스페이스
    NS_HWPVML = "http://www.hancom.co.kr/hwpml/2011/vml"
    NS_HWPP = "http://www.hancom.co.kr/hwpml/2011/hwpp"
    NS_OOXML_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
    NS_OOXML_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
    
    def __init__(self):
        self.image_counter = 0
        self.images = {}  # 이미지 ID -> 바이너리 데이터
        
    def convert_html_to_hwpx(self, pages: List[Dict], year: int, quarter: int) -> bytes:
        """
        여러 HTML 페이지를 하나의 HWPX 파일로 변환
        
        Args:
            pages: [{"html": "...", "title": "..."}, ...] 형식의 페이지 리스트
            year: 연도
            quarter: 분기
            
        Returns:
            HWPX 파일의 바이너리 데이터 (ZIP 형식)
        """
        self.image_counter = 0
        self.images = {}
        
        # 메모리 내 ZIP 파일 생성
        zip_buffer = io.BytesIO()
        
        with ZipFile(zip_buffer, 'w', ZIP_DEFLATED) as zip_file:
            # 1. Content Types 파일 생성
            content_types = self._create_content_types()
            zip_file.writestr('[Content_Types].xml', content_types)
            
            # 2. Relationships 파일 생성
            rels = self._create_relationships()
            zip_file.writestr('_rels/.rels', rels)
            
            # 3. 문서 내용 생성
            section_xml, image_files = self._create_section_content(pages, year, quarter)
            zip_file.writestr('Contents/section0.xml', section_xml)
            
            # 4. 이미지 파일 추가
            for image_id, image_data in image_files.items():
                zip_file.writestr(f'BinData/{image_id}', image_data)
            
            # 5. Manifest 파일 생성
            manifest = self._create_manifest(list(image_files.keys()))
            zip_file.writestr('manifest.xml', manifest)
        
        zip_buffer.seek(0)
        return zip_buffer.read()
    
    def _create_content_types(self) -> str:
        """Content Types XML 생성"""
        root = ET.Element('Types', xmlns=self.NS_OOXML_CT)
        
        # 기본 타입들
        ET.SubElement(root, 'Default', Extension='rels', ContentType='application/vnd.openxmlformats-package.relationships+xml')
        ET.SubElement(root, 'Default', Extension='xml', ContentType='application/xml')
        ET.SubElement(root, 'Default', Extension='png', ContentType='image/png')
        ET.SubElement(root, 'Default', Extension='jpg', ContentType='image/jpeg')
        ET.SubElement(root, 'Default', Extension='jpeg', ContentType='image/jpeg')
        
        # HWPX 문서 타입
        ET.SubElement(root, 'Override', PartName='/Contents/section0.xml', 
                     ContentType='application/vnd.hancom.hwpml+xml')
        ET.SubElement(root, 'Override', PartName='/manifest.xml',
                     ContentType='application/vnd.hancom.hwpml.manifest+xml')
        
        return self._pretty_xml(root)
    
    def _create_relationships(self) -> str:
        """Relationships XML 생성"""
        root = ET.Element('Relationships', xmlns=self.NS_OOXML_REL)
        
        ET.SubElement(root, 'Relationship', Id='rId1', Type='http://www.hancom.co.kr/hwpml/2011/hwpp', 
                     Target='Contents/section0.xml')
        
        return self._pretty_xml(root)
    
    def _create_section_content(self, pages: List[Dict], year: int, quarter: int) -> Tuple[str, Dict]:
        """
        섹션 콘텐츠 XML 생성
        
        Returns:
            (XML 문자열, 이미지 파일 딕셔너리)
        """
        # HWPML 루트 요소
        root = ET.Element('HWPML', xmlns=self.NS_HWPP)
        
        # 문서 정보
        doc_info = ET.SubElement(root, 'DocumentInfo')
        ET.SubElement(doc_info, 'Version').text = '5.0'
        
        # 본문
        body = ET.SubElement(root, 'Body')
        section = ET.SubElement(body, 'Section')
        
        # 페이지 설정
        page_def = ET.SubElement(section, 'PageDef')
        page_def.set('paperType', 'A4')
        page_def.set('width', '210000')  # A4 너비 (HWP 단위: 1/7200 inch)
        page_def.set('height', '297000')  # A4 높이
        page_def.set('leftMargin', '15000')  # 왼쪽 여백
        page_def.set('rightMargin', '15000')  # 오른쪽 여백
        page_def.set('topMargin', '12000')  # 위 여백
        page_def.set('bottomMargin', '15000')  # 아래 여백
        
        # 각 페이지 처리
        for page_idx, page in enumerate(pages):
            html_content = page.get('html', '')
            title = page.get('title', f'페이지 {page_idx + 1}')
            
            if page_idx > 0:
                # 페이지 나누기 (첫 페이지 제외)
                para = ET.SubElement(section, 'Paragraph')
                run = ET.SubElement(para, 'Run')
                br = ET.SubElement(run, 'Break')
                br.set('type', 'page')
            
            # 제목 추가
            if title:
                para = self._create_paragraph(title, is_heading=True)
                section.append(para)
            
            # HTML 콘텐츠 변환
            if html_content:
                soup = BeautifulSoup(html_content, 'html.parser')
                
                # body 또는 직접 콘텐츠 찾기
                body = soup.find('body')
                if body:
                    content_root = body
                else:
                    # body가 없으면 직접 주요 콘텐츠 찾기
                    content_root = soup
                
                # 주요 콘텐츠 요소들 처리
                self._process_html_content(content_root, section)
        
        xml_str = self._pretty_xml(root)
        return xml_str, self.images
    
    def _process_html_content(self, root_element, parent_section):
        """HTML 콘텐츠를 재귀적으로 처리하여 HWPX 요소로 변환"""
        if not root_element:
            return
        
        # 직접 자식 요소들을 순회
        for element in root_element.children:
            if not hasattr(element, 'name'):
                continue
            
            # 스크립트, 스타일 태그는 건너뛰기
            if element.name in ['script', 'style', 'meta', 'link']:
                continue
            
            # 표 처리
            if element.name == 'table':
                table_elem = self._convert_table(element)
                if table_elem is not None:
                    parent_section.append(table_elem)
            
            # 이미지 처리
            elif element.name == 'img':
                img_elem = self._convert_image(element)
                if img_elem is not None:
                    parent_section.append(img_elem)
            
            # 제목 요소
            elif element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                text = element.get_text(strip=True)
                if text:
                    para = self._create_paragraph(text, is_heading=True)
                    parent_section.append(para)
            
            # 문단 요소
            elif element.name == 'p':
                text = element.get_text(strip=True)
                if text:
                    para = self._create_paragraph_with_style(text, element)
                    parent_section.append(para)
            
            # div 요소 (블록 콘텐츠)
            elif element.name == 'div':
                # div 내부 콘텐츠 재귀 처리
                self._process_html_content(element, parent_section)
            
            # 리스트 처리
            elif element.name in ['ul', 'ol']:
                for li in element.find_all('li', recursive=False):
                    text = li.get_text(strip=True)
                    if text:
                        para = self._create_paragraph(f"• {text}")
                        parent_section.append(para)
            
            # 기타 블록 요소
            else:
                text = element.get_text(strip=True)
                if text and element.name not in ['span', 'a', 'strong', 'em', 'b', 'i']:
                    para = self._create_paragraph(text)
                    parent_section.append(para)
                elif element.name in ['span', 'a', 'strong', 'em', 'b', 'i']:
                    # 인라인 요소는 부모 문단에 포함되도록 처리
                    # 여기서는 간단히 텍스트로만 처리
                    text = element.get_text(strip=True)
                    if text:
                        para = self._create_paragraph(text)
                        parent_section.append(para)
    
    def _create_paragraph_with_style(self, text: str, html_element) -> ET.Element:
        """HTML 요소의 스타일을 고려한 문단 생성"""
        para = self._create_paragraph(text)
        
        # 인라인 스타일에서 정렬 추출
        style = html_element.get('style', '')
        if 'text-align: center' in style or 'text-align:center' in style:
            para.find('ParaShape').set('alignment', 'center')
        elif 'text-align: right' in style or 'text-align:right' in style:
            para.find('ParaShape').set('alignment', 'right')
        
        return para
    
    def _create_paragraph(self, text: str, is_heading: bool = False) -> ET.Element:
        """문단 요소 생성"""
        para = ET.Element('Paragraph')
        
        run = ET.SubElement(para, 'Run')
        
        # 텍스트 스타일 설정
        char_prop = ET.SubElement(run, 'CharShape')
        if is_heading:
            char_prop.set('size', '1800')  # 18pt
            char_prop.set('bold', 'true')
        else:
            char_prop.set('size', '1000')  # 10pt
        
        text_elem = ET.SubElement(run, 'Text')
        text_elem.text = text
        
        # 문단 스타일
        para_prop = ET.SubElement(para, 'ParaShape')
        if is_heading:
            para_prop.set('alignment', 'center')  # 제목은 가운데 정렬
        else:
            para_prop.set('alignment', 'left')
        
        return para
    
    def _convert_table(self, table_element) -> Optional[ET.Element]:
        """HTML 테이블을 HWPX 표로 변환"""
        rows = table_element.find_all('tr')
        if not rows:
            return None
        
        table = ET.Element('Table')
        
        # 표 속성
        table_prop = ET.SubElement(table, 'TableProp')
        table_prop.set('rowCount', str(len(rows)))
        
        # 열 너비 설정 (첫 행 기준)
        first_row = rows[0]
        cols = first_row.find_all(['th', 'td'])
        col_count = len(cols)
        table_prop.set('colCount', str(col_count))
        
        # 각 행 처리
        for row_idx, tr in enumerate(rows):
            row = ET.SubElement(table, 'Row')
            
            cells = tr.find_all(['th', 'td'])
            for col_idx, cell in enumerate(cells):
                cell_elem = ET.SubElement(row, 'Cell')
                
                # 셀 병합 처리
                colspan = int(cell.get('colspan', 1))
                rowspan = int(cell.get('rowspan', 1))
                
                if colspan > 1:
                    cell_elem.set('colSpan', str(colspan))
                if rowspan > 1:
                    cell_elem.set('rowSpan', str(rowspan))
                
                # 셀 내용
                cell_text = cell.get_text(strip=True)
                if cell_text:
                    para = ET.SubElement(cell_elem, 'Paragraph')
                    run = ET.SubElement(para, 'Run')
                    text = ET.SubElement(run, 'Text')
                    text.text = cell_text
                    
                    # 헤더 셀 스타일
                    if cell.name == 'th':
                        char_prop = ET.SubElement(run, 'CharShape')
                        char_prop.set('bold', 'true')
        
        return table
    
    def _convert_image(self, img_element) -> Optional[ET.Element]:
        """이미지를 HWPX 이미지 객체로 변환"""
        src = img_element.get('src', '')
        if not src:
            return None
        
        # base64 이미지 처리
        if src.startswith('data:image'):
            match = re.match(r'data:image/([^;]+);base64,(.+)', src)
            if match:
                img_format = match.group(1).lower()
                img_data = base64.b64decode(match.group(2))
                
                # 이미지 ID 생성
                self.image_counter += 1
                image_id = f'image{self.image_counter}.{img_format if img_format in ["png", "jpg", "jpeg"] else "png"}'
                
                # 이미지 데이터 저장
                self.images[image_id] = img_data
                
                # 이미지 요소 생성
                para = ET.Element('Paragraph')
                
                # 이미지 객체
                pic = ET.SubElement(para, 'Picture')
                pic.set('binID', image_id)
                
                # 이미지 크기 정보 (기본값)
                pic.set('width', '100000')  # 약 3.5cm
                pic.set('height', '75000')  # 비율 유지
                
                # 실제 이미지 크기 가져오기
                try:
                    img = Image.open(io.BytesIO(img_data))
                    width, height = img.size
                    # HWP 단위로 변환 (1 pixel ≈ 1270 HWP units)
                    hwp_width = int(width * 1270)
                    hwp_height = int(height * 1270)
                    pic.set('width', str(hwp_width))
                    pic.set('height', str(hwp_height))
                except:
                    pass  # 기본값 사용
                
                return para
        
        return None
    
    def _create_manifest(self, image_ids: List[str]) -> str:
        """Manifest XML 생성"""
        root = ET.Element('Manifest', xmlns=self.NS_HWPP)
        
        # 문서 항목
        doc_item = ET.SubElement(root, 'FileEntry')
        doc_item.set('full-path', 'Contents/section0.xml')
        doc_item.set('mediatype', 'application/vnd.hancom.hwpml+xml')
        
        # 이미지 항목들
        for image_id in image_ids:
            img_item = ET.SubElement(root, 'FileEntry')
            img_item.set('full-path', f'BinData/{image_id}')
            img_item.set('mediatype', f'image/{image_id.split(".")[-1]}')
        
        return self._pretty_xml(root)
    
    def _pretty_xml(self, elem: ET.Element) -> str:
        """XML을 보기 좋게 포맷팅"""
        rough_string = ET.tostring(elem, encoding='utf-8')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ", encoding='utf-8').decode('utf-8')

