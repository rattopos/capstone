# -*- coding: utf-8 -*-
"""
HWPX 파일 분석 도구
원본 HWPX 파일의 구조를 분석하여 정확한 재현 방법을 파악합니다.
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List
import json


class HWPXAnalyzer:
    """HWPX 파일 구조 분석 클래스"""
    
    def __init__(self, hwpx_path: str):
        """
        Args:
            hwpx_path: 분석할 HWPX 파일 경로
        """
        self.hwpx_path = Path(hwpx_path)
        self.structure = {}
        
    def analyze(self) -> Dict:
        """
        HWPX 파일의 전체 구조를 분석
        
        Returns:
            분석 결과 딕셔너리
        """
        if not self.hwpx_path.exists():
            raise FileNotFoundError(f"HWPX 파일을 찾을 수 없습니다: {self.hwpx_path}")
        
        with zipfile.ZipFile(self.hwpx_path, 'r') as zip_file:
            # 1. 파일 목록 확인
            self.structure['files'] = zip_file.namelist()
            
            # 2. Content Types 분석
            if '[Content_Types].xml' in zip_file.namelist():
                content_types_xml = zip_file.read('[Content_Types].xml')
                self.structure['content_types'] = self._parse_content_types(content_types_xml)
            
            # 3. Relationships 분석
            if '_rels/.rels' in zip_file.namelist():
                rels_xml = zip_file.read('_rels/.rels')
                self.structure['relationships'] = self._parse_relationships(rels_xml)
            
            # 4. 메인 섹션 분석
            section_files = [f for f in zip_file.namelist() if 'Contents/section' in f and f.endswith('.xml')]
            self.structure['sections'] = {}
            for section_file in section_files:
                section_xml = zip_file.read(section_file)
                self.structure['sections'][section_file] = self._parse_section(section_xml)
            
            # 5. Manifest 분석
            if 'manifest.xml' in zip_file.namelist():
                manifest_xml = zip_file.read('manifest.xml')
                self.structure['manifest'] = self._parse_manifest(manifest_xml)
            
            # 6. 스타일 분석
            style_files = [f for f in zip_file.namelist() if 'Styles' in f or 'styles' in f]
            if style_files:
                self.structure['styles'] = {}
                for style_file in style_files:
                    style_xml = zip_file.read(style_file)
                    self.structure['styles'][style_file] = self._parse_styles(style_xml)
        
        return self.structure
    
    def _parse_content_types(self, xml_data: bytes) -> Dict:
        """Content Types XML 파싱"""
        root = ET.fromstring(xml_data)
        result = {
            'defaults': [],
            'overrides': []
        }
        
        for elem in root:
            if elem.tag.endswith('Default'):
                result['defaults'].append({
                    'extension': elem.get('Extension'),
                    'content_type': elem.get('ContentType')
                })
            elif elem.tag.endswith('Override'):
                result['overrides'].append({
                    'part_name': elem.get('PartName'),
                    'content_type': elem.get('ContentType')
                })
        
        return result
    
    def _parse_relationships(self, xml_data: bytes) -> Dict:
        """Relationships XML 파싱"""
        root = ET.fromstring(xml_data)
        relationships = []
        
        for rel in root:
            relationships.append({
                'id': rel.get('Id'),
                'type': rel.get('Type'),
                'target': rel.get('Target')
            })
        
        return relationships
    
    def _parse_section(self, xml_data: bytes) -> Dict:
        """섹션 XML 파싱"""
        root = ET.fromstring(xml_data)
        
        # 네임스페이스 처리
        ns = {'hwpp': 'http://www.hancom.co.kr/hwpml/2011/hwpp'}
        
        result = {
            'document_info': {},
            'page_def': {},
            'paragraphs': [],
            'tables': [],
            'images': []
        }
        
        # DocumentInfo
        doc_info = root.find('.//hwpp:DocumentInfo', ns)
        if doc_info is not None:
            version = doc_info.find('hwpp:Version', ns)
            if version is not None:
                result['document_info']['version'] = version.text
        
        # PageDef
        page_def = root.find('.//hwpp:PageDef', ns)
        if page_def is not None:
            result['page_def'] = {
                'paper_type': page_def.get('paperType'),
                'width': page_def.get('width'),
                'height': page_def.get('height'),
                'margins': {
                    'left': page_def.get('leftMargin'),
                    'right': page_def.get('rightMargin'),
                    'top': page_def.get('topMargin'),
                    'bottom': page_def.get('bottomMargin')
                }
            }
        
        # Paragraphs
        for para in root.findall('.//hwpp:Paragraph', ns):
            para_info = self._extract_paragraph_info(para, ns)
            if para_info:
                result['paragraphs'].append(para_info)
        
        # Tables
        for table in root.findall('.//hwpp:Table', ns):
            table_info = self._extract_table_info(table, ns)
            if table_info:
                result['tables'].append(table_info)
        
        # Images
        for pic in root.findall('.//hwpp:Picture', ns):
            pic_info = {
                'bin_id': pic.get('binID'),
                'width': pic.get('width'),
                'height': pic.get('height')
            }
            result['images'].append(pic_info)
        
        return result
    
    def _extract_paragraph_info(self, para_elem, ns: Dict) -> Dict:
        """문단 정보 추출"""
        info = {
            'runs': []
        }
        
        # ParaShape
        para_shape = para_elem.find('hwpp:ParaShape', ns)
        if para_shape is not None:
            info['para_shape'] = dict(para_shape.attrib)
        
        # Runs
        for run in para_elem.findall('hwpp:Run', ns):
            run_info = {}
            
            # CharShape
            char_shape = run.find('hwpp:CharShape', ns)
            if char_shape is not None:
                run_info['char_shape'] = dict(char_shape.attrib)
            
            # Text
            text_elem = run.find('hwpp:Text', ns)
            if text_elem is not None:
                run_info['text'] = text_elem.text or ''
            
            info['runs'].append(run_info)
        
        return info if info['runs'] else None
    
    def _extract_table_info(self, table_elem, ns: Dict) -> Dict:
        """표 정보 추출"""
        info = {}
        
        # TableProp
        table_prop = table_elem.find('hwpp:TableProp', ns)
        if table_prop is not None:
            info['table_prop'] = dict(table_prop.attrib)
        
        # Rows
        rows = []
        for row in table_elem.findall('hwpp:Row', ns):
            row_info = {
                'cells': []
            }
            
            for cell in row.findall('hwpp:Cell', ns):
                cell_info = dict(cell.attrib)
                
                # 셀 내부 문단
                para = cell.find('hwpp:Paragraph', ns)
                if para is not None:
                    cell_info['paragraph'] = self._extract_paragraph_info(para, ns)
                
                row_info['cells'].append(cell_info)
            
            rows.append(row_info)
        
        info['rows'] = rows
        return info
    
    def _parse_manifest(self, xml_data: bytes) -> Dict:
        """Manifest XML 파싱"""
        root = ET.fromstring(xml_data)
        files = []
        
        ns = {'hwpp': 'http://www.hancom.co.kr/hwpml/2011/hwpp'}
        for file_entry in root.findall('hwpp:FileEntry', ns):
            files.append({
                'full_path': file_entry.get('full-path'),
                'media_type': file_entry.get('mediatype')
            })
        
        return files
    
    def _parse_styles(self, xml_data: bytes) -> Dict:
        """스타일 XML 파싱"""
        # 스타일 정보 추출
        root = ET.fromstring(xml_data)
        return {
            'raw_xml': ET.tostring(root, encoding='utf-8').decode('utf-8')
        }
    
    def export_analysis(self, output_path: str):
        """분석 결과를 JSON 파일로 저장"""
        output = Path(output_path)
        with open(output, 'w', encoding='utf-8') as f:
            json.dump(self.structure, f, ensure_ascii=False, indent=2)
        print(f"분석 결과가 저장되었습니다: {output_path}")
    
    def print_summary(self):
        """분석 결과 요약 출력"""
        print("=" * 60)
        print("HWPX 파일 구조 분석 결과")
        print("=" * 60)
        
        print(f"\n[파일 목록] ({len(self.structure.get('files', []))}개)")
        for file in self.structure.get('files', [])[:10]:
            print(f"  - {file}")
        if len(self.structure.get('files', [])) > 10:
            print(f"  ... 외 {len(self.structure.get('files', [])) - 10}개")
        
        if 'content_types' in self.structure:
            print(f"\n[Content Types]")
            for override in self.structure['content_types'].get('overrides', []):
                print(f"  {override['part_name']}: {override['content_type']}")
        
        if 'sections' in self.structure:
            print(f"\n[섹션 정보]")
            for section_file, section_data in self.structure['sections'].items():
                print(f"  {section_file}:")
                print(f"    - 문단: {len(section_data.get('paragraphs', []))}개")
                print(f"    - 표: {len(section_data.get('tables', []))}개")
                print(f"    - 이미지: {len(section_data.get('images', []))}개")
                
                if 'page_def' in section_data:
                    pd = section_data['page_def']
                    print(f"    - 용지: {pd.get('paper_type')} "
                          f"({pd.get('width')} x {pd.get('height')})")


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("사용법: python hwpx_analyzer.py <hwpx_file> [output_json]")
        sys.exit(1)
    
    hwpx_file = sys.argv[1]
    analyzer = HWPXAnalyzer(hwpx_file)
    
    print(f"HWPX 파일 분석 중: {hwpx_file}")
    analyzer.analyze()
    analyzer.print_summary()
    
    if len(sys.argv) >= 3:
        analyzer.export_analysis(sys.argv[2])

