"""
PDF를 읽어서 Word 템플릿으로 변환하는 모듈
PDF의 모든 페이지를 읽고 디자인을 그대로 재현한 Word 파일 생성
"""

import os
from pathlib import Path
from typing import List, Dict, Optional, Tuple
from PIL import Image
import io

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import easyocr
    EASYOCR_AVAILABLE = True
except ImportError:
    EASYOCR_AVAILABLE = False

try:
    import pytesseract
    PYTESSERACT_AVAILABLE = True
except ImportError:
    PYTESSERACT_AVAILABLE = False

try:
    import cv2
    import numpy as np
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False


class PDFToWordConverter:
    """PDF를 읽어서 Word 템플릿으로 변환하는 클래스"""
    
    def __init__(self, use_easyocr: bool = True, dpi: int = 300):
        """
        PDF 변환기 초기화
        
        Args:
            use_easyocr: True면 easyocr 사용, False면 pytesseract 사용
            dpi: PDF를 이미지로 변환할 때의 DPI (기본값: 300)
        """
        self.use_easyocr = use_easyocr and EASYOCR_AVAILABLE
        self.dpi = dpi
        self.reader = None
        
        if self.use_easyocr:
            try:
                self.reader = easyocr.Reader(['ko', 'en'], gpu=False)
            except Exception as e:
                print(f"EasyOCR 초기화 실패, pytesseract 사용: {e}")
                self.use_easyocr = False
                self.reader = None
    
    def extract_text_from_pdf_direct(self, pdf_path: str) -> List[List[Dict]]:
        """
        PDF에서 직접 텍스트와 레이아웃 정보를 추출합니다 (이미지 변환 없이).
        표 정보도 함께 추출합니다.
        
        Args:
            pdf_path: PDF 파일 경로
            
        Returns:
            페이지별 텍스트 정보 리스트. 각 페이지는 텍스트 정보 딕셔너리 리스트를 포함:
            - 'text': 추출된 텍스트
            - 'bbox': 바운딩 박스 좌표 (x1, y1, x2, y2)
            - 'font_size': 폰트 크기
            - 'is_bold': 굵은 글씨 여부
            - 'page_num': 페이지 번호
        """
        if not PYMUPDF_AVAILABLE:
            return []
        
        pages_text_data = []
        
        try:
            doc = fitz.open(pdf_path)
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                page_text_data = []
                
                # PyMuPDF의 표 감지 시도 (최신 버전)
                try:
                    # find_tables() 메서드가 있는지 확인
                    if hasattr(page, 'find_tables'):
                        tables = page.find_tables()
                        if tables:
                            print(f"[DEBUG PDFToWord] 페이지 {page_num + 1}에서 {len(tables.tables)}개 표 발견 (PyMuPDF find_tables)")
                            # 표 데이터는 별도로 처리하므로 여기서는 텍스트만 추출
                except Exception as e:
                    # find_tables가 없거나 실패한 경우 무시
                    pass
                
                # PDF에서 텍스트 딕셔너리 형식으로 추출
                text_dict = page.get_text("dict")
                
                # 텍스트 블록 처리
                for block in text_dict.get("blocks", []):
                    if "lines" not in block:
                        continue
                    
                    for line in block["lines"]:
                        for span in line.get("spans", []):
                            text = span.get("text", "").strip()
                            if not text:
                                continue
                            
                            # 바운딩 박스
                            bbox = span.get("bbox", [0, 0, 0, 0])
                            
                            # 폰트 정보
                            font_size = span.get("size", 12)
                            font_flags = span.get("flags", 0)
                            is_bold = (font_flags & 16) != 0  # bit 4는 bold
                            
                            page_text_data.append({
                                'text': text,
                                'bbox': tuple(bbox),
                                'font_size': font_size,
                                'is_bold': is_bold,
                                'confidence': 1.0,  # 직접 추출이므로 신뢰도 100%
                                'page_num': page_num + 1
                            })
                
                pages_text_data.append(page_text_data)
            
            doc.close()
            
        except Exception as e:
            print(f"PDF 직접 텍스트 추출 실패: {e}")
            return []
        
        return pages_text_data
    
    def extract_images_from_pdf(self, pdf_path: str) -> List[List[Dict]]:
        """
        PDF에서 이미지를 추출합니다 (그래프, 인포그래픽 등).
        
        Args:
            pdf_path: PDF 파일 경로
            
        Returns:
            페이지별 이미지 정보 리스트. 각 페이지는 이미지 정보 딕셔너리 리스트를 포함:
            - 'image_data': 이미지 바이너리 데이터
            - 'bbox': 바운딩 박스 좌표 (x1, y1, x2, y2)
            - 'width': 이미지 너비
            - 'height': 이미지 높이
            - 'page_num': 페이지 번호
        """
        if not PYMUPDF_AVAILABLE:
            return []
        
        pages_images = []
        
        try:
            doc = fitz.open(pdf_path)
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                page_images = []
                
                # 페이지에서 이미지 리스트 가져오기
                image_list = page.get_images()
                
                for img_idx, img in enumerate(image_list):
                    try:
                        # 이미지 정보
                        xref = img[0]  # 이미지 참조 번호
                        base_image = doc.extract_image(xref)
                        image_bytes = base_image["image"]
                        image_ext = base_image["ext"]
                        
                        # 이미지의 위치 정보 가져오기
                        # 이미지가 포함된 블록 찾기
                        image_rects = page.get_image_rects(xref)
                        
                        for rect in image_rects:
                            # 바운딩 박스 (PDF 좌표계)
                            bbox = (rect.x0, rect.y0, rect.x1, rect.y1)
                            
                            # 이미지 크기
                            width = rect.width
                            height = rect.height
                            
                            # 최소 크기 필터링 (너무 작은 아이콘 제외)
                            if width < 50 or height < 50:
                                continue
                            
                            page_images.append({
                                'image_data': image_bytes,
                                'image_ext': image_ext,
                                'bbox': bbox,
                                'width': width,
                                'height': height,
                                'page_num': page_num + 1
                            })
                    except Exception as e:
                        print(f"[DEBUG PDFToWord] 이미지 {img_idx} 추출 실패: {e}")
                        continue
                
                pages_images.append(page_images)
                if page_images:
                    print(f"[DEBUG PDFToWord] 페이지 {page_num + 1}에서 {len(page_images)}개 이미지 발견")
            
            doc.close()
            
        except Exception as e:
            print(f"PDF 이미지 추출 실패: {e}")
            return []
        
        return pages_images
    
    def pdf_to_images(self, pdf_path: str) -> List[Image.Image]:
        """
        PDF 파일을 이미지 리스트로 변환합니다 (스캔된 PDF용).
        
        Args:
            pdf_path: PDF 파일 경로
            
        Returns:
            PIL Image 객체 리스트
        """
        images = []
        
        if PYMUPDF_AVAILABLE:
            try:
                doc = fitz.open(pdf_path)
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    zoom = self.dpi / 72.0
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat)
                    img_data = pix.tobytes("png")
                    img = Image.open(io.BytesIO(img_data))
                    images.append(img)
                doc.close()
                return images
            except Exception as e:
                print(f"PyMuPDF 변환 실패: {e}")
                raise ValueError(f"PDF를 이미지로 변환할 수 없습니다: {e}")
        
        raise ValueError("PDF를 이미지로 변환할 수 있는 라이브러리가 설치되어 있지 않습니다. PyMuPDF를 설치해주세요.")
    
    def extract_text_and_layout(self, image: Image.Image) -> List[Dict]:
        """
        이미지에서 텍스트와 레이아웃 정보를 추출합니다.
        
        Args:
            image: PIL Image 객체
            
        Returns:
            텍스트 정보 딕셔너리 리스트. 각 딕셔너리는 다음 키를 포함:
            - 'text': 추출된 텍스트
            - 'bbox': 바운딩 박스 좌표 (x1, y1, x2, y2)
            - 'confidence': 신뢰도 (0-1)
            - 'font_size': 추정된 폰트 크기
            - 'is_bold': 굵은 글씨 여부 (추정)
        """
        text_data = []
        
        if CV2_AVAILABLE:
            img_array = np.array(image)
            if len(img_array.shape) == 3:
                gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            else:
                gray = img_array
            
            denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
            _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            processed_image = Image.fromarray(binary)
        else:
            processed_image = image.convert('L')
        
        if self.use_easyocr and self.reader:
            try:
                if CV2_AVAILABLE:
                    img_array = np.array(processed_image)
                    results = self.reader.readtext(img_array)
                else:
                    results = self.reader.readtext(processed_image)
                
                for (bbox, text, confidence) in results:
                    x_coords = [point[0] for point in bbox]
                    y_coords = [point[1] for point in bbox]
                    x1, y1 = min(x_coords), min(y_coords)
                    x2, y2 = max(x_coords), max(y_coords)
                    
                    # 폰트 크기 추정 (높이 기준)
                    font_size = max(8, int((y2 - y1) * 0.75))
                    
                    text_data.append({
                        'text': text.strip(),
                        'bbox': (x1, y1, x2, y2),
                        'confidence': confidence,
                        'font_size': font_size,
                        'is_bold': False  # OCR로는 정확히 판단 어려움
                    })
            except Exception as e:
                print(f"EasyOCR 처리 중 오류: {e}")
                if PYTESSERACT_AVAILABLE:
                    return self._extract_with_pytesseract(processed_image)
        
        elif PYTESSERACT_AVAILABLE:
            return self._extract_with_pytesseract(processed_image)
        else:
            raise ValueError("OCR 라이브러리가 설치되어 있지 않습니다. easyocr 또는 pytesseract를 설치해주세요.")
        
        return text_data
    
    def _extract_with_pytesseract(self, image: Image.Image) -> List[Dict]:
        """pytesseract를 사용하여 텍스트 추출"""
        text_data = []
        try:
            data = pytesseract.image_to_data(image, lang='kor+eng', output_type=pytesseract.Output.DICT)
            
            for i in range(len(data['text'])):
                text = data['text'][i].strip()
                if text and int(data['conf'][i]) > 0:
                    x1 = data['left'][i]
                    y1 = data['top'][i]
                    x2 = x1 + data['width'][i]
                    y2 = y1 + data['height'][i]
                    confidence = float(data['conf'][i]) / 100.0
                    font_size = max(8, int(data['height'][i] * 0.75))
                    
                    text_data.append({
                        'text': text,
                        'bbox': (x1, y1, x2, y2),
                        'confidence': confidence,
                        'font_size': font_size,
                        'is_bold': False
                    })
        except Exception as e:
            raise ValueError(f"OCR 처리 중 오류 발생: {e}")
        
        return text_data
    
    def organize_text_by_layout(self, text_data: List[Dict], page_width: int, page_height: int) -> List[Dict]:
        """
        텍스트 데이터를 레이아웃에 따라 정리합니다.
        표를 감지하여 별도로 처리합니다.
        
        Args:
            text_data: extract_text_and_layout()의 결과
            page_width: 페이지 너비
            page_height: 페이지 높이
            
        Returns:
            정리된 텍스트 구조 리스트 (단락, 제목, 표 등으로 구분)
        """
        if not text_data:
            return []
        
        # 표 감지 시도
        tables = self._detect_tables(text_data, page_width, page_height)
        
        # y 좌표 기준으로 정렬
        sorted_text = sorted(text_data, key=lambda x: (x['bbox'][1], x['bbox'][0]))
        
        organized = []
        current_paragraph = []
        current_y = None
        line_threshold = page_height * 0.02
        
        # 표 영역에 포함된 텍스트는 제외 (bbox 좌표로 판단)
        text_in_tables = set()
        for table in tables:
            table_bbox = table.get('bbox', (0, 0, page_width, page_height))
            for idx, item in enumerate(sorted_text):
                item_bbox = item['bbox']
                # 텍스트가 표 영역 안에 있는지 확인 (약간의 여유 공간 포함)
                margin = page_width * 0.01  # 1% 여유
                if (table_bbox[0] - margin <= item_bbox[0] <= table_bbox[2] + margin and
                    table_bbox[1] - margin <= item_bbox[1] <= table_bbox[3] + margin):
                    text_in_tables.add(idx)
        
        for idx, item in enumerate(sorted_text):
            # 표에 포함된 텍스트는 건너뛰기
            if idx in text_in_tables:
                continue
            
            y = item['bbox'][1]
            
            if current_y is None or abs(y - current_y) < line_threshold:
                # 같은 줄
                current_paragraph.append(item)
                if current_y is None:
                    current_y = y
            else:
                # 새 줄
                if current_paragraph:
                    current_paragraph.sort(key=lambda x: x['bbox'][0])
                    organized.append({
                        'type': 'paragraph',
                        'items': current_paragraph,
                        'y': current_y,
                        'font_size': max([item['font_size'] for item in current_paragraph], default=12)
                    })
                
                current_paragraph = [item]
                current_y = y
        
        # 마지막 단락 처리
        if current_paragraph:
            current_paragraph.sort(key=lambda x: x['bbox'][0])
            organized.append({
                'type': 'paragraph',
                'items': current_paragraph,
                'y': current_y,
                'font_size': max([item['font_size'] for item in current_paragraph], default=12)
            })
        
        # 표를 organized 리스트에 추가 (y 좌표 기준으로 정렬)
        for table in tables:
            organized.append({
                'type': 'table',
                'table_data': table,
                'y': table.get('bbox', [0, 0, 0, 0])[1]
            })
        
        # y 좌표 기준으로 다시 정렬
        organized.sort(key=lambda x: x.get('y', 0))
        
        return organized
    
    def _detect_tables(self, text_data: List[Dict], page_width: int, page_height: int) -> List[Dict]:
        """
        텍스트 데이터에서 표를 감지합니다.
        
        Args:
            text_data: 텍스트 데이터 리스트
            page_width: 페이지 너비
            page_height: 페이지 높이
            
        Returns:
            표 정보 리스트. 각 표는 다음 키를 포함:
            - 'rows': 행 리스트
            - 'bbox': 표의 바운딩 박스
        """
        if not text_data:
            return []
        
        # y 좌표 기준으로 그룹화 (행 감지)
        y_groups = {}
        y_threshold = page_height * 0.015  # 1.5% 높이 차이를 같은 행으로 간주
        
        for item in text_data:
            y = item['bbox'][1]
            # 가장 가까운 y 그룹 찾기
            matched_group = None
            for group_y in y_groups.keys():
                if abs(y - group_y) < y_threshold:
                    matched_group = group_y
                    break
            
            if matched_group is not None:
                y_groups[matched_group].append(item)
            else:
                y_groups[y] = [item]
        
        # 각 행을 x 좌표 기준으로 정렬
        rows = []
        for y, items in sorted(y_groups.items()):
            items_sorted = sorted(items, key=lambda x: x['bbox'][0])
            rows.append(items_sorted)
        
        # 표 감지: 여러 행이 있고, 각 행에 여러 셀이 있는 경우
        tables = []
        if len(rows) >= 2:  # 최소 2행 이상
            # 연속된 행들을 그룹화하여 표로 인식
            current_table_rows = []
            min_x = min(item['bbox'][0] for row in rows for item in row)
            max_x = max(item['bbox'][2] for row in rows for item in row)
            min_y = rows[0][0]['bbox'][1]
            max_y = rows[-1][0]['bbox'][3]
            
            # 각 행의 셀 개수 확인
            cells_per_row = [len(row) for row in rows]
            avg_cells = sum(cells_per_row) / len(cells_per_row) if cells_per_row else 0
            
            # 평균적으로 2개 이상의 셀이 있는 행이 3개 이상이면 표로 간주
            rows_with_multiple_cells = sum(1 for count in cells_per_row if count >= 2)
            
            if rows_with_multiple_cells >= 3 and avg_cells >= 2:
                # 표 데이터 구성
                table_rows = []
                for row in rows:
                    table_row = []
                    for item in row:
                        table_row.append({
                            'text': item['text'],
                            'bbox': item['bbox'],
                            'font_size': item.get('font_size', 12),
                            'is_bold': item.get('is_bold', False)
                        })
                    table_rows.append(table_row)
                
                tables.append({
                    'rows': table_rows,
                    'bbox': (min_x, min_y, max_x, max_y),
                    'num_rows': len(table_rows),
                    'num_cols': max(cells_per_row) if cells_per_row else 0
                })
        
        return tables
    
    def convert_pdf_to_word(self, pdf_path: str, output_path: str) -> str:
        """
        PDF 파일을 Word 템플릿으로 변환합니다.
        텍스트 기반 PDF는 직접 추출하고, 스캔된 PDF는 OCR을 사용합니다.
        
        Args:
            pdf_path: PDF 파일 경로
            output_path: 출력 Word 파일 경로
            
        Returns:
            생성된 Word 파일 경로
        """
        if not DOCX_AVAILABLE:
            raise ValueError("python-docx가 설치되어 있지 않습니다. pip install python-docx를 실행해주세요.")
        
        if not PYMUPDF_AVAILABLE:
            raise ValueError("PyMuPDF가 설치되어 있지 않습니다. pip install PyMuPDF를 실행해주세요.")
        
        print(f"PDF 파일 로딩 중: {pdf_path}")
        
        # 1단계: PDF에서 직접 텍스트 추출 시도 (텍스트 기반 PDF)
        print("텍스트 기반 추출 시도 중...")
        pages_text_data = self.extract_text_from_pdf_direct(pdf_path)
        
        # 1-1단계: PDF에서 이미지 추출 (그래프, 인포그래픽)
        print("이미지 추출 중...")
        pages_images = self.extract_images_from_pdf(pdf_path)
        total_images = sum(len(imgs) for imgs in pages_images)
        print(f"[DEBUG PDFToWord] 총 {total_images}개 이미지 발견")
        
        # 텍스트가 충분히 추출되었는지 확인
        total_text_chars = sum(len(item['text']) for page_data in pages_text_data for item in page_data)
        total_pages_with_text = len([p for p in pages_text_data if len(p) > 0])
        use_ocr = total_text_chars < 100  # 100자 미만이면 스캔된 PDF로 간주
        
        print(f"[DEBUG PDFToWord] 추출된 텍스트 통계:")
        print(f"  - 총 문자 수: {total_text_chars}")
        print(f"  - 텍스트가 있는 페이지: {total_pages_with_text}/{len(pages_text_data)}")
        print(f"  - 페이지별 텍스트 수:")
        for i, page_data in enumerate(pages_text_data[:5]):  # 처음 5페이지만
            page_chars = sum(len(item['text']) for item in page_data)
            print(f"    페이지 {i+1}: {page_chars}자, {len(page_data)}개 텍스트 블록")
        
        if use_ocr:
            print(f"[DEBUG PDFToWord] 텍스트가 부족합니다 ({total_text_chars}자). 스캔된 PDF로 판단하여 OCR을 사용합니다.")
        else:
            print(f"[DEBUG PDFToWord] 텍스트 기반 PDF로 확인되었습니다 ({total_text_chars}자 추출). 직접 추출 방식을 사용합니다.")
        
        # PDF 열기 (페이지 크기 정보를 위해)
        pdf_doc = fitz.open(pdf_path)
        total_pages = len(pdf_doc)
        
        # 첫 페이지의 크기로 Word 문서 설정
        first_page = pdf_doc[0]
        first_page_rect = first_page.rect
        pdf_page_width_pt = first_page_rect.width
        pdf_page_height_pt = first_page_rect.height
        
        # 포인트를 인치로 변환 (72 DPI 기준)
        pdf_page_width_inch = pdf_page_width_pt / 72.0
        pdf_page_height_inch = pdf_page_height_pt / 72.0
        
        # Word 문서 생성
        doc = Document()
        
        # 페이지 설정 (PDF 페이지 크기에 맞춤)
        section = doc.sections[0]
        section.page_height = Inches(pdf_page_height_inch)
        section.page_width = Inches(pdf_page_width_inch)
        
        # 마진 설정 (PDF와 유사하게)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        
        print(f"[DEBUG PDFToWord] PDF 페이지 크기: {pdf_page_width_inch:.2f} x {pdf_page_height_inch:.2f} 인치")
        
        # 각 페이지 처리
        for page_num in range(total_pages):
            print(f"페이지 {page_num + 1}/{total_pages} 처리 중...")
            
            if page_num > 0:
                # 페이지 나누기
                doc.add_page_break()
            
            page = pdf_doc[page_num]
            page_rect = page.rect
            page_width = page_rect.width
            page_height = page_rect.height
            
            if use_ocr:
                # OCR 방식: 이미지 변환 후 OCR
                zoom = self.dpi / 72.0
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                image = Image.open(io.BytesIO(img_data))
                
                text_data = self.extract_text_and_layout(image)
                organized_text = self.organize_text_by_layout(text_data, image.size[0], image.size[1]) if text_data else []
                
                # OCR 방식에서도 이미지 추출 (그래프, 인포그래픽)
                # 페이지 전체를 이미지로 사용하지 않고, 개별 이미지 객체만 추출
                page_images_for_ocr = []
                try:
                    image_list = page.get_images()
                    for img_idx, img in enumerate(image_list):
                        try:
                            xref = img[0]
                            base_image = pdf_doc.extract_image(xref)
                            image_bytes = base_image["image"]
                            image_ext = base_image["ext"]
                            
                            image_rects = page.get_image_rects(xref)
                            for rect in image_rects:
                                bbox = (rect.x0, rect.y0, rect.x1, rect.y1)
                                width = rect.width
                                height = rect.height
                                
                                # 최소 크기 필터링
                                if width >= 50 and height >= 50:
                                    page_images_for_ocr.append({
                                        'image_data': image_bytes,
                                        'image_ext': image_ext,
                                        'bbox': bbox,
                                        'width': width,
                                        'height': height,
                                        'page_num': page_num + 1
                                    })
                        except Exception as e:
                            print(f"[DEBUG PDFToWord] OCR 모드에서 이미지 {img_idx} 추출 실패: {e}")
                            continue
                    
                    # OCR 모드의 이미지를 pages_images에 추가
                    if page_num < len(pages_images):
                        pages_images[page_num].extend(page_images_for_ocr)
                    else:
                        # 페이지가 없으면 새로 추가
                        while len(pages_images) <= page_num:
                            pages_images.append([])
                        pages_images[page_num] = page_images_for_ocr
                except Exception as e:
                    print(f"[DEBUG PDFToWord] OCR 모드에서 이미지 추출 실패: {e}")
            else:
                # 직접 추출 방식: 이미 추출된 텍스트 사용
                page_text_data = pages_text_data[page_num] if page_num < len(pages_text_data) else []
                
                if not page_text_data:
                    para = doc.add_paragraph()
                    para.add_run(f"(페이지 {page_num + 1} - 텍스트를 추출할 수 없습니다)")
                    continue
                
                # 직접 추출된 텍스트를 organize_text_by_layout 형식으로 변환
                organized_text = self.organize_text_by_layout(page_text_data, page_width, page_height)
            
            if not organized_text:
                # 텍스트가 없으면 빈 단락 추가
                print(f"[DEBUG PDFToWord] ⚠️ 경고: 페이지 {page_num + 1}에서 텍스트를 추출할 수 없습니다")
                para = doc.add_paragraph()
                para.add_run(f"(페이지 {page_num + 1} - 텍스트를 추출할 수 없습니다)")
                continue
            
            # 디버그: 첫 페이지의 텍스트 샘플 출력
            if page_num == 0:
                print(f"[DEBUG PDFToWord] 첫 페이지 텍스트 샘플 (처음 3개 요소):")
                for i, item in enumerate(organized_text[:3]):
                    item_type = item.get('type', 'unknown')
                    if item_type == 'paragraph' and 'items' in item:
                        text_sample = ' '.join([ti.get('text', '') for ti in item.get('items', [])])
                        print(f"  요소 {i+1} (단락): {repr(text_sample[:100])}")
                    elif item_type == 'table':
                        print(f"  요소 {i+1} (표): {item.get('table_data', {}).get('num_rows', 0)}행 x {item.get('table_data', {}).get('num_cols', 0)}열")
                    else:
                        print(f"  요소 {i+1} (타입: {item_type})")
            
            # 페이지의 이미지 가져오기
            page_images = pages_images[page_num] if page_num < len(pages_images) else []
            
            # 모든 요소 (텍스트, 표, 이미지)를 y 좌표 기준으로 정렬
            all_elements = []
            
            # 텍스트/표 요소 추가
            for item in organized_text:
                # 안전하게 타입과 y 좌표 가져오기
                item_type = item.get('type', 'paragraph')
                item_y = item.get('y', 0)
                
                all_elements.append({
                    'type': item_type,
                    'data': item,
                    'y': item_y
                })
            
            # 이미지 요소 추가
            for img_info in page_images:
                img_y = img_info['bbox'][1]  # 이미지의 상단 y 좌표
                all_elements.append({
                    'type': 'image',
                    'data': img_info,
                    'y': img_y
                })
            
            # y 좌표 기준으로 정렬
            all_elements.sort(key=lambda x: x.get('y', 0))
            
            # Word 문서에 추가
            prev_y = None
            for element in all_elements:
                try:
                    element_y = element.get('y', 0)
                    
                    # 큰 간격이 있으면 단락 구분
                    if prev_y is not None and (element_y - prev_y) > page_height * 0.05:
                        doc.add_paragraph()  # 빈 단락 추가
                    
                    # 이미지인 경우
                    if element.get('type') == 'image':
                    img_info = element.get('data', {})
                    try:
                        # 이미지를 임시 파일로 저장
                        import tempfile
                        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{img_info.get('image_ext', 'png')}") as tmp_file:
                            tmp_file.write(img_info.get('image_data', b''))
                            tmp_path = tmp_file.name
                        
                        # 이미지 크기 계산 (PDF 좌표를 인치로 변환)
                        # PDF 좌표는 포인트 단위 (72 DPI 기준)
                        img_width_pt = img_info.get('width', 0)
                        img_height_pt = img_info.get('height', 0)
                        
                        # 포인트를 인치로 변환
                        img_width_inch = img_width_pt / 72.0
                        img_height_inch = img_height_pt / 72.0
                        
                        # 페이지 너비에 맞게 조정 (너무 크면 축소)
                        max_width_inch = 6.5  # A4 페이지 너비에서 마진 제외
                        if img_width_inch > max_width_inch:
                            ratio = max_width_inch / img_width_inch
                            img_width_inch = max_width_inch
                            img_height_inch = img_height_inch * ratio
                        
                        # 단락 생성 및 이미지 추가
                        para = doc.add_paragraph()
                        run = para.add_run()
                        run.add_picture(tmp_path, width=Inches(img_width_inch))
                        
                        # 이미지 정렬 (중앙 정렬)
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # 임시 파일 삭제
                        import os
                        if os.path.exists(tmp_path):
                            os.unlink(tmp_path)
                        
                        print(f"[DEBUG PDFToWord] 이미지 추가: {img_width_inch:.2f} x {img_height_inch:.2f} 인치")
                        prev_y = element_y
                        continue
                    except Exception as e:
                        print(f"[DEBUG PDFToWord] 이미지 추가 실패: {e}")
                        continue
                
                # 표인 경우
                if element.get('type') == 'table':
                    item = element.get('data', {})
                    table_data = item.get('table_data', {})
                    table_rows = table_data.get('rows', [])
                    
                    if table_rows:
                        # Word 표 생성
                        num_rows = len(table_rows)
                        num_cols = max(len(row) for row in table_rows) if table_rows else 1
                        
                        # 표 생성
                        table = doc.add_table(rows=num_rows, cols=num_cols)
                        table.style = 'Light Grid Accent 1'  # 표 스타일
                        
                        # 표에 데이터 채우기
                        for row_idx, row_data in enumerate(table_rows):
                            for col_idx, cell_data in enumerate(row_data):
                                if row_idx < num_rows and col_idx < num_cols:
                                    cell = table.rows[row_idx].cells[col_idx]
                                    cell_text = cell_data.get('text', '')
                                    
                                    # 숫자나 데이터를 의미 기반 마커로 변환
                                    processed_text = self._convert_data_to_semantic_markers(
                                        cell_text, [cell_data]
                                    )
                                    
                                    # 셀에 텍스트 추가
                                    cell_para = cell.paragraphs[0]
                                    cell_para.clear()
                                    run = cell_para.add_run(processed_text)
                                    
                                    # 폰트 설정
                                    font_size = cell_data.get('font_size', 10)
                                    run.font.size = Pt(font_size)
                                    run.font.name = '맑은 고딕'
                                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
                                    
                                    if cell_data.get('is_bold', False):
                                        run.bold = True
                                    
                                    # 셀 정렬 (첫 번째 행은 중앙 정렬, 나머지는 왼쪽 정렬)
                                    if row_idx == 0:
                                        cell_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    else:
                                        cell_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        
                        print(f"[DEBUG PDFToWord] 표 추가: {num_rows}행 x {num_cols}열")
                        doc.add_paragraph()  # 표 다음에 빈 단락 추가
                        prev_y = element_y
                        continue
                
                # 단락인 경우
                if element.get('type') == 'paragraph':
                    item = element.get('data', {})
                    if not item:
                        continue
                    
                    para = doc.add_paragraph()
                    
                    # 텍스트 아이템들을 순서대로 추가 (폰트 정보 유지)
                    items = item.get('items', [])
                    if items and isinstance(items, list):
                        for idx, text_item in enumerate(items):
                            if not isinstance(text_item, dict):
                                continue
                            
                            text = text_item.get('text', '')
                            if not text:
                                continue
                            
                            # 숫자나 데이터를 의미 기반 마커로 변환 (단일 아이템 기준)
                            processed_text = self._convert_data_to_semantic_markers(text, [text_item])
                            
                            # 각 텍스트 아이템을 별도 Run으로 추가 (폰트 정보 유지)
                            run = para.add_run(processed_text)
                            
                            # 폰트 크기 설정 (정확한 크기 유지)
                            font_size = text_item.get('font_size', 12)
                            if font_size:
                                run.font.size = Pt(font_size)
                            
                            # 폰트 이름 설정 (한글 지원)
                            run.font.name = '맑은 고딕'
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
                            
                            # 굵은 글씨 설정
                            if text_item.get('is_bold', False):
                                run.bold = True
                            
                            # 다음 아이템과 공백 추가 (같은 줄의 경우)
                            if idx < len(items) - 1 and isinstance(items[idx + 1], dict):
                                # 다음 아이템과의 x 좌표 차이 확인
                                current_bbox = text_item.get('bbox', [0, 0, 0, 0])
                                next_bbox = items[idx + 1].get('bbox', [0, 0, 0, 0])
                                if len(current_bbox) >= 3 and len(next_bbox) >= 1:
                                    current_x = current_bbox[2]
                                    next_x = next_bbox[0]
                                    gap = next_x - current_x
                                    
                                    # 적절한 간격이면 공백 추가
                                    if 5 < gap < 100:  # 5~100 픽셀 간격
                                        run.add_text(' ')
                    else:
                        # 아이템이 없는 경우 빈 단락
                        para.add_run('')
                    
                    # 단락 정렬 (왼쪽 정렬 기본)
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                    # 줄 간격 설정 (원본과 유사하게)
                    para.paragraph_format.line_spacing = 1.15
                    
                    # 단락 간격 설정 (y 좌표 차이 기반)
                    if prev_y is not None:
                        y_diff = element_y - prev_y
                        # 큰 간격이면 단락 간격 추가
                        if y_diff > page_height * 0.02:
                            para.paragraph_format.space_before = Pt(max(6, int(y_diff * 0.1)))
                    
                    prev_y = element_y
                except Exception as e:
                    print(f"[DEBUG PDFToWord] 요소 처리 중 오류 발생: {e}")
                    import traceback
                    traceback.print_exc()
                    # 오류가 발생해도 계속 진행
                    continue
        
        pdf_doc.close()
        
        # Word 파일 저장
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_file))
        print(f"Word 템플릿이 생성되었습니다: {output_path}")
        
        return str(output_file)
    
    def _convert_data_to_semantic_markers(self, text: str, text_items: List[Dict]) -> str:
        """
        텍스트에서 숫자나 데이터를 의미 기반 마커로 변환합니다.
        하드코딩된 시트명 대신 의미 기반 키워드를 사용합니다.
        
        Args:
            text: 원본 텍스트
            text_items: 텍스트 아이템 리스트 (컨텍스트 분석용)
            
        Returns:
            마커가 포함된 텍스트
        """
        import re
        
        # 숫자 패턴 (퍼센트, 소수점, 콤마 포함)
        number_pattern = r'[\d,]+\.?\d*%?'
        
        # 숫자가 포함된 경우 주변 컨텍스트 분석
        def replace_with_marker(match):
            number = match.group(0)
            
            # 주변 텍스트에서 키워드 추출
            start_pos = match.start()
            end_pos = match.end()
            
            # 앞뒤 20자 정도의 컨텍스트 추출
            context_start = max(0, start_pos - 20)
            context_end = min(len(text), end_pos + 20)
            context = text[context_start:context_end]
            
            # 의미 기반 키워드 추출
            sheet_keyword = self._extract_sheet_keyword_from_context(context)
            data_keyword = self._extract_data_keyword_from_context(context, number)
            
            # 마커 생성: {의미키워드:데이터키워드}
            if sheet_keyword and data_keyword:
                return f"{{{sheet_keyword}:{data_keyword}}}"
            elif sheet_keyword:
                return f"{{{sheet_keyword}:값}}"
            else:
                # 키워드를 추출하지 못한 경우 원본 숫자 유지
                return number
        
        # 숫자 패턴을 마커로 변환
        result = re.sub(number_pattern, replace_with_marker, text)
        
        return result
    
    def _extract_sheet_keyword_from_context(self, context: str) -> Optional[str]:
        """컨텍스트에서 시트 키워드 추출"""
        import re
        
        # 키워드 매핑 (의미 기반)
        keyword_patterns = {
            '경제지표': r'경제|지표|gdp|생산|소비|투자',
            '생산': r'생산|제조|manufacturing|production',
            '소비': r'소비|소매|consumption|retail',
            '건설': r'건설|construction|공정',
            '광업': r'광업|mining',
            '제조업': r'제조|manufacturing',
            '서비스업': r'서비스|service',
            '소매업': r'소매|retail',
            '지역': r'지역|region|시도|시군구',
            '전국': r'전국|national|전체',
        }
        
        context_lower = context.lower()
        
        for keyword, pattern in keyword_patterns.items():
            if re.search(pattern, context_lower, re.IGNORECASE):
                return keyword
        
        return None
    
    def _extract_data_keyword_from_context(self, context: str, number: str) -> Optional[str]:
        """컨텍스트에서 데이터 키워드 추출"""
        import re
        
        # 데이터 키워드 패턴
        data_patterns = {
            '증감률': r'증감|증가|감소|변화|변동|growth|change',
            '증가율': r'증가|상승|up|rise',
            '감소율': r'감소|하락|down|fall',
            '전국': r'전국|national|전체',
            '서울': r'서울|seoul',
            '부산': r'부산|busan',
            '대구': r'대구|daegu',
            '인천': r'인천|incheon',
            '광주': r'광주|gwangju',
            '대전': r'대전|daejeon',
            '울산': r'울산|ulsan',
        }
        
        context_lower = context.lower()
        
        keywords = []
        for keyword, pattern in data_patterns.items():
            if re.search(pattern, context_lower, re.IGNORECASE):
                keywords.append(keyword)
        
        if keywords:
            # 여러 키워드가 있으면 언더스코어로 연결
            return '_'.join(keywords)
        
        # 키워드를 찾지 못한 경우 숫자 주변 텍스트에서 추출
        # 예: "전국 2.5%" -> "전국_값"
        words = re.findall(r'[가-힣]+', context)
        if words:
            return '_'.join(words[:2])  # 최대 2개 단어
        
        return None

