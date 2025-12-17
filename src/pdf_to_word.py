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
    from docx.shared import Inches, Pt, RGBColor
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
        
        Args:
            text_data: extract_text_and_layout()의 결과
            page_width: 페이지 너비
            page_height: 페이지 높이
            
        Returns:
            정리된 텍스트 구조 리스트 (단락, 제목, 표 등으로 구분)
        """
        if not text_data:
            return []
        
        # y 좌표 기준으로 정렬
        sorted_text = sorted(text_data, key=lambda x: (x['bbox'][1], x['bbox'][0]))
        
        organized = []
        current_paragraph = []
        current_y = None
        line_threshold = page_height * 0.02
        
        for item in sorted_text:
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
        
        return organized
    
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
        
        # 텍스트가 충분히 추출되었는지 확인
        total_text_chars = sum(len(item['text']) for page_data in pages_text_data for item in page_data)
        use_ocr = total_text_chars < 100  # 100자 미만이면 스캔된 PDF로 간주
        
        if use_ocr:
            print(f"텍스트가 부족합니다 ({total_text_chars}자). 스캔된 PDF로 판단하여 OCR을 사용합니다.")
        else:
            print(f"텍스트 기반 PDF로 확인되었습니다 ({total_text_chars}자 추출). 직접 추출 방식을 사용합니다.")
        
        # Word 문서 생성
        doc = Document()
        
        # 페이지 설정 (A4 크기)
        section = doc.sections[0]
        section.page_height = Inches(11.69)  # A4 높이
        section.page_width = Inches(8.27)    # A4 너비
        section.left_margin = Inches(0.79)
        section.right_margin = Inches(0.79)
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        
        # PDF 열기 (페이지 크기 정보를 위해)
        pdf_doc = fitz.open(pdf_path)
        total_pages = len(pdf_doc)
        
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
                organized_text = self.organize_text_by_layout(text_data, image.size[0], image.size[1])
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
                para = doc.add_paragraph()
                para.add_run(f"(페이지 {page_num + 1} - 텍스트를 추출할 수 없습니다)")
                continue
            
            # Word 문서에 추가
            prev_y = None
            for item in organized_text:
                # 큰 간격이 있으면 단락 구분
                if prev_y is not None and (item['y'] - prev_y) > page_height * 0.05:
                    doc.add_paragraph()  # 빈 단락 추가
                
                # 단락 생성
                para = doc.add_paragraph()
                
                # 텍스트 아이템들을 순서대로 추가
                for text_item in item['items']:
                    run = para.add_run(text_item['text'])
                    
                    # 폰트 크기 설정
                    font_size = text_item.get('font_size', 12)
                    run.font.size = Pt(font_size)
                    
                    # 폰트 이름 설정 (한글 지원)
                    run.font.name = '맑은 고딕'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
                    
                    # 굵은 글씨 설정
                    if text_item.get('is_bold', False):
                        run.bold = True
                    
                    # 공백 추가 (같은 줄의 다음 텍스트)
                    if text_item != item['items'][-1]:
                        run.add_text(' ')
                
                # 단락 정렬 (왼쪽 정렬 기본)
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                # 줄 간격 설정
                para.paragraph_format.line_spacing = 1.15
                
                prev_y = item['y']
        
        pdf_doc.close()
        
        # Word 파일 저장
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_file))
        print(f"Word 템플릿이 생성되었습니다: {output_path}")
        
        return str(output_file)

