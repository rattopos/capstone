"""
PDF를 이미지로 변환하고 OCR을 통해 HTML로 생성하는 모듈
"""

import os
from pathlib import Path
from typing import List, Dict, Optional
from PIL import Image
import io

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

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


class PDFToHTMLConverter:
    """PDF를 이미지로 변환하고 OCR을 통해 HTML로 생성하는 클래스"""
    
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
    
    def pdf_to_images(self, pdf_path: str) -> List[Image.Image]:
        """
        PDF 파일을 이미지 리스트로 변환합니다.
        
        Args:
            pdf_path: PDF 파일 경로
            
        Returns:
            PIL Image 객체 리스트
        """
        images = []
        
        if PYMUPDF_AVAILABLE:
            # PyMuPDF 사용 (더 빠르고 간단)
            try:
                doc = fitz.open(pdf_path)
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    # DPI에 맞춰 변환
                    zoom = self.dpi / 72.0  # 기본 DPI는 72
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat)
                    img_data = pix.tobytes("png")
                    img = Image.open(io.BytesIO(img_data))
                    images.append(img)
                doc.close()
                return images
            except Exception as e:
                print(f"PyMuPDF 변환 실패: {e}")
                if not PDF2IMAGE_AVAILABLE:
                    raise ValueError(f"PDF를 이미지로 변환할 수 없습니다: {e}")
        
        if PDF2IMAGE_AVAILABLE:
            # pdf2image 사용 (poppler 필요)
            try:
                images = convert_from_path(pdf_path, dpi=self.dpi)
                return images
            except Exception as e:
                raise ValueError(f"PDF를 이미지로 변환할 수 없습니다: {e}")
        
        raise ValueError("PDF를 이미지로 변환할 수 있는 라이브러리가 설치되어 있지 않습니다. PyMuPDF 또는 pdf2image를 설치해주세요.")
    
    def extract_text_from_image(self, image: Image.Image) -> List[Dict]:
        """
        이미지에서 텍스트와 위치 정보를 추출합니다.
        
        Args:
            image: PIL Image 객체
            
        Returns:
            텍스트 정보 딕셔너리 리스트. 각 딕셔너리는 다음 키를 포함:
            - 'text': 추출된 텍스트
            - 'bbox': 바운딩 박스 좌표 (x1, y1, x2, y2)
            - 'confidence': 신뢰도 (0-1)
        """
        text_data = []
        
        if CV2_AVAILABLE:
            # OpenCV를 사용한 이미지 전처리
            img_array = np.array(image)
            if len(img_array.shape) == 3:
                gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
            else:
                gray = img_array
            
            # 노이즈 제거
            denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
            
            # 이진화
            _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            processed_image = Image.fromarray(binary)
        else:
            # PIL만 사용
            processed_image = image.convert('L')
        
        if self.use_easyocr and self.reader:
            # EasyOCR 사용
            try:
                if CV2_AVAILABLE:
                    img_array = np.array(processed_image)
                    results = self.reader.readtext(img_array)
                else:
                    results = self.reader.readtext(processed_image)
                
                for (bbox, text, confidence) in results:
                    # bbox는 [[x1, y1], [x2, y1], [x2, y2], [x1, y2]] 형식
                    x_coords = [point[0] for point in bbox]
                    y_coords = [point[1] for point in bbox]
                    x1, y1 = min(x_coords), min(y_coords)
                    x2, y2 = max(x_coords), max(y_coords)
                    
                    text_data.append({
                        'text': text.strip(),
                        'bbox': (x1, y1, x2, y2),
                        'confidence': confidence
                    })
            except Exception as e:
                print(f"EasyOCR 처리 중 오류: {e}")
                # pytesseract로 폴백
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
                    
                    text_data.append({
                        'text': text,
                        'bbox': (x1, y1, x2, y2),
                        'confidence': confidence
                    })
        except Exception as e:
            raise ValueError(f"OCR 처리 중 오류 발생: {e}")
        
        return text_data
    
    def organize_text_by_layout(self, text_data: List[Dict], page_width: int, page_height: int) -> List[Dict]:
        """
        텍스트 데이터를 레이아웃에 따라 정리합니다.
        
        Args:
            text_data: extract_text_from_image()의 결과
            page_width: 페이지 너비
            page_height: 페이지 높이
            
        Returns:
            정리된 텍스트 구조 리스트
        """
        if not text_data:
            return []
        
        # y 좌표 기준으로 정렬
        sorted_text = sorted(text_data, key=lambda x: (x['bbox'][1], x['bbox'][0]))
        
        organized = []
        current_line = []
        current_y = None
        line_threshold = page_height * 0.02  # 페이지 높이의 2% 차이면 다른 줄
        
        for item in sorted_text:
            y = item['bbox'][1]
            
            if current_y is None or abs(y - current_y) < line_threshold:
                # 같은 줄
                current_line.append(item)
                if current_y is None:
                    current_y = y
            else:
                # 새 줄
                if current_line:
                    # 현재 줄을 x 좌표 기준으로 정렬
                    current_line.sort(key=lambda x: x['bbox'][0])
                    line_text = ' '.join([item['text'] for item in current_line])
                    organized.append({
                        'text': line_text,
                        'y': current_y,
                        'items': current_line
                    })
                
                current_line = [item]
                current_y = y
        
        # 마지막 줄 처리
        if current_line:
            current_line.sort(key=lambda x: x['bbox'][0])
            line_text = ' '.join([item['text'] for item in current_line])
            organized.append({
                'text': line_text,
                'y': current_y,
                'items': current_line
            })
        
        return organized
    
    def generate_html_from_pdf(self, pdf_path: str, output_path: Optional[str] = None) -> str:
        """
        PDF 파일을 OCR을 통해 HTML로 변환합니다.
        
        Args:
            pdf_path: PDF 파일 경로
            output_path: 출력 HTML 파일 경로 (None이면 반환만)
            
        Returns:
            생성된 HTML 문자열
        """
        print(f"PDF 파일 로딩 중: {pdf_path}")
        
        # PDF를 이미지로 변환
        images = self.pdf_to_images(pdf_path)
        print(f"총 {len(images)}페이지를 이미지로 변환했습니다.")
        
        # HTML 생성
        html_parts = []
        html_parts.append('<!DOCTYPE html>')
        html_parts.append('<html lang="ko">')
        html_parts.append('<head>')
        html_parts.append('    <meta charset="UTF-8">')
        html_parts.append('    <meta name="viewport" content="width=device-width, initial-scale=1.0">')
        html_parts.append('    <title>지역경제동향 보도자료</title>')
        html_parts.append('    <style>')
        html_parts.append(self._generate_css())
        html_parts.append('    </style>')
        html_parts.append('</head>')
        html_parts.append('<body>')
        
        # 각 페이지 처리
        for page_num, image in enumerate(images, 1):
            print(f"페이지 {page_num}/{len(images)} OCR 처리 중...")
            
            # 텍스트 추출
            text_data = self.extract_text_from_image(image)
            
            if not text_data:
                html_parts.append(f'    <div class="page-break"></div>')
                html_parts.append(f'    <div class="page-number">페이지 {page_num}</div>')
                html_parts.append(f'    <div class="content-text">(텍스트를 추출할 수 없습니다)</div>')
                continue
            
            # 레이아웃 정리
            page_width, page_height = image.size
            organized_text = self.organize_text_by_layout(text_data, page_width, page_height)
            
            # 페이지 구분
            if page_num > 1:
                html_parts.append('    <div class="page-break"></div>')
            
            html_parts.append(f'    <div class="page-number">페이지 {page_num}</div>')
            
            # 텍스트를 HTML로 변환
            prev_y = None
            for item in organized_text:
                text = item['text']
                y = item['y']
                
                # 큰 간격이 있으면 단락 구분
                if prev_y is not None and (y - prev_y) > page_height * 0.05:
                    html_parts.append('    <div class="paragraph-spacing"></div>')
                
                # 제목인지 본문인지 판단
                if self._is_title(text):
                    html_parts.append(f'    <div class="section-title">{self._escape_html(text)}</div>')
                elif self._is_table_row(text):
                    html_parts.append(f'    <div class="table-row">{self._escape_html(text)}</div>')
                else:
                    html_parts.append(f'    <div class="content-text">{self._escape_html(text)}</div>')
                
                prev_y = y
        
        html_parts.append('</body>')
        html_parts.append('</html>')
        
        html_content = '\n'.join(html_parts)
        
        # 파일로 저장
        if output_path:
            output_file = Path(output_path)
            output_file.parent.mkdir(parents=True, exist_ok=True)
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            print(f"HTML 파일이 생성되었습니다: {output_path}")
        
        return html_content
    
    def _is_title(self, text: str) -> bool:
        """텍스트가 제목인지 판단"""
        # 짧고 굵은 글씨 패턴, 또는 특정 키워드 포함
        title_keywords = ['제', '장', '절', 'I.', 'II.', 'III.', 'IV.', 'V.', '목차', '요약']
        if len(text) < 100 and any(keyword in text for keyword in title_keywords):
            return True
        return False
    
    def _is_table_row(self, text: str) -> bool:
        """텍스트가 테이블 행인지 판단"""
        # 여러 숫자나 구분자가 있는 경우
        import re
        parts = re.split(r'\s+', text.strip())
        if len(parts) >= 3:
            # 숫자나 퍼센트가 많으면 테이블 행일 가능성
            number_count = sum(1 for part in re.findall(r'-?\d+\.?\d*%?', ' '.join(parts)))
            if number_count >= 2:
                return True
        return False
    
    def _escape_html(self, text: str) -> str:
        """HTML 특수문자 이스케이프"""
        return (text.replace('&', '&amp;')
                   .replace('<', '&lt;')
                   .replace('>', '&gt;')
                   .replace('"', '&quot;')
                   .replace("'", '&#39;'))
    
    def _generate_css(self) -> str:
        """CSS 스타일 생성"""
        return """        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Malgun Gothic', '맑은 고딕', 'Apple SD Gothic Neo', sans-serif;
            line-height: 1.6;
            color: #000;
            background-color: #fff;
            padding: 30px 40px;
            max-width: 1200px;
            margin: 0 auto;
            font-size: 14px;
        }
        
        .page-break {
            page-break-before: always;
            border-top: 2px solid #ccc;
            margin-top: 40px;
            padding-top: 20px;
        }
        
        .page-number {
            text-align: center;
            font-size: 12px;
            color: #666;
            margin-bottom: 20px;
        }
        
        .section-title {
            font-size: 18px;
            font-weight: bold;
            margin: 25px 0 15px 0;
            color: #000;
            line-height: 1.4;
        }
        
        .content-text {
            font-size: 14px;
            margin-bottom: 10px;
            text-align: justify;
            line-height: 1.7;
            color: #000;
        }
        
        .table-row {
            font-size: 13px;
            margin-bottom: 5px;
            font-family: 'Courier New', monospace;
            white-space: pre;
        }
        
        .paragraph-spacing {
            height: 15px;
        }"""

