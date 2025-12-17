"""
템플릿 생성 모듈
이미지에서 HTML 템플릿을 자동 생성하는 기능 제공
"""

import re
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from PIL import Image
import pytesseract

try:
    import easyocr
    EASYOCR_AVAILABLE = True
except ImportError:
    EASYOCR_AVAILABLE = False

try:
    import cv2
    import numpy as np
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False


class TemplateGenerator:
    """이미지에서 HTML 템플릿을 생성하는 클래스"""
    
    def __init__(self, use_easyocr: bool = True):
        """
        템플릿 생성기 초기화
        
        Args:
            use_easyocr: True면 easyocr 사용, False면 pytesseract 사용
        """
        self.use_easyocr = use_easyocr and EASYOCR_AVAILABLE
        self.reader = None
        
        if self.use_easyocr:
            try:
                self.reader = easyocr.Reader(['ko', 'en'], gpu=False)
            except Exception as e:
                print(f"EasyOCR 초기화 실패, pytesseract 사용: {e}")
                self.use_easyocr = False
                self.reader = None
    
    def extract_text_from_image(self, image_path: str) -> List[Dict]:
        """
        이미지에서 텍스트와 위치 정보를 추출합니다.
        
        Args:
            image_path: 이미지 파일 경로
            
        Returns:
            텍스트 정보 딕셔너리 리스트. 각 딕셔너리는 다음 키를 포함:
            - 'text': 추출된 텍스트
            - 'bbox': 바운딩 박스 좌표 (x1, y1, x2, y2)
            - 'confidence': 신뢰도 (0-1)
        """
        if CV2_AVAILABLE:
            image = cv2.imread(image_path)
            if image is None:
                raise ValueError(f"이미지를 읽을 수 없습니다: {image_path}")
            
            # 이미지 전처리
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            
            # 노이즈 제거
            denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
            
            # 이진화
            _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        else:
            # PIL만 사용
            img = Image.open(image_path)
            binary = img.convert('L')
        
        text_data = []
        
        if self.use_easyocr and self.reader and CV2_AVAILABLE:
            # EasyOCR 사용
            results = self.reader.readtext(binary)
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
        else:
            # pytesseract 사용
            try:
                # 이미지를 PIL Image로 변환
                if CV2_AVAILABLE:
                    pil_image = Image.fromarray(binary)
                else:
                    pil_image = binary
                
                # OCR 수행
                data = pytesseract.image_to_data(pil_image, lang='kor+eng', output_type=pytesseract.Output.DICT)
                
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
    
    def analyze_layout(self, text_data: List[Dict], image_width: int, image_height: int) -> Dict:
        """
        텍스트 데이터를 분석하여 레이아웃 구조를 파악합니다.
        
        Args:
            text_data: extract_text_from_image()의 결과
            image_width: 이미지 너비
            image_height: 이미지 높이
            
        Returns:
            레이아웃 정보 딕셔너리
        """
        if not text_data:
            return {
                'sections': [],
                'tables': [],
                'paragraphs': []
            }
        
        # 텍스트를 y 좌표 기준으로 정렬
        sorted_text = sorted(text_data, key=lambda x: x['bbox'][1])
        
        # 섹션 구분 (y 좌표 차이가 큰 경우)
        sections = []
        current_section = []
        prev_y = None
        
        for item in sorted_text:
            y = item['bbox'][1]
            if prev_y is not None and y - prev_y > image_height * 0.05:  # 5% 이상 차이면 새 섹션
                if current_section:
                    sections.append(current_section)
                current_section = [item]
            else:
                current_section.append(item)
            prev_y = y
        
        if current_section:
            sections.append(current_section)
        
        # 테이블 감지 (정렬된 텍스트가 여러 줄에 걸쳐 있는 경우)
        tables = []
        # 간단한 테이블 감지 로직 (개선 가능)
        
        return {
            'sections': sections,
            'tables': tables,
            'paragraphs': sorted_text
        }
    
    def detect_data_markers(self, text: str) -> List[str]:
        """
        텍스트에서 데이터 마커로 변환할 수 있는 부분을 감지합니다.
        
        Args:
            text: 분석할 텍스트
            
        Returns:
            마커로 변환 가능한 텍스트 패턴 리스트
        """
        markers = []
        
        # 숫자 패턴 (퍼센트, 소수점 등)
        number_patterns = [
            r'-?\d+\.?\d*%',  # 퍼센트
            r'-?\d+\.?\d*',    # 일반 숫자
            r'\d{4}년',        # 연도
            r'\d{1,2}분기',    # 분기
        ]
        
        for pattern in number_patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                markers.append(match.group())
        
        # 지역명 패턴 (한글 지역명)
        region_pattern = r'[가-힣]+(?:시|도|군|구|시|읍|면)'
        region_matches = re.finditer(region_pattern, text)
        for match in region_matches:
            markers.append(match.group())
        
        return markers
    
    def generate_html_template(
        self, 
        image_path: str, 
        template_name: str,
        default_sheet_name: str = "시트1"
    ) -> str:
        """
        이미지에서 HTML 템플릿을 생성합니다.
        
        Args:
            image_path: 이미지 파일 경로
            template_name: 생성할 템플릿 이름
            default_sheet_name: 기본 시트명
            
        Returns:
            생성된 HTML 템플릿 문자열
        """
        # 이미지 크기 가져오기
        img = Image.open(image_path)
        image_width, image_height = img.size
        
        # 텍스트 추출
        text_data = self.extract_text_from_image(image_path)
        
        if not text_data:
            # 텍스트가 없으면 기본 템플릿 생성
            return self._generate_default_template(template_name)
        
        # 레이아웃 분석
        layout = self.analyze_layout(text_data, image_width, image_height)
        
        # HTML 생성
        html_parts = []
        html_parts.append('<!DOCTYPE html>')
        html_parts.append('<html lang="ko">')
        html_parts.append('<head>')
        html_parts.append('    <meta charset="UTF-8">')
        html_parts.append('    <meta name="viewport" content="width=device-width, initial-scale=1.0">')
        html_parts.append(f'    <title>{template_name}</title>')
        html_parts.append('    <style>')
        html_parts.append(self._generate_css())
        html_parts.append('    </style>')
        html_parts.append('</head>')
        html_parts.append('<body>')
        
        # 본문 생성
        html_parts.append(self._generate_body_content(text_data, layout, default_sheet_name))
        
        html_parts.append('</body>')
        html_parts.append('</html>')
        
        return '\n'.join(html_parts)
    
    def _generate_css(self) -> str:
        """기본 CSS 스타일 생성"""
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
            max-width: 1000px;
            margin: 0 auto;
            font-size: 14px;
        }
        
        .document-title {
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 20px;
            color: #000;
            line-height: 1.4;
        }
        
        .section-title {
            font-size: 16px;
            font-weight: bold;
            margin: 25px 0 12px 0;
            color: #000;
            line-height: 1.4;
        }
        
        .content-text {
            font-size: 14px;
            margin-bottom: 12px;
            text-align: justify;
            line-height: 1.7;
            color: #000;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 25px 0;
            font-size: 13px;
            border: 1px solid #000;
        }
        
        th, td {
            padding: 8px 6px;
            text-align: center;
            border: 1px solid #000;
            font-size: 12px;
        }
        
        th {
            background-color: #f5f5f5;
            font-weight: bold;
        }"""
    
    def _generate_body_content(
        self, 
        text_data: List[Dict], 
        layout: Dict,
        default_sheet_name: str
    ) -> str:
        """본문 내용 생성"""
        content_parts = []
        
        # 텍스트를 y 좌표 기준으로 정렬
        sorted_text = sorted(text_data, key=lambda x: (x['bbox'][1], x['bbox'][0]))
        
        current_y = None
        current_line = []
        
        for item in sorted_text:
            y = item['bbox'][1]
            text = item['text']
            
            # 같은 줄인지 판단 (y 좌표 차이가 작으면)
            if current_y is None or abs(y - current_y) < 20:
                current_line.append(item)
                if current_y is None:
                    current_y = y
            else:
                # 이전 줄 처리
                if current_line:
                    line_html = self._process_line(current_line, default_sheet_name)
                    if line_html:
                        content_parts.append(line_html)
                
                # 새 줄 시작
                current_line = [item]
                current_y = y
        
        # 마지막 줄 처리
        if current_line:
            line_html = self._process_line(current_line, default_sheet_name)
            if line_html:
                content_parts.append(line_html)
        
        return '\n    '.join(content_parts) if content_parts else '    <div class="content-text">템플릿 내용이 여기에 표시됩니다.</div>'
    
    def _process_line(self, line_items: List[Dict], default_sheet_name: str) -> str:
        """한 줄의 텍스트를 HTML로 변환"""
        if not line_items:
            return ''
        
        # 텍스트 합치기
        line_text = ' '.join([item['text'] for item in line_items])
        
        # 숫자나 데이터로 보이는 부분을 마커로 변환
        processed_text = self._replace_with_markers(line_text, default_sheet_name)
        
        # 제목인지 본문인지 판단 (폰트 크기나 위치 기반으로, 여기서는 간단히 처리)
        if len(line_text) < 50 and any(keyword in line_text for keyword in ['제목', '제', '장', '절']):
            return f'    <div class="section-title">{processed_text}</div>'
        else:
            return f'    <div class="content-text">{processed_text}</div>'
    
    def _replace_with_markers(self, text: str, sheet_name: str) -> str:
        """텍스트의 숫자 부분을 마커로 변환"""
        # 숫자 패턴 찾기
        patterns = [
            (r'-?\d+\.?\d*%', f'{{{sheet_name}:A1}}'),  # 퍼센트
            (r'-?\d+\.?\d*', f'{{{sheet_name}:A1}}'),   # 일반 숫자
        ]
        
        processed = text
        for pattern, marker in patterns:
            processed = re.sub(pattern, marker, processed)
        
        return processed
    
    def _generate_default_template(self, template_name: str) -> str:
        """기본 템플릿 생성 (텍스트 추출 실패 시)"""
        return f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{template_name}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Malgun Gothic', '맑은 고딕', 'Apple SD Gothic Neo', sans-serif;
            line-height: 1.6;
            color: #000;
            background-color: #fff;
            padding: 30px 40px;
            max-width: 1000px;
            margin: 0 auto;
            font-size: 14px;
        }}
        
        .content-text {{
            font-size: 14px;
            margin-bottom: 12px;
            text-align: justify;
            line-height: 1.7;
            color: #000;
        }}
    </style>
</head>
<body>
    <div class="content-text">템플릿 내용을 수동으로 편집해주세요.</div>
    <div class="content-text">데이터 마커 형식: {{시트명:셀주소}} 또는 {{시트명:셀주소:계산식}}</div>
</body>
</html>"""

