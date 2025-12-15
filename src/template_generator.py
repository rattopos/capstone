"""
스크린샷 기반 템플릿 자동 생성 모듈
스크린샷 이미지를 분석하여 광공업생산 템플릿과 유사한 구조의 HTML 템플릿 생성
"""

import re
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple

try:
    from PIL import Image
    import pytesseract
    import cv2
    import numpy as np
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


class TemplateGenerator:
    """스크린샷에서 템플릿을 자동 생성하는 클래스"""
    
    def __init__(self):
        """템플릿 생성기 초기화"""
        if not OCR_AVAILABLE:
            raise ImportError(
                "OCR 라이브러리가 설치되지 않았습니다. "
                "다음 명령어로 설치하세요: pip install Pillow pytesseract opencv-python"
            )
        # Tesseract OCR 설정 (한국어 + 영어)
        self.tesseract_config = '--oem 3 --psm 6 -l kor+eng'
    
    def load_image(self, image_path: str) -> Image.Image:
        """
        이미지를 로드합니다.
        
        Args:
            image_path: 이미지 파일 경로
            
        Returns:
            PIL Image 객체
        """
        return Image.open(image_path)
    
    def preprocess_image(self, image: Image.Image) -> np.ndarray:
        """
        이미지를 전처리합니다 (OCR 정확도 향상).
        
        Args:
            image: PIL Image 객체
            
        Returns:
            전처리된 이미지 (numpy array)
        """
        # PIL Image를 numpy array로 변환
        img_array = np.array(image)
        
        # 그레이스케일 변환
        if len(img_array.shape) == 3:
            gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
        else:
            gray = img_array
        
        # 노이즈 제거
        denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
        
        # 이진화 (thresholding)
        _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        return binary
    
    def extract_text(self, image: Image.Image) -> str:
        """
        이미지에서 텍스트를 추출합니다 (OCR).
        
        Args:
            image: PIL Image 객체
            
        Returns:
            추출된 텍스트
        """
        # 이미지 전처리
        processed = self.preprocess_image(image)
        
        # numpy array를 PIL Image로 변환
        processed_image = Image.fromarray(processed)
        
        # OCR 수행
        try:
            text = pytesseract.image_to_string(processed_image, config=self.tesseract_config)
            return text
        except Exception as e:
            # Tesseract가 설치되지 않았거나 오류 발생 시
            raise RuntimeError(f"OCR 오류: {e}. Tesseract OCR이 설치되어 있는지 확인하세요.")
    
    def parse_structure(self, text: str) -> Dict[str, Any]:
        """
        추출된 텍스트에서 문서 구조를 파악합니다.
        
        Args:
            text: 추출된 텍스트
            
        Returns:
            구조 정보 딕셔너리
        """
        lines = text.split('\n')
        structure = {
            'title': None,
            'sections': [],
            'tables': [],
            'text_blocks': []
        }
        
        current_section = None
        current_text = []
        
        for line in lines:
            line = line.strip()
            if not line:
                if current_text:
                    structure['text_blocks'].append(' '.join(current_text))
                    current_text = []
                continue
            
            # 제목 감지 (짧고 큰 폰트로 가정)
            if len(line) < 50 and not current_section:
                if any(keyword in line for keyword in ['부문별', '지역경제', '동향', '보도자료']):
                    structure['title'] = line
                    continue
            
            # 섹션 제목 감지 (숫자로 시작하는 패턴)
            section_match = re.match(r'^(\d+\.?\s*[가-힣]+)', line)
            if section_match:
                if current_text:
                    structure['text_blocks'].append(' '.join(current_text))
                    current_text = []
                current_section = section_match.group(1)
                structure['sections'].append({
                    'title': current_section,
                    'content': []
                })
                continue
            
            # 하위 섹션 감지
            subsection_match = re.match(r'^([가-힣]\.\s*[가-힣]+)', line)
            if subsection_match:
                if current_section:
                    structure['sections'][-1]['content'].append({
                        'type': 'subsection',
                        'title': subsection_match.group(1),
                        'text': []
                    })
                continue
            
            # 텍스트 블록에 추가
            current_text.append(line)
        
        if current_text:
            structure['text_blocks'].append(' '.join(current_text))
        
        return structure
    
    def generate_markers(self, text: str, sheet_name: str) -> List[Tuple[str, str]]:
        """
        텍스트에서 마커를 생성합니다.
        
        Args:
            text: 추출된 텍스트
            sheet_name: 시트 이름
            
        Returns:
            (원본 텍스트, 마커) 튜플 리스트
        """
        markers = []
        
        # 숫자 패턴 찾기 (퍼센트 포함)
        percent_pattern = r'([+-]?\d+\.?\d*)\s*%'
        number_pattern = r'(\d{1,3}(?:,\d{3})*(?:\.\d+)?)'
        
        # 지역명 패턴
        region_pattern = r'(전국|서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주)'
        
        # 퍼센트 값 마커 생성
        for match in re.finditer(percent_pattern, text):
            value = match.group(1)
            # 동적 마커 생성 (예: {시트명:지역명_증감률})
            marker = f"{{{sheet_name}:동적_증감률}}"
            markers.append((match.group(0), marker))
        
        # 지역명 마커 생성
        for match in re.finditer(region_pattern, text):
            region = match.group(1)
            marker = f"{{{sheet_name}:{region}_이름}}"
            markers.append((region, marker))
        
        return markers
    
    def generate_template(self, image_path: str, sheet_name: str, 
                         base_template_path: Optional[str] = None) -> str:
        """
        스크린샷에서 템플릿을 생성합니다.
        
        Args:
            image_path: 스크린샷 이미지 경로
            sheet_name: 시트 이름
            base_template_path: 기본 템플릿 경로 (스타일 참조용, 선택적)
            
        Returns:
            생성된 HTML 템플릿
        """
        # 이미지 로드
        image = self.load_image(image_path)
        
        # 텍스트 추출
        text = self.extract_text(image)
        
        # 구조 파악
        structure = self.parse_structure(text)
        
        # 기본 템플릿 스타일 로드 (있는 경우)
        base_style = ""
        if base_template_path and Path(base_template_path).exists():
            with open(base_template_path, 'r', encoding='utf-8') as f:
                base_content = f.read()
                # CSS 스타일 추출
                style_match = re.search(r'<style[^>]*>(.*?)</style>', base_content, re.DOTALL)
                if style_match:
                    base_style = style_match.group(1)
        
        # HTML 생성
        html = self._generate_html(structure, sheet_name, base_style)
        
        return html
    
    def _generate_html(self, structure: Dict[str, Any], sheet_name: str, base_style: str) -> str:
        """
        구조 정보로부터 HTML을 생성합니다.
        
        Args:
            structure: 구조 정보 딕셔너리
            sheet_name: 시트 이름
            base_style: 기본 CSS 스타일
            
        Returns:
            HTML 문자열
        """
        html_parts = [
            '<!DOCTYPE html>',
            '<html lang="ko">',
            '<head>',
            '    <meta charset="UTF-8">',
            '    <meta name="viewport" content="width=device-width, initial-scale=1.0">',
            f'    <title>{structure.get("title", "보도자료")}</title>',
            '    <style>',
        ]
        
        # 기본 스타일 추가
        if base_style:
            html_parts.append(base_style)
        else:
            # 기본 스타일 (광공업생산 템플릿과 유사)
            html_parts.append("""
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Malgun Gothic', '맑은 고딕', 'Apple SD Gothic Neo', sans-serif;
            line-height: 1.8;
            color: #333;
            background-color: #fff;
            padding: 40px 20px;
            max-width: 1200px;
            margin: 0 auto;
        }
        
        .document-title {
            font-size: 24px;
            font-weight: bold;
            margin-bottom: 20px;
            color: #1a1a1a;
        }
        
        .section-title {
            font-size: 20px;
            font-weight: bold;
            margin: 30px 0 15px 0;
            color: #2c3e50;
        }
        
        .subsection-title {
            font-size: 18px;
            font-weight: bold;
            margin: 25px 0 15px 0;
            color: #34495e;
        }
        
        .content-text {
            font-size: 14px;
            margin-bottom: 15px;
            text-align: justify;
            line-height: 1.9;
        }
            """)
        
        html_parts.extend([
            '    </style>',
            '</head>',
            '<body>',
        ])
        
        # 제목 추가
        if structure.get('title'):
            html_parts.append(f'    <div class="document-title">{structure["title"]}</div>')
        
        # 섹션 추가
        for section in structure.get('sections', []):
            html_parts.append(f'    <div class="section-title">{section["title"]}</div>')
            
            for content in section.get('content', []):
                if content.get('type') == 'subsection':
                    html_parts.append(f'        <div class="subsection-title">{content["title"]}</div>')
        
        # 텍스트 블록 추가
        for text_block in structure.get('text_blocks', []):
            # 마커 생성 및 적용
            markers = self.generate_markers(text_block, sheet_name)
            processed_text = text_block
            for original, marker in markers:
                processed_text = processed_text.replace(original, marker)
            
            html_parts.append(f'    <div class="content-text">{processed_text}</div>')
        
        html_parts.extend([
            '</body>',
            '</html>',
        ])
        
        return '\n'.join(html_parts)
    
    def generate_from_screenshot(self, screenshot_path: str, sheet_name: str,
                                 output_path: Optional[str] = None,
                                 base_template_path: Optional[str] = None) -> str:
        """
        스크린샷에서 템플릿을 생성하고 저장합니다.
        
        Args:
            screenshot_path: 스크린샷 이미지 경로
            sheet_name: 시트 이름
            output_path: 출력 파일 경로 (선택적)
            base_template_path: 기본 템플릿 경로 (선택적)
            
        Returns:
            생성된 HTML 템플릿 문자열
        """
        template_html = self.generate_template(screenshot_path, sheet_name, base_template_path)
        
        if output_path:
            Path(output_path).parent.mkdir(parents=True, exist_ok=True)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(template_html)
        
        return template_html

