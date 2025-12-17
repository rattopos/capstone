"""
스크린샷 기반 템플릿 생성 모듈
correct_answer 폴더의 스크린샷들을 분석하여 템플릿을 자동 생성
엑셀 파일의 시트를 자동으로 매핑
"""

import re
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Any
from PIL import Image
import pytesseract
import cv2
import numpy as np

try:
    import easyocr
    EASYOCR_AVAILABLE = True
except ImportError:
    EASYOCR_AVAILABLE = False

from .excel_extractor import ExcelExtractor
from .excel_header_parser import ExcelHeaderParser
from .flexible_mapper import FlexibleMapper


class ScreenshotTemplateGenerator:
    """스크린샷에서 템플릿을 생성하고 엑셀 시트를 자동 매핑하는 클래스"""
    
    def __init__(self, excel_path: str, use_easyocr: bool = True):
        """
        템플릿 생성기 초기화
        
        Args:
            excel_path: 엑셀 파일 경로
            use_easyocr: True면 easyocr 사용, False면 pytesseract 사용
        """
        self.excel_path = excel_path
        self.excel_extractor = ExcelExtractor(excel_path)
        self.excel_extractor.load_workbook()
        self.header_parser = ExcelHeaderParser(excel_path)
        self.flexible_mapper = FlexibleMapper(self.excel_extractor)
        
        self.use_easyocr = use_easyocr and EASYOCR_AVAILABLE
        self.reader = None
        
        if self.use_easyocr:
            try:
                self.reader = easyocr.Reader(['ko', 'en'], gpu=False)
            except Exception as e:
                print(f"EasyOCR 초기화 실패, pytesseract 사용: {e}")
                self.use_easyocr = False
                self.reader = None
        
        # 엑셀 시트 정보 캐시
        self.sheet_info_cache = {}
        self._cache_sheet_info()
    
    def _cache_sheet_info(self):
        """모든 시트의 헤더 정보를 캐시에 저장"""
        sheet_names = self.excel_extractor.get_sheet_names()
        for sheet_name in sheet_names:
            if sheet_name == '완료체크':
                continue
            try:
                headers_info = self.header_parser.parse_sheet_headers(sheet_name)
                self.sheet_info_cache[sheet_name] = headers_info
            except Exception as e:
                print(f"시트 '{sheet_name}' 정보 캐시 실패: {e}")
    
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
        image = cv2.imread(image_path)
        if image is None:
            raise ValueError(f"이미지를 읽을 수 없습니다: {image_path}")
        
        # 이미지 전처리
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
        _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        text_data = []
        
        if self.use_easyocr and self.reader:
            # EasyOCR 사용
            results = self.reader.readtext(binary)
            for (bbox, text, confidence) in results:
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
            pil_image = Image.fromarray(binary)
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
        
        return text_data
    
    def find_matching_sheet(self, extracted_text: str, text_data: List[Dict]) -> Optional[str]:
        """
        추출된 텍스트를 기반으로 적합한 엑셀 시트를 찾습니다.
        
        Args:
            extracted_text: 추출된 전체 텍스트
            text_data: 텍스트 데이터 리스트
            
        Returns:
            매칭되는 시트명 또는 None
        """
        # 키워드 기반 매핑
        keyword_mapping = {
            '광공업생산': ['광공업', '광업', '제조업', '생산'],
            '서비스업생산': ['서비스업', '서비스'],
            '소비(소매, 추가)': ['소비', '소매', '판매', '백화점', '면세점', '대형마트'],
            '고용': ['고용', '취업'],
            '고용(kosis)': ['고용', 'kosis'],
            '고용률': ['고용률'],
            '실업자 수': ['실업자', '실업'],
            '지출목적별 물가': ['지출목적', '물가'],
            '품목성질별 물가': ['품목성질', '물가'],
            '건설 (공표자료)': ['건설', '수주', '공정'],
            '수출': ['수출'],
            '수입': ['수입'],
            '연령별 인구이동': ['인구이동', '연령별'],
            '시도 간 이동': ['시도', '이동'],
            '시군구인구이동': ['시군구', '인구이동']
        }
        
        # 텍스트에서 키워드 찾기
        text_lower = extracted_text.lower()
        scores = {}
        
        for sheet_name, keywords in keyword_mapping.items():
            score = 0
            for keyword in keywords:
                if keyword in text_lower:
                    score += 1
            if score > 0:
                scores[sheet_name] = score
        
        if scores:
            # 가장 높은 점수의 시트 반환
            best_sheet = max(scores.items(), key=lambda x: x[1])[0]
            return best_sheet
        
        # 키워드 매칭이 실패하면 시트명 직접 매칭 시도
        sheet_names = self.excel_extractor.get_sheet_names()
        for sheet_name in sheet_names:
            if sheet_name == '완료체크':
                continue
            if sheet_name in extracted_text:
                return sheet_name
        
        return None
    
    def detect_numbers_in_text(self, text: str) -> List[Dict]:
        """
        텍스트에서 숫자 패턴을 찾아 반환합니다.
        
        Args:
            text: 분석할 텍스트
            
        Returns:
            숫자 정보 리스트. 각 딕셔너리는:
            - 'value': 숫자 값 (문자열)
            - 'type': 숫자 타입 ('percent', 'number', 'year', 'quarter')
            - 'context': 주변 텍스트 (50자)
        """
        numbers = []
        
        # 퍼센트 패턴
        percent_pattern = r'-?\d+\.?\d*%'
        for match in re.finditer(percent_pattern, text):
            numbers.append({
                'value': match.group(),
                'type': 'percent',
                'context': text[max(0, match.start()-50):min(len(text), match.end()+50)]
            })
        
        # 일반 숫자 패턴 (퍼센트 제외)
        number_pattern = r'-?\d+\.?\d*'
        for match in re.finditer(number_pattern, text):
            value = match.group()
            # 퍼센트가 아니고, 연도나 분기가 아닌 경우
            if '%' not in value and not re.match(r'^\d{4}$', value) and not re.match(r'^\d{1,2}분기$', value):
                numbers.append({
                    'value': value,
                    'type': 'number',
                    'context': text[max(0, match.start()-50):min(len(text), match.end()+50)]
                })
        
        # 연도 패턴
        year_pattern = r'\d{4}년?'
        for match in re.finditer(year_pattern, text):
            numbers.append({
                'value': match.group(),
                'type': 'year',
                'context': text[max(0, match.start()-50):min(len(text), match.end()+50)]
            })
        
        # 분기 패턴
        quarter_pattern = r'\d{1,2}분기'
        for match in re.finditer(quarter_pattern, text):
            numbers.append({
                'value': match.group(),
                'type': 'quarter',
                'context': text[max(0, match.start()-50):min(len(text), match.end()+50)]
            })
        
        return numbers
    
    def find_column_for_number(self, sheet_name: str, number_info: Dict, context: str) -> Optional[str]:
        """
        숫자 정보를 기반으로 적절한 컬럼을 찾습니다.
        
        Args:
            sheet_name: 시트명
            number_info: 숫자 정보 딕셔너리
            context: 주변 텍스트
            
        Returns:
            컬럼 문자 (예: 'A', 'B') 또는 None
        """
        if sheet_name not in self.sheet_info_cache:
            return None
        
        headers_info = self.sheet_info_cache[sheet_name]
        
        # 컨텍스트에서 지역명 찾기
        region_pattern = r'([가-힣]+(?:시|도|군|구|시|읍|면))'
        region_match = re.search(region_pattern, context)
        
        # 컨텍스트에서 헤더 키워드 찾기
        for col_info in headers_info['columns']:
            header_text = col_info['header'].lower()
            context_lower = context.lower()
            
            # 헤더 키워드가 컨텍스트에 포함되어 있는지 확인
            if any(keyword in context_lower for keyword in ['증감', '증감률', '증가', '감소', '생산', '수주']):
                if number_info['type'] == 'percent':
                    # 퍼센트 관련 헤더 찾기
                    if '증감' in header_text or '증감률' in header_text or '%' in header_text:
                        return col_info['letter']
                else:
                    # 일반 숫자 헤더 찾기
                    if '증감' not in header_text and '%' not in header_text:
                        return col_info['letter']
        
        return None
    
    def generate_template_from_screenshot(
        self, 
        screenshot_path: str, 
        template_name: str
    ) -> str:
        """
        스크린샷에서 HTML 템플릿을 생성합니다.
        
        Args:
            screenshot_path: 스크린샷 이미지 파일 경로
            template_name: 생성할 템플릿 이름
            
        Returns:
            생성된 HTML 템플릿 문자열
        """
        # 이미지 크기 가져오기
        img = Image.open(screenshot_path)
        image_width, image_height = img.size
        
        # 텍스트 추출
        text_data = self.extract_text_from_image(screenshot_path)
        
        if not text_data:
            return self._generate_default_template(template_name)
        
        # 전체 텍스트 합치기
        full_text = ' '.join([item['text'] for item in text_data])
        
        # 적합한 시트 찾기
        matched_sheet = self.find_matching_sheet(full_text, text_data)
        if not matched_sheet:
            # 기본값: 첫 번째 시트 사용
            sheet_names = self.excel_extractor.get_sheet_names()
            matched_sheet = sheet_names[0] if sheet_names else '시트1'
        
        # 텍스트를 y 좌표 기준으로 정렬
        sorted_text = sorted(text_data, key=lambda x: (x['bbox'][1], x['bbox'][0]))
        
        # 숫자 감지
        numbers = self.detect_numbers_in_text(full_text)
        
        # HTML 생성
        html_parts = []
        html_parts.append('<!DOCTYPE html>')
        html_parts.append('<html lang="ko">')
        html_parts.append('<head>')
        html_parts.append('    <meta charset="UTF-8">')
        html_parts.append('    <meta name="viewport" content="width=device-width, initial-scale=1.0">')
        html_parts.append(f'    <title>{template_name}</title>')
        html_parts.append('    <style>')
        html_parts.append(self._generate_css(image_width, image_height))
        html_parts.append('    </style>')
        html_parts.append('</head>')
        html_parts.append('<body>')
        
        # 본문 생성
        html_parts.append(self._generate_body_content(
            sorted_text, 
            numbers, 
            matched_sheet,
            image_width,
            image_height
        ))
        
        html_parts.append('</body>')
        html_parts.append('</html>')
        
        return '\n'.join(html_parts)
    
    def _generate_css(self, image_width: int, image_height: int) -> str:
        """이미지 크기를 기반으로 CSS 스타일 생성"""
        # 이미지 비율 계산
        aspect_ratio = image_width / image_height if image_height > 0 else 1.0
        
        return f"""        * {{
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
            max-width: {min(image_width, 1200)}px;
            margin: 0 auto;
            font-size: 14px;
        }}
        
        .document-title {{
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 20px;
            color: #000;
            line-height: 1.4;
            text-align: center;
        }}
        
        .section-title {{
            font-size: 16px;
            font-weight: bold;
            margin: 25px 0 12px 0;
            color: #000;
            line-height: 1.4;
        }}
        
        .content-text {{
            font-size: 14px;
            margin-bottom: 12px;
            text-align: justify;
            line-height: 1.7;
            color: #000;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 25px 0;
            font-size: 13px;
            border: 1px solid #000;
        }}
        
        th, td {{
            padding: 8px 6px;
            text-align: center;
            border: 1px solid #000;
            font-size: 12px;
        }}
        
        th {{
            background-color: #f5f5f5;
            font-weight: bold;
        }}"""
    
    def _detect_table_structure(self, sorted_text: List[Dict]) -> List[List[Dict]]:
        """
        텍스트 데이터에서 테이블 구조를 감지합니다.
        
        Returns:
            테이블 행 리스트 (각 행은 텍스트 아이템 리스트)
        """
        if not sorted_text:
            return []
        
        # y 좌표 기준으로 그룹화
        rows = []
        current_row = []
        current_y = None
        y_threshold = 20  # 같은 줄로 간주할 y 좌표 차이
        
        for item in sorted_text:
            y = item['bbox'][1]
            
            if current_y is None or abs(y - current_y) < y_threshold:
                current_row.append(item)
                if current_y is None:
                    current_y = y
            else:
                if current_row:
                    # x 좌표 기준으로 정렬
                    current_row.sort(key=lambda x: x['bbox'][0])
                    rows.append(current_row)
                current_row = [item]
                current_y = y
        
        if current_row:
            current_row.sort(key=lambda x: x['bbox'][0])
            rows.append(current_row)
        
        # 테이블인지 판단 (여러 행에 걸쳐 비슷한 x 좌표 패턴이 있는지)
        if len(rows) >= 3:
            # 첫 3개 행의 x 좌표 패턴 확인
            first_row_x = [item['bbox'][0] for item in rows[0]]
            second_row_x = [item['bbox'][0] for item in rows[1]]
            third_row_x = [item['bbox'][0] for item in rows[2]]
            
            # x 좌표가 비슷한 패턴을 보이면 테이블로 간주
            if len(first_row_x) >= 2 and len(second_row_x) >= 2:
                return rows
        
        return rows
    
    def _generate_body_content(
        self,
        sorted_text: List[Dict],
        numbers: List[Dict],
        sheet_name: str,
        image_width: int,
        image_height: int
    ) -> str:
        """본문 내용 생성"""
        content_parts = []
        
        # 테이블 구조 감지
        table_rows = self._detect_table_structure(sorted_text)
        
        # 테이블로 보이면 테이블 HTML 생성
        if len(table_rows) >= 3:
            table_html = self._generate_table_html(table_rows, sheet_name)
            if table_html:
                content_parts.append(table_html)
        else:
            # 일반 텍스트로 처리
            current_y = None
            current_line = []
            
            for item in sorted_text:
                y = item['bbox'][1]
                
                # 같은 줄인지 판단 (y 좌표 차이가 작으면)
                if current_y is None or abs(y - current_y) < 20:
                    current_line.append(item)
                    if current_y is None:
                        current_y = y
                else:
                    # 이전 줄 처리
                    if current_line:
                        line_html = self._process_line(
                            current_line, 
                            numbers, 
                            0,
                            sheet_name
                        )
                        if line_html:
                            content_parts.append(line_html)
                    
                    # 새 줄 시작
                    current_line = [item]
                    current_y = y
            
            # 마지막 줄 처리
            if current_line:
                line_html = self._process_line(current_line, numbers, 0, sheet_name)
                if line_html:
                    content_parts.append(line_html)
        
        return '\n    '.join(content_parts) if content_parts else '    <div class="content-text">템플릿 내용이 여기에 표시됩니다.</div>'
    
    def _generate_table_html(self, table_rows: List[List[Dict]], sheet_name: str) -> str:
        """테이블 HTML 생성"""
        if not table_rows:
            return ''
        
        html_parts = []
        html_parts.append('    <table>')
        
        for row_idx, row_items in enumerate(table_rows):
            # 행 텍스트 합치기
            row_text = ' '.join([item['text'] for item in row_items])
            
            # 마커로 변환
            processed_row = self._replace_with_header_markers(row_text, sheet_name)
            
            # 셀 분리 (간단한 방법: 공백으로 분리, 실제로는 x 좌표 기반으로 분리해야 함)
            cells = row_text.split() if len(row_items) == 1 else [item['text'] for item in row_items]
            
            html_parts.append('        <tr>')
            
            for cell_text in cells:
                # 셀 텍스트도 마커로 변환
                processed_cell = self._replace_with_header_markers(cell_text, sheet_name)
                
                # 첫 번째 행이면 th, 아니면 td
                if row_idx == 0:
                    html_parts.append(f'            <th>{processed_cell}</th>')
                else:
                    html_parts.append(f'            <td>{processed_cell}</td>')
            
            html_parts.append('        </tr>')
        
        html_parts.append('    </table>')
        
        return '\n'.join(html_parts)
    
    def _process_line(
        self,
        line_items: List[Dict],
        numbers: List[Dict],
        number_index: int,
        sheet_name: str
    ) -> str:
        """한 줄의 텍스트를 HTML로 변환하고 숫자를 마커로 치환"""
        if not line_items:
            return ''
        
        # 텍스트 합치기
        line_text = ' '.join([item['text'] for item in line_items])
        
        # 헤더 파서를 사용하여 마커 생성
        processed_text = self._replace_with_header_markers(line_text, sheet_name)
        
        # 제목인지 본문인지 판단
        if len(line_text) < 50 and any(keyword in line_text for keyword in ['제목', '제', '장', '절']):
            return f'    <div class="section-title">{processed_text}</div>'
        else:
            return f'    <div class="content-text">{processed_text}</div>'
    
    def _replace_with_header_markers(self, text: str, sheet_name: str) -> str:
        """
        헤더 기반으로 마커를 생성하여 텍스트의 숫자 부분을 변환
        더 정확한 마커 생성: 동적 마커 패턴 사용
        """
        try:
            if sheet_name not in self.sheet_info_cache:
                return text
            
            headers_info = self.sheet_info_cache[sheet_name]
            
            # 숫자 패턴 찾기 (퍼센트, 일반 숫자, 연도, 분기)
            patterns = [
                (r'-?\d+\.?\d*%', 'percent'),  # 퍼센트
                (r'\d{4}년?', 'year'),  # 연도 (마커로 변환하지 않음)
                (r'\d{1,2}분기', 'quarter'),  # 분기 (마커로 변환하지 않음)
                (r'-?\d+\.?\d*', 'number'),  # 일반 숫자
            ]
            
            processed = text
            all_matches = []
            
            # 모든 패턴의 매치 찾기
            for pattern, num_type in patterns:
                for match in re.finditer(pattern, text):
                    all_matches.append({
                        'match': match,
                        'type': num_type,
                        'text': match.group()
                    })
            
            # 위치 기준으로 정렬 (역순)
            all_matches.sort(key=lambda x: x['match'].start(), reverse=True)
            
            # 역순으로 치환 (뒤에서부터 치환하여 인덱스 문제 방지)
            for match_info in all_matches:
                match = match_info['match']
                number_text = match_info['text']
                num_type = match_info['type']
                
                # 연도나 분기는 마커로 변환하지 않음
                if num_type in ['year', 'quarter']:
                    continue
                
                # 컨텍스트 기반으로 적절한 헤더 찾기
                start_pos = max(0, match.start() - 100)
                end_pos = min(len(text), match.end() + 100)
                context = text[start_pos:end_pos]
                
                # 지역명이나 헤더 키워드 찾기
                header_key = self._find_header_from_context(context, headers_info, number_text, num_type)
                
                if header_key:
                    marker = "{" + sheet_name + ":" + header_key + "}"
                else:
                    # 동적 마커 사용 (템플릿 필러가 처리)
                    if num_type == 'percent':
                        # 퍼센트는 증감률로 추정
                        marker = "{" + sheet_name + ":전국_증감률}"
                    else:
                        # 일반 숫자는 동적 마커 사용
                        marker = "{" + sheet_name + ":value}"
                
                processed = processed[:match.start()] + marker + processed[match.end():]
            
            return processed
        except Exception as e:
            # 오류 발생 시 원본 텍스트 반환
            print(f"마커 생성 오류: {e}")
            return text
    
    def _find_header_from_context(self, context: str, headers_info: Dict, number_text: str, num_type: str = 'number') -> Optional[str]:
        """
        컨텍스트에서 헤더를 찾습니다.
        동적 마커 패턴을 우선 사용합니다.
        """
        # 지역명 패턴 찾기
        region_pattern = r'([가-힣]+(?:시|도|군|구|시|읍|면))'
        region_match = re.search(region_pattern, context)
        
        # 동적 마커 패턴 우선 사용
        # 전국 관련
        if '전국' in context:
            if num_type == 'percent' or '%' in number_text:
                return '전국_증감률'
            else:
                return '전국_이름'  # 기본값
        
        # 지역별 증감률
        if region_match:
            region_name = region_match.group(1)
            normalized_region = self._normalize_name(region_name)
            
            if num_type == 'percent' or '%' in number_text:
                # 지역별 증감률 마커
                return f'{normalized_region}_증감률'
            else:
                # 지역별 이름 마커 (기본값)
                return f'{normalized_region}_이름'
        
        # 상위/하위 시도 패턴
        if '상위' in context or '하위' in context:
            if '상위' in context:
                # 상위시도1_증감률, 상위시도2_증감률 등
                idx_match = re.search(r'상위시도(\d+)', context)
                if idx_match:
                    idx = idx_match.group(1)
                    if num_type == 'percent' or '%' in number_text:
                        return f'상위시도{idx}_증감률'
                    else:
                        return f'상위시도{idx}_이름'
            if '하위' in context:
                idx_match = re.search(r'하위시도(\d+)', context)
                if idx_match:
                    idx = idx_match.group(1)
                    if num_type == 'percent' or '%' in number_text:
                        return f'하위시도{idx}_증감률'
                    else:
                        return f'하위시도{idx}_이름'
        
        # 업태/산업 패턴
        if '업태' in context or '산업' in context:
            # 전국_업태1_증감률, 상위시도1_업태1_증감률 등
            if '전국' in context:
                industry_match = re.search(r'(업태|산업)(\d+)', context)
                if industry_match:
                    idx = industry_match.group(2)
                    if num_type == 'percent' or '%' in number_text:
                        return f'전국_업태{idx}_증감률'
                    else:
                        return f'전국_업태{idx}_이름'
            elif region_match:
                normalized_region = self._normalize_name(region_match.group(1))
                industry_match = re.search(r'(업태|산업)(\d+)', context)
                if industry_match:
                    idx = industry_match.group(2)
                    if num_type == 'percent' or '%' in number_text:
                        return f'{normalized_region}_업태{idx}_증감률'
                    else:
                        return f'{normalized_region}_업태{idx}_이름'
        
        # 헤더 정보에서 직접 찾기
        if region_match:
            region_name = region_match.group(1)
            normalized_region = self._normalize_name(region_name)
            
            # 행 헤더에서 찾기
            if normalized_region in headers_info.get('row_map', {}):
                # 열 헤더도 찾기
                if num_type == 'percent' or '%' in number_text:
                    # 퍼센트 관련 헤더 찾기
                    for col in headers_info.get('columns', []):
                        if '증감' in col['header'] or '증감률' in col['header'] or '%' in col['header']:
                            return f"{normalized_region}_{col['normalized']}"
                else:
                    # 일반 숫자 헤더 찾기
                    for col in headers_info.get('columns', []):
                        if col['normalized'] not in ['시도', '시·도']:
                            return f"{normalized_region}_{col['normalized']}"
        
        # 헤더 키워드로 찾기
        for col in headers_info.get('columns', []):
            if col['header'] in context or any(keyword in context for keyword in ['증감', '증감률', '생산', '수주']):
                return col['normalized']
        
        return None
    
    def _normalize_name(self, name: str) -> str:
        """이름을 정규화합니다."""
        if not name:
            return ""
        
        # 공백, 특수문자 제거 및 언더스코어로 대체
        normalized = re.sub(r'[^\w가-힣]', '_', str(name))
        # 연속된 언더스코어를 하나로
        normalized = re.sub(r'_+', '_', normalized)
        # 앞뒤 언더스코어 제거
        normalized = normalized.strip('_').lower()
        
        return normalized
    
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
    <div class="content-text">데이터 마커 형식: {{시트명:셀주소}} 또는 {{시트명:헤더기반키}}</div>
</body>
</html>"""
    
    def close(self):
        """리소스 정리"""
        if self.header_parser:
            self.header_parser.close()
        if self.excel_extractor:
            self.excel_extractor.close()

