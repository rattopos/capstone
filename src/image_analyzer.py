"""
이미지 분석 모듈
OCR을 사용하여 이미지에서 텍스트 추출 및 레이아웃 분석
"""

import re
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Any
import cv2
import numpy as np
from PIL import Image
import easyocr
import pytesseract


def get_available_device():
    """
    사용 가능한 디바이스를 감지합니다.
    
    Returns:
        'cuda', 'mps', 'cpu' 중 하나
    """
    try:
        import torch
        
        # CUDA 사용 가능 여부 확인
        if torch.cuda.is_available():
            return 'cuda'
        
        # MPS 사용 가능 여부 확인 (macOS)
        if hasattr(torch.backends, 'mps') and torch.backends.mps.is_available():
            return 'mps'
        
        # CPU 사용
        return 'cpu'
    except ImportError:
        # PyTorch가 설치되지 않은 경우 CPU 사용
        return 'cpu'


def setup_pytorch_device(device: str):
    """
    PyTorch의 기본 디바이스를 설정합니다.
    EasyOCR이 내부적으로 PyTorch를 사용하므로, MPS를 사용할 수 있도록 환경을 설정합니다.
    
    Args:
        device: 디바이스 타입 ('cuda', 'mps', 'cpu')
    """
    try:
        import torch
        import os
        
        if device == 'mps' and hasattr(torch.backends, 'mps') and torch.backends.mps.is_available():
            # MPS 디바이스 설정
            # PyTorch 2.0+에서는 set_default_device 사용
            try:
                if hasattr(torch, 'set_default_device'):
                    torch.set_default_device('mps')
                    print(f"PyTorch 디바이스가 MPS로 설정되었습니다.")
                else:
                    # PyTorch 1.x의 경우 환경 변수 사용
                    os.environ['PYTORCH_ENABLE_MPS_FALLBACK'] = '1'
                    print(f"MPS 사용을 위한 환경 변수가 설정되었습니다.")
            except Exception as e:
                print(f"MPS 설정 중 오류 (계속 진행): {e}")
        elif device == 'cuda' and torch.cuda.is_available():
            try:
                if hasattr(torch, 'set_default_device'):
                    torch.set_default_device('cuda')
                    print(f"PyTorch 디바이스가 CUDA로 설정되었습니다.")
            except Exception:
                pass
    except ImportError:
        # PyTorch가 설치되지 않은 경우 무시
        pass
    except Exception as e:
        print(f"PyTorch 디바이스 설정 중 오류 (계속 진행): {e}")


def should_use_gpu(device: str) -> bool:
    """
    디바이스가 GPU인지 확인합니다.
    
    Args:
        device: 디바이스 타입 ('cuda', 'mps', 'cpu')
        
    Returns:
        GPU 사용 여부
    """
    return device in ['cuda', 'mps']


class ImageAnalyzer:
    """이미지에서 텍스트와 구조를 추출하는 클래스"""
    
    def __init__(self, use_easyocr: bool = True, device: Optional[str] = None):
        """
        이미지 분석기 초기화
        
        Args:
            use_easyocr: EasyOCR 사용 여부 (True면 EasyOCR, False면 Tesseract)
            device: 사용할 디바이스 ('cuda', 'mps', 'cpu', None이면 자동 감지)
        """
        self.use_easyocr = use_easyocr
        
        # 디바이스 자동 감지
        if device is None:
            detected_device = get_available_device()
            # EasyOCR은 MPS를 안정적으로 지원하지 않으므로, MPS 감지 시 CPU로 폴백
            if detected_device == 'mps':
                print(f"⚠️ MPS 디바이스가 감지되었지만, EasyOCR의 MPS 지원이 불안정하여 CPU를 사용합니다.")
                self.device = 'cpu'
            else:
                self.device = detected_device
        else:
            # 사용자가 명시적으로 MPS를 요청해도 EasyOCR 호환성을 위해 CPU로 폴백
            if device == 'mps':
                print(f"⚠️ MPS 디바이스가 요청되었지만, EasyOCR의 MPS 지원이 불안정하여 CPU를 사용합니다.")
                self.device = 'cpu'
            else:
                self.device = device
        
        # GPU 사용 여부 결정 (CUDA만 직접 지원)
        # EasyOCR의 gpu 파라미터는 CUDA만 직접 지원하므로, CUDA가 아니면 CPU 사용
        use_gpu = (self.device == 'cuda')
        
        if use_easyocr:
            try:
                # EasyOCR 초기화 (한글 + 영어 지원)
                # EasyOCR의 gpu 파라미터는 CUDA만 직접 지원
                # MPS는 지원하지 않으므로 CPU 사용
                self.reader = easyocr.Reader(['ko', 'en'], gpu=use_gpu)
                if use_gpu:
                    print(f"✓ EasyOCR 초기화 완료 (디바이스: CUDA)")
                else:
                    print(f"✓ EasyOCR 초기화 완료 (디바이스: CPU)")
            except Exception as e:
                print(f"✗ EasyOCR 초기화 실패, Tesseract 사용: {e}")
                self.use_easyocr = False
                self.reader = None
                self.device = 'cpu'
    
    def analyze_image(self, image_path: str) -> Dict[str, Any]:
        """
        이미지를 분석하여 텍스트와 구조를 추출합니다.
        
        Args:
            image_path: 이미지 파일 경로
            
        Returns:
            분석 결과 딕셔너리:
            - 'text_regions': 텍스트 영역 리스트
            - 'full_text': 전체 추출된 텍스트
            - 'tables': 테이블 구조 (있는 경우)
            - 'layout': 레이아웃 정보
        """
        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"이미지 파일을 찾을 수 없습니다: {image_path}")
        
        # 이미지 로드
        image = cv2.imread(str(image_path))
        if image is None:
            raise ValueError(f"이미지를 읽을 수 없습니다: {image_path}")
        
        # 텍스트 추출
        if self.use_easyocr and self.reader:
            text_regions = self._extract_text_easyocr(image)
        else:
            text_regions = self._extract_text_tesseract(image)
        
        # 전체 텍스트 추출
        full_text = ' '.join([region['text'] for region in text_regions])
        
        # 레이아웃 분석
        layout = self._analyze_layout(image, text_regions)
        
        # 테이블 구조 추출 (간단한 버전)
        tables = self._detect_tables(image, text_regions)
        
        return {
            'text_regions': text_regions,
            'full_text': full_text,
            'tables': tables,
            'layout': layout,
            'image_path': str(image_path)
        }
    
    def _extract_text_easyocr(self, image: np.ndarray) -> List[Dict[str, Any]]:
        """
        EasyOCR을 사용하여 텍스트 추출
        
        Args:
            image: OpenCV 이미지 배열
            
        Returns:
            텍스트 영역 리스트. 각 딕셔너리는 다음 키를 포함:
            - 'text': 추출된 텍스트
            - 'bbox': 바운딩 박스 좌표 [x1, y1, x2, y2]
            - 'confidence': 신뢰도 (0-1)
        """
        results = self.reader.readtext(image)
        text_regions = []
        
        for (bbox, text, confidence) in results:
            # bbox는 [[x1, y1], [x2, y2], [x3, y3], [x4, y4]] 형식
            # 간단하게 [x1, y1, x2, y2] 형식으로 변환
            x_coords = [point[0] for point in bbox]
            y_coords = [point[1] for point in bbox]
            x1, y1 = int(min(x_coords)), int(min(y_coords))
            x2, y2 = int(max(x_coords)), int(max(y_coords))
            
            text_regions.append({
                'text': text.strip(),
                'bbox': [x1, y1, x2, y2],
                'confidence': confidence
            })
        
        return text_regions
    
    def _extract_text_tesseract(self, image: np.ndarray) -> List[Dict[str, Any]]:
        """
        Tesseract OCR을 사용하여 텍스트 추출
        
        Args:
            image: OpenCV 이미지 배열
            
        Returns:
            텍스트 영역 리스트
        """
        # PIL 이미지로 변환
        pil_image = Image.fromarray(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
        
        # Tesseract로 텍스트 추출 (한글 지원)
        try:
            # 박스 정보와 함께 추출
            data = pytesseract.image_to_data(pil_image, lang='kor+eng', output_type=pytesseract.Output.DICT)
        except Exception:
            # 한글 모델이 없으면 영어만
            data = pytesseract.image_to_data(pil_image, lang='eng', output_type=pytesseract.Output.DICT)
        
        text_regions = []
        n_boxes = len(data['text'])
        
        for i in range(n_boxes):
            text = data['text'][i].strip()
            if text and int(data['conf'][i]) > 0:
                x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
                confidence = float(data['conf'][i]) / 100.0  # 0-100을 0-1로 변환
                
                text_regions.append({
                    'text': text,
                    'bbox': [x, y, x + w, y + h],
                    'confidence': confidence
                })
        
        return text_regions
    
    def _analyze_layout(self, image: np.ndarray, text_regions: List[Dict]) -> Dict[str, Any]:
        """
        이미지 레이아웃 분석
        
        Args:
            image: OpenCV 이미지 배열
            text_regions: 텍스트 영역 리스트
            
        Returns:
            레이아웃 정보 딕셔너리
        """
        if not text_regions:
            return {
                'width': image.shape[1],
                'height': image.shape[0],
                'regions': []
            }
        
        # 텍스트 영역을 Y 좌표 기준으로 정렬하여 행 구분
        sorted_regions = sorted(text_regions, key=lambda r: r['bbox'][1])
        
        # 간단한 행 그룹화 (Y 좌표가 비슷한 것들을 같은 행으로)
        rows = []
        current_row = [sorted_regions[0]]
        row_y_threshold = 20  # 픽셀 단위
        
        for region in sorted_regions[1:]:
            y_diff = abs(region['bbox'][1] - current_row[0]['bbox'][1])
            if y_diff < row_y_threshold:
                current_row.append(region)
            else:
                rows.append(current_row)
                current_row = [region]
        if current_row:
            rows.append(current_row)
        
        # 각 행을 X 좌표 기준으로 정렬
        for row in rows:
            row.sort(key=lambda r: r['bbox'][0])
        
        return {
            'width': image.shape[1],
            'height': image.shape[0],
            'rows': rows,
            'num_rows': len(rows)
        }
    
    def _detect_tables(self, image: np.ndarray, text_regions: List[Dict]) -> List[Dict]:
        """
        이미지에서 테이블 구조 감지 (간단한 버전)
        
        Args:
            image: OpenCV 이미지 배열
            text_regions: 텍스트 영역 리스트
            
        Returns:
            테이블 정보 리스트
        """
        # 간단한 구현: 정렬된 텍스트 영역을 기반으로 테이블 추정
        # 실제로는 더 정교한 테이블 감지 알고리즘이 필요할 수 있음
        
        if not text_regions:
            return []
        
        # 레이아웃 분석 결과 사용
        layout = self._analyze_layout(image, text_regions)
        
        tables = []
        # 여러 행이 있고 각 행에 여러 텍스트가 있으면 테이블로 간주
        if layout['num_rows'] > 2:
            # 첫 번째 행을 헤더로 간주
            if len(layout['rows']) > 0:
                header_row = layout['rows'][0]
                data_rows = layout['rows'][1:]
                
                if len(header_row) > 1 and len(data_rows) > 0:
                    tables.append({
                        'header': [r['text'] for r in header_row],
                        'rows': [[r['text'] for r in row] for row in data_rows],
                        'bbox': self._calculate_table_bbox(text_regions)
                    })
        
        return tables
    
    def _calculate_table_bbox(self, text_regions: List[Dict]) -> List[int]:
        """텍스트 영역들로부터 테이블 바운딩 박스 계산"""
        if not text_regions:
            return [0, 0, 0, 0]
        
        x_coords = []
        y_coords = []
        
        for region in text_regions:
            bbox = region['bbox']
            x_coords.extend([bbox[0], bbox[2]])
            y_coords.extend([bbox[1], bbox[3]])
        
        return [min(x_coords), min(y_coords), max(x_coords), max(y_coords)]
    
    def extract_template_structure(self, image_path: str) -> Dict[str, Any]:
        """
        이미지에서 템플릿 구조를 추출합니다.
        
        Args:
            image_path: 이미지 파일 경로
            
        Returns:
            템플릿 구조 딕셔너리:
            - 'fields': 데이터 필드 리스트
            - 'markers': 마커 후보 리스트
            - 'structure': 구조 정보
        """
        analysis = self.analyze_image(image_path)
        
        # 텍스트에서 데이터 필드 추출
        fields = self._extract_fields(analysis['text_regions'])
        
        # 마커 후보 생성
        markers = self._generate_marker_candidates(fields, analysis)
        
        return {
            'fields': fields,
            'markers': markers,
            'structure': analysis['layout'],
            'full_text': analysis['full_text']
        }
    
    def _extract_fields(self, text_regions: List[Dict]) -> List[Dict[str, Any]]:
        """
        텍스트 영역에서 데이터 필드 추출
        
        Args:
            text_regions: 텍스트 영역 리스트
            
        Returns:
            필드 리스트. 각 필드는 다음 키를 포함:
            - 'field_id': 필드 식별자
            - 'text_in_image': 이미지의 텍스트
            - 'type': 필드 타입 (예: 'number', 'text', 'percentage')
            - 'bbox': 바운딩 박스
        """
        fields = []
        
        # 숫자 패턴 (퍼센트 포함)
        number_pattern = re.compile(r'[-+]?\d*\.?\d+%?')
        # 지역명 패턴 (한글 지역명)
        region_pattern = re.compile(r'(전국|서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주)')
        
        for region in text_regions:
            text = region['text']
            
            # 숫자 필드 감지
            if number_pattern.search(text):
                field_type = 'percentage' if '%' in text else 'number'
                fields.append({
                    'field_id': f"field_{len(fields)}",
                    'text_in_image': text,
                    'type': field_type,
                    'bbox': region['bbox'],
                    'confidence': region.get('confidence', 0.0)
                })
            # 지역명 필드 감지
            elif region_pattern.search(text):
                fields.append({
                    'field_id': f"field_{len(fields)}",
                    'text_in_image': text,
                    'type': 'region',
                    'bbox': region['bbox'],
                    'confidence': region.get('confidence', 0.0)
                })
            # 기타 텍스트 필드
            elif len(text.strip()) > 1:
                fields.append({
                    'field_id': f"field_{len(fields)}",
                    'text_in_image': text,
                    'type': 'text',
                    'bbox': region['bbox'],
                    'confidence': region.get('confidence', 0.0)
                })
        
        return fields
    
    def _generate_marker_candidates(self, fields: List[Dict], analysis: Dict) -> List[Dict[str, Any]]:
        """
        필드로부터 마커 후보 생성
        
        Args:
            fields: 필드 리스트
            analysis: 이미지 분석 결과
            
        Returns:
            마커 후보 리스트
        """
        markers = []
        
        # 텍스트에서 키워드 추출하여 마커 생성
        keywords = ['전국', '증감률', '증가', '감소', '상위', '하위', '시도', '업태', '산업']
        
        for field in fields:
            text = field['text_in_image']
            
            # 키워드 기반 마커 생성
            for keyword in keywords:
                if keyword in text:
                    # 마커 ID 생성 (예: 전국_증감률, 상위시도1_이름 등)
                    marker_id = self._generate_marker_id(text, keyword)
                    if marker_id:
                        markers.append({
                            'marker_id': marker_id,
                            'field_id': field['field_id'],
                            'text_in_image': text,
                            'type': field['type']
                        })
        
        return markers
    
    def _generate_marker_id(self, text: str, keyword: str) -> Optional[str]:
        """텍스트에서 마커 ID 생성"""
        # 간단한 규칙 기반 생성
        if '전국' in text:
            if '증감률' in text or '%' in text:
                return '전국_증감률'
            elif '이름' in text or '명' in text:
                return '전국_이름'
        elif '상위' in text and '시도' in text:
            # 숫자 추출
            numbers = re.findall(r'\d+', text)
            idx = numbers[0] if numbers else '1'
            if '증감률' in text or '%' in text:
                return f'상위시도{idx}_증감률'
            else:
                return f'상위시도{idx}_이름'
        elif '하위' in text and '시도' in text:
            numbers = re.findall(r'\d+', text)
            idx = numbers[0] if numbers else '1'
            if '증감률' in text or '%' in text:
                return f'하위시도{idx}_증감률'
            else:
                return f'하위시도{idx}_이름'
        
        return None

