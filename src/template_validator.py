"""
템플릿 검증 모듈
생성된 템플릿과 원본 이미지를 비교하여 일치 여부 확인
"""

import cv2
import numpy as np
from pathlib import Path
from typing import Dict, Any, Tuple, Optional, List
from PIL import Image
import io
import base64
import tempfile
import shutil

try:
    from html2image import Html2Image
    HTML2IMAGE_AVAILABLE = True
except ImportError:
    HTML2IMAGE_AVAILABLE = False


class TemplateValidator:
    """생성된 템플릿과 원본 이미지를 비교하는 클래스"""
    
    def __init__(self):
        """템플릿 검증기 초기화"""
        if HTML2IMAGE_AVAILABLE:
            try:
                self.html2image = Html2Image()
            except Exception:
                self.html2image = None
        else:
            self.html2image = None
    
    def validate_template(
        self,
        template_html: str,
        reference_image_path: str,
        threshold: float = 0.85
    ) -> Dict[str, Any]:
        """
        생성된 템플릿과 원본 이미지를 비교합니다.
        
        Args:
            template_html: 생성된 HTML 템플릿
            reference_image_path: 정답 이미지 경로
            threshold: 일치 기준 (0-1)
            
        Returns:
            검증 결과 딕셔너리:
            - 'is_match': 일치 여부
            - 'similarity_score': 유사도 점수 (0-1)
            - 'differences': 차이점 리스트
            - 'rendered_image_path': 렌더링된 이미지 경로
        """
        # HTML을 이미지로 렌더링
        rendered_image_path = self._render_html_to_image(template_html)
        
        if not rendered_image_path:
            return {
                'is_match': False,
                'similarity_score': 0.0,
                'differences': ['HTML 렌더링 실패'],
                'rendered_image_path': None
            }
        
        # 이미지 비교
        similarity_score = self._compare_images(
            reference_image_path,
            rendered_image_path
        )
        
        # 차이점 분석
        differences = self._analyze_differences(
            reference_image_path,
            rendered_image_path
        )
        
        is_match = similarity_score >= threshold
        
        return {
            'is_match': is_match,
            'similarity_score': similarity_score,
            'differences': differences,
            'rendered_image_path': str(rendered_image_path),
            'threshold': threshold
        }
    
    def _render_html_to_image(self, html_content: str) -> Optional[Path]:
        """HTML을 이미지로 렌더링"""
        try:
            # 임시 HTML 파일 생성
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
                f.write(html_content)
                temp_html = f.name
            
            # 출력 디렉토리 생성
            output_dir = tempfile.mkdtemp()
            
            # 이미지로 렌더링
            if self.html2image:
                try:
                    self.html2image.screenshot(
                        html_file=temp_html,
                        save_as='rendered.png',
                        size=(1000, 1400)  # 기본 크기
                    )
                    
                    # 렌더링된 이미지 찾기 (html2image는 현재 디렉토리에 저장)
                    rendered_path = Path('rendered.png')
                    if rendered_path.exists():
                        # 출력 디렉토리로 이동
                        final_path = Path(output_dir) / 'rendered.png'
                        shutil.move(str(rendered_path), str(final_path))
                        return final_path
                except Exception as e:
                    # html2image 실패 시 대체 방법 시도
                    print(f"html2image 렌더링 실패, 대체 방법 시도: {e}")
                    return self._render_html_alternative(html_content, output_dir)
            else:
                # html2image가 없으면 대체 방법 사용
                return self._render_html_alternative(html_content, output_dir)
            
            return None
        except Exception as e:
            print(f"HTML 렌더링 오류: {e}")
            return None
    
    def _render_html_alternative(self, html_content: str, output_dir: str) -> Optional[Path]:
        """대체 HTML 렌더링 방법 (간단한 텍스트 기반)"""
        try:
            # HTML에서 텍스트만 추출하여 간단한 이미지 생성
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(html_content, 'html.parser')
            text = soup.get_text()
            
            # PIL로 텍스트 이미지 생성
            from PIL import Image, ImageDraw, ImageFont
            
            # 이미지 생성
            img = Image.new('RGB', (1000, 1400), color='white')
            draw = ImageDraw.Draw(img)
            
            # 텍스트 그리기 (간단한 버전)
            try:
                font = ImageFont.truetype("/System/Library/Fonts/AppleGothic.ttf", 20)
            except:
                font = ImageFont.load_default()
            
            y = 20
            for line in text.split('\n')[:50]:  # 최대 50줄
                if line.strip():
                    draw.text((20, y), line[:80], fill='black', font=font)
                    y += 30
                    if y > 1350:
                        break
            
            # 저장
            output_path = Path(output_dir) / 'rendered.png'
            img.save(str(output_path))
            return output_path
        except Exception as e:
            print(f"대체 렌더링 오류: {e}")
            return None
    
    def _compare_images(
        self,
        image1_path: str,
        image2_path: str
    ) -> float:
        """
        두 이미지를 비교하여 유사도 점수를 반환합니다.
        
        Args:
            image1_path: 첫 번째 이미지 경로
            image2_path: 두 번째 이미지 경로
            
        Returns:
            유사도 점수 (0-1)
        """
        try:
            # 이미지 로드
            img1 = cv2.imread(str(image1_path))
            img2 = cv2.imread(str(image2_path))
            
            if img1 is None or img2 is None:
                return 0.0
            
            # 크기 맞추기
            h1, w1 = img1.shape[:2]
            h2, w2 = img2.shape[:2]
            
            if h1 != h2 or w1 != w2:
                img2 = cv2.resize(img2, (w1, h1))
            
            # 구조적 유사도 (SSIM) 계산
            ssim_score = self._calculate_ssim(img1, img2)
            
            # 픽셀 유사도 계산
            pixel_similarity = self._calculate_pixel_similarity(img1, img2)
            
            # 종합 점수 (가중 평균)
            similarity = (ssim_score * 0.7) + (pixel_similarity * 0.3)
            
            return float(similarity)
        except Exception as e:
            print(f"이미지 비교 오류: {e}")
            return 0.0
    
    def _calculate_ssim(self, img1: np.ndarray, img2: np.ndarray) -> float:
        """구조적 유사도 (SSIM) 계산"""
        try:
            from skimage.metrics import structural_similarity as ssim
            
            # 그레이스케일로 변환
            gray1 = cv2.cvtColor(img1, cv2.COLOR_BGR2GRAY)
            gray2 = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
            
            # SSIM 계산
            score = ssim(gray1, gray2)
            return max(0.0, min(1.0, score))
        except ImportError:
            # scikit-image가 없으면 간단한 유사도 계산
            return self._simple_similarity(img1, img2)
        except Exception:
            return 0.0
    
    def _calculate_pixel_similarity(self, img1: np.ndarray, img2: np.ndarray) -> float:
        """픽셀 유사도 계산"""
        try:
            # 이미지를 float로 변환
            img1_float = img1.astype(np.float32)
            img2_float = img2.astype(np.float32)
            
            # 차이 계산
            diff = np.abs(img1_float - img2_float)
            
            # 평균 차이
            mean_diff = np.mean(diff) / 255.0
            
            # 유사도 (차이가 적을수록 높음)
            similarity = 1.0 - mean_diff
            
            return max(0.0, min(1.0, similarity))
        except Exception:
            return 0.0
    
    def _simple_similarity(self, img1: np.ndarray, img2: np.ndarray) -> float:
        """간단한 유사도 계산 (SSIM 대체)"""
        try:
            # 히스토그램 비교
            hist1 = cv2.calcHist([img1], [0, 1, 2], None, [50, 50, 50], [0, 256, 0, 256, 0, 256])
            hist2 = cv2.calcHist([img2], [0, 1, 2], None, [50, 50, 50], [0, 256, 0, 256, 0, 256])
            
            # 상관관계 계산
            correlation = cv2.compareHist(hist1, hist2, cv2.HISTCMP_CORREL)
            
            return max(0.0, min(1.0, correlation))
        except Exception:
            return 0.0
    
    def _analyze_differences(
        self,
        reference_path: str,
        rendered_path: str
    ) -> List[str]:
        """이미지 차이점 분석"""
        differences = []
        
        try:
            ref_img = cv2.imread(str(reference_path))
            ren_img = cv2.imread(str(rendered_path))
            
            if ref_img is None or ren_img is None:
                differences.append("이미지를 로드할 수 없습니다.")
                return differences
            
            # 크기 차이
            h1, w1 = ref_img.shape[:2]
            h2, w2 = ren_img.shape[:2]
            if h1 != h2 or w1 != w2:
                differences.append(f"크기 차이: 참조({w1}x{h1}) vs 렌더링({w2}x{h2})")
            
            # 밝기 차이
            ref_brightness = np.mean(ref_img)
            ren_brightness = np.mean(ren_img)
            if abs(ref_brightness - ren_brightness) > 20:
                differences.append(f"밝기 차이: {abs(ref_brightness - ren_brightness):.1f}")
            
            # 색상 분포 차이
            ref_hist = cv2.calcHist([ref_img], [0, 1, 2], None, [8, 8, 8], [0, 256, 0, 256, 0, 256])
            ren_hist = cv2.calcHist([ren_img], [0, 1, 2], None, [8, 8, 8], [0, 256, 0, 256, 0, 256])
            hist_diff = cv2.compareHist(ref_hist, ren_hist, cv2.HISTCMP_BHATTACHARYYA)
            if hist_diff > 0.3:
                differences.append(f"색상 분포 차이: {hist_diff:.2f}")
            
        except Exception as e:
            differences.append(f"차이점 분석 오류: {str(e)}")
        
        return differences

