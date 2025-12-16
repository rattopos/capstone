"""
템플릿 개선 모듈
검증 결과를 바탕으로 템플릿을 자동으로 개선
"""

import re
from typing import Dict, Any, List, Optional
from .template_generator import TemplateGenerator
from .image_analyzer import ImageAnalyzer


class TemplateImprover:
    """템플릿을 자동으로 개선하는 클래스"""
    
    def __init__(self, template_generator: Optional[TemplateGenerator] = None):
        """
        템플릿 개선기 초기화
        
        Args:
            template_generator: 템플릿 생성기 인스턴스
        """
        self.template_generator = template_generator or TemplateGenerator()
        self.improvement_history = []
        self.recent_strategies = []  # 최근 사용한 전략 추적 (같은 전략 반복 방지)
    
    def improve_template(
        self,
        template_html: str,
        validation_result: Dict[str, Any],
        image_path: str,
        max_iterations: int = 10,
        excel_file_path: Optional[str] = None,
        sheet_name: Optional[str] = None,
        year: Optional[int] = None,
        quarter: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        템플릿을 개선합니다.
        
        Args:
            template_html: 현재 템플릿 HTML
            validation_result: 검증 결과
            image_path: 원본 이미지 경로
            max_iterations: 최대 개선 반복 횟수
            
        Returns:
            개선 결과 딕셔너리:
            - 'improved_html': 개선된 HTML
            - 'improvements': 적용된 개선 사항 리스트
            - 'iterations': 반복 횟수
        """
        improvements = []
        current_html = template_html
        iteration = 0
        
        # 최근 전략 기록 초기화 (너무 길어지지 않도록 최근 5개만 유지)
        if len(self.recent_strategies) > 10:
            self.recent_strategies = self.recent_strategies[-5:]
        
        while iteration < max_iterations:
            iteration += 1
            
            # 개선 전략 선택 및 적용
            improvement_strategy = self._select_improvement_strategy(
                validation_result,
                current_html
            )
            
            if not improvement_strategy:
                break
            
            # 개선 적용
            improved_html = self._apply_improvement(
                current_html,
                improvement_strategy,
                image_path,
                excel_file_path=excel_file_path,
                sheet_name=sheet_name,
                year=year,
                quarter=quarter
            )
            
            if improved_html == current_html:
                # 더 이상 개선되지 않음
                break
            
            improvements.append({
                'iteration': iteration,
                'strategy': improvement_strategy['type'],
                'description': improvement_strategy.get('description', '')
            })
            
            current_html = improved_html
            
            # 개선 기록
            self.improvement_history.append({
                'iteration': iteration,
                'strategy': improvement_strategy,
                'html': current_html
            })
        
        return {
            'improved_html': current_html,
            'improvements': improvements,
            'iterations': iteration
        }
    
    def _select_improvement_strategy(
        self,
        validation_result: Dict[str, Any],
        template_html: str
    ) -> Optional[Dict[str, Any]]:
        """개선 전략 선택 (매핑 기반 개선 우선, 같은 전략 반복 방지)"""
        differences = validation_result.get('differences', [])
        similarity_score = validation_result.get('similarity_score', 0.0)
        
        # 최근 3번의 전략 확인 (같은 전략 반복 방지)
        recent_types = [s.get('type') for s in self.recent_strategies[-3:]]
        
        # 전략 후보 리스트 (우선순위 순)
        strategy_candidates = []
        
        # 1. 매핑 기반 개선 우선 (데이터 매핑 후 비교)
        if 'improve_with_mapping' not in recent_types:
            strategy_candidates.append({
                'type': 'improve_with_mapping',
                'description': '매핑 데이터 적용 후 비교',
                'priority': 10,  # 가장 높은 우선순위
                'similarity_score': similarity_score
            })
        
        # 2. 마커 개선 (유사도가 낮을 때)
        if similarity_score < 0.7 and 'improve_markers' not in recent_types:
            strategy_candidates.append({
                'type': 'improve_markers',
                'description': '마커 정확도 개선',
                'priority': 9,
                'similarity_score': similarity_score
            })
        
        # 3. 레이아웃 개선
        if similarity_score < 0.85 and 'improve_layout' not in recent_types:
            strategy_candidates.append({
                'type': 'improve_layout',
                'description': '레이아웃 개선',
                'priority': 8,
                'similarity_score': similarity_score
            })
        
        # 4. 차이점 기반 전략 (크기 조정 제외하고 우선)
        for diff in differences:
            if '밝기 차이' in diff and 'adjust_brightness' not in recent_types:
                strategy_candidates.append({
                    'type': 'adjust_brightness',
                    'description': '밝기 조정',
                    'priority': 7,
                    'difference': diff
                })
            elif '색상 분포 차이' in diff and 'adjust_colors' not in recent_types:
                strategy_candidates.append({
                    'type': 'adjust_colors',
                    'description': '색상 조정',
                    'priority': 6,
                    'difference': diff
                })
        
        # 5. 크기 조정 (최하위 우선순위, 최근에 사용하지 않은 경우에만)
        if 'adjust_size' not in recent_types:
            for diff in differences:
                if '크기 차이' in diff:
                    strategy_candidates.append({
                        'type': 'adjust_size',
                        'description': '템플릿 크기 조정',
                        'priority': 3,  # 낮은 우선순위
                        'difference': diff
                    })
                    break
        
        # 우선순위가 높은 전략 선택
        if strategy_candidates:
            selected = max(strategy_candidates, key=lambda x: x['priority'])
            # 최근 전략 기록에 추가
            self.recent_strategies.append(selected)
            return selected
        
        return None
    
    def _apply_improvement(
        self,
        template_html: str,
        strategy: Dict[str, Any],
        image_path: str,
        excel_file_path: Optional[str] = None,
        sheet_name: Optional[str] = None,
        year: Optional[int] = None,
        quarter: Optional[int] = None
    ) -> str:
        """개선 전략 적용"""
        strategy_type = strategy['type']
        
        if strategy_type == 'improve_with_mapping':
            return self._improve_with_mapping(
                template_html, 
                image_path,
                excel_file_path,
                sheet_name,
                year,
                quarter
            )
        elif strategy_type == 'adjust_size':
            return self._adjust_template_size(template_html, strategy)
        elif strategy_type == 'adjust_brightness':
            return self._adjust_brightness(template_html, strategy)
        elif strategy_type == 'adjust_colors':
            return self._adjust_colors(template_html, strategy)
        elif strategy_type == 'improve_markers':
            return self._improve_markers(template_html, image_path)
        elif strategy_type == 'improve_layout':
            return self._improve_layout(template_html, image_path)
        
        return template_html
    
    def _adjust_template_size(self, html: str, strategy: Dict[str, Any]) -> str:
        """템플릿 크기 조정"""
        # CSS에서 크기 관련 스타일 수정
        # body 태그의 max-width 조정
        if 'max-width' in html:
            # 기존 max-width 값 추출 및 조정
            pattern = r'max-width:\s*(\d+)px'
            match = re.search(pattern, html)
            if match:
                current_width = int(match.group(1))
                # 차이에 따라 조정 (간단한 휴리스틱)
                new_width = current_width + 50
                html = re.sub(pattern, f'max-width: {new_width}px', html)
        
        return html
    
    def _adjust_brightness(self, html: str, strategy: Dict[str, Any]) -> str:
        """밝기 조정"""
        # CSS에서 배경색 조정
        if 'background-color' in html:
            # 배경색을 약간 밝게 또는 어둡게 조정
            pattern = r'background-color:\s*([^;]+)'
            # 간단한 조정 (실제로는 더 정교한 로직 필요)
            html = re.sub(
                r'background-color:\s*#fff',
                'background-color: #fafafa',
                html
            )
        
        return html
    
    def _adjust_colors(self, html: str, strategy: Dict[str, Any]) -> str:
        """색상 조정"""
        # 텍스트 색상 조정
        # 실제 구현은 더 정교해야 함
        return html
    
    def _improve_markers(self, html: str, image_path: str) -> str:
        """마커 정확도 개선"""
        # 이미지를 다시 분석하여 마커 위치 개선
        try:
            analysis = self.template_generator.image_analyzer.analyze_image(image_path)
            template_structure = self.template_generator.image_analyzer.extract_template_structure(image_path)
            
            # 새로운 마커로 업데이트
            # 실제 구현은 더 복잡함
            return html
        except Exception:
            return html
    
    def _improve_layout(self, html: str, image_path: str) -> str:
        """레이아웃 개선"""
        # 레이아웃 분석 결과를 바탕으로 HTML 구조 개선
        try:
            analysis = self.template_generator.image_analyzer.analyze_image(image_path)
            layout = analysis.get('layout', {})
            
            # 행 수에 맞게 레이아웃 조정
            num_rows = layout.get('num_rows', 0)
            if num_rows > 0:
                # HTML 구조 개선
                # 실제 구현은 더 정교해야 함
                pass
            
            return html
        except Exception:
            return html
    
    def _improve_with_mapping(
        self,
        html: str,
        image_path: str,
        excel_file_path: Optional[str],
        sheet_name: Optional[str],
        year: Optional[int],
        quarter: Optional[int]
    ) -> str:
        """매핑 데이터 적용 후 이미지와 비교하여 개선"""
        # 이미지를 재분석하여 템플릿 재생성 (가장 효과적인 방법)
        try:
            # 이미지에서 시트명 추론
            from pathlib import Path
            image_name = Path(image_path).stem
            
            # 이미지 재분석을 통해 템플릿 재생성
            # 이렇게 하면 OCR 결과가 개선되어 더 정확한 템플릿이 생성될 수 있음
            improved_result = self.template_generator.generate_template_from_image(
                image_path,
                template_name=image_name,  # 임시 이름
                sheet_name=sheet_name
            )
            
            improved_html = improved_result['template_html']
            return improved_html
        except Exception as e:
            print(f"이미지 재분석 기반 개선 실패: {e}")
            return html

