"""
템플릿 자동 학습 모듈
전체 자동화 루틴 관리: 이미지 저장 → 템플릿 생성 → 검증 → 개선 반복
"""

from pathlib import Path
from typing import Dict, Any, Optional, List
from .image_cache import ImageCache
from .template_generator import TemplateGenerator
from .template_validator import TemplateValidator
from .template_improver import TemplateImprover
from .mapping_trainer import MappingTrainer
from .template_storage import TemplateStorage
from .excel_extractor import ExcelExtractor
from .training_status import training_status_manager


class TemplateAutoTrainer:
    """템플릿 자동 학습 및 개선 시스템"""
    
    def __init__(
        self,
        image_cache: Optional[ImageCache] = None,
        template_generator: Optional[TemplateGenerator] = None,
        template_validator: Optional[TemplateValidator] = None,
        template_improver: Optional[TemplateImprover] = None
    ):
        """
        자동 학습기 초기화
        
        Args:
            image_cache: 이미지 캐시 인스턴스
            template_generator: 템플릿 생성기 인스턴스
            template_validator: 템플릿 검증기 인스턴스
            template_improver: 템플릿 개선기 인스턴스
        """
        self.image_cache = image_cache or ImageCache()
        self.template_generator = template_generator or TemplateGenerator()
        self.template_validator = template_validator or TemplateValidator()
        self.template_improver = template_improver or TemplateImprover(self.template_generator)
        self.training_history = []
    
    def train_template(
        self,
        reference_image_path: str,
        template_name: str,
        excel_file_path: Optional[str] = None,
        sheet_name: Optional[str] = None,
        year: int = 2025,
        quarter: int = 2,
        max_iterations: int = 10,
        similarity_threshold: float = 0.85,
        status_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        템플릿 자동 학습 루틴 실행
        
        Args:
            reference_image_path: 정답 이미지 경로
            template_name: 템플릿 이름
            excel_file_path: 엑셀 파일 경로 (매핑 학습용, 선택)
            sheet_name: 시트명 (선택)
            year: 연도
            quarter: 분기
            max_iterations: 최대 개선 반복 횟수
            similarity_threshold: 유사도 임계값
            
        Returns:
            학습 결과 딕셔너리:
            - 'template_html': 최종 템플릿 HTML
            - 'final_similarity': 최종 유사도 점수
            - 'iterations': 반복 횟수
            - 'improvements': 개선 내역
            - 'mapping_data': 매핑 데이터 (엑셀 파일이 제공된 경우)
        """
        # 상태 업데이트 초기화
        if status_id:
            training_status_manager.update_status(
                status_id,
                status='running',
                current_step=1,
                total_steps=5,
                max_iterations=max_iterations,
                message='정답 이미지를 캐시에 저장 중...',
                progress_percentage=0
            )
        
        # 1. 정답 이미지를 캐시에 저장
        print(f"[1/5] 정답 이미지를 캐시에 저장 중...")
        cache_path = self.image_cache.save_reference_image(
            reference_image_path,
            template_name,
            {'year': year, 'quarter': quarter}
        )
        print(f"✓ 이미지 저장 완료: {cache_path}")
        
        if status_id:
            if training_status_manager.is_stop_requested(status_id):
                training_status_manager.update_status(status_id, status='stopped', message='중단됨')
                return {'template_html': '', 'final_similarity': 0.0, 'iterations': 0, 'improvements': [], 'stopped': True}
            
            training_status_manager.update_status(
                status_id,
                current_step=2,
                message='템플릿 생성 중...',
                progress_percentage=20
            )
        
        # 2. 템플릿 생성
        print(f"[2/5] 템플릿 생성 중...")
        template_result = self.template_generator.generate_template_from_image(
            str(cache_path),
            template_name,
            sheet_name
        )
        template_html = template_result['template_html']
        print(f"✓ 템플릿 생성 완료 (마커 수: {len(template_result['markers'])})")
        
        if status_id:
            training_status_manager.update_status(
                status_id,
                current_step=3,
                message='템플릿 검증 및 개선 중...',
                progress_percentage=40
            )
        
        # 3-5. 검증 및 개선 반복
        print(f"[3-5/5] 템플릿 검증 및 개선 중...")
        iteration = 0
        improvements = []
        current_html = template_html
        best_similarity = 0.0
        best_html = template_html
        
        while iteration < max_iterations:
            # 중단 확인
            if status_id and training_status_manager.is_stop_requested(status_id):
                print(f"  중단 요청됨. 현재까지의 최고 템플릿 저장...")
                training_status_manager.update_status(
                    status_id,
                    status='stopped',
                    message='중단됨 - 최고 템플릿 저장 중...',
                    progress_percentage=90
                )
                break
            
            iteration += 1
            print(f"  반복 {iteration}/{max_iterations}...")
            
            if status_id:
                training_status_manager.update_status(
                    status_id,
                    current_iteration=iteration,
                    message=f'검증 중... (반복 {iteration}/{max_iterations})',
                    progress_percentage=40 + (iteration * 50 / max_iterations)
                )
            
            # 3. 검증
            validation_result = self.template_validator.validate_template(
                current_html,
                str(cache_path),
                threshold=similarity_threshold
            )
            
            similarity = validation_result['similarity_score']
            is_match = validation_result['is_match']
            
            print(f"    유사도: {similarity:.3f} {'✓' if is_match else '✗'}")
            
            # 최고 점수 업데이트 및 저장
            if similarity > best_similarity:
                best_similarity = similarity
                best_html = current_html
                
                # 매번 최고 템플릿 저장
                try:
                    template_storage = TemplateStorage()
                    mapping_data = None
                    if excel_file_path and excel_file_path is not None:
                        from pathlib import Path
                        excel_path_obj = Path(excel_file_path)
                        if excel_path_obj.exists():
                            excel_extractor = ExcelExtractor(str(excel_path_obj))
                            excel_extractor.load_workbook()
                            mapping_trainer = MappingTrainer(excel_extractor)
                            mapping_data = mapping_trainer.train_mapping(
                                best_html,
                                template_name,
                                template_result['sheet_name'],
                                year,
                                quarter
                            )
                            excel_extractor.close()
                    
                    template_storage.save_template(
                        template_name,
                        best_html,
                        mapping_data or {},
                        {
                            'sheet_name': template_result['sheet_name'],
                            'final_similarity': best_similarity,
                            'iterations': iteration,
                            'auto_trained': True,
                            'interrupted': status_id and training_status_manager.is_stop_requested(status_id) if status_id else False
                        }
                    )
                except Exception as e:
                    print(f"템플릿 저장 오류 (무시하고 계속): {e}")
            
            if status_id:
                training_status_manager.update_status(
                    status_id,
                    similarity_score=similarity,
                    message=f'유사도: {similarity:.3f} ({iteration}/{max_iterations})',
                    progress_percentage=40 + (iteration * 50 / max_iterations)
                )
            
            # 일치하면 종료
            if is_match:
                print(f"  ✓ 목표 유사도 달성! (유사도: {similarity:.3f})")
                if status_id:
                    training_status_manager.update_status(
                        status_id,
                        status='completed',
                        message=f'완료! 유사도: {similarity:.3f}',
                        progress_percentage=100
                    )
                break
            
            # 4-5. 개선
            if status_id:
                training_status_manager.update_status(
                    status_id,
                    message=f'개선 중... (반복 {iteration}/{max_iterations})'
                )
            
            improvement_result = self.template_improver.improve_template(
                current_html,
                validation_result,
                str(cache_path),
                max_iterations=1,  # 한 번에 하나씩 개선
                excel_file_path=excel_file_path,
                sheet_name=template_result.get('sheet_name'),
                year=year,
                quarter=quarter
            )
            
            if improvement_result['improvements']:
                improvements.extend(improvement_result['improvements'])
                current_html = improvement_result['improved_html']
                improvement_info = improvement_result['improvements'][0]
                print(f"    개선 적용: {improvement_info['description']}")
                
                if status_id:
                    training_status_manager.update_status(
                        status_id,
                        improvement={
                            'iteration': iteration,
                            'description': improvement_info['description']
                        }
                    )
            else:
                print(f"    더 이상 개선할 수 없습니다.")
                if status_id:
                    training_status_manager.update_status(
                        status_id,
                        status='completed',
                        message='더 이상 개선할 수 없음',
                        progress_percentage=100
                    )
                break
        
        # 최종 템플릿 저장 (중단된 경우에도 최고 템플릿 저장)
        if best_html:
            try:
                template_storage = TemplateStorage()
                mapping_data = None
                if excel_file_path and excel_file_path is not None:
                    from pathlib import Path
                    excel_path_obj = Path(excel_file_path)
                    if excel_path_obj.exists():
                        excel_extractor = ExcelExtractor(str(excel_path_obj))
                        excel_extractor.load_workbook()
                        mapping_trainer = MappingTrainer(excel_extractor)
                        mapping_data = mapping_trainer.train_mapping(
                            best_html,
                            template_name,
                            template_result['sheet_name'],
                            year,
                            quarter
                        )
                        excel_extractor.close()
                
                template_storage.save_template(
                    template_name,
                    best_html,
                    mapping_data or {},
                    {
                        'sheet_name': template_result['sheet_name'],
                        'final_similarity': best_similarity,
                        'iterations': iteration,
                        'auto_trained': True,
                        'interrupted': status_id and training_status_manager.is_stop_requested(status_id) if status_id else False
                    }
                )
                print(f"✓ 템플릿 '{template_name}' 저장 완료 (유사도: {best_similarity:.3f})")
            except Exception as e:
                import traceback
                traceback.print_exc()
                print(f"✗ 최종 템플릿 저장 실패: {e}")
        
        # 학습 기록 저장
        training_record = {
            'template_name': template_name,
            'iterations': iteration,
            'final_similarity': best_similarity,
            'improvements_count': len(improvements),
            'improvements': improvements,
            'stopped': status_id and training_status_manager.is_stop_requested(status_id) if status_id else False
        }
        self.training_history.append(training_record)
        
        result = {
            'template_html': best_html,
            'final_similarity': best_similarity,
            'iterations': iteration,
            'improvements': improvements,
            'mapping_data': mapping_data,
            'template_result': template_result,
            'training_record': training_record
        }
        
        if status_id:
            training_status_manager.update_status(
                status_id,
                status='completed' if not training_status_manager.is_stop_requested(status_id) else 'stopped',
                result=result,
                message=f'완료 (유사도: {best_similarity:.3f})' if not training_status_manager.is_stop_requested(status_id) else f'중단됨 (유사도: {best_similarity:.3f})',
                progress_percentage=100
            )
        
        return result
    
    def get_training_history(self, template_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """학습 기록 가져오기"""
        if template_name:
            return [
                record for record in self.training_history
                if record.get('template_name') == template_name
            ]
        return self.training_history

