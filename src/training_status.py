"""
학습 진행 상태 관리 모듈
"""

from typing import Dict, Any, Optional
import threading
from datetime import datetime
import uuid


class TrainingStatus:
    """학습 진행 상태를 관리하는 클래스"""
    
    def __init__(self):
        """상태 관리자 초기화"""
        self._statuses: Dict[str, Dict[str, Any]] = {}
        self._lock = threading.Lock()
    
    def create_status(self, template_name: str) -> str:
        """새로운 학습 상태 생성"""
        status_id = str(uuid.uuid4())
        
        with self._lock:
            self._statuses[status_id] = {
                'status_id': status_id,
                'template_name': template_name,
                'status': 'running',  # running, completed, stopped, error
                'current_step': 0,
                'total_steps': 0,
                'current_iteration': 0,
                'max_iterations': 0,
                'similarity_score': 0.0,
                'message': '',
                'progress_percentage': 0,
                'improvements': [],
                'created_at': datetime.now().isoformat(),
                'updated_at': datetime.now().isoformat(),
                'stop_requested': False,
                'result': None
            }
        
        return status_id
    
    def update_status(
        self,
        status_id: str,
        status: Optional[str] = None,
        current_step: Optional[int] = None,
        total_steps: Optional[int] = None,
        current_iteration: Optional[int] = None,
        max_iterations: Optional[int] = None,
        similarity_score: Optional[float] = None,
        message: Optional[str] = None,
        progress_percentage: Optional[float] = None,
        improvement: Optional[Dict[str, Any]] = None,
        result: Optional[Dict[str, Any]] = None
    ):
        """상태 업데이트"""
        with self._lock:
            if status_id not in self._statuses:
                return
            
            status_data = self._statuses[status_id]
            
            if status:
                status_data['status'] = status
            if current_step is not None:
                status_data['current_step'] = current_step
            if total_steps is not None:
                status_data['total_steps'] = total_steps
            if current_iteration is not None:
                status_data['current_iteration'] = current_iteration
            if max_iterations is not None:
                status_data['max_iterations'] = max_iterations
            if similarity_score is not None:
                status_data['similarity_score'] = similarity_score
            if message is not None:
                status_data['message'] = message
            if progress_percentage is not None:
                status_data['progress_percentage'] = progress_percentage
            if improvement:
                status_data['improvements'].append(improvement)
            if result:
                status_data['result'] = result
            
            status_data['updated_at'] = datetime.now().isoformat()
    
    def request_stop(self, status_id: str):
        """중단 요청"""
        with self._lock:
            if status_id in self._statuses:
                self._statuses[status_id]['stop_requested'] = True
                self._statuses[status_id]['message'] = '중단 요청됨...'
    
    def is_stop_requested(self, status_id: str) -> bool:
        """중단 요청 확인"""
        with self._lock:
            if status_id in self._statuses:
                return self._statuses[status_id].get('stop_requested', False)
            return False
    
    def get_status(self, status_id: str) -> Optional[Dict[str, Any]]:
        """상태 가져오기"""
        with self._lock:
            return self._statuses.get(status_id, {}).copy()
    
    def remove_status(self, status_id: str):
        """상태 제거"""
        with self._lock:
            if status_id in self._statuses:
                del self._statuses[status_id]


# 전역 상태 관리자 인스턴스
training_status_manager = TrainingStatus()

