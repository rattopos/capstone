"""
이미지 캐시 저장소 모듈
정답 이미지를 저장하고 관리
"""

import shutil
from pathlib import Path
from typing import Optional, Dict, Any
from datetime import datetime
import json


class ImageCache:
    """정답 이미지를 저장하고 관리하는 클래스"""
    
    def __init__(self, cache_dir: str = 'image_cache'):
        """
        이미지 캐시 초기화
        
        Args:
            cache_dir: 캐시 디렉토리 경로
        """
        self.cache_dir = Path(cache_dir)
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        
        # 메타데이터 파일
        self.metadata_file = self.cache_dir / 'cache_metadata.json'
    
    def save_reference_image(
        self,
        image_path: str,
        template_name: str,
        metadata: Optional[Dict[str, Any]] = None
    ) -> Path:
        """
        정답 이미지를 캐시에 저장합니다.
        
        Args:
            image_path: 원본 이미지 파일 경로
            template_name: 템플릿 이름
            metadata: 추가 메타데이터
            
        Returns:
            저장된 이미지 파일 경로
        """
        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"이미지 파일을 찾을 수 없습니다: {image_path}")
        
        # 캐시 파일명 생성
        cache_filename = f"{template_name}_reference.png"
        cache_path = self.cache_dir / cache_filename
        
        # 이미지 복사
        shutil.copy2(image_path, cache_path)
        
        # 메타데이터 저장
        self._update_metadata(template_name, {
            'reference_image': str(cache_path),
            'original_path': str(image_path),
            'saved_at': datetime.now().isoformat(),
            **(metadata or {})
        })
        
        return cache_path
    
    def get_reference_image(self, template_name: str) -> Optional[Path]:
        """
        정답 이미지를 가져옵니다.
        
        Args:
            template_name: 템플릿 이름
            
        Returns:
            정답 이미지 파일 경로 또는 None
        """
        metadata = self._get_metadata(template_name)
        if metadata and 'reference_image' in metadata:
            ref_path = Path(metadata['reference_image'])
            if ref_path.exists():
                return ref_path
        
        # 파일명으로 직접 찾기
        cache_path = self.cache_dir / f"{template_name}_reference.png"
        if cache_path.exists():
            return cache_path
        
        return None
    
    def has_reference_image(self, template_name: str) -> bool:
        """정답 이미지가 있는지 확인"""
        return self.get_reference_image(template_name) is not None
    
    def _update_metadata(self, template_name: str, metadata: Dict[str, Any]):
        """메타데이터 업데이트"""
        if self.metadata_file.exists():
            try:
                with open(self.metadata_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except Exception:
                data = {}
        else:
            data = {}
        
        if 'templates' not in data:
            data['templates'] = {}
        
        data['templates'][template_name] = {
            'template_name': template_name,
            **metadata
        }
        
        with open(self.metadata_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    
    def _get_metadata(self, template_name: str) -> Optional[Dict[str, Any]]:
        """메타데이터 가져오기"""
        if not self.metadata_file.exists():
            return None
        
        try:
            with open(self.metadata_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('templates', {}).get(template_name)
        except Exception:
            return None

