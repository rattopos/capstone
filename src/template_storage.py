"""
템플릿 저장 모듈
생성된 템플릿과 매핑 정보를 저장하고 관리
"""

import json
from pathlib import Path
from typing import Dict, List, Optional, Any
from datetime import datetime


class TemplateStorage:
    """템플릿과 매핑 정보를 저장하고 관리하는 클래스"""
    
    def __init__(self, storage_dir: str = 'templates_generated'):
        """
        템플릿 저장소 초기화
        
        Args:
            storage_dir: 템플릿이 저장될 디렉토리
        """
        self.storage_dir = Path(storage_dir)
        self.storage_dir.mkdir(parents=True, exist_ok=True)
        
        # 템플릿 메타데이터 파일
        self.metadata_file = self.storage_dir / 'templates_metadata.json'
    
    def save_template(
        self,
        template_name: str,
        template_html: str,
        mapping_data: Dict[str, Any],
        metadata: Optional[Dict[str, Any]] = None
    ) -> Path:
        """
        템플릿과 매핑 정보를 저장합니다.
        
        Args:
            template_name: 템플릿 이름
            template_html: HTML 템플릿 내용
            mapping_data: 매핑 데이터
            metadata: 추가 메타데이터
            
        Returns:
            저장된 템플릿 파일 경로
        """
        # 템플릿 파일 저장
        template_file = self.storage_dir / f"{template_name}.html"
        with open(template_file, 'w', encoding='utf-8') as f:
            f.write(template_html)
        
        # 매핑 파일 저장
        mapping_file = self.storage_dir / f"{template_name}_mapping.json"
        with open(mapping_file, 'w', encoding='utf-8') as f:
            json.dump(mapping_data, f, ensure_ascii=False, indent=2)
        
        # 메타데이터 업데이트
        self._update_metadata(template_name, metadata or {})
        
        return template_file
    
    def load_template(self, template_name: str) -> Optional[Dict[str, Any]]:
        """
        저장된 템플릿을 로드합니다.
        
        Args:
            template_name: 템플릿 이름
            
        Returns:
            템플릿 정보 딕셔너리 또는 None:
            - 'template_html': HTML 템플릿
            - 'mapping_data': 매핑 데이터
            - 'metadata': 메타데이터
        """
        template_file = self.storage_dir / f"{template_name}.html"
        mapping_file = self.storage_dir / f"{template_name}_mapping.json"
        
        if not template_file.exists() or not mapping_file.exists():
            return None
        
        try:
            # 템플릿 HTML 로드
            with open(template_file, 'r', encoding='utf-8') as f:
                template_html = f.read()
            
            # 매핑 데이터 로드
            with open(mapping_file, 'r', encoding='utf-8') as f:
                mapping_data = json.load(f)
            
            # 메타데이터 로드
            metadata = self._get_metadata(template_name)
            
            return {
                'template_html': template_html,
                'mapping_data': mapping_data,
                'metadata': metadata,
                'template_name': template_name
            }
        except Exception as e:
            print(f"템플릿 로드 실패: {e}")
            return None
    
    def list_templates(self) -> List[Dict[str, Any]]:
        """
        저장된 모든 템플릿 목록을 반환합니다.
        
        Returns:
            템플릿 메타데이터 리스트
        """
        templates = []
        
        if self.metadata_file.exists():
            try:
                with open(self.metadata_file, 'r', encoding='utf-8') as f:
                    metadata = json.load(f)
                    templates = list(metadata.get('templates', {}).values())
            except Exception:
                pass
        
        # 파일 시스템에서도 확인
        for template_file in self.storage_dir.glob('*.html'):
            template_name = template_file.stem
            if not any(t['name'] == template_name for t in templates):
                templates.append({
                    'name': template_name,
                    'created_at': datetime.fromtimestamp(template_file.stat().st_mtime).isoformat(),
                    'file_path': str(template_file)
                })
        
        return sorted(templates, key=lambda x: x.get('created_at', ''), reverse=True)
    
    def delete_template(self, template_name: str) -> bool:
        """
        템플릿을 삭제합니다.
        
        Args:
            template_name: 템플릿 이름
            
        Returns:
            삭제 성공 여부
        """
        template_file = self.storage_dir / f"{template_name}.html"
        mapping_file = self.storage_dir / f"{template_name}_mapping.json"
        
        try:
            if template_file.exists():
                template_file.unlink()
            if mapping_file.exists():
                mapping_file.unlink()
            
            # 메타데이터에서 제거
            self._remove_from_metadata(template_name)
            
            return True
        except Exception as e:
            print(f"템플릿 삭제 실패: {e}")
            return False
    
    def _update_metadata(self, template_name: str, metadata: Dict[str, Any]):
        """메타데이터 파일 업데이트"""
        if self.metadata_file.exists():
            try:
                with open(self.metadata_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except Exception:
                data = {'templates': {}}
        else:
            data = {'templates': {}}
        
        # 템플릿 메타데이터 추가/업데이트
        data['templates'][template_name] = {
            'name': template_name,
            'created_at': datetime.now().isoformat(),
            'updated_at': datetime.now().isoformat(),
            **metadata
        }
        
        with open(self.metadata_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    
    def _get_metadata(self, template_name: str) -> Dict[str, Any]:
        """템플릿 메타데이터 가져오기"""
        if not self.metadata_file.exists():
            return {}
        
        try:
            with open(self.metadata_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return data.get('templates', {}).get(template_name, {})
        except Exception:
            return {}
    
    def _remove_from_metadata(self, template_name: str):
        """메타데이터에서 템플릿 제거"""
        if not self.metadata_file.exists():
            return
        
        try:
            with open(self.metadata_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if template_name in data.get('templates', {}):
                del data['templates'][template_name]
            
            with open(self.metadata_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

