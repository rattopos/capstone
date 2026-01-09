"""
매핑 관리 모듈
이미지-엑셀 매핑 설정 파일 관리
"""

import json
from pathlib import Path
from typing import Dict, List, Optional, Any


class MappingManager:
    """이미지-엑셀 매핑 설정을 관리하는 클래스"""
    
    def __init__(self, mappings_dir: str = 'image_mappings'):
        """
        매핑 관리자 초기화
        
        Args:
            mappings_dir: 매핑 설정 파일이 저장될 디렉토리
        """
        self.mappings_dir = Path(mappings_dir)
        self.mappings_dir.mkdir(parents=True, exist_ok=True)
    
    def save_mapping(self, image_name: str, mapping_data: Dict[str, Any]) -> Path:
        """
        매핑 설정을 파일로 저장합니다.
        
        Args:
            image_name: 이미지 파일명 (확장자 제외 또는 포함)
            mapping_data: 매핑 데이터 딕셔너리
            
        Returns:
            저장된 파일 경로
        """
        # 이미지 이름에서 확장자 제거
        image_base = Path(image_name).stem
        
        # JSON 파일 경로
        mapping_file = self.mappings_dir / f"{image_base}.json"
        
        # 매핑 데이터에 이미지 이름 추가
        mapping_data['image_name'] = image_name
        
        # JSON 파일로 저장
        with open(mapping_file, 'w', encoding='utf-8') as f:
            json.dump(mapping_data, f, ensure_ascii=False, indent=2)
        
        return mapping_file
    
    def load_mapping(self, image_name: str) -> Optional[Dict[str, Any]]:
        """
        매핑 설정을 파일에서 로드합니다.
        
        Args:
            image_name: 이미지 파일명
            
        Returns:
            매핑 데이터 딕셔너리 또는 None (파일이 없을 경우)
        """
        image_base = Path(image_name).stem
        mapping_file = self.mappings_dir / f"{image_base}.json"
        
        if not mapping_file.exists():
            return None
        
        try:
            with open(mapping_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"매핑 파일 로드 실패: {e}")
            return None
    
    def mapping_exists(self, image_name: str) -> bool:
        """
        매핑 설정 파일이 존재하는지 확인합니다.
        
        Args:
            image_name: 이미지 파일명
            
        Returns:
            파일 존재 여부
        """
        image_base = Path(image_name).stem
        mapping_file = self.mappings_dir / f"{image_base}.json"
        return mapping_file.exists()
    
    def list_mappings(self) -> List[str]:
        """
        저장된 모든 매핑 파일 목록을 반환합니다.
        
        Returns:
            이미지 파일명 리스트
        """
        mappings = []
        for mapping_file in self.mappings_dir.glob('*.json'):
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    image_name = data.get('image_name', mapping_file.stem)
                    mappings.append(image_name)
            except Exception:
                continue
        
        return mappings
    
    def create_mapping_from_analysis(
        self,
        image_name: str,
        image_analysis: Dict[str, Any],
        excel_structure: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        이미지 분석 결과로부터 매핑 설정을 생성합니다.
        
        Args:
            image_name: 이미지 파일명
            image_analysis: 이미지 분석 결과
            excel_structure: 엑셀 구조 정보 (선택적)
            
        Returns:
            매핑 데이터 딕셔너리
        """
        mapping_data = {
            'image_name': image_name,
            'template_fields': [],
            'detected_structure': {
                'text_regions': image_analysis.get('text_regions', []),
                'tables': image_analysis.get('tables', []),
                'layout': image_analysis.get('layout', {})
            }
        }
        
        # 필드 정보 추가
        if 'fields' in image_analysis:
            for field in image_analysis['fields']:
                field_mapping = {
                    'field_id': field.get('field_id', ''),
                    'text_in_image': field.get('text_in_image', ''),
                    'type': field.get('type', 'text'),
                    'excel_mapping': {
                        'sheet_name': None,
                        'cell_or_range': None,
                        'column_pattern': None,
                        'row_pattern': None
                    }
                }
                mapping_data['template_fields'].append(field_mapping)
        
        # 마커 정보 추가
        if 'markers' in image_analysis:
            mapping_data['markers'] = image_analysis['markers']
        
        return mapping_data
    
    def update_field_mapping(
        self,
        image_name: str,
        field_id: str,
        excel_mapping: Dict[str, Any]
    ) -> bool:
        """
        특정 필드의 엑셀 매핑을 업데이트합니다.
        
        Args:
            image_name: 이미지 파일명
            field_id: 필드 ID
            excel_mapping: 엑셀 매핑 정보
            
        Returns:
            업데이트 성공 여부
        """
        mapping_data = self.load_mapping(image_name)
        if not mapping_data:
            return False
        
        # 필드 찾기 및 업데이트
        for field in mapping_data.get('template_fields', []):
            if field.get('field_id') == field_id:
                field['excel_mapping'] = excel_mapping
                self.save_mapping(image_name, mapping_data)
                return True
        
        return False
    
    def get_field_mapping(self, image_name: str, field_id: str) -> Optional[Dict[str, Any]]:
        """
        특정 필드의 엑셀 매핑을 가져옵니다.
        
        Args:
            image_name: 이미지 파일명
            field_id: 필드 ID
            
        Returns:
            엑셀 매핑 정보 또는 None
        """
        mapping_data = self.load_mapping(image_name)
        if not mapping_data:
            return None
        
        for field in mapping_data.get('template_fields', []):
            if field.get('field_id') == field_id:
                return field.get('excel_mapping')
        
        return None

