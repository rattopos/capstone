"""
스키마 로더 모듈
JSON 스키마 파일을 로드하고 관리하는 기능 제공
"""

import json
from pathlib import Path
from typing import Dict, Any, Optional


class SchemaLoader:
    """스키마 파일을 로드하고 관리하는 클래스"""
    
    def __init__(self, schemas_dir: Optional[str] = None):
        """
        스키마 로더 초기화
        
        Args:
            schemas_dir: 스키마 디렉토리 경로 (기본값: 프로젝트 루트의 schemas/)
        """
        if schemas_dir is None:
            current_file = Path(__file__)
            project_root = current_file.parent.parent.parent
            schemas_dir = project_root / 'schemas'
        else:
            schemas_dir = Path(schemas_dir)
        
        self.schemas_dir = Path(schemas_dir)
    
    def load_name_mappings(self) -> Dict[str, Dict[str, str]]:
        """이름 매핑 스키마를 로드합니다."""
        mappings_file = self.schemas_dir / 'name_mappings.json'
        if not mappings_file.exists():
            raise FileNotFoundError(f"이름 매핑 파일을 찾을 수 없습니다: {mappings_file}")
        
        with open(mappings_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def get_name_mapping(self, mapping_name: str) -> Optional[Dict[str, str]]:
        """특정 이름 매핑을 가져옵니다."""
        mappings = self.load_name_mappings()
        return mappings.get(mapping_name)
    
    def load_sheet_config(self, sheet_name: str) -> Dict[str, Any]:
        """특정 시트의 설정을 로드합니다."""
        sheets_dir = self.schemas_dir / 'sheets'
        sheet_file = sheets_dir / f"{sheet_name}.json"
        
        if not sheet_file.exists():
            default_file = sheets_dir / 'default.json'
            if default_file.exists():
                with open(default_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            else:
                config = {
                    "category_column": 6,
                    "base_year": 2023,
                    "base_quarter": 1,
                    "base_col": 56,
                    "name_mapping": "industry_name_mapping",
                    "national_priorities": None,
                    "region_priorities": {}
                }
        else:
            with open(sheet_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
        
        if config.get('name_mapping') and isinstance(config['name_mapping'], str):
            mapping_name = config['name_mapping']
            name_mapping = self.get_name_mapping(mapping_name)
            if name_mapping:
                config['name_mapping'] = name_mapping
            else:
                config['name_mapping'] = {}
        elif config.get('name_mapping') is None:
            config['name_mapping'] = {}
        
        return config
    
    def load_template_mapping(self) -> Dict[str, Dict[str, str]]:
        """템플릿 매핑 스키마를 로드합니다."""
        mapping_file = self.schemas_dir / 'template_mapping.json'
        if not mapping_file.exists():
            raise FileNotFoundError(f"템플릿 매핑 파일을 찾을 수 없습니다: {mapping_file}")
        
        with open(mapping_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def get_template_for_sheet(self, sheet_name: str) -> Optional[Dict[str, str]]:
        """시트명에 해당하는 템플릿 정보를 반환합니다."""
        mapping = self.load_template_mapping()
        
        if sheet_name in mapping:
            return mapping[sheet_name]
        
        sheet_lower = sheet_name.lower()
        for key, value in mapping.items():
            if key.lower() in sheet_lower or sheet_lower in key.lower():
                return value
        
        if '소비' in sheet_name or '소매' in sheet_name:
            return mapping.get('소비(소매, 추가)')
        
        return None
    
    def reload(self):
        """스키마 캐시를 초기화합니다."""
        pass
    
    def load_base_format(self) -> Dict[str, Any]:
        """기본 출력 형식 스키마를 로드합니다."""
        base_file = self.schemas_dir / 'output_formats' / 'base.json'
        if not base_file.exists():
            return {
                "direction_expressions": {
                    "increase": {"rate": "증가", "production": "늘어", "result": "증가"},
                    "decrease": {"rate": "감소", "production": "줄어", "result": "감소"},
                    "rise": {"rate": "상승", "change": "올라", "result": "상승"},
                    "fall": {"rate": "하락", "change": "내려", "result": "하락"}
                },
                "format_rules": {
                    "percentage": {"decimal_places": 1, "suffix": "%"},
                    "percentage_point": {"decimal_places": 1, "suffix": "%p"},
                    "count": {"decimal_places": 0, "use_comma": True, "suffix": ""},
                    "population": {"decimal_places": 0, "use_comma": True, "suffix": "명"}
                },
                "bullet_styles": {"diamond": "◆", "square": "□", "circle": "○", "dot": "·"}
            }
        else:
            with open(base_file, 'r', encoding='utf-8') as f:
                return json.load(f)
    
    def load_output_format(self, display_name: str) -> Optional[Dict[str, Any]]:
        """특정 템플릿의 출력 형식 스키마를 로드합니다."""
        output_formats_dir = self.schemas_dir / 'output_formats'
        if not output_formats_dir.exists():
            return None
        
        format_file = output_formats_dir / f"{display_name}.json"
        if not format_file.exists():
            return None
        
        with open(format_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def get_output_format_for_sheet(self, sheet_name: str) -> Optional[Dict[str, Any]]:
        """시트명에 해당하는 출력 형식 스키마를 반환합니다."""
        template_info = self.get_template_for_sheet(sheet_name)
        if template_info and 'display_name' in template_info:
            return self.load_output_format(template_info['display_name'])
        return self.load_output_format(sheet_name)
    
    def get_direction_expression(self, direction_type: str, expression_key: str) -> str:
        """방향 표현 문자열을 반환합니다."""
        base_format = self.load_base_format()
        direction_expressions = base_format.get('direction_expressions', {})
        direction = direction_expressions.get(direction_type, {})
        return direction.get(expression_key, '')
    
    def get_format_rule(self, rule_name: str) -> Dict[str, Any]:
        """포맷 규칙을 반환합니다."""
        base_format = self.load_base_format()
        format_rules = base_format.get('format_rules', {})
        return format_rules.get(rule_name, {})
    
    def list_available_output_formats(self) -> list:
        """사용 가능한 출력 형식 목록을 반환합니다."""
        output_formats_dir = self.schemas_dir / 'output_formats'
        if not output_formats_dir.exists():
            return []
        
        formats = []
        for format_file in output_formats_dir.glob('*.json'):
            if format_file.name != 'base.json':
                formats.append(format_file.stem)
        
        return sorted(formats)
    
    def list_available_sheets(self) -> list:
        """사용 가능한 시트 설정 목록을 반환합니다."""
        sheets_dir = self.schemas_dir / 'sheets'
        if not sheets_dir.exists():
            return []
        
        sheets = []
        for sheet_file in sheets_dir.glob('*.json'):
            if sheet_file.name not in ('default.json', 'weight_ranking_config.json'):
                sheets.append(sheet_file.stem)
        
        return sorted(sheets)
    
    def get_weight_config(self, sheet_name: str) -> Dict[str, Any]:
        """특정 시트의 가중치 설정을 반환합니다."""
        config = self.load_sheet_config(sheet_name)
        
        return {
            'weight_column': config.get('weight_column'),
            'weight_default': config.get('weight_default', 1),
            'classification_column': config.get('classification_column', 2),
            'max_classification_level': config.get('max_classification_level', 2),
            'use_weighted_ranking': config.get('use_weighted_ranking', True),
            'region_column': config.get('region_column'),
            'category_column': config.get('category_column')
        }
    
    def get_weight_value(self, sheet_name: str, row_weight_value: Any) -> float:
        """행의 가중치 값을 결정합니다."""
        weight_config = self.get_weight_config(sheet_name)
        weight_column = weight_config.get('weight_column')
        weight_default = weight_config.get('weight_default', 1)
        
        if weight_column is None:
            return 1.0
        
        if row_weight_value is None:
            return 100.0 if weight_default is None else float(weight_default)
        
        if isinstance(row_weight_value, str) and not row_weight_value.strip():
            return 100.0 if weight_default is None else float(weight_default)
        
        try:
            return float(row_weight_value)
        except (ValueError, TypeError):
            return 100.0 if weight_default is None else float(weight_default)
    
    def should_include_classification_level(self, sheet_name: str, classification_level: Any) -> bool:
        """분류단계가 포함되어야 하는지 확인합니다."""
        weight_config = self.get_weight_config(sheet_name)
        max_level = weight_config.get('max_classification_level')
        
        if max_level is None:
            return True
        
        if classification_level is None:
            return True
        
        try:
            level = float(classification_level)
            return level <= max_level
        except (ValueError, TypeError):
            return True

