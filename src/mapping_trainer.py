"""
매핑 학습 모듈
템플릿과 엑셀 데이터 간의 매핑을 학습하고 파인튜닝
"""

from typing import Dict, List, Any, Optional, Tuple
from .excel_extractor import ExcelExtractor
from .auto_mapper import AutoMapper
from .template_manager import TemplateManager


class MappingTrainer:
    """템플릿과 엑셀 데이터 간의 매핑을 학습하는 클래스"""
    
    def __init__(self, excel_extractor: ExcelExtractor):
        """
        매핑 학습기 초기화
        
        Args:
            excel_extractor: 엑셀 추출기 인스턴스
        """
        self.excel_extractor = excel_extractor
        self.auto_mapper = AutoMapper(excel_extractor)
    
    def train_mapping(
        self,
        template_html: str,
        template_name: str,
        sheet_name: str,
        year: int,
        quarter: int
    ) -> Dict[str, Any]:
        """
        템플릿과 엑셀 데이터 간의 매핑을 학습합니다.
        
        Args:
            template_html: HTML 템플릿
            template_name: 템플릿 이름
            sheet_name: 엑셀 시트명
            year: 연도
            quarter: 분기
            
        Returns:
            학습된 매핑 정보 딕셔너리:
            - 'mappings': 필드별 매핑 정보
            - 'confidence_scores': 신뢰도 점수
            - 'validation_results': 검증 결과
        """
        # 템플릿에서 마커 추출
        template_manager = TemplateManager(template_path=None)
        template_manager.template_content = template_html
        markers = template_manager.extract_markers()
        
        # 각 마커에 대한 매핑 학습
        mappings = {}
        confidence_scores = {}
        
        for marker in markers:
            marker_info = {
                'sheet_name': marker['sheet_name'],
                'cell_address': marker['cell_address'],
                'operation': marker.get('operation')
            }
            
            # 마커 타입에 따라 다른 학습 전략 적용
            mapping_result = self._learn_marker_mapping(
                marker_info,
                sheet_name,
                year,
                quarter
            )
            
            if mapping_result:
                marker_key = marker['full_match']
                mappings[marker_key] = mapping_result['mapping']
                confidence_scores[marker_key] = mapping_result['confidence']
        
        # 매핑 검증
        validation_results = self._validate_mappings(
            template_html,
            mappings,
            sheet_name,
            year,
            quarter
        )
        
        return {
            'mappings': mappings,
            'confidence_scores': confidence_scores,
            'validation_results': validation_results,
            'template_name': template_name,
            'sheet_name': sheet_name
        }
    
    def _learn_marker_mapping(
        self,
        marker_info: Dict[str, Any],
        sheet_name: str,
        year: int,
        quarter: int
    ) -> Optional[Dict[str, Any]]:
        """개별 마커에 대한 매핑 학습"""
        marker_key = marker_info.get('cell_address', '')
        
        # 동적 마커인 경우 (셀 주소가 아닌 키워드)
        if not self._is_cell_address(marker_key):
            return self._learn_dynamic_marker_mapping(
                marker_info,
                sheet_name,
                year,
                quarter
            )
        
        # 정적 마커인 경우 (셀 주소)
        return self._learn_static_marker_mapping(
            marker_info,
            sheet_name
        )
    
    def _is_cell_address(self, address: str) -> bool:
        """셀 주소 형식인지 확인"""
        import re
        return bool(re.match(r'^[A-Z]+\d+', address))
    
    def _learn_dynamic_marker_mapping(
        self,
        marker_info: Dict[str, Any],
        sheet_name: str,
        year: int,
        quarter: int
    ) -> Optional[Dict[str, Any]]:
        """동적 마커 매핑 학습"""
        marker_key = marker_info.get('cell_address', '')
        
        # 마커 키에서 패턴 추출
        # 예: "전국_증감률", "상위시도1_이름" 등
        
        # 전국 관련 마커
        if marker_key.startswith('전국'):
            if '증감률' in marker_key:
                return {
                    'mapping': {
                        'type': 'dynamic',
                        'pattern': 'national_growth_rate',
                        'sheet_name': sheet_name,
                        'year': year,
                        'quarter': quarter
                    },
                    'confidence': 0.9
                }
            elif '이름' in marker_key:
                return {
                    'mapping': {
                        'type': 'dynamic',
                        'pattern': 'national_name',
                        'sheet_name': sheet_name
                    },
                    'confidence': 0.9
                }
        
        # 시도 관련 마커
        if '시도' in marker_key:
            # 상위/하위 시도 추출
            if '상위' in marker_key:
                idx_match = self._extract_number(marker_key)
                idx = int(idx_match) if idx_match else 1
                
                if '증감률' in marker_key:
                    return {
                        'mapping': {
                            'type': 'dynamic',
                            'pattern': 'top_region_growth_rate',
                            'index': idx,
                            'sheet_name': sheet_name,
                            'year': year,
                            'quarter': quarter
                        },
                        'confidence': 0.85
                    }
                elif '이름' in marker_key:
                    return {
                        'mapping': {
                            'type': 'dynamic',
                            'pattern': 'top_region_name',
                            'index': idx,
                            'sheet_name': sheet_name
                        },
                        'confidence': 0.85
                    }
        
        # 산업/업태 관련 마커
        if '산업' in marker_key or '업태' in marker_key:
            return self._learn_industry_marker_mapping(
                marker_key,
                sheet_name,
                year,
                quarter
            )
        
        # 기본값
        return {
            'mapping': {
                'type': 'dynamic',
                'pattern': 'unknown',
                'sheet_name': sheet_name
            },
            'confidence': 0.5
        }
    
    def _learn_static_marker_mapping(
        self,
        marker_info: Dict[str, Any],
        sheet_name: str
    ) -> Optional[Dict[str, Any]]:
        """정적 마커 매핑 학습 (셀 주소 직접 사용)"""
        cell_address = marker_info.get('cell_address', '')
        operation = marker_info.get('operation')
        
        return {
            'mapping': {
                'type': 'static',
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'operation': operation
            },
            'confidence': 0.95  # 정적 마커는 높은 신뢰도
        }
    
    def _learn_industry_marker_mapping(
        self,
        marker_key: str,
        sheet_name: str,
        year: int,
        quarter: int
    ) -> Optional[Dict[str, Any]]:
        """산업/업태 마커 매핑 학습"""
        # 예: "전국_산업1_이름", "상위시도1_업태1_증감률" 등
        
        if '전국' in marker_key:
            idx_match = self._extract_number(marker_key)
            idx = int(idx_match) if idx_match else 1
            
            if '이름' in marker_key:
                return {
                    'mapping': {
                        'type': 'dynamic',
                        'pattern': 'national_industry_name',
                        'index': idx,
                        'sheet_name': sheet_name
                    },
                    'confidence': 0.8
                }
            elif '증감률' in marker_key:
                return {
                    'mapping': {
                        'type': 'dynamic',
                        'pattern': 'national_industry_growth_rate',
                        'index': idx,
                        'sheet_name': sheet_name,
                        'year': year,
                        'quarter': quarter
                    },
                    'confidence': 0.8
                }
        
        return None
    
    def _extract_number(self, text: str) -> Optional[str]:
        """텍스트에서 숫자 추출"""
        import re
        match = re.search(r'\d+', text)
        return match.group(0) if match else None
    
    def _validate_mappings(
        self,
        template_html: str,
        mappings: Dict[str, Any],
        sheet_name: str,
        year: int,
        quarter: int
    ) -> Dict[str, Any]:
        """매핑 검증"""
        validation_results = {
            'total_markers': len(mappings),
            'valid_mappings': 0,
            'invalid_mappings': 0,
            'errors': []
        }
        
        for marker_key, mapping in mappings.items():
            try:
                # 매핑이 유효한지 테스트
                is_valid = self._test_mapping(
                    mapping,
                    sheet_name,
                    year,
                    quarter
                )
                
                if is_valid:
                    validation_results['valid_mappings'] += 1
                else:
                    validation_results['invalid_mappings'] += 1
                    validation_results['errors'].append({
                        'marker': marker_key,
                        'error': 'Mapping test failed'
                    })
            except Exception as e:
                validation_results['invalid_mappings'] += 1
                validation_results['errors'].append({
                    'marker': marker_key,
                    'error': str(e)
                })
        
        return validation_results
    
    def _test_mapping(
        self,
        mapping: Dict[str, Any],
        sheet_name: str,
        year: int,
        quarter: int
    ) -> bool:
        """매핑이 유효한지 테스트"""
        try:
            if mapping['type'] == 'static':
                # 정적 마커: 셀 주소로 직접 테스트
                cell_address = mapping.get('cell_address')
                if cell_address:
                    value = self.excel_extractor.extract_value(sheet_name, cell_address)
                    return value is not None
            elif mapping['type'] == 'dynamic':
                # 동적 마커: 패턴 기반 테스트
                pattern = mapping.get('pattern')
                if pattern:
                    # 패턴이 유효한지 확인 (시트 구조 확인)
                    return self._validate_dynamic_pattern(
                        pattern,
                        sheet_name,
                        mapping
                    )
        except Exception:
            return False
        
        return False
    
    def _validate_dynamic_pattern(
        self,
        pattern: str,
        sheet_name: str,
        mapping: Dict[str, Any]
    ) -> bool:
        """동적 패턴 검증"""
        try:
            sheet = self.excel_extractor.get_sheet(sheet_name)
            # 시트가 존재하고 데이터가 있는지 확인
            return sheet.max_row > 0 and sheet.max_column > 0
        except Exception:
            return False

