# -*- coding: utf-8 -*-
"""
섹션 복제기
서울특별시 주요지표 섹션을 17개 시도로 복제합니다.
"""

import re
from typing import Dict, List, Tuple
from pathlib import Path
import json


class SectionReplicator:
    """서울 섹션을 다른 시도로 복제하는 클래스"""
    
    def __init__(self):
        """설정 로드"""
        config_path = Path(__file__).parent.parent / "config" / "sido_config.json"
        
        if config_path.exists():
            with open(config_path, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
        else:
            # 기본 설정
            self.config = self._get_default_config()
        
        self.sido_list = self.config.get("sido_list", [])
        self.indicators = self.config.get("indicators", [])
    
    def _get_default_config(self) -> Dict:
        """기본 설정 반환"""
        return {
            "sido_list": [
                {"code": "11", "name": "서울", "full_name": "서울특별시"},
                {"code": "21", "name": "부산", "full_name": "부산광역시"},
                {"code": "22", "name": "대구", "full_name": "대구광역시"},
                {"code": "23", "name": "인천", "full_name": "인천광역시"},
                {"code": "24", "name": "광주", "full_name": "광주광역시"},
                {"code": "25", "name": "대전", "full_name": "대전광역시"},
                {"code": "26", "name": "울산", "full_name": "울산광역시"},
                {"code": "29", "name": "세종", "full_name": "세종특별자치시"},
                {"code": "31", "name": "경기", "full_name": "경기도"},
                {"code": "32", "name": "강원", "full_name": "강원특별자치도"},
                {"code": "33", "name": "충북", "full_name": "충청북도"},
                {"code": "34", "name": "충남", "full_name": "충청남도"},
                {"code": "35", "name": "전북", "full_name": "전북특별자치도"},
                {"code": "36", "name": "전남", "full_name": "전라남도"},
                {"code": "37", "name": "경북", "full_name": "경상북도"},
                {"code": "38", "name": "경남", "full_name": "경상남도"},
                {"code": "39", "name": "제주", "full_name": "제주특별자치도"},
            ],
            "indicators": [
                {"id": "manufacturing", "name": "광공업생산", "unit": "%"},
                {"id": "service", "name": "서비스업생산", "unit": "%"},
                {"id": "retail", "name": "소매판매", "unit": "%"},
                {"id": "construction", "name": "건설수주", "unit": "%"},
                {"id": "export", "name": "수출", "unit": "%"},
                {"id": "import", "name": "수입", "unit": "%"},
                {"id": "price", "name": "소비자물가", "unit": "%"},
                {"id": "employment", "name": "고용률", "unit": "%p"},
                {"id": "migration", "name": "인구순이동", "unit": "천명"},
            ]
        }
    
    def get_sido_names(self) -> List[str]:
        """시도명 목록 반환"""
        return [s["name"] for s in self.sido_list]
    
    def get_sido_full_names(self) -> List[str]:
        """시도 전체명 목록 반환"""
        return [s["full_name"] for s in self.sido_list]
    
    def get_sido_info(self, name: str) -> Dict:
        """시도명으로 정보 조회"""
        for sido in self.sido_list:
            if sido["name"] == name or sido["full_name"] == name:
                return sido
        return None
    
    def extract_section_boundaries(self, xml_content: str, sido_name: str = "서울특별시") -> Tuple[int, int]:
        """특정 시도 주요지표 섹션의 경계 찾기
        
        Args:
            xml_content: section0.xml 내용
            sido_name: 찾을 시도명 (기본값: 서울특별시)
            
        Returns:
            (시작 위치, 끝 위치)
        """
        # 시도명 주요지표 마커 찾기
        marker = f"{sido_name} 주요지표"
        marker_pos = xml_content.find(marker)
        
        if marker_pos == -1:
            raise ValueError(f"{marker}를 찾을 수 없습니다")
        
        # 이전 《 문자 찾기 (제목 시작)
        before = xml_content[:marker_pos]
        title_start = before.rfind("《")
        
        if title_start == -1:
            # 대안: 가장 가까운 <hp:p 태그
            para_starts = [m.start() for m in re.finditer(r'<hp:p[^>]*>', before)]
            if para_starts:
                title_start = para_starts[-1]
            else:
                title_start = marker_pos - 100
        
        # 섹션 끝 찾기: 다음 시도 또는 "참고" 섹션
        after = xml_content[marker_pos:]
        
        # 끝 마커들
        end_markers = [
            "특별시 주요지표",
            "광역시 주요지표", 
            "특별자치시 주요지표",
            "특별자치도 주요지표",
            "도 주요지표",
            "참고",
        ]
        
        section_end = len(xml_content)
        for end_marker in end_markers:
            if end_marker in sido_name:
                continue
            
            end_pos = after.find(end_marker)
            if end_pos != -1 and end_pos > 100:  # 최소 거리
                # 해당 마커 이전의 》 찾기
                before_end = after[:end_pos]
                bracket_pos = before_end.rfind("》")
                if bracket_pos != -1:
                    potential_end = marker_pos + bracket_pos + 1
                    # </hp:p> 태그 끝까지 확장
                    para_end = xml_content[potential_end:].find("</hp:p>")
                    if para_end != -1:
                        potential_end += para_end + len("</hp:p>")
                    section_end = min(section_end, potential_end)
        
        return title_start, section_end
    
    def create_placeholder_template(self, section_xml: str) -> str:
        """섹션 XML을 플레이스홀더 템플릿으로 변환
        
        Args:
            section_xml: 서울특별시 주요지표 섹션 XML
            
        Returns:
            플레이스홀더가 삽입된 템플릿
        """
        template = section_xml
        
        # 1. 시도명 플레이스홀더
        # 서울특별시 → {{SIDO_FULL_NAME}}
        # 서울 → {{SIDO_NAME}}
        for sido in self.sido_list:
            # 전체명 먼저 (더 긴 문자열)
            template = template.replace(sido["full_name"], "{{SIDO_FULL_NAME}}")
        
        for sido in self.sido_list:
            # 단축명은 단어 경계 고려
            template = re.sub(
                rf'(?<![가-힣]){re.escape(sido["name"])}(?![가-힣])',
                "{{SIDO_NAME}}",
                template
            )
        
        # 2. 숫자 데이터 플레이스홀더
        # 테이블 내 숫자들을 순서대로 플레이스홀더로 변환
        quarters = ["23.2/4", "24.2/4", "25.1/4", "25.2/4"]
        indicators = [ind["id"] for ind in self.indicators]
        
        for q_idx, quarter in enumerate(quarters):
            q_key = quarter.replace(".", "_").replace("/", "Q")
            
            # 해당 분기 이후의 숫자들 찾기
            q_pattern = rf"(<hp:t>'{re.escape(quarter)}</hp:t>)"
            q_match = re.search(q_pattern, template)
            
            if not q_match:
                # 따옴표 없는 버전 시도
                q_pattern = rf"(<hp:t>{re.escape(quarter)}</hp:t>)"
                q_match = re.search(q_pattern, template)
            
            if q_match:
                q_end = q_match.end()
                
                # 다음 분기까지의 범위 찾기
                next_q_pos = len(template)
                for next_q in quarters[q_idx + 1:]:
                    for pattern in [f"'{next_q}", next_q]:
                        next_match = template[q_end:].find(f"<hp:t>{pattern}</hp:t>")
                        if next_match != -1:
                            next_q_pos = min(next_q_pos, q_end + next_match)
                            break
                
                # 해당 범위 내 숫자들을 플레이스홀더로 변환
                range_content = template[q_end:next_q_pos]
                
                ind_idx = [0]  # 클로저를 위한 리스트
                def replace_num(match):
                    if ind_idx[0] < len(indicators):
                        indicator = indicators[ind_idx[0]]
                        placeholder = f"{{{{DATA_{q_key}_{indicator}}}}}"
                        ind_idx[0] += 1
                        return f"<hp:t>{placeholder}</hp:t>"
                    return match.group(0)
                
                # 숫자 패턴
                num_pattern = r'<hp:t>(-?\d+\.?\d*)</hp:t>'
                new_range = re.sub(num_pattern, replace_num, range_content)
                
                template = template[:q_end] + new_range + template[next_q_pos:]
        
        return template
    
    def apply_sido_data(self, template: str, sido_info: Dict, data: Dict) -> str:
        """템플릿에 시도 데이터 적용
        
        Args:
            template: 플레이스홀더 템플릿
            sido_info: 시도 정보 {"name": "부산", "full_name": "부산광역시", ...}
            data: 해당 시도의 데이터 딕셔너리
            
        Returns:
            데이터가 적용된 섹션 XML
        """
        result = template
        
        # 시도명 치환
        result = result.replace("{{SIDO_FULL_NAME}}", sido_info["full_name"])
        result = result.replace("{{SIDO_NAME}}", sido_info["name"])
        
        # 데이터 플레이스홀더 치환
        for key, value in data.items():
            placeholder = f"{{{{{key}}}}}"
            
            # 숫자 포맷팅
            if isinstance(value, float):
                formatted = f"{value:.1f}"
            elif isinstance(value, int):
                formatted = str(value)
            elif value is None:
                formatted = "N/A"
            else:
                formatted = str(value)
            
            result = result.replace(placeholder, formatted)
        
        # 남은 플레이스홀더를 N/A로 치환
        result = re.sub(r'\{\{DATA_[^}]+\}\}', 'N/A', result)
        
        return result
    
    def replicate_for_all_sido(self, template: str, all_data: Dict[str, Dict], 
                               selected: List[str] = None) -> List[str]:
        """모든 시도에 대해 섹션 복제
        
        Args:
            template: 플레이스홀더 템플릿
            all_data: 시도별 데이터 {"서울": {...}, "부산": {...}, ...}
            selected: 생성할 시도 목록 (None이면 전체)
            
        Returns:
            시도별 섹션 XML 리스트
        """
        sections = []
        
        for sido in self.sido_list:
            if selected is not None and sido["name"] not in selected:
                continue
            
            sido_data = all_data.get(sido["name"], {})
            section = self.apply_sido_data(template, sido, sido_data)
            sections.append(section)
        
        return sections
