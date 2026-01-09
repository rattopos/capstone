# -*- coding: utf-8 -*-
"""
HWPX 데이터 주입기
양식.hwpx 템플릿에 데이터를 주입하고 17개 시도 섹션을 생성합니다.
"""

import os
import re
import json
import shutil
import zipfile
import tempfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from lxml import etree


class HWPXDataInjector:
    """HWPX 템플릿에 데이터를 주입하는 클래스"""
    
    # HWPX XML 네임스페이스
    NAMESPACES = {
        'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
        'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
        'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    }
    
    # 서울특별시 주요지표 테이블 ID (양식.hwpx에서 추출)
    SEOUL_TABLE_ID = "1992325869"
    
    # 17개 시도 정보
    SIDO_LIST = [
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
    ]
    
    # 지표 순서 (테이블 열 순서)
    INDICATOR_ORDER = [
        "manufacturing",  # 광공업생산
        "service",        # 서비스업생산
        "retail",         # 소매판매
        "construction",   # 건설수주
        "export",         # 수출
        "import",         # 수입
        "price",          # 소비자물가
        "employment",     # 고용률
        "migration_20_29",  # 인구순이동 20-29
        "migration_all",    # 인구순이동 전체
    ]
    
    def __init__(self, template_path: str = None):
        """
        Args:
            template_path: 양식.hwpx 템플릿 파일 경로
        """
        if template_path is None:
            # 기본 경로: utils/양식.hwpx
            base_dir = Path(__file__).parent.parent.parent
            template_path = base_dir / "utils" / "양식.hwpx"
        
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {self.template_path}")
        
        self.temp_dir = None
        self.section_xml = None
        self.seoul_section_template = None
    
    def extract_hwpx(self) -> Path:
        """HWPX 파일을 임시 디렉토리에 압축 해제
        
        Returns:
            압축 해제된 디렉토리 경로
        """
        self.temp_dir = Path(tempfile.mkdtemp(prefix="hwpx_"))
        
        with zipfile.ZipFile(self.template_path, 'r') as zf:
            zf.extractall(self.temp_dir)
        
        return self.temp_dir
    
    def load_section_xml(self) -> str:
        """section0.xml 파일 로드
        
        Returns:
            section0.xml 내용
        """
        if self.temp_dir is None:
            self.extract_hwpx()
        
        section_path = self.temp_dir / "Contents" / "section0.xml"
        with open(section_path, 'r', encoding='utf-8') as f:
            self.section_xml = f.read()
        
        return self.section_xml
    
    def find_seoul_section(self) -> Tuple[int, int, str]:
        """서울특별시 주요지표 섹션을 찾아서 추출
        
        Returns:
            (시작 위치, 끝 위치, 섹션 XML)
        """
        if self.section_xml is None:
            self.load_section_xml()
        
        # "서울특별시 주요지표" 텍스트가 포함된 문단 찾기
        seoul_marker = "서울특별시 주요지표"
        seoul_pos = self.section_xml.find(seoul_marker)
        
        if seoul_pos == -1:
            raise ValueError("서울특별시 주요지표 섹션을 찾을 수 없습니다")
        
        # 해당 위치 이전의 가장 가까운 문단 시작 태그 찾기
        # 《 서울특별시 주요지표 》 형태의 제목 문단
        before_content = self.section_xml[:seoul_pos]
        
        # <hp:p 태그로 시작하는 문단 찾기 (제목 포함)
        # 제목 문단 시작 위치 찾기
        para_starts = [m.start() for m in re.finditer(r'<hp:p[^>]*>', before_content)]
        if not para_starts:
            raise ValueError("문단 시작을 찾을 수 없습니다")
        
        section_start = para_starts[-1]
        
        # 섹션 끝 찾기: "참고" 또는 다음 시도 섹션 시작 전까지
        after_seoul = self.section_xml[seoul_pos:]
        
        # "참고" 섹션 또는 "*" 각주 이후의 문단 찾기
        end_markers = [
            "참고",
            "* 광공업생산지수",
        ]
        
        section_end = len(self.section_xml)
        for marker in end_markers:
            marker_pos = after_seoul.find(marker)
            if marker_pos != -1:
                # 해당 마커가 포함된 문단의 끝 찾기
                after_marker = after_seoul[marker_pos:]
                para_end = after_marker.find("</hp:p>")
                if para_end != -1:
                    potential_end = seoul_pos + marker_pos + para_end + len("</hp:p>")
                    section_end = min(section_end, potential_end)
        
        # 추출된 섹션
        section_content = self.section_xml[section_start:section_end]
        
        return section_start, section_end, section_content
    
    def create_sido_template(self, seoul_section: str) -> str:
        """서울 섹션을 플레이스홀더 템플릿으로 변환
        
        Args:
            seoul_section: 서울특별시 주요지표 섹션 XML
            
        Returns:
            플레이스홀더가 삽입된 템플릿 XML
        """
        template = seoul_section
        
        # 1. 시도명 플레이스홀더
        template = template.replace("서울특별시", "{{SIDO_FULL_NAME}}")
        template = template.replace("서울", "{{SIDO_NAME}}")
        
        # 2. 테이블 내 숫자 데이터를 플레이스홀더로 변환
        # 숫자 패턴: -숫자 또는 숫자 (소수점 포함)
        # 분기별 데이터 행 식별
        quarters = ["23.2/4", "24.2/4", "25.1/4", "25.2/4"]
        
        # 각 분기별로 데이터 플레이스홀더 생성
        for q_idx, quarter in enumerate(quarters):
            q_key = quarter.replace(".", "_").replace("/", "Q")
            
            # 해당 분기 다음에 오는 숫자들을 플레이스홀더로 변환
            # 패턴: 분기 문자열 이후의 연속된 숫자 셀들
            pattern = rf"(<hp:t>{re.escape(quarter)}</hp:t>)"
            
            # 분기 위치 찾기
            q_match = re.search(pattern, template)
            if q_match:
                # 해당 분기 이후 다음 분기 또는 각주까지의 숫자들 찾기
                q_pos = q_match.end()
                
                # 다음 분기 또는 섹션 끝까지
                next_q_pos = len(template)
                for next_q in quarters[q_idx + 1:]:
                    next_match = re.search(rf"<hp:t>{re.escape(next_q)}</hp:t>", template[q_pos:])
                    if next_match:
                        next_q_pos = q_pos + next_match.start()
                        break
                
                # 해당 범위 내 숫자들을 플레이스홀더로
                range_content = template[q_pos:next_q_pos]
                
                # 숫자 패턴 찾기 및 치환
                indicator_idx = 0
                def replace_number(match):
                    nonlocal indicator_idx
                    if indicator_idx < len(self.INDICATOR_ORDER):
                        indicator = self.INDICATOR_ORDER[indicator_idx]
                        placeholder = f"{{{{DATA_{q_key}_{indicator}}}}}"
                        indicator_idx += 1
                        return f"<hp:t>{placeholder}</hp:t>"
                    return match.group(0)
                
                # 숫자 셀 패턴: <hp:t>-?숫자.숫자</hp:t>
                number_pattern = r'<hp:t>(-?\d+\.?\d*)</hp:t>'
                new_range = re.sub(number_pattern, replace_number, range_content)
                
                template = template[:q_pos] + new_range + template[next_q_pos:]
        
        return template
    
    def generate_sido_section(self, template: str, sido_info: Dict, data: Dict) -> str:
        """템플릿에 특정 시도 데이터를 삽입
        
        Args:
            template: 플레이스홀더 템플릿
            sido_info: 시도 정보 딕셔너리
            data: 해당 시도의 지표 데이터
            
        Returns:
            데이터가 삽입된 섹션 XML
        """
        result = template
        
        # 시도명 치환
        result = result.replace("{{SIDO_FULL_NAME}}", sido_info["full_name"])
        result = result.replace("{{SIDO_NAME}}", sido_info["name"])
        
        # 데이터 치환
        for placeholder_match in re.finditer(r'\{\{DATA_([^}]+)\}\}', result):
            placeholder = placeholder_match.group(0)
            key = placeholder_match.group(1)
            
            # 데이터에서 해당 키 찾기
            value = data.get(key, "N/A")
            if isinstance(value, (int, float)):
                value = f"{value:.1f}" if isinstance(value, float) else str(value)
            
            result = result.replace(placeholder, str(value))
        
        return result
    
    def inject_all_sido(self, all_data: Dict[str, Dict], selected_sido: List[str] = None) -> str:
        """모든 시도 섹션을 생성하여 문서에 삽입
        
        Args:
            all_data: 시도별 데이터 딕셔너리 {"서울": {...}, "부산": {...}, ...}
            selected_sido: 생성할 시도 목록 (None이면 전체)
            
        Returns:
            완성된 section0.xml 내용
        """
        if self.section_xml is None:
            self.load_section_xml()
        
        # 서울 섹션 추출
        start, end, seoul_section = self.find_seoul_section()
        
        # 템플릿 생성
        template = self.create_sido_template(seoul_section)
        
        # 선택된 시도 목록
        if selected_sido is None:
            selected_sido = [s["name"] for s in self.SIDO_LIST]
        
        # 각 시도별 섹션 생성
        all_sections = []
        for sido_info in self.SIDO_LIST:
            if sido_info["name"] in selected_sido:
                sido_data = all_data.get(sido_info["name"], {})
                section = self.generate_sido_section(template, sido_info, sido_data)
                all_sections.append(section)
        
        # 원본 문서에서 서울 섹션을 모든 시도 섹션으로 교체
        combined_sections = "\n".join(all_sections)
        new_xml = self.section_xml[:start] + combined_sections + self.section_xml[end:]
        
        return new_xml
    
    def save_hwpx(self, output_path: str, new_section_xml: str) -> bool:
        """수정된 내용으로 HWPX 파일 저장
        
        Args:
            output_path: 출력 파일 경로
            new_section_xml: 수정된 section0.xml 내용
            
        Returns:
            성공 여부
        """
        if self.temp_dir is None:
            raise RuntimeError("먼저 extract_hwpx()를 호출하세요")
        
        # section0.xml 업데이트
        section_path = self.temp_dir / "Contents" / "section0.xml"
        with open(section_path, 'w', encoding='utf-8') as f:
            f.write(new_section_xml)
        
        # HWPX로 다시 압축
        output_path = Path(output_path)
        
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(self.temp_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(self.temp_dir)
                    zf.write(file_path, arcname)
        
        return True
    
    def cleanup(self):
        """임시 파일 정리"""
        if self.temp_dir and self.temp_dir.exists():
            shutil.rmtree(self.temp_dir)
            self.temp_dir = None
    
    def inject(self, all_data: Dict[str, Dict], output_path: str, 
               selected_sido: List[str] = None) -> bool:
        """전체 워크플로우 실행
        
        Args:
            all_data: 시도별 데이터
            output_path: 출력 HWPX 파일 경로
            selected_sido: 생성할 시도 목록
            
        Returns:
            성공 여부
        """
        try:
            # 1. 템플릿 압축 해제
            self.extract_hwpx()
            
            # 2. section0.xml 로드
            self.load_section_xml()
            
            # 3. 모든 시도 섹션 생성
            new_xml = self.inject_all_sido(all_data, selected_sido)
            
            # 4. 결과 저장
            result = self.save_hwpx(output_path, new_xml)
            
            return result
            
        except Exception as e:
            print(f"HWPX 생성 오류: {e}")
            raise
        finally:
            self.cleanup()


# 테스트 코드
if __name__ == "__main__":
    # 테스트용 더미 데이터
    test_data = {
        "서울": {
            "DATA_23_2Q4_manufacturing": -9.7,
            "DATA_23_2Q4_service": 6.6,
            "DATA_23_2Q4_retail": -2.4,
            # ... 나머지 데이터
        }
    }
    
    injector = HWPXDataInjector()
    
    # 서울 섹션 분석
    injector.extract_hwpx()
    injector.load_section_xml()
    start, end, seoul_section = injector.find_seoul_section()
    
    print(f"서울 섹션 위치: {start} ~ {end}")
    print(f"섹션 길이: {end - start} 문자")
    print(f"\n서울 섹션 내용 (처음 500자):\n{seoul_section[:500]}")
    
    injector.cleanup()
