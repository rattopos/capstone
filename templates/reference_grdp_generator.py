#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
참고 GRDP (지역내총생산) 보고서 생성기
- GRDP 데이터는 결측치이므로 플레이스홀더로 생성
- 사용자가 직접 입력할 수 있도록 편집 가능한 필드 제공
"""

import json
from pathlib import Path
from jinja2 import Template


class 참고_GRDP_Generator:
    """참고 GRDP 보고서 생성기"""
    
    # 지역 순서 (차트 및 테이블 순서)
    REGION_ORDER = [
        {"group": None, "region": "전국"},
        {"group": "경인", "region": "서울"},
        {"group": "경인", "region": "인천"},
        {"group": "경인", "region": "경기"},
        {"group": "충청", "region": "대전"},
        {"group": "충청", "region": "세종"},
        {"group": "충청", "region": "충북"},
        {"group": "충청", "region": "충남"},
        {"group": "호남", "region": "광주"},
        {"group": "호남", "region": "전북"},
        {"group": "호남", "region": "전남"},
        {"group": "호남", "region": "제주"},
        {"group": "동북", "region": "대구"},
        {"group": "동북", "region": "경북"},
        {"group": "동북", "region": "강원"},
        {"group": "동남", "region": "부산"},
        {"group": "동남", "region": "울산"},
        {"group": "동남", "region": "경남"},
    ]
    
    def __init__(self, excel_path=None):
        """
        Args:
            excel_path: 분석표 엑셀 파일 경로 (GRDP는 결측치이므로 사용하지 않음)
        """
        self.excel_path = excel_path
    
    def generate_placeholder_data(self, year=2025, quarter=2):
        """플레이스홀더 데이터 생성 (모든 값이 결측치)"""
        
        # 전국 요약 데이터 (플레이스홀더)
        national_summary = {
            "growth_rate": 0.0,
            "direction": "증가",
            "contributions": {
                "service": 0.0,
                "manufacturing": 0.0,
                "other": 0.0,
                "construction": 0.0
            },
            "placeholder": True
        }
        
        # 성장률 1위 지역 (플레이스홀더)
        top_region = {
            "name": "[지역명]",
            "growth_rate": 0.0,
            "contributions": {
                "manufacturing": 0.0,
                "service": 0.0,
                "other": 0.0,
                "construction": 0.0
            },
            "placeholder": True
        }
        
        # 시도별 데이터 (모두 플레이스홀더)
        regional_data = []
        for region_info in self.REGION_ORDER:
            regional_data.append({
                "region_group": region_info["group"],
                "region": region_info["region"],
                "growth_rate": 0.0,
                "manufacturing": 0.0,
                "construction": 0.0,
                "service": 0.0,
                "other": 0.0,
                "placeholder": True
            })
        
        return {
            "report_info": {
                "year": year,
                "quarter": quarter,
                "page_number": 20
            },
            "national_summary": national_summary,
            "top_region": top_region,
            "regional_data": regional_data,
            "chart_config": {
                "y_axis": {
                    "min": -6,
                    "max": 8,
                    "step": 2
                }
            }
        }
    
    def generate_sample_data(self, year=2025, quarter=2):
        """이미지 기준 샘플 데이터 생성 (테스트용)"""
        
        # 전국 요약 데이터
        national_summary = {
            "growth_rate": 0.4,
            "direction": "증가",
            "contributions": {
                "service": 0.7,
                "manufacturing": 0.5,
                "other": -0.2,
                "construction": -0.6
            },
            "placeholder": False
        }
        
        # 성장률 1위 지역 (충북)
        top_region = {
            "name": "충북",
            "growth_rate": 5.8,
            "contributions": {
                "manufacturing": 5.1,
                "service": 0.8,
                "other": 0.5,
                "construction": -0.6
            },
            "placeholder": False
        }
        
        # 시도별 데이터 (이미지 기준)
        sample_values = {
            "전국": {"growth_rate": 0.4, "manufacturing": 0.5, "construction": -0.6, "service": 0.7, "other": -0.2},
            "서울": {"growth_rate": 1.2, "manufacturing": -0.1, "construction": -0.1, "service": 1.5, "other": -0.1},
            "인천": {"growth_rate": -1.6, "manufacturing": -1.2, "construction": -0.4, "service": 0.5, "other": -0.5},
            "경기": {"growth_rate": 2.7, "manufacturing": 2.5, "construction": -0.8, "service": 1.0, "other": 0.0},
            "대전": {"growth_rate": -0.6, "manufacturing": -0.4, "construction": -0.5, "service": 0.5, "other": -0.2},
            "세종": {"growth_rate": -0.3, "manufacturing": 0.0, "construction": -1.3, "service": 1.2, "other": -0.2},
            "충북": {"growth_rate": 5.8, "manufacturing": 5.1, "construction": -0.6, "service": 0.8, "other": 0.5},
            "충남": {"growth_rate": -3.9, "manufacturing": -2.3, "construction": -0.4, "service": 0.3, "other": -1.5},
            "광주": {"growth_rate": -0.9, "manufacturing": 0.0, "construction": -0.7, "service": 0.0, "other": -0.2},
            "전북": {"growth_rate": -0.9, "manufacturing": 0.5, "construction": -0.8, "service": -0.4, "other": -0.2},
            "전남": {"growth_rate": -3.2, "manufacturing": -1.2, "construction": -1.2, "service": 0.1, "other": -1.0},
            "제주": {"growth_rate": -3.7, "manufacturing": -0.1, "construction": -1.1, "service": -2.9, "other": 0.4},
            "대구": {"growth_rate": -3.2, "manufacturing": -1.0, "construction": -1.2, "service": -0.7, "other": -0.3},
            "경북": {"growth_rate": 1.9, "manufacturing": 2.7, "construction": -1.1, "service": 0.1, "other": 0.3},
            "강원": {"growth_rate": -0.5, "manufacturing": 0.0, "construction": -1.0, "service": 0.2, "other": 0.3},
            "부산": {"growth_rate": 0.7, "manufacturing": -0.8, "construction": -0.2, "service": 1.9, "other": -0.3},
            "울산": {"growth_rate": -1.0, "manufacturing": -0.6, "construction": -0.3, "service": 0.3, "other": -0.4},
            "경남": {"growth_rate": -2.2, "manufacturing": -0.4, "construction": -0.7, "service": -0.5, "other": -0.7},
        }
        
        regional_data = []
        for region_info in self.REGION_ORDER:
            region = region_info["region"]
            values = sample_values.get(region, {})
            regional_data.append({
                "region_group": region_info["group"],
                "region": region,
                "growth_rate": values.get("growth_rate", 0.0),
                "manufacturing": values.get("manufacturing", 0.0),
                "construction": values.get("construction", 0.0),
                "service": values.get("service", 0.0),
                "other": values.get("other", 0.0),
                "placeholder": False
            })
        
        return {
            "report_info": {
                "year": year,
                "quarter": quarter,
                "page_number": 20
            },
            "national_summary": national_summary,
            "top_region": top_region,
            "regional_data": regional_data,
            "chart_config": {
                "y_axis": {
                    "min": -6,
                    "max": 8,
                    "step": 2
                }
            }
        }
    
    def render_html(self, template_path, year=2025, quarter=2, use_sample=False):
        """HTML 렌더링
        
        Args:
            template_path: 템플릿 파일 경로
            year: 연도
            quarter: 분기
            use_sample: True면 샘플 데이터, False면 플레이스홀더
        """
        # 데이터 생성
        if use_sample:
            data = self.generate_sample_data(year, quarter)
        else:
            data = self.generate_placeholder_data(year, quarter)
        
        # 템플릿 로드 및 렌더링
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        return template.render(**data)
    
    def save_json(self, output_path, year=2025, quarter=2, use_sample=False):
        """JSON 데이터 저장"""
        if use_sample:
            data = self.generate_sample_data(year, quarter)
        else:
            data = self.generate_placeholder_data(year, quarter)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        return data


def generate_report_data(excel_path, year=2025, quarter=2, use_sample=False):
    """보고서 데이터 생성 함수 (app.py에서 호출)
    
    우선순위:
    1. 추출된 GRDP JSON 파일이 있으면 사용
    2. use_sample=True면 샘플 데이터 사용
    3. 그 외 플레이스홀더 데이터 사용
    """
    import json
    
    # 1. 추출된 GRDP JSON 파일 확인
    grdp_json_path = Path(__file__).parent / 'grdp_extracted.json'
    if grdp_json_path.exists():
        try:
            with open(grdp_json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            # 연도/분기 업데이트
            data['report_info']['year'] = year
            data['report_info']['quarter'] = quarter
            print(f"[GRDP Generator] JSON에서 GRDP 데이터 로드 (전국 {data['national_summary']['growth_rate']}%)")
            return data
        except Exception as e:
            print(f"[GRDP Generator] JSON 로드 실패: {e}")
    
    # 2. Generator 사용
    generator = 참고_GRDP_Generator(excel_path)
    
    if use_sample:
        return generator.generate_sample_data(year, quarter)
    else:
        return generator.generate_placeholder_data(year, quarter)


if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='참고 GRDP 보고서 생성')
    parser.add_argument('--template', type=str, default='templates/reference_grdp_template.html',
                        help='템플릿 파일 경로')
    parser.add_argument('--output', type=str, default='templates/reference_grdp_output.html',
                        help='출력 HTML 파일 경로')
    parser.add_argument('--json', type=str, default='templates/reference_grdp_data.json',
                        help='출력 JSON 파일 경로')
    parser.add_argument('--year', type=int, default=2025, help='연도')
    parser.add_argument('--quarter', type=int, default=2, help='분기')
    parser.add_argument('--sample', action='store_true', help='샘플 데이터 사용')
    
    args = parser.parse_args()
    
    # Generator 생성
    generator = 참고_GRDP_Generator()
    
    # JSON 저장
    data = generator.save_json(args.json, args.year, args.quarter, args.sample)
    print(f"JSON 저장 완료: {args.json}")
    
    # HTML 렌더링
    html = generator.render_html(args.template, args.year, args.quarter, args.sample)
    
    with open(args.output, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"HTML 저장 완료: {args.output}")
    
    print(f"\n=== 생성 완료 ===")
    print(f"모드: {'샘플 데이터' if args.sample else '플레이스홀더 (결측치)'}")
    print(f"전국 성장률: {data['national_summary']['growth_rate']}%")
    print(f"1위 지역: {data['top_region']['name']} ({data['top_region']['growth_rate']}%)")

