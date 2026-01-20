#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
통합 보도자료 생성기
모든 개별 Generator들을 통합하여 관리하고, 데이터 추출 및 HTML 생성을 처리합니다.
"""

from __future__ import annotations


import sys
import json
import importlib.util
from pathlib import Path
from typing import Any

from jinja2 import Template
import pandas as pd

from config.settings import BASE_DIR, TEMPLATES_DIR, TEMP_OUTPUT_DIR, TEMP_REGIONAL_OUTPUT_DIR
from config.reports import REPORT_ORDER, SECTOR_REPORTS, SUMMARY_REPORTS, REGIONAL_REPORTS
from services.report_generator import generate_regional_report_html, generate_report_html
from utils.filters import is_missing, format_value
from utils.text_utils import get_josa, get_terms, get_comparative_terms


class ReportGenerator:
    """통합 보도자료 생성기"""
    
    def __init__(self, excel_path: str) -> None:
        """
        초기화
        Args:
            excel_path: 엑셀 파일 경로
        """
        self.excel_path = Path(excel_path)
        if not self.excel_path.exists():
            raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {excel_path}")
        self._modules: dict[str, Any] = {}
        # 파일명에서 연/분기 자동 추정 (예: '분석표_25년 3분기_...xlsx')
        self._year_quarter: tuple[int, int] | None = self._infer_period_from_filename(
            self.excel_path.name
        )
    
    def _load_module(self, generator_name: str) -> Any:
        """Generator 모듈 동적 로드"""
        if generator_name in self._modules:
            return self._modules[generator_name]
        
        generator_path = TEMPLATES_DIR / generator_name
        if not generator_path.exists():
            raise FileNotFoundError(f"Generator 파일을 찾을 수 없습니다: {generator_path}")
        
        spec = importlib.util.spec_from_file_location(
            generator_name.replace('.py', ''),
            str(generator_path)
        )
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        
        self._modules[generator_name] = module
        return module

    @staticmethod
    def _infer_period_from_filename(filename: str) -> tuple[int, int] | None:
        """파일명에서 연도/분기 추정. 없으면 None 반환.
        허용 패턴 예시:
        - '분석표_25년 3분기_캡스톤.xlsx' -> (2025, 3)
        - '보고서_2025년3분기.xlsx' -> (2025, 3)
        - '2025_3Q.xlsx' -> (2025, 3)
        """
        import re
        s = filename
        # 1) '2025년 3분기' 또는 '25년 3분기'
        m = re.search(r"(\d{2,4})\s*년\s*(\d)\s*분기", s)
        if m:
            y = int(m.group(1))
            if y < 100:  # 두 자리 연도 처리 (예: 25 -> 2025)
                y += 2000
            q = int(m.group(2))
            return (y, q)
        # 2) '2025 3/4' 또는 '25 3/4'
        m = re.search(r"(\d{2,4})\s*(\d)\s*/\s*4", s)
        if m:
            y = int(m.group(1))
            if y < 100:
                y += 2000
            q = int(m.group(2))
            return (y, q)
        # 3) '2025 3Q' 또는 '25 Q3'
        m = re.search(r"(\d{2,4})\s*[Qq]?\s*(\d)\s*[Qq]", s)
        if m:
            y = int(m.group(1))
            if y < 100:
                y += 2000
            q = int(m.group(2))
            return (y, q)
        # 4) 'Q3 2025' 등 역순
        m = re.search(r"[Qq]\s*(\d)\s*(\d{2,4})", s)
        if m:
            q = int(m.group(1))
            y = int(m.group(2))
            if y < 100:
                y += 2000
            return (y, q)
        return None

    def _summary_table_labels(self, report_id: str | None = None) -> tuple[list[str], list[str], list[str]]:
        """요약 표 라벨(증감률/지수/율)을 현재 연·분기에 맞게 구성"""
        age_label = "15-29세"
        if report_id == "employment" or report_id == "summary_employment":
            age_label = "20-29세"
        elif report_id == "unemployment":
            age_label = "15-29세"

        if not self._year_quarter:
            growth_cols = ["{Y-2}. {Q}/4", "{Y-1}. {Q}/4", "{Y}. {Q-1}/4", "{Y}. {Q}/4"]
            index_cols = ["{Y-1}. {Q}/4", "{Y}. {Q}/4"]
            rate_cols = ["{Y-1}. {Q}/4", "{Y}. {Q}/4", age_label]
            return growth_cols, index_cols, rate_cols
        year, quarter = self._year_quarter
        if quarter <= 1:
            prev_q_year, prev_q = year - 1, 4
        else:
            prev_q_year, prev_q = year, quarter - 1
        growth_cols = [
            f"{year-2}. {quarter}/4",
            f"{year-1}. {quarter}/4",
            f"{prev_q_year}. {prev_q}/4",
            f"{year}. {quarter}/4",
        ]
        index_cols = [f"{year-1}. {quarter}/4", f"{year}. {quarter}/4"]
        rate_cols = [f"{year-1}. {quarter}/4", f"{year}. {quarter}/4", age_label]
        return growth_cols, index_cols, rate_cols
    
    def extract_data(self, report_id: str) -> dict[str, Any]:
        """
        특정 보도자료의 데이터 추출
        Args:
            report_id: 보도자료 ID
        Returns:
            추출된 데이터 딕셔너리
        """
        # REPORT_ORDER(부문별+요약)에서 report_id로 검색
        all_reports = [*REPORT_ORDER]
        config = next((r for r in all_reports if r.get('id') == report_id), None)
        if not config:
            raise ValueError(f"알 수 없는 보도자료 ID: {report_id}")
        module = self._load_module(config['generator'])
        # 함수 기반 Generator (고용률, 실업률 등)
        if config.get('uses_functions'):
            return self._extract_with_functions(module, config)
        # 클래스 기반 Generator
        return self._extract_with_class(module, config)
    
    def _extract_with_class(self, module: Any, config: dict[str, Any]) -> dict[str, Any]:
        """클래스 기반 Generator로 데이터 추출"""
        class_name = config.get('class_name')
        generator_class = None
        for name in dir(module):
            obj = getattr(module, name)
            if isinstance(obj, type):
                if class_name and name == class_name:
                    generator_class = obj
                    break
                elif name.endswith('Generator'):
                    generator_class = obj
        if not generator_class:
            raise ValueError(f"Generator 클래스를 찾을 수 없습니다: {config['generator']}")
        # 연/분기 추정치가 있으면 전달하여 헤더 탐색 실패를 방지
        if self._year_quarter:
            y, q = self._year_quarter
            generator = generator_class(str(self.excel_path), year=y, quarter=q)
        else:
            generator = generator_class(str(self.excel_path))
        return generator.extract_all_data()
    
    def _extract_with_functions(self, module: Any, config: dict[str, Any]) -> dict[str, Any]:
        """함수 기반 Generator로 데이터 추출"""
        data: dict[str, Any] = {}
        
        # load_data 함수로 데이터 로드
        if hasattr(module, 'load_data'):
            df_analysis, df_index = module.load_data(str(self.excel_path))
            
            # 전국 데이터
            if hasattr(module, 'get_nationwide_data'):
                data['nationwide_data'] = module.get_nationwide_data(df_analysis, df_index)
            
            # 지역 데이터
            if hasattr(module, 'get_regional_data'):
                data['regional_data'] = module.get_regional_data(df_analysis, df_index)
            
            # 요약 박스 데이터
            if hasattr(module, 'get_summary_box_data') and 'regional_data' in data:
                data['summary_box'] = module.get_summary_box_data(data['regional_data'])
            
            # 표 데이터
            if hasattr(module, 'get_table_data'):
                growth_cols, index_cols, rate_cols = self._summary_table_labels(config.get('id'))
                data['summary_table'] = {
                    'columns': {
                        'change_columns': growth_cols,
                        'rate_columns': rate_cols
                    },
                    'regions': module.get_table_data(df_analysis, df_index)
                }
            
            # Top3 증가/감소 지역
            if 'regional_data' in data:
                top3_increase = []
                for r in data['regional_data'].get('increase_regions', [])[:3]:
                    top3_increase.append({
                        'region': r['region'],
                        'change': r.get('change', 0),
                        'age_groups': r.get('top_age_groups', [])
                    })
                
                top3_decrease = []
                for r in data['regional_data'].get('decrease_regions', [])[:3]:
                    top3_decrease.append({
                        'region': r['region'],
                        'change': r.get('change', 0),
                        'age_groups': r.get('top_age_groups', [])
                    })
                
                data['top3_increase_regions'] = top3_increase
                data['top3_decrease_regions'] = top3_decrease
        
        return data
    
    def generate_html(self, report_id: str, custom_data: dict[str, Any] | None = None) -> str:
        """
        보도자료 HTML 생성
        Args:
            report_id: 보도자료 ID
            custom_data: 사용자 정의 데이터 (결측치 대체용)
        Returns:
            생성된 HTML 문자열
        """
        all_reports = [*REPORT_ORDER]
        config = next((r for r in all_reports if r.get('id') == report_id), None)
        if not config:
            raise ValueError(f"알 수 없는 보도자료 ID: {report_id}")
        if config.get('generator') is None:
            year = None
            quarter = None
            if self._year_quarter:
                year, quarter = self._year_quarter
            html, error, _ = generate_report_html(
                str(self.excel_path),
                config,
                year,
                quarter,
                custom_data=custom_data
            )
            if error:
                raise ValueError(error)
            if html is None:
                raise ValueError("요약 보도자료 HTML 생성 결과가 None입니다")
            return html
        # 데이터 추출
        data = self.extract_data(report_id)
        # 커스텀 데이터 병합
        if custom_data:
            data = self._merge_custom_data(data, custom_data)
        data['get_terms'] = get_terms
        data['get_comparative_terms'] = get_comparative_terms
        # 템플릿 렌더링
        template_path = TEMPLATES_DIR / config['template']
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        # 템플릿 필터 등록 (템플릿에서 사용되는 커스텀 필터들)
        try:
            template.environment.filters['format_value'] = format_value
        except Exception:
            pass
        try:
            template.environment.filters['is_missing'] = is_missing
        except Exception:
            pass
        try:
            template.environment.filters['josa'] = get_josa
        except Exception:
            pass
        return template.render(**data)
    
    def _merge_custom_data(self, data: dict[str, Any], custom_data: dict[str, Any]) -> dict[str, Any]:
        """커스텀 데이터를 기존 데이터에 병합"""
        for key, value in custom_data.items():
            keys = key.split('.')
            obj = data
            for k in keys[:-1]:
                if '[' in k:
                    name, idx = k.replace(']', '').split('[')
                    obj = obj[name][int(idx)]
                else:
                    obj = obj.get(k, {})
            
            final_key = keys[-1]
            if '[' in final_key:
                name, idx = final_key.replace(']', '').split('[')
                if name in obj and isinstance(obj[name], list):
                    obj[name][int(idx)] = value
            else:
                obj[final_key] = value
        
        return data
    
    def check_missing_data(self, data: dict[str, Any]) -> list[str]:
        """데이터에서 결측치 확인"""
        missing_fields = []
        
        def traverse(obj, path=''):
            if obj is None:
                missing_fields.append(path)
            elif isinstance(obj, dict):
                for key, value in obj.items():
                    new_path = f"{path}.{key}" if path else key
                    traverse(value, new_path)
            elif isinstance(obj, list):
                for idx, item in enumerate(obj):
                    new_path = f"{path}[{idx}]"
                    traverse(item, new_path)
            elif isinstance(obj, float) and pd.isna(obj):
                missing_fields.append(path)
            elif obj == '':
                missing_fields.append(path)
        
        traverse(data)
        return missing_fields
    
    def save_report(self, report_id: str, output_path: str = None, custom_data: dict = None) -> str:
        """
        보도자료를 파일로 저장
        Args:
            report_id: 보도자료 ID
            output_path: 출력 파일 경로 (미지정 시 기본 경로)
            custom_data: 사용자 정의 데이터
        Returns:
            저장된 파일 경로
        """
        all_reports = [*REPORT_ORDER]
        config = next((r for r in all_reports if r.get('id') == report_id), None)
        if not config:
            raise ValueError(f"알 수 없는 보도자료 ID: {report_id}")
        if output_path is None:
            TEMP_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_path = TEMP_OUTPUT_DIR / f"{config['name']}_output.html"
        html = self.generate_html(report_id, custom_data)
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html)
        return str(output_path)
    
    def generate_all(self, custom_data_by_report: dict = None) -> dict:
        """
        모든 보도자료 생성
        Args:
            custom_data_by_report: 보도자료별 커스텀 데이터
        Returns:
            생성 결과 딕셔너리
        """
        results = {
            'success': [],
            'errors': []
        }
        custom_data_by_report = custom_data_by_report or {}
        year = None
        quarter = None
        if self._year_quarter:
            year, quarter = self._year_quarter

        # 1) 부문별 보도자료 먼저 생성
        for report in SECTOR_REPORTS:
            report_id = report.get('id')
            try:
                custom_data = custom_data_by_report.get(report_id, {})
                output_path = self.save_report(report_id, custom_data=custom_data)
                results['success'].append({
                    'report_id': report_id,
                    'name': report.get('name'),
                    'path': output_path
                })
            except Exception as e:
                results['errors'].append({
                    'report_id': report_id,
                    'name': report.get('name'),
                    'error': str(e)
                })

        # 2) 시도별 보도자료 생성 (부문별 결과를 활용)
        TEMP_REGIONAL_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_dir = TEMP_REGIONAL_OUTPUT_DIR
        for region in REGIONAL_REPORTS:
            region_name = region.get('name', region.get('id', 'Unknown'))
            region_id = region.get('id', 'Unknown')
            try:
                html_content, error = generate_regional_report_html(
                    str(self.excel_path),
                    region_name,
                    is_reference=False,
                    year=year,
                    quarter=quarter
                )
                if error:
                    results['errors'].append({
                        'report_id': region_id,
                        'name': f"시도별-{region_name}",
                        'error': str(error)
                    })
                    continue
                if html_content is None:
                    results['errors'].append({
                        'report_id': region_id,
                        'name': f"시도별-{region_name}",
                        'error': 'HTML 내용이 None입니다'
                    })
                    continue

                region_name_safe = region_name.replace('/', '_').replace('\\', '_').replace('..', '_')
                output_path = output_dir / f"{region_name_safe}_output.html"
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write(html_content if html_content else '<!-- Empty content -->')
                results['success'].append({
                    'report_id': region_id,
                    'name': f"시도별-{region_name}",
                    'path': str(output_path)
                })
            except Exception as e:
                results['errors'].append({
                    'report_id': region_id,
                    'name': f"시도별-{region_name}",
                    'error': str(e)
                })

        # 3) 요약 보도자료 생성 (마지막)
        for report in SUMMARY_REPORTS:
            report_id = report.get('id')
            try:
                custom_data = custom_data_by_report.get(report_id, {})
                output_path = self.save_report(report_id, custom_data=custom_data)
                results['success'].append({
                    'report_id': report_id,
                    'name': report.get('name'),
                    'path': output_path
                })
            except Exception as e:
                results['errors'].append({
                    'report_id': report_id,
                    'name': report.get('name'),
                    'error': str(e)
                })
        return results


def main():
    """CLI 실행"""
    import argparse
    
    parser = argparse.ArgumentParser(description='통합 보도자료 생성기')
    parser.add_argument('--excel', '-e', required=True, help='엑셀 파일 경로')
    parser.add_argument('--report', '-r', help='생성할 보도자료 ID (미지정 시 전체)')
    parser.add_argument('--output', '-o', help='출력 파일 경로')
    parser.add_argument('--list', '-l', action='store_true', help='사용 가능한 보도자료 목록')
    
    args = parser.parse_args()
    
    if args.list:
        print("사용 가능한 보도자료:")
        for report_id, config in ReportGenerator.REPORT_CONFIGS.items():
            print(f"  {report_id}: {config['name']} ({config['sheet']})")
        return
    
    generator = ReportGenerator(args.excel)
    
    if args.report:
        output_path = generator.save_report(args.report, args.output)
        print(f"보도자료 생성 완료: {output_path}")
    else:
        results = generator.generate_all()
        print(f"\n성공: {len(results['success'])}개")
        for r in results['success']:
            print(f"  ✓ {r['name']}: {r['path']}")
        
        if results['errors']:
            print(f"\n실패: {len(results['errors'])}개")
            for r in results['errors']:
                print(f"  ✕ {r['name']}: {r['error']}")


if __name__ == '__main__':
    main()

