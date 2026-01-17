#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
통합 보도자료 생성기
모든 개별 Generator들을 통합하여 관리하고, 데이터 추출 및 HTML 생성을 처리합니다.
"""

import sys
import json
import importlib.util
from pathlib import Path
from jinja2 import Template
import pandas as pd

BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / 'templates'


class ReportGenerator:
    """통합 보도자료 생성기"""
    
    # 보도자료 설정
    REPORT_CONFIGS = {
        'manufacturing': {
            'name': '광공업생산',
            'generator': 'mining_manufacturing_generator.py',
            'template': 'mining_manufacturing_template.html',
            'sheet': 'A 분석',
            'class_name': 'MiningManufacturingGenerator'
        },
        'service': {
            'name': '서비스업생산',
            'generator': 'service_industry_generator.py',
            'template': 'service_industry_template.html',
            'sheet': 'B 분석',
            'class_name': 'ServiceIndustryGenerator'
        },
        'consumption': {
            'name': '소비동향',
            'generator': 'consumption_generator.py',
            'template': 'consumption_template.html',
            'sheet': 'C 분석',
            'class_name': 'ConsumptionGenerator'
        },
        'employment': {
            'name': '고용률',
            'generator': 'employment_rate_generator.py',
            'template': 'employment_rate_template.html',
            'sheet': 'D(고용률)분석',
            'uses_functions': True
        },
        'unemployment': {
            'name': '실업률',
            'generator': 'unemployment_generator.py',
            'template': 'unemployment_template.html',
            'sheet': 'D(실업)분석',
            'uses_functions': True
        },
        'price': {
            'name': '물가동향',
            'generator': 'price_trend_generator.py',
            'template': 'price_trend_template.html',
            'sheet': 'E(품목성질물가)분석',
            'class_name': 'PriceTrendGenerator'
        },
        'export': {
            'name': '수출',
            'generator': 'export_generator.py',
            'template': 'export_template.html',
            'sheet': 'G 분석',
            'class_name': 'ExportGenerator'
        },
        'import': {
            'name': '수입',
            'generator': 'import_generator.py',
            'template': 'import_template.html',
            'sheet': 'H 분석',
            'class_name': 'ImportGenerator'
        },
        'migration': {
            'name': '국내인구이동',
            'generator': 'domestic_migration_generator.py',
            'template': 'domestic_migration_template.html',
            'sheet': 'I(순인구이동)집계',
            'class_name': 'DomesticMigrationGenerator'
        }
    }
    
    def __init__(self, excel_path: str):
        """
        초기화
        
        Args:
            excel_path: 엑셀 파일 경로
        """
        self.excel_path = Path(excel_path)
        if not self.excel_path.exists():
            raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {excel_path}")
        
        self._modules = {}
    
    def _load_module(self, generator_name: str):
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
    
    def extract_data(self, report_id: str) -> dict:
        """
        특정 보도자료의 데이터 추출
        
        Args:
            report_id: 보도자료 ID
            
        Returns:
            추출된 데이터 딕셔너리
        """
        if report_id not in self.REPORT_CONFIGS:
            raise ValueError(f"알 수 없는 보도자료 ID: {report_id}")
        
        config = self.REPORT_CONFIGS[report_id]
        module = self._load_module(config['generator'])
        
        # 함수 기반 Generator (고용률, 실업률)
        if config.get('uses_functions'):
            return self._extract_with_functions(module, config)
        
        # 클래스 기반 Generator
        return self._extract_with_class(module, config)
    
    def _extract_with_class(self, module, config: dict) -> dict:
        """클래스 기반 Generator로 데이터 추출"""
        class_name = config.get('class_name')
        
        # 클래스 찾기
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
        
        generator = generator_class(str(self.excel_path))
        return generator.extract_all_data()
    
    def _extract_with_functions(self, module, config: dict) -> dict:
        """함수 기반 Generator로 데이터 추출"""
        data = {}
        
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
                data['summary_table'] = {
                    'columns': {
                        'change_columns': ['2023.2/4', '2024.2/4', '2025.1/4', '2025.2/4'],
                        'rate_columns': ['2024.2/4', '2025.2/4', '20-29세']
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
    
    def generate_html(self, report_id: str, custom_data: dict = None) -> str:
        """
        보도자료 HTML 생성
        
        Args:
            report_id: 보도자료 ID
            custom_data: 사용자 정의 데이터 (결측치 대체용)
            
        Returns:
            생성된 HTML 문자열
        """
        if report_id not in self.REPORT_CONFIGS:
            raise ValueError(f"알 수 없는 보도자료 ID: {report_id}")
        
        config = self.REPORT_CONFIGS[report_id]
        
        # 데이터 추출
        data = self.extract_data(report_id)
        
        # 커스텀 데이터 병합
        if custom_data:
            data = self._merge_custom_data(data, custom_data)
        
        # 템플릿 렌더링
        template_path = TEMPLATES_DIR / config['template']
        with open(template_path, 'r', encoding='utf-8') as f:
            template = Template(f.read())
        
        return template.render(**data)
    
    def _merge_custom_data(self, data: dict, custom_data: dict) -> dict:
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
    
    def check_missing_data(self, data: dict) -> list:
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
        config = self.REPORT_CONFIGS[report_id]
        
        if output_path is None:
            output_path = TEMPLATES_DIR / f"{config['name']}_output.html"
        
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
        
        for report_id in self.REPORT_CONFIGS:
            try:
                custom_data = custom_data_by_report.get(report_id, {})
                output_path = self.save_report(report_id, custom_data=custom_data)
                results['success'].append({
                    'report_id': report_id,
                    'name': self.REPORT_CONFIGS[report_id]['name'],
                    'path': output_path
                })
            except Exception as e:
                results['errors'].append({
                    'report_id': report_id,
                    'name': self.REPORT_CONFIGS[report_id]['name'],
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

