#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
국내인구이동 보도자료 생성기
I(순인구이동)집계 및 I 참고 시트에서 데이터를 추출하여 HTML 보도자료를 생성합니다.
"""

import pandas as pd
import json
from jinja2 import Environment, FileSystemLoader
import os

# 시도 순서
SIDO_ORDER = [
    "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종",
    "경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"
]

# 권역별 그룹
REGION_GROUPS = {
    "경인": ["서울", "인천", "경기"],
    "충청": ["대전", "세종", "충북", "충남"],
    "호남": ["광주", "전북", "전남", "제주"],
    "동북": ["대구", "경북", "강원"],
    "동남": ["부산", "울산", "경남"]
}


def safe_float(value, default=0.0):
    """안전하게 float로 변환합니다."""
    try:
        if pd.isna(value):
            return default
        return float(value)
    except (ValueError, TypeError):
        return default


def normalize_age_group(age_str):
    """연령대 표기를 정규화합니다.
    
    예: '00~04세' → '0~4세', '05~09세' → '5~9세'
    """
    if not age_str or not isinstance(age_str, str):
        return age_str
    
    import re
    # '00~04세' 패턴 매칭
    match = re.match(r'^(\d+)~(\d+)세$', age_str)
    if match:
        start = int(match.group(1))  # 앞의 0 제거
        end = int(match.group(2))    # 앞의 0 제거
        return f"{start}~{end}세"
    
    return age_str


def load_data(excel_path):
    """엑셀 파일에서 데이터를 로드합니다."""
    xl = pd.ExcelFile(excel_path)
    sheet_names = xl.sheet_names
    
    # 집계 시트 찾기
    summary_sheet = None
    for name in ['I(순인구이동)집계', 'I(순인구이동) 집계', '시도 간 이동']:
        if name in sheet_names:
            summary_sheet = name
            if name == '시도 간 이동':
                print(f"[시트 대체] 'I(순인구이동)집계' → '시도 간 이동' (기초자료)")
            break
    
    if not summary_sheet:
        raise ValueError(f"국내인구이동 집계 시트를 찾을 수 없습니다. 시트 목록: {sheet_names}")
    
    # 참고 시트 찾기 (없으면 집계 시트 사용)
    reference_sheet = None
    for name in ['I 참고', 'I참고']:
        if name in sheet_names:
            reference_sheet = name
            break
    if not reference_sheet:
        reference_sheet = summary_sheet
    
    summary_df = pd.read_excel(excel_path, sheet_name=summary_sheet, header=None)
    reference_df = pd.read_excel(excel_path, sheet_name=reference_sheet, header=None)
    return summary_df, reference_df

def get_sido_data(summary_df):
    """시도별 순인구이동 데이터를 추출합니다."""
    sido_data = {}
    
    for i in range(3, len(summary_df)):
        row = summary_df.iloc[i]
        age_group = row[7]
        
        if age_group == '합계':
            sido = row[4]
            if sido in SIDO_ORDER:
                net_migration_2025_24 = safe_float(row[25], 0)  # 2025.2/4 순이동
                net_migration_2025_14 = safe_float(row[24], 0)  # 2025.1/4
                net_migration_2024_24 = safe_float(row[21], 0)  # 2024.2/4
                net_migration_2023_24 = safe_float(row[17], 0)  # 2023.2/4
                
                sido_data[sido] = {
                    'net_migration': net_migration_2025_24,
                    'net_migrations': [net_migration_2023_24, net_migration_2024_24, net_migration_2025_14, net_migration_2025_24]
                }
    
    return sido_data

def get_sido_age_data(summary_df):
    """시도별 연령별 순인구이동 데이터를 추출합니다."""
    sido_age_data = {}
    current_sido = None
    
    for i in range(3, len(summary_df)):
        row = summary_df.iloc[i]
        sido = row[4]
        age_group = row[7]
        rank = row[6]
        level = row[5]
        net_migration = safe_float(row[25], 0)  # 2025.2/4 순이동
        
        if sido in SIDO_ORDER and age_group == '합계':
            current_sido = sido
            sido_age_data[current_sido] = {
                'total': net_migration,
                'ages': [],
                'age_20_29': 0,
                'age_other': 0
            }
        
        # rank 또는 level로 연령별 데이터 판별 (rank가 없는 경우 level=1로 판별)
        is_age_row = (pd.notna(rank) and age_group != '합계') or (current_sido and str(level) == '1' and age_group != '합계')
        
        if current_sido and is_age_row:
            try:
                rank_num = int(rank) if pd.notna(rank) else i  # rank가 없으면 행 인덱스 사용
                # 연령대 표기 정규화 (00~04세 → 0~4세)
                normalized_age = normalize_age_group(age_group)
                sido_age_data[current_sido]['ages'].append({
                    'rank': rank_num,
                    'name': normalized_age,
                    'value': net_migration
                })
                
                # 20~29세 합계 계산
                if age_group in ['20~24세', '25~29세']:
                    sido_age_data[current_sido]['age_20_29'] += net_migration
                else:
                    sido_age_data[current_sido]['age_other'] += net_migration
                    
            except (ValueError, TypeError):
                pass
    
    # 순위순 정렬
    for sido in sido_age_data:
        sido_age_data[sido]['ages'].sort(key=lambda x: x['rank'])
    
    return sido_age_data

def get_regional_data(sido_data, sido_age_data):
    """순유입/순유출 지역을 분류합니다."""
    inflow_regions = []
    outflow_regions = []
    
    for sido in SIDO_ORDER:
        sido_info = sido_data.get(sido, {})
        age_info = sido_age_data.get(sido, {})
        net_migration = sido_info.get('net_migration', 0)
        
        if pd.isna(net_migration):
            continue
        
        # 상위 3개 연령대
        ages = age_info.get('ages', [])[:3]
        
        region_info = {
            'name': sido,
            'net_migration': net_migration,
            'ages': ages
        }
        
        if net_migration > 0:
            inflow_regions.append(region_info)
        elif net_migration < 0:
            outflow_regions.append(region_info)
    
    # 정렬 (유입은 높은 순, 유출은 낮은 순)
    inflow_regions.sort(key=lambda x: -x['net_migration'])
    outflow_regions.sort(key=lambda x: x['net_migration'])
    
    return inflow_regions, outflow_regions

def generate_summary_box(inflow_regions, outflow_regions):
    """요약 박스 텍스트를 생성합니다."""
    inflow_count = len(inflow_regions)
    outflow_count = len(outflow_regions)
    
    # 상위 3개 지역
    top_inflow = inflow_regions[:3] if inflow_regions else []
    top_outflow = outflow_regions[:3] if outflow_regions else []
    
    # 헤드라인
    inflow_names = ", ".join([f"<span class='bold highlight'>{r['name']}</span>" for r in top_inflow[:2]])
    outflow_names = ", ".join([f"<span class='bold highlight'>{r['name']}</span>" for r in top_outflow[:2]])
    
    headline = f"◆ 국내 인구이동은 {inflow_names} 등 <span class='bold'>{inflow_count}개</span> 시도 인구 <span class='bold'>순유입</span>, {outflow_names} 등 <span class='bold'>{outflow_count}개</span> 시도 인구 <span class='bold'>순유출</span>"
    
    # 유입 요약
    inflow_str = ", ".join([f"<span class='bold'>{r['name']}</span>({r['net_migration']:,.0f}명)" for r in top_inflow])
    inflow_summary = f"국내 인구 <span class='bold'>순유입지역</span>은 {inflow_str} 등 {inflow_count}개 시도임"
    
    # 유출 요약
    outflow_str = ", ".join([f"<span class='bold'>{r['name']}</span>({r['net_migration']:,.0f}명)" for r in top_outflow])
    outflow_summary = f"국내 인구 <span class='bold'>순유출지역</span>은 {outflow_str} 등 {outflow_count}개 시도임"
    
    return {
        'headline': headline,
        'inflow_summary': inflow_summary,
        'outflow_summary': outflow_summary
    }

def generate_summary_table(sido_data, sido_age_data):
    """요약 테이블 데이터를 생성합니다."""
    rows = []
    
    # 권역별 시도
    region_group_order = ['경인', '충청', '호남', '동북', '동남']
    
    for group_name in region_group_order:
        sidos = REGION_GROUPS[group_name]
        for idx, sido in enumerate(sidos):
            sido_info = sido_data.get(sido, {})
            age_info = sido_age_data.get(sido, {})
            
            # 순이동 데이터 (천명 단위로 변환)
            net_migrations = sido_info.get('net_migrations', [None, None, None, None])
            net_migrations_k = [v / 1000 if v is not None else None for v in net_migrations]
            
            # 20~29세 및 그 외 (천명 단위)
            age_20_29 = age_info.get('age_20_29', 0) / 1000
            age_other = age_info.get('age_other', 0) / 1000
            
            row_data = {
                'sido': ' '.join(sido) if len(sido) == 2 else sido,
                'net_migrations': net_migrations_k,
                'age_20_29': age_20_29,
                'age_other': age_other
            }
            
            # 첫 번째 시도에만 region_group과 rowspan 추가
            if idx == 0:
                row_data['region_group'] = group_name
                row_data['rowspan'] = len(sidos)
            else:
                row_data['region_group'] = None
            
            rows.append(row_data)
    
    return {
        'rows': rows,
        'columns': {
            'current_quarter': '2025.2/4',
            'quarter_columns': ['2023.2/4', '2024.2/4', '2025.1/4', '2025.2/4']
        }
    }

class DomesticMigrationGenerator:
    """국내인구이동 보도자료 생성 클래스"""
    
    def __init__(self, excel_path: str):
        """
        초기화
        
        Args:
            excel_path: 엑셀 파일 경로
        """
        self.excel_path = excel_path
        self.summary_df = None
        self.reference_df = None
        self.sido_data = None
        self.sido_age_data = None
    
    def load_data(self):
        """엑셀 파일에서 데이터를 로드합니다."""
        self.summary_df, self.reference_df = load_data(self.excel_path)
    
    def extract_all_data(self) -> dict:
        """모든 데이터 추출"""
        self.load_data()
        
        # 시도별 데이터 추출
        self.sido_data = get_sido_data(self.summary_df)
        self.sido_age_data = get_sido_age_data(self.summary_df)
        
        # 시도별 분류
        inflow_regions, outflow_regions = get_regional_data(self.sido_data, self.sido_age_data)
        
        # Top 3 지역
        top3_inflow = inflow_regions[:3]
        top3_outflow = outflow_regions[:3]
        
        # 요약 박스
        summary_box = generate_summary_box(inflow_regions, outflow_regions)
        
        # 요약 테이블
        summary_table = generate_summary_table(self.sido_data, self.sido_age_data)
        
        return {
            'report_info': {
                'main_section_number': '7',
                'main_section_title': '국내 인구이동',
                'page_number': '- 15 -'
            },
            'inflow_regions': inflow_regions,
            'outflow_regions': outflow_regions,
            'summary_box': summary_box,
            'top3_inflow_regions': top3_inflow,
            'top3_outflow_regions': top3_outflow,
            'summary_table': summary_table
        }


def generate_report_data(excel_path, raw_excel_path=None, year=None, quarter=None):
    """보도자료 데이터를 생성합니다. (하위 호환성 유지)
    
    Args:
        excel_path: 분석표 엑셀 파일 경로
        raw_excel_path: 기초자료 엑셀 파일 경로 (선택사항, 향후 기초자료 직접 추출 지원 예정)
        year: 현재 연도 (선택사항)
        quarter: 현재 분기 (선택사항)
    """
    generator = DomesticMigrationGenerator(excel_path)
    return generator.extract_all_data()

def render_template(data, template_path, output_path):
    """Jinja2 템플릿을 렌더링합니다."""
    template_dir = os.path.dirname(template_path)
    template_name = os.path.basename(template_path)
    
    env = Environment(loader=FileSystemLoader(template_dir))
    template = env.get_template(template_name)
    
    html_content = template.render(data=data)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"HTML 보도자료가 생성되었습니다: {output_path}")

def main():
    # 경로 설정
    base_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(base_dir, '..', '분석표_25년 2분기_캡스톤.xlsx')
    template_path = os.path.join(base_dir, 'domestic_migration_template.html')
    output_path = os.path.join(base_dir, 'domestic_migration_output.html')
    json_output_path = os.path.join(base_dir, 'domestic_migration_data.json')
    
    # 데이터 생성
    print("데이터 추출 중...")
    data = generate_report_data(excel_path)
    
    # JSON 저장
    with open(json_output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"JSON 데이터가 저장되었습니다: {json_output_path}")
    
    # HTML 렌더링
    print("HTML 렌더링 중...")
    render_template(data, template_path, output_path)
    
    # 결과 요약 출력
    print(f"\n=== 데이터 요약 ===")
    print(f"순유입 지역 수: {len(data['inflow_regions'])}")
    print(f"순유출 지역 수: {len(data['outflow_regions'])}")
    
    print("\n=== Top 3 순유입 지역 ===")
    for region in data['top3_inflow_regions']:
        ages_str = ", ".join([f"{a['name']}({a['value']:,.0f}명)" for a in region['ages'][:3]]) if region['ages'] else "N/A"
        print(f"  {region['name']}({region['net_migration']:,.0f}명): {ages_str}")
    
    print("\n=== Top 3 순유출 지역 ===")
    for region in data['top3_outflow_regions']:
        ages_str = ", ".join([f"{a['name']}({a['value']:,.0f}명)" for a in region['ages'][:3]]) if region['ages'] else "N/A"
        print(f"  {region['name']}({region['net_migration']:,.0f}명): {ages_str}")

if __name__ == '__main__':
    main()

