# -*- coding: utf-8 -*-
"""
GRDP 관련 서비스 함수
"""

import json
import pandas as pd
from pathlib import Path


# 기본 기여율 JSON 경로
DEFAULT_CONTRIBUTIONS_PATH = Path(__file__).parent.parent / 'templates' / 'default_contributions.json'


def load_default_contributions():
    """기본 기여율 데이터 로드"""
    try:
        if DEFAULT_CONTRIBUTIONS_PATH.exists():
            with open(DEFAULT_CONTRIBUTIONS_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"[GRDP] 기본 기여율 로드 실패: {e}")
    return None


def save_extracted_contributions(grdp_data):
    """추출된 기여율을 JSON으로 저장 (향후 기본값으로 사용)"""
    try:
        if not grdp_data or grdp_data.get('data_missing'):
            return False
        
        contributions_data = {
            "_comment": "분석표/KOSIS에서 추출한 실제 기여율 데이터",
            "_last_updated": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
            "_source": grdp_data.get('source', 'unknown'),
            
            "national": {
                "growth_rate": grdp_data.get('national_summary', {}).get('growth_rate', 0.0),
                "contributions": grdp_data.get('national_summary', {}).get('contributions', {}),
                "is_placeholder": False
            },
            
            "regional": {}
        }
        
        for region_data in grdp_data.get('regional_data', []):
            region = region_data.get('region')
            if region and region != '전국':
                contributions_data['regional'][region] = {
                    "growth_rate": region_data.get('growth_rate', 0.0),
                    "manufacturing": region_data.get('manufacturing', 0.0),
                    "construction": region_data.get('construction', 0.0),
                    "service": region_data.get('service', 0.0),
                    "other": region_data.get('other', 0.0),
                    "is_placeholder": region_data.get('placeholder', False)
                }
        
        # 저장 경로
        save_path = Path(__file__).parent.parent / 'templates' / 'extracted_contributions.json'
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(contributions_data, f, ensure_ascii=False, indent=2)
        
        print(f"[GRDP] 기여율 데이터 저장 완료: {save_path}")
        return True
    except Exception as e:
        print(f"[GRDP] 기여율 저장 실패: {e}")
        return False


def get_kosis_grdp_download_info():
    """KOSIS GRDP 다운로드 정보 반환"""
    return {
        'url': 'https://kosis.kr/statisticsList/experimentStatistical.do#down23',
        'name': '분기 지역내총생산',
        'source': 'KOSIS 실험적통계',
        'description': '국가데이터처에서 제공하는 분기별 지역내총생산(GRDP) 통계입니다.',
        'instruction': '''GRDP 데이터가 누락되었습니다. 다음 단계를 따라 데이터를 추가하세요:
1. KOSIS 실험적통계 페이지(https://kosis.kr/statisticsList/experimentStatistical.do#down23)에 접속
2. "분기 지역내총생산" 항목의 "다운로드" 버튼 클릭
3. 다운로드한 엑셀 파일을 기초자료 수집표와 함께 업로드하거나,
   기초자료 수집표의 "분기 GRDP" 시트에 데이터를 추가'''
    }


def check_grdp_in_raw_data(raw_excel_path):
    """기초자료 수집표에 GRDP 시트가 있는지 확인"""
    try:
        xl = pd.ExcelFile(raw_excel_path)
        grdp_sheets = [s for s in xl.sheet_names if 'GRDP' in s.upper() or '지역내총생산' in s]
        return len(grdp_sheets) > 0
    except:
        return False


def parse_kosis_grdp_file(file_path, year=None, quarter=None):
    """KOSIS에서 다운로드한 GRDP 엑셀 파일 파싱 (성장률 및 기여도 시트 사용)"""
    try:
        print(f"[KOSIS] GRDP 파일 파싱: {file_path}")
        
        xl = pd.ExcelFile(file_path)
        print(f"[KOSIS] 시트 목록: {xl.sheet_names}")
        
        # 연도/분기 기본값
        target_year = year or 2025
        target_quarter = quarter or 2
        
        regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        region_groups = {
            '서울': '경인', '인천': '경인', '경기': '경인',
            '대전': '충청', '세종': '충청', '충북': '충청', '충남': '충청',
            '광주': '호남', '전북': '호남', '전남': '호남', '제주': '호남',
            '대구': '동북', '경북': '동북', '강원': '동북',
            '부산': '동남', '울산': '동남', '경남': '동남'
        }
        
        # 성장률 시트와 기여도 시트가 있는지 확인
        has_growth_sheet = '성장률' in xl.sheet_names
        has_contrib_sheet = '기여도' in xl.sheet_names
        
        if has_growth_sheet and has_contrib_sheet:
            # 성장률 시트와 기여도 시트에서 직접 값 추출
            result = _parse_grdp_from_sheets(file_path, target_year, target_quarter, regions, region_groups)
            if result:
                return result
        
        # 시트가 없거나 파싱 실패시 실질금액에서 계산
        return _parse_grdp_from_values(file_path, target_year, target_quarter, regions, region_groups)
        
    except Exception as e:
        import traceback
        print(f"[KOSIS] GRDP 파일 파싱 오류: {e}")
        traceback.print_exc()
        return None


def _parse_grdp_from_sheets(file_path, year, quarter, regions, region_groups):
    """성장률 및 기여도 시트에서 직접 데이터 추출"""
    try:
        df_growth = pd.read_excel(file_path, sheet_name='성장률', header=None)
        df_contrib = pd.read_excel(file_path, sheet_name='기여도', header=None)
        
        # 당분기 컬럼 찾기
        header_row = 4
        current_pattern = f"{year}.{quarter}/4"
        
        growth_col = -1
        contrib_col = -1
        
        for col_idx in range(len(df_growth.columns)):
            cell = str(df_growth.iloc[header_row, col_idx]).strip()
            if current_pattern in cell:
                growth_col = col_idx
                break
        
        for col_idx in range(len(df_contrib.columns)):
            cell = str(df_contrib.iloc[header_row, col_idx]).strip()
            if current_pattern in cell:
                contrib_col = col_idx
                break
        
        if growth_col == -1 or contrib_col == -1:
            print(f"[KOSIS] 해당 분기 컬럼을 찾을 수 없음: {current_pattern}")
            return None
        
        print(f"[KOSIS] 성장률 컬럼: {growth_col}, 기여도 컬럼: {contrib_col} ({current_pattern})")
        
        # 지역별 데이터 추출
        regional_data = []
        national_data = None
        
        def safe_float(val, default=0.0):
            try:
                if pd.isna(val):
                    return default
                return round(float(val), 1)
            except:
                return default
        
        current_region = None
        region_values = {}
        
        for i in range(5, len(df_growth)):
            region_cell = df_growth.iloc[i, 1]
            item_cell = str(df_growth.iloc[i, 2]).strip() if pd.notna(df_growth.iloc[i, 2]) else ''
            
            # 새 지역 시작
            if pd.notna(region_cell):
                region_name = str(region_cell).strip()
                if region_name in regions:
                    current_region = region_name
                    region_values[current_region] = {'growth_rate': 0.0, 'manufacturing': 0.0, 
                                                     'construction': 0.0, 'service': 0.0, 'other': 0.0}
            
            if current_region is None or current_region not in region_values:
                continue
            
            # 성장률 시트에서 총 성장률 가져오기
            if '지역내총생산' in item_cell or '시장가격' in item_cell:
                region_values[current_region]['growth_rate'] = safe_float(df_growth.iloc[i, growth_col])
        
        # 기여도 시트에서 산업별 기여도 추출
        current_region = None
        for i in range(5, len(df_contrib)):
            region_cell = df_contrib.iloc[i, 1]
            item_cell = str(df_contrib.iloc[i, 2]).strip() if pd.notna(df_contrib.iloc[i, 2]) else ''
            
            # 새 지역 시작
            if pd.notna(region_cell):
                region_name = str(region_cell).strip()
                if region_name in regions:
                    current_region = region_name
            
            if current_region is None or current_region not in region_values:
                continue
            
            val = safe_float(df_contrib.iloc[i, contrib_col])
            
            # 광업, 제조업
            if '광업' in item_cell and '제조업' in item_cell:
                region_values[current_region]['manufacturing'] = val
            # 건설업 (정확히 매칭)
            elif item_cell.strip() == '건설업' or item_cell == ' 건설업':
                region_values[current_region]['construction'] = val
            # 서비스업 (정확히 매칭 - 하위 서비스업 제외)
            elif item_cell.strip() == '서비스업' or item_cell == ' 서비스업':
                region_values[current_region]['service'] = val
            # 기타산업 및 순생산물세
            elif '기타산업' in item_cell or '순생산물세' in item_cell:
                region_values[current_region]['other'] = val
        
        # regional_data 구성
        for region in regions:
            if region not in region_values:
                continue
            
            rv = region_values[region]
            region_entry = {
                'region': region,
                'region_group': region_groups.get(region, ''),
                'growth_rate': rv['growth_rate'],
                'manufacturing': rv['manufacturing'],
                'construction': rv['construction'],
                'service': rv['service'],
                'other': rv['other']
            }
            
            if region == '전국':
                national_data = region_entry.copy()
            
            regional_data.append(region_entry)
        
        if not regional_data:
            return None
        
        # 성장률 1위 지역 찾기
        non_national = [r for r in regional_data if r['region'] != '전국']
        if non_national:
            top_region = max(non_national, key=lambda x: x['growth_rate'])
        else:
            top_region = {'region': '-', 'growth_rate': 0.0, 'manufacturing': 0.0, 
                         'construction': 0.0, 'service': 0.0, 'other': 0.0}
        
        if national_data is None:
            national_data = {'growth_rate': 0.0, 'manufacturing': 0.0, 'construction': 0.0, 
                           'service': 0.0, 'other': 0.0}
        
        result = {
            'report_info': {
                'year': year,
                'quarter': quarter,
                'page_number': ''
            },
            'national_summary': {
                'growth_rate': national_data.get('growth_rate', 0.0),
                'direction': '증가' if national_data.get('growth_rate', 0) >= 0 else '감소',
                'contributions': {
                    'manufacturing': national_data.get('manufacturing', 0.0),
                    'construction': national_data.get('construction', 0.0),
                    'service': national_data.get('service', 0.0),
                    'other': national_data.get('other', 0.0)
                }
            },
            'top_region': {
                'name': top_region.get('region', '-'),
                'growth_rate': top_region.get('growth_rate', 0.0),
                'contributions': {
                    'manufacturing': top_region.get('manufacturing', 0.0),
                    'construction': top_region.get('construction', 0.0),
                    'service': top_region.get('service', 0.0),
                    'other': top_region.get('other', 0.0)
                }
            },
            'regional_data': regional_data,
            'chart_config': {
                'y_axis': {
                    'min': -6,
                    'max': 8,
                    'step': 2
                }
            },
            'source': 'KOSIS 실험적통계'
        }
        
        print(f"[KOSIS] GRDP 시트 파싱 완료")
        print(f"  - 전국: 성장률 {national_data.get('growth_rate', 0)}%, 기여도: 광제조 {national_data.get('manufacturing', 0)}%p, 서비스 {national_data.get('service', 0)}%p")
        print(f"  - 1위: {top_region.get('region', '-')} ({top_region.get('growth_rate', 0)}%)")
        return result
        
    except Exception as e:
        import traceback
        print(f"[KOSIS] 시트 파싱 오류: {e}")
        traceback.print_exc()
        return None


def _parse_grdp_from_values(file_path, year, quarter, regions, region_groups):
    """실질금액 시트에서 성장률 및 기여도 계산"""
    try:
        df = pd.read_excel(file_path, header=None)
        
        # 당분기 및 전년동기 컬럼 찾기
        header_row = 4
        current_pattern = f"{year}.{quarter}/4"
        prev_year_pattern = f"{year - 1}.{quarter}/4"
        
        current_col = -1
        prev_col = -1
        
        for col_idx in range(len(df.columns)):
            cell = str(df.iloc[header_row, col_idx]).strip()
            if current_pattern in cell:
                current_col = col_idx
            elif prev_year_pattern in cell:
                prev_col = col_idx
        
        if current_col == -1:
            current_col = len(df.columns) - 1
            prev_col = current_col - 4
        
        print(f"[KOSIS] 실질금액 계산 - 당분기: {current_col}, 전년동기: {prev_col}")
        
        regional_data = []
        national_data = None
        
        current_region = None
        region_industries = {}
        region_total_prev = {}
        
        for i in range(5, len(df)):
            region_cell = df.iloc[i, 1]
            item_cell = df.iloc[i, 2]
            
            if pd.notna(region_cell):
                region_name = str(region_cell).strip()
                if region_name in regions:
                    current_region = region_name
                    region_industries[current_region] = {}
            
            if current_region is None:
                continue
            
            item = str(item_cell).strip() if pd.notna(item_cell) else ''
            
            try:
                current_val = float(df.iloc[i, current_col])
                prev_val = float(df.iloc[i, prev_col])
            except:
                continue
            
            if '지역내총생산' in item:
                region_total_prev[current_region] = prev_val
                growth = round(((current_val - prev_val) / prev_val) * 100, 1) if prev_val > 0 else 0.0
                region_industries[current_region]['total'] = {'current': current_val, 'prev': prev_val, 'growth': growth}
            elif '광업' in item and '제조업' in item:
                region_industries[current_region]['manufacturing'] = {'current': current_val, 'prev': prev_val}
            elif item.strip() in ['건설업', ' 건설업']:
                region_industries[current_region]['construction'] = {'current': current_val, 'prev': prev_val}
            elif item.strip() in ['서비스업', ' 서비스업']:
                region_industries[current_region]['service'] = {'current': current_val, 'prev': prev_val}
            elif '기타산업' in item:
                region_industries[current_region]['other'] = {'current': current_val, 'prev': prev_val}
        
        for region in regions:
            if region not in region_industries:
                continue
            
            ind = region_industries[region]
            total_prev = region_total_prev.get(region, 0)
            
            if total_prev == 0:
                continue
            
            def calc_contrib(key):
                d = ind.get(key, {})
                if total_prev > 0:
                    return round(((d.get('current', 0) - d.get('prev', 0)) / total_prev) * 100, 1)
                return 0.0
            
            entry = {
                'region': region,
                'region_group': region_groups.get(region, ''),
                'growth_rate': ind.get('total', {}).get('growth', 0.0),
                'manufacturing': calc_contrib('manufacturing'),
                'construction': calc_contrib('construction'),
                'service': calc_contrib('service'),
                'other': calc_contrib('other')
            }
            
            if region == '전국':
                national_data = entry.copy()
            
            regional_data.append(entry)
        
        if not regional_data:
            return None
        
        non_national = [r for r in regional_data if r['region'] != '전국']
        top_region = max(non_national, key=lambda x: x['growth_rate']) if non_national else {}
        
        if national_data is None:
            national_data = {'growth_rate': 0.0, 'manufacturing': 0.0, 'construction': 0.0, 'service': 0.0, 'other': 0.0}
        
        print(f"[KOSIS] GRDP 계산 완료 - 전국: {national_data.get('growth_rate', 0)}%")
        
        return {
            'report_info': {'year': year, 'quarter': quarter, 'page_number': ''},
            'national_summary': {
                'growth_rate': national_data.get('growth_rate', 0.0),
                'direction': '증가' if national_data.get('growth_rate', 0) >= 0 else '감소',
                'contributions': {
                    'manufacturing': national_data.get('manufacturing', 0.0),
                    'construction': national_data.get('construction', 0.0),
                    'service': national_data.get('service', 0.0),
                    'other': national_data.get('other', 0.0)
                }
            },
            'top_region': {
                'name': top_region.get('region', '-'),
                'growth_rate': top_region.get('growth_rate', 0.0),
                'contributions': {
                    'manufacturing': top_region.get('manufacturing', 0.0),
                    'construction': top_region.get('construction', 0.0),
                    'service': top_region.get('service', 0.0),
                    'other': top_region.get('other', 0.0)
                }
            },
            'regional_data': regional_data,
            'chart_config': {'y_axis': {'min': -6, 'max': 8, 'step': 2}},
            'source': 'KOSIS 실험적통계'
        }
    except Exception as e:
        import traceback
        print(f"[KOSIS] 실질금액 파싱 오류: {e}")
        traceback.print_exc()
        return None


def get_default_grdp_data(year, quarter, use_default_contributions=True):
    """기본 GRDP 데이터 (기본 기여율 포함)
    
    Args:
        year: 연도
        quarter: 분기
        use_default_contributions: True면 default_contributions.json의 기본값 사용
    
    Returns:
        GRDP 데이터 딕셔너리 (placeholder 플래그 포함)
    """
    regions = ['전국', '서울', '인천', '경기', '대전', '세종', '충북', '충남',
               '광주', '전북', '전남', '제주', '대구', '경북', '강원', '부산', '울산', '경남']
    
    region_groups = {
        '서울': '경인', '인천': '경인', '경기': '경인',
        '대전': '충청', '세종': '충청', '충북': '충청', '충남': '충청',
        '광주': '호남', '전북': '호남', '전남': '호남', '제주': '호남',
        '대구': '동북', '경북': '동북', '강원': '동북',
        '부산': '동남', '울산': '동남', '경남': '동남'
    }
    
    # 기본 기여율 로드
    default_contributions = None
    if use_default_contributions:
        default_contributions = load_default_contributions()
        
        # extracted_contributions.json이 있으면 우선 사용
        extracted_path = Path(__file__).parent.parent / 'templates' / 'extracted_contributions.json'
        if extracted_path.exists():
            try:
                with open(extracted_path, 'r', encoding='utf-8') as f:
                    extracted = json.load(f)
                    # extracted가 placeholder가 아닌 경우에만 사용
                    if extracted.get('national', {}).get('is_placeholder') == False:
                        default_contributions = extracted
                        print(f"[GRDP] 추출된 기여율 사용: {extracted_path}")
            except:
                pass
    
    regional_data = []
    
    for region in regions:
        if default_contributions and region != '전국':
            region_contrib = default_contributions.get('regional', {}).get(region, {})
            regional_data.append({
                'region': region,
                'region_group': region_groups.get(region, ''),
                'growth_rate': region_contrib.get('growth_rate'),
                'manufacturing': region_contrib.get('manufacturing'),
                'construction': region_contrib.get('construction'),
                'service': region_contrib.get('service'),
                'other': region_contrib.get('other'),
                'placeholder': region_contrib.get('is_placeholder', True),
                'needs_review': region_contrib.get('is_placeholder', True)  # 수정 필요 표시
            })
        else:
            regional_data.append({
                'region': region,
                'region_group': region_groups.get(region, ''),
                'growth_rate': None,
                'manufacturing': None,
                'construction': None,
                'service': None,
                'other': None,
                'placeholder': True,
                'needs_review': True
            })
    
    # 전국 기여율 기본값
    national_growth = None
    national_contributions = {'manufacturing': None, 'construction': None, 'service': None, 'other': None}
    national_placeholder = True
    
    if default_contributions:
        national = default_contributions.get('national', {})
        national_growth = national.get('growth_rate')
        national_contributions = national.get('contributions', national_contributions)
        national_placeholder = national.get('is_placeholder', True)
    
    # 1위 지역 찾기
    non_national = [r for r in regional_data if r['region'] != '전국' and r.get('growth_rate') is not None]
    if non_national:
        top_region = max(non_national, key=lambda x: x.get('growth_rate', 0) or 0)
    else:
        top_region = {'region': '-', 'growth_rate': None, 'manufacturing': None, 
                     'construction': None, 'service': None, 'other': None, 'placeholder': True}
    
    kosis_info = get_kosis_grdp_download_info()
    
    return {
        'report_info': {
            'year': year,
            'quarter': quarter,
            'page_number': ''
        },
        'national_summary': {
            'growth_rate': national_growth,
            'direction': '증가' if national_growth is not None and national_growth >= 0 else ('감소' if national_growth is not None else None),
            'contributions': national_contributions,
            'placeholder': national_placeholder,
            'needs_review': national_placeholder
        },
        'top_region': {
            'name': top_region.get('region', '-'),
            'growth_rate': top_region.get('growth_rate', 0.0),
            'contributions': {
                'manufacturing': top_region.get('manufacturing', 0.0),
                'construction': top_region.get('construction', 0.0),
                'service': top_region.get('service', 0.0),
                'other': top_region.get('other', 0.0)
            },
            'placeholder': top_region.get('placeholder', True),
            'needs_review': top_region.get('placeholder', True)
        },
        'regional_data': regional_data,
        'chart_config': {
            'y_axis': {
                'min': -6,
                'max': 8,
                'step': 2
            }
        },
        'kosis_info': kosis_info,
        'data_missing': national_placeholder,
        'needs_review': national_placeholder  # 전체 데이터가 수정 필요한지 표시
    }

