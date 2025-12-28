# -*- coding: utf-8 -*-
"""
GRDP 관련 서비스 함수
"""

import pandas as pd


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
    """KOSIS에서 다운로드한 GRDP 엑셀 파일 파싱"""
    try:
        print(f"[KOSIS] GRDP 파일 파싱: {file_path}")
        
        xl = pd.ExcelFile(file_path)
        print(f"[KOSIS] 시트 목록: {xl.sheet_names}")
        
        df = pd.read_excel(file_path, header=None)
        
        regions = ['전국', '서울', '부산', '대구', '인천', '광주', '대전', '울산', '세종',
                   '경기', '강원', '충북', '충남', '전북', '전남', '경북', '경남', '제주']
        
        region_groups = {
            '서울': '경인', '인천': '경인', '경기': '경인',
            '대전': '충청', '세종': '충청', '충북': '충청', '충남': '충청',
            '광주': '호남', '전북': '호남', '전남': '호남', '제주': '호남',
            '대구': '동북', '경북': '동북', '강원': '동북',
            '부산': '동남', '울산': '동남', '경남': '동남'
        }
        
        regional_data = []
        national_growth = 0.0
        top_region = {'name': '-', 'growth_rate': 0.0}
        
        for i, row in df.iterrows():
            for j, val in enumerate(row):
                if pd.notna(val) and str(val).strip() in regions:
                    region_name = str(val).strip()
                    growth_rate = 0.0
                    for k in range(j+1, min(j+10, len(row))):
                        try:
                            growth_rate = float(row[k])
                            break
                        except:
                            continue
                    
                    if region_name == '전국':
                        national_growth = growth_rate
                    else:
                        regional_data.append({
                            'region': region_name,
                            'region_group': region_groups.get(region_name, ''),
                            'growth_rate': growth_rate,
                            'manufacturing': 0.0,
                            'construction': 0.0,
                            'service': 0.0,
                            'other': 0.0
                        })
                        
                        if growth_rate > top_region['growth_rate']:
                            top_region = {'name': region_name, 'growth_rate': growth_rate}
        
        if not regional_data:
            print("[KOSIS] GRDP 데이터 추출 실패: 지역 데이터를 찾을 수 없음")
            return None
        
        result = {
            'report_info': {
                'year': year or 2025,
                'quarter': quarter or 2,
                'page_number': ''
            },
            'national_summary': {
                'growth_rate': national_growth,
                'direction': '증가' if national_growth > 0 else '감소',
                'contributions': {
                    'manufacturing': 0.0,
                    'construction': 0.0,
                    'service': 0.0,
                    'other': 0.0
                }
            },
            'top_region': {
                'name': top_region['name'],
                'growth_rate': top_region['growth_rate'],
                'contributions': {
                    'manufacturing': 0.0,
                    'construction': 0.0,
                    'service': 0.0,
                    'other': 0.0
                }
            },
            'regional_data': regional_data,
            'source': 'KOSIS 실험적통계'
        }
        
        print(f"[KOSIS] GRDP 파싱 완료 - 전국: {national_growth}%, 최고: {top_region['name']}({top_region['growth_rate']}%)")
        return result
        
    except Exception as e:
        import traceback
        print(f"[KOSIS] GRDP 파일 파싱 오류: {e}")
        traceback.print_exc()
        return None


def get_default_grdp_data(year, quarter):
    """기본 GRDP 데이터 (KOSIS 안내 정보 포함)"""
    regions = ['전국', '서울', '인천', '경기', '대전', '세종', '충북', '충남',
               '광주', '전북', '전남', '제주', '대구', '경북', '강원', '부산', '울산', '경남']
    
    regional_data = []
    region_groups = {
        '서울': '경인', '인천': '경인', '경기': '경인',
        '대전': '충청', '세종': '충청', '충북': '충청', '충남': '충청',
        '광주': '호남', '전북': '호남', '전남': '호남', '제주': '호남',
        '대구': '동북', '경북': '동북', '강원': '동북',
        '부산': '동남', '울산': '동남', '경남': '동남'
    }
    
    for region in regions:
        regional_data.append({
            'region': region,
            'region_group': region_groups.get(region, ''),
            'growth_rate': 0.0,
            'manufacturing': 0.0,
            'construction': 0.0,
            'service': 0.0,
            'other': 0.0,
            'placeholder': True
        })
    
    kosis_info = get_kosis_grdp_download_info()
    
    return {
        'report_info': {
            'year': year,
            'quarter': quarter,
            'page_number': ''
        },
        'national_summary': {
            'growth_rate': 0.0,
            'direction': '증가',
            'contributions': {
                'manufacturing': 0.0,
                'construction': 0.0,
                'service': 0.0,
                'other': 0.0
            },
            'placeholder': True
        },
        'top_region': {
            'name': '-',
            'growth_rate': 0.0,
            'contributions': {
                'manufacturing': 0.0,
                'construction': 0.0,
                'service': 0.0,
                'other': 0.0
            },
            'placeholder': True
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
        'data_missing': True
    }

