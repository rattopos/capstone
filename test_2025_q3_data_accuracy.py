#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
25년 3분기 데이터 정확성 테스트

목적: 동적 매핑 시스템이 엑셀에서 2025년 3분기 데이터를 정확히 추출하는지 검증
"""

import sys
import pandas as pd
from pathlib import Path

# 프로젝트 루트를 sys.path에 추가
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from templates.mining_manufacturing_generator import MiningManufacturingGenerator


class Color:
    """터미널 색상"""
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    BOLD = '\033[1m'
    END = '\033[0m'


def print_header(text):
    """헤더 출력"""
    print(f"\n{Color.BOLD}{Color.CYAN}{'='*80}{Color.END}")
    print(f"{Color.BOLD}{Color.CYAN}{text}{Color.END}")
    print(f"{Color.BOLD}{Color.CYAN}{'='*80}{Color.END}\n")


def print_success(text):
    """성공 메시지"""
    print(f"{Color.GREEN}✅ {text}{Color.END}")


def print_error(text):
    """에러 메시지"""
    print(f"{Color.RED}❌ {text}{Color.END}")


def print_warning(text):
    """경고 메시지"""
    print(f"{Color.YELLOW}⚠️  {text}{Color.END}")


def print_info(text):
    """정보 메시지"""
    print(f"{Color.BLUE}ℹ️  {text}{Color.END}")


def verify_excel_raw_data(excel_path):
    """엑셀 파일에서 직접 2025년 3분기 데이터 확인"""
    print_header("1단계: 엑셀 원본 데이터 직접 확인")
    
    # A 분석 시트 읽기
    df = pd.read_excel(excel_path, sheet_name='A 분석', header=None)
    
    print_info(f"시트 크기: {df.shape[0]}행 × {df.shape[1]}열")
    
    # 헤더 행 찾기
    header_row_idx = 2  # 일반적으로 행 2
    header_row = df.iloc[header_row_idx]
    
    print_info(f"헤더 행 인덱스: {header_row_idx}")
    print_info(f"헤더 샘플 (첫 25개): {list(header_row[:25])}")
    
    # 2025 3/4 컬럼 찾기
    target_col = None
    for col_idx, cell_value in enumerate(header_row):
        if pd.notna(cell_value):
            cell_str = str(cell_value).strip()
            if '2025' in cell_str and '3/4' in cell_str:
                target_col = col_idx
                print_success(f"2025년 3분기 컬럼 발견: 인덱스 {col_idx} ('{cell_str}')")
                break
    
    if target_col is None:
        print_error("2025년 3분기 컬럼을 찾을 수 없습니다!")
        return None
    
    # 전국 총지수 데이터 추출 (데이터 시작 행은 3)
    data_start_row = 3
    
    # 전국 BCD (총지수) 찾기
    nationwide_total_row = None
    for row_idx in range(data_start_row, min(data_start_row + 20, len(df))):
        row = df.iloc[row_idx]
        region_code = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
        industry_code = str(row.iloc[6]) if pd.notna(row.iloc[6]) else ""
        
        if region_code == "00" and industry_code == "BCD":
            nationwide_total_row = row_idx
            growth_rate = row.iloc[target_col]
            print_success(f"전국 총지수(BCD) 발견: 행 {row_idx}")
            print_info(f"  지역코드: {region_code}")
            print_info(f"  산업코드: {industry_code}")
            print_info(f"  2025 3분기 증감률: {growth_rate}%")
            
            return {
                "target_col": target_col,
                "data_row": nationwide_total_row,
                "growth_rate": growth_rate,
                "header_row_idx": header_row_idx
            }
    
    print_error("전국 총지수 데이터를 찾을 수 없습니다!")
    return None


def verify_generator_extraction(excel_path):
    """Generator를 통한 데이터 추출 검증"""
    print_header("2단계: Generator를 통한 데이터 추출")
    
    try:
        generator = MiningManufacturingGenerator(
            excel_path=excel_path,
            year=2025,
            quarter=3
        )
        
        print_success("Generator 인스턴스 생성 성공")
        
        # 전체 데이터 추출 (시트 로드 포함)
        print_info("전체 데이터 추출 중...")
        all_data = generator.extract_all_data()
        
        # 전국 데이터 확인
        nationwide_data = all_data.get('nationwide_data')
        
        if nationwide_data:
            print_success("전국 데이터 추출 성공!")
            print_info(f"  증감률: {nationwide_data.get('growth_rate')}%")
            print_info(f"  나레이션: {nationwide_data.get('narrative', 'N/A')[:100]}...")
            
            # 증가/감소 업종 확인
            increase_industries = nationwide_data.get('increase_industries', [])
            decrease_industries = nationwide_data.get('decrease_industries', [])
            
            print_info(f"  증가 업종 수: {len(increase_industries)}")
            print_info(f"  감소 업종 수: {len(decrease_industries)}")
            
            if increase_industries:
                print_info(f"  주요 증가 업종: {increase_industries[0].get('name', 'N/A')}")
            if decrease_industries:
                print_info(f"  주요 감소 업종: {decrease_industries[0].get('name', 'N/A')}")
            
            return nationwide_data
        else:
            print_error("전국 데이터가 비어있습니다!")
            return None
            
    except Exception as e:
        print_error(f"Generator 실행 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return None


def compare_results(raw_data, generator_data):
    """원본 데이터와 Generator 추출 데이터 비교"""
    print_header("3단계: 데이터 정확성 비교")
    
    if raw_data is None or generator_data is None:
        print_error("비교할 데이터가 없습니다!")
        return False
    
    raw_growth_rate = raw_data.get('growth_rate')
    generator_growth_rate = generator_data.get('growth_rate')
    
    print_info(f"엑셀 원본 증감률: {raw_growth_rate}%")
    print_info(f"Generator 추출 증감률: {generator_growth_rate}%")
    
    # 소수점 2자리까지 비교
    try:
        raw_value = float(raw_growth_rate) if pd.notna(raw_growth_rate) else None
        gen_value = float(generator_growth_rate) if pd.notna(generator_growth_rate) else None
        
        if raw_value is None or gen_value is None:
            print_error("증감률 값이 None입니다!")
            return False
        
        difference = abs(raw_value - gen_value)
        
        if difference < 0.01:  # 0.01% 이내 차이는 허용
            print_success(f"✨ 데이터 일치! (차이: {difference:.4f}%)")
            return True
        else:
            print_warning(f"데이터 불일치! (차이: {difference:.4f}%)")
            return False
            
    except Exception as e:
        print_error(f"비교 중 오류: {e}")
        return False


def test_column_detection(excel_path):
    """컬럼 감지 로직 테스트"""
    print_header("4단계: 동적 컬럼 감지 로직 검증")
    
    try:
        generator = MiningManufacturingGenerator(
            excel_path=excel_path,
            year=2025,
            quarter=3
        )
        
        # 데이터 추출 (시트 로드 포함)
        try:
            all_data = generator.extract_all_data()
        except Exception as e:
            print_error(f"extract_all_data 실행 실패: {e}")
            return False
        
        # 분석 시트의 헤더 확인
        if hasattr(generator, 'df_analysis') and generator.df_analysis is not None:
            df = generator.df_analysis
            header_row = df.iloc[2]  # 일반적으로 행 2
            
            # find_target_col_index 직접 호출
            from templates.base_generator import BaseGenerator
            
            # BaseGenerator의 메서드 사용
            if hasattr(generator, 'find_target_col_index'):
                target_col = generator.find_target_col_index(header_row, 2025, 3)
                print_success(f"find_target_col_index() 결과: 컬럼 {target_col}")
                
                # 해당 컬럼의 헤더 값 확인
                if target_col is not None and target_col < len(header_row):
                    header_value = header_row.iloc[target_col]
                    print_info(f"  컬럼 헤더: '{header_value}'")
                    
                    if '2025' in str(header_value) and '3/4' in str(header_value):
                        print_success("✨ 정확한 컬럼을 찾았습니다!")
                        return True
                    else:
                        print_error(f"잘못된 컬럼을 찾았습니다: '{header_value}'")
                        return False
                else:
                    print_error("컬럼 인덱스가 유효하지 않습니다!")
                    return False
            else:
                print_error("find_target_col_index 메서드를 찾을 수 없습니다!")
                return False
        else:
            print_error("df_analysis를 찾을 수 없습니다!")
            return False
            
    except Exception as e:
        print_error(f"컬럼 감지 테스트 중 오류: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """메인 테스트 실행"""
    print_header("🧪 25년 3분기 데이터 정확성 테스트")
    
    excel_path = project_root / "uploads" / "분석표_25년_3분기_캡스톤업데이트_ee0197ea.xlsx"
    
    if not excel_path.exists():
        print_error(f"엑셀 파일을 찾을 수 없습니다: {excel_path}")
        return False
    
    print_success(f"엑셀 파일 발견: {excel_path.name}")
    
    # 테스트 실행
    test_results = []
    
    # 1. 원본 데이터 확인
    raw_data = verify_excel_raw_data(str(excel_path))
    test_results.append(("엑셀 원본 데이터 확인", raw_data is not None))
    
    # 2. Generator 추출
    generator_data = verify_generator_extraction(str(excel_path))
    test_results.append(("Generator 데이터 추출", generator_data is not None))
    
    # 3. 데이터 비교
    if raw_data and generator_data:
        comparison_result = compare_results(raw_data, generator_data)
        test_results.append(("데이터 정확성 비교", comparison_result))
    else:
        test_results.append(("데이터 정확성 비교", False))
    
    # 4. 컬럼 감지 로직 테스트
    column_detection_result = test_column_detection(str(excel_path))
    test_results.append(("동적 컬럼 감지", column_detection_result))
    
    # 최종 결과 출력
    print_header("📊 테스트 결과 요약")
    
    total_tests = len(test_results)
    passed_tests = sum(1 for _, result in test_results if result)
    
    for test_name, result in test_results:
        if result:
            print_success(f"{test_name}: PASS")
        else:
            print_error(f"{test_name}: FAIL")
    
    print(f"\n{Color.BOLD}총 테스트: {total_tests}, 성공: {passed_tests}, 실패: {total_tests - passed_tests}{Color.END}")
    
    if passed_tests == total_tests:
        print(f"\n{Color.GREEN}{Color.BOLD}🎉 모든 테스트 통과! 25년 3분기 데이터가 정확히 추출됩니다.{Color.END}")
        return True
    else:
        print(f"\n{Color.RED}{Color.BOLD}❌ 일부 테스트 실패. 코드 검토가 필요합니다.{Color.END}")
        return False


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
