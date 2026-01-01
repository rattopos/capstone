# -*- coding: utf-8 -*-
"""
Excel 파일 전처리 서비스

업로드된 엑셀 파일의 수식을 계산하여 분석 시트의 값을 채웁니다.
"""

import os
import shutil
from pathlib import Path
from typing import Optional, Tuple
import platform


def preprocess_excel(excel_path: str, output_path: Optional[str] = None) -> Tuple[str, bool, str]:
    """
    엑셀 파일의 수식을 계산하여 새 파일로 저장
    
    Args:
        excel_path: 원본 엑셀 파일 경로
        output_path: 출력 파일 경로 (None이면 원본 덮어쓰기)
    
    Returns:
        Tuple[str, bool, str]: (처리된 파일 경로, 성공 여부, 메시지)
    """
    if output_path is None:
        output_path = excel_path
    
    # 1. xlwings 시도 (Mac/Windows - Excel 앱 필요)
    result_path, success, message = _try_xlwings(excel_path, output_path)
    if success:
        return result_path, True, message
    
    # 2. formulas 라이브러리 시도 (순수 Python)
    result_path, success, message = _try_formulas(excel_path, output_path)
    if success:
        return result_path, True, message
    
    # 3. openpyxl로 직접 계산 시도 (기존 방식)
    result_path, success, message = _try_openpyxl_calculation(excel_path, output_path)
    if success:
        return result_path, True, message
    
    # 4. 모두 실패 시 원본 반환 (fallback 로직 사용)
    print(f"[전처리] 수식 계산 실패 - 원본 파일 사용, generator fallback 로직 활성화")
    return excel_path, False, "수식 계산 라이브러리 없음 - fallback 모드"


def _try_xlwings(excel_path: str, output_path: str) -> Tuple[str, bool, str]:
    """xlwings를 사용하여 Excel 앱으로 수식 계산"""
    try:
        import xlwings as xw
        
        print(f"[xlwings] Excel 앱으로 수식 계산 시작...")
        
        # Excel 앱을 백그라운드에서 실행
        app = xw.App(visible=False)
        
        try:
            # 파일 열기
            wb = app.books.open(excel_path)
            
            # 모든 시트의 수식 강제 재계산
            wb.app.calculate()
            
            # 저장 (수식 결과가 캐시됨)
            wb.save(output_path)
            wb.close()
            
            print(f"[xlwings] 수식 계산 완료: {output_path}")
            return output_path, True, "xlwings로 수식 계산 완료"
            
        finally:
            app.quit()
            
    except ImportError:
        return excel_path, False, "xlwings 미설치"
    except Exception as e:
        print(f"[xlwings] 오류: {e}")
        return excel_path, False, f"xlwings 오류: {str(e)}"


def _try_formulas(excel_path: str, output_path: str) -> Tuple[str, bool, str]:
    """formulas 라이브러리를 사용하여 순수 Python으로 수식 계산"""
    try:
        import formulas
        
        print(f"[formulas] 순수 Python으로 수식 계산 시작...")
        
        # 수식 모델 생성
        xl_model = formulas.ExcelModel().loads(excel_path).finish()
        
        # 모든 수식 계산
        xl_model.calculate()
        
        # 결과 저장
        xl_model.write(output_path)
        
        print(f"[formulas] 수식 계산 완료: {output_path}")
        return output_path, True, "formulas로 수식 계산 완료"
        
    except ImportError:
        return excel_path, False, "formulas 미설치"
    except Exception as e:
        print(f"[formulas] 오류: {e}")
        return excel_path, False, f"formulas 오류: {str(e)}"


def _try_openpyxl_calculation(excel_path: str, output_path: str) -> Tuple[str, bool, str]:
    """
    openpyxl을 사용하여 시트 간 참조 수식 계산
    
    분석 시트의 수식이 집계 시트를 참조하는 경우:
    ='시트이름'!셀주소 형태의 간단한 참조만 처리
    """
    try:
        import openpyxl
        import re
        
        print(f"[openpyxl] 시트 간 참조 수식 계산 시작...")
        
        # 분석 시트 → 집계 시트 매핑
        analysis_aggregate_mapping = {
            'A 분석': 'A(광공업생산)집계',
            'B 분석': 'B(서비스업생산)집계',
            'C 분석': 'C(소비)집계',
            'D(고용률)분석': 'D(고용률)집계',
            'D(실업)분석': 'D(실업)집계',
            'E(지출목적물가) 분석': 'E(지출목적물가)집계',
            'E(품목성질물가)분석': 'E(품목성질물가)집계',
            "F'분석": "F'(건설)집계",
            'G 분석': 'G(수출)집계',
            'H 분석': 'H(수입)집계',
            'I(순인구이동)분석': 'I(순인구이동)집계',
        }
        
        # 원본 파일 복사 (원본 보존)
        if excel_path != output_path:
            shutil.copy2(excel_path, output_path)
        
        # 수식 포함 모드로 열기
        wb = openpyxl.load_workbook(output_path, data_only=False)
        
        # 집계 시트 데이터 캐싱 (값 모드로 다시 열어서)
        wb_data = openpyxl.load_workbook(output_path, data_only=True)
        aggregate_cache = {}
        
        for analysis_sheet, aggregate_sheet in analysis_aggregate_mapping.items():
            if aggregate_sheet in wb_data.sheetnames:
                ws_agg = wb_data[aggregate_sheet]
                sheet_data = {}
                for row in ws_agg.iter_rows(min_row=1, max_row=ws_agg.max_row):
                    for cell in row:
                        if cell.value is not None:
                            sheet_data[(cell.row, cell.column)] = cell.value
                aggregate_cache[aggregate_sheet] = sheet_data
        
        wb_data.close()
        
        # 분석 시트의 수식을 값으로 대체
        calculated_count = 0
        formula_count = 0
        
        for analysis_sheet, aggregate_sheet in analysis_aggregate_mapping.items():
            if analysis_sheet not in wb.sheetnames:
                continue
            if aggregate_sheet not in aggregate_cache:
                continue
            
            ws_analysis = wb[analysis_sheet]
            agg_data = aggregate_cache[aggregate_sheet]
            
            for row in ws_analysis.iter_rows(min_row=1, max_row=ws_analysis.max_row):
                for cell in row:
                    if cell.value is None:
                        continue
                    
                    val = str(cell.value)
                    
                    # 수식인 경우 (=로 시작)
                    if val.startswith('='):
                        formula_count += 1
                        
                        # 집계 시트 참조 파싱: ='시트이름'!셀주소
                        match = re.match(r"^='?([^'!]+)'?!([A-Z]+)(\d+)$", val)
                        if match:
                            ref_sheet = match.group(1)
                            ref_col_letter = match.group(2)
                            ref_row = int(match.group(3))
                            
                            # 열 문자를 숫자로 변환 (A=1, B=2, ...)
                            ref_col = 0
                            for i, c in enumerate(reversed(ref_col_letter)):
                                ref_col += (ord(c) - ord('A') + 1) * (26 ** i)
                            
                            # 해당 집계 시트에서 값 가져오기
                            if ref_sheet in aggregate_cache:
                                ref_value = aggregate_cache[ref_sheet].get((ref_row, ref_col))
                                if ref_value is not None:
                                    cell.value = ref_value
                                    calculated_count += 1
        
        wb.save(output_path)
        wb.close()
        
        if calculated_count > 0:
            print(f"[openpyxl] 수식 계산 완료: {calculated_count}개 셀 (총 {formula_count}개 수식)")
            return output_path, True, f"openpyxl로 {calculated_count}개 수식 계산 완료"
        else:
            return excel_path, False, "계산할 수식 없음"
        
    except ImportError:
        return excel_path, False, "openpyxl 미설치"
    except Exception as e:
        print(f"[openpyxl] 오류: {e}")
        import traceback
        traceback.print_exc()
        return excel_path, False, f"openpyxl 오류: {str(e)}"


def check_available_methods() -> dict:
    """사용 가능한 수식 계산 방법 확인"""
    methods = {
        'xlwings': False,
        'formulas': False,
        'openpyxl': False,
        'excel_installed': False,
    }
    
    # xlwings 확인
    try:
        import xlwings as xw
        methods['xlwings'] = True
        
        # Excel 설치 여부 확인 (Mac/Windows)
        if platform.system() == 'Darwin':  # macOS
            methods['excel_installed'] = os.path.exists('/Applications/Microsoft Excel.app')
        elif platform.system() == 'Windows':
            try:
                app = xw.App(visible=False)
                app.quit()
                methods['excel_installed'] = True
            except:
                pass
    except ImportError:
        pass
    
    # formulas 확인
    try:
        import formulas
        methods['formulas'] = True
    except ImportError:
        pass
    
    # openpyxl 확인 (항상 있어야 함)
    try:
        import openpyxl
        methods['openpyxl'] = True
    except ImportError:
        pass
    
    return methods


def get_recommended_method() -> str:
    """권장 수식 계산 방법 반환"""
    methods = check_available_methods()
    
    if methods['xlwings'] and methods['excel_installed']:
        return 'xlwings (Excel 앱 사용 - 가장 정확)'
    elif methods['formulas']:
        return 'formulas (순수 Python - 복잡한 수식 지원)'
    elif methods['openpyxl']:
        return 'openpyxl (기본 참조 수식만 지원)'
    else:
        return 'fallback (generator에서 집계 시트 직접 사용)'


if __name__ == '__main__':
    # 테스트
    print("=== 사용 가능한 수식 계산 방법 ===")
    methods = check_available_methods()
    for method, available in methods.items():
        status = "✓ 사용 가능" if available else "✗ 사용 불가"
        print(f"  {method}: {status}")
    
    print(f"\n권장 방법: {get_recommended_method()}")

