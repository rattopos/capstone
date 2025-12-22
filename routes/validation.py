"""
검증 관련 라우트
/api/validate, /api/check-default-file
"""

import logging
from pathlib import Path
from typing import Optional

from flask import Blueprint, request, jsonify, current_app
from werkzeug.utils import secure_filename

from .common import (
    allowed_file, DEFAULT_EXCEL_FILE,
    ValidationError, ValidationResult, ErrorCodes,
    ValidationCache, error_response
)
from .validator import TemplateValidator

# 로거 설정
logger = logging.getLogger(__name__)

validation_bp = Blueprint('validation', __name__, url_prefix='/api')


# ============================================================
# API 엔드포인트
# ============================================================

@validation_bp.route('/check-default-file', methods=['GET'])
def check_default_file():
    """
    기본 엑셀 파일 존재 여부 확인
    
    Returns:
        JSON: exists (bool), filename (str) 또는 message (str)
    """
    if DEFAULT_EXCEL_FILE.exists():
        return jsonify({
            'success': True,
            'exists': True,
            'filename': DEFAULT_EXCEL_FILE.name,
            'path': str(DEFAULT_EXCEL_FILE)
        })
    else:
        return jsonify({
            'success': True,
            'exists': False,
            'message': '기본 엑셀 파일이 없습니다. 파일을 업로드해주세요.'
        })


@validation_bp.route('/validate', methods=['POST'])
def validate_template():
    """
    템플릿 유효성 검증
    
    Request:
        - template_name: 템플릿 파일명 (필수)
        - year: 연도 (선택)
        - quarter: 분기 (선택)
        - excel_file: 엑셀 파일 (선택, 없으면 기본 파일 사용)
    
    Returns:
        JSON: ValidationResult 형식의 검증 결과
    """
    try:
        # 1. 요청 파라미터 추출
        template_name = request.form.get('template_name', '').strip()
        year_str = request.form.get('year', '')
        quarter_str = request.form.get('quarter', '')
        excel_file = request.files.get('excel_file')
        
        # 2. 템플릿 선택 검증
        if not template_name:
            return error_response(
                ValidationError(
                    code=ErrorCodes.TEMPLATE_NOT_SELECTED,
                    message='템플릿을 선택해주세요.'
                ),
                status_code=400
            )
        
        # 3. 연도/분기 파싱
        year: Optional[int] = int(year_str) if year_str else None
        quarter: Optional[int] = int(quarter_str) if quarter_str else None
        
        # 4. 엑셀 파일 경로 결정
        excel_path = _get_excel_path(excel_file)
        if isinstance(excel_path, tuple):
            # 에러 응답 반환
            return excel_path
        
        # 5. 템플릿 경로 확인
        template_path = Path('templates') / template_name
        if not template_path.exists():
            return error_response(
                ValidationError(
                    code=ErrorCodes.TEMPLATE_NOT_FOUND,
                    message='템플릿 파일을 찾을 수 없습니다.',
                    detail=template_name
                ),
                status_code=404
            )
        
        # 6. 캐시 확인
        cache_key = ValidationCache.get_cache_key(
            str(excel_path), template_name, year or 0, quarter or 0
        )
        cached_result = ValidationCache.get(cache_key)
        if cached_result:
            logger.debug(f"캐시 히트: {cache_key}")
            response = cached_result.to_dict()
            response['cached'] = True
            return jsonify(response)
        
        # 7. 검증 수행
        validator = TemplateValidator(excel_path, template_path)
        result = validator.validate_all(year, quarter)
        
        # 8. 결과 캐싱
        ValidationCache.set(cache_key, result)
        
        # 9. 응답 반환
        return jsonify(result.to_dict())
        
    except ValueError as e:
        logger.warning(f"잘못된 입력값: {str(e)}")
        return error_response(
            ValidationError(
                code=ErrorCodes.VALIDATION_ERROR,
                message='잘못된 입력값입니다.',
                detail=str(e)
            ),
            status_code=400
        )
        
    except Exception as e:
        logger.error(f"검증 중 오류: {str(e)}", exc_info=True)
        return error_response(
            ValidationError(
                code=ErrorCodes.VALIDATION_ERROR,
                message='검증 중 오류가 발생했습니다.',
                detail=str(e)
            ),
            status_code=500
        )


@validation_bp.route('/validate/cache/clear', methods=['POST'])
def clear_validation_cache():
    """검증 캐시 초기화"""
    try:
        ValidationCache.clear()
        return jsonify({
            'success': True,
            'message': '캐시가 초기화되었습니다.'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@validation_bp.route('/validate/cache/cleanup', methods=['POST'])
def cleanup_validation_cache():
    """만료된 캐시 정리"""
    try:
        removed_count = ValidationCache.cleanup_expired()
        return jsonify({
            'success': True,
            'removed_count': removed_count,
            'message': f'{removed_count}개의 만료된 캐시가 정리되었습니다.'
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


# ============================================================
# 헬퍼 함수
# ============================================================

def _get_excel_path(excel_file):
    """
    엑셀 파일 경로를 결정합니다.
    
    Args:
        excel_file: 업로드된 파일 객체 (또는 None)
        
    Returns:
        Path: 엑셀 파일 경로
        또는
        Tuple: 에러 응답 (jsonify 결과, 상태 코드)
    """
    if excel_file and excel_file.filename:
        # 파일 형식 검증
        if not allowed_file(excel_file.filename):
            return error_response(
                ValidationError(
                    code=ErrorCodes.INVALID_FILE_TYPE,
                    message='지원하지 않는 파일 형식입니다.',
                    detail='xlsx, xls 파일만 지원됩니다.'
                ),
                status_code=400
            )
        
        # 파일 저장
        excel_filename = secure_filename(excel_file.filename)
        upload_folder = Path(current_app.config['UPLOAD_FOLDER'])
        upload_folder.mkdir(parents=True, exist_ok=True)
        
        excel_path = upload_folder / excel_filename
        excel_file.save(str(excel_path))
        
        return excel_path
    else:
        # 기본 파일 사용
        if not DEFAULT_EXCEL_FILE.exists():
            return error_response(
                ValidationError(
                    code=ErrorCodes.EXCEL_NOT_FOUND,
                    message='기본 엑셀 파일을 찾을 수 없습니다.',
                    detail=str(DEFAULT_EXCEL_FILE.name)
                ),
                status_code=400
            )
        return DEFAULT_EXCEL_FILE
