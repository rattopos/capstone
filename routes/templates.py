"""
템플릿 관련 라우트
/api/templates, /api/template-sheets
"""

from pathlib import Path
from flask import Blueprint, request, jsonify

from .common import get_template_manager, get_schema_loader

templates_bp = Blueprint('templates', __name__, url_prefix='/api')

# 성능 최적화: 템플릿 목록 캐시
_templates_cache = None
_templates_cache_mtime = None


@templates_bp.route('/templates', methods=['GET'])
def get_templates():
    """사용 가능한 템플릿 목록 반환"""
    global _templates_cache, _templates_cache_mtime
    
    templates_dir = Path('templates')
    schema_loader = get_schema_loader()
    
    # 캐시 유효성 확인
    if templates_dir.exists():
        current_mtime = templates_dir.stat().st_mtime
        for f in templates_dir.glob('*.html'):
            file_mtime = f.stat().st_mtime
            if file_mtime > current_mtime:
                current_mtime = file_mtime
        
        if _templates_cache is not None and _templates_cache_mtime == current_mtime:
            return jsonify({'templates': _templates_cache})
    
    templates = []
    
    if templates_dir.exists():
        html_files = list(templates_dir.glob('*.html'))
        template_mapping = schema_loader.load_template_mapping()
        
        for template_path in html_files:
            template_name = template_path.name
            
            try:
                template_manager = get_template_manager(template_path)
                markers = template_manager.extract_markers()
                
                required_sheets = set()
                for marker in markers:
                    sheet_name = marker.get('sheet_name', '').strip()
                    if sheet_name:
                        required_sheets.add(sheet_name)
                
                display_name = template_name.replace('.html', '')
                for sheet_name, info in template_mapping.items():
                    if info['template'] == template_name:
                        display_name = info['display_name']
                        break
                
                templates.append({
                    'name': template_name,
                    'path': str(template_path),
                    'display_name': display_name,
                    'required_sheets': list(required_sheets)
                })
            except Exception:
                display_name = template_name.replace('.html', '')
                for sheet_name, info in template_mapping.items():
                    if info['template'] == template_name:
                        display_name = info['display_name']
                        break
                
                templates.append({
                    'name': template_name,
                    'path': str(template_path),
                    'display_name': display_name,
                    'required_sheets': []
                })
    
    # 서울 관련 템플릿 제외
    excluded_templates = {'서울.html', '서울주요지표.html'}
    templates = [t for t in templates if t['name'] not in excluded_templates]
    
    templates.sort(key=lambda x: x['display_name'])
    
    _templates_cache = templates
    _templates_cache_mtime = current_mtime if templates_dir.exists() else None
    
    return jsonify({'templates': templates})


@templates_bp.route('/template-sheets', methods=['POST'])
def get_template_sheets():
    """템플릿이 필요한 시트 목록 반환"""
    try:
        template_name = request.form.get('template_name', '')
        if not template_name:
            return jsonify({'error': '템플릿명이 필요합니다.'}), 400
        
        template_path = Path('templates') / template_name
        if not template_path.exists():
            return jsonify({'error': f'템플릿 파일을 찾을 수 없습니다: {template_name}'}), 404
        
        template_manager = get_template_manager(template_path)
        markers = template_manager.extract_markers()
        
        required_sheets = set()
        for marker in markers:
            sheet_name = marker.get('sheet_name', '').strip()
            if sheet_name:
                required_sheets.add(sheet_name)
        
        return jsonify({
            'template_name': template_name,
            'required_sheets': list(required_sheets)
        })
    except Exception as e:
        return jsonify({
            'error': f'템플릿 분석 중 오류가 발생했습니다: {str(e)}'
        }), 500

