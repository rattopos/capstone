# -*- coding: utf-8 -*-
"""
디버그용 HTML 보도자료 생성 라우트
모든 페이지를 A4 크기로 순차적으로 이어붙여 출력합니다.
"""

from pathlib import Path
from datetime import datetime
import re

from flask import Blueprint, request, jsonify, session, render_template_string

from config.settings import TEMPLATES_DIR, UPLOAD_FOLDER, DEBUG_FOLDER


def extract_body_content(html_content):
    """
    완전한 HTML 문서에서 body 내용과 스타일을 추출합니다.
    스타일은 scoped style 태그로 컨텐츠에 포함됩니다.
    """
    if not html_content:
        return html_content, ""
    
    # body 태그 내용 추출
    body_match = re.search(r'<body[^>]*>(.*?)</body>', html_content, re.DOTALL | re.IGNORECASE)
    if body_match:
        body_content = body_match.group(1)
    else:
        # body 태그가 없으면 원본 반환
        body_content = html_content
    
    # script 태그 분리 (차트 등에 필요)
    scripts = re.findall(r'<script[^>]*>.*?</script>', body_content, re.DOTALL | re.IGNORECASE)
    
    # style 태그 추출 (head에서)
    style_matches = re.findall(r'<style[^>]*>(.*?)</style>', html_content, re.DOTALL | re.IGNORECASE)
    inline_style = "\n".join(style_matches) if style_matches else ""
    
    # body 내부에서 불필요한 래퍼 제거
    # 컨테이너 클래스 패턴 (page, cover-container, summary-container 등)
    # page 클래스는 page-number, page-title 등 제외 (단독 또는 공백으로 구분된 경우만)
    container_patterns = [
        r'<div[^>]*class="page"[^>]*>(.*)</div>\s*$',  # class="page" 단독
        r'<div[^>]*class="page\s[^"]*"[^>]*>(.*)</div>\s*$',  # class="page ..." 시작
        r'<div[^>]*class="[^"]*\spage"[^>]*>(.*)</div>\s*$',  # class="... page" 끝
        r'<div[^>]*class="[^"]*\spage\s[^"]*"[^>]*>(.*)</div>\s*$',  # class="... page ..." 중간
        r'<div[^>]*class="[^"]*-container[^"]*"[^>]*>(.*)</div>\s*$',  # *-container 패턴
    ]
    
    inner_content = None
    for pattern in container_patterns:
        match = re.search(pattern, body_content, re.DOTALL | re.IGNORECASE)
        if match:
            inner_content = match.group(1).strip()
            break
    
    if inner_content is None:
        inner_content = body_content.strip()
    
    # script 태그 제거 (나중에 별도로 추가)
    inner_content = re.sub(r'<script[^>]*>.*?</script>', '', inner_content, flags=re.DOTALL | re.IGNORECASE)
    
    # 최종 컨텐츠 구성: 스타일 + 본문 + 스크립트
    result_content = ""
    if inline_style:
        # 스타일을 scoped 형태로 추가 (중복 방지를 위해 각 페이지별 고유 스타일 유지)
        result_content += f"<style>{inline_style}</style>\n"
    result_content += inner_content
    
    # script 태그 추가
    for script in scripts:
        result_content += "\n" + script
    
    return result_content, inline_style
from config.reports import (
    REPORT_ORDER, SUMMARY_REPORTS, SECTOR_REPORTS, REGIONAL_REPORTS, STATISTICS_REPORTS,
    PAGE_CONFIG, TOC_SECTOR_ITEMS, TOC_REGION_ITEMS
)
from services.report_generator import (
    generate_report_html,
    generate_regional_report_html,
    generate_individual_statistics_html
)
from services.summary_data import (
    get_summary_overview_data,
    get_summary_table_data,
    get_production_summary_data,
    get_consumption_construction_data,
    get_trade_price_data,
    get_employment_population_data
)
from utils.excel_utils import load_generator_module
from jinja2 import Template

debug_bp = Blueprint('debug', __name__, url_prefix='/debug')


@debug_bp.route('/set-session', methods=['POST'])
def set_debug_session():
    """디버그용 세션 설정 - 파일 업로드 없이 경로 직접 설정"""
    data = request.get_json() or {}
    
    # uploads 폴더에서 최신 분석표 파일 찾기
    excel_files = sorted(UPLOAD_FOLDER.glob('분석표*.xlsx'), key=lambda x: x.stat().st_mtime, reverse=True)
    if not excel_files:
        # 프로젝트 루트에서 찾기
        from config.settings import BASE_DIR
        excel_files = sorted(BASE_DIR.glob('분석표*.xlsx'), key=lambda x: x.stat().st_mtime, reverse=True)
    
    if excel_files:
        excel_path = str(excel_files[0])
        session['excel_path'] = excel_path
        session['year'] = data.get('year', 2025)
        session['quarter'] = data.get('quarter', 2)
        
        # 기초자료 수집표 찾기
        raw_files = sorted(UPLOAD_FOLDER.glob('기초자료*.xlsx'), key=lambda x: x.stat().st_mtime, reverse=True)
        if not raw_files:
            from config.settings import BASE_DIR
            raw_files = sorted(BASE_DIR.glob('기초자료*.xlsx'), key=lambda x: x.stat().st_mtime, reverse=True)
        if raw_files:
            session['raw_excel_path'] = str(raw_files[0])
        
        return jsonify({
            'success': True, 
            'excel_path': excel_path,
            'raw_excel_path': session.get('raw_excel_path'),
            'year': session['year'],
            'quarter': session['quarter']
        })
    
    return jsonify({'success': False, 'error': '분석표 엑셀 파일을 찾을 수 없습니다'})


# ===== 디버그 페이지 템플릿 =====
DEBUG_PAGE_TEMPLATE = '''
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>🐛 디버그 - 지역경제동향 보도자료</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&display=swap');
        
        * { margin: 0; padding: 0; box-sizing: border-box; }
        
        body {
            font-family: 'Noto Sans KR', sans-serif;
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
            min-height: 100vh;
            color: #e8e8e8;
        }
        
        .debug-container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 40px 20px;
        }
        
        .debug-header {
            text-align: center;
            margin-bottom: 40px;
            padding: 30px;
            background: rgba(255, 255, 255, 0.05);
            border-radius: 20px;
            border: 1px solid rgba(255, 255, 255, 0.1);
            backdrop-filter: blur(10px);
        }
        
        .debug-header h1 {
            font-size: 2.5rem;
            font-weight: 700;
            margin-bottom: 10px;
            background: linear-gradient(120deg, #e94560, #533483);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }
        
        .debug-header p {
            color: #a0a0a0;
            font-size: 1rem;
        }
        
        .debug-status {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }
        
        .status-card {
            background: rgba(255, 255, 255, 0.05);
            border-radius: 15px;
            padding: 20px;
            border: 1px solid rgba(255, 255, 255, 0.1);
            text-align: center;
        }
        
        .status-card .label {
            font-size: 0.85rem;
            color: #888;
            margin-bottom: 8px;
        }
        
        .status-card .value {
            font-size: 1.4rem;
            font-weight: 600;
            color: #e94560;
        }
        
        .status-card .value.ok { color: #4ade80; }
        .status-card .value.warn { color: #facc15; }
        .status-card .value.error { color: #f87171; }
        
        .quick-setup {
            background: rgba(74, 222, 128, 0.1);
            border: 1px solid rgba(74, 222, 128, 0.3);
            border-radius: 10px;
            padding: 15px 20px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        
        .quick-setup p {
            color: #4ade80;
            margin: 0;
        }
        
        .debug-actions {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 25px;
            margin-bottom: 40px;
        }
        
        .action-card {
            background: rgba(255, 255, 255, 0.08);
            border-radius: 20px;
            padding: 30px;
            border: 1px solid rgba(255, 255, 255, 0.1);
            transition: all 0.3s ease;
        }
        
        .action-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 20px 40px rgba(233, 69, 96, 0.2);
            border-color: rgba(233, 69, 96, 0.3);
        }
        
        .action-card h3 {
            font-size: 1.3rem;
            margin-bottom: 15px;
            color: #fff;
        }
        
        .action-card p {
            color: #a0a0a0;
            font-size: 0.9rem;
            margin-bottom: 20px;
            line-height: 1.6;
        }
        
        .action-btn {
            display: inline-block;
            background: linear-gradient(135deg, #e94560, #533483);
            color: #fff;
            padding: 12px 30px;
            border-radius: 30px;
            text-decoration: none;
            font-weight: 500;
            transition: all 0.3s ease;
            border: none;
            cursor: pointer;
            font-size: 1rem;
        }
        
        .action-btn:hover {
            transform: scale(1.05);
            box-shadow: 0 10px 30px rgba(233, 69, 96, 0.4);
        }
        
        .action-btn.secondary {
            background: transparent;
            border: 2px solid #e94560;
        }
        
        .action-btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
            transform: none;
        }
        
        .report-list {
            background: rgba(255, 255, 255, 0.05);
            border-radius: 20px;
            padding: 30px;
            border: 1px solid rgba(255, 255, 255, 0.1);
        }
        
        .report-list h3 {
            font-size: 1.3rem;
            margin-bottom: 20px;
            color: #fff;
        }
        
        .report-sections {
            display: grid;
            gap: 15px;
        }
        
        .report-section {
            background: rgba(255, 255, 255, 0.03);
            border-radius: 10px;
            padding: 15px 20px;
            border: 1px solid rgba(255, 255, 255, 0.05);
        }
        
        .report-section h4 {
            font-size: 1rem;
            margin-bottom: 10px;
            color: #e94560;
        }
        
        .report-items {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
        }
        
        .report-item {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            background: rgba(255, 255, 255, 0.05);
            padding: 6px 12px;
            border-radius: 20px;
            font-size: 0.85rem;
            color: #ccc;
        }
        
        .report-item .icon { font-size: 1rem; }
        
        .loading-overlay {
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.8);
            z-index: 1000;
            justify-content: center;
            align-items: center;
            flex-direction: column;
        }
        
        .loading-overlay.active { display: flex; }
        
        .spinner {
            width: 60px;
            height: 60px;
            border: 4px solid rgba(233, 69, 96, 0.3);
            border-top-color: #e94560;
            border-radius: 50%;
            animation: spin 1s linear infinite;
        }
        
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        
        .loading-text {
            margin-top: 20px;
            color: #fff;
            font-size: 1.1rem;
        }
        
        .progress-info {
            margin-top: 10px;
            color: #a0a0a0;
            font-size: 0.9rem;
        }
        
        .debug-log {
            background: #0d0d1a;
            border-radius: 15px;
            padding: 20px;
            margin-top: 30px;
            max-height: 300px;
            overflow-y: auto;
            font-family: 'Courier New', monospace;
            font-size: 0.85rem;
        }
        
        .log-entry {
            padding: 5px 0;
            border-bottom: 1px solid rgba(255, 255, 255, 0.05);
        }
        
        .log-entry.info { color: #4ade80; }
        .log-entry.warn { color: #facc15; }
        .log-entry.error { color: #f87171; }
        .log-entry .timestamp { color: #666; margin-right: 10px; }
        
        .footer-info {
            text-align: center;
            margin-top: 40px;
            padding: 20px;
            color: #666;
            font-size: 0.85rem;
        }
    </style>
</head>
<body>
    <div class="debug-container">
        <div class="debug-header">
            <h1>🐛 디버그 모드</h1>
            <p>지역경제동향 보도자료 HTML 생성 및 레이아웃 테스트</p>
        </div>
        
        {% if not excel_loaded %}
        <div class="quick-setup">
            <p>💡 엑셀 파일이 없습니다. 자동으로 분석표를 찾아 세션을 설정합니다.</p>
            <button class="action-btn" onclick="quickSetup()">빠른 설정</button>
        </div>
        {% endif %}
        
        <div class="debug-status">
            <div class="status-card">
                <div class="label">엑셀 파일</div>
                <div class="value {{ 'ok' if excel_loaded else 'error' }}">
                    {{ '✓ 로드됨' if excel_loaded else '✗ 미로드' }}
                </div>
            </div>
            <div class="status-card">
                <div class="label">연도/분기</div>
                <div class="value ok">{{ year }}년 {{ quarter }}분기</div>
            </div>
            <div class="status-card">
                <div class="label">총 보도자료</div>
                <div class="value">{{ total_reports }}개</div>
            </div>
            <div class="status-card">
                <div class="label">총 페이지</div>
                <div class="value">{{ total_pages }}+</div>
            </div>
        </div>
        
        <div class="debug-actions">
            <div class="action-card">
                <h3>📄 전체 보도자료 생성</h3>
                <p>모든 섹션을 A4 크기로 순차적으로 이어붙인 HTML 파일을 생성합니다. 
                   디버그 주석과 페이지 정보가 포함됩니다.</p>
                <button class="action-btn" onclick="generateFullReport()" {{ 'disabled' if not excel_loaded else '' }}>
                    전체 HTML 생성
                </button>
            </div>
            
            <div class="action-card">
                <h3>📊 요약 섹션만</h3>
                <p>표지, 일러두기, 목차, 인포그래픽, 요약 페이지만 생성합니다.</p>
                <button class="action-btn secondary" onclick="generateSection('summary')" {{ 'disabled' if not excel_loaded else '' }}>
                    요약 섹션 생성
                </button>
            </div>
            
            <div class="action-card">
                <h3>🏭 부문별 섹션만</h3>
                <p>광공업생산, 서비스업생산, 소비동향 등 부문별 보도자료만 생성합니다.</p>
                <button class="action-btn secondary" onclick="generateSection('sector')" {{ 'disabled' if not excel_loaded else '' }}>
                    부문별 섹션 생성
                </button>
            </div>
            
            <div class="action-card">
                <h3>🗺️ 시도별 섹션만</h3>
                <p>17개 시도별 경제동향 보도자료와 참고 GRDP를 생성합니다.</p>
                <button class="action-btn secondary" onclick="generateSection('regional')" {{ 'disabled' if not excel_loaded else '' }}>
                    시도별 섹션 생성
                </button>
            </div>
            
            <div class="action-card">
                <h3>📈 통계표 섹션만</h3>
                <p>통계표 목차, 개별 통계표, 부록을 생성합니다.</p>
                <button class="action-btn secondary" onclick="generateSection('statistics')" {{ 'disabled' if not excel_loaded else '' }}>
                    통계표 섹션 생성
                </button>
            </div>
            
            <div class="action-card">
                <h3>🔍 개별 페이지 테스트</h3>
                <p>특정 보도자료 ID를 입력하여 개별 페이지만 테스트합니다.</p>
                <input type="text" id="single-report-id" placeholder="예: manufacturing" 
                       style="width: 100%; padding: 10px; margin-bottom: 10px; border-radius: 10px; 
                              border: 1px solid rgba(255,255,255,0.2); background: rgba(0,0,0,0.3); color: #fff;">
                <button class="action-btn secondary" onclick="generateSingleReport()" {{ 'disabled' if not excel_loaded else '' }}>
                    개별 페이지 생성
                </button>
            </div>
        </div>
        
        <div class="report-list">
            <h3>📋 보도자료 구성 목록</h3>
            <div class="report-sections">
                <div class="report-section">
                    <h4>요약 보도자료 ({{ summary_reports|length }}개)</h4>
                    <div class="report-items">
                        {% for r in summary_reports %}
                        <span class="report-item">
                            <span class="icon">{{ r.icon }}</span>
                            {{ r.name }}
                        </span>
                        {% endfor %}
                    </div>
                </div>
                
                <div class="report-section">
                    <h4>부문별 보도자료 ({{ sector_reports|length }}개)</h4>
                    <div class="report-items">
                        {% for r in sector_reports %}
                        <span class="report-item">
                            <span class="icon">{{ r.icon }}</span>
                            {{ r.name }}
                        </span>
                        {% endfor %}
                    </div>
                </div>
                
                <div class="report-section">
                    <h4>시도별 보도자료 ({{ regional_reports|length }}개)</h4>
                    <div class="report-items">
                        {% for r in regional_reports %}
                        <span class="report-item">
                            <span class="icon">{{ r.icon }}</span>
                            {{ r.name }}
                        </span>
                        {% endfor %}
                    </div>
                </div>
                
                <div class="report-section">
                    <h4>통계표 ({{ statistics_reports|length }}개)</h4>
                    <div class="report-items">
                        {% for r in statistics_reports %}
                        <span class="report-item">
                            <span class="icon">{{ r.icon }}</span>
                            {{ r.name }}
                        </span>
                        {% endfor %}
                    </div>
                </div>
            </div>
        </div>
        
        <div class="debug-log" id="debug-log">
            <div class="log-entry info">
                <span class="timestamp">[시작]</span>
                디버그 페이지가 로드되었습니다.
            </div>
        </div>
        
        <div class="footer-info">
            <p>국가데이터처 지역경제동향 보도자료 생성 시스템 | 디버그 모드</p>
        </div>
    </div>
    
    <div class="loading-overlay" id="loading-overlay">
        <div class="spinner"></div>
        <div class="loading-text" id="loading-text">보도자료 생성 중...</div>
        <div class="progress-info" id="progress-info"></div>
    </div>
    
    <script>
        async function quickSetup() {
            addLog('자동 세션 설정 중...', 'info');
            showLoading('분석표 파일 찾는 중...');
            
            try {
                const response = await fetch('/debug/set-session', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ year: 2025, quarter: 2 })
                });
                
                const result = await response.json();
                
                if (result.success) {
                    addLog('✓ 세션 설정 완료: ' + result.excel_path, 'info');
                    location.reload();  // 페이지 새로고침
                } else {
                    addLog('✗ 오류: ' + result.error, 'error');
                }
            } catch (error) {
                addLog('✗ 요청 실패: ' + error.message, 'error');
            } finally {
                hideLoading();
            }
        }
        
        function addLog(message, type = 'info') {
            const log = document.getElementById('debug-log');
            const entry = document.createElement('div');
            entry.className = 'log-entry ' + type;
            const now = new Date().toLocaleTimeString('ko-KR');
            entry.innerHTML = '<span class="timestamp">[' + now + ']</span> ' + message;
            log.appendChild(entry);
            log.scrollTop = log.scrollHeight;
        }
        
        function showLoading(text, progress = '') {
            document.getElementById('loading-overlay').classList.add('active');
            document.getElementById('loading-text').textContent = text;
            document.getElementById('progress-info').textContent = progress;
        }
        
        function hideLoading() {
            document.getElementById('loading-overlay').classList.remove('active');
        }
        
        async function generateFullReport() {
            addLog('전체 보도자료 생성 시작...', 'info');
            showLoading('전체 보도자료 생성 중...', '모든 섹션을 처리합니다');
            
            try {
                const response = await fetch('/debug/generate-full-html', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' }
                });
                
                const result = await response.json();
                
                if (result.success) {
                    addLog('✓ 보도자료 생성 완료: ' + result.filename, 'info');
                    addLog('총 ' + result.page_count + '개 페이지, 생성시간: ' + result.generation_time, 'info');
                    
                    // 새 탭에서 열기
                    window.open(result.view_url, '_blank');
                } else {
                    addLog('✗ 오류: ' + result.error, 'error');
                }
            } catch (error) {
                addLog('✗ 요청 실패: ' + error.message, 'error');
            } finally {
                hideLoading();
            }
        }
        
        async function generateSection(section) {
            addLog(section + ' 섹션 생성 시작...', 'info');
            showLoading(section + ' 섹션 생성 중...');
            
            try {
                const response = await fetch('/debug/generate-section-html', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ section: section })
                });
                
                const result = await response.json();
                
                if (result.success) {
                    addLog('✓ ' + section + ' 섹션 생성 완료', 'info');
                    window.open(result.view_url, '_blank');
                } else {
                    addLog('✗ 오류: ' + result.error, 'error');
                }
            } catch (error) {
                addLog('✗ 요청 실패: ' + error.message, 'error');
            } finally {
                hideLoading();
            }
        }
        
        async function generateSingleReport() {
            const reportId = document.getElementById('single-report-id').value.trim();
            if (!reportId) {
                addLog('보도자료 ID를 입력하세요', 'warn');
                return;
            }
            
            addLog(reportId + ' 개별 보도자료 생성 시작...', 'info');
            showLoading(reportId + ' 생성 중...');
            
            try {
                const response = await fetch('/debug/generate-single-html', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ report_id: reportId })
                });
                
                const result = await response.json();
                
                if (result.success) {
                    addLog('✓ ' + reportId + ' 생성 완료', 'info');
                    window.open(result.view_url, '_blank');
                } else {
                    addLog('✗ 오류: ' + result.error, 'error');
                }
            } catch (error) {
                addLog('✗ 요청 실패: ' + error.message, 'error');
            } finally {
                hideLoading();
            }
        }
    </script>
</body>
</html>
'''


# ===== A4 통합 HTML 템플릿 =====
A4_FULL_REPORT_TEMPLATE = '''
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>지역경제동향 {{ year }}년 {{ quarter }}분기 - 디버그 출력</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    
    <!-- ===== DEBUG INFO ===== -->
    <!-- 
    [DEBUG] 생성시간: {{ generation_time }}
    [DEBUG] 총 페이지: {{ page_count }}개
    [DEBUG] 섹션 구성:
    {% for section in sections %}
    - {{ section.name }}: {{ section.count }}개 페이지
    {% endfor %}
    -->
    
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700;900&display=swap');
        
        /* ===== 기본 리셋 ===== */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        /* ===== A4 페이지 설정 (통일된 크기와 여백) ===== */
        @page {
            size: A4;
            margin: 15mm 20mm;
        }
        
        body {
            font-family: '바탕', 'Batang', 'Noto Sans KR', 'Malgun Gothic', sans-serif;
            font-size: 10.5pt;
            line-height: 1.4;
            color: #000;
            background: #f5f5f5;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }
        
        /* ===== A4 페이지 컨테이너 (통일된 여백: 15mm 상하, 20mm 좌우) ===== */
        .a4-page {
            width: 210mm;
            height: 297mm;
            max-height: 297mm;
            background: #fff;
            margin: 20px auto;
            padding: 15mm 20mm;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
            position: relative;
            page-break-after: always;
            page-break-inside: avoid;
            overflow: hidden;
        }
        
        .a4-page:last-child {
            page-break-after: auto;
        }
        
        /* ===== 디버그 오버레이 (화면에만 표시, 인쇄 시 숨김) ===== */
        .debug-overlay {
            position: absolute;
            top: 5mm;
            right: 5mm;
            background: rgba(233, 69, 96, 0.9);
            color: #fff;
            padding: 5px 12px;
            border-radius: 4px;
            font-size: 8pt;
            font-weight: 500;
            z-index: 100;
            font-family: 'Courier New', monospace;
        }
        
        .debug-page-info {
            position: absolute;
            bottom: 5mm;
            left: 5mm;
            background: rgba(0, 0, 0, 0.85);
            color: #fff;
            padding: 6px 12px;
            border-radius: 4px;
            font-size: 7pt;
            font-family: 'Courier New', monospace;
            z-index: 100;
            max-width: 60%;
            line-height: 1.4;
        }
        
        .debug-page-info .debug-id {
            color: #4fc3f7;
        }
        
        .debug-page-info .debug-name {
            color: #fff;
            font-weight: 500;
        }
        
        .debug-page-info .debug-template {
            color: #81c784;
            font-size: 6.5pt;
        }
        
        .debug-page-info .debug-error {
            color: #ef5350;
            font-size: 6.5pt;
        }
        
        /* ===== 페이지 내용 컨테이너 ===== */
        .page-content {
            width: 100%;
            height: 100%;
            font-family: '바탕', 'Batang', 'Times New Roman', serif;
            font-size: 10.5pt;
            line-height: 1.4;
        }
        
        /* 중복된 .page, .cover-container 스타일 무효화 */
        .page-content .page,
        .page-content .cover-container {
            width: auto !important;
            min-height: auto !important;
            height: auto !important;
            padding: 0 !important;
            margin: 0 !important;
            box-shadow: none !important;
        }
        
        .page-content > * {
            max-width: 100%;
        }
        
        /* ===== 공통 섹션 스타일 ===== */
        .section-main-title {
            font-family: '돋움', 'Dotum', sans-serif;
            font-size: 14pt;
            font-weight: bold;
            text-align: center;
            padding: 6px 40px;
            background: #e0e0e0;
            margin-bottom: 18px;
            letter-spacing: 3px;
        }
        
        .section-title {
            font-family: '돋움', 'Dotum', sans-serif;
            font-size: 13pt;
            font-weight: bold;
            margin-bottom: 12px;
        }
        
        .subsection-title {
            font-family: '돋움', 'Dotum', sans-serif;
            font-size: 11pt;
            font-weight: bold;
            margin-bottom: 10px;
        }
        
        /* 요약 박스 */
        .summary-box {
            border: 1px dotted #555;
            padding: 8px 12px;
            margin-bottom: 12px;
            background-color: transparent;
            line-height: 1.6;
        }
        
        /* 증가/감소 표시 */
        .increase { color: #d32f2f; font-weight: bold; }
        .decrease { color: #1976d2; font-weight: bold; }
        
        /* 플레이스홀더 */
        .editable-placeholder {
            background-color: #fff3cd;
            border: 1px dashed #ffc107;
            padding: 0 4px;
            color: #856404;
            min-width: 30px;
            display: inline-block;
        }
        
        /* ===== 섹션 구분선 ===== */
        .section-divider {
            width: 210mm;
            margin: 0 auto;
            padding: 15px 20mm;
            background: linear-gradient(135deg, #1a1a2e, #16213e);
            color: #fff;
            text-align: center;
        }
        
        .section-divider h2 {
            font-size: 1.2rem;
            font-weight: 600;
        }
        
        .section-divider p {
            font-size: 0.85rem;
            opacity: 0.7;
            margin-top: 5px;
        }
        
        /* ===== 인쇄 시 설정 ===== */
        @media print {
            body {
                background: #fff;
            }
            
            .a4-page {
                margin: 0;
                padding: 15mm 20mm;
                box-shadow: none;
                page-break-after: always;
                height: auto;
                max-height: none;
            }
            
            .debug-overlay,
            .debug-page-info,
            .section-divider {
                display: none !important;
            }
        }
        
        /* ===== 디버그 네비게이션 (화면용) ===== */
        .debug-nav {
            position: fixed;
            top: 20px;
            left: 20px;
            background: rgba(26, 26, 46, 0.95);
            border-radius: 15px;
            padding: 15px;
            z-index: 1000;
            max-height: 80vh;
            overflow-y: auto;
            width: 200px;
            box-shadow: 0 10px 40px rgba(0, 0, 0, 0.3);
        }
        
        .debug-nav h4 {
            color: #e94560;
            font-size: 0.9rem;
            margin-bottom: 10px;
            padding-bottom: 8px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
        }
        
        .debug-nav-section {
            margin-bottom: 15px;
        }
        
        .debug-nav-section h5 {
            color: #888;
            font-size: 0.75rem;
            margin-bottom: 5px;
            text-transform: uppercase;
        }
        
        .debug-nav a {
            display: block;
            color: #ccc;
            text-decoration: none;
            padding: 4px 8px;
            font-size: 0.8rem;
            border-radius: 4px;
            transition: all 0.2s;
        }
        
        .debug-nav a:hover {
            background: rgba(233, 69, 96, 0.2);
            color: #e94560;
        }
        
        .debug-toggle {
            position: fixed;
            top: 20px;
            left: 20px;
            background: #e94560;
            color: #fff;
            border: none;
            padding: 10px 15px;
            border-radius: 8px;
            cursor: pointer;
            z-index: 1001;
            font-size: 0.85rem;
        }
        
        .debug-nav.hidden {
            display: none;
        }
        
        @media print {
            .debug-nav,
            .debug-toggle {
                display: none !important;
            }
        }
        
        /* ===== 페이지 내 기본 스타일 재정의 ===== */
        .page-content table {
            border-collapse: collapse;
            width: 100%;
        }
        
        .page-content th,
        .page-content td {
            border: 1px solid #000;
            padding: 4px 6px;
            text-align: center;
            font-size: 9pt;
        }
        
        .page-content th {
            background-color: #e3f2fd;
            font-weight: 500;
        }
        
        .page-content img {
            max-width: 100%;
            height: auto;
        }
        
        /* 차트 컨테이너 */
        .page-content .chart-container {
            position: relative;
            width: 100%;
            max-height: 200px;
        }
        
        .page-content canvas {
            max-width: 100%;
        }
    </style>
</head>
<body>
    <!-- 디버그 네비게이션 -->
    <button class="debug-toggle" onclick="toggleNav()">📋 목차</button>
    <nav class="debug-nav hidden" id="debug-nav">
        <h4>🐛 페이지 네비게이션</h4>
        {% for section in sections %}
        <div class="debug-nav-section">
            <h5>{{ section.name }} ({{ section.count }})</h5>
            {% for page in section.pages %}
            <a href="#page-{{ page.id }}">{{ page.name }}</a>
            {% endfor %}
        </div>
        {% endfor %}
    </nav>
    
    <!-- 페이지 내용 -->
    {% for page in pages %}
    <!-- 
    ===== [DEBUG] 페이지 {{ loop.index }}/{{ page_count }} =====
    ID: {{ page.id }}
    이름: {{ page.name }}
    섹션: {{ page.section }}
    템플릿: {{ page.template or 'N/A' }}
    생성기: {{ page.generator or 'N/A' }}
    {% if page.error %}오류: {{ page.error }}{% endif %}
    ================================
    -->
    <div class="a4-page" id="page-{{ page.id }}">
        <div class="debug-overlay">{{ page.section }} #{{ loop.index }}</div>
        <div class="debug-page-info">
            <span class="debug-id">ID: {{ page.id }}</span> | 
            <span class="debug-name">{{ page.name }}</span>
            {% if page.template %}<br><span class="debug-template">📄 {{ page.template }}</span>{% endif %}
            {% if page.error %}<br><span class="debug-error">⚠️ {{ page.error }}</span>{% endif %}
        </div>
        <div class="page-content">
            {{ page.content|safe }}
        </div>
    </div>
    {% endfor %}
    
    <!-- 
    ===== DEBUG SUMMARY =====
    생성 시간: {{ generation_time }}
    총 페이지 수: {{ page_count }}
    섹션별 페이지:
    {% for section in sections %}
    - {{ section.name }}: {{ section.count }}개
    {% endfor %}
    
    페이지 상세:
    {% for page in pages %}
    {{ loop.index }}. [{{ page.section }}] {{ page.id }} - {{ page.name }}{% if page.template %} ({{ page.template }}){% endif %}{% if page.error %} ❌ ERROR: {{ page.error }}{% endif %}
    {% endfor %}
    ========================
    -->
    
    <script>
        function toggleNav() {
            const nav = document.getElementById('debug-nav');
            nav.classList.toggle('hidden');
        }
        
        // 키보드 단축키 (D: 디버그 네비게이션 토글)
        document.addEventListener('keydown', function(e) {
            if (e.key === 'd' || e.key === 'D') {
                toggleNav();
            }
        });
        
        // 콘솔에 디버그 정보 출력
        console.log('%c🐛 디버그 모드', 'color: #e94560; font-size: 16px; font-weight: bold;');
        console.log('총 페이지: {{ page_count }}개');
        console.log('섹션:', {{ sections|tojson }});
    </script>
</body>
</html>
'''


@debug_bp.route('/')
def debug_page():
    """디버그 페이지 메인"""
    excel_path = session.get('excel_path')
    excel_loaded = excel_path and Path(excel_path).exists()
    year = session.get('year', 2025)
    quarter = session.get('quarter', 2)
    
    total_reports = len(SUMMARY_REPORTS) + len(SECTOR_REPORTS) + len(REGIONAL_REPORTS) + len(STATISTICS_REPORTS)
    # 대략적인 페이지 수 계산 (각 보도자료당 평균 2페이지)
    total_pages = total_reports * 2
    
    return render_template_string(
        DEBUG_PAGE_TEMPLATE,
        excel_loaded=excel_loaded,
        year=year,
        quarter=quarter,
        total_reports=total_reports,
        total_pages=total_pages,
        summary_reports=SUMMARY_REPORTS,
        sector_reports=SECTOR_REPORTS,
        regional_reports=REGIONAL_REPORTS,
        statistics_reports=STATISTICS_REPORTS
    )


@debug_bp.route('/generate-full-html', methods=['POST'])
def generate_full_html():
    """전체 보도자료 HTML 생성"""
    start_time = datetime.now()
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    year = session.get('year', 2025)
    quarter = session.get('quarter', 2)
    raw_excel_path = session.get('raw_excel_path')
    
    try:
        pages = []
        sections = []
        
        # 1. 요약 보도자료
        summary_pages = _generate_summary_pages(excel_path, year, quarter)
        pages.extend(summary_pages)
        sections.append({'name': '요약', 'count': len(summary_pages), 'pages': summary_pages})
        
        # 2. 부문별 보도자료
        sector_pages = _generate_sector_pages(excel_path, year, quarter, raw_excel_path)
        pages.extend(sector_pages)
        sections.append({'name': '부문별', 'count': len(sector_pages), 'pages': sector_pages})
        
        # 3. 시도별 보도자료
        regional_pages = _generate_regional_pages(excel_path, year, quarter)
        pages.extend(regional_pages)
        sections.append({'name': '시도별', 'count': len(regional_pages), 'pages': regional_pages})
        
        # 4. 통계표
        statistics_pages = _generate_statistics_pages(excel_path, year, quarter, raw_excel_path)
        pages.extend(statistics_pages)
        sections.append({'name': '통계표', 'count': len(statistics_pages), 'pages': statistics_pages})
        
        # HTML 생성
        generation_time = (datetime.now() - start_time).total_seconds()
        
        full_html = render_template_string(
            A4_FULL_REPORT_TEMPLATE,
            year=year,
            quarter=quarter,
            pages=pages,
            sections=sections,
            page_count=len(pages),
            generation_time=f"{generation_time:.2f}초"
        )
        
        # 파일 저장 (debug 폴더)
        # 파일명 형식: YYYYMMDD_HHMMSS_full_연도Q분기.html (시간순 정렬 가능)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{timestamp}_full_{year}Q{quarter}.html"
        output_path = DEBUG_FOLDER / filename
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_html)
        
        return jsonify({
            'success': True,
            'filename': filename,
            'view_url': f'/view/{filename}',
            'page_count': len(pages),
            'generation_time': f"{generation_time:.2f}초"
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


@debug_bp.route('/generate-section-html', methods=['POST'])
def generate_section_html():
    """섹션별 HTML 생성"""
    data = request.get_json()
    section = data.get('section')
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    year = session.get('year', 2025)
    quarter = session.get('quarter', 2)
    raw_excel_path = session.get('raw_excel_path')
    
    start_time = datetime.now()
    
    try:
        pages = []
        sections = []
        
        if section == 'summary':
            pages = _generate_summary_pages(excel_path, year, quarter)
            sections.append({'name': '요약', 'count': len(pages), 'pages': pages})
        elif section == 'sector':
            pages = _generate_sector_pages(excel_path, year, quarter, raw_excel_path)
            sections.append({'name': '부문별', 'count': len(pages), 'pages': pages})
        elif section == 'regional':
            pages = _generate_regional_pages(excel_path, year, quarter)
            sections.append({'name': '시도별', 'count': len(pages), 'pages': pages})
        elif section == 'statistics':
            pages = _generate_statistics_pages(excel_path, year, quarter, raw_excel_path)
            sections.append({'name': '통계표', 'count': len(pages), 'pages': pages})
        else:
            return jsonify({'success': False, 'error': f'알 수 없는 섹션: {section}'})
        
        generation_time = (datetime.now() - start_time).total_seconds()
        
        full_html = render_template_string(
            A4_FULL_REPORT_TEMPLATE,
            year=year,
            quarter=quarter,
            pages=pages,
            sections=sections,
            page_count=len(pages),
            generation_time=f"{generation_time:.2f}초"
        )
        
        # 파일명 형식: YYYYMMDD_HHMMSS_섹션명_연도Q분기.html (시간순 정렬 가능)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{timestamp}_{section}_{year}Q{quarter}.html"
        output_path = DEBUG_FOLDER / filename
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_html)
        
        return jsonify({
            'success': True,
            'filename': filename,
            'view_url': f'/view/{filename}',
            'page_count': len(pages),
            'generation_time': f"{generation_time:.2f}초"
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


@debug_bp.route('/generate-single-html', methods=['POST'])
def generate_single_html():
    """개별 보도자료 HTML 생성"""
    data = request.get_json()
    report_id = data.get('report_id')
    
    excel_path = session.get('excel_path')
    if not excel_path or not Path(excel_path).exists():
        return jsonify({'success': False, 'error': '엑셀 파일을 먼저 업로드하세요'})
    
    year = session.get('year', 2025)
    quarter = session.get('quarter', 2)
    raw_excel_path = session.get('raw_excel_path')
    
    try:
        pages = []
        section_name = '개별'
        
        # 요약 보도자료에서 찾기
        report_config = next((r for r in SUMMARY_REPORTS if r['id'] == report_id), None)
        if report_config:
            section_name = '요약'
            html, error, _ = _generate_single_summary(excel_path, report_config, year, quarter)
            if html:
                pages.append({'id': report_id, 'name': report_config['name'], 'section': section_name, 'content': html})
        
        # 부문별 보도자료에서 찾기
        if not pages:
            report_config = next((r for r in SECTOR_REPORTS if r['id'] == report_id), None)
            if report_config:
                section_name = '부문별'
                html, error, _ = generate_report_html(excel_path, report_config, year, quarter, None, raw_excel_path)
                if html:
                    pages.append({'id': report_id, 'name': report_config['name'], 'section': section_name, 'content': html})
        
        # 시도별 보도자료에서 찾기
        if not pages:
            region_config = next((r for r in REGIONAL_REPORTS if r['id'] == report_id), None)
            if region_config:
                section_name = '시도별'
                is_reference = region_config.get('is_reference', False)
                html, error = generate_regional_report_html(excel_path, region_config['name'], is_reference)
                if html:
                    pages.append({'id': report_id, 'name': region_config['name'], 'section': section_name, 'content': html})
        
        # 통계표에서 찾기
        if not pages:
            stat_config = next((s for s in STATISTICS_REPORTS if s['id'] == report_id), None)
            if stat_config:
                section_name = '통계표'
                html, error = generate_individual_statistics_html(excel_path, stat_config, year, quarter, raw_excel_path)
                if html:
                    pages.append({'id': report_id, 'name': stat_config['name'], 'section': section_name, 'content': html})
        
        if not pages:
            return jsonify({'success': False, 'error': f'보도자료를 찾을 수 없습니다: {report_id}'})
        
        sections = [{'name': section_name, 'count': len(pages), 'pages': pages}]
        
        full_html = render_template_string(
            A4_FULL_REPORT_TEMPLATE,
            year=year,
            quarter=quarter,
            pages=pages,
            sections=sections,
            page_count=len(pages),
            generation_time="0.1초"
        )
        
        # 파일명 형식: YYYYMMDD_HHMMSS_single_보도자료ID.html (시간순 정렬 가능)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{timestamp}_single_{report_id}.html"
        output_path = DEBUG_FOLDER / filename
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_html)
        
        return jsonify({
            'success': True,
            'filename': filename,
            'view_url': f'/view/{filename}',
            'page_count': len(pages)
        })
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)})


def _generate_summary_pages(excel_path, year, quarter):
    """요약 보도자료 페이지 생성"""
    pages = []
    
    for report in SUMMARY_REPORTS:
        try:
            html, error, _ = _generate_single_summary(excel_path, report, year, quarter)
            if html:
                # HTML 컨텐츠 정제 (body 내용만 추출)
                content, _ = extract_body_content(html)
                pages.append({
                    'id': report['id'],
                    'name': report['name'],
                    'section': '요약',
                    'template': report.get('template', ''),
                    'generator': report.get('generator', ''),
                    'content': content
                })
            else:
                # 에러 발생 시 플레이스홀더 페이지
                pages.append({
                    'id': report['id'],
                    'name': report['name'],
                    'section': '요약',
                    'template': report.get('template', ''),
                    'generator': report.get('generator', ''),
                    'error': error or '생성 실패',
                    'content': f'<div style="padding: 50px; text-align: center; color: #999;"><h3>⚠️ {report["name"]}</h3><p>{error or "생성 실패"}</p></div>'
                })
        except Exception as e:
            pages.append({
                'id': report['id'],
                'name': report['name'],
                'section': '요약',
                'template': report.get('template', ''),
                'generator': report.get('generator', ''),
                'error': str(e),
                'content': f'<div style="padding: 50px; text-align: center; color: #f00;"><h3>❌ {report["name"]}</h3><p>오류: {str(e)}</p></div>'
            })
    
    return pages


def _generate_single_summary(excel_path, report_config, year, quarter):
    """단일 요약 보도자료 생성 (preview.py와 동일한 로직 사용)"""
    try:
        template_name = report_config['template']
        generator_name = report_config.get('generator')
        report_id = report_config['id']
        
        # 템플릿 파일 존재 확인
        template_path = TEMPLATES_DIR / template_name
        if not template_path.exists():
            error_msg = f"템플릿 파일을 찾을 수 없습니다: {template_name}"
            print(f"[DEBUG] {error_msg}")
            return None, error_msg, []
        
        report_data = {
            'report_info': {
                'year': year,
                'quarter': quarter,
                'organization': '국가데이터처',
                'department': '경제통계심의관'
            }
        }
        
        # Generator를 통한 데이터 생성 (인포그래픽 등)
        if generator_name:
            try:
                module = load_generator_module(generator_name)
                if module is None:
                    error_msg = f"Generator 모듈을 로드할 수 없습니다: {generator_name}"
                    print(f"[DEBUG] {error_msg}")
                    return None, error_msg, []
                
                if hasattr(module, 'generate_report_data'):
                    try:
                        generated_data = module.generate_report_data(excel_path)
                        if generated_data:
                            report_data.update(generated_data)
                            print(f"[DEBUG] Generator 데이터 생성 성공: {generator_name}")
                        else:
                            print(f"[DEBUG] Generator가 빈 데이터를 반환했습니다: {generator_name}")
                    except Exception as e:
                        import traceback
                        error_msg = f"Generator 데이터 생성 오류 ({generator_name}): {str(e)}"
                        print(f"[DEBUG] {error_msg}")
                        traceback.print_exc()
                        return None, error_msg, []
                else:
                    print(f"[DEBUG] Generator에 generate_report_data 함수가 없습니다: {generator_name}")
            except Exception as e:
                import traceback
                error_msg = f"Generator 모듈 로드 오류 ({generator_name}): {str(e)}"
                print(f"[DEBUG] {error_msg}")
                traceback.print_exc()
                return None, error_msg, []
        
        # 템플릿별 데이터 제공 (preview.py와 동일)
        if report_id == 'toc':
            # 목차는 고정된 HTML 템플릿 사용 (동적 계산 없음)
            print(f"[DEBUG] 목차 템플릿 로드 (고정 페이지 번호)")
        elif report_id == 'guide':
            try:
                report_data.update(_get_guide_data(year, quarter))
                print(f"[DEBUG] 일러두기 데이터 생성 완료")
            except Exception as e:
                import traceback
                error_msg = f"일러두기 데이터 생성 오류: {str(e)}"
                print(f"[DEBUG] {error_msg}")
                traceback.print_exc()
                return None, error_msg, []
        elif report_id == 'summary_overview':
            try:
                report_data['summary'] = get_summary_overview_data(excel_path, year, quarter)
                report_data['table_data'] = get_summary_table_data(excel_path, year, quarter)
                report_data['page_number'] = 1
            except Exception as e:
                import traceback
                error_msg = f"요약-지역경제동향 데이터 생성 오류: {str(e)}"
                print(f"[DEBUG] {error_msg}")
                traceback.print_exc()
                return None, error_msg, []
        elif report_id == 'summary_production':
            try:
                report_data.update(get_production_summary_data(excel_path, year, quarter))
                report_data['page_number'] = 2
            except Exception as e:
                import traceback
                error_msg = f"요약-생산 데이터 생성 오류: {str(e)}"
                print(f"[DEBUG] {error_msg}")
                traceback.print_exc()
                return None, error_msg, []
        elif report_id == 'summary_consumption':
            try:
                report_data.update(get_consumption_construction_data(excel_path, year, quarter))
                report_data['page_number'] = 3
            except Exception as e:
                import traceback
                error_msg = f"요약-소비건설 데이터 생성 오류: {str(e)}"
                print(f"[DEBUG] {error_msg}")
                traceback.print_exc()
                return None, error_msg, []
        elif report_id == 'summary_trade_price':
            try:
                report_data.update(get_trade_price_data(excel_path, year, quarter))
                report_data['page_number'] = 4
            except Exception as e:
                import traceback
                error_msg = f"요약-수출물가 데이터 생성 오류: {str(e)}"
                print(f"[DEBUG] {error_msg}")
                traceback.print_exc()
                return None, error_msg, []
        elif report_id == 'summary_employment':
            try:
                report_data.update(get_employment_population_data(excel_path, year, quarter))
                report_data['page_number'] = 5
            except Exception as e:
                import traceback
                error_msg = f"요약-고용인구 데이터 생성 오류: {str(e)}"
                print(f"[DEBUG] {error_msg}")
                traceback.print_exc()
                return None, error_msg, []
        
        # 기본 연락처 정보
        report_data['release_info'] = {
            'release_datetime': f'{year}. 8. 12.(화) 12:00',
            'distribution_datetime': f'{year}. 8. 12.(화) 08:30'
        }
        report_data['contact_info'] = {
            'department': '국가데이터처 경제통계국',
            'division': '소득통계과',
            'manager_title': '과 장',
            'manager_name': '정선경',
            'manager_phone': '042-481-2206',
            'staff_title': '사무관',
            'staff_name': '윤민희',
            'staff_phone': '042-481-2226'
        }
        
        # 템플릿 렌더링
        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                template = Template(f.read())
            
            html_content = template.render(**report_data)
            print(f"[DEBUG] {report_id} 템플릿 렌더링 완료: {template_name}")
            return html_content, None, []
        except Exception as e:
            import traceback
            error_msg = f"템플릿 렌더링 오류 ({template_name}): {str(e)}"
            print(f"[DEBUG] {error_msg}")
            traceback.print_exc()
            return None, error_msg, []
        
    except Exception as e:
        import traceback
        error_msg = f"보도자료 생성 오류 ({report_config.get('name', 'unknown')}): {str(e)}"
        print(f"[DEBUG] {error_msg}")
        traceback.print_exc()
        return None, error_msg, []


def _generate_sector_pages(excel_path, year, quarter, raw_excel_path=None):
    """부문별 보도자료 페이지 생성"""
    pages = []
    
    for report in SECTOR_REPORTS:
        try:
            html, error, _ = generate_report_html(excel_path, report, year, quarter, None, raw_excel_path)
            if html:
                # HTML 컨텐츠 정제 (body 내용만 추출)
                content, _ = extract_body_content(html)
                pages.append({
                    'id': report['id'],
                    'name': report['name'],
                    'section': '부문별',
                    'template': report.get('template', ''),
                    'generator': report.get('generator', ''),
                    'content': content
                })
            else:
                pages.append({
                    'id': report['id'],
                    'name': report['name'],
                    'section': '부문별',
                    'template': report.get('template', ''),
                    'generator': report.get('generator', ''),
                    'error': error or '생성 실패',
                    'content': f'<div style="padding: 50px; text-align: center; color: #999;"><h3>⚠️ {report["name"]}</h3><p>{error or "생성 실패"}</p></div>'
                })
        except Exception as e:
            pages.append({
                'id': report['id'],
                'name': report['name'],
                'section': '부문별',
                'template': report.get('template', ''),
                'generator': report.get('generator', ''),
                'error': str(e),
                'content': f'<div style="padding: 50px; text-align: center; color: #f00;"><h3>❌ {report["name"]}</h3><p>오류: {str(e)}</p></div>'
            })
    
    return pages


def _generate_regional_pages(excel_path, year, quarter):
    """시도별 보도자료 페이지 생성"""
    pages = []
    
    for region in REGIONAL_REPORTS:
        try:
            is_reference = region.get('is_reference', False)
            html, error = generate_regional_report_html(excel_path, region['name'], is_reference)
            if html:
                # HTML 컨텐츠 정제 (body 내용만 추출)
                content, _ = extract_body_content(html)
                pages.append({
                    'id': region['id'],
                    'name': region['name'],
                    'section': '시도별',
                    'template': 'regional_template.html' if not is_reference else 'grdp_reference_template.html',
                    'is_reference': is_reference,
                    'content': content
                })
            else:
                pages.append({
                    'id': region['id'],
                    'name': region['name'],
                    'section': '시도별',
                    'template': 'regional_template.html' if not is_reference else 'grdp_reference_template.html',
                    'is_reference': is_reference,
                    'error': error or '생성 실패',
                    'content': f'<div style="padding: 50px; text-align: center; color: #999;"><h3>⚠️ {region["name"]}</h3><p>{error or "생성 실패"}</p></div>'
                })
        except Exception as e:
            pages.append({
                'id': region['id'],
                'name': region['name'],
                'section': '시도별',
                'template': 'regional_template.html' if not region.get('is_reference', False) else 'grdp_reference_template.html',
                'is_reference': region.get('is_reference', False),
                'error': str(e),
                'content': f'<div style="padding: 50px; text-align: center; color: #f00;"><h3>❌ {region["name"]}</h3><p>오류: {str(e)}</p></div>'
            })
    
    return pages


def _generate_statistics_pages(excel_path, year, quarter, raw_excel_path=None):
    """통계표 페이지 생성"""
    pages = []
    
    for stat in STATISTICS_REPORTS:
        try:
            html, error = generate_individual_statistics_html(excel_path, stat, year, quarter, raw_excel_path)
            if html:
                # HTML 컨텐츠 정제 (body 내용만 추출)
                content, _ = extract_body_content(html)
                pages.append({
                    'id': stat['id'],
                    'name': stat['name'],
                    'section': '통계표',
                    'template': stat.get('template', ''),
                    'generator': stat.get('generator', ''),
                    'content': content
                })
            else:
                pages.append({
                    'id': stat['id'],
                    'name': stat['name'],
                    'section': '통계표',
                    'template': stat.get('template', ''),
                    'generator': stat.get('generator', ''),
                    'error': error or '생성 실패',
                    'content': f'<div style="padding: 50px; text-align: center; color: #999;"><h3>⚠️ {stat["name"]}</h3><p>{error or "생성 실패"}</p></div>'
                })
        except Exception as e:
            pages.append({
                'id': stat['id'],
                'name': stat['name'],
                'section': '통계표',
                'template': stat.get('template', ''),
                'generator': stat.get('generator', ''),
                'error': str(e),
                'content': f'<div style="padding: 50px; text-align: center; color: #f00;"><h3>❌ {stat["name"]}</h3><p>오류: {str(e)}</p></div>'
            })
    
    return pages


def _get_toc_sections():
    """목차 섹션 데이터 - 페이지 단위로 동적 계산
    
    같은 항목이 여러 페이지인 경우 (1), (2) 등으로 구분
    """
    
    # 현재 페이지 번호 (요약부터 1페이지 시작)
    current_page = 1
    
    # 요약 섹션 시작 페이지
    summary_page = current_page
    summary_pages = sum(PAGE_CONFIG['summary'].values())
    current_page += summary_pages
    
    # 부문별 섹션 시작 페이지
    sector_page = current_page
    
    # 부문별 각 항목의 시작 페이지 계산 (페이지 단위)
    sector_entries = []
    sector_config = PAGE_CONFIG['sector']
    sector_order = ['manufacturing', 'service', 'consumption', 'construction', 
                    'export', 'import', 'price', 'employment', 'unemployment', 'population']
    
    # TOC_SECTOR_ITEMS 기반으로 통합 항목 처리
    item_pages_map = {}  # start_from -> [페이지번호들]
    for sector_id in sector_order:
        pages_count = sector_config.get(sector_id, 1)
        if sector_id not in item_pages_map:
            item_pages_map[sector_id] = []
        for i in range(pages_count):
            item_pages_map[sector_id].append(current_page + i)
        current_page += pages_count
    
    # 부문별 목차 항목 생성 (페이지 단위)
    entry_number = 1
    for item in TOC_SECTOR_ITEMS:
        start_from = item.get('start_from')
        pages = item_pages_map.get(start_from, [])
        
        if len(pages) == 1:
            # 1페이지짜리 항목
            sector_entries.append({
                'number': entry_number,
                'name': item['name'],
                'page': pages[0]
            })
            entry_number += 1
        else:
            # 여러 페이지인 경우 (1), (2) 등으로 구분
            for idx, page in enumerate(pages, 1):
                sector_entries.append({
                    'number': entry_number,
                    'name': f"{item['name']} ({idx})",
                    'page': page
                })
                entry_number += 1
    
    # 시도별 섹션 시작 페이지
    region_page = current_page
    
    # 시도별 목차 항목 생성 (페이지 단위)
    region_entries = []
    regional_pages = PAGE_CONFIG['regional']  # 각 시도당 페이지 수 (2)
    entry_number = 1
    
    for item in TOC_REGION_ITEMS:
        if regional_pages == 1:
            # 1페이지짜리 항목
            region_entries.append({
                'number': entry_number,
                'name': item['name'],
                'page': current_page
            })
            entry_number += 1
            current_page += 1
        else:
            # 여러 페이지인 경우 (1), (2) 등으로 구분
            for idx in range(1, regional_pages + 1):
                region_entries.append({
                    'number': entry_number,
                    'name': f"{item['name']} ({idx})",
                    'page': current_page
                })
                entry_number += 1
                current_page += 1
    
    # 참고 GRDP 페이지 (2페이지인 경우 구분)
    reference_page = current_page
    grdp_pages = PAGE_CONFIG['reference_grdp']
    reference_entries = []
    if grdp_pages > 1:
        for idx in range(1, grdp_pages + 1):
            reference_entries.append({
                'name': f'분기GRDP ({idx})',
                'page': current_page
            })
            current_page += 1
    else:
        current_page += grdp_pages
    
    # 통계표 섹션 시작 페이지
    statistics_page = current_page
    stat_config = PAGE_CONFIG['statistics']
    current_page += stat_config['toc']
    
    # 통계표 목차 항목 생성 (페이지 단위, 각 통계표 2페이지)
    statistics_entries = []
    stat_names = ['광공업생산지수', '서비스업생산지수', '소매판매액지수', '건설수주액',
                  '고용률', '실업률', '국내인구이동', '수출액', '수입액', '소비자물가지수', 'GRDP']
    entry_number = 1
    for stat_name in stat_names:
        pages_per_table = 2 if stat_name != 'GRDP' else 2  # 모든 통계표 2페이지
        for idx in range(1, pages_per_table + 1):
            statistics_entries.append({
                'number': entry_number,
                'name': f'{stat_name} ({idx})',
                'page': current_page
            })
            entry_number += 1
            current_page += 1
    
    # 부록 페이지
    appendix_page = current_page
    
    return {
        'summary': {'page': summary_page},
        'sector': {
            'page': sector_page,
            'entries': sector_entries
        },
        'region': {
            'page': region_page,
            'entries': region_entries
        },
        'reference': {
            'name': '분기GRDP', 
            'page': reference_page,
            'entries': reference_entries if reference_entries else None
        },
        'statistics': {
            'page': statistics_page,
            'entries': statistics_entries
        },
        'appendix': {'page': appendix_page}
    }


def _get_guide_data(year, quarter):
    """일러두기 데이터"""
    return {
        'intro': {
            'background': '지역경제동향은 시·도별 경제 현황을 생산, 소비, 건설, 수출입, 물가, 고용, 인구 등의 주요 경제지표를 통하여 분석한 자료입니다.',
            'purpose': '지역경제의 동향 파악과 지역개발정책 수립 및 평가의 기초자료로 활용하고자 작성합니다.'
        },
        'content': {
            'description': f'본 보도자료는 {year}년 {quarter}/4분기 시·도별 지역경제동향을 수록하였습니다.',
            'indicator_note': '수록 지표는 총 7개 부문으로 다음과 같습니다.',
            'indicators': [
                {'type': '생산', 'stat_items': ['광공업생산지수', '서비스업생산지수']},
                {'type': '소비', 'stat_items': ['소매판매액지수']},
                {'type': '건설', 'stat_items': ['건설수주액']},
                {'type': '수출입', 'stat_items': ['수출액', '수입액']},
                {'type': '물가', 'stat_items': ['소비자물가지수']},
                {'type': '고용', 'stat_items': ['고용률', '실업률']},
                {'type': '인구', 'stat_items': ['국내인구이동']}
            ]
        },
        'contacts': [
            {'category': '생산', 'statistics_name': '광공업생산지수', 'department': '광업제조업동향과', 'phone': '042-481-2183'},
            {'category': '생산', 'statistics_name': '서비스업생산지수', 'department': '서비스업동향과', 'phone': '042-481-2196'},
            {'category': '소비', 'statistics_name': '소매판매액지수', 'department': '서비스업동향과', 'phone': '042-481-2199'},
            {'category': '건설', 'statistics_name': '건설수주액', 'department': '건설동향과', 'phone': '042-481-2556'},
            {'category': '수출입', 'statistics_name': '수출입액', 'department': '관세청 정보데이터기획담당관', 'phone': '042-481-7845'},
            {'category': '물가', 'statistics_name': '소비자물가지수', 'department': '물가동향과', 'phone': '042-481-2532'},
            {'category': '고용', 'statistics_name': '고용률, 실업률', 'department': '고용통계과', 'phone': '042-481-2264'},
            {'category': '인구', 'statistics_name': '국내인구이동', 'department': '인구동향과', 'phone': '042-481-2252'}
        ],
        'references': [
            {'content': '본 자료는 국가데이터처 홈페이지(http://kostat.go.kr)에서 확인하실 수 있습니다.'},
            {'content': '관련 통계표는 KOSIS(국가통계포털, http://kosis.kr)에서 이용하실 수 있습니다.'}
        ],
        'notes': [
            '자료에 수록된 값은 잠정치이므로 추후 수정될 수 있습니다.'
        ]
    }

