# -*- coding: utf-8 -*-
"""
마크다운을 A4 양면인쇄용 HTML로 변환하는 스크립트

요구사항:
1. 차트는 비율을 유지한 이미지 파일로 삽입
2. 출력시 레이아웃 무너지면 안됨
3. 그레이스케일 출력
4. 표나 차트가 쪼개지면 강제 개행
5. 공백 부분 최소화
6. 기본 12pt 나눔고딕
7. HTML과 출력 모두 레이아웃 유지

Mermaid 변환 방법:
1. mermaid-cli (mmdc) - npm install -g @mermaid-js/mermaid-cli
2. kroki.io API - 네트워크 필요
3. playwright - pip install playwright && playwright install chromium
"""

import re
import os
import base64
import hashlib
import subprocess
import tempfile
import shutil
from pathlib import Path
from datetime import datetime

# Mermaid 변환 설정
MERMAID_CONVERTER = None  # 'mmdc', 'kroki', 'playwright', None(자동 감지)
MERMAID_IMAGES_DIR = None  # 이미지 저장 디렉토리


def check_mermaid_cli():
    """mermaid-cli (mmdc) 설치 여부 확인"""
    try:
        result = subprocess.run(['mmdc', '--version'], capture_output=True, text=True, timeout=5)
        return result.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def check_playwright():
    """playwright 설치 여부 확인"""
    try:
        from playwright.sync_api import sync_playwright
        return True
    except ImportError:
        return False


def convert_mermaid_with_mmdc(mermaid_code, output_path):
    """mermaid-cli로 변환"""
    with tempfile.NamedTemporaryFile(mode='w', suffix='.mmd', delete=False, encoding='utf-8') as f:
        f.write(mermaid_code)
        input_path = f.name
    
    try:
        result = subprocess.run([
            'mmdc', '-i', input_path, '-o', str(output_path),
            '-b', 'white', '-t', 'default', '-w', '800'
        ], capture_output=True, text=True, timeout=30)
        
        if result.returncode == 0 and Path(output_path).exists():
            return True
        print(f"  ⚠️ mmdc 변환 실패: {result.stderr}")
        return False
    except subprocess.TimeoutExpired:
        print("  ⚠️ mmdc 타임아웃")
        return False
    finally:
        os.unlink(input_path)


def convert_mermaid_with_kroki(mermaid_code, output_path):
    """kroki.io API로 변환"""
    try:
        import urllib.request
        import zlib
        
        # Mermaid 코드를 압축하고 base64 인코딩
        compressed = zlib.compress(mermaid_code.encode('utf-8'), 9)
        encoded = base64.urlsafe_b64encode(compressed).decode('ascii')
        
        # Kroki API 호출
        url = f'https://kroki.io/mermaid/png/{encoded}'
        
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=30) as response:
            with open(output_path, 'wb') as f:
                f.write(response.read())
        
        if Path(output_path).exists() and Path(output_path).stat().st_size > 0:
            return True
        return False
    except Exception as e:
        print(f"  ⚠️ kroki.io 변환 실패: {e}")
        return False


def convert_mermaid_with_playwright(mermaid_code, output_path):
    """playwright로 브라우저 렌더링 후 스크린샷"""
    try:
        from playwright.sync_api import sync_playwright
        
        html_content = f'''<!DOCTYPE html>
<html>
<head>
    <script src="https://cdn.jsdelivr.net/npm/mermaid/dist/mermaid.min.js"></script>
    <style>
        body {{ margin: 0; padding: 20px; background: white; }}
        .mermaid {{ background: white; }}
    </style>
</head>
<body>
    <div class="mermaid">
{mermaid_code}
    </div>
    <script>
        mermaid.initialize({{ startOnLoad: true, theme: 'default' }});
    </script>
</body>
</html>'''
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as f:
            f.write(html_content)
            html_path = f.name
        
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch()
                page = browser.new_page()
                page.goto(f'file://{html_path}')
                page.wait_for_timeout(2000)  # Mermaid 렌더링 대기
                
                # 다이어그램 요소 찾기
                element = page.query_selector('.mermaid svg')
                if element:
                    element.screenshot(path=str(output_path))
                else:
                    page.screenshot(path=str(output_path), full_page=True)
                
                browser.close()
            
            return Path(output_path).exists()
        finally:
            os.unlink(html_path)
    except Exception as e:
        print(f"  ⚠️ playwright 변환 실패: {e}")
        return False


def convert_mermaid_to_image(mermaid_code, output_path, method=None):
    """Mermaid 다이어그램을 이미지로 변환
    
    Args:
        mermaid_code: Mermaid 다이어그램 코드
        output_path: 출력 이미지 경로
        method: 변환 방법 ('mmdc', 'kroki', 'playwright', None=자동)
    
    Returns:
        bool: 변환 성공 여부
    """
    global MERMAID_CONVERTER
    
    if method is None:
        method = MERMAID_CONVERTER
    
    # 자동 감지
    if method is None:
        if check_mermaid_cli():
            method = 'mmdc'
        elif check_playwright():
            method = 'playwright'
        else:
            method = 'kroki'  # 기본값 (네트워크 필요)
    
    print(f"  🔄 Mermaid 변환 중 ({method})...")
    
    if method == 'mmdc':
        return convert_mermaid_with_mmdc(mermaid_code, output_path)
    elif method == 'playwright':
        return convert_mermaid_with_playwright(mermaid_code, output_path)
    elif method == 'kroki':
        return convert_mermaid_with_kroki(mermaid_code, output_path)
    else:
        print(f"  ⚠️ 알 수 없는 변환 방법: {method}")
        return False

def escape_html(text):
    """HTML 특수문자 이스케이프"""
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))

def parse_markdown_to_html(md_content, images_dir=None, convert_mermaid=True):
    """마크다운을 HTML로 변환
    
    Args:
        md_content: 마크다운 내용
        images_dir: Mermaid 이미지 저장 디렉토리 (Path 객체)
        convert_mermaid: Mermaid 다이어그램 이미지 변환 여부
    """
    
    lines = md_content.split('\n')
    html_parts = []
    in_code_block = False
    code_block_lang = ''
    code_content = []
    in_table = False
    table_rows = []
    in_list = False
    list_type = None
    list_items = []
    in_blockquote = False
    blockquote_content = []
    mermaid_counter = 0
    
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # 코드 블록 처리
        if line.startswith('```'):
            if in_code_block:
                # 코드 블록 종료
                code_text = '\n'.join(code_content)
                
                if code_block_lang == 'mermaid':
                    mermaid_counter += 1
                    
                    # Mermaid 이미지 변환 시도
                    image_converted = False
                    image_html = ''
                    
                    if convert_mermaid and images_dir:
                        image_filename = f'mermaid_{mermaid_counter}.png'
                        image_path = images_dir / image_filename
                        
                        print(f"📊 다이어그램 {mermaid_counter} 변환 중...")
                        
                        if convert_mermaid_to_image(code_text, image_path):
                            # 이미지를 base64로 인코딩하여 HTML에 포함
                            with open(image_path, 'rb') as img_file:
                                img_data = base64.b64encode(img_file.read()).decode('utf-8')
                            
                            image_html = f'''
<div class="mermaid-image" style="page-break-inside: avoid; margin: 15px 0; text-align: center;">
    <img src="data:image/png;base64,{img_data}" alt="다이어그램 {mermaid_counter}" style="max-width: 100%; height: auto; border: 1px solid #ddd;">
</div>'''
                            image_converted = True
                            print(f"  ✅ 다이어그램 {mermaid_counter} 변환 완료")
                        else:
                            print(f"  ❌ 다이어그램 {mermaid_counter} 변환 실패 - 플레이스홀더 사용")
                    
                    if not image_converted:
                        # 플레이스홀더 (변환 실패 또는 비활성화)
                        image_html = f'''
<div class="mermaid-placeholder" style="page-break-inside: avoid;">
    <div style="background: #f5f5f5; border: 2px dashed #999; padding: 20px; text-align: center; margin: 15px 0;">
        <p style="color: #666; font-size: 11pt; margin: 0;">[다이어그램 {mermaid_counter}]</p>
        <p style="color: #999; font-size: 9pt; margin: 5px 0 0 0;">Mermaid 다이어그램</p>
        <pre style="text-align: left; font-size: 8pt; background: #fff; padding: 10px; margin-top: 10px; overflow: auto; max-height: 200px;">{escape_html(code_text)}</pre>
    </div>
</div>'''
                    
                    html_parts.append(image_html)
                else:
                    # 일반 코드 블록
                    html_parts.append(f'''
<div class="code-block" style="page-break-inside: avoid;">
    <pre><code class="language-{code_block_lang}">{escape_html(code_text)}</code></pre>
</div>''')
                
                in_code_block = False
                code_block_lang = ''
                code_content = []
            else:
                # 코드 블록 시작
                in_code_block = True
                code_block_lang = line[3:].strip() or 'text'
            i += 1
            continue
        
        if in_code_block:
            code_content.append(line)
            i += 1
            continue
        
        # 빈 줄 처리
        if not line.strip():
            # 리스트 종료
            if in_list:
                tag = 'ol' if list_type == 'ol' else 'ul'
                html_parts.append(f'<{tag}>{"".join(list_items)}</{tag}>')
                in_list = False
                list_items = []
            # 인용 종료
            if in_blockquote:
                html_parts.append(f'<blockquote>{"".join(blockquote_content)}</blockquote>')
                in_blockquote = False
                blockquote_content = []
            # 표 종료
            if in_table:
                html_parts.append(build_table(table_rows))
                in_table = False
                table_rows = []
            i += 1
            continue
        
        # 인용문 처리
        if line.startswith('>'):
            quote_text = line[1:].strip()
            if not in_blockquote:
                in_blockquote = True
            blockquote_content.append(f'<p>{process_inline(quote_text)}</p>')
            i += 1
            continue
        
        if in_blockquote:
            html_parts.append(f'<blockquote>{"".join(blockquote_content)}</blockquote>')
            in_blockquote = False
            blockquote_content = []
        
        # 표 처리
        if '|' in line and line.strip().startswith('|'):
            if not in_table:
                in_table = True
            table_rows.append(line)
            i += 1
            continue
        
        if in_table:
            html_parts.append(build_table(table_rows))
            in_table = False
            table_rows = []
        
        # 수평선
        if re.match(r'^-{3,}$|^\*{3,}$|^_{3,}$', line.strip()):
            if in_list:
                tag = 'ol' if list_type == 'ol' else 'ul'
                html_parts.append(f'<{tag}>{"".join(list_items)}</{tag}>')
                in_list = False
                list_items = []
            html_parts.append('<hr class="page-break-suggestion">')
            i += 1
            continue
        
        # 제목 처리
        heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
        if heading_match:
            if in_list:
                tag = 'ol' if list_type == 'ol' else 'ul'
                html_parts.append(f'<{tag}>{"".join(list_items)}</{tag}>')
                in_list = False
                list_items = []
            
            level = len(heading_match.group(1))
            text = heading_match.group(2)
            # ID 생성 (한글 포함)
            heading_id = re.sub(r'[^\w\s가-힣-]', '', text.lower()).replace(' ', '-')
            
            # h1, h2는 페이지 나눔 고려
            page_break = 'page-break-before: auto;' if level <= 2 else ''
            html_parts.append(f'<h{level} id="{heading_id}" style="{page_break}">{process_inline(text)}</h{level}>')
            i += 1
            continue
        
        # 리스트 처리
        list_match = re.match(r'^(\s*)[-*+]\s+(.+)$', line)
        ol_match = re.match(r'^(\s*)(\d+)\.\s+(.+)$', line)
        
        if list_match:
            indent = len(list_match.group(1))
            item_text = list_match.group(2)
            if not in_list or list_type != 'ul':
                if in_list:
                    tag = 'ol' if list_type == 'ol' else 'ul'
                    html_parts.append(f'<{tag}>{"".join(list_items)}</{tag}>')
                    list_items = []
                in_list = True
                list_type = 'ul'
            list_items.append(f'<li>{process_inline(item_text)}</li>')
            i += 1
            continue
        
        if ol_match:
            indent = len(ol_match.group(1))
            item_text = ol_match.group(3)
            if not in_list or list_type != 'ol':
                if in_list:
                    tag = 'ol' if list_type == 'ol' else 'ul'
                    html_parts.append(f'<{tag}>{"".join(list_items)}</{tag}>')
                    list_items = []
                in_list = True
                list_type = 'ol'
            list_items.append(f'<li>{process_inline(item_text)}</li>')
            i += 1
            continue
        
        # 리스트 종료 후 일반 텍스트
        if in_list:
            tag = 'ol' if list_type == 'ol' else 'ul'
            html_parts.append(f'<{tag}>{"".join(list_items)}</{tag}>')
            in_list = False
            list_items = []
        
        # 일반 문단
        if line.strip():
            html_parts.append(f'<p>{process_inline(line)}</p>')
        
        i += 1
    
    # 마지막 리스트/인용/표 처리
    if in_list:
        tag = 'ol' if list_type == 'ol' else 'ul'
        html_parts.append(f'<{tag}>{"".join(list_items)}</{tag}>')
    if in_blockquote:
        html_parts.append(f'<blockquote>{"".join(blockquote_content)}</blockquote>')
    if in_table:
        html_parts.append(build_table(table_rows))
    if in_code_block:
        code_text = '\n'.join(code_content)
        html_parts.append(f'<pre><code>{escape_html(code_text)}</code></pre>')
    
    return '\n'.join(html_parts)

def process_inline(text):
    """인라인 마크다운 처리"""
    # 굵은 글씨 (** 또는 __)
    text = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)
    text = re.sub(r'__(.+?)__', r'<strong>\1</strong>', text)
    
    # 기울임 (* 또는 _)
    text = re.sub(r'\*(.+?)\*', r'<em>\1</em>', text)
    text = re.sub(r'_(.+?)_', r'<em>\1</em>', text)
    
    # 인라인 코드
    text = re.sub(r'`([^`]+)`', r'<code class="inline-code">\1</code>', text)
    
    # 링크
    text = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'<a href="\2">\1</a>', text)
    
    # 이미지
    text = re.sub(r'!\[([^\]]*)\]\(([^)]+)\)', r'<img src="\2" alt="\1" style="max-width: 100%; height: auto;">', text)
    
    # 취소선
    text = re.sub(r'~~(.+?)~~', r'<del>\1</del>', text)
    
    return text

def build_table(rows):
    """표 생성"""
    if len(rows) < 2:
        return ''
    
    html = ['<div class="table-container" style="page-break-inside: avoid;"><table>']
    
    # 헤더 행
    header_cells = [cell.strip() for cell in rows[0].split('|')[1:-1]]
    html.append('<thead><tr>')
    for cell in header_cells:
        html.append(f'<th>{process_inline(cell)}</th>')
    html.append('</tr></thead>')
    
    # 구분선 행 건너뛰기 (rows[1])
    
    # 데이터 행
    html.append('<tbody>')
    for row in rows[2:]:
        cells = [cell.strip() for cell in row.split('|')[1:-1]]
        html.append('<tr>')
        for cell in cells:
            html.append(f'<td>{process_inline(cell)}</td>')
        html.append('</tr>')
    html.append('</tbody>')
    
    html.append('</table></div>')
    return '\n'.join(html)

def generate_print_html(md_content, title="발표자료", images_dir=None, convert_mermaid=True):
    """A4 양면인쇄용 HTML 생성
    
    Args:
        md_content: 마크다운 내용
        title: 문서 제목
        images_dir: Mermaid 이미지 저장 디렉토리
        convert_mermaid: Mermaid 다이어그램 이미지 변환 여부
    """
    
    body_content = parse_markdown_to_html(md_content, images_dir, convert_mermaid)
    
    html = f'''<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        /* 나눔고딕 폰트 */
        @import url('https://fonts.googleapis.com/css2?family=Nanum+Gothic:wght@400;700;800&display=swap');
        
        /* A4 페이지 설정 */
        @page {{
            size: A4 portrait;
            margin: 20mm 15mm 20mm 20mm; /* 상 우 하 좌 - 좌철 여백 */
        }}
        
        /* 짝수 페이지 (뒷면) - 양면 인쇄 시 */
        @page :left {{
            margin: 20mm 20mm 20mm 15mm; /* 좌철 기준 */
        }}
        
        @page :right {{
            margin: 20mm 15mm 20mm 20mm;
        }}
        
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        html {{
            font-size: 12pt;
        }}
        
        body {{
            font-family: 'Nanum Gothic', '나눔고딕', '맑은 고딕', sans-serif;
            font-size: 12pt;
            line-height: 1.7;
            color: #000;
            background: #fff;
            max-width: 210mm;
            margin: 0 auto;
            padding: 0;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
        }}
        
        /* 제목 스타일 */
        h1 {{
            font-size: 20pt;
            font-weight: 800;
            margin: 30px 0 20px 0;
            padding-bottom: 10px;
            border-bottom: 3px solid #333;
            page-break-after: avoid;
            page-break-inside: avoid;
        }}
        
        h2 {{
            font-size: 16pt;
            font-weight: 700;
            margin: 25px 0 15px 0;
            padding: 8px 0 8px 12px;
            border-left: 4px solid #333;
            background: #f0f0f0;
            page-break-after: avoid;
            page-break-inside: avoid;
        }}
        
        h3 {{
            font-size: 14pt;
            font-weight: 700;
            margin: 20px 0 12px 0;
            page-break-after: avoid;
            page-break-inside: avoid;
        }}
        
        h4 {{
            font-size: 13pt;
            font-weight: 700;
            margin: 15px 0 10px 0;
            page-break-after: avoid;
            page-break-inside: avoid;
        }}
        
        h5, h6 {{
            font-size: 12pt;
            font-weight: 700;
            margin: 12px 0 8px 0;
            page-break-after: avoid;
        }}
        
        /* 문단 */
        p {{
            margin: 8px 0;
            text-align: justify;
            orphans: 3;
            widows: 3;
        }}
        
        /* 표 스타일 - 쪼개지지 않도록 */
        .table-container {{
            page-break-inside: avoid;
            margin: 15px 0;
            overflow: hidden;
        }}
        
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 10pt;
            margin: 0;
            table-layout: fixed;
        }}
        
        th, td {{
            border: 1px solid #333;
            padding: 6px 8px;
            text-align: center;
            vertical-align: middle;
            word-wrap: break-word;
        }}
        
        th {{
            background: #e0e0e0;
            font-weight: 700;
        }}
        
        /* 코드 블록 */
        .code-block {{
            page-break-inside: avoid;
            margin: 15px 0;
        }}
        
        pre {{
            background: #f5f5f5;
            border: 1px solid #ccc;
            border-radius: 4px;
            padding: 12px;
            overflow-x: auto;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 9pt;
            line-height: 1.4;
            white-space: pre-wrap;
            word-wrap: break-word;
        }}
        
        code {{
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 9pt;
        }}
        
        .inline-code {{
            background: #e8e8e8;
            padding: 2px 5px;
            border-radius: 3px;
            font-size: 10pt;
        }}
        
        /* 인용문 */
        blockquote {{
            margin: 15px 0;
            padding: 12px 15px;
            border-left: 4px solid #666;
            background: #f8f8f8;
            font-style: italic;
            page-break-inside: avoid;
        }}
        
        blockquote p {{
            margin: 5px 0;
        }}
        
        /* 리스트 */
        ul, ol {{
            margin: 10px 0 10px 25px;
            padding: 0;
        }}
        
        li {{
            margin: 5px 0;
            line-height: 1.6;
        }}
        
        /* 수평선 */
        hr {{
            border: none;
            border-top: 1px solid #999;
            margin: 20px 0;
        }}
        
        .page-break-suggestion {{
            page-break-after: auto;
        }}
        
        /* 링크 - 인쇄 시 URL 표시 안 함 */
        a {{
            color: #333;
            text-decoration: underline;
        }}
        
        /* 이미지 - 비율 유지 */
        img {{
            max-width: 100%;
            height: auto;
            display: block;
            margin: 15px auto;
            page-break-inside: avoid;
        }}
        
        /* 다이어그램 플레이스홀더 */
        .mermaid-placeholder {{
            page-break-inside: avoid;
            margin: 15px 0;
        }}
        
        /* 강조 */
        strong {{
            font-weight: 700;
        }}
        
        em {{
            font-style: italic;
        }}
        
        /* 화면 표시용 */
        @media screen {{
            body {{
                padding: 30px;
                max-width: 210mm;
                margin: 0 auto;
                background: #f0f0f0;
                filter: none;
                -webkit-filter: none;
            }}
            
            .page-wrapper {{
                background: #fff;
                padding: 20mm 15mm;
                box-shadow: 0 2px 10px rgba(0,0,0,0.15);
            }}
        }}
        
        /* 인쇄 스타일 */
        @media print {{
            body {{
                padding: 0;
                margin: 0;
                background: #fff;
                -webkit-filter: grayscale(100%);
                filter: grayscale(100%);
            }}
            
            .page-wrapper {{
                padding: 0;
                box-shadow: none;
            }}
            
            /* 표와 차트가 페이지 사이에 쪼개지지 않도록 */
            table, .table-container, .code-block, .mermaid-placeholder, blockquote {{
                page-break-inside: avoid !important;
            }}
            
            /* 제목 뒤에 바로 페이지 나눔 방지 */
            h1, h2, h3, h4, h5, h6 {{
                page-break-after: avoid !important;
            }}
            
            /* 첫 번째 제목 앞에서 페이지 나눔 방지 */
            h1:first-of-type {{
                page-break-before: avoid !important;
            }}
            
            /* URL 숨기기 */
            a[href]:after {{
                content: none !important;
            }}
            
            /* 페이지 번호 (브라우저 설정 필요) */
            @page {{
                @bottom-center {{
                    content: counter(page);
                }}
            }}
        }}
    </style>
</head>
<body>
    <div class="page-wrapper">
        {body_content}
    </div>
</body>
</html>'''
    
    return html

def convert_md_to_print_html(input_path, output_path=None, convert_mermaid=True):
    """마크다운 파일을 인쇄용 HTML로 변환
    
    Args:
        input_path: 입력 마크다운 파일 경로
        output_path: 출력 HTML 파일 경로 (None이면 자동 생성)
        convert_mermaid: Mermaid 다이어그램 이미지 변환 여부
    """
    
    input_path = Path(input_path)
    if output_path is None:
        output_path = input_path.with_suffix('.print.html')
    else:
        output_path = Path(output_path)
    
    # 이미지 저장 디렉토리 생성
    images_dir = output_path.parent / f'{output_path.stem}_images'
    if convert_mermaid:
        images_dir.mkdir(exist_ok=True)
        print(f"📁 이미지 디렉토리: {images_dir}")
    
    # 마크다운 읽기
    with open(input_path, 'r', encoding='utf-8') as f:
        md_content = f.read()
    
    # 제목 추출
    title_match = re.search(r'^#\s+(.+)$', md_content, re.MULTILINE)
    title = title_match.group(1) if title_match else input_path.stem
    
    # 사용 가능한 Mermaid 변환기 확인
    if convert_mermaid:
        print("\n🔍 Mermaid 변환기 확인 중...")
        if check_mermaid_cli():
            print("  ✅ mermaid-cli (mmdc) 사용 가능")
        elif check_playwright():
            print("  ✅ playwright 사용 가능")
        else:
            print("  ℹ️ kroki.io API 사용 (네트워크 필요)")
        print()
    
    # HTML 생성
    html_content = generate_print_html(md_content, title, images_dir if convert_mermaid else None, convert_mermaid)
    
    # 저장
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"\n✅ 변환 완료: {output_path}")
    print(f"   - 입력: {input_path}")
    print(f"   - 출력: {output_path}")
    if convert_mermaid:
        print(f"   - 이미지: {images_dir}")
    print(f"\n📌 인쇄 설정:")
    print(f"   - 용지: A4")
    print(f"   - 양면인쇄: 긴 가장자리로 넘김 (좌철)")
    print(f"   - 여백: 기본값 또는 사용자 정의")
    print(f"   - 그레이스케일: 자동 적용됨")
    
    return output_path


def print_help():
    """도움말 출력"""
    print("""
마크다운을 A4 양면인쇄용 HTML로 변환

사용법:
    python md_to_print_html.py <입력파일.md> [출력파일.html] [옵션]

옵션:
    --no-mermaid    Mermaid 다이어그램 변환 안 함 (플레이스홀더 사용)
    --help, -h      도움말 표시

Mermaid 다이어그램 변환기 (우선순위 순):
    1. mermaid-cli (mmdc) - npm install -g @mermaid-js/mermaid-cli
    2. playwright         - pip install playwright && playwright install chromium
    3. kroki.io API       - 네트워크 필요 (기본값)

예시:
    python md_to_print_html.py docs/PRESENTATION.md
    python md_to_print_html.py docs/PRESENTATION.md output.html
    python md_to_print_html.py docs/PRESENTATION.md --no-mermaid
""")


if __name__ == '__main__':
    import sys
    
    args = sys.argv[1:]
    
    # 도움말
    if '--help' in args or '-h' in args:
        print_help()
        sys.exit(0)
    
    # 옵션 파싱
    convert_mermaid = '--no-mermaid' not in args
    args = [a for a in args if not a.startswith('--')]
    
    if len(args) < 1:
        # 기본 파일 변환
        input_file = Path(__file__).parent.parent / 'docs' / 'PRESENTATION.md'
        output_file = Path(__file__).parent.parent / 'docs' / 'PRESENTATION_PRINT.html'
    else:
        input_file = args[0]
        output_file = args[1] if len(args) > 1 else None
    
    convert_md_to_print_html(input_file, output_file, convert_mermaid)

