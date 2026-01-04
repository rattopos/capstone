#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
노션에 프로젝트 문서 업로드 스크립트

사용 방법:
1. 노션에서 Integration 생성: https://www.notion.so/my-integrations
2. Integration Token 복사
3. 업로드할 노션 페이지에 Integration 추가 (페이지 → ... → Connections → Integration 선택)
4. 노션 페이지 URL에서 페이지 ID 추출 (32자리 hex 문자열)
5. 아래 설정에 토큰과 페이지 ID 입력
6. python upload_to_notion.py 실행
"""

import os
import json
from pathlib import Path
from typing import Dict, List, Optional
import re

try:
    from notion_client import Client
    NOTION_AVAILABLE = True
except ImportError:
    NOTION_AVAILABLE = False
    print("[경고] notion-client 라이브러리가 설치되지 않았습니다.")
    print("      pip install notion-client 를 실행하세요.")

# ============================================
# 설정: 여기에 노션 정보를 입력하세요
# ============================================

# 노션 Integration Token (https://www.notion.so/my-integrations 에서 생성)
NOTION_TOKEN = os.getenv("NOTION_TOKEN", "")

# 업로드할 노션 페이지 ID (페이지 URL에서 추출)
# 예: https://www.notion.so/My-Page-abc123def456... → abc123def456...
NOTION_PAGE_ID = os.getenv("NOTION_PAGE_ID", "")

# 문서가 있는 디렉토리
DOCS_DIR = Path(__file__).parent / "docs"

# 업로드할 문서 목록 (우선순위 순서)
DOCUMENTS = [
    {
        "file": "PROJECT_PROGRESS.md",
        "title": "📊 프로젝트 진행 현황",
        "description": "프로젝트 개요, 개발 타임라인, 시스템 아키텍처, 현재 상태"
    },
    {
        "file": "STUDY_GUIDE.md",
        "title": "📚 종합 학습 가이드",
        "description": "프로젝트 학습, 용어 정리, AI 개발 히스토리, 기술 문서"
    },
    {
        "file": "DEPLOYMENT_GUIDE.md",
        "title": "📦 배포 및 설치 가이드",
        "description": "빠른 시작, 상세 설치 가이드, 문제 해결, Windows 배포 이슈"
    },
    {
        "file": "PRESENTATION.md",
        "title": "🎤 발표 자료",
        "description": "프로젝트 발표를 위한 상세 자료"
    },
    {
        "file": "DEVELOPMENT_CHALLENGES.md",
        "title": "🔴 개발 지연 원인 분석",
        "description": "프로젝트 개발 과정에서 겪은 어려움과 지연 원인 분석"
    },
    {
        "file": "GLOSSARY.md",
        "title": "📖 용어 사전",
        "description": "프로젝트에서 사용되는 모든 용어 정리"
    },
    {
        "file": "DEBUG_LOG.md",
        "title": "🐛 디버그 로그",
        "description": "프로젝트의 모든 디버그 작업 추적 및 기록"
    },
    {
        "file": "EXCEL_FLEXIBILITY_PLAN.md",
        "title": "📋 엑셀 구조 유연성 고도화 전략",
        "description": "엑셀 구조 변경에 유연하게 대응하는 방안"
    },
    {
        "file": "REPORT_GENERATION_FIX.md",
        "title": "🔧 보도자료 생성 오류 분석 및 해결",
        "description": "광공업생산 보도자료 생성 오류 분석 및 해결 과정"
    },
    {
        "file": "genspark_prompt_rule_based.md",
        "title": "💬 규칙기반 구현 프롬프트",
        "description": "규칙기반 시스템 구현을 위한 프롬프트 예시"
    },
]


def markdown_to_notion_blocks(markdown_text: str) -> List[Dict]:
    """마크다운 텍스트를 노션 블록 형식으로 변환"""
    blocks = []
    lines = markdown_text.split('\n')
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # 빈 줄
        if not line:
            i += 1
            continue
        
        # 제목 처리
        if line.startswith('# '):
            blocks.append({
                "object": "block",
                "type": "heading_1",
                "heading_1": {
                    "rich_text": [{"type": "text", "text": {"content": line[2:]}}]
                }
            })
        elif line.startswith('## '):
            blocks.append({
                "object": "block",
                "type": "heading_2",
                "heading_2": {
                    "rich_text": [{"type": "text", "text": {"content": line[3:]}}]
                }
            })
        elif line.startswith('### '):
            blocks.append({
                "object": "block",
                "type": "heading_3",
                "heading_3": {
                    "rich_text": [{"type": "text", "text": {"content": line[4:]}}]
                }
            })
        elif line.startswith('#### '):
            blocks.append({
                "object": "block",
                "type": "heading_3",
                "heading_3": {
                    "rich_text": [{"type": "text", "text": {"content": line[5:]}}]
                }
            })
        # 코드 블록
        elif line.startswith('```'):
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            if code_lines:
                blocks.append({
                    "object": "block",
                    "type": "code",
                    "code": {
                        "rich_text": [{"type": "text", "text": {"content": "\n".join(code_lines)}}],
                        "language": "plain text"
                    }
                })
        # 리스트
        elif line.startswith('- ') or line.startswith('* '):
            blocks.append({
                "object": "block",
                "type": "bulleted_list_item",
                "bulleted_list_item": {
                    "rich_text": [{"type": "text", "text": {"content": line[2:]}}]
                }
            })
        elif re.match(r'^\d+\.\s', line):
            blocks.append({
                "object": "block",
                "type": "numbered_list_item",
                "numbered_list_item": {
                    "rich_text": [{"type": "text", "text": {"content": re.sub(r'^\d+\.\s', '', line)}}]
                }
            })
        # 테이블 (간단한 처리)
        elif line.startswith('|') and '|' in line[1:]:
            # 테이블은 별도 처리 필요 (여기서는 단순 텍스트로)
            blocks.append({
                "object": "block",
                "type": "paragraph",
                "paragraph": {
                    "rich_text": [{"type": "text", "text": {"content": line}}]
                }
            })
        # 일반 텍스트
        else:
            # 링크 처리
            rich_text = []
            parts = re.split(r'(\[.*?\]\(.*?\))', line)
            for part in parts:
                if re.match(r'\[.*?\]\(.*?\)', part):
                    match = re.match(r'\[(.*?)\]\((.*?)\)', part)
                    if match:
                        rich_text.append({
                            "type": "text",
                            "text": {"content": match.group(1)},
                            "annotations": {"link": {"url": match.group(2)}}
                        })
                elif part:
                    rich_text.append({"type": "text", "text": {"content": part}})
            
            if not rich_text:
                rich_text = [{"type": "text", "text": {"content": line}}]
            
            blocks.append({
                "object": "block",
                "type": "paragraph",
                "paragraph": {"rich_text": rich_text}
            })
        
        i += 1
    
    return blocks


def upload_document_to_notion(notion, parent_page_id: str, doc_info: Dict) -> Optional[str]:
    """단일 문서를 노션에 업로드"""
    file_path = DOCS_DIR / doc_info["file"]
    
    if not file_path.exists():
        print(f"❌ 파일을 찾을 수 없습니다: {doc_info['file']}")
        return None
    
    print(f"📄 업로드 중: {doc_info['title']}...")
    
    # 파일 읽기
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        print(f"❌ 파일 읽기 실패: {e}")
        return None
    
    # 노션 페이지 생성
    try:
        # 페이지 생성
        page = notion.pages.create(
            parent={"page_id": parent_page_id},
            properties={
                "title": {
                    "title": [
                        {"text": {"content": doc_info["title"]}}
                    ]
                }
            }
        )
        
        page_id = page["id"]
        
        # 설명 추가 (있는 경우)
        if doc_info.get("description"):
            notion.blocks.children.append(
                block_id=page_id,
                children=[{
                    "object": "block",
                    "type": "paragraph",
                    "paragraph": {
                        "rich_text": [{"type": "text", "text": {"content": doc_info["description"]}}]
                    }
                }]
            )
        
        # 구분선 추가
        notion.blocks.children.append(
            block_id=page_id,
            children=[{
                "object": "block",
                "type": "divider",
                "divider": {}
            }]
        )
        
        # 마크다운을 노션 블록으로 변환하여 추가
        blocks = markdown_to_notion_blocks(content)
        
        # 노션은 한 번에 최대 100개 블록만 추가 가능
        chunk_size = 100
        for i in range(0, len(blocks), chunk_size):
            chunk = blocks[i:i + chunk_size]
            notion.blocks.children.append(
                block_id=page_id,
                children=chunk
            )
        
        print(f"✅ 업로드 완료: {doc_info['title']}")
        return page_id
        
    except Exception as e:
        print(f"❌ 업로드 실패: {e}")
        return None


def main():
    """메인 함수"""
    if not NOTION_AVAILABLE:
        print("\n❌ notion-client 라이브러리가 필요합니다.")
        print("   다음 명령어로 설치하세요: pip install notion-client\n")
        return
    
    # 토큰 입력 받기 (환경 변수에 없으면)
    notion_token = NOTION_TOKEN
    if not notion_token:
        print("\n📝 노션 Integration Token이 필요합니다.")
        print("   https://www.notion.so/my-integrations 에서 생성하세요.\n")
        notion_token = input("노션 Integration Token을 입력하세요: ").strip()
        if not notion_token:
            print("\n❌ 토큰이 입력되지 않았습니다. 종료합니다.\n")
            return
    
    # 페이지 ID 입력 받기 (환경 변수에 없으면)
    notion_page_id = NOTION_PAGE_ID
    if not notion_page_id:
        print("\n📝 노션 페이지 ID가 필요합니다.")
        print("   페이지 URL에서 32자리 hex 문자열을 추출하세요.")
        print("   예: https://www.notion.so/My-Page-abc123... → abc123...\n")
        notion_page_id = input("노션 페이지 ID를 입력하세요: ").strip()
        if not notion_page_id:
            print("\n❌ 페이지 ID가 입력되지 않았습니다. 종료합니다.\n")
            return
    
    # 노션 클라이언트 초기화
    try:
        notion = Client(auth=notion_token)
    except Exception as e:
        print(f"\n❌ 노션 클라이언트 초기화 실패: {e}")
        print("   토큰을 확인하세요.\n")
        return
    
    # 부모 페이지 확인
    try:
        parent_page = notion.pages.retrieve(notion_page_id)
        page_title = "Unknown"
        if 'properties' in parent_page:
            title_prop = parent_page['properties'].get('title', {})
            if 'title' in title_prop and title_prop['title']:
                page_title = title_prop['title'][0].get('plain_text', 'Unknown')
        print(f"\n📌 부모 페이지: {page_title}")
    except Exception as e:
        print(f"\n❌ 부모 페이지 접근 실패: {e}")
        print("   페이지 ID와 Integration 권한을 확인하세요.")
        print("   Integration이 페이지에 연결되어 있는지 확인하세요.\n")
        return
    
    print(f"\n🚀 총 {len(DOCUMENTS)}개 문서 업로드 시작...\n")
    
    # 각 문서 업로드
    uploaded = []
    failed = []
    
    for doc_info in DOCUMENTS:
        page_id = upload_document_to_notion(notion, notion_page_id, doc_info)
        if page_id:
            uploaded.append(doc_info["title"])
        else:
            failed.append(doc_info["title"])
        print()  # 빈 줄
    
    # 결과 요약
    print("=" * 60)
    print("📊 업로드 결과")
    print("=" * 60)
    print(f"✅ 성공: {len(uploaded)}개")
    for title in uploaded:
        print(f"   - {title}")
    
    if failed:
        print(f"\n❌ 실패: {len(failed)}개")
        for title in failed:
            print(f"   - {title}")
    
    print("\n✨ 완료!\n")


if __name__ == "__main__":
    main()

