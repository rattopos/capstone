#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
앱 보고서 생성 및 내보내기 End-to-End 테스트 (서버 없이 Flask test_client 사용)
"""
import io
import os
from pathlib import Path
from typing import List, Dict

from app import create_app
from config.settings import UPLOAD_FOLDER, TEMPLATES_DIR, EXPORT_FOLDER, BASE_DIR

EXCEL_FILE = Path(BASE_DIR) / "분석표_25년 3분기_캡스톤(업데이트).xlsx"


def _collect_pages_from_generated_outputs() -> List[Dict]:
    """templates 폴더에서 *_output.html 과 regional_output/*.html 읽어 pages 리스트 구성"""
    pages: List[Dict] = []
    # 일반 보고서들
    for fpath in sorted(TEMPLATES_DIR.glob("*_output.html")):
        try:
            html = fpath.read_text(encoding="utf-8")
            title = fpath.stem.replace("_output", "")
            pages.append({
                "title": title,
                "category": "report",
                "html": html
            })
        except Exception:
            continue
    # 시도별 보고서들
    regional_dir = TEMPLATES_DIR / "regional_output"
    if regional_dir.exists():
        for fpath in sorted(regional_dir.glob("*_output.html")):
            try:
                html = fpath.read_text(encoding="utf-8")
                title = fpath.stem.replace("_output", "")
                pages.append({
                    "title": f"시도별-{title}",
                    "category": "regional",
                    "html": html
                })
            except Exception:
                continue
    return pages


def main():
    print("="*60)
    print("앱 End-to-End 테스트: 업로드 → 생성 → 내보내기")
    print("="*60)

    if not EXCEL_FILE.exists():
        print(f"❌ 분석표 파일이 없습니다: {EXCEL_FILE}")
        return 1

    app = create_app()
    app.testing = True
    client = app.test_client()

    # 1) 업로드
    print("\n[1] /api/upload 업로드")
    with open(EXCEL_FILE, "rb") as f:
        data = {
            "file": (io.BytesIO(f.read()), EXCEL_FILE.name)
        }
        resp = client.post("/api/upload", data=data, content_type="multipart/form-data")
    if resp.status_code != 200:
        print(f"❌ 업로드 실패: status={resp.status_code}, body={resp.data[:200]}")
        return 1
    upload_json = resp.get_json()
    if not upload_json or not upload_json.get("success"):
        print(f"❌ 업로드 응답 실패: {upload_json}")
        return 1
    print(f"  ✅ 업로드 성공: 저장 파일명={upload_json.get('filename')}")

    # 2) 모든 보도자료 생성
    print("\n[2] /api/generate-all 전체 생성")
    resp = client.post("/api/generate-all", json={"year": 2025, "quarter": 3, "cleanup_after": False})
    if resp.status_code != 200:
        print(f"❌ generate-all 실패: status={resp.status_code}, body={resp.data[:200]}")
        return 1
    gen_json = resp.get_json()
    if not gen_json or not gen_json.get("success"):
        print(f"❌ 생성 응답 실패: {gen_json}")
        return 1
    generated = gen_json.get("generated", [])
    print(f"  ✅ 생성 성공: {len(generated)}개 파일")

    # 3) pages 수집
    pages = _collect_pages_from_generated_outputs()
    if not pages:
        print("❌ 생성된 페이지를 찾지 못했습니다.")
        return 1
    print(f"  ✅ pages 수집: {len(pages)}개")

    # 4) export-final (standalone=False)
    print("\n[3] /api/export-final HTML 합치기")
    resp = client.post("/api/export-final", json={
        "pages": pages,
        "year": 2025,
        "quarter": 3,
        "standalone": False
    })
    if resp.status_code != 200:
        print(f"❌ export-final 실패: status={resp.status_code}, body={resp.data[:200]}")
        return 1
    final_json = resp.get_json()
    if not final_json or not final_json.get("success"):
        print(f"❌ export-final 응답 실패: {final_json}")
        return 1
    final_name = final_json.get("filename")
    print(f"  ✅ export-final 성공: {final_name}")
    final_download = Path(UPLOAD_FOLDER) / final_name
    print(f"  - 저장 경로: {final_download}")
    if not final_download.exists():
        print("  ⚠️ 파일이 uploads에 없지만 응답 HTML은 포함되어 있습니다.")

    # 5) export-xlsx
    print("\n[4] /api/export-xlsx XLSX 내보내기")
    resp = client.post("/api/export-xlsx", json={
        "pages": pages,
        "year": 2025,
        "quarter": 3
    })
    if resp.status_code != 200:
        print(f"❌ export-xlsx 실패: status={resp.status_code}, body={resp.data[:200]}")
        return 1
    xlsx_json = resp.get_json()
    if not xlsx_json or not xlsx_json.get("success"):
        print(f"❌ export-xlsx 응답 실패: {xlsx_json}")
        return 1
    print(f"  ✅ export-xlsx 성공: {xlsx_json.get('filename')} (이미지 {xlsx_json.get('image_count')}개)")

    # 6) export-hwp-ready
    print("\n[5] /api/export-hwp-ready HWP 복붙용 내보내기")
    resp = client.post("/api/export-hwp-ready", json={
        "pages": pages,
        "year": 2025,
        "quarter": 3
    })
    if resp.status_code != 200:
        print(f"❌ export-hwp-ready 실패: status={resp.status_code}, body={resp.data[:200]}")
        return 1
    hwp_json = resp.get_json()
    if not hwp_json or not hwp_json.get("success"):
        print(f"❌ export-hwp-ready 응답 실패: {hwp_json}")
        return 1
    hwp_file = hwp_json.get("filename")
    hwp_path = Path(EXPORT_FOLDER) / hwp_file if hwp_file else None
    print(f"  ✅ export-hwp-ready 성공: {hwp_file}")
    if hwp_path and hwp_path.exists():
        print(f"  - 저장 경로: {hwp_path}")
    else:
        print(f"  ⚠️ 파일을 exports 폴더에서 찾지 못했습니다.")

    print("\n✅ End-to-End 테스트 완료")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())
