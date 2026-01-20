# -*- coding: utf-8 -*-
"""
데이터표 위치(고정 범위) 로더
"""

from __future__ import annotations

from pathlib import Path
import re
from typing import Any

from config.settings import BASE_DIR

TABLE_LOCATIONS_PATH = BASE_DIR / "data_table_locations.md"


def _parse_range(range_text: str) -> dict[str, Any] | None:
    if not range_text:
        return None
    cleaned = range_text.strip().replace(" ", "")
    match = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", cleaned)
    if not match:
        return None
    start_col, start_row, end_col, end_row = match.groups()
    return {
        "start_col": start_col,
        "start_row": int(start_row),
        "end_col": end_col,
        "end_row": int(end_row),
    }


def load_table_locations(path: Path | None = None) -> dict[str, dict[str, Any]]:
    """
    data_table_locations.md에서 표 위치 정보를 읽어 딕셔너리로 반환
    """
    target = path or TABLE_LOCATIONS_PATH
    if not target.exists():
        return {}

    sections: dict[str, dict[str, Any]] = {}
    current_section: str | None = None

    for raw_line in target.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if line.startswith("## "):
            current_section = line.replace("## ", "", 1).strip()
            sections[current_section] = {}
            continue
        if not current_section or not line.startswith("-"):
            continue
        # "- 키: 값" 파싱
        parts = line[1:].strip().split(":", 1)
        if len(parts) != 2:
            continue
        key = parts[0].strip()
        value = parts[1].strip()
        if key == "파일":
            sections[current_section]["file"] = value
        elif key == "시트":
            sections[current_section]["sheet"] = value
        elif key == "범위":
            sections[current_section]["range"] = value
            parsed = _parse_range(value)
            if parsed:
                sections[current_section]["range_dict"] = parsed
        elif key == "헤더 포함":
            sections[current_section]["header_included"] = value in {"예", "true", "True", "Y", "yes"}
        elif key == "템플릿":
            sections[current_section]["template"] = value

    return sections
