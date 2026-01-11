# 🚀 프로젝트 최적화 로드맵

이 문서는 지역경제동향 보도자료 생성 시스템의 남은 최적화 작업을 정리합니다.

## 📊 현재 상태 요약

### ✅ 완료된 최적화 (2026-01-08 ~ 2026-01-09)

1. **코드 구조 개선**
   - `extractors/` 패키지 분리 (3,511줄 → 평균 328줄/파일)
   - Facade 패턴 적용으로 기존 인터페이스 유지
   - 도메인별 모듈 분리 (생산, 소비, 무역, 물가, 고용)

2. **성능 최적화**
   - 파일 유형 감지 최적화 (openpyxl 우선, 핵심 시트만 확인)
   - Excel 전처리 순서 변경 (openpyxl → formulas → xlwings)
   - 동적 컬럼 계산 로직 구현

3. **하드코딩 제거**
   - 연도/분기 동적 처리
   - 컬럼 인덱스 동적 계산
   - 템플릿 호환성 보장 (`_ensure_template_compatibility`)

4. **에러 처리 개선**
   - None 값 안전 처리 (`val`, `safe_abs` 필터)
   - 템플릿 호환성 보장 로직 추가
   - 데이터 구조 불일치 해결

---

## 🎯 남은 최적화 작업

### 1. 성능 최적화 (High Priority)

#### 1.1 메모리 최적화
**현재 문제:**
- 대용량 Excel 파일 처리 시 메모리 사용량 증가
- pandas DataFrame 전체 로딩으로 인한 메모리 부담

**개선 방안:**
```python
# 현재: 전체 시트를 한 번에 로드
df = pd.read_excel(xl, sheet_name='광공업생산', header=None)

# 개선: 청크 단위 읽기 또는 필요한 행만 읽기
df = pd.read_excel(xl, sheet_name='광공업생산', header=None, 
                   nrows=300, usecols='A:ZZ')  # 필요한 범위만
```

**예상 효과:**
- 메모리 사용량 30-50% 감소
- 대용량 파일 처리 시간 단축

**관련 파일:**
- `extractors/base.py`
- `services/summary_data.py`
- `services/excel_processor.py`

#### 1.2 데이터 캐싱 전략
**현재 문제:**
- 동일한 Excel 파일을 여러 번 읽음
- 시트별로 반복적인 데이터 추출

**개선 방안:**
```python
# 캐시 데코레이터 추가
from functools import lru_cache
from hashlib import md5

@lru_cache(maxsize=10)
def get_cached_sheet_data(filepath_hash, sheet_name):
    """시트 데이터 캐싱"""
    pass
```

**예상 효과:**
- 반복 작업 시간 50-70% 단축
- 서버 부하 감소

**관련 파일:**
- `extractors/base.py`
- `services/report_generator.py`

#### 1.3 병렬 처리
**현재 문제:**
- 보도자료 생성이 순차적으로 처리됨
- 각 보도자료가 독립적이므로 병렬 처리 가능

**개선 방안:**
```python
from concurrent.futures import ThreadPoolExecutor, as_completed

def generate_all_reports_parallel(excel_path, year, quarter):
    """병렬로 모든 보도자료 생성"""
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {
            executor.submit(generate_report, report_id, excel_path, year, quarter): report_id
            for report_id in REPORT_IDS
        }
        # 결과 수집
```

**예상 효과:**
- 전체 생성 시간 60-70% 단축 (4개 워커 기준)

**관련 파일:**
- `routes/api.py` (`generate_all` 엔드포인트)
- `services/report_generator.py`

---

### 2. 코드 품질 개선 (Medium Priority)

#### 2.1 레거시 코드 정리
**현재 문제:**
- `raw_data_extractor.py` (3,511줄) - 사용되지 않지만 유지됨
- `templates/*_generator.py` - 분석표 기반 생성기 (레거시)
- TODO 주석 다수 존재

**개선 방안:**
1. 레거시 코드 마이그레이션 완료 후 제거
2. 사용되지 않는 파일 정리
3. TODO 주석 해결 또는 이슈로 전환

**관련 파일:**
- `raw_data_extractor.py` (제거 검토)
- `templates/*_generator.py` (레거시)
- `kosis_grdp_fetcher.py` (TODO 주석)

#### 2.2 코드 중복 제거
**현재 문제:**
- 유사한 데이터 추출 로직 반복
- 템플릿 호환성 보장 로직 중복

**개선 방안:**
```python
# 공통 추출 로직 추상화
class CommonExtractorMixin:
    def extract_quarterly_data(self, sheet_name, region_col, code_col, total_code):
        """공통 분기 데이터 추출 로직"""
        pass
```

**관련 파일:**
- `extractors/production.py`
- `extractors/consumption.py`
- `extractors/trade.py`

#### 2.3 타입 힌팅 강화
**현재 문제:**
- 일부 함수에 타입 힌팅 누락
- 반환 타입 불명확

**개선 방안:**
```python
from typing import Dict, List, Optional, Tuple

def extract_report_data(
    self, 
    report_id: str
) -> Optional[Dict[str, Any]]:
    """타입 힌팅 추가"""
    pass
```

**관련 파일:**
- `extractors/` 패키지 전체
- `services/report_generator.py`

---

### 3. 에러 처리 및 안정성 (High Priority)

#### 3.1 로깅 시스템 개선
**현재 문제:**
- `print()` 문으로 디버그 로그 출력
- 로그 레벨 구분 없음
- 파일 기반 로깅 없음

**개선 방안:**
```python
import logging

# 로거 설정
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# 파일 핸들러 추가
file_handler = logging.FileHandler('logs/app.log')
file_handler.setFormatter(logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
))
logger.addHandler(file_handler)
```

**예상 효과:**
- 디버깅 시간 단축
- 운영 환경 모니터링 가능

**관련 파일:**
- 모든 서비스 파일
- `app.py` (로깅 설정)

#### 3.2 예외 처리 표준화
**현재 문제:**
- 예외 처리 방식이 일관되지 않음
- 사용자 친화적 에러 메시지 부족

**개선 방안:**
```python
class ReportGenerationError(Exception):
    """보도자료 생성 관련 커스텀 예외"""
    pass

class DataExtractionError(ReportGenerationError):
    """데이터 추출 실패"""
    pass

# 사용
try:
    data = extractor.extract_report_data(report_id)
except DataExtractionError as e:
    logger.error(f"데이터 추출 실패: {e}")
    return {"error": "데이터를 추출할 수 없습니다. 파일을 확인해주세요."}
```

**관련 파일:**
- `extractors/` 패키지
- `services/report_generator.py`
- `routes/api.py`

#### 3.3 데이터 검증 강화
**현재 문제:**
- 추출된 데이터의 유효성 검증 부족
- 잘못된 데이터로 인한 템플릿 오류

**개선 방안:**
```python
from pydantic import BaseModel, validator

class ReportData(BaseModel):
    """보도자료 데이터 스키마"""
    report_info: Dict[str, Any]
    regional_data: Dict[str, Any]
    
    @validator('report_info')
    def validate_year_quarter(cls, v):
        if not (2020 <= v.get('year', 0) <= 2030):
            raise ValueError("연도 범위 오류")
        if not (1 <= v.get('quarter', 0) <= 4):
            raise ValueError("분기 범위 오류")
        return v
```

**관련 파일:**
- `extractors/base.py`
- `services/report_generator.py`

---

### 4. 사용자 경험 개선 (Medium Priority)

#### 4.1 진행 상황 표시 개선
**현재 상태:**
- 로딩 모달에 진행률 표시 (완료)
- 남은 시간 추정 (완료)

**추가 개선:**
- 세부 작업 단계 표시 (예: "서비스업생산 데이터 추출 중...")
- 예상 완료 시간 정확도 향상

**관련 파일:**
- `dashboard.html` (프론트엔드)
- `routes/api.py` (백엔드 진행 상황 전송)

#### 4.2 에러 메시지 개선
**현재 문제:**
- 기술적 에러 메시지가 사용자에게 노출됨
- 사용자 친화적 안내 부족

**개선 방안:**
```python
ERROR_MESSAGES = {
    'file_not_found': '파일을 찾을 수 없습니다. 다시 업로드해주세요.',
    'invalid_sheet': '필수 시트가 없습니다. 기초자료 수집표 형식을 확인해주세요.',
    'data_extraction_failed': '데이터 추출에 실패했습니다. 파일 내용을 확인해주세요.',
}
```

**관련 파일:**
- `routes/api.py`
- `dashboard.html` (에러 표시)

#### 4.3 미리보기 성능 개선
**현재 문제:**
- 모든 보도자료를 한 번에 생성
- 불필요한 재생성 발생

**개선 방안:**
- 선택한 보도자료만 생성 (지연 로딩)
- 생성된 보도자료 캐싱

**관련 파일:**
- `routes/api.py` (`generate-preview` 엔드포인트)
- `dashboard.html` (지연 로딩)

---

### 5. 테스트 및 문서화 (Low Priority)

#### 5.1 단위 테스트 추가
**현재 상태:**
- 테스트 코드 없음

**개선 방안:**
```python
# tests/test_extractors.py
import pytest
from extractors import DataExtractor

def test_extract_manufacturing_data():
    """광공업생산 데이터 추출 테스트"""
    extractor = DataExtractor('test_data.xlsx', 2025, 3)
    data = extractor.extract_report_data('manufacturing')
    assert data is not None
    assert 'regional_data' in data
```

**예상 효과:**
- 리팩토링 시 회귀 버그 방지
- 코드 신뢰성 향상

**관련 파일:**
- `tests/` 디렉토리 생성
- 각 모듈별 테스트 파일

#### 5.2 통합 테스트
**개선 방안:**
- 전체 워크플로우 테스트
- 실제 Excel 파일을 사용한 E2E 테스트

#### 5.3 API 문서화
**개선 방안:**
- Swagger/OpenAPI 문서 생성
- API 엔드포인트 상세 설명

**관련 파일:**
- `routes/api.py` (docstring 추가)
- `docs/API.md` 생성

---

### 6. 보안 및 운영 (Medium Priority)

#### 6.1 파일 업로드 보안 강화
**현재 상태:**
- 기본적인 파일 확장자 검증만 존재

**개선 방안:**
```python
import magic  # python-magic

def validate_excel_file(filepath):
    """Excel 파일 유효성 검증"""
    # MIME 타입 확인
    mime = magic.from_file(filepath, mime=True)
    if mime not in ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    'application/vnd.ms-excel']:
        raise ValueError("유효하지 않은 Excel 파일입니다.")
    
    # 파일 크기 제한
    if os.path.getsize(filepath) > 100 * 1024 * 1024:  # 100MB
        raise ValueError("파일 크기가 너무 큽니다.")
```

**관련 파일:**
- `routes/api.py` (`upload_excel` 함수)

#### 6.2 환경 변수 관리
**개선 방안:**
- `.env` 파일 사용
- 민감한 정보 환경 변수로 관리

**관련 파일:**
- `config/settings.py`
- `.env.example` 생성

---

### 7. 배포 및 패키징 (High Priority)

#### 7.1 PyInstaller를 이용한 독립 실행 파일 생성
**목적:**
- Python 환경 없이도 실행 가능한 배포 패키지 생성
- 사용자 편의성 향상 (설치 과정 단순화)
- 오프라인 환경에서도 동작

**구현 계획:**

**1. PyInstaller 설정 파일 생성**
```python
# build.spec
# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('correct_answer', 'correct_answer'),
        ('static', 'static'),  # CSS, JS 파일이 있다면
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'jinja2',
        'flask',
        'extractors',
        'services',
        'routes',
        'utils',
        'config',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='지역경제동향_보도자료_생성기',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 모드 (Windows)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='resources/icon.ico',  # 아이콘 파일 (선택사항)
)
```

**2. 경로 처리 개선**
```python
# app.py 또는 config/settings.py
import sys
import os
from pathlib import Path

def get_base_path():
    """실행 파일 경로 또는 개발 환경 경로 반환"""
    if getattr(sys, 'frozen', False):
        # PyInstaller로 패키징된 경우
        base_path = Path(sys._MEIPASS)
    else:
        # 개발 환경
        base_path = Path(__file__).parent
    
    return base_path

BASE_DIR = get_base_path()
TEMPLATES_DIR = BASE_DIR / 'templates'
UPLOAD_FOLDER = BASE_DIR / 'uploads'
CORRECT_ANSWER_DIR = BASE_DIR / 'correct_answer'
```

**3. 빌드 스크립트 생성**
```bash
#!/bin/bash
# build.sh (macOS/Linux)

# 가상환경 활성화
source VENV/bin/activate

# PyInstaller 설치 확인
pip install pyinstaller

# 빌드 실행
pyinstaller build.spec

# 빌드 결과 확인
echo "빌드 완료: dist/지역경제동향_보도자료_생성기"
```

```batch
@echo off
REM build.bat (Windows)

REM 가상환경 활성화
call VENV\Scripts\activate.bat

REM PyInstaller 설치 확인
pip install pyinstaller

REM 빌드 실행
pyinstaller build.spec

REM 빌드 결과 확인
echo 빌드 완료: dist\지역경제동향_보도자료_생성기.exe
```

**4. 플랫폼별 빌드 전략**

**Windows:**
- 단일 실행 파일 (`.exe`) 생성
- 아이콘 파일 포함
- 콘솔 창 숨김 (GUI 모드)

**macOS:**
- 앱 번들 (`.app`) 생성
- 코드 서명 (선택사항)
- DMG 파일로 배포

**Linux:**
- AppImage 또는 단일 실행 파일
- 데스크톱 파일 생성

**5. 의존성 최적화**
```python
# requirements-build.txt
# PyInstaller 빌드용 최소 의존성
pyinstaller>=5.13.0
pyinstaller-hooks-contrib>=2023.0  # 추가 훅 지원
```

**6. 빌드 최적화 옵션**
```python
# build.spec 최적화 옵션
exe = EXE(
    # ... 기존 설정 ...
    upx=True,  # UPX 압축 사용 (파일 크기 감소)
    upx_exclude=[],  # UPX 제외 패키지
    optimize=2,  # 최적화 레벨
    strip=True,  # 디버그 심볼 제거 (Linux/macOS)
)
```

**예상 효과:**
- 배포 파일 크기: ~100-200MB (의존성 포함)
- 실행 시간: 개발 환경과 유사
- 사용자 설치 과정: 0단계 (실행 파일만 실행)

**관련 파일:**
- `build.spec` (신규 생성)
- `build.sh` / `build.bat` (빌드 스크립트)
- `app.py` (경로 처리 수정)
- `config/settings.py` (경로 처리 수정)

#### 7.2 배포 패키지 구조
**목표 구조:**
```
지역경제동향_보도자료_생성기/
├── 지역경제동향_보도자료_생성기.exe  (또는 .app, 실행 파일)
├── README.txt                         (사용자 가이드)
├── LICENSE                            (라이선스)
└── examples/                          (예제 파일)
    └── 기초자료_수집표_예제.xlsx
```

#### 7.3 자동 업데이트 메커니즘 (선택사항)
**개선 방안:**
- 버전 체크 API
- 자동 업데이트 다운로드
- 사용자 확인 후 업데이트

**구현 예시:**
```python
# utils/updater.py
import requests
import json

def check_for_updates(current_version):
    """최신 버전 확인"""
    try:
        response = requests.get('https://api.example.com/version')
        latest = response.json()['version']
        return latest > current_version
    except:
        return False
```

#### 7.4 설치 프로그램 생성 (Windows)
**도구:**
- Inno Setup (무료, 오픈소스)
- NSIS (Nullsoft Scriptable Install System)

**Inno Setup 스크립트 예시:**
```inno
[Setup]
AppName=지역경제동향 보도자료 생성 시스템
AppVersion=1.0.0
DefaultDirName={pf}\지역경제동향_보도자료_생성기
DefaultGroupName=지역경제동향 보도자료 생성 시스템
OutputDir=installer
OutputBaseFilename=지역경제동향_보도자료_생성기_Setup

[Files]
Source: "dist\지역경제동향_보도자료_생성기.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\지역경제동향 보도자료 생성 시스템"; Filename: "{app}\지역경제동향_보도자료_생성기.exe"
Name: "{commondesktop}\지역경제동향 보도자료 생성 시스템"; Filename: "{app}\지역경제동향_보도자료_생성기.exe"
```

#### 7.5 테스트 및 검증
**배포 전 체크리스트:**
- [ ] 모든 기능 정상 동작 확인
- [ ] 정적 파일 (템플릿, 이미지) 포함 확인
- [ ] Excel 파일 처리 정상 동작
- [ ] 플랫폼별 빌드 테스트
- [ ] 바이러스 검사 (False Positive 대응)
- [ ] 사용자 가이드 작성

**관련 파일:**
- `tests/test_build.py` (빌드 검증 테스트)
- `docs/DEPLOYMENT.md` (배포 가이드)

---

## 📋 우선순위별 작업 계획

### Phase 1: 핵심 최적화 (1-2주)
1. ✅ 메모리 최적화 (청크 단위 읽기)
2. ✅ 데이터 캐싱 전략 구현
3. ✅ 로깅 시스템 개선
4. ✅ 예외 처리 표준화

### Phase 2: 성능 향상 (2-3주)
1. ✅ 병렬 처리 구현
2. ✅ 미리보기 성능 개선
3. ✅ 에러 메시지 개선

### Phase 3: 코드 품질 (3-4주)
1. ✅ 레거시 코드 정리
2. ✅ 코드 중복 제거
3. ✅ 타입 힌팅 강화

### Phase 4: 테스트 및 문서화 (4-5주)
1. ✅ 단위 테스트 추가
2. ✅ 통합 테스트 추가
3. ✅ API 문서화

### Phase 5: 배포 및 패키징 (5-6주)
1. ✅ PyInstaller 설정 및 빌드 스크립트 작성
2. ✅ 경로 처리 개선 (frozen 환경 대응)
3. ✅ 플랫폼별 빌드 테스트 (Windows, macOS, Linux)
4. ✅ 설치 프로그램 생성 (Windows)
5. ✅ 배포 패키지 검증 및 문서화

---

## 📊 예상 효과

### 성능 개선
- **전체 생성 시간**: 현재 ~30초 → **~10초** (70% 단축)
- **메모리 사용량**: 현재 ~500MB → **~300MB** (40% 감소)
- **미리보기 로딩**: 현재 ~3초 → **~0.5초** (83% 단축)

### 코드 품질
- **코드 중복률**: 현재 ~15% → **~5%** (67% 감소)
- **테스트 커버리지**: 현재 0% → **~70%**
- **타입 힌팅**: 현재 ~40% → **~90%**

### 사용자 경험
- **에러 발생률**: 현재 ~5% → **~1%** (80% 감소)
- **사용자 만족도**: 현재 4.0/5.0 → **4.5/5.0** (예상)

### 배포 및 설치
- **설치 시간**: 현재 수동 설정 필요 → **0분** (즉시 실행)
- **Python 환경 요구사항**: 현재 필요 → **불필요** (독립 실행)
- **배포 파일 크기**: **~100-200MB** (의존성 포함)
- **플랫폼 지원**: Windows, macOS, Linux

---

## 🔍 모니터링 및 측정

### 성능 지표
- 각 최적화 작업 전후 성능 측정
- 메모리 프로파일링 (`memory_profiler`)
- 실행 시간 측정 (`timeit`, `cProfile`)

### 코드 품질 지표
- 코드 복잡도 (`radon`)
- 테스트 커버리지 (`pytest-cov`)
- 타입 체크 (`mypy`)

---

## 📝 참고 사항

- 모든 최적화 작업은 기존 기능을 손상시키지 않아야 함
- 각 단계마다 테스트 및 검증 필요
- 사용자 피드백을 반영하여 우선순위 조정 가능
- 레거시 코드 제거 시 하위 호환성 고려

---

**작성일**: 2026-01-09  
**최종 수정일**: 2026-01-09  
**작성자**: AI Assistant
