# Windows EXE 배포 시 발생 가능한 문제점

이 문서는 이 프로젝트를 Windows 배포용 EXE 파일로 패키징할 때 발생할 수 있는 주요 문제점과 해결 방안을 정리한 것입니다.

## 주요 문제점

### 1. `__file__` 경로 문제 (가장 중요)

**문제:**
- `config/settings.py`에서 `BASE_DIR = Path(__file__).parent.parent`로 경로를 설정하고 있음
- PyInstaller나 cx_Freeze로 패키징할 때 `__file__`이 임시 디렉토리(예: `C:\Users\...\AppData\Local\Temp\_MEIxxxxx\`)를 가리킴
- 리소스 파일(templates, config 등)을 찾지 못하는 문제 발생

**영향을 받는 코드:**
```python
# config/settings.py
BASE_DIR = Path(__file__).parent.parent
TEMPLATES_DIR = BASE_DIR / 'templates'
UPLOAD_FOLDER = BASE_DIR / 'uploads'
DEBUG_FOLDER = BASE_DIR / 'debug'
EXPORT_FOLDER = BASE_DIR / 'exports'
```

**해결 방법:**
- `sys._MEIPASS` 또는 PyInstaller의 `sys.frozen` 체크 필요
- 리소스 파일은 별도 디렉토리에 복사하거나 `--add-data` 옵션으로 포함

---

### 2. 리소스 파일 접근 불가

**문제:**
- `templates/` 디렉토리의 HTML, JSON, Python 파일들이 EXE에 번들링되어 접근 방식이 달라짐
- 동적 모듈 로드(`importlib.util`)가 실패할 수 있음

**영향을 받는 코드:**
- `services/report_generator.py`: `TEMPLATES_DIR / 'regional_generator.py'` 등의 경로
- `utils/excel_utils.py`: `load_generator_module()` 함수
- 모든 generator 모듈들의 동적 로드

**예시:**
```python
# services/report_generator.py
generator_path = TEMPLATES_DIR / 'regional_generator.py'
spec = importlib.util.spec_from_file_location('regional_generator', str(generator_path))
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
```

---

### 3. 업로드/출력 폴더 경로 문제

**문제:**
- `UPLOAD_FOLDER`, `DEBUG_FOLDER`, `EXPORT_FOLDER`가 BASE_DIR 기준으로 생성됨
- EXE 패키징 환경에서는 임시 디렉토리를 가리키거나, EXE 종료 시 삭제될 수 있음
- 사용자가 업로드한 파일이나 생성된 파일이 손실될 수 있음

**영향을 받는 코드:**
```python
# config/settings.py
UPLOAD_FOLDER.mkdir(exist_ok=True)
DEBUG_FOLDER.mkdir(exist_ok=True)
EXPORT_FOLDER.mkdir(exist_ok=True)
```

**해결 방법:**
- 사용자 문서 폴더나 EXE 실행 파일 위치 기준으로 경로 설정 필요
- 예: `Path.home() / 'Documents' / 'capstone_data'` 또는 `Path(sys.executable).parent / 'data'`

---

### 4. 환경 변수 파일(.env) 경로

**문제:**
- `kosis_collector.py`에서 `load_dotenv()`를 사용하여 `.env` 파일을 로드
- 프로젝트 루트에서 `.env` 파일을 찾지만, EXE 환경에서는 찾지 못할 수 있음

**영향을 받는 코드:**
```python
# kosis_collector.py
if env_path:
    load_dotenv(env_path)
else:
    load_dotenv()  # 프로젝트 루트에서 찾음
```

**해결 방법:**
- `.env` 파일 경로를 명시적으로 지정하거나 EXE 위치 기준 경로 사용
- 또는 환경 변수로 직접 전달

---

### 5. Flask 정적/템플릿 폴더 경로

**문제:**
- `app.py`에서 Flask의 `template_folder`와 `static_folder`가 `BASE_DIR` 기준으로 설정됨
- EXE 환경에서 BASE_DIR이 임시 디렉토리를 가리키면 템플릿/정적 파일을 찾지 못함

**영향을 받는 코드:**
```python
# app.py
app = Flask(
    __name__, 
    template_folder=str(BASE_DIR),
    static_folder=str(BASE_DIR)
)
```

---

### 6. 한글 파일명 인코딩 문제

**문제:**
- Windows에서 한글 파일명 처리 시 인코딩 문제 발생 가능
- `routes/api.py`의 `safe_filename()` 함수와 파일 다운로드 로직에서 주의 필요
- 업로드된 파일명이나 생성된 파일명이 깨질 수 있음

**영향을 받는 코드:**
- `routes/api.py`: `safe_filename()`, `send_file_with_korean_filename()`
- 파일 업로드/다운로드 관련 모든 코드

---

### 7. 동적 모듈 임포트 실패

**문제:**
- `templates/` 디렉토리의 Python 파일들을 동적으로 로드하는 구조
- EXE 패키징 시 이 방식이 동작하지 않을 수 있음

**영향을 받는 코드:**
```python
# utils/excel_utils.py
def load_generator_module(generator_name):
    generator_path = TEMPLATES_DIR / generator_name
    spec = importlib.util.spec_from_file_location(...)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module
```

**해결 방법:**
- 패키징 시 모든 generator 모듈을 미리 포함
- 또는 정적 임포트 방식으로 변경

---

### 8. 경로 구분자 문제 (부분 해결됨)

**현재 상태:**
- `pathlib.Path`를 사용하고 있어서 대부분 해결됨
- 일부 하드코딩된 경로(슬래시 `/` 사용)가 있다면 문제 가능

---

### 9. 실행 파일 위치 기반 경로 결정 필요

**권장 해결책:**
```python
import sys
import os
from pathlib import Path

def get_base_dir():
    """EXE 패키징 환경과 개발 환경을 모두 지원하는 BASE_DIR 반환"""
    if getattr(sys, 'frozen', False):
        # EXE 실행 환경 (PyInstaller)
        if hasattr(sys, '_MEIPASS'):
            # 리소스는 sys._MEIPASS에 있고
            # 데이터 폴더는 EXE 실행 파일 위치에 생성
            return Path(sys.executable).parent
        else:
            # cx_Freeze 등
            return Path(sys.executable).parent
    else:
        # 개발 환경
        return Path(__file__).parent.parent

def get_resource_dir():
    """리소스 파일(templates, config 등)이 있는 디렉토리 반환"""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # PyInstaller 실행 환경
        return Path(sys._MEIPASS)
    else:
        # 개발 환경
        return Path(__file__).parent.parent
```

---

### 10. 리소스 파일 번들링

**필요한 작업:**
- PyInstaller 사용 시 `--add-data` 옵션으로 templates, config 등 디렉토리 포함 필요
- 리소스 접근 시 `sys._MEIPASS`를 경로에 포함

**예시 (PyInstaller spec 파일):**
```python
# app.spec
a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('config', 'config'),
        ('kosis_config.json', '.'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
```

---

## 권장 해결 방안 요약

### 1. `config/settings.py` 수정
- EXE 환경 감지 로직 추가
- `get_base_dir()`, `get_resource_dir()` 함수 구현
- 데이터 폴더는 EXE 위치 기준, 리소스 폴더는 `sys._MEIPASS` 기준

### 2. 리소스 파일 접근 방식 변경
- 모든 리소스 접근 시 `get_resource_dir()` 사용
- `TEMPLATES_DIR` 계산 방식 변경

### 3. 데이터 폴더 경로 변경
- `UPLOAD_FOLDER`, `DEBUG_FOLDER`, `EXPORT_FOLDER`를 EXE 실행 파일 위치 기준으로 변경
- 또는 사용자 문서 폴더 사용

### 4. 환경 변수 처리
- `.env` 파일 경로를 EXE 위치 기준으로 설정
- 또는 환경 변수 직접 전달 방식으로 변경

### 5. 동적 모듈 로드 처리
- 패키징 시 모든 generator 모듈을 미리 포함
- 또는 정적 임포트 방식으로 전환 검토

---

## 참고 사항

- **PyInstaller**: 가장 널리 사용되는 패키징 도구
  - `--onefile`: 단일 EXE 파일 생성 (느린 시작)
  - `--onedir`: 디렉토리 구조 유지 (빠른 시작)
  
- **cx_Freeze**: 대안 패키징 도구

- **테스트 환경**: 
  - 개발 환경과 EXE 환경을 모두 테스트 필요
  - Windows 가상 머신에서 테스트 권장

---

## 작업 체크리스트

리팩토링 시 다음 항목들을 확인하세요:

- [ ] `config/settings.py`에 EXE 환경 감지 로직 추가
- [ ] `get_base_dir()`, `get_resource_dir()` 함수 구현
- [ ] 모든 리소스 파일 경로를 `get_resource_dir()` 기반으로 변경
- [ ] 데이터 폴더 경로를 EXE 위치 또는 사용자 문서 폴더로 변경
- [ ] `.env` 파일 경로 처리 수정
- [ ] Flask `template_folder`, `static_folder` 경로 수정
- [ ] 동적 모듈 로드 테스트
- [ ] 한글 파일명 처리 테스트
- [ ] PyInstaller spec 파일 작성
- [ ] EXE 빌드 및 테스트
- [ ] 업로드/다운로드 기능 테스트
- [ ] 보도자료 생성 기능 테스트

