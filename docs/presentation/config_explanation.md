# ⚙️ Config 모듈 상세 설명

발표 시 "Config가 뭐야?", "왜 Config를 분리했어?" 질문에 대비합니다.

---

## 🎯 Config란 무엇인가?

### 정의
> **Config(Configuration, 설정)**은 애플리케이션의 **동작 방식을 결정하는 값들을 코드와 분리하여 관리**하는 모듈입니다.

### 일반적인 Config의 역할
```
┌─────────────────────────────────────────────────────┐
│                    Config 영역                       │
├─────────────────────────────────────────────────────┤
│  • 경로 설정 (파일 저장 위치, 템플릿 위치)              │
│  • 상수 값 (최대 파일 크기, 비밀 키)                   │
│  • 환경 설정 (개발/운영 모드, 디버그 설정)              │
│  • 비즈니스 규칙 (보고서 목록, 시트 매핑)               │
└─────────────────────────────────────────────────────┘
```

### Config를 분리하는 이유
| 장점 | 설명 |
|------|------|
| **유지보수성** | 설정만 변경해도 프로그램 동작 변경 가능 |
| **가독성** | 핵심 로직과 설정 분리로 코드 이해 쉬움 |
| **재사용성** | 여러 모듈에서 동일 설정 참조 가능 |
| **확장성** | 새 기능 추가 시 설정만 추가하면 됨 |
| **환경 분리** | 개발/테스트/운영 환경별 다른 설정 적용 가능 |

---

## 📁 이 프로젝트의 Config 구조

```
config/
├── __init__.py      # 모듈 초기화, 외부 export 정의
├── settings.py      # 시스템 기본 설정 (경로, 상수)
└── reports.py       # 보고서 정의 (50개 보고서 설정)
```

---

## 1. config/settings.py - 시스템 설정

### 역할
- 프로젝트 경로 설정
- Flask 기본 설정
- 업로드 제한 설정

### 코드 설명
```python
from pathlib import Path

# 프로젝트 루트 설정 (이 파일 위치 기준 상위 폴더)
BASE_DIR = Path(__file__).parent.parent
# 결과: /Users/topos/Desktop/capstone

TEMPLATES_DIR = BASE_DIR / 'templates'
# 결과: /Users/topos/Desktop/capstone/templates

UPLOAD_FOLDER = BASE_DIR / 'uploads'
# 결과: /Users/topos/Desktop/capstone/uploads

# 업로드 폴더 자동 생성
UPLOAD_FOLDER.mkdir(exist_ok=True)

# Flask 보안 설정
SECRET_KEY = 'capstone_secret_key_2025'

# 최대 업로드 파일 크기: 50MB
MAX_CONTENT_LENGTH = 50 * 1024 * 1024
```

### 질문 대비
> **Q: 왜 하드코딩 안 하고 Path로 계산해?**  
> A: 다른 컴퓨터에서 실행해도 자동으로 경로가 맞춰집니다. 하드코딩하면 환경마다 수정해야 해요.

> **Q: `Path(__file__).parent.parent`가 뭐야?**  
> A: `__file__`은 현재 파일 경로, `.parent`는 상위 폴더입니다.  
> `config/settings.py` → `.parent` → `config/` → `.parent` → 프로젝트 루트

---

## 2. config/reports.py - 보고서 정의 ⭐

### 역할
- **50개 보고서**의 메타데이터 관리
- 보고서 ID, 이름, 엑셀 시트, Generator, Template 매핑

### 구조 개요
```python
# 4개 카테고리로 분류
SUMMARY_REPORTS = [...]      # 요약 보고서 9개
SECTOR_REPORTS = [...]       # 부문별 보고서 10개
REGIONAL_REPORTS = [...]     # 시도별 보고서 18개
STATISTICS_REPORTS = [...]   # 통계표 13개

# 전체 순서
REPORT_ORDER = SUMMARY_REPORTS + SECTOR_REPORTS
```

### 보고서 정의 예시
```python
SECTOR_REPORTS = [
    {
        'id': 'manufacturing',                     # 고유 식별자
        'name': '광공업생산',                        # 화면 표시 이름
        'sheet': 'A 분석',                         # 엑셀 시트명
        'generator': 'mining_manufacturing_generator.py',  # 데이터 추출 모듈
        'template': 'mining_manufacturing_template.html',  # HTML 템플릿
        'icon': '🏭',                              # 아이콘
        'category': 'production'                   # 카테고리
    },
    {
        'id': 'service',
        'name': '서비스업생산',
        'sheet': 'B 분석',
        'generator': 'service_industry_generator.py',
        'template': 'service_industry_template.html',
        'icon': '🏢',
        'category': 'production'
    },
    # ... 총 10개
]
```

### 시도별 보고서 정의
```python
REGIONAL_REPORTS = [
    {'id': 'region_seoul', 'name': '서울', 'full_name': '서울특별시', 'index': 1, 'icon': '🏙️'},
    {'id': 'region_busan', 'name': '부산', 'full_name': '부산광역시', 'index': 2, 'icon': '🌊'},
    {'id': 'region_daegu', 'name': '대구', 'full_name': '대구광역시', 'index': 3, 'icon': '🏛️'},
    # ... 17개 시도 + 참고_GRDP
]
```

---

## 3. config/__init__.py - 모듈 인터페이스

### 역할
- 외부에서 import 할 수 있는 항목 정의
- 모듈 사용 방법 단순화

### 코드
```python
from .reports import (
    SUMMARY_REPORTS,
    SECTOR_REPORTS,
    STATISTICS_REPORTS,
    REGIONAL_REPORTS,
    REPORT_ORDER
)
from .settings import (
    BASE_DIR,
    TEMPLATES_DIR,
    UPLOAD_FOLDER,
    SECRET_KEY,
    MAX_CONTENT_LENGTH
)

__all__ = [
    'SUMMARY_REPORTS', 'SECTOR_REPORTS', 'STATISTICS_REPORTS',
    'REGIONAL_REPORTS', 'REPORT_ORDER',
    'BASE_DIR', 'TEMPLATES_DIR', 'UPLOAD_FOLDER',
    'SECRET_KEY', 'MAX_CONTENT_LENGTH'
]
```

### 사용 효과
```python
# __init__.py 덕분에 간단하게 import 가능
from config import REPORT_ORDER, BASE_DIR

# __init__.py 없으면 이렇게 해야 함
from config.reports import REPORT_ORDER
from config.settings import BASE_DIR
```

---

## 🔗 실제 사용 예시

### 1️⃣ app.py - Flask 앱 생성
```python
from config.settings import BASE_DIR, SECRET_KEY, MAX_CONTENT_LENGTH, UPLOAD_FOLDER

def create_app():
    app = Flask(__name__, 
        template_folder=str(BASE_DIR),      # Config에서 경로 가져옴
        static_folder=str(BASE_DIR)
    )
    app.secret_key = SECRET_KEY             # Config에서 비밀키 가져옴
    app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)
    app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH
    return app
```

### 2️⃣ routes/main.py - 대시보드 페이지
```python
from config.reports import REPORT_ORDER, REGIONAL_REPORTS

@main_bp.route('/')
def index():
    # Config에서 보고서 목록을 가져와서 템플릿에 전달
    return render_template('dashboard.html', 
        reports=REPORT_ORDER,           # 요약+부문별 보고서 목록
        regional_reports=REGIONAL_REPORTS  # 시도별 보고서 목록
    )
```

### 3️⃣ routes/api.py - 보고서 생성
```python
from config.reports import REPORT_ORDER, REGIONAL_REPORTS, SUMMARY_REPORTS

@api_bp.route('/generate-preview', methods=['POST'])
def generate_preview():
    report_id = request.json.get('report_id')
    
    # Config에서 보고서 찾기
    for report in REPORT_ORDER:
        if report['id'] == report_id:
            # report 설정에 따라 적절한 generator와 template 사용
            generator_name = report['generator']
            template_name = report['template']
            break
```

### 4️⃣ services/report_generator.py - 보고서 HTML 생성
```python
from config.settings import TEMPLATES_DIR

def generate_report_html(excel_path, report_config, year, quarter):
    # Config의 경로를 사용해서 템플릿 파일 찾기
    template_path = TEMPLATES_DIR / report_config['template']
    generator_path = TEMPLATES_DIR / report_config['generator']
    
    # 동적으로 generator 모듈 로드
    module = load_generator_module(report_config['generator'])
    # ...
```

---

## 📊 Config 데이터 흐름

```
┌──────────────────────────────────────────────────────────┐
│                    config/reports.py                      │
│  ┌─────────────────────────────────────────────────────┐ │
│  │ SECTOR_REPORTS = [                                  │ │
│  │   {'id': 'manufacturing', 'generator': '...', ...}  │ │
│  │   {'id': 'service', 'generator': '...', ...}        │ │
│  │ ]                                                   │ │
│  └─────────────────────────────────────────────────────┘ │
└──────────────────────┬───────────────────────────────────┘
                       │
        ┌──────────────┼──────────────┐
        ▼              ▼              ▼
┌───────────────┐ ┌───────────────┐ ┌───────────────┐
│  routes/      │ │  services/    │ │  대시보드     │
│  main.py      │ │  report_gen.. │ │  dashboard    │
├───────────────┤ ├───────────────┤ ├───────────────┤
│ 보고서 목록   │ │ generator/    │ │ 보고서 버튼   │
│ 전달          │ │ template 결정 │ │ 렌더링        │
└───────────────┘ └───────────────┘ └───────────────┘
```

---

## 💡 Config 분리의 장점 (발표 포인트)

### 1. 새 보고서 추가가 쉬움
```python
# reports.py에 한 줄만 추가하면 됨
SECTOR_REPORTS.append({
    'id': 'new_report',
    'name': '새로운 보고서',
    'generator': 'new_report_generator.py',
    'template': 'new_report_template.html',
    'icon': '📝',
    'category': 'other'
})
# → 다른 코드 수정 필요 없음!
```

### 2. 순서 변경이 쉬움
```python
# REPORT_ORDER 순서만 바꾸면 출력 순서 변경
REPORT_ORDER = SUMMARY_REPORTS + SECTOR_REPORTS
# ↓
REPORT_ORDER = SECTOR_REPORTS + SUMMARY_REPORTS  # 부문별 먼저
```

### 3. 환경별 설정 분리 가능
```python
# 개발 환경
DEBUG = True
MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100MB

# 운영 환경
DEBUG = False
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB
```

### 4. 코드 중복 제거
```python
# Config 없이 하드코딩하면...
# main.py
template_folder='/Users/topos/Desktop/capstone/templates'

# api.py
template_folder='/Users/topos/Desktop/capstone/templates'  # 중복!

# preview.py
template_folder='/Users/topos/Desktop/capstone/templates'  # 또 중복!

# Config 사용하면...
# 모든 파일에서
from config import TEMPLATES_DIR
# → 한 곳에서 관리, 수정 시 한 번만
```

---

## 🤔 예상 질문 & 답변

### Q1: 왜 설정을 코드에 직접 안 쓰고 분리했나요?
> **A:** 설정이 바뀔 때 모든 파일을 수정하면 실수 가능성이 높습니다. 한 곳에서 관리하면 유지보수가 쉽고 일관성이 유지됩니다.

### Q2: 보고서 50개를 어떻게 관리하나요?
> **A:** `config/reports.py`에 딕셔너리 리스트로 정의합니다. 각 보고서의 ID, 이름, 사용할 generator와 template을 명시하고, 필요한 곳에서 이 Config를 참조합니다.

### Q3: Config와 환경변수의 차이는?
> **A:** 환경변수는 OS 레벨 설정(비밀키, API 키 등), Config는 애플리케이션 레벨 설정(보고서 목록, 경로 등)입니다. 이 프로젝트는 보안 민감 정보가 없어 Config 파일로 충분합니다.

### Q4: 새 보고서 추가하려면?
> **A:** `config/reports.py`에 보고서 정의 추가 → `templates/`에 generator와 template 파일 생성 → 끝! 다른 코드 수정 불필요.

---

## 📋 Config 요약표

| 파일 | 역할 | 주요 내용 |
|------|------|----------|
| `settings.py` | 시스템 설정 | 경로, 파일 크기, 비밀키 |
| `reports.py` | 보고서 정의 | 50개 보고서 메타데이터 |
| `__init__.py` | 모듈 인터페이스 | 외부 import 단순화 |

---

## 🎯 발표 시 핵심 메시지

> "Config 분리로 **보고서 추가/수정이 설정 파일 수정만으로 가능**합니다.  
> 50개 보고서를 관리하면서도 코드 변경 없이 유연하게 확장할 수 있습니다."

---

*마지막 업데이트: 2025년 12월*

