# Vercel 배포 문제 해결 가이드

## 🔍 배포 실패 진단 방법

### 1. Vercel 대시보드에서 로그 확인

1. Vercel 대시보드 → 프로젝트 선택
2. **Deployments** 탭 클릭
3. 실패한 배포 클릭
4. **Build Logs** 또는 **Function Logs** 확인

### 2. 일반적인 오류와 해결 방법

#### ❌ 오류: "ModuleNotFoundError"

**원인**: 필요한 패키지가 `requirements.txt`에 없거나 설치 실패

**해결책**:
```bash
# 로컬에서 의존성 확인
pip freeze > requirements-check.txt

# requirements.txt에 누락된 패키지 추가
```

#### ❌ 오류: "Cannot find module 'app'"

**원인**: `api/index.py`에서 `app` 객체를 찾을 수 없음

**해결책**: `api/index.py` 파일 확인
- `from app import app` 구문이 있는지 확인
- 프로젝트 루트가 Python 경로에 추가되었는지 확인

#### ❌ 오류: "Build timeout" 또는 "Function timeout"

**원인**: 빌드 시간 또는 함수 실행 시간 초과

**해결책**: `vercel.json`에서 타임아웃 증가
```json
{
  "functions": {
    "api/index.py": {
      "maxDuration": 60
    }
  }
}
```

#### ❌ 오류: "Memory limit exceeded"

**원인**: 메모리 사용량 초과

**해결책**: `vercel.json`에서 메모리 증가
```json
{
  "functions": {
    "api/index.py": {
      "memory": 3008
    }
  }
}
```

#### ❌ 오류: "xlwings" 또는 "playwright" 설치 실패

**원인**: 서버리스 환경에서 작동하지 않는 패키지

**해결책**: 
1. `requirements-vercel.txt` 사용 (이미 생성됨)
2. 또는 `requirements.txt`에서 해당 패키지 제거

#### ❌ 오류: "FileNotFoundError" 또는 경로 오류

**원인**: Vercel 환경에서 파일 경로 문제

**해결책**: 
- `config/settings.py`에서 Vercel 환경 감지 로직 확인
- `/tmp` 디렉토리 사용 확인

## 🛠️ 배포 전 체크리스트

배포 전 다음을 확인하세요:

- [ ] `api/index.py` 파일이 존재하는가?
- [ ] `vercel.json` 파일이 존재하는가?
- [ ] `requirements.txt`에 모든 필수 의존성이 포함되어 있는가?
- [ ] 로컬에서 `python app.py`가 정상 실행되는가?
- [ ] Git에 모든 변경사항이 커밋되어 있는가?

## 🔧 수동 배포 테스트

### 로컬에서 Vercel CLI로 테스트

```bash
# Vercel CLI 설치
npm i -g vercel

# 개발 환경 배포 (테스트)
vercel

# 프로덕션 배포
vercel --prod
```

### 배포 전 로컬 테스트

```bash
# 의존성 설치 테스트
pip install -r requirements.txt

# 앱 실행 테스트
python app.py

# api/index.py import 테스트
python -c "import sys; sys.path.insert(0, '.'); from api.index import app; print('OK')"
```

## 📝 배포 설정 파일 확인

### `vercel.json` 확인

```json
{
  "version": 2,
  "builds": [
    {
      "src": "api/index.py",
      "use": "@vercel/python"
    }
  ],
  "routes": [
    {
      "src": "/(.*)",
      "dest": "api/index.py"
    }
  ],
  "functions": {
    "api/index.py": {
      "maxDuration": 60,
      "memory": 3008
    }
  }
}
```

### `api/index.py` 확인

다음 내용이 포함되어 있어야 합니다:
- 프로젝트 루트를 Python 경로에 추가
- Vercel 환경 감지 및 `/tmp` 디렉토리 설정
- `from app import app` 구문

## 🚨 긴급 해결 방법

### 방법 1: 최소 의존성으로 배포

`requirements-vercel.txt`를 사용하여 배포:

1. Vercel 대시보드 → Settings → Environment Variables
2. `REQUIREMENTS_FILE` 환경 변수 추가: `requirements-vercel.txt`
3. 또는 `requirements.txt`를 `requirements-vercel.txt`로 교체

### 방법 2: 단계적 배포

1. 먼저 기본 Flask 앱만 배포
2. 의존성을 하나씩 추가하며 테스트
3. 문제가 되는 패키지 식별

### 방법 3: 다른 플랫폼 사용

Vercel의 제약사항으로 인해 배포가 어려운 경우:
- **Railway**: 파일 시스템 지원, 쉬운 배포
- **Render**: 무료 플랜, 파일 시스템 지원
- **Heroku**: 전통적인 PaaS

## 📞 추가 도움

배포가 계속 실패하는 경우:

1. **Vercel 로그 전체 복사**: Build Logs와 Function Logs 모두
2. **에러 메시지 확인**: 정확한 오류 메시지와 스택 트레이스
3. **환경 정보 확인**: Python 버전, Node.js 버전 등

## 🔄 배포 재시도

설정을 수정한 후:

1. Git에 커밋 및 푸시
2. Vercel이 자동으로 재배포 시도
3. 또는 수동으로 "Redeploy" 클릭

