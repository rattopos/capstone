# Vercel 배포 가이드

이 문서는 지역경제동향 보도자료 생성 시스템을 Vercel에 배포하는 방법을 설명합니다.

## 📋 사전 준비사항

1. **Vercel 계정**: [vercel.com](https://vercel.com)에서 계정 생성
2. **Vercel CLI 설치** (선택사항, 웹 대시보드에서도 배포 가능):
   ```bash
   npm i -g vercel
   ```
3. **Git 저장소**: 프로젝트가 Git 저장소에 연결되어 있어야 합니다

## ⚠️ 중요 제약사항

Vercel은 서버리스 플랫폼이므로 다음 제약사항이 있습니다:

### 1. 파일 시스템 제한
- **읽기 전용**: 프로젝트 파일은 읽기만 가능
- **임시 쓰기만 가능**: `/tmp` 디렉토리에만 파일 쓰기 가능
- **영구 저장 불가**: 업로드된 파일은 함수 실행 후 삭제됨

### 2. 세션 제한
- 서버리스 환경에서는 메모리 기반 세션이 제한적
- 업로드된 파일 경로를 세션에 저장하는 방식은 동작하지만, 파일 자체는 임시 저장소에만 존재

### 3. 실행 시간 제한
- 무료 플랜: 최대 10초 (설정에서 60초로 확장 가능)
- Pro 플랜: 최대 60초
- 대용량 Excel 파일 처리 시 시간 초과 가능성

### 4. 메모리 제한
- 무료 플랜: 최대 1GB
- Pro 플랜: 최대 3GB (설정에서 확장 가능)

## 🚀 배포 방법

### 방법 1: Vercel CLI 사용 (권장)

1. **프로젝트 디렉토리에서 로그인**:
   ```bash
   vercel login
   ```

2. **배포**:
   ```bash
   vercel
   ```
   
   첫 배포 시 다음 질문에 답변:
   - Set up and deploy? → **Y**
   - Which scope? → 본인의 계정 선택
   - Link to existing project? → **N** (새 프로젝트)
   - Project name? → 원하는 프로젝트 이름 입력
   - Directory? → **.** (현재 디렉토리)
   - Override settings? → **N**

3. **프로덕션 배포**:
   ```bash
   vercel --prod
   ```

### 방법 2: Vercel 웹 대시보드 사용

1. [Vercel 대시보드](https://vercel.com/dashboard) 접속
2. **"Add New..." → "Project"** 클릭
3. Git 저장소 연결 (GitHub, GitLab, Bitbucket)
4. 프로젝트 설정:
   - **Framework Preset**: Other
   - **Root Directory**: `.` (기본값)
   - **Build Command**: (비워둠)
   - **Output Directory**: (비워둠)
5. **Environment Variables** 설정 (필요시):
   - `SECRET_KEY`: Flask 시크릿 키 (선택사항)
6. **Deploy** 클릭

## 📁 배포 파일 구조

배포에 필요한 파일들:

```
capstone/
├── api/
│   └── index.py          # Vercel 서버리스 함수 진입점
├── vercel.json           # Vercel 설정 파일
├── .vercelignore         # 배포 제외 파일 목록
├── requirements.txt      # Python 의존성
└── ... (기타 프로젝트 파일)
```

## ⚙️ 설정 파일 설명

### `vercel.json`
- Vercel 배포 설정
- Python 런타임 버전, 함수 타임아웃, 메모리 제한 등 설정

### `api/index.py`
- Vercel 서버리스 함수 진입점
- Flask 애플리케이션을 WSGI 핸들러로 래핑
- Vercel 환경 감지 및 `/tmp` 디렉토리 설정

### `.vercelignore`
- 배포 시 제외할 파일/디렉토리 목록
- 가상환경, 업로드 파일, 캐시 등 제외

## 🔧 환경 변수 설정

Vercel 대시보드에서 환경 변수를 설정할 수 있습니다:

1. 프로젝트 → **Settings** → **Environment Variables**
2. 다음 변수 추가 (선택사항):
   - `SECRET_KEY`: Flask 시크릿 키 (기본값 사용 가능)
   - `VERCEL`: 자동 설정됨 (수동 설정 불필요)

## 🐛 문제 해결

### 1. 타임아웃 오류
**증상**: 함수 실행 시간 초과

**해결책**:
- `vercel.json`에서 `maxDuration` 값 증가 (최대 60초)
- Excel 파일 크기 줄이기
- 데이터 처리 로직 최적화

### 2. 메모리 부족 오류
**증상**: 메모리 초과 오류

**해결책**:
- `vercel.json`에서 `memory` 값 증가 (최대 3008MB)
- Pro 플랜으로 업그레이드

### 3. 파일 업로드 실패
**증상**: 업로드된 파일을 찾을 수 없음

**원인**: Vercel 서버리스 환경에서는 파일이 임시 저장소에만 존재하며, 함수 실행 후 삭제됨

**해결책**:
- 업로드 후 즉시 처리 (다운로드 또는 변환)
- 외부 스토리지 사용 (AWS S3, Cloudinary 등)

### 4. 모듈 import 오류
**증상**: `ModuleNotFoundError`

**해결책**:
- `requirements.txt`에 모든 의존성 포함 확인
- 로컬에서 `pip freeze > requirements.txt` 실행하여 의존성 확인

## 🔄 업데이트 배포

코드 변경 후 재배포:

```bash
# 개발 환경 배포
vercel

# 프로덕션 배포
vercel --prod
```

또는 Git에 푸시하면 자동 배포됩니다 (연결된 경우).

## 📊 모니터링

Vercel 대시보드에서 다음을 확인할 수 있습니다:
- 배포 상태 및 로그
- 함수 실행 시간 및 메모리 사용량
- 에러 로그
- 트래픽 통계

## 🔐 보안 고려사항

1. **시크릿 키**: 프로덕션에서는 환경 변수로 설정
2. **파일 업로드**: 파일 크기 및 형식 검증 필수
3. **CORS**: 필요시 CORS 설정 추가

## 💡 대안 플랫폼

Vercel의 제약사항으로 인해 다음 플랫폼도 고려할 수 있습니다:

1. **Railway**: 파일 시스템 쓰기 지원, 쉬운 배포
2. **Render**: 무료 플랜 제공, 파일 시스템 지원
3. **Heroku**: 전통적인 PaaS, 파일 시스템 지원
4. **AWS Lambda + API Gateway**: 더 많은 제어권, 복잡한 설정

## 📝 참고 자료

- [Vercel Python 문서](https://vercel.com/docs/concepts/functions/serverless-functions/runtimes/python)
- [Vercel CLI 문서](https://vercel.com/docs/cli)
- [Flask 배포 가이드](https://flask.palletsprojects.com/en/latest/deploying/)

