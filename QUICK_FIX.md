# Vercel 배포 빠른 수정 가이드

## 🚨 배포가 안 될 때 즉시 확인할 사항

### 1단계: 에러 로그 확인

Vercel 대시보드에서:
1. 프로젝트 → **Deployments** 탭
2. 실패한 배포 클릭
3. **Build Logs** 확인
4. 에러 메시지 복사

### 2단계: 일반적인 해결 방법

#### 방법 A: requirements.txt 교체 (xlwings/playwright 문제 시)

```bash
# vercel-requirements.txt를 requirements.txt로 백업 후 교체
cp requirements.txt requirements-backup.txt
cp vercel-requirements.txt requirements.txt

# Git에 커밋
git add requirements.txt
git commit -m "Fix: Use Vercel-compatible requirements"
git push
```

#### 방법 B: vercel.json 수정

현재 `vercel.json`이 올바르게 설정되어 있는지 확인:
- `api/index.py`가 `src`에 있는지
- `routes`가 모든 경로를 `api/index.py`로 라우팅하는지

#### 방법 C: api/index.py 확인

`api/index.py` 파일이 다음을 포함하는지 확인:
```python
from app import app
```

### 3단계: 로컬 테스트

```bash
# 의존성 설치 테스트
pip install -r vercel-requirements.txt

# 앱 실행 테스트
python app.py

# api/index.py import 테스트
cd api
python -c "import sys; sys.path.insert(0, '..'); from index import app; print('✅ OK')"
```

### 4단계: Vercel CLI로 직접 배포

```bash
# Vercel CLI 설치
npm i -g vercel

# 로그인
vercel login

# 배포 (에러 메시지 확인)
vercel
```

## 📋 체크리스트

배포 전 확인:
- [ ] `api/index.py` 파일 존재
- [ ] `vercel.json` 파일 존재  
- [ ] `requirements.txt` 또는 `vercel-requirements.txt` 사용
- [ ] 로컬에서 `python app.py` 정상 실행
- [ ] Git에 커밋 및 푸시 완료

## 🔍 에러별 해결책

### "ModuleNotFoundError: No module named 'xlwings'"
→ `vercel-requirements.txt` 사용 (xlwings 제외)

### "ModuleNotFoundError: No module named 'playwright'"
→ `vercel-requirements.txt` 사용 (playwright 제외)

### "Cannot find module 'app'"
→ `api/index.py`에서 `from app import app` 확인

### "Build timeout"
→ `vercel.json`의 `maxDuration` 증가 (최대 60)

## 💡 빠른 해결

가장 빠른 해결 방법:

1. **requirements.txt를 vercel-requirements.txt로 교체**
2. **Git에 커밋 및 푸시**
3. **Vercel 자동 재배포 대기**

```bash
cp vercel-requirements.txt requirements.txt
git add requirements.txt
git commit -m "Fix: Use Vercel-compatible requirements"
git push
```

