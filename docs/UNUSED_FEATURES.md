# 프론트엔드에서 사용하지 않는 기능 정리

이 문서는 현재 프론트엔드(`dashboard.html`)에서 사용되지 않는 기능들을 정리한 목록입니다.

## 1. 업로드 모달 - Step 3 (분석표 변환)

### 위치
- `dashboard.html` 라인 1755-1759: Step 3 HTML 요소
- `dashboard.html` 라인 2969-2973: Step 3 숨김 처리 코드

### 상태
- **사용 안 함**: 분석표를 직접 업로드하는 방식으로 변경되어, 기초자료 → 분석표 변환 단계가 필요 없어짐
- Step 3은 항상 `display: none`으로 숨겨짐
- 코드에서 Step 3을 건너뛰는 로직 포함

### 관련 코드
```1755:1759:dashboard.html
<div class="upload-step" id="uploadStep3">
    <span class="step-icon">⏳</span>
    <span class="step-text">분석표 변환</span>
    <span class="step-status"></span>
</div>
```

```2969:2973:dashboard.html
// Step 3 숨기기 (분석표 직접 업로드 시 변환 불필요)
const step3 = document.getElementById('uploadStep3');
if (step3) {
    step3.style.display = 'none';
}
```

---

## 2. 분석표 생성 및 다운로드 관련 함수들

### 2.1 `generateAnalysisInBackground()` 함수

### 위치
- `dashboard.html` 라인 4140-4233

### 상태
- **사용 안 함**: 함수가 정의되어 있지만 호출되는 곳이 없음
- API 엔드포인트(`/api/generate-analysis-with-weights`)가 주석 처리되어 있어 동작 불가

### 관련 코드
```4140:4233:dashboard.html
async function generateAnalysisInBackground() {
    // ... (전체 함수 코드)
    const response = await fetch('/api/generate-analysis-with-weights', {
        // ...
    });
    // ...
}
```

### 2.2 `downloadAnalysisFile()` 함수

### 위치
- `dashboard.html` 라인 4236-4289

### 상태
- **사용 안 함**: 함수가 정의되어 있지만 호출되는 곳이 없음
- API 엔드포인트(`/api/download-analysis`)가 주석 처리되어 있어 동작 불가

### 관련 코드
```4236:4289:dashboard.html
async function downloadAnalysisFile() {
    // ...
    downloadLink.href = '/api/download-analysis';
    // ...
}
```

### 2.3 `updateAnalysisButton()` 함수

### 위치
- `dashboard.html` 라인 4085-4113

### 상태
- **사용 안 함**: 함수가 정의되어 있지만 실제로 호출되는 곳이 없음
- `downloadAnalysisBtn` 요소가 HTML에 존재하지 않음
- `generateAnalysisInBackground()` 내부에서만 호출되지만, 해당 함수 자체가 호출되지 않음

### 관련 코드
```4085:4113:dashboard.html
function updateAnalysisButton(ready, text, hide = false) {
    const btn = document.getElementById('downloadAnalysisBtn');
    // ...
}
```

---

## 3. 사용하지 않는 API 엔드포인트

### 3.1 `/api/generate-analysis-with-weights`

### 위치
- `routes/api.py` 라인 767-811 (주석 처리됨)

### 상태
- **사용 안 함**: API 엔드포인트가 주석 처리되어 비활성화됨
- 프론트엔드에서 호출하려는 코드는 남아있지만 동작하지 않음

### 관련 코드
```767:811:routes/api.py
# 레거시 엔드포인트 - data_converter 모듈이 제거되어 비활성화됨
# @api_bp.route('/generate-analysis-with-weights', methods=['POST'])
# def generate_analysis_with_weights():
#     """분석표 생성 + 다운로드 (가중치 기본값 제거, 결측치는 N/A로 표시)"""
#     ...
```

### 3.2 `/api/download-analysis`

### 위치
- `routes/api.py` 라인 692-763 (주석 처리됨)

### 상태
- **사용 안 함**: API 엔드포인트가 주석 처리되어 비활성화됨
- 프론트엔드에서 호출하려는 코드는 남아있지만 동작하지 않음

### 관련 코드
```692:763:routes/api.py
# 레거시 엔드포인트 - data_converter 모듈이 제거되어 비활성화됨
# @api_bp.route('/download-analysis', methods=['GET'])
# def download_analysis():
#     """분석표 다운로드 (다운로드 시점에 생성 + 수식 계산)"""
#     ...
```

---

## 4. 존재하지 않는 HTML 요소 참조

### 4.1 `downloadAnalysisBtn`

### 위치
- `dashboard.html` 여러 곳에서 참조됨

### 상태
- **요소 없음**: `document.getElementById('downloadAnalysisBtn')`으로 참조하지만 실제 HTML 요소가 존재하지 않음
- 모든 참조가 `null`을 반환하여 실제로는 동작하지 않음

### 참조 위치
- `dashboard.html` 라인 4086: `updateAnalysisButton()` 함수 내
- `dashboard.html` 라인 4122: `updateUIForFileType()` 함수 내
- `dashboard.html` 라인 4131: 버튼 숨김 처리 코드

---

## 5. 사용하지 않는 상태 변수

### 5.1 `state.analysisFilename`

### 위치
- `dashboard.html` 상태 객체 내

### 상태
- **사용 안 함**: `generateAnalysisInBackground()` 함수에서만 사용되지만, 해당 함수가 호출되지 않음

### 5.2 `state.analysisReady`

### 위치
- `dashboard.html` 상태 객체 내

### 상태
- **사용 안 함**: `downloadAnalysisFile()` 함수에서만 사용되지만, 해당 함수가 호출되지 않음

---

## 정리 권장 사항

### 즉시 제거 가능한 항목
1. **업로드 모달 Step 3 HTML 요소** - 항상 숨겨지므로 제거 가능
2. **분석표 관련 함수들** (`generateAnalysisInBackground`, `downloadAnalysisFile`, `updateAnalysisButton`) - 호출되지 않음
3. **상태 변수** (`state.analysisFilename`, `state.analysisReady`) - 사용되지 않음
4. **HTML 요소 참조 코드** - `downloadAnalysisBtn` 참조 코드 (요소가 존재하지 않음)

### 유지 고려 사항
- API 엔드포인트는 이미 주석 처리되어 있으므로, 프론트엔드 코드만 정리하면 됨
- 향후 기초자료 → 분석표 변환 기능이 다시 필요할 수 있으므로, 백엔드 코드는 주석으로 남겨두는 것도 고려 가능

---

## 참고 사항

현재 시스템은 **분석표를 직접 업로드하는 방식**으로 변경되었습니다. 따라서:
- 기초자료 → 분석표 변환 기능은 더 이상 필요하지 않음
- 분석표 생성 및 다운로드 기능도 필요하지 않음
- 사용자는 분석표 파일을 직접 업로드하여 보도자료를 생성함
