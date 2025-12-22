// 전역 변수
let selectedExcelFile = null;
let currentOutputFilename = null;
let selectedTemplate = null;
let templatesList = [];
let selectedPdfExcelFile = null;
let currentPdfFilename = null;
let selectedDocxExcelFile = null;
let currentDocxFilename = null;

// 토스트 알림 시스템
function showToast(message, type = 'info', duration = 5000) {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.setAttribute('role', 'alert');
    toast.setAttribute('aria-live', 'assertive');
    
    const icons = {
        success: '✓',
        error: '✕',
        info: 'ℹ'
    };
    
    toast.innerHTML = `
        <span class="toast-icon">${icons[type] || icons.info}</span>
        <span class="toast-message">${message}</span>
        <button class="toast-close" aria-label="닫기" onclick="this.parentElement.remove()">×</button>
    `;
    
    container.appendChild(toast);
    
    // 자동 제거
    if (duration > 0) {
        setTimeout(() => {
            if (toast.parentElement) {
                toast.style.animation = 'slideInRight 0.3s ease-out reverse';
                setTimeout(() => toast.remove(), 300);
            }
        }, duration);
    }
}

// 진행률 추적 변수
let progressTracker = {
    startTime: null,
    currentStep: 0,
    stepStartTime: null,
    stepProgress: 0,
    stepTimes: {}, // 각 단계의 시작 시간 기록
    stepEndTimes: {}, // 각 단계의 종료 시간 기록
    stepPercentages: {}, // 각 단계의 진행률 기록
    averageTimes: {}, // 각 단계의 평균 소요 시간 (localStorage에서 로드)
    stepNames: {
        0: '파일 준비',
        1: '데이터 분석',
        2: '템플릿 채우기',
        3: '결과 생성'
    }
};

// localStorage에서 평균 시간 로드
function loadAverageTimes() {
    try {
        const saved = localStorage.getItem('stepAverageTimes');
        if (saved) {
            progressTracker.averageTimes = JSON.parse(saved);
        }
    } catch (e) {
        console.warn('평균 시간 로드 실패:', e);
    }
}

// 평균 시간 저장
function saveAverageTime(step, actualTime) {
    if (!progressTracker.averageTimes[step]) {
        progressTracker.averageTimes[step] = [];
    }
    progressTracker.averageTimes[step].push(actualTime);
    
    // 최근 5개 기록만 유지
    if (progressTracker.averageTimes[step].length > 5) {
        progressTracker.averageTimes[step].shift();
    }
    
    try {
        localStorage.setItem('stepAverageTimes', JSON.stringify(progressTracker.averageTimes));
    } catch (e) {
        console.warn('평균 시간 저장 실패:', e);
    }
}

// 평균 시간 계산
function getAverageTime(step) {
    const times = progressTracker.averageTimes[step];
    if (!times || times.length === 0) {
        return null;
    }
    const sum = times.reduce((a, b) => a + b, 0);
    return Math.ceil(sum / times.length);
}

// 초기화 시 평균 시간 로드
loadAverageTimes();

// 진행률 표시 (개선된 버전)
function showProgress(percentage, text = null, step = null, subStep = null) {
    const container = document.getElementById('progressContainer');
    const bar = document.getElementById('progressBar');
    const header = container?.querySelector('.progress-header');
    const textEl = document.getElementById('progressText');
    const percentageEl = document.getElementById('progressPercentage');
    const stepsEl = document.getElementById('progressSteps');
    
    if (container && bar) {
        container.classList.add('active');
        bar.style.setProperty('--progress-width', `${Math.min(100, Math.max(0, percentage))}%`);
        
        if (header) header.style.display = 'flex';
        
        // 상세 텍스트 구성
        let displayText = text || '처리 중...';
        if (subStep) {
            displayText += ` - ${subStep}`;
        }
        if (textEl) textEl.textContent = displayText;
        
        if (percentageEl) percentageEl.textContent = `${Math.round(percentage)}%`;
        if (stepsEl) stepsEl.style.display = 'flex';
        
        // 단계 업데이트
        if (step !== null) {
            // 단계 시작 시간 기록
            if (!progressTracker.stepTimes[step]) {
                progressTracker.stepTimes[step] = Date.now();
            }
            // 진행률 기록
            progressTracker.stepPercentages[step] = percentage;
            updateProgressSteps(step, subStep, percentage);
        }
    }
}

// 시간 포맷팅 (초 -> "N초" 또는 "N분 N초")
function formatSeconds(seconds) {
    if (seconds < 60) {
        return `${seconds}초`;
    } else {
        const minutes = Math.floor(seconds / 60);
        const secs = seconds % 60;
        return `${minutes}분 ${secs}초`;
    }
}

function updateProgressSteps(activeStep, subStep = null, currentPercentage = 0) {
    const steps = ['progressStep1', 'progressStep2', 'progressStep3', 'progressStep4'];
    const stepLabels = [
        '파일 업로드 및 검증',
        '데이터 분석',
        '템플릿 채우기',
        '결과 생성'
    ];
    
    // 이전 단계가 완료되었을 때 종료 시간 기록 및 평균 시간 업데이트
    if (progressTracker.currentStep < activeStep) {
        // 이전 단계의 종료 시간 기록
        for (let i = 0; i < activeStep; i++) {
            if (progressTracker.stepTimes[i] && !progressTracker.stepEndTimes[i]) {
                progressTracker.stepEndTimes[i] = Date.now();
                // 실제 소요 시간 계산 및 평균 시간 업데이트
                const actualTime = Math.ceil((progressTracker.stepEndTimes[i] - progressTracker.stepTimes[i]) / 1000);
                saveAverageTime(i, actualTime);
            }
        }
        progressTracker.currentStep = activeStep;
    }
    
    steps.forEach((stepId, index) => {
        const stepEl = document.getElementById(stepId);
        if (stepEl) {
            stepEl.classList.remove('active', 'completed');
            
            let stepText = stepLabels[index];
            let timeInfo = '';
            
            if (index < activeStep) {
                // 완료된 단계: 실제 소요 시간만 표시
                stepEl.classList.add('completed');
                stepText = '✓ ' + stepText;
                
                const stepStartTime = progressTracker.stepTimes[index];
                const stepEndTime = progressTracker.stepEndTimes[index];
                
                if (stepStartTime && stepEndTime) {
                    const actualTime = Math.ceil((stepEndTime - stepStartTime) / 1000);
                    timeInfo = ` (${actualTime}초)`;
                }
            } else if (index === activeStep) {
                // 진행 중인 단계: 경과 시간만 표시
                stepEl.classList.add('active');
                stepText = '⏳ ' + stepText;
                
                // 서브 단계 정보 추가
                if (subStep) {
                    stepText += ` - ${subStep}`;
                }
                
                const stepStartTime = progressTracker.stepTimes[index];
                
                if (stepStartTime) {
                    const elapsed = Math.ceil((Date.now() - stepStartTime) / 1000);
                    if (elapsed > 0) {
                        timeInfo = ` (경과: ${elapsed}초)`;
                    }
                }
            }
            
            stepEl.textContent = stepText + timeInfo;
        }
    });
}

function hideProgress() {
    const container = document.getElementById('progressContainer');
    if (container) {
        container.classList.remove('active');
        setTimeout(() => {
            const bar = document.getElementById('progressBar');
            if (bar) bar.style.setProperty('--progress-width', '0%');
            const header = container.querySelector('.progress-header');
            if (header) header.style.display = 'none';
            const stepsEl = document.getElementById('progressSteps');
            if (stepsEl) stepsEl.style.display = 'none';
        }, 300);
    }
    
    // 진행률 추적 초기화
    progressTracker.startTime = null;
    progressTracker.currentStep = 0;
    progressTracker.stepProgress = 0;
    progressTracker.stepTimes = {};
    progressTracker.stepEndTimes = {};
    progressTracker.stepPercentages = {};
}

// DOM 로드 완료 시 초기화
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

// 앱 초기화
function initializeApp() {
    checkDefaultFile();
    loadTemplates();
    setupFileUpload();
    setupTemplateSelect();
    setupProcessButton();
    setupTabNavigation();
    setupPdfFileUpload();
    setupPdfGenerateButton();
    setupDocxFileUpload();
    setupDocxGenerateButton();
    setupCompareButtons();
}

// 파일 업로드 설정
function setupFileUpload() {
    const uploadArea = document.getElementById('excelUploadArea');
    const fileInput = document.getElementById('excelFile');
    const fileInfo = document.getElementById('excelFileInfo');

    // 클릭 이벤트
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // 파일 선택 이벤트
    fileInput.addEventListener('change', async (e) => {
        await handleFileSelect(e.target.files[0]);
    });

    // 드래그 앤 드롭 이벤트
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', async (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) {
            await handleFileSelect(file);
        }
    });
}

// 파일 선택 처리
async function handleFileSelect(file) {
    if (!file) return;

    // 파일 크기 검증 (100MB = 100 * 1024 * 1024 bytes)
    const maxFileSize = 100 * 1024 * 1024;
    if (file.size > maxFileSize) {
        showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        return;
    }

    // 파일 형식 검증
    const allowedTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    const allowedExtensions = ['.xlsx', '.xls'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();

    if (!allowedExtensions.includes(fileExtension) && !allowedTypes.includes(file.type)) {
        showError('지원하지 않는 파일 형식입니다. .xlsx 또는 .xls 파일만 업로드 가능합니다.');
        return;
    }
    
    // 파일 크기 포맷팅
    const fileSizeMB = (file.size / (1024 * 1024)).toFixed(2);
    showToast(`파일 업로드: ${file.name} (${fileSizeMB} MB)`, 'info', 3000);

    selectedExcelFile = file;
    displayFileInfo(file);
    
    // 필요한 시트 확인
    if (selectedTemplate) {
        const validation = await validateRequiredSheets(file);
        if (!validation.valid) {
            showError(validation.error);
            selectedExcelFile = null;
            document.getElementById('excelFileInfo').style.display = 'none';
            return;
        }
    }
    
    updateProcessButton();
}

// 파일 정보 표시
function displayFileInfo(file) {
    const fileInfo = document.getElementById('excelFileInfo');
    const fileName = fileInfo.querySelector('.file-name');
    
    fileName.textContent = file.name;
    fileInfo.style.display = 'flex';
}

// 엑셀 파일 제거
function removeExcelFile() {
    selectedExcelFile = null;
    document.getElementById('excelFile').value = '';
    document.getElementById('excelFileInfo').style.display = 'none';
    
    // 시트 확인 정보 숨기기
    document.getElementById('requiredSheetsInfo').style.display = 'none';
    
    updateProcessButton();
}

// 기본 파일 존재 여부 확인
async function checkDefaultFile() {
    try {
        const response = await fetch('/api/check-default-file');
        const data = await response.json();
        
        if (!data.exists) {
            showError(data.message || `기본 엑셀 파일을 찾을 수 없습니다: ${data.filename || '기초자료 수집표_2025년 2분기_캡스톤.xlsx'}`);
        }
    } catch (error) {
        console.error('기본 파일 확인 오류:', error);
        // 오류가 발생해도 앱은 계속 실행되도록 함
    }
}

// 템플릿 목록 로드
async function loadTemplates() {
    const templateSelect = document.getElementById('templateSelect');
    templateSelect.disabled = true;
    templateSelect.innerHTML = '<option value="">템플릿 목록을 불러오는 중...</option>';
    
    try {
        const response = await fetch('/api/templates');
        const data = await response.json();
        
        if (response.ok && data.templates) {
            templatesList = data.templates;
            templateSelect.innerHTML = '<option value="">템플릿을 선택하세요</option>';
            
            data.templates.forEach(template => {
                const option = document.createElement('option');
                option.value = template.name;
                option.textContent = template.display_name || template.name;
                option.dataset.requiredSheets = JSON.stringify(template.required_sheets || []);
                templateSelect.appendChild(option);
            });
            
            templateSelect.disabled = false;
            if (data.templates.length > 0) {
                showToast(`${data.templates.length}개의 템플릿을 불러왔습니다.`, 'success', 3000);
            }
        } else {
            templateSelect.innerHTML = '<option value="">템플릿을 불러올 수 없습니다</option>';
            showError('템플릿 목록을 불러올 수 없습니다.');
        }
    } catch (error) {
        console.error('템플릿 목록 로드 오류:', error);
        templateSelect.innerHTML = '<option value="">템플릿을 불러올 수 없습니다</option>';
        showError('템플릿 목록을 불러오는 중 오류가 발생했습니다.');
    }
}

// 템플릿 선택 설정
function setupTemplateSelect() {
    const templateSelect = document.getElementById('templateSelect');
    if (templateSelect) {
        templateSelect.addEventListener('change', function() {
            selectedTemplate = this.value;
            updateRequiredSheetsInfo();
            updateProcessButton();
        });
    }
}

// 필요한 시트 정보 업데이트
function updateRequiredSheetsInfo() {
    const templateSelect = document.getElementById('templateSelect');
    const requiredSheetsInfo = document.getElementById('requiredSheetsInfo');
    const requiredSheetsList = document.getElementById('requiredSheetsList');
    
    if (!templateSelect.value) {
        requiredSheetsInfo.style.display = 'none';
        return;
    }
    
    const selectedOption = templateSelect.options[templateSelect.selectedIndex];
    const requiredSheets = JSON.parse(selectedOption.dataset.requiredSheets || '[]');
    
    if (requiredSheets.length > 0) {
        requiredSheetsList.textContent = requiredSheets.join(', ');
        requiredSheetsInfo.style.display = 'block';
    } else {
        requiredSheetsInfo.style.display = 'none';
    }
}

// 엑셀 파일 업로드 후 필요한 시트 확인
async function validateRequiredSheets(file) {
    if (!selectedTemplate) {
        return { valid: true };
    }
    
    try {
        const formData = new FormData();
        formData.append('excel_file', file);
        
        const response = await fetch('/api/validate', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok && data.valid && data.sheet_names) {
            const templateSelect = document.getElementById('templateSelect');
            const selectedOption = templateSelect.options[templateSelect.selectedIndex];
            const requiredSheets = JSON.parse(selectedOption.dataset.requiredSheets || '[]');
            
            // 필요한 시트가 모두 있는지 확인
            const missingSheets = requiredSheets.filter(sheet => {
                // 유연한 매칭: 정확한 매칭 또는 부분 매칭
                return !data.sheet_names.some(availableSheet => {
                    const normalizedRequired = sheet.toLowerCase().trim();
                    const normalizedAvailable = availableSheet.toLowerCase().trim();
                    return normalizedRequired === normalizedAvailable ||
                           normalizedRequired.includes(normalizedAvailable) ||
                           normalizedAvailable.includes(normalizedRequired);
                });
            });
            
            if (missingSheets.length > 0) {
                return {
                    valid: false,
                    error: `필요한 시트를 찾을 수 없습니다: ${missingSheets.join(', ')}`
                };
            }
            
            // 시트별 연도/분기 정보 저장
            if (data.sheets_info) {
                window.sheetsInfo = data.sheets_info;
                // 첫 번째 필요한 시트의 연도/분기 업데이트
                if (requiredSheets.length > 0) {
                    const firstRequiredSheet = requiredSheets[0];
                    const matchedSheet = data.sheet_names.find(sheet => {
                        const normalizedRequired = firstRequiredSheet.toLowerCase().trim();
                        const normalizedAvailable = sheet.toLowerCase().trim();
                        return normalizedRequired === normalizedAvailable ||
                               normalizedRequired.includes(normalizedAvailable) ||
                               normalizedAvailable.includes(normalizedRequired);
                    });
                    if (matchedSheet && window.sheetsInfo[matchedSheet]) {
                        updateYearQuarterOptions(matchedSheet);
                    }
                }
            }
            
            return { valid: true };
        } else {
            return {
                valid: false,
                error: data.error || '엑셀 파일 검증에 실패했습니다.'
            };
        }
    } catch (error) {
        console.error('시트 검증 오류:', error);
        return {
            valid: false,
            error: '시트 검증 중 오류가 발생했습니다.'
        };
    }
}


// 처리 버튼 설정
function setupProcessButton() {
    const processBtn = document.getElementById('processBtn');
    processBtn.addEventListener('click', handleProcess);
}

// 처리 버튼 상태 업데이트
function updateProcessButton() {
    const processBtn = document.getElementById('processBtn');
    const templateSelect = document.getElementById('templateSelect');
    
    // 템플릿만 선택되어 있으면 활성화 (엑셀 파일은 선택사항)
    if (templateSelect.value) {
        processBtn.disabled = false;
    } else {
        processBtn.disabled = true;
    }
}

// 보도자료 생성 처리
async function handleProcess() {
    // 템플릿, 연도 및 분기 가져오기
    const templateSelect = document.getElementById('templateSelect');
    const yearSelect = document.getElementById('yearSelect');
    const quarterSelect = document.getElementById('quarterSelect');
    
    const templateName = templateSelect.value;
    const year = yearSelect.value;
    const quarter = quarterSelect.value;
    
    if (!templateName) {
        showError('템플릿을 선택해주세요.');
        return;
    }
    
    // 엑셀 파일이 선택된 경우에만 시트 검증
    if (selectedExcelFile) {
        const validation = await validateRequiredSheets(selectedExcelFile);
        if (!validation.valid) {
            showError(validation.error);
            return;
        }
    }

    // UI 업데이트
    const processBtn = document.getElementById('processBtn');
    const btnText = processBtn.querySelector('.btn-text');
    const btnLoader = processBtn.querySelector('.btn-loader');
    
    processBtn.disabled = true;
    btnText.style.display = 'none';
    btnLoader.style.display = 'inline-block';
    
    hideError();
    hideResult();
    
    // 진행률 추적 초기화
    progressTracker.startTime = Date.now();
    progressTracker.currentStep = 0;
    progressTracker.stepProgress = 0;
    
    // 단계별 진행률 시뮬레이션
    const progressSteps = [
        { percentage: 5, text: '파일 준비 중...', step: 0, subStep: '파일 검증' },
        { percentage: 10, text: '파일 준비 중...', step: 0, subStep: '파일 로드' },
        { percentage: 15, text: '데이터 분석 중...', step: 1, subStep: '엑셀 파일 열기' },
        { percentage: 25, text: '데이터 분석 중...', step: 1, subStep: '시트 목록 확인' },
        { percentage: 35, text: '데이터 분석 중...', step: 1, subStep: '연도/분기 감지' },
        { percentage: 45, text: '데이터 분석 중...', step: 1, subStep: '필요한 시트 매핑' },
        { percentage: 55, text: '템플릿 채우는 중...', step: 2, subStep: '템플릿 로드' },
        { percentage: 65, text: '템플릿 채우는 중...', step: 2, subStep: '마커 추출' },
        { percentage: 75, text: '템플릿 채우는 중...', step: 2, subStep: '데이터 추출 및 치환' },
        { percentage: 85, text: '템플릿 채우는 중...', step: 2, subStep: '포맷팅 처리' },
        { percentage: 90, text: '결과 생성 중...', step: 3, subStep: '파일 저장' },
        { percentage: 95, text: '결과 생성 중...', step: 3, subStep: '최종 검증' },
        { percentage: 100, text: '완료!', step: 3, subStep: null }
    ];
    
    let currentStepIndex = 0;
    
    // 진행률 업데이트 함수
    const updateProgress = () => {
        if (currentStepIndex < progressSteps.length) {
            const step = progressSteps[currentStepIndex];
            showProgress(step.percentage, step.text, step.step, step.subStep);
            currentStepIndex++;
        }
    };
    
    // 초기 진행률 표시
    showProgress(5, '파일 준비 중...', 0, '시작');
    if (selectedExcelFile) {
        showToast('파일을 업로드하고 처리 중입니다...', 'info', 2000);
    } else {
        showToast('기본 엑셀 파일을 사용하여 처리 중입니다...', 'info', 2000);
    }

    try {
        // FormData 생성
        const formData = new FormData();
        if (selectedExcelFile) {
            formData.append('excel_file', selectedExcelFile);
        }
        formData.append('template_name', templateName);
        formData.append('year', year);
        formData.append('quarter', quarter);

        // 진행률 시뮬레이션 시작 (비동기로 진행)
        const progressInterval = setInterval(() => {
            if (currentStepIndex < progressSteps.length - 1) {
                updateProgress();
            } else {
                clearInterval(progressInterval);
            }
        }, 800); // 0.8초마다 업데이트
        
        // API 호출
        const response = await fetch('/api/process', {
            method: 'POST',
            body: formData
        });

        // 진행률 시뮬레이션 중지
        clearInterval(progressInterval);
        
        // 마지막 단계로 진행
        showProgress(85, '템플릿 채우는 중...', 2, '데이터 처리 완료');

        const data = await response.json();
        
        // 결측치가 있는 경우 처리
        if (data.missing_values && data.missing_values.length > 0) {
            await handleMissingValues(data.missing_values, formData);
        }
        
        showProgress(95, '결과 생성 중...', 3, '최종 처리');

        if (response.ok && data.success) {
            currentOutputFilename = data.output_filename;
            showProgress(100, '완료!', 3, null);
            setTimeout(() => {
                hideProgress();
                showResult(data.message);
                showSuccess('보도자료가 성공적으로 생성되었습니다!');
            }, 500);
        } else {
            hideProgress();
            // 413 에러 (파일 크기 초과) 처리
            if (response.status === 413) {
                showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
            } else {
                showError(data.error || '처리 중 오류가 발생했습니다.');
            }
        }
    } catch (error) {
        console.error('처리 오류:', error);
        hideProgress();
        // 네트워크 오류나 파일 크기 초과 등의 경우
        if (error.message && error.message.includes('413')) {
            showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        } else {
            showError('서버와 통신하는 중 오류가 발생했습니다. 네트워크 연결을 확인해주세요.');
        }
    } finally {
        // UI 복원
        processBtn.disabled = false;
        btnText.style.display = 'inline';
        btnLoader.style.display = 'none';
        updateProcessButton();
    }
}

// 결과 표시
function showResult(message) {
    const resultSection = document.getElementById('resultSection');
    const resultMessage = document.getElementById('resultMessage');
    
    resultMessage.textContent = message;
    resultSection.style.display = 'block';
    
    // 미리보기 및 다운로드 버튼 설정
    setupResultButtons();
}

// 결과 숨기기
function hideResult() {
    document.getElementById('resultSection').style.display = 'none';
    hidePreview();
}

// 결과 버튼 설정
function setupResultButtons() {
    const previewBtn = document.getElementById('previewBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const compareBtn = document.getElementById('compareBtn');
    const closePreviewBtn = document.getElementById('closePreviewBtn');
    
    previewBtn.onclick = () => {
        if (currentOutputFilename) {
            showPreview(currentOutputFilename);
        }
    };
    
    downloadBtn.onclick = () => {
        if (currentOutputFilename) {
            showToast('파일을 다운로드합니다...', 'info', 2000);
            window.location.href = `/api/download/${currentOutputFilename}`;
        }
    };
    
    // 정답 비교 버튼
    if (compareBtn) {
        compareBtn.onclick = async () => {
            if (currentOutputFilename && selectedTemplate) {
                await showCompare(selectedTemplate, currentOutputFilename);
            }
        };
    }
    
    // 미리보기 닫기 버튼
    if (closePreviewBtn) {
        closePreviewBtn.onclick = () => {
            hidePreview();
        };
    }
}

// 정답 비교 표시
async function showCompare(templateName, outputFilename) {
    try {
        showToast('정답 파일을 불러오는 중...', 'info', 2000);
        
        const formData = new FormData();
        formData.append('template_name', templateName);
        formData.append('output_filename', outputFilename);
        
        const response = await fetch('/api/compare-answer', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok && data.success) {
            const compareSection = document.getElementById('compareSection');
            const answerImage = document.getElementById('answerImage');
            const comparePreviewFrame = document.getElementById('comparePreviewFrame');
            
            // 정답 이미지 로드
            answerImage.src = `/api/answer-image/${data.answer_file}`;
            
            // 생성된 결과 미리보기
            comparePreviewFrame.src = `/api/preview/${outputFilename}`;
            
            // 비교 섹션 표시
            compareSection.style.display = 'block';
            
            // 스크롤 이동
            compareSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
            
            showToast('정답 파일과 비교할 준비가 되었습니다.', 'success', 3000);
        } else {
            showError(data.error || '정답 파일을 불러올 수 없습니다.');
        }
    } catch (error) {
        console.error('비교 오류:', error);
        showError('정답 파일을 불러오는 중 오류가 발생했습니다.');
    }
}

// 비교 섹션 닫기
function hideCompare() {
    const compareSection = document.getElementById('compareSection');
    if (compareSection) {
        compareSection.style.display = 'none';
    }
}

// 비교 닫기 버튼 설정
function setupCompareButtons() {
    const closeCompareBtn = document.getElementById('closeCompareBtn');
    if (closeCompareBtn) {
        closeCompareBtn.onclick = () => {
            hideCompare();
        };
    }
}

// 미리보기 표시
function showPreview(filename) {
    const previewSection = document.getElementById('previewSection');
    const previewFrame = document.getElementById('previewFrame');
    
    if (previewSection && previewFrame) {
        previewFrame.src = `/api/preview/${filename}`;
        previewSection.style.display = 'block';
        
        // 미리보기 영역으로 스크롤
        setTimeout(() => {
            previewSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 100);
        
        showToast('미리보기를 표시합니다...', 'info', 2000);
    }
}

// 미리보기 숨기기
function hidePreview() {
    const previewSection = document.getElementById('previewSection');
    const previewFrame = document.getElementById('previewFrame');
    
    if (previewSection) {
        previewSection.style.display = 'none';
        if (previewFrame) {
            previewFrame.src = '';
        }
    }
}

// 연도/분기 옵션 업데이트
function updateYearQuarterOptions(sheetName) {
    if (!window.sheetsInfo || !window.sheetsInfo[sheetName]) {
        return;
    }
    
    const sheetInfo = window.sheetsInfo[sheetName];
    const yearSelect = document.getElementById('yearSelect');
    const quarterSelect = document.getElementById('quarterSelect');
    
    // 연도 옵션 업데이트
    yearSelect.innerHTML = '';
    for (let year = sheetInfo.min_year; year <= sheetInfo.max_year; year++) {
        const option = document.createElement('option');
        option.value = year;
        option.textContent = year;
        if (year === sheetInfo.default_year) {
            option.selected = true;
        }
        yearSelect.appendChild(option);
    }
    
    // 분기 옵션 업데이트 (항상 1-4)
    quarterSelect.innerHTML = '';
    for (let quarter = 1; quarter <= 4; quarter++) {
        const option = document.createElement('option');
        option.value = quarter;
        option.textContent = quarter + '분기';
        if (quarter === sheetInfo.default_quarter) {
            option.selected = true;
        }
        quarterSelect.appendChild(option);
    }
}

// 템플릿 선택 변경 시 처리 버튼 상태 업데이트
document.addEventListener('DOMContentLoaded', function() {
    const templateSelect = document.getElementById('templateSelect');
    if (templateSelect) {
        templateSelect.addEventListener('change', function() {
            updateProcessButton();
            updateRequiredSheetsInfo();
            // 엑셀 파일이 이미 업로드되어 있으면 시트 검증
            if (selectedExcelFile) {
                validateRequiredSheets(selectedExcelFile).then(validation => {
                    if (!validation.valid) {
                        showError(validation.error);
                    }
                });
            }
        });
    }
});

// 에러 표시
function showError(message) {
    const errorSection = document.getElementById('errorSection');
    const errorMessage = document.getElementById('errorMessage');
    
    if (errorSection && errorMessage) {
        errorMessage.textContent = message;
        errorSection.style.display = 'block';
        
        // 스크롤하여 에러 메시지가 보이도록
        errorSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
    
    // 토스트 알림도 표시
    showToast(message, 'error', 7000);
}

// 에러 숨기기
function hideError() {
    const errorSection = document.getElementById('errorSection');
    if (errorSection) {
        errorSection.style.display = 'none';
    }
}

// 성공 메시지 표시
function showSuccess(message) {
    showToast(message, 'success', 4000);
}

// 탭 네비게이션 설정
function setupTabNavigation() {
    const htmlTabBtn = document.getElementById('htmlTabBtn');
    const pdfTabBtn = document.getElementById('pdfTabBtn');
    const docxTabBtn = document.getElementById('docxTabBtn');
    const htmlTab = document.getElementById('html-tab');
    const pdfTab = document.getElementById('pdf-tab');
    const docxTab = document.getElementById('docx-tab');
    
    htmlTabBtn.addEventListener('click', () => {
        htmlTabBtn.classList.add('active');
        pdfTabBtn.classList.remove('active');
        docxTabBtn.classList.remove('active');
        htmlTab.classList.add('active');
        pdfTab.classList.remove('active');
        docxTab.classList.remove('active');
    });
    
    pdfTabBtn.addEventListener('click', () => {
        pdfTabBtn.classList.add('active');
        htmlTabBtn.classList.remove('active');
        docxTabBtn.classList.remove('active');
        pdfTab.classList.add('active');
        htmlTab.classList.remove('active');
        docxTab.classList.remove('active');
    });
    
    docxTabBtn.addEventListener('click', () => {
        docxTabBtn.classList.add('active');
        htmlTabBtn.classList.remove('active');
        pdfTabBtn.classList.remove('active');
        docxTab.classList.add('active');
        htmlTab.classList.remove('active');
        pdfTab.classList.remove('active');
    });
}

// PDF 엑셀 파일 업로드 설정
function setupPdfFileUpload() {
    const uploadArea = document.getElementById('pdfExcelUploadArea');
    const fileInput = document.getElementById('pdfExcelFile');
    
    if (!uploadArea || !fileInput) return;
    
    // 클릭 이벤트
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });
    
    // 파일 선택 이벤트
    fileInput.addEventListener('change', async (e) => {
        await handlePdfFileSelect(e.target.files[0]);
    });
    
    // 드래그 앤 드롭 이벤트
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });
    
    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('dragover');
    });
    
    uploadArea.addEventListener('drop', async (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) {
            await handlePdfFileSelect(file);
        }
    });
}

// PDF 파일 선택 처리
async function handlePdfFileSelect(file) {
    if (!file) return;
    
    // 파일 크기 검증 (100MB = 100 * 1024 * 1024 bytes)
    const maxFileSize = 100 * 1024 * 1024;
    if (file.size > maxFileSize) {
        showPdfError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        return;
    }
    
    // 파일 형식 검증
    const allowedExtensions = ['.xlsx', '.xls'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
    
    if (!allowedExtensions.includes(fileExtension)) {
        showPdfError('지원하지 않는 파일 형식입니다. .xlsx 또는 .xls 파일만 업로드 가능합니다.');
        return;
    }
    
    // 파일 크기 포맷팅
    const fileSizeMB = (file.size / (1024 * 1024)).toFixed(2);
    showToast(`파일 업로드: ${file.name} (${fileSizeMB} MB)`, 'info', 3000);
    
    selectedPdfExcelFile = file;
    displayPdfFileInfo(file);
}

// PDF 파일 정보 표시
function displayPdfFileInfo(file) {
    const fileInfo = document.getElementById('pdfExcelFileInfo');
    const fileName = fileInfo.querySelector('.file-name');
    
    if (fileName) {
        fileName.textContent = file.name;
        fileInfo.style.display = 'flex';
    }
}

// PDF 엑셀 파일 제거
function removePdfExcelFile() {
    selectedPdfExcelFile = null;
    const fileInput = document.getElementById('pdfExcelFile');
    if (fileInput) fileInput.value = '';
    const fileInfo = document.getElementById('pdfExcelFileInfo');
    if (fileInfo) fileInfo.style.display = 'none';
}

// PDF 생성 버튼 설정
function setupPdfGenerateButton() {
    const generatePdfBtn = document.getElementById('generatePdfBtn');
    if (generatePdfBtn) {
        generatePdfBtn.addEventListener('click', handlePdfGenerate);
    }
}

// PDF 생성 처리
async function handlePdfGenerate() {
    const yearSelect = document.getElementById('pdfYearSelect');
    const quarterSelect = document.getElementById('pdfQuarterSelect');
    
    const year = yearSelect.value;
    const quarter = quarterSelect.value;
    
    if (!year || !quarter) {
        showPdfError('연도와 분기를 선택해주세요.');
        return;
    }
    
    // UI 업데이트
    const generatePdfBtn = document.getElementById('generatePdfBtn');
    const btnText = generatePdfBtn.querySelector('.btn-text');
    const btnLoader = generatePdfBtn.querySelector('.btn-loader');
    
    generatePdfBtn.disabled = true;
    btnText.style.display = 'none';
    btnLoader.style.display = 'inline-block';
    
    hidePdfError();
    hidePdfResult();
    
    // PDF 진행률 추적 초기화
    pdfProgressTracker.startTime = Date.now();
    pdfProgressTracker.currentTemplate = 0;
    
    // PDF 생성 단계별 진행률 시뮬레이션
    const pdfProgressSteps = [
        { percentage: 5, text: '파일 준비 중...', templateIndex: 0, subStep: '파일 검증' },
        { percentage: 10, text: '파일 준비 중...', templateIndex: 0, subStep: '파일 로드' },
        { percentage: 15, text: '템플릿 처리 중...', templateIndex: 0, subStep: '템플릿 목록 확인' },
        { percentage: 20, text: '템플릿 처리 중...', templateIndex: 1, subStep: '템플릿 1/10 처리' },
        { percentage: 30, text: '템플릿 처리 중...', templateIndex: 3, subStep: '템플릿 3/10 처리' },
        { percentage: 40, text: '템플릿 처리 중...', templateIndex: 5, subStep: '템플릿 5/10 처리' },
        { percentage: 50, text: '템플릿 처리 중...', templateIndex: 7, subStep: '템플릿 7/10 처리' },
        { percentage: 60, text: '템플릿 처리 중...', templateIndex: 9, subStep: '템플릿 9/10 처리' },
        { percentage: 70, text: '템플릿 처리 중...', templateIndex: 10, subStep: '템플릿 처리 완료' },
        { percentage: 80, text: 'PDF 생성 중...', templateIndex: 10, subStep: 'HTML 변환' },
        { percentage: 90, text: 'PDF 생성 중...', templateIndex: 10, subStep: 'PDF 렌더링' },
        { percentage: 95, text: 'PDF 생성 중...', templateIndex: 10, subStep: '최종 검증' },
        { percentage: 100, text: '완료!', templateIndex: 10, subStep: null }
    ];
    
    let currentPdfStepIndex = 0;
    
    // 진행률 업데이트 함수
    const updatePdfProgress = () => {
        if (currentPdfStepIndex < pdfProgressSteps.length) {
            const step = pdfProgressSteps[currentPdfStepIndex];
            showPdfProgress(step.percentage, step.text, step.templateIndex, step.subStep);
            currentPdfStepIndex++;
        }
    };
    
    // 초기 진행률 표시
    showPdfProgress(5, '파일 준비 중...', 0, '시작');
    if (selectedPdfExcelFile) {
        showToast('파일을 업로드하고 PDF를 생성 중입니다...', 'info', 2000);
    } else {
        showToast('기본 엑셀 파일을 사용하여 PDF를 생성 중입니다...', 'info', 2000);
    }
    
    try {
        // FormData 생성
        const formData = new FormData();
        if (selectedPdfExcelFile) {
            formData.append('excel_file', selectedPdfExcelFile);
        }
        formData.append('year', year);
        formData.append('quarter', quarter);
        
        // 진행률 시뮬레이션 시작 (비동기로 진행)
        const pdfProgressInterval = setInterval(() => {
            if (currentPdfStepIndex < pdfProgressSteps.length - 1) {
                updatePdfProgress();
            } else {
                clearInterval(pdfProgressInterval);
            }
        }, 1200); // 1.2초마다 업데이트
        
        // API 호출
        const response = await fetch('/api/generate-pdf', {
            method: 'POST',
            body: formData
        });
        
        // 진행률 시뮬레이션 중지
        clearInterval(pdfProgressInterval);
        
        // 마지막 단계로 진행
        showPdfProgress(85, 'PDF 생성 중...', 10, '데이터 처리 완료');
        
        const data = await response.json();
        showPdfProgress(95, 'PDF 생성 중...', 10, '최종 처리');
        
        if (response.ok && data.success) {
            currentPdfFilename = data.output_filename;
            showPdfProgress(100, '완료!', 10, null);
            setTimeout(() => {
                hidePdfProgress();
                showPdfResult(data.message);
                showSuccess('PDF가 성공적으로 생성되었습니다!');
            }, 500);
        } else {
            hidePdfProgress();
            // 413 에러 (파일 크기 초과) 처리
            if (response.status === 413) {
                showPdfError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
            } else {
                showPdfError(data.error || 'PDF 생성 중 오류가 발생했습니다.');
            }
        }
    } catch (error) {
        console.error('PDF 생성 오류:', error);
        hidePdfProgress();
        // 네트워크 오류나 파일 크기 초과 등의 경우
        if (error.message && error.message.includes('413')) {
            showPdfError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        } else {
            showPdfError('서버와 통신하는 중 오류가 발생했습니다. 네트워크 연결을 확인해주세요.');
        }
    } finally {
        // UI 복원
        generatePdfBtn.disabled = false;
        btnText.style.display = 'inline';
        btnLoader.style.display = 'none';
    }
}

// PDF 진행률 추적 변수
let pdfProgressTracker = {
    startTime: null,
    currentTemplate: 0,
    totalTemplates: 10,
    stepTimes: {}, // 각 단계의 시작 시간 기록
    stepEndTimes: {}, // 각 단계의 종료 시간 기록
    stepPercentages: {}, // 각 단계의 진행률 기록
    averageTimes: {} // 각 단계의 평균 소요 시간 (localStorage에서 로드)
};

// PDF 평균 시간 로드
function loadPdfAverageTimes() {
    try {
        const saved = localStorage.getItem('pdfStepAverageTimes');
        if (saved) {
            pdfProgressTracker.averageTimes = JSON.parse(saved);
        }
    } catch (e) {
        console.warn('PDF 평균 시간 로드 실패:', e);
    }
}

// PDF 평균 시간 저장
function savePdfAverageTime(step, actualTime) {
    if (!pdfProgressTracker.averageTimes[step]) {
        pdfProgressTracker.averageTimes[step] = [];
    }
    pdfProgressTracker.averageTimes[step].push(actualTime);
    
    // 최근 5개 기록만 유지
    if (pdfProgressTracker.averageTimes[step].length > 5) {
        pdfProgressTracker.averageTimes[step].shift();
    }
    
    try {
        localStorage.setItem('pdfStepAverageTimes', JSON.stringify(pdfProgressTracker.averageTimes));
    } catch (e) {
        console.warn('PDF 평균 시간 저장 실패:', e);
    }
}

// PDF 평균 시간 계산
function getPdfAverageTime(step) {
    const times = pdfProgressTracker.averageTimes[step];
    if (!times || times.length === 0) {
        return null;
    }
    const sum = times.reduce((a, b) => a + b, 0);
    return Math.ceil(sum / times.length);
}

// 초기화 시 PDF 평균 시간 로드
loadPdfAverageTimes();

// PDF 진행률 표시 (개선된 버전)
function showPdfProgress(percentage, text = null, templateIndex = null, subStep = null) {
    const container = document.getElementById('pdfProgressContainer');
    const bar = document.getElementById('pdfProgressBar');
    const header = container?.querySelector('.progress-header');
    const textEl = document.getElementById('pdfProgressText');
    const percentageEl = document.getElementById('pdfProgressPercentage');
    const stepsEl = document.getElementById('pdfProgressSteps');
    
    if (container && bar) {
        container.classList.add('active');
        bar.style.setProperty('--progress-width', `${Math.min(100, Math.max(0, percentage))}%`);
        
        if (header) header.style.display = 'flex';
        
        // 상세 텍스트 구성
        let displayText = text || '처리 중...';
        if (subStep) {
            displayText += ` - ${subStep}`;
        }
        if (textEl) textEl.textContent = displayText;
        
        if (percentageEl) percentageEl.textContent = `${Math.round(percentage)}%`;
        if (stepsEl) stepsEl.style.display = 'flex';
        
        // 진행률 기록
        pdfProgressTracker.stepPercentages[pdfProgressTracker.currentTemplate] = percentage;
        
        // 단계별 시간 정보 업데이트
        updatePdfProgressSteps(templateIndex, subStep, percentage);
    }
}

// PDF 진행 단계 업데이트
function updatePdfProgressSteps(templateIndex, subStep, percentage) {
    const steps = ['pdfProgressStep1', 'pdfProgressStep2', 'pdfProgressStep3'];
    const stepLabels = [
        '파일 업로드 및 검증',
        '템플릿 처리 중',
        'PDF 생성'
    ];
    
    // 현재 단계 판단
    let currentStep = 0;
    if (percentage >= 70) {
        currentStep = 2; // PDF 생성
    } else if (percentage >= 15) {
        currentStep = 1; // 템플릿 처리
    } else {
        currentStep = 0; // 파일 준비
    }
    
    // 단계 시작 시간 기록
    if (!pdfProgressTracker.stepTimes[currentStep]) {
        pdfProgressTracker.stepTimes[currentStep] = Date.now();
    }
    
    // 이전 단계가 완료되었을 때 종료 시간 기록 및 평균 시간 업데이트
    if (pdfProgressTracker.currentTemplate < currentStep) {
        // 이전 단계의 종료 시간 기록
        for (let i = 0; i < currentStep; i++) {
            if (pdfProgressTracker.stepTimes[i] && !pdfProgressTracker.stepEndTimes[i]) {
                pdfProgressTracker.stepEndTimes[i] = Date.now();
                // 실제 소요 시간 계산 및 평균 시간 업데이트
                const actualTime = Math.ceil((pdfProgressTracker.stepEndTimes[i] - pdfProgressTracker.stepTimes[i]) / 1000);
                savePdfAverageTime(i, actualTime);
            }
        }
        pdfProgressTracker.currentTemplate = currentStep;
    }
    
    steps.forEach((stepId, index) => {
        const stepEl = document.getElementById(stepId);
        if (stepEl) {
            stepEl.classList.remove('active', 'completed');
            
            let stepText = stepLabels[index];
            let timeInfo = '';
            
            if (index < currentStep) {
                // 완료된 단계: 실제 소요 시간만 표시
                stepEl.classList.add('completed');
                stepText = '✓ ' + stepText;
                
                const stepStartTime = pdfProgressTracker.stepTimes[index];
                const stepEndTime = pdfProgressTracker.stepEndTimes[index];
                
                if (stepStartTime && stepEndTime) {
                    const actualTime = Math.ceil((stepEndTime - stepStartTime) / 1000);
                    timeInfo = ` (${actualTime}초)`;
                }
            } else if (index === currentStep) {
                // 진행 중인 단계: 경과 시간만 표시
                stepEl.classList.add('active');
                stepText = '⏳ ' + stepText;
                
                if (index === 1 && templateIndex !== null) {
                    stepText += ` (${templateIndex}/10)`;
                    if (subStep) {
                        stepText += ` - ${subStep}`;
                    }
                } else if (subStep) {
                    stepText += ` - ${subStep}`;
                }
                
                const stepStartTime = pdfProgressTracker.stepTimes[index];
                
                if (stepStartTime) {
                    const elapsed = Math.ceil((Date.now() - stepStartTime) / 1000);
                    if (elapsed > 0) {
                        timeInfo = ` (경과: ${elapsed}초)`;
                    }
                }
            }
            
            stepEl.textContent = stepText + timeInfo;
        }
    });
}

function hidePdfProgress() {
    const container = document.getElementById('pdfProgressContainer');
    if (container) {
        container.classList.remove('active');
        setTimeout(() => {
            const bar = document.getElementById('pdfProgressBar');
            if (bar) bar.style.setProperty('--progress-width', '0%');
            const header = container.querySelector('.progress-header');
            if (header) header.style.display = 'none';
            const stepsEl = document.getElementById('pdfProgressSteps');
            if (stepsEl) stepsEl.style.display = 'none';
        }, 300);
    }
    
    // 진행률 추적 초기화
    pdfProgressTracker.startTime = null;
    pdfProgressTracker.currentTemplate = 0;
    pdfProgressTracker.stepTimes = {};
    pdfProgressTracker.stepEndTimes = {};
    pdfProgressTracker.stepPercentages = {};
}

// PDF 결과 표시
function showPdfResult(message) {
    const resultSection = document.getElementById('pdfResultSection');
    const resultMessage = document.getElementById('pdfResultMessage');
    
    if (resultMessage) {
        resultMessage.textContent = message;
    }
    if (resultSection) {
        resultSection.style.display = 'block';
        
        // 다운로드 버튼 설정
        const pdfDownloadBtn = document.getElementById('pdfDownloadBtn');
        if (pdfDownloadBtn) {
            pdfDownloadBtn.onclick = () => {
                if (currentPdfFilename) {
                    showToast('PDF를 다운로드합니다...', 'info', 2000);
                    window.location.href = `/api/download/${currentPdfFilename}`;
                }
            };
        }
    }
}

// PDF 결과 숨기기
function hidePdfResult() {
    const resultSection = document.getElementById('pdfResultSection');
    if (resultSection) {
        resultSection.style.display = 'none';
    }
}

// PDF 에러 표시
function showPdfError(message) {
    const errorSection = document.getElementById('pdfErrorSection');
    const errorMessage = document.getElementById('pdfErrorMessage');
    
    if (errorSection && errorMessage) {
        errorMessage.textContent = message;
        errorSection.style.display = 'block';
        
        // 스크롤하여 에러 메시지가 보이도록
        errorSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
    
    // 토스트 알림도 표시
    showToast(message, 'error', 7000);
}

// PDF 에러 숨기기
function hidePdfError() {
    const errorSection = document.getElementById('pdfErrorSection');
    if (errorSection) {
        errorSection.style.display = 'none';
    }
}

// DOCX 엑셀 파일 업로드 설정
function setupDocxFileUpload() {
    const uploadArea = document.getElementById('docxExcelUploadArea');
    const fileInput = document.getElementById('docxExcelFile');
    
    if (!uploadArea || !fileInput) return;
    
    // 클릭 이벤트
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });
    
    // 파일 선택 이벤트
    fileInput.addEventListener('change', async (e) => {
        await handleDocxFileSelect(e.target.files[0]);
    });
    
    // 드래그 앤 드롭 이벤트
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });
    
    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('dragover');
    });
    
    uploadArea.addEventListener('drop', async (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) {
            await handleDocxFileSelect(file);
        }
    });
}

// DOCX 파일 선택 처리
async function handleDocxFileSelect(file) {
    if (!file) return;

    // 파일 크기 검증
    const maxFileSize = 100 * 1024 * 1024;
    if (file.size > maxFileSize) {
        showDocxError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        return;
    }

    // 파일 형식 검증
    const allowedExtensions = ['.xlsx', '.xls'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();

    if (!allowedExtensions.includes(fileExtension)) {
        showDocxError('지원하지 않는 파일 형식입니다. .xlsx 또는 .xls 파일만 업로드 가능합니다.');
        return;
    }
    
    const fileSizeMB = (file.size / (1024 * 1024)).toFixed(2);
    showToast(`파일 업로드: ${file.name} (${fileSizeMB} MB)`, 'info', 3000);

    selectedDocxExcelFile = file;
    displayDocxFileInfo(file);
}

// DOCX 파일 정보 표시
function displayDocxFileInfo(file) {
    const fileInfo = document.getElementById('docxExcelFileInfo');
    const fileName = fileInfo?.querySelector('.file-name');
    
    if (fileInfo && fileName) {
        fileName.textContent = file.name;
        fileInfo.style.display = 'flex';
    }
}

// DOCX 엑셀 파일 제거
function removeDocxExcelFile() {
    selectedDocxExcelFile = null;
    const fileInput = document.getElementById('docxExcelFile');
    if (fileInput) fileInput.value = '';
    const fileInfo = document.getElementById('docxExcelFileInfo');
    if (fileInfo) fileInfo.style.display = 'none';
}

// DOCX 생성 버튼 설정
function setupDocxGenerateButton() {
    const generateDocxBtn = document.getElementById('generateDocxBtn');
    if (generateDocxBtn) {
        generateDocxBtn.addEventListener('click', handleDocxGenerate);
    }
}

// DOCX 생성 처리
async function handleDocxGenerate() {
    const yearSelect = document.getElementById('docxYearSelect');
    const quarterSelect = document.getElementById('docxQuarterSelect');
    
    const year = yearSelect.value;
    const quarter = quarterSelect.value;
    
    if (!year || !quarter) {
        showDocxError('연도와 분기를 선택해주세요.');
        return;
    }
    
    // UI 업데이트
    const generateDocxBtn = document.getElementById('generateDocxBtn');
    const btnText = generateDocxBtn.querySelector('.btn-text');
    const btnLoader = generateDocxBtn.querySelector('.btn-loader');
    
    generateDocxBtn.disabled = true;
    btnText.style.display = 'none';
    btnLoader.style.display = 'inline-block';
    
    hideDocxError();
    hideDocxResult();
    
    // DOCX 진행률 추적 초기화
    docxProgressTracker.startTime = Date.now();
    docxProgressTracker.currentTemplate = 0;
    
    // DOCX 생성 단계별 진행률 시뮬레이션
    const docxProgressSteps = [
        { percentage: 5, text: '파일 준비 중...', templateIndex: 0, subStep: '파일 검증' },
        { percentage: 10, text: '파일 준비 중...', templateIndex: 0, subStep: '파일 로드' },
        { percentage: 15, text: '템플릿 처리 중...', templateIndex: 0, subStep: '템플릿 목록 확인' },
        { percentage: 20, text: '템플릿 처리 중...', templateIndex: 1, subStep: '템플릿 1/10 처리' },
        { percentage: 30, text: '템플릿 처리 중...', templateIndex: 3, subStep: '템플릿 3/10 처리' },
        { percentage: 40, text: '템플릿 처리 중...', templateIndex: 5, subStep: '템플릿 5/10 처리' },
        { percentage: 50, text: '템플릿 처리 중...', templateIndex: 7, subStep: '템플릿 7/10 처리' },
        { percentage: 60, text: '템플릿 처리 중...', templateIndex: 9, subStep: '템플릿 9/10 처리' },
        { percentage: 70, text: '템플릿 처리 중...', templateIndex: 10, subStep: '템플릿 처리 완료' },
        { percentage: 80, text: 'DOCX 생성 중...', templateIndex: 10, subStep: '문서 구조 생성' },
        { percentage: 90, text: 'DOCX 생성 중...', templateIndex: 10, subStep: '콘텐츠 삽입' },
        { percentage: 95, text: 'DOCX 생성 중...', templateIndex: 10, subStep: '최종 검증' },
        { percentage: 100, text: '완료!', templateIndex: 10, subStep: null }
    ];
    
    let currentDocxStepIndex = 0;
    
    // 진행률 업데이트 함수
    const updateDocxProgress = () => {
        if (currentDocxStepIndex < docxProgressSteps.length) {
            const step = docxProgressSteps[currentDocxStepIndex];
            showDocxProgress(step.percentage, step.text, step.templateIndex, step.subStep);
            currentDocxStepIndex++;
        }
    };
    
    // 초기 진행률 표시
    showDocxProgress(5, '파일 준비 중...', 0, '시작');
    if (selectedDocxExcelFile) {
        showToast('파일을 업로드하고 DOCX를 생성 중입니다...', 'info', 2000);
    } else {
        showToast('기본 엑셀 파일을 사용하여 DOCX를 생성 중입니다...', 'info', 2000);
    }
    
    try {
        // FormData 생성
        const formData = new FormData();
        if (selectedDocxExcelFile) {
            formData.append('excel_file', selectedDocxExcelFile);
        }
        formData.append('year', year);
        formData.append('quarter', quarter);
        
        // 진행률 시뮬레이션 시작 (비동기로 진행)
        const docxProgressInterval = setInterval(() => {
            if (currentDocxStepIndex < docxProgressSteps.length - 1) {
                updateDocxProgress();
            } else {
                clearInterval(docxProgressInterval);
            }
        }, 1200); // 1.2초마다 업데이트
        
        // API 호출
        const response = await fetch('/api/generate-docx', {
            method: 'POST',
            body: formData
        });
        
        // 진행률 시뮬레이션 중지
        clearInterval(docxProgressInterval);
        
        // 마지막 단계로 진행
        showDocxProgress(85, 'DOCX 생성 중...', 10, '데이터 처리 완료');
        
        const data = await response.json();
        showDocxProgress(95, 'DOCX 생성 중...', 10, '최종 처리');
        
        if (response.ok && data.success) {
            currentDocxFilename = data.output_filename;
            showDocxProgress(100, '완료!', 10, null);
            setTimeout(() => {
                hideDocxProgress();
                showDocxResult(data.message);
                showSuccess('DOCX가 성공적으로 생성되었습니다!');
            }, 500);
        } else {
            hideDocxProgress();
            // 413 에러 (파일 크기 초과) 처리
            if (response.status === 413) {
                showDocxError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
            } else {
                showDocxError(data.error || 'DOCX 생성 중 오류가 발생했습니다.');
            }
        }
    } catch (error) {
        console.error('DOCX 생성 오류:', error);
        hideDocxProgress();
        // 네트워크 오류나 파일 크기 초과 등의 경우
        if (error.message && error.message.includes('413')) {
            showDocxError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        } else {
            showDocxError('서버와 통신하는 중 오류가 발생했습니다. 네트워크 연결을 확인해주세요.');
        }
    } finally {
        // UI 복원
        generateDocxBtn.disabled = false;
        btnText.style.display = 'inline';
        btnLoader.style.display = 'none';
    }
}

// DOCX 진행률 추적 변수
let docxProgressTracker = {
    startTime: null,
    currentTemplate: 0,
    totalTemplates: 10,
    stepTimes: {}, // 각 단계의 시작 시간 기록
    stepEndTimes: {}, // 각 단계의 종료 시간 기록
    stepPercentages: {}, // 각 단계의 진행률 기록
    averageTimes: {} // 각 단계의 평균 소요 시간 (localStorage에서 로드)
};

// DOCX 평균 시간 로드
function loadDocxAverageTimes() {
    try {
        const saved = localStorage.getItem('docxStepAverageTimes');
        if (saved) {
            docxProgressTracker.averageTimes = JSON.parse(saved);
        }
    } catch (e) {
        console.warn('DOCX 평균 시간 로드 실패:', e);
    }
}

// DOCX 평균 시간 저장
function saveDocxAverageTime(step, actualTime) {
    if (!docxProgressTracker.averageTimes[step]) {
        docxProgressTracker.averageTimes[step] = [];
    }
    docxProgressTracker.averageTimes[step].push(actualTime);
    
    // 최근 5개 기록만 유지
    if (docxProgressTracker.averageTimes[step].length > 5) {
        docxProgressTracker.averageTimes[step].shift();
    }
    
    try {
        localStorage.setItem('docxStepAverageTimes', JSON.stringify(docxProgressTracker.averageTimes));
    } catch (e) {
        console.warn('DOCX 평균 시간 저장 실패:', e);
    }
}

// DOCX 평균 시간 계산
function getDocxAverageTime(step) {
    const times = docxProgressTracker.averageTimes[step];
    if (!times || times.length === 0) {
        return null;
    }
    const sum = times.reduce((a, b) => a + b, 0);
    return Math.ceil(sum / times.length);
}

// 초기화 시 DOCX 평균 시간 로드
loadDocxAverageTimes();

// DOCX 진행률 표시
function showDocxProgress(percentage, text = null, templateIndex = null, subStep = null) {
    const container = document.getElementById('docxProgressContainer');
    const bar = document.getElementById('docxProgressBar');
    const header = container?.querySelector('.progress-header');
    const textEl = document.getElementById('docxProgressText');
    const percentageEl = document.getElementById('docxProgressPercentage');
    const stepsEl = document.getElementById('docxProgressSteps');
    
    if (container && bar) {
        container.classList.add('active');
        bar.style.setProperty('--progress-width', `${Math.min(100, Math.max(0, percentage))}%`);
        
        if (header) header.style.display = 'flex';
        
        // 상세 텍스트 구성
        let displayText = text || '처리 중...';
        if (subStep) {
            displayText += ` - ${subStep}`;
        }
        if (textEl) textEl.textContent = displayText;
        
        if (percentageEl) percentageEl.textContent = `${Math.round(percentage)}%`;
        if (stepsEl) stepsEl.style.display = 'flex';
        
        // 진행률 기록
        docxProgressTracker.stepPercentages[docxProgressTracker.currentTemplate] = percentage;
        
        // 단계별 시간 정보 업데이트
        updateDocxProgressSteps(templateIndex, subStep, percentage);
    }
}

// DOCX 진행 단계 업데이트
function updateDocxProgressSteps(templateIndex, subStep, percentage) {
    const steps = ['docxProgressStep1', 'docxProgressStep2', 'docxProgressStep3'];
    const stepLabels = [
        '파일 업로드 및 검증',
        '템플릿 처리 중',
        'DOCX 생성'
    ];
    
    // 현재 단계 판단
    let currentStep = 0;
    if (percentage >= 70) {
        currentStep = 2; // DOCX 생성
    } else if (percentage >= 15) {
        currentStep = 1; // 템플릿 처리
    } else {
        currentStep = 0; // 파일 준비
    }
    
    // 단계 시작 시간 기록
    if (!docxProgressTracker.stepTimes[currentStep]) {
        docxProgressTracker.stepTimes[currentStep] = Date.now();
    }
    
    // 이전 단계가 완료되었을 때 종료 시간 기록 및 평균 시간 업데이트
    if (docxProgressTracker.currentTemplate < currentStep) {
        // 이전 단계의 종료 시간 기록
        for (let i = 0; i < currentStep; i++) {
            if (docxProgressTracker.stepTimes[i] && !docxProgressTracker.stepEndTimes[i]) {
                docxProgressTracker.stepEndTimes[i] = Date.now();
                // 실제 소요 시간 계산 및 평균 시간 업데이트
                const actualTime = Math.ceil((docxProgressTracker.stepEndTimes[i] - docxProgressTracker.stepTimes[i]) / 1000);
                saveDocxAverageTime(i, actualTime);
            }
        }
        docxProgressTracker.currentTemplate = currentStep;
    }
    
    steps.forEach((stepId, index) => {
        const stepEl = document.getElementById(stepId);
        if (stepEl) {
            stepEl.classList.remove('active', 'completed');
            
            let stepText = stepLabels[index];
            let timeInfo = '';
            
            if (index < currentStep) {
                // 완료된 단계: 실제 소요 시간만 표시
                stepEl.classList.add('completed');
                stepText = '✓ ' + stepText;
                
                const stepStartTime = docxProgressTracker.stepTimes[index];
                const stepEndTime = docxProgressTracker.stepEndTimes[index];
                
                if (stepStartTime && stepEndTime) {
                    const actualTime = Math.ceil((stepEndTime - stepStartTime) / 1000);
                    timeInfo = ` (${actualTime}초)`;
                }
            } else if (index === currentStep) {
                // 진행 중인 단계: 경과 시간만 표시
                stepEl.classList.add('active');
                stepText = '⏳ ' + stepText;
                
                if (index === 1 && templateIndex !== null) {
                    stepText += ` (${templateIndex}/10)`;
                    if (subStep) {
                        stepText += ` - ${subStep}`;
                    }
                } else if (subStep) {
                    stepText += ` - ${subStep}`;
                }
                
                const stepStartTime = docxProgressTracker.stepTimes[index];
                
                if (stepStartTime) {
                    const elapsed = Math.ceil((Date.now() - stepStartTime) / 1000);
                    if (elapsed > 0) {
                        timeInfo = ` (경과: ${elapsed}초)`;
                    }
                }
            }
            
            stepEl.textContent = stepText + timeInfo;
        }
    });
}

function hideDocxProgress() {
    const container = document.getElementById('docxProgressContainer');
    if (container) {
        container.classList.remove('active');
        setTimeout(() => {
            const bar = document.getElementById('docxProgressBar');
            if (bar) bar.style.setProperty('--progress-width', '0%');
            const header = container.querySelector('.progress-header');
            if (header) header.style.display = 'none';
            const stepsEl = document.getElementById('docxProgressSteps');
            if (stepsEl) stepsEl.style.display = 'none';
        }, 300);
    }
    
    // 진행률 추적 초기화
    docxProgressTracker.startTime = null;
    docxProgressTracker.currentTemplate = 0;
    docxProgressTracker.stepTimes = {};
    docxProgressTracker.stepEndTimes = {};
    docxProgressTracker.stepPercentages = {};
}

// DOCX 결과 표시
function showDocxResult(message) {
    const resultSection = document.getElementById('docxResultSection');
    const resultMessage = document.getElementById('docxResultMessage');
    
    if (resultMessage) {
        resultMessage.textContent = message;
    }
    if (resultSection) {
        resultSection.style.display = 'block';
        
        // 다운로드 버튼 설정
        const docxDownloadBtn = document.getElementById('docxDownloadBtn');
        if (docxDownloadBtn) {
            docxDownloadBtn.onclick = () => {
                if (currentDocxFilename) {
                    showToast('DOCX를 다운로드합니다...', 'info', 2000);
                    window.location.href = `/api/download/${currentDocxFilename}`;
                }
            };
        }
    }
}

// DOCX 결과 숨기기
function hideDocxResult() {
    const resultSection = document.getElementById('docxResultSection');
    if (resultSection) {
        resultSection.style.display = 'none';
    }
}

// DOCX 에러 표시
function showDocxError(message) {
    const errorSection = document.getElementById('docxErrorSection');
    const errorMessage = document.getElementById('docxErrorMessage');
    
    if (errorSection && errorMessage) {
        errorMessage.textContent = message;
        errorSection.style.display = 'block';
        
        // 스크롤하여 에러 메시지가 보이도록
        errorSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
    
    // 토스트 알림도 표시
    showToast(message, 'error', 7000);
}

// DOCX 에러 숨기기
function hideDocxError() {
    const errorSection = document.getElementById('docxErrorSection');
    if (errorSection) {
        errorSection.style.display = 'none';
    }
}

// 결측치 모달 관련 변수
let missingValueQueue = [];
let currentMissingValue = null;
let missingValueResolve = null;

// 결측치 모달 표시
function showMissingValueModal(missingInfo) {
    const modal = document.getElementById('missingValueModal');
    const sheetNameEl = document.getElementById('missingSheetName');
    const regionNameEl = document.getElementById('missingRegionName');
    const categoryNameEl = document.getElementById('missingCategoryName');
    const periodEl = document.getElementById('missingPeriod');
    const inputEl = document.getElementById('missingValueInput');
    
    if (modal && missingInfo) {
        if (sheetNameEl) sheetNameEl.textContent = missingInfo.sheet || 'N/A';
        if (regionNameEl) regionNameEl.textContent = missingInfo.region || 'N/A';
        if (categoryNameEl) categoryNameEl.textContent = missingInfo.category || 'N/A';
        if (periodEl) periodEl.textContent = `${missingInfo.year}년 ${missingInfo.quarter}분기`;
        if (inputEl) {
            inputEl.value = '';
            inputEl.focus();
        }
        
        currentMissingValue = missingInfo;
        modal.style.display = 'flex';
        
        // Enter 키로 제출
        inputEl?.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                submitMissingValue();
            }
        });
    }
}

// 결측치 모달 닫기
function closeMissingValueModal() {
    const modal = document.getElementById('missingValueModal');
    if (modal) {
        modal.style.display = 'none';
        currentMissingValue = null;
    }
}

// 결측치 값 제출
function submitMissingValue() {
    const inputEl = document.getElementById('missingValueInput');
    const value = inputEl?.value;
    
    if (value === '' || value === null) {
        showToast('값을 입력해주세요.', 'error', 3000);
        return;
    }
    
    const numValue = parseFloat(value);
    if (isNaN(numValue)) {
        showToast('올바른 숫자를 입력해주세요.', 'error', 3000);
        return;
    }
    
    if (currentMissingValue && missingValueResolve) {
        missingValueResolve(numValue);
        closeMissingValueModal();
        missingValueResolve = null;
    }
}

// 결측치 건너뛰기
function skipMissingValue() {
    if (currentMissingValue && missingValueResolve) {
        missingValueResolve(null);
        closeMissingValueModal();
        missingValueResolve = null;
    }
}

// 결측치 입력 요청 (Promise 기반)
function requestMissingValue(missingInfo) {
    return new Promise((resolve) => {
        missingValueResolve = resolve;
        showMissingValueModal(missingInfo);
    });
}

// 결측치 처리
async function handleMissingValues(missingValues, formData) {
    const missingValueMap = {};
    
    for (const missing of missingValues) {
        const value = await requestMissingValue({
            sheet: missing.sheet,
            region: missing.region,
            category: missing.category,
            year: missing.year,
            quarter: missing.quarter
        });
        
        if (value !== null) {
            const key = `${missing.sheet}_${missing.region}_${missing.category}_${missing.year}_${missing.quarter}`;
            missingValueMap[key] = value;
        }
    }
    
    // 결측치 값을 formData에 추가
    if (Object.keys(missingValueMap).length > 0) {
        formData.append('missing_values', JSON.stringify(missingValueMap));
    }
    
    return missingValueMap;
}

