// 전역 변수
let selectedPdfFile = null;
let selectedExcelFile = null;
let currentOutputFilename = null;
let currentOutputFormat = 'pdf';
let sheetsInfo = {};

// 진행 상황 폴링 관련 변수
let progressPollingInterval = null;
let currentSessionId = null;

// 시간 추정 관련 변수
let stepStartTimes = {};
let stepDurations = {
    step1: [], // PDF to Word 변환 시간들
    step2: [], // 시트 감지 시간들
    step3: [], // 데이터 채우기 시간들
    step4: []  // 최종 변환 시간들
};
let pageOcrTimes = {}; // 페이지별 OCR 시간 추적
let currentStep = null;
let currentStepStartTime = null;

// DOM 로드 완료 시 초기화
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

// 앱 초기화
function initializeApp() {
    setupPdfUpload();
    setupExcelUpload();
    setupProcessButton();
    setupWorkflowSteps();
}

// 워크플로우 단계 설정
function setupWorkflowSteps() {
    // 더 이상 시트 선택이 없으므로 이 함수는 비워둠
}

// 워크플로우 단계 업데이트
function updateWorkflowStep(step) {
    // 모든 단계 비활성화
    document.querySelectorAll('.workflow-steps .step').forEach((s, index) => {
        if (index + 1 <= step) {
            s.classList.add('active');
        } else {
            s.classList.remove('active');
        }
    });
}

// PDF 파일 업로드 설정
function setupPdfUpload() {
    const uploadArea = document.getElementById('pdfUploadArea');
    const fileInput = document.getElementById('pdfFile');
    const fileInfo = document.getElementById('pdfFileInfo');

    if (!uploadArea || !fileInput) return;

    // 클릭 이벤트
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // 파일 선택 이벤트
    fileInput.addEventListener('change', (e) => {
        handlePdfSelect(e.target.files[0]);
    });

    // 드래그 앤 드롭 이벤트
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('dragover');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('dragover');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) {
            handlePdfSelect(file);
        }
    });
}

// PDF 파일 선택 처리
function handlePdfSelect(file) {
    if (!file) return;

    // 파일 크기 검증
    const maxFileSize = 100 * 1024 * 1024;
    if (file.size > maxFileSize) {
        showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        return;
    }

    // 파일 형식 검증
    if (!file.name.toLowerCase().endsWith('.pdf')) {
        showError('지원하지 않는 파일 형식입니다. PDF 파일만 업로드 가능합니다.');
        return;
    }

    selectedPdfFile = file;
    displayPdfFileInfo(file);
    updateProcessButton();
}

// PDF 파일 정보 표시
function displayPdfFileInfo(file) {
    const fileInfo = document.getElementById('pdfFileInfo');
    const fileName = fileInfo.querySelector('.file-name');
    
    fileName.textContent = file.name;
    fileInfo.style.display = 'flex';
}

// PDF 파일 제거
function removePdfFile() {
    selectedPdfFile = null;
    document.getElementById('pdfFile').value = '';
    document.getElementById('pdfFileInfo').style.display = 'none';
    updateProcessButton();
}

// Excel 파일 업로드 설정
function setupExcelUpload() {
    const uploadArea = document.getElementById('excelUploadArea');
    const fileInput = document.getElementById('excelFile');
    const fileInfo = document.getElementById('excelFileInfo');

    if (!uploadArea || !fileInput) return;

    // 클릭 이벤트
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // 파일 선택 이벤트
    fileInput.addEventListener('change', async (e) => {
        await handleExcelSelect(e.target.files[0]);
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
            await handleExcelSelect(file);
        }
    });
}

// Excel 파일 선택 처리
async function handleExcelSelect(file) {
    if (!file) return;

    // 파일 크기 검증
    const maxFileSize = 100 * 1024 * 1024;
    if (file.size > maxFileSize) {
        showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        return;
    }

    // 파일 형식 검증
    const allowedExtensions = ['.xlsx', '.xls'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();

    if (!allowedExtensions.includes(fileExtension)) {
        showError('지원하지 않는 파일 형식입니다. .xlsx 또는 .xls 파일만 업로드 가능합니다.');
        return;
    }

    selectedExcelFile = file;
    displayExcelFileInfo(file);
    
    // 연도/분기 옵션 업데이트 (시트는 자동 감지되므로 시트 목록 로드 불필요)
    await updateYearQuarterFromExcel(file);
    
    // 연도/분기 섹션 표시
    document.getElementById('periodSection').style.display = 'block';
    document.getElementById('formatSection').style.display = 'block';
    updateWorkflowStep(2);
    
    updateProcessButton();
}

// Excel 파일 정보 표시
function displayExcelFileInfo(file) {
    const fileInfo = document.getElementById('excelFileInfo');
    const fileName = fileInfo.querySelector('.file-name');
    
    fileName.textContent = file.name;
    fileInfo.style.display = 'flex';
}

// Excel 파일 제거
function removeExcelFile() {
    selectedExcelFile = null;
    document.getElementById('excelFile').value = '';
    document.getElementById('excelFileInfo').style.display = 'none';
    
    // 섹션 숨기기
    document.getElementById('periodSection').style.display = 'none';
    document.getElementById('formatSection').style.display = 'none';
    
    updateProcessButton();
    updateWorkflowStep(1);
}

// 엑셀 파일에서 연도/분기 정보 가져오기
async function updateYearQuarterFromExcel(file) {
    try {
        const formData = new FormData();
        formData.append('excel_file', file);
        
        const response = await fetch('/api/validate', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok && data.valid && data.sheets_info) {
            sheetsInfo = data.sheets_info;
            
            // 첫 번째 시트의 연도/분기 정보 사용 (백엔드에서 자동으로 필요한 시트를 찾을 것)
            const firstSheetName = Object.keys(data.sheets_info)[0];
            if (firstSheetName && data.sheets_info[firstSheetName]) {
                updateYearQuarterOptions(firstSheetName);
            }
        }
    } catch (error) {
        console.error('연도/분기 정보 로드 오류:', error);
        // 에러가 발생해도 기본값 사용
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
    
    if (selectedPdfFile && selectedExcelFile) {
        processBtn.disabled = false;
    } else {
        processBtn.disabled = true;
    }
}

// 보도자료 생성 처리
async function handleProcess() {
    if (!selectedPdfFile || !selectedExcelFile) {
        showError('PDF 파일과 엑셀 파일을 모두 업로드해주세요.');
        return;
    }

    // 연도 및 분기 가져오기 (시트는 백엔드에서 자동 감지)
    const yearSelect = document.getElementById('yearSelect');
    const quarterSelect = document.getElementById('quarterSelect');
    
    const year = yearSelect.value;
    const quarter = quarterSelect.value;
    
    // 출력 포맷 가져오기
    const formatRadio = document.querySelector('input[name="outputFormat"]:checked');
    const outputFormat = formatRadio ? formatRadio.value : 'pdf';
    
    // 진행 상황 텍스트를 포맷에 맞게 업데이트
    updateProgressTexts(outputFormat);

    // UI 업데이트
    const processBtn = document.getElementById('processBtn');
    const btnText = processBtn.querySelector('.btn-text');
    const btnLoader = processBtn.querySelector('.btn-loader');
    
    processBtn.disabled = true;
    btnText.style.display = 'none';
    btnLoader.style.display = 'inline-block';
    
    hideError();
    hideResult();
    
    // 진행 상황 섹션 표시
    const progressSection = document.getElementById('progressSection');
    progressSection.style.display = 'block';
    
    // 시간 추정 초기화
    stepStartTimes = {};
    currentStep = null;
    currentStepStartTime = null;
    
    // 첫 번째 단계 시작
    startStep('step1');
    
    updateProgress(0);

    try {
        // FormData 생성 (시트명은 백엔드에서 자동 감지)
        const formData = new FormData();
        formData.append('pdf_file', selectedPdfFile);
        formData.append('excel_file', selectedExcelFile);
        formData.append('year', year);
        formData.append('quarter', quarter);
        formData.append('output_format', outputFormat);

        // API 호출 (비동기 처리 시작)
        // 주의: 현재 구조상 process-word-template이 동기적으로 실행되므로
        // 응답이 올 때는 이미 완료된 상태입니다.
        // 하지만 진행 상황을 실시간으로 보여주기 위해 폴링을 사용합니다.
        
        // 백그라운드에서 처리 시작 (실제로는 동기적이지만 진행 상황은 폴링으로 확인)
        fetch('/api/process-word-template', {
            method: 'POST',
            body: formData
        })
        .then(response => response.json())
        .then(data => {
            if (data.success && data.session_id) {
                // 성공적으로 완료되었고 세션 ID도 있는 경우
                // 결과 정보 저장
                currentOutputFilename = data.output_filename;
                currentOutputFormat = data.output_format || outputFormat;
                currentSessionId = data.session_id;
                
                // 폴링 시작 (이미 완료되었을 수 있지만 최신 진행 상황 확인)
                startProgressPolling(data.session_id);
                
                // 폴링이 완료 상태를 확인하면 결과 표시
            } else if (data.success) {
                // 즉시 완료된 경우 (세션 ID 없이 완료)
                stopProgressPolling();
                currentOutputFilename = data.output_filename;
                currentOutputFormat = data.output_format || outputFormat;
                updateProgress(100, 7);
                setTimeout(() => {
                    progressSection.style.display = 'none';
                    showResult(data.message, currentOutputFormat);
                    updateWorkflowStep(3);
                }, 1000);
            } else if (data.session_id) {
                // 세션 ID만 있고 아직 완료되지 않은 경우 (이론적으로는 발생하지 않음)
                currentSessionId = data.session_id;
                startProgressPolling(data.session_id);
            } else {
                // 에러 발생
                stopProgressPolling();
                progressSection.style.display = 'none';
                showError(data.error || '처리 중 오류가 발생했습니다.');
            }
        })
        .catch(error => {
            console.error('처리 오류:', error);
            stopProgressPolling();
            progressSection.style.display = 'none';
            showError('서버와 통신하는 중 오류가 발생했습니다.');
        });
    } catch (error) {
        console.error('처리 오류:', error);
        stopProgressPolling();
        progressSection.style.display = 'none';
        if (error.message && error.message.includes('413')) {
            showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        } else {
            showError('서버와 통신하는 중 오류가 발생했습니다.');
        }
    } finally {
        // UI 복원
        processBtn.disabled = false;
        btnText.style.display = 'inline';
        btnLoader.style.display = 'none';
        updateProcessButton();
    }
}

// 진행 상황 폴링 시작
function startProgressPolling(sessionId) {
    // 기존 폴링 중지
    stopProgressPolling();
    
    // 폴링 시작 (500ms 간격)
    progressPollingInterval = setInterval(async () => {
        try {
            const response = await fetch(`/api/progress/${sessionId}`);
            const data = await response.json();
            
            if (response.ok && !data.error) {
                // 진행 상황 업데이트
                updateProgressFromBackend(data);
                
                // 완료 확인
                if (data.progress >= 100 && data.result) {
                    stopProgressPolling();
                    
                    // 결과 정보 가져오기
                    const result = data.result;
                    currentOutputFilename = result.output_filename;
                    currentOutputFormat = result.output_format || currentOutputFormat;
                    
                    // 진행 상황 숨기고 결과 표시
                    setTimeout(() => {
                        const progressSection = document.getElementById('progressSection');
                        if (progressSection) {
                            progressSection.style.display = 'none';
                        }
                        showResult(result.message, currentOutputFormat);
                        updateWorkflowStep(3);
                    }, 1000);
                } else if (data.progress >= 100) {
                    // 완료되었지만 결과 정보가 아직 없는 경우 (약간 대기)
                    // 폴링 계속 (결과 정보가 추가될 때까지)
                } else if (data.step === 0) {
                    // 에러 발생
                    stopProgressPolling();
                    const progressSection = document.getElementById('progressSection');
                    if (progressSection) {
                        progressSection.style.display = 'none';
                    }
                    showError(data.message || '처리 중 오류가 발생했습니다.');
                }
            } else {
                // 진행 상황을 찾을 수 없음 (타임아웃 또는 완료)
                if (data.error && data.error.includes('만료')) {
                    stopProgressPolling();
                    const progressSection = document.getElementById('progressSection');
                    progressSection.style.display = 'none';
                    showError('처리 시간이 초과되었습니다. 다시 시도해주세요.');
                }
            }
        } catch (error) {
            console.error('진행 상황 조회 오류:', error);
            // 폴링은 계속 진행 (일시적 네트워크 오류일 수 있음)
        }
    }, 500);
}

// 진행 상황 폴링 중지
function stopProgressPolling() {
    if (progressPollingInterval) {
        clearInterval(progressPollingInterval);
        progressPollingInterval = null;
    }
}

// 백엔드에서 받은 진행 상황으로 UI 업데이트
function updateProgressFromBackend(progressData) {
    const progress = progressData.progress || 0;
    const step = progressData.step || 1;
    const stepName = progressData.step_name || '';
    const message = progressData.message || '';
    const pageInfo = progressData.page_info || {current: 0, total: 0};
    const ocrProgress = progressData.ocr_progress;
    const ocrTimes = progressData.ocr_times || {};
    
    // OCR 시간 정보 업데이트
    Object.keys(ocrTimes).forEach(pageNum => {
        if (!pageOcrTimes[pageNum]) {
            pageOcrTimes[pageNum] = [];
        }
        const time = ocrTimes[pageNum];
        // 중복 방지: 최근 값과 다를 때만 추가
        if (pageOcrTimes[pageNum].length === 0 || 
            pageOcrTimes[pageNum][pageOcrTimes[pageNum].length - 1] !== time) {
            pageOcrTimes[pageNum].push(time);
            // 최근 10개만 유지
            if (pageOcrTimes[pageNum].length > 10) {
                pageOcrTimes[pageNum].shift();
            }
        }
    });
    
    // 진행률 업데이트 (백엔드 단계 정보 포함)
    updateProgress(progress, step, progressData);
    
    // 단계별 텍스트 업데이트 (OCR 진행률 포함)
    updateStepTexts(step, stepName, message, pageInfo, ocrProgress);
    
    // 완료 시 결과 처리는 폴링 루프에서 처리
}

// 단계별 텍스트 동적 업데이트
function updateStepTexts(step, stepName, message, pageInfo, ocrProgress) {
    // 단계별 요소 찾기
    const stepElements = {
        1: { text: document.getElementById('step1Text'), time: document.getElementById('step1Time') },
        2: { text: document.getElementById('step2Text'), time: document.getElementById('step2Time') },
        3: { text: document.getElementById('step3Text'), time: document.getElementById('step3Time') },
        4: { text: document.getElementById('step4Text'), time: document.getElementById('step4Time') }
    };
    
    // 모든 단계 업데이트
    Object.keys(stepElements).forEach(stepNum => {
        const stepNumInt = parseInt(stepNum);
        const element = stepElements[stepNumInt];
        
        if (element && element.text) {
            // 현재 단계인 경우
            if (stepNumInt === step) {
                let displayText = stepName;
                
                // 페이지 정보가 있으면 추가
                if (pageInfo && pageInfo.total > 0) {
                    displayText += ` (${pageInfo.current}/${pageInfo.total})`;
                }
                
                // OCR 진행률이 있으면 추가 (step1인 경우)
                if (stepNumInt === 1 && ocrProgress !== undefined && ocrProgress !== null) {
                    displayText += ` - OCR ${ocrProgress}%`;
                }
                
                // 메시지가 있으면 추가 (stepName과 다를 때만)
                if (message && message !== stepName) {
                    displayText = message;
                    if (pageInfo && pageInfo.total > 0) {
                        displayText += ` (${pageInfo.current}/${pageInfo.total})`;
                    }
                    // OCR 진행률이 있으면 추가
                    if (stepNumInt === 1 && ocrProgress !== undefined && ocrProgress !== null && ocrProgress < 100) {
                        displayText += ` (${ocrProgress}%)`;
                    }
                }
                
                element.text.textContent = displayText;
            }
        }
    });
}

// 단계 시작
function startStep(stepId) {
    if (currentStep && currentStep !== stepId) {
        // 이전 단계 종료 시간 기록
        endStep(currentStep);
    }
    currentStep = stepId;
    currentStepStartTime = Date.now();
    stepStartTimes[stepId] = currentStepStartTime;
}

// 단계 종료
function endStep(stepId) {
    if (stepStartTimes[stepId]) {
        const duration = Date.now() - stepStartTimes[stepId];
        if (stepDurations[stepId]) {
            stepDurations[stepId].push(duration);
            // 최근 5개만 유지
            if (stepDurations[stepId].length > 5) {
                stepDurations[stepId].shift();
            }
        }
    }
}

// 평균 시간 계산
function getAverageTime(stepId) {
    const times = stepDurations[stepId] || [];
    if (times.length === 0) return null;
    return times.reduce((a, b) => a + b, 0) / times.length;
}

// 가중 이동 평균(EMA) 계산
function calculateEMA(times, alpha = 0.3) {
    if (times.length === 0) return null;
    let ema = times[0];
    for (let i = 1; i < times.length; i++) {
        ema = alpha * times[i] + (1 - alpha) * ema;
    }
    return ema;
}

// 페이지별 OCR 평균 시간 계산
function getAverageOcrTime(pageNum) {
    const times = pageOcrTimes[pageNum] || [];
    if (times.length === 0) return null;
    // EMA 사용 (최근 데이터에 더 높은 가중치)
    return calculateEMA(times, 0.4);
}

// 전체 OCR 평균 시간 계산 (모든 페이지)
function getOverallAverageOcrTime() {
    const allTimes = [];
    Object.values(pageOcrTimes).forEach(times => {
        allTimes.push(...times);
    });
    if (allTimes.length === 0) return null;
    return calculateEMA(allTimes, 0.3);
}

// 남은 시간 추정 (개선된 알고리즘)
function estimateRemainingTime(currentStepId, currentProgress, progressData = null) {
    const steps = ['step1', 'step2', 'step3', 'step4'];
    const currentIndex = steps.indexOf(currentStepId);
    
    if (currentIndex === -1) return null;
    
    let remainingTime = 0;
    
    // step1 (PDF to Word 변환)인 경우 OCR 시간 기반 추정
    if (currentStepId === 'step1' && progressData) {
        const pageInfo = progressData.page_info || {};
        const ocrProgress = progressData.ocr_progress || 0;
        const ocrTimes = progressData.ocr_times || {};
        const currentPage = pageInfo.current || 0;
        const totalPages = pageInfo.total || 1;
        
        if (currentPage > 0 && totalPages > 0) {
            // 현재 페이지의 OCR 진행률 고려
            const currentPageOcrAvg = getAverageOcrTime(currentPage);
            const overallOcrAvg = getOverallAverageOcrTime() || currentPageOcrAvg;
            
            if (currentPageOcrAvg) {
                // 현재 페이지 남은 OCR 시간
                const currentPageRemaining = currentPageOcrAvg * (1 - ocrProgress / 100);
                remainingTime += currentPageRemaining;
            } else if (overallOcrAvg) {
                // 전체 평균 사용
                const currentPageRemaining = overallOcrAvg * (1 - ocrProgress / 100);
                remainingTime += currentPageRemaining;
            }
            
            // 남은 페이지들의 예상 OCR 시간
            const remainingPages = totalPages - currentPage;
            if (remainingPages > 0) {
                const avgOcrTime = currentPageOcrAvg || overallOcrAvg;
                if (avgOcrTime) {
                    remainingTime += remainingPages * avgOcrTime;
                } else {
                    // 기본 OCR 시간 (페이지당 5초)
                    remainingTime += remainingPages * 5000;
                }
            }
            
            // Word 문서 생성 시간 추가 (페이지당 1초)
            remainingTime += totalPages * 1000;
            
            return remainingTime;
        }
    }
    
    // 현재 단계 남은 시간 (EMA 사용)
    if (currentStepStartTime) {
        const elapsed = Date.now() - currentStepStartTime;
        const avgTime = calculateEMA(stepDurations[currentStepId] || [], 0.3);
        
        if (avgTime) {
            const estimatedTotal = avgTime;
            const remaining = Math.max(0, estimatedTotal - elapsed);
            remainingTime += remaining;
        } else {
            // 평균 시간이 없으면 현재 진행률 기반 추정
            if (currentProgress > 0) {
                const estimatedTotal = elapsed / (currentProgress / 100);
                const remaining = Math.max(0, estimatedTotal - elapsed);
                remainingTime += remaining;
            }
        }
    }
    
    // 남은 단계들의 예상 시간 (EMA 사용)
    for (let i = currentIndex + 1; i < steps.length; i++) {
        const stepId = steps[i];
        const avgTime = calculateEMA(stepDurations[stepId] || [], 0.3);
        if (avgTime) {
            remainingTime += avgTime;
        } else {
            // 기본 추정 시간 (초)
            const defaultTimes = {
                step1: 30000, // 30초
                step2: 5000,  // 5초
                step3: 15000, // 15초
                step4: 10000  // 10초
            };
            remainingTime += defaultTimes[stepId] || 10000;
        }
    }
    
    return remainingTime;
}

// 시간 포맷팅
function formatTime(ms) {
    if (!ms || ms < 0) return '';
    const seconds = Math.ceil(ms / 1000);
    if (seconds < 60) {
        return `약 ${seconds}초`;
    }
    const minutes = Math.floor(seconds / 60);
    const remainingSeconds = seconds % 60;
    if (remainingSeconds === 0) {
        return `약 ${minutes}분`;
    }
    return `약 ${minutes}분 ${remainingSeconds}초`;
}

// 진행 상황 텍스트 업데이트 (포맷에 따라)
function updateProgressTexts(format) {
    const step1Text = document.getElementById('step1Text');
    const step4Text = document.getElementById('step4Text');
    
    if (step1Text) {
        step1Text.textContent = 'PDF를 Word 템플릿으로 변환 중...';
    }
    
    if (step4Text) {
        if (format === 'word') {
            step4Text.textContent = 'Word 파일 생성 중...';
        } else {
            step4Text.textContent = 'PDF로 변환 중...';
        }
    }
}

// 진행 상황 업데이트
function updateProgress(percentage, backendStep = null, progressData = null) {
    const progressBar = document.getElementById('progressBar');
    const progressPercentage = document.getElementById('progressPercentage');
    
    progressBar.style.width = percentage + '%';
    if (progressPercentage) {
        progressPercentage.textContent = Math.round(percentage) + '%';
    }
    
    // 백엔드에서 받은 단계 정보가 있으면 사용, 없으면 진행률 기반 추정
    let currentStepNum = backendStep;
    if (!currentStepNum) {
        // 진행률 기반으로 단계 추정
        if (percentage < 25) currentStepNum = 1;
        else if (percentage < 50) currentStepNum = 2;
        else if (percentage < 75) currentStepNum = 3;
        else currentStepNum = 4;
    }
    
    // 단계별 아이콘 및 시간 업데이트
    const steps = [
        { id: 'step1', stepNum: 1, threshold: 25 },
        { id: 'step2', stepNum: 2, threshold: 50 },
        { id: 'step3', stepNum: 3, threshold: 75 },
        { id: 'step4', stepNum: 4, threshold: 100 }
    ];
    
    let activeStepId = null;
    
    steps.forEach((step, index) => {
        const stepElement = document.getElementById(step.id);
        const icon = stepElement.querySelector('.progress-icon');
        const timeElement = document.getElementById(step.id + 'Time');
        
        // 백엔드 단계 정보를 우선 사용
        if (currentStepNum && step.stepNum === currentStepNum) {
            // 현재 진행 중인 단계
            if (!activeStepId) {
                activeStepId = step.id;
                startStep(step.id);
            }
            icon.textContent = '⏳';
            stepElement.classList.add('active');
            stepElement.classList.remove('completed');
            
            // 남은 시간 추정
            if (timeElement && currentStepStartTime) {
                const remaining = estimateRemainingTime(step.id, percentage, progressData);
                if (remaining !== null) {
                    timeElement.textContent = formatTime(remaining) + ' 남음';
                }
            }
        } else if (currentStepNum && step.stepNum < currentStepNum) {
            // 완료된 단계
            icon.textContent = '✅';
            stepElement.classList.add('completed');
            stepElement.classList.remove('active');
            if (timeElement) {
                const duration = stepDurations[step.id]?.[stepDurations[step.id].length - 1];
                if (duration) {
                    timeElement.textContent = `완료 (${formatTime(duration)})`;
                } else {
                    timeElement.textContent = '완료';
                }
            }
            endStep(step.id);
        } else if (percentage >= step.threshold) {
            // 진행률 기반 완료 판단 (백엔드 정보가 없을 때)
            icon.textContent = '✅';
            stepElement.classList.add('completed');
            stepElement.classList.remove('active');
            if (timeElement) {
                const duration = stepDurations[step.id]?.[stepDurations[step.id].length - 1];
                if (duration) {
                    timeElement.textContent = `완료 (${formatTime(duration)})`;
                } else {
                    timeElement.textContent = '완료';
                }
            }
            endStep(step.id);
        } else if (percentage >= step.threshold - 10 && !currentStepNum) {
            // 진행 중인 단계 (백엔드 정보 없을 때만)
            if (!activeStepId) {
                activeStepId = step.id;
                startStep(step.id);
            }
            icon.textContent = '⏳';
            stepElement.classList.add('active');
            stepElement.classList.remove('completed');
            
            // 남은 시간 추정
            if (timeElement && currentStepStartTime) {
                const remaining = estimateRemainingTime(step.id, percentage, progressData);
                if (remaining !== null) {
                    timeElement.textContent = formatTime(remaining) + ' 남음';
                }
            }
        } else {
            // 대기 중인 단계
            icon.textContent = '⏸️';
            stepElement.classList.remove('active', 'completed');
            if (timeElement) {
                const avgTime = getAverageTime(step.id);
                if (avgTime) {
                    timeElement.textContent = `예상: ${formatTime(avgTime)}`;
                } else {
                    timeElement.textContent = '';
                }
            }
        }
    });
    
    // 전체 남은 시간 표시
    const timeEstimate = document.getElementById('progressTimeEstimate');
    if (timeEstimate && activeStepId) {
        const remaining = estimateRemainingTime(activeStepId, percentage, progressData);
        if (remaining !== null && remaining > 0) {
            timeEstimate.textContent = `⏱️ 예상 남은 시간: ${formatTime(remaining)}`;
        } else {
            timeEstimate.textContent = '';
        }
    }
}

// 결과 표시
function showResult(message, format = 'pdf') {
    const resultSection = document.getElementById('resultSection');
    const resultMessage = document.getElementById('resultMessage');
    
    resultMessage.textContent = message;
    resultSection.style.display = 'block';
    
    // 다운로드 버튼 설정
    setupDownloadButton(format);
    
    // 결과 섹션으로 스크롤
    resultSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// 결과 숨기기
function hideResult() {
    document.getElementById('resultSection').style.display = 'none';
}

// 다운로드 버튼 설정
function setupDownloadButton(format = 'pdf') {
    const downloadBtn = document.getElementById('downloadBtn');
    
    // 버튼 텍스트 업데이트
    const formatText = format === 'word' ? 'Word' : 'PDF';
    downloadBtn.innerHTML = `<span>📥 ${formatText} 다운로드</span>`;
    
    downloadBtn.onclick = () => {
        if (currentOutputFilename) {
            window.location.href = `/api/download/${encodeURIComponent(currentOutputFilename)}`;
        }
    };
}

// 연도/분기 옵션 업데이트
function updateYearQuarterOptions(sheetName) {
    if (!sheetsInfo || !sheetsInfo[sheetName]) {
        return;
    }
    
    const sheetInfo = sheetsInfo[sheetName];
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
    
    // 분기 옵션 업데이트
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

// 에러 표시
function showError(message) {
    const errorSection = document.getElementById('errorSection');
    const errorMessage = document.getElementById('errorMessage');
    
    errorMessage.textContent = message;
    errorSection.style.display = 'block';
    
    // 스크롤하여 에러 메시지가 보이도록
    errorSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

// 에러 숨기기
function hideError() {
    document.getElementById('errorSection').style.display = 'none';
}
