// 전역 변수
let selectedPdfFile = null;
let selectedExcelFile = null;
let currentOutputFilename = null;
let currentOutputFormat = 'pdf';
let sheetsInfo = {};

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
    updateProgress(0);

    try {
        // FormData 생성 (시트명은 백엔드에서 자동 감지)
        const formData = new FormData();
        formData.append('pdf_file', selectedPdfFile);
        formData.append('excel_file', selectedExcelFile);
        formData.append('year', year);
        formData.append('quarter', quarter);
        formData.append('output_format', outputFormat);

        // 진행 상황 시뮬레이션
        simulateProgress();

        // API 호출
        const response = await fetch('/api/process-word-template', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (response.ok && data.success) {
            currentOutputFilename = data.output_filename;
            currentOutputFormat = data.output_format || outputFormat;
            updateProgress(100);
            setTimeout(() => {
                progressSection.style.display = 'none';
                showResult(data.message, currentOutputFormat);
                updateWorkflowStep(3);
            }, 500);
        } else {
            progressSection.style.display = 'none';
            if (response.status === 413) {
                showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
            } else {
                showError(data.error || '처리 중 오류가 발생했습니다.');
            }
        }
    } catch (error) {
        console.error('처리 오류:', error);
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

// 진행 상황 시뮬레이션
function simulateProgress() {
    let progress = 0;
    const interval = setInterval(() => {
        progress += 5;
        if (progress <= 90) {
            updateProgress(progress);
        } else {
            clearInterval(interval);
        }
    }, 500);
}

// 진행 상황 업데이트
function updateProgress(percentage) {
    const progressBar = document.getElementById('progressBar');
    progressBar.style.width = percentage + '%';
    
    // 단계별 아이콘 업데이트
    const steps = [
        { id: 'step1', threshold: 25 },
        { id: 'step2', threshold: 50 },
        { id: 'step3', threshold: 75 },
        { id: 'step4', threshold: 100 }
    ];
    
    steps.forEach((step, index) => {
        const stepElement = document.getElementById(step.id);
        const icon = stepElement.querySelector('.progress-icon');
        const text = stepElement.querySelector('.progress-text');
        
        if (percentage >= step.threshold) {
            icon.textContent = '✅';
            stepElement.classList.add('completed');
        } else if (percentage >= step.threshold - 10) {
            icon.textContent = '⏳';
            stepElement.classList.add('active');
        } else {
            icon.textContent = '⏸️';
            stepElement.classList.remove('active', 'completed');
        }
    });
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
