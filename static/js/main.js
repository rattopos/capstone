// 전역 변수
let selectedExcelFile = null;
let selectedImageFile = null;
let selectedExcelFileImage = null;
let currentOutputFilename = null;
let currentMode = 'sheet'; // 'sheet' or 'image'
let currentTrainingStatusId = null;
let trainingStatusInterval = null;

// DOM 로드 완료 시 초기화
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

// 앱 초기화
function initializeApp() {
    setupFileUpload();
    setupProcessButton();
    setupTemplateCreate();
    setupTemplateUse();
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

    selectedExcelFile = file;
    displayFileInfo(file);
    
    // 시트 목록 로드
    await loadSheetNames(file);
    
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
    
    // 시트 선택 초기화
    const sheetSelect = document.getElementById('sheetSelect');
    sheetSelect.innerHTML = '<option value="">엑셀 파일을 업로드하세요</option>';
    sheetSelect.disabled = true;
    
    updateProcessButton();
}

// 시트 목록 로드
async function loadSheetNames(file) {
    const sheetSelect = document.getElementById('sheetSelect');
    sheetSelect.disabled = true;
    sheetSelect.innerHTML = '<option value="">로딩 중...</option>';
    
    try {
        const formData = new FormData();
        formData.append('excel_file', file);
        
        const response = await fetch('/api/validate', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok && data.valid && data.sheet_names) {
            // 시트 목록 채우기
            sheetSelect.innerHTML = '';
            const defaultSheetName = '광공업생산';
            
            data.sheet_names.forEach(sheetName => {
                const option = document.createElement('option');
                option.value = sheetName;
                option.textContent = sheetName;
                sheetSelect.appendChild(option);
            });
            
            // 기본 시트 "광공업생산"이 있으면 선택, 없으면 첫 번째 시트 선택
            if (data.sheet_names.includes(defaultSheetName)) {
                sheetSelect.value = defaultSheetName;
            } else if (data.sheet_names.length > 0) {
                sheetSelect.value = data.sheet_names[0];
            }
            
            sheetSelect.disabled = false;
            
            // 시트별 연도/분기 정보 저장
            if (data.sheets_info) {
                window.sheetsInfo = data.sheets_info;
                // 선택된 시트의 연도/분기 업데이트
                updateYearQuarterOptions(sheetSelect.value);
            }
        } else {
            sheetSelect.innerHTML = '<option value="">시트를 불러올 수 없습니다</option>';
            showError(data.error || '시트 목록을 불러올 수 없습니다.');
        }
    } catch (error) {
        console.error('시트 목록 로드 오류:', error);
        sheetSelect.innerHTML = '<option value="">시트를 불러올 수 없습니다</option>';
        showError('시트 목록을 불러오는 중 오류가 발생했습니다.');
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
    const sheetSelect = document.getElementById('sheetSelect');
    
    if (selectedExcelFile && sheetSelect.value) {
        processBtn.disabled = false;
    } else {
        processBtn.disabled = true;
    }
}

// 보도자료 생성 처리
async function handleProcess() {
    if (!selectedExcelFile) {
        showError('엑셀 파일을 선택해주세요.');
        return;
    }

    // 시트, 연도 및 분기 가져오기
    const sheetSelect = document.getElementById('sheetSelect');
    const yearSelect = document.getElementById('yearSelect');
    const quarterSelect = document.getElementById('quarterSelect');
    
    const sheetName = sheetSelect.value;
    const year = yearSelect.value;
    const quarter = quarterSelect.value;
    
    if (!sheetName) {
        showError('시트를 선택해주세요.');
        return;
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

    try {
        // FormData 생성
        const formData = new FormData();
        formData.append('excel_file', selectedExcelFile);
        formData.append('sheet_name', sheetName);
        formData.append('year', year);
        formData.append('quarter', quarter);

        // API 호출
        const response = await fetch('/api/process', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (response.ok && data.success) {
            currentOutputFilename = data.output_filename;
            showResult(data.message);
        } else {
            // 413 에러 (파일 크기 초과) 처리
            if (response.status === 413) {
                showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
            } else {
                showError(data.error || '처리 중 오류가 발생했습니다.');
            }
        }
    } catch (error) {
        console.error('처리 오류:', error);
        // 네트워크 오류나 파일 크기 초과 등의 경우
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
}

// 결과 버튼 설정
function setupResultButtons() {
    const previewBtn = document.getElementById('previewBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    
    previewBtn.onclick = () => {
        if (currentOutputFilename) {
            showPreviewModal(`/api/preview/${currentOutputFilename}`);
        }
    };
    
    downloadBtn.onclick = () => {
        if (currentOutputFilename) {
            window.location.href = `/api/download/${currentOutputFilename}`;
        }
    };
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

// 시트 선택 변경 시 처리 버튼 상태 업데이트 및 연도/분기 옵션 업데이트
document.addEventListener('DOMContentLoaded', function() {
    const sheetSelect = document.getElementById('sheetSelect');
    if (sheetSelect) {
        sheetSelect.addEventListener('change', function() {
            updateProcessButton();
            if (window.sheetsInfo && this.value) {
                updateYearQuarterOptions(this.value);
            }
        });
    }
});

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

// 모드 전환
function switchMode(mode) {
    currentMode = mode;
    
    // 탭 버튼 업데이트
    document.querySelectorAll('.mode-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.dataset.mode === mode) {
            btn.classList.add('active');
        }
    });
    
    // 콘텐츠 표시/숨김
    const sheetMode = document.getElementById('sheetMode');
    const createMode = document.getElementById('createMode');
    const useMode = document.getElementById('useMode');
    
    if (sheetMode) sheetMode.style.display = mode === 'sheet' ? 'block' : 'none';
    if (createMode) createMode.style.display = mode === 'create' ? 'block' : 'none';
    if (useMode) {
        useMode.style.display = mode === 'use' ? 'block' : 'none';
        if (mode === 'use') {
            loadTemplateList();
        }
    }
}

// 템플릿 생성 모드 설정
function setupTemplateCreate() {
    const imageUploadArea = document.getElementById('imageUploadArea');
    const imageFileInput = document.getElementById('imageFile');
    const excelUploadAreaCreate = document.getElementById('excelUploadAreaCreate');
    const excelFileInputCreate = document.getElementById('excelFileCreate');
    const createBtn = document.getElementById('createTemplateBtn');
    
    // 이미지 파일 업로드
    if (imageUploadArea && imageFileInput) {
        imageUploadArea.addEventListener('click', () => imageFileInput.click());
        imageFileInput.addEventListener('change', async (e) => {
            await handleImageFileSelect(e.target.files[0]);
        });
        
        // 드래그 앤 드롭
        imageUploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            imageUploadArea.classList.add('dragover');
        });
        imageUploadArea.addEventListener('dragleave', () => {
            imageUploadArea.classList.remove('dragover');
        });
        imageUploadArea.addEventListener('drop', async (e) => {
            e.preventDefault();
            imageUploadArea.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) await handleImageFileSelect(file);
        });
    }
    
    // 엑셀 파일 업로드
    if (excelUploadAreaCreate && excelFileInputCreate) {
        excelUploadAreaCreate.addEventListener('click', () => excelFileInputCreate.click());
        excelFileInputCreate.addEventListener('change', async (e) => {
            await handleExcelFileSelectCreate(e.target.files[0]);
        });
        
        // 드래그 앤 드롭
        excelUploadAreaCreate.addEventListener('dragover', (e) => {
            e.preventDefault();
            excelUploadAreaCreate.classList.add('dragover');
        });
        excelUploadAreaCreate.addEventListener('dragleave', () => {
            excelUploadAreaCreate.classList.remove('dragover');
        });
        excelUploadAreaCreate.addEventListener('drop', async (e) => {
            e.preventDefault();
            excelUploadAreaCreate.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) await handleExcelFileSelectCreate(file);
        });
    }
    
    // 생성 버튼
    if (createBtn) {
        createBtn.addEventListener('click', handleCreateTemplate);
    }
}

// 템플릿 사용 모드 설정
function setupTemplateUse() {
    const excelUploadAreaUse = document.getElementById('excelUploadAreaUse');
    const excelFileInputUse = document.getElementById('excelFileUse');
    const processBtn = document.getElementById('processTemplateBtn');
    
    // 엑셀 파일 업로드
    if (excelUploadAreaUse && excelFileInputUse) {
        excelUploadAreaUse.addEventListener('click', () => excelFileInputUse.click());
        excelFileInputUse.addEventListener('change', async (e) => {
            await handleExcelFileSelectUse(e.target.files[0]);
        });
        
        // 드래그 앤 드롭
        excelUploadAreaUse.addEventListener('dragover', (e) => {
            e.preventDefault();
            excelUploadAreaUse.classList.add('dragover');
        });
        excelUploadAreaUse.addEventListener('dragleave', () => {
            excelUploadAreaUse.classList.remove('dragover');
        });
        excelUploadAreaUse.addEventListener('drop', async (e) => {
            e.preventDefault();
            excelUploadAreaUse.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) await handleExcelFileSelectUse(file);
        });
    }
    
    // 처리 버튼
    if (processBtn) {
        processBtn.addEventListener('click', handleProcessTemplate);
    }
    
    // 템플릿 선택 변경 시
    const templateSelect = document.getElementById('templateSelect');
    if (templateSelect) {
        templateSelect.addEventListener('change', updateProcessTemplateButton);
    }
}

// 템플릿 목록 로드
async function loadTemplateList() {
    const templateSelect = document.getElementById('templateSelect');
    if (!templateSelect) return;
    
    templateSelect.disabled = true;
    templateSelect.innerHTML = '<option value="">로딩 중...</option>';
    
    try {
        const response = await fetch('/api/templates/list');
        const data = await response.json();
        
        if (response.ok && data.success) {
            templateSelect.innerHTML = '<option value="">템플릿을 선택하세요</option>';
            
            data.templates.forEach(template => {
                const option = document.createElement('option');
                option.value = template.name;
                option.textContent = template.name;
                templateSelect.appendChild(option);
            });
            
            templateSelect.disabled = false;
        } else {
            templateSelect.innerHTML = '<option value="">템플릿을 불러올 수 없습니다</option>';
        }
    } catch (error) {
        console.error('템플릿 목록 로드 오류:', error);
        templateSelect.innerHTML = '<option value="">템플릿을 불러올 수 없습니다</option>';
    }
}

// 엑셀 파일 선택 처리 (생성 모드)
async function handleExcelFileSelectCreate(file) {
    if (!file) return;
    
    const maxFileSize = 100 * 1024 * 1024;
    if (file.size > maxFileSize) {
        showError('파일 크기가 너무 큽니다.');
        return;
    }
    
    const allowedExtensions = ['.xlsx', '.xls'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
    
    if (!allowedExtensions.includes(fileExtension)) {
        showError('지원하지 않는 파일 형식입니다.');
        return;
    }
    
    selectedExcelFileImage = file;
    displayExcelFileInfoCreate(file);
    updateCreateTemplateButton();
}

function displayExcelFileInfoCreate(file) {
    const fileInfo = document.getElementById('excelFileInfoCreate');
    const fileName = fileInfo?.querySelector('.file-name');
    if (fileName) {
        fileName.textContent = file.name;
        fileInfo.style.display = 'flex';
    }
}

function removeExcelFileCreate() {
    selectedExcelFileImage = null;
    document.getElementById('excelFileCreate').value = '';
    document.getElementById('excelFileInfoCreate').style.display = 'none';
    updateCreateTemplateButton();
}

// 엑셀 파일 선택 처리 (사용 모드)
async function handleExcelFileSelectUse(file) {
    if (!file) return;
    
    const maxFileSize = 100 * 1024 * 1024;
    if (file.size > maxFileSize) {
        showError('파일 크기가 너무 큽니다.');
        return;
    }
    
    const allowedExtensions = ['.xlsx', '.xls'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
    
    if (!allowedExtensions.includes(fileExtension)) {
        showError('지원하지 않는 파일 형식입니다.');
        return;
    }
    
    selectedExcelFileImage = file;
    displayExcelFileInfoUse(file);
    updateProcessTemplateButton();
}

function displayExcelFileInfoUse(file) {
    const fileInfo = document.getElementById('excelFileInfoUse');
    const fileName = fileInfo?.querySelector('.file-name');
    if (fileName) {
        fileName.textContent = file.name;
        fileInfo.style.display = 'flex';
    }
}

function removeExcelFileUse() {
    selectedExcelFileImage = null;
    document.getElementById('excelFileUse').value = '';
    document.getElementById('excelFileInfoUse').style.display = 'none';
    updateProcessTemplateButton();
}

// 템플릿 생성 버튼 상태 업데이트
function updateCreateTemplateButton() {
    const btn = document.getElementById('createTemplateBtn');
    const autoTrainCheckbox = document.getElementById('autoTrainCheckbox');
    const useAutoTrain = autoTrainCheckbox?.checked || false;
    
    if (btn) {
        if (useAutoTrain) {
            // 자동 학습 모드: 이미지만 필요
            btn.disabled = !selectedImageFile;
        } else {
            // 일반 모드: 이미지와 엑셀 파일 모두 필요
            btn.disabled = !(selectedImageFile && selectedExcelFileImage);
        }
    }
}

// 자동 학습 체크박스 변경 이벤트
document.addEventListener('DOMContentLoaded', function() {
    const autoTrainCheckbox = document.getElementById('autoTrainCheckbox');
    const autoTrainOptions = document.getElementById('autoTrainOptions');
    
    if (autoTrainCheckbox && autoTrainOptions) {
        autoTrainCheckbox.addEventListener('change', function() {
            autoTrainOptions.style.display = this.checked ? 'block' : 'none';
            updateCreateTemplateButton();
        });
        
        // 초기 상태 설정
        autoTrainOptions.style.display = autoTrainCheckbox.checked ? 'block' : 'none';
    }
});

// 템플릿 처리 버튼 상태 업데이트
function updateProcessTemplateButton() {
    const btn = document.getElementById('processTemplateBtn');
    const templateSelect = document.getElementById('templateSelect');
    if (btn && templateSelect) {
        btn.disabled = !(templateSelect.value && selectedExcelFileImage);
    }
}

// 템플릿 생성 처리
async function handleCreateTemplate() {
    if (!selectedImageFile) {
        showError('이미지 파일을 선택해주세요.');
        return;
    }
    
    const templateNameInput = document.getElementById('templateNameInput');
    const templateName = templateNameInput?.value || selectedImageFile.name.split('.').slice(0, -1).join('.');
    
    if (!templateName) {
        showError('템플릿 이름을 입력해주세요.');
        return;
    }
    
    const yearSelect = document.getElementById('yearSelectCreate');
    const quarterSelect = document.getElementById('quarterSelectCreate');
    const year = yearSelect?.value || '2025';
    const quarter = quarterSelect?.value || '2';
    
    // 자동 학습 옵션 확인
    const autoTrainCheckbox = document.getElementById('autoTrainCheckbox');
    const useAutoTrain = autoTrainCheckbox?.checked || false;
    const maxIterations = document.getElementById('maxIterationsInput')?.value || '10';
    const similarityThreshold = document.getElementById('similarityThresholdInput')?.value || '0.85';
    
    const btn = document.getElementById('createTemplateBtn');
    const btnText = btn.querySelector('.btn-text');
    const btnLoader = btn.querySelector('.btn-loader');
    
    btn.disabled = true;
    btnText.style.display = 'none';
    btnLoader.style.display = 'inline-block';
    
    hideError();
    hideResultCreate();
    
    try {
        const formData = new FormData();
        
            if (useAutoTrain) {
                // 자동 학습 모드: 정답 이미지로 사용
                formData.append('reference_image', selectedImageFile);
                if (selectedExcelFileImage) {
                    formData.append('excel_file', selectedExcelFileImage);
                }
                formData.append('template_name', templateName);
                formData.append('year', year);
                formData.append('quarter', quarter);
                formData.append('max_iterations', maxIterations);
                formData.append('similarity_threshold', similarityThreshold);
                
                const response = await fetch('/api/auto_train', {
                    method: 'POST',
                    body: formData
                });
                
                const data = await response.json();
                
                if (response.ok && data.success) {
                    // 상태 ID 저장 및 진행 상황 모니터링 시작
                    currentTrainingStatusId = data.status_id;
                    showTrainingProgress();
                    startTrainingStatusPolling(data.status_id);
                    
                    // 중단 버튼 이벤트
                    const stopBtn = document.getElementById('stopTrainingBtn');
                    if (stopBtn) {
                        stopBtn.onclick = () => stopTraining(data.status_id);
                    }
                } else {
                    showError(data.error || '자동 학습 시작 중 오류가 발생했습니다.');
                }
            } else {
            // 기존 모드: 일반 템플릿 생성
            if (!selectedExcelFileImage) {
                showError('엑셀 파일을 선택해주세요.');
                return;
            }
            
            formData.append('image_file', selectedImageFile);
            formData.append('excel_file', selectedExcelFileImage);
            formData.append('template_name', templateName);
            formData.append('year', year);
            formData.append('quarter', quarter);
            
            const response = await fetch('/api/create_template', {
                method: 'POST',
                body: formData
            });
            
            const data = await response.json();
            
            if (response.ok && data.success) {
                showResultCreate(data.message);
            } else {
                showError(data.error || '템플릿 생성 중 오류가 발생했습니다.');
            }
        }
    } catch (error) {
        console.error('템플릿 생성 오류:', error);
        showError('서버와 통신하는 중 오류가 발생했습니다.');
    } finally {
        btn.disabled = false;
        btnText.style.display = 'inline';
        btnLoader.style.display = 'none';
        updateCreateTemplateButton();
    }
}

// 템플릿 사용 처리
async function handleProcessTemplate() {
    const templateSelect = document.getElementById('templateSelect');
    const templateName = templateSelect?.value;
    
    if (!templateName) {
        showError('템플릿을 선택해주세요.');
        return;
    }
    
    if (!selectedExcelFileImage) {
        showError('엑셀 파일을 선택해주세요.');
        return;
    }
    
    const yearSelect = document.getElementById('yearSelectUse');
    const quarterSelect = document.getElementById('quarterSelectUse');
    const year = yearSelect?.value || '2025';
    const quarter = quarterSelect?.value || '2';
    
    const btn = document.getElementById('processTemplateBtn');
    const btnText = btn.querySelector('.btn-text');
    const btnLoader = btn.querySelector('.btn-loader');
    
    btn.disabled = true;
    btnText.style.display = 'none';
    btnLoader.style.display = 'inline-block';
    
    hideError();
    hideResultUse();
    
    try {
        const formData = new FormData();
        formData.append('excel_file', selectedExcelFileImage);
        formData.append('template_name', templateName);
        formData.append('year', year);
        formData.append('quarter', quarter);
        
        const response = await fetch('/api/process_template', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok && data.success) {
            currentOutputFilename = data.output_filename;
            showResultUse(data.message);
        } else {
            showError(data.error || '처리 중 오류가 발생했습니다.');
        }
    } catch (error) {
        console.error('처리 오류:', error);
        showError('서버와 통신하는 중 오류가 발생했습니다.');
    } finally {
        btn.disabled = false;
        btnText.style.display = 'inline';
        btnLoader.style.display = 'none';
        updateProcessTemplateButton();
    }
}

function showResultCreate(message) {
    // 진행 상황 섹션 숨기기
    const progressSection = document.getElementById('trainingProgressSection');
    if (progressSection) {
        progressSection.style.display = 'none';
    }
    
    // 결과 섹션 표시
    const resultSection = document.getElementById('resultSectionCreate');
    const resultMessage = document.getElementById('resultMessageCreate');
    if (resultMessage) {
        resultMessage.textContent = message;
        resultMessage.style.whiteSpace = 'pre-line'; // 줄바꿈 지원
    }
    if (resultSection) {
        resultSection.style.display = 'block';
    }
}

function hideResultCreate() {
    const resultSection = document.getElementById('resultSectionCreate');
    if (resultSection) resultSection.style.display = 'none';
    
    const progressSection = document.getElementById('trainingProgressSection');
    if (progressSection) progressSection.style.display = 'none';
}

function showResultUse(message) {
    const resultSection = document.getElementById('resultSectionUse');
    const resultMessage = document.getElementById('resultMessageUse');
    if (resultMessage) resultMessage.textContent = message;
    if (resultSection) resultSection.style.display = 'block';
    setupResultButtonsUse();
}

function hideResultUse() {
    const resultSection = document.getElementById('resultSectionUse');
    if (resultSection) resultSection.style.display = 'none';
}

function setupResultButtonsUse() {
    const previewBtn = document.getElementById('previewBtnUse');
    const downloadBtn = document.getElementById('downloadBtnUse');
    
    if (previewBtn) {
        previewBtn.onclick = () => {
            if (currentOutputFilename) {
                showPreviewModal(`/api/preview/${currentOutputFilename}`);
            }
        };
    }
    
    if (downloadBtn) {
        downloadBtn.onclick = () => {
            if (currentOutputFilename) {
                window.location.href = `/api/download/${currentOutputFilename}`;
            }
        };
    }
}

// 학습 진행 상황 표시
function showTrainingProgress() {
    hideResultCreate();
    const progressSection = document.getElementById('trainingProgressSection');
    if (progressSection) {
        progressSection.style.display = 'block';
    }
}

// 학습 진행 상황 폴링
function startTrainingStatusPolling(statusId) {
    // 기존 인터벌 정리
    if (trainingStatusInterval) {
        clearInterval(trainingStatusInterval);
    }
    
    // 주기적으로 상태 확인
    trainingStatusInterval = setInterval(async () => {
        try {
            const response = await fetch(`/api/training_status/${statusId}`);
            const data = await response.json();
            
            if (response.ok && data.success) {
                updateTrainingProgress(data.status);
                
                // 완료 또는 중단 시 폴링 중지
                if (data.status.status === 'completed' || 
                    data.status.status === 'stopped' || 
                    data.status.status === 'error') {
                    clearInterval(trainingStatusInterval);
                    trainingStatusInterval = null;
                    
                    // 결과 표시
                    if (data.status.result) {
                        const result = data.status.result;
                        let message = `템플릿 "${data.status.template_name}" 자동 학습 완료\n`;
                        message += `최종 유사도: ${result.final_similarity.toFixed(3)}\n`;
                        message += `반복 횟수: ${result.iterations}회\n`;
                        if (result.improvements.length > 0) {
                            message += `개선 사항: ${result.improvements.length}개`;
                        }
                        showResultCreate(message);
                        
                        // 미리보기 버튼 설정 (템플릿 이름으로)
                        const previewBtn = document.getElementById('previewBtnCreate');
                        if (previewBtn) {
                            previewBtn.onclick = () => {
                                // 저장된 템플릿 미리보기
                                loadAndPreviewTemplate(data.status.template_name);
                            };
                        }
                    }
                }
            }
        } catch (error) {
            console.error('상태 확인 오류:', error);
        }
    }, 1000); // 1초마다 확인
}

// 학습 진행 상황 업데이트
function updateTrainingProgress(status) {
    const statusText = document.getElementById('progressStatusText');
    const percentage = document.getElementById('progressPercentage');
    const progressBar = document.getElementById('progressBar');
    const progressDetails = document.getElementById('progressDetails');
    const improvementsList = document.getElementById('improvementsList');
    const improvementsUl = document.getElementById('improvementsUl');
    
    if (statusText) {
        statusText.textContent = status.message || '처리 중...';
    }
    
    const progress = status.progress_percentage || 0;
    if (percentage) {
        percentage.textContent = `${Math.round(progress)}%`;
    }
    if (progressBar) {
        progressBar.style.width = `${progress}%`;
    }
    
    if (progressDetails) {
        let details = '';
        if (status.current_iteration > 0) {
            details += `반복: ${status.current_iteration}/${status.max_iterations} `;
        }
        if (status.similarity_score > 0) {
            details += `| 유사도: ${status.similarity_score.toFixed(3)}`;
        }
        progressDetails.textContent = details;
    }
    
    // 개선 내역 표시
    if (status.improvements && status.improvements.length > 0) {
        if (improvementsList) improvementsList.style.display = 'block';
        if (improvementsUl) {
            improvementsUl.innerHTML = '';
            status.improvements.forEach((improvement, index) => {
                const li = document.createElement('li');
                li.textContent = `[${improvement.iteration}회] ${improvement.description}`;
                improvementsUl.appendChild(li);
            });
        }
    }
}

// 학습 중단
async function stopTraining(statusId) {
    try {
        const response = await fetch(`/api/training_stop/${statusId}`, {
            method: 'POST'
        });
        
        const data = await response.json();
        
        if (response.ok && data.success) {
            // 상태 업데이트는 폴링에서 처리됨
            const stopBtn = document.getElementById('stopTrainingBtn');
            if (stopBtn) {
                stopBtn.disabled = true;
                stopBtn.textContent = '중단 중...';
            }
        } else {
            showError(data.error || '중단 요청 실패');
        }
    } catch (error) {
        console.error('중단 요청 오류:', error);
        showError('중단 요청 중 오류가 발생했습니다.');
    }
}

// 미리보기 모달 표시
function showPreviewModal(url) {
    const modal = document.getElementById('previewModal');
    const iframe = document.getElementById('previewIframe');
    
    if (modal && iframe) {
        iframe.src = url;
        modal.style.display = 'block';
    }
}

// 미리보기 모달 닫기
function closePreviewModal() {
    const modal = document.getElementById('previewModal');
    const iframe = document.getElementById('previewIframe');
    
    if (modal) {
        modal.style.display = 'none';
    }
    if (iframe) {
        iframe.src = '';
    }
}

// 모달 외부 클릭 시 닫기
window.onclick = function(event) {
    const modal = document.getElementById('previewModal');
    if (event.target === modal) {
        closePreviewModal();
    }
}

// 템플릿 로드 및 미리보기
async function loadAndPreviewTemplate(templateName) {
    try {
        // 템플릿 HTML 가져오기 (한글 템플릿 이름을 URL 인코딩)
        const encodedName = encodeURIComponent(templateName);
        const response = await fetch(`/api/templates/${encodedName}/html`);
        if (response.ok) {
            const html = await response.text();
            // Blob URL로 변환하여 미리보기
            const blob = new Blob([html], { type: 'text/html' });
            const url = URL.createObjectURL(blob);
            showPreviewModal(url);
        } else {
            const errorData = await response.json().catch(() => ({ error: '알 수 없는 오류' }));
            showError(`템플릿을 불러올 수 없습니다: ${errorData.error || '알 수 없는 오류'}`);
        }
    } catch (error) {
        console.error('템플릿 로드 오류:', error);
        showError('템플릿 미리보기 중 오류가 발생했습니다.');
    }
}

// 이미지 업로드 설정 (기존 함수 유지)
function setupImageUpload() {
    const imageUploadArea = document.getElementById('imageUploadArea');
    const imageFileInput = document.getElementById('imageFile');
    const excelUploadAreaImage = document.getElementById('excelUploadAreaImage');
    const excelFileInputImage = document.getElementById('excelFileImage');
    
    // 이미지 파일 업로드
    if (imageUploadArea && imageFileInput) {
        imageUploadArea.addEventListener('click', () => {
            imageFileInput.click();
        });
        
        imageFileInput.addEventListener('change', async (e) => {
            await handleImageFileSelect(e.target.files[0]);
        });
        
        // 드래그 앤 드롭
        imageUploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            imageUploadArea.classList.add('dragover');
        });
        
        imageUploadArea.addEventListener('dragleave', () => {
            imageUploadArea.classList.remove('dragover');
        });
        
        imageUploadArea.addEventListener('drop', async (e) => {
            e.preventDefault();
            imageUploadArea.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) {
                await handleImageFileSelect(file);
            }
        });
    }
    
    // 엑셀 파일 업로드 (이미지 모드)
    if (excelUploadAreaImage && excelFileInputImage) {
        excelUploadAreaImage.addEventListener('click', () => {
            excelFileInputImage.click();
        });
        
        excelFileInputImage.addEventListener('change', async (e) => {
            await handleExcelFileSelectImage(e.target.files[0]);
        });
        
        // 드래그 앤 드롭
        excelUploadAreaImage.addEventListener('dragover', (e) => {
            e.preventDefault();
            excelUploadAreaImage.classList.add('dragover');
        });
        
        excelUploadAreaImage.addEventListener('dragleave', () => {
            excelUploadAreaImage.classList.remove('dragover');
        });
        
        excelUploadAreaImage.addEventListener('drop', async (e) => {
            e.preventDefault();
            excelUploadAreaImage.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file) {
                await handleExcelFileSelectImage(file);
            }
        });
    }
}

// 이미지 파일 선택 처리
async function handleImageFileSelect(file) {
    if (!file) return;
    
    // 파일 크기 검증
    const maxFileSize = 100 * 1024 * 1024;
    if (file.size > maxFileSize) {
        showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        return;
    }
    
    // 파일 형식 검증
    const allowedExtensions = ['.png', '.jpg', '.jpeg', '.gif', '.bmp'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
    
    if (!allowedExtensions.includes(fileExtension)) {
        showError('지원하지 않는 이미지 형식입니다. .png, .jpg, .jpeg, .gif, .bmp 파일만 업로드 가능합니다.');
        return;
    }
    
    selectedImageFile = file;
    displayImageFileInfo(file);
    
    // 이미지 미리보기
    const reader = new FileReader();
    reader.onload = (e) => {
        const previewImg = document.getElementById('previewImg');
        const imagePreview = document.getElementById('imagePreview');
        if (previewImg && imagePreview) {
            previewImg.src = e.target.result;
            imagePreview.style.display = 'block';
        }
    };
    reader.readAsDataURL(file);
    
    // 이미지 분석 (선택적)
    // await analyzeImage(file);
    
    updateImageProcessButton();
}

// 이미지 파일 정보 표시
function displayImageFileInfo(file) {
    const fileInfo = document.getElementById('imageFileInfo');
    const fileName = fileInfo.querySelector('.file-name');
    
    if (fileName) {
        fileName.textContent = file.name;
        fileInfo.style.display = 'flex';
    }
}

// 이미지 파일 제거
function removeImageFile() {
    selectedImageFile = null;
    document.getElementById('imageFile').value = '';
    document.getElementById('imageFileInfo').style.display = 'none';
    document.getElementById('imagePreview').style.display = 'none';
    updateImageProcessButton();
}

// 엑셀 파일 선택 처리 (이미지 모드)
async function handleExcelFileSelectImage(file) {
    if (!file) return;
    
    // 파일 크기 검증
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
    
    selectedExcelFileImage = file;
    displayExcelFileInfoImage(file);
    updateImageProcessButton();
}

// 엑셀 파일 정보 표시 (이미지 모드)
function displayExcelFileInfoImage(file) {
    const fileInfo = document.getElementById('excelFileInfoImage');
    const fileName = fileInfo.querySelector('.file-name');
    
    if (fileName) {
        fileName.textContent = file.name;
        fileInfo.style.display = 'flex';
    }
}

// 엑셀 파일 제거 (이미지 모드)
function removeExcelFileImage() {
    selectedExcelFileImage = null;
    document.getElementById('excelFileImage').value = '';
    document.getElementById('excelFileInfoImage').style.display = 'none';
    updateImageProcessButton();
}

// 이미지 처리 버튼 설정
function setupImageProcessButton() {
    const processBtn = document.getElementById('processImageBtn');
    if (processBtn) {
        processBtn.addEventListener('click', handleImageProcess);
    }
}

// 이미지 처리 버튼 상태 업데이트
function updateImageProcessButton() {
    const processBtn = document.getElementById('processImageBtn');
    if (processBtn) {
        if (selectedImageFile && selectedExcelFileImage) {
            processBtn.disabled = false;
        } else {
            processBtn.disabled = true;
        }
    }
}

// 이미지 기반 보도자료 생성 처리
async function handleImageProcess() {
    if (!selectedImageFile) {
        showError('이미지 파일을 선택해주세요.');
        return;
    }
    
    if (!selectedExcelFileImage) {
        showError('엑셀 파일을 선택해주세요.');
        return;
    }
    
    // 연도 및 분기 가져오기
    const yearSelect = document.getElementById('yearSelectImage');
    const quarterSelect = document.getElementById('quarterSelectImage');
    
    const year = yearSelect ? yearSelect.value : '2025';
    const quarter = quarterSelect ? quarterSelect.value : '2';
    
    // UI 업데이트
    const processBtn = document.getElementById('processImageBtn');
    const btnText = processBtn.querySelector('.btn-text');
    const btnLoader = processBtn.querySelector('.btn-loader');
    
    processBtn.disabled = true;
    btnText.style.display = 'none';
    btnLoader.style.display = 'inline-block';
    
    hideError();
    hideResultImage();
    
    try {
        // FormData 생성
        const formData = new FormData();
        formData.append('image_file', selectedImageFile);
        formData.append('excel_file', selectedExcelFileImage);
        formData.append('year', year);
        formData.append('quarter', quarter);
        
        // API 호출
        const response = await fetch('/api/process_image', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok && data.success) {
            currentOutputFilename = data.output_filename;
            showResultImage(data.message);
        } else {
            if (response.status === 413) {
                showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
            } else {
                showError(data.error || '처리 중 오류가 발생했습니다.');
            }
        }
    } catch (error) {
        console.error('처리 오류:', error);
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
        updateImageProcessButton();
    }
}

// 결과 표시 (이미지 모드)
function showResultImage(message) {
    const resultSection = document.getElementById('resultSectionImage');
    const resultMessage = document.getElementById('resultMessageImage');
    
    if (resultMessage) {
        resultMessage.textContent = message;
    }
    if (resultSection) {
        resultSection.style.display = 'block';
    }
    
    // 미리보기 및 다운로드 버튼 설정
    setupResultButtonsImage();
}

// 결과 숨기기 (이미지 모드)
function hideResultImage() {
    const resultSection = document.getElementById('resultSectionImage');
    if (resultSection) {
        resultSection.style.display = 'none';
    }
}

// 결과 버튼 설정 (이미지 모드)
function setupResultButtonsImage() {
    const previewBtn = document.getElementById('previewBtnImage');
    const downloadBtn = document.getElementById('downloadBtnImage');
    
    if (previewBtn) {
        previewBtn.onclick = () => {
            if (currentOutputFilename) {
                showPreviewModal(`/api/preview/${currentOutputFilename}`);
            }
        };
    }
    
    if (downloadBtn) {
        downloadBtn.onclick = () => {
            if (currentOutputFilename) {
                window.location.href = `/api/download/${currentOutputFilename}`;
            }
        };
    }
}

