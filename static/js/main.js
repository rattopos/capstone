// 전역 변수
let selectedExcelFile = null;
let currentOutputFilename = null;
let selectedTemplate = null;
let templatesList = [];

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

// 진행률 표시
function showProgress(percentage) {
    const container = document.getElementById('progressContainer');
    const bar = document.getElementById('progressBar');
    
    if (container && bar) {
        container.classList.add('active');
        bar.style.width = `${Math.min(100, Math.max(0, percentage))}%`;
    }
}

function hideProgress() {
    const container = document.getElementById('progressContainer');
    if (container) {
        container.classList.remove('active');
        setTimeout(() => {
            const bar = document.getElementById('progressBar');
            if (bar) bar.style.width = '0%';
        }, 300);
    }
}

// DOM 로드 완료 시 초기화
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

// 앱 초기화
function initializeApp() {
    loadTemplates();
    setupFileUpload();
    setupTemplateSelect();
    setupProcessButton();
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
    
    if (selectedExcelFile && templateSelect.value) {
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
    
    // 필요한 시트 확인
    const validation = await validateRequiredSheets(selectedExcelFile);
    if (!validation.valid) {
        showError(validation.error);
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
    showProgress(10);
    showToast('파일을 업로드하고 처리 중입니다...', 'info', 2000);

    try {
        // FormData 생성
        const formData = new FormData();
        formData.append('excel_file', selectedExcelFile);
        formData.append('template_name', templateName);
        formData.append('year', year);
        formData.append('quarter', quarter);

        // 진행률 업데이트
        showProgress(30);
        
        // API 호출
        const response = await fetch('/api/process', {
            method: 'POST',
            body: formData
        });
        
        showProgress(60);

        const data = await response.json();
        showProgress(90);

        if (response.ok && data.success) {
            currentOutputFilename = data.output_filename;
            showProgress(100);
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
}

// 결과 버튼 설정
function setupResultButtons() {
    const previewBtn = document.getElementById('previewBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    
    previewBtn.onclick = () => {
        if (currentOutputFilename) {
            showToast('미리보기 창을 엽니다...', 'info', 2000);
            window.open(`/api/preview/${currentOutputFilename}`, '_blank');
        }
    };
    
    downloadBtn.onclick = () => {
        if (currentOutputFilename) {
            showToast('파일을 다운로드합니다...', 'info', 2000);
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

// 템플릿 생성 관련 함수들
let selectedImageFile = null;
let selectedExcelFileForTemplate = null;

// 이미지 업로드 설정
function setupImageUpload() {
    const uploadArea = document.getElementById('imageUploadArea');
    const fileInput = document.getElementById('imageFile');
    const fileInfo = document.getElementById('imageFileInfo');

    if (!uploadArea || !fileInput) return;

    // 클릭 이벤트
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // 파일 선택 이벤트
    fileInput.addEventListener('change', (e) => {
        handleImageSelect(e.target.files[0]);
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
            handleImageSelect(file);
        }
    });
}

// 이미지 파일 선택 처리
function handleImageSelect(file) {
    if (!file) return;

    // 파일 크기 검증
    const maxFileSize = 100 * 1024 * 1024;
    if (file.size > maxFileSize) {
        showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        return;
    }

    // 이미지 파일 형식 검증
    const allowedTypes = ['image/png', 'image/jpeg', 'image/jpg', 'image/gif', 'image/bmp', 'image/webp'];
    if (!allowedTypes.includes(file.type)) {
        showError('지원하지 않는 이미지 형식입니다. PNG, JPG, JPEG, GIF, BMP, WEBP 파일만 업로드 가능합니다.');
        return;
    }

    selectedImageFile = file;
    displayImageFileInfo(file);
    updateCreateTemplateButton();
}

// 이미지 파일 정보 표시
function displayImageFileInfo(file) {
    const fileInfo = document.getElementById('imageFileInfo');
    const fileName = fileInfo.querySelector('.file-name');
    
    fileName.textContent = file.name;
    fileInfo.style.display = 'flex';
}

// 이미지 파일 제거
function removeImageFile() {
    selectedImageFile = null;
    document.getElementById('imageFile').value = '';
    document.getElementById('imageFileInfo').style.display = 'none';
    document.getElementById('templateCreateResult').style.display = 'none';
    updateCreateTemplateButton();
}

// 템플릿 생성 버튼 상태 업데이트
function updateCreateTemplateButton() {
    const createBtn = document.getElementById('createTemplateBtn');
    if (createBtn) {
        createBtn.disabled = !selectedImageFile;
    }
}

// 템플릿 생성 처리
async function handleCreateTemplate() {
    if (!selectedImageFile) {
        showError('이미지 파일을 선택해주세요.');
        return;
    }

    const templateName = document.getElementById('templateName').value.trim();
    const sheetName = document.getElementById('sheetName').value.trim() || '시트1';

    // UI 업데이트
    const createBtn = document.getElementById('createTemplateBtn');
    const btnText = createBtn.querySelector('.btn-text');
    const btnLoader = createBtn.querySelector('.btn-loader');
    
    createBtn.disabled = true;
    btnText.style.display = 'none';
    btnLoader.style.display = 'inline-block';
    
    hideError();
    document.getElementById('templateCreateResult').style.display = 'none';
    showProgress(10);
    showToast('템플릿을 생성하는 중입니다...', 'info', 2000);

    try {
        // FormData 생성
        const formData = new FormData();
        formData.append('image_file', selectedImageFile);
        
        // 엑셀 파일이 있으면 추가
        if (selectedExcelFileForTemplate) {
            formData.append('excel_file', selectedExcelFileForTemplate);
        }
        
        if (templateName) {
            formData.append('template_name', templateName);
        }
        formData.append('sheet_name', sheetName);

        // 진행률 업데이트
        showProgress(30);
        
        // API 호출
        const response = await fetch('/api/create-template', {
            method: 'POST',
            body: formData
        });
        
        showProgress(60);

        const data = await response.json();
        showProgress(90);

        if (response.ok && data.success) {
            showProgress(100);
            setTimeout(() => {
                hideProgress();
                const resultDiv = document.getElementById('templateCreateResult');
                const messageDiv = document.getElementById('templateCreateMessage');
                messageDiv.innerHTML = `
                    <strong>✅ ${data.message}</strong><br>
                    <small>템플릿 파일: ${data.template_name}</small>
                `;
                resultDiv.style.display = 'block';
                resultDiv.style.backgroundColor = '#e8f5e9';
                showSuccess('템플릿이 성공적으로 생성되었습니다!');
                
                // 템플릿 목록 새로고침
                setTimeout(() => {
                    loadTemplates();
                }, 1000);
            }, 500);
        } else {
            hideProgress();
            showError(data.error || '템플릿 생성 중 오류가 발생했습니다.');
        }
    } catch (error) {
        console.error('템플릿 생성 오류:', error);
        hideProgress();
        showError('서버와 통신하는 중 오류가 발생했습니다. 네트워크 연결을 확인해주세요.');
    } finally {
        // UI 복원
        createBtn.disabled = false;
        btnText.style.display = 'inline';
        btnLoader.style.display = 'none';
        updateCreateTemplateButton();
    }
}

// 엑셀 파일 업로드 설정 (템플릿 생성용)
function setupExcelUploadForTemplate() {
    const uploadArea = document.getElementById('excelUploadAreaForTemplate');
    const fileInput = document.getElementById('excelFileForTemplate');
    const fileInfo = document.getElementById('excelFileInfoForTemplate');

    if (!uploadArea || !fileInput) return;

    // 클릭 이벤트
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // 파일 선택 이벤트
    fileInput.addEventListener('change', (e) => {
        handleExcelSelectForTemplate(e.target.files[0]);
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
            handleExcelSelectForTemplate(file);
        }
    });
}

// 엑셀 파일 선택 처리 (템플릿 생성용)
async function handleExcelSelectForTemplate(file) {
    if (!file) return;

    // 파일 크기 검증
    const maxFileSize = 100 * 1024 * 1024;
    if (file.size > maxFileSize) {
        showError('파일 크기가 너무 큽니다. 최대 100MB까지 업로드 가능합니다.');
        return;
    }

    // 엑셀 파일 형식 검증
    const allowedExtensions = ['.xlsx', '.xls'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();

    if (!allowedExtensions.includes(fileExtension)) {
        showError('지원하지 않는 파일 형식입니다. .xlsx 또는 .xls 파일만 업로드 가능합니다.');
        return;
    }

    selectedExcelFileForTemplate = file;
    displayExcelFileInfoForTemplate(file);
    
    // 엑셀 파일의 시트 목록 가져오기
    await loadSheetNamesForTemplate(file);
}

// 엑셀 파일 정보 표시 (템플릿 생성용)
function displayExcelFileInfoForTemplate(file) {
    const fileInfo = document.getElementById('excelFileInfoForTemplate');
    const fileName = fileInfo.querySelector('.file-name');
    
    fileName.textContent = file.name;
    fileInfo.style.display = 'flex';
}

// 엑셀 파일 제거 (템플릿 생성용)
function removeExcelFileForTemplate() {
    selectedExcelFileForTemplate = null;
    document.getElementById('excelFileForTemplate').value = '';
    document.getElementById('excelFileInfoForTemplate').style.display = 'none';
    document.getElementById('sheetName').value = '';
    updateCreateTemplateButton();
}

// 시트 목록 로드 (템플릿 생성용)
async function loadSheetNamesForTemplate(file) {
    const sheetNameInput = document.getElementById('sheetName');
    
    try {
        const formData = new FormData();
        formData.append('excel_file', file);
        
        const response = await fetch('/api/validate', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (response.ok && data.valid && data.sheet_names && data.sheet_names.length > 0) {
            // 첫 번째 시트를 기본값으로 설정
            sheetNameInput.value = data.sheet_names[0];
            sheetNameInput.placeholder = `사용 가능한 시트: ${data.sheet_names.join(', ')}`;
        }
    } catch (error) {
        console.error('시트 목록 로드 오류:', error);
    }
}

// 초기화 시 이미지 업로드 설정
document.addEventListener('DOMContentLoaded', function() {
    setupImageUpload();
    setupExcelUploadForTemplate();
    
    // 템플릿 생성 버튼 이벤트
    const createBtn = document.getElementById('createTemplateBtn');
    if (createBtn) {
        createBtn.addEventListener('click', handleCreateTemplate);
    }
});

