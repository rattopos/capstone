// 전역 변수
let selectedExcelFile = null;
let currentOutputFilename = null;

// DOM 로드 완료 시 초기화
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

// 앱 초기화
function initializeApp() {
    setupFileUpload();
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
            window.open(`/api/preview/${currentOutputFilename}`, '_blank');
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

        // API 호출
        const response = await fetch('/api/create-template', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();

        if (response.ok && data.success) {
            const resultDiv = document.getElementById('templateCreateResult');
            const messageDiv = document.getElementById('templateCreateMessage');
            messageDiv.innerHTML = `
                <strong>✅ ${data.message}</strong><br>
                <small>템플릿 파일: ${data.template_name}</small>
            `;
            resultDiv.style.display = 'block';
            resultDiv.style.backgroundColor = '#e8f5e9';
        } else {
            showError(data.error || '템플릿 생성 중 오류가 발생했습니다.');
        }
    } catch (error) {
        console.error('템플릿 생성 오류:', error);
        showError('서버와 통신하는 중 오류가 발생했습니다.');
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

