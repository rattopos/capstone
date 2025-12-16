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
    loadTemplates();
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
    fileInput.addEventListener('change', (e) => {
        handleFileSelect(e.target.files[0]);
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
            handleFileSelect(file);
        }
    });
}

// 파일 선택 처리
function handleFileSelect(file) {
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
    updateProcessButton();
}

// 템플릿 목록 로드
async function loadTemplates() {
    try {
        const response = await fetch('/api/templates');
        const data = await response.json();
        
        const select = document.getElementById('templateSelect');
        select.innerHTML = '';

        if (data.templates && data.templates.length > 0) {
            data.templates.forEach(template => {
                const option = document.createElement('option');
                option.value = template.name;
                option.textContent = template.name;
                if (template.name === 'mining_manufacturing_production.html') {
                    option.selected = true;
                }
                select.appendChild(option);
            });
        } else {
            const option = document.createElement('option');
            option.value = '';
            option.textContent = '템플릿을 찾을 수 없습니다';
            select.appendChild(option);
        }
    } catch (error) {
        console.error('템플릿 로드 오류:', error);
        showError('템플릿을 불러오는 중 오류가 발생했습니다.');
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

// 템플릿 선택 변경 이벤트
document.getElementById('templateSelect').addEventListener('change', updateProcessButton);

// 보도자료 생성 처리
async function handleProcess() {
    if (!selectedExcelFile) {
        showError('엑셀 파일을 선택해주세요.');
        return;
    }

    const templateSelect = document.getElementById('templateSelect');
    if (!templateSelect.value) {
        showError('템플릿을 선택해주세요.');
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
        formData.append('template', templateSelect.value);

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

