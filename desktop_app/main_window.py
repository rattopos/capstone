# -*- coding: utf-8 -*-
"""
PyQt6 메인 윈도우
지역경제동향 보도자료 생성기 GUI
"""

import sys
import os
from pathlib import Path
from typing import Optional, List, Dict

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QComboBox, QCheckBox, QGroupBox,
    QFileDialog, QProgressBar, QStatusBar, QMessageBox,
    QScrollArea, QFrame, QSplitter, QApplication
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QMimeData
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QFont, QIcon

# WebEngine은 선택적으로 임포트
try:
    from PyQt6.QtWebEngineWidgets import QWebEngineView
    HAS_WEBENGINE = True
except ImportError:
    HAS_WEBENGINE = False
    QWebEngineView = None


class FileDropWidget(QFrame):
    """드래그 앤 드롭 파일 선택 위젯"""
    
    file_dropped = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setup_ui()
    
    def setup_ui(self):
        self.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.setMinimumHeight(80)
        self.setStyleSheet("""
            FileDropWidget {
                background-color: #f5f5f5;
                border: 2px dashed #ccc;
                border-radius: 8px;
            }
            FileDropWidget:hover {
                border-color: #2196F3;
                background-color: #e3f2fd;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.icon_label = QLabel("📁")
        self.icon_label.setFont(QFont("Segoe UI Emoji", 24))
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.text_label = QLabel("엑셀 파일을 여기에 드래그하거나\n클릭하여 선택하세요")
        self.text_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.text_label.setStyleSheet("color: #666;")
        
        self.file_label = QLabel("")
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_label.setStyleSheet("color: #2196F3; font-weight: bold;")
        self.file_label.hide()
        
        layout.addWidget(self.icon_label)
        layout.addWidget(self.text_label)
        layout.addWidget(self.file_label)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].toLocalFile().endswith(('.xlsx', '.xls')):
                event.acceptProposedAction()
                self.setStyleSheet("""
                    FileDropWidget {
                        background-color: #e3f2fd;
                        border: 2px dashed #2196F3;
                        border-radius: 8px;
                    }
                """)
    
    def dragLeaveEvent(self, event):
        self.setStyleSheet("""
            FileDropWidget {
                background-color: #f5f5f5;
                border: 2px dashed #ccc;
                border-radius: 8px;
            }
            FileDropWidget:hover {
                border-color: #2196F3;
                background-color: #e3f2fd;
            }
        """)
    
    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.endswith(('.xlsx', '.xls')):
                self.set_file(file_path)
                self.file_dropped.emit(file_path)
        
        self.dragLeaveEvent(event)
    
    def mousePressEvent(self, event):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "엑셀 파일 선택",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.set_file(file_path)
            self.file_dropped.emit(file_path)
    
    def set_file(self, file_path: str):
        """선택된 파일 표시"""
        filename = os.path.basename(file_path)
        self.icon_label.setText("✅")
        self.text_label.hide()
        self.file_label.setText(filename)
        self.file_label.show()


class SidoCheckboxGroup(QGroupBox):
    """17개 시도 체크박스 그룹"""
    
    def __init__(self, parent=None):
        super().__init__("📋 생성할 시도", parent)
        self.checkboxes: Dict[str, QCheckBox] = {}
        self.setup_ui()
    
    def setup_ui(self):
        layout = QGridLayout(self)
        
        # 전체 선택 체크박스
        self.select_all = QCheckBox("전체 선택")
        self.select_all.setChecked(True)
        self.select_all.stateChanged.connect(self.toggle_all)
        layout.addWidget(self.select_all, 0, 0, 1, 2)
        
        # 17개 시도 체크박스
        sido_list = [
            "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종",
            "경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"
        ]
        
        for i, sido in enumerate(sido_list):
            cb = QCheckBox(sido)
            cb.setChecked(True)
            cb.stateChanged.connect(self.update_select_all)
            self.checkboxes[sido] = cb
            
            row = (i // 2) + 1
            col = i % 2
            layout.addWidget(cb, row, col)
    
    def toggle_all(self, state):
        """전체 선택/해제"""
        checked = state == Qt.CheckState.Checked.value
        for cb in self.checkboxes.values():
            cb.blockSignals(True)
            cb.setChecked(checked)
            cb.blockSignals(False)
    
    def update_select_all(self):
        """개별 체크박스 변경 시 전체 선택 상태 업데이트"""
        all_checked = all(cb.isChecked() for cb in self.checkboxes.values())
        self.select_all.blockSignals(True)
        self.select_all.setChecked(all_checked)
        self.select_all.blockSignals(False)
    
    def get_selected(self) -> List[str]:
        """선택된 시도 목록 반환"""
        return [name for name, cb in self.checkboxes.items() if cb.isChecked()]


class GeneratorThread(QThread):
    """HWPX 생성 작업 스레드"""
    
    progress = pyqtSignal(int, str)  # (진행률, 메시지)
    finished = pyqtSignal(bool, str)  # (성공 여부, 결과 메시지)
    
    def __init__(self, excel_path: str, output_path: str, 
                 year: int, quarter: int, selected_sido: List[str]):
        super().__init__()
        self.excel_path = excel_path
        self.output_path = output_path
        self.year = year
        self.quarter = quarter
        self.selected_sido = selected_sido
    
    def run(self):
        try:
            self.progress.emit(10, "데이터 추출 중...")
            
            # 데이터 추출 (기존 로직 활용)
            all_data = self.extract_data()
            
            self.progress.emit(40, "HWPX 템플릿 로드 중...")
            
            # HWPX 생성
            from desktop_app.core.hwpx_injector import HWPXDataInjector
            injector = HWPXDataInjector()
            
            self.progress.emit(60, "시도별 섹션 생성 중...")
            
            success = injector.inject(all_data, self.output_path, self.selected_sido)
            
            self.progress.emit(100, "완료!")
            self.finished.emit(True, f"파일이 생성되었습니다:\n{self.output_path}")
            
        except Exception as e:
            self.finished.emit(False, f"오류 발생: {str(e)}")
    
    def extract_data(self) -> Dict[str, Dict]:
        """엑셀에서 시도별 데이터 추출"""
        # TODO: 기존 RawDataExtractor 연결
        # 현재는 더미 데이터 반환
        dummy_data = {}
        for sido in self.selected_sido:
            dummy_data[sido] = {
                "DATA_23_2Q4_manufacturing": -5.5,
                "DATA_23_2Q4_service": 3.2,
                "DATA_23_2Q4_retail": 1.5,
                "DATA_23_2Q4_construction": 10.2,
                "DATA_23_2Q4_export": 5.5,
                "DATA_23_2Q4_import": -2.3,
                "DATA_23_2Q4_price": 2.8,
                "DATA_23_2Q4_employment": 0.5,
                "DATA_23_2Q4_migration": -3.2,
            }
        return dummy_data


class MainWindow(QMainWindow):
    """메인 윈도우"""
    
    def __init__(self):
        super().__init__()
        self.excel_path: Optional[str] = None
        self.generator_thread: Optional[GeneratorThread] = None
        self.setup_ui()
    
    def setup_ui(self):
        """UI 초기화"""
        self.setWindowTitle("지역경제동향 보도자료 생성기 v1.0")
        self.setMinimumSize(900, 700)
        
        # 중앙 위젯
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 메인 레이아웃 (스플리터)
        main_layout = QHBoxLayout(central_widget)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)
        
        # 좌측 패널
        left_panel = self.create_left_panel()
        splitter.addWidget(left_panel)
        
        # 우측 패널 (미리보기)
        right_panel = self.create_right_panel()
        splitter.addWidget(right_panel)
        
        # 스플리터 비율 설정
        splitter.setSizes([300, 600])
        
        # 상태바
        self.statusBar().showMessage("준비됨")
        
        # 스타일시트
        self.setStyleSheet("""
            QMainWindow {
                background-color: #fafafa;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ddd;
                border-radius: 6px;
                margin-top: 12px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #ccc;
            }
            QPushButton#generateBtn {
                background-color: #4CAF50;
                font-size: 14px;
            }
            QPushButton#generateBtn:hover {
                background-color: #388E3C;
            }
        """)
    
    def create_left_panel(self) -> QWidget:
        """좌측 패널 생성"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setSpacing(15)
        
        # 파일 선택
        file_group = QGroupBox("📁 입력 파일")
        file_layout = QVBoxLayout(file_group)
        self.file_drop = FileDropWidget()
        self.file_drop.file_dropped.connect(self.on_file_selected)
        file_layout.addWidget(self.file_drop)
        layout.addWidget(file_group)
        
        # 기준 설정
        settings_group = QGroupBox("📅 기준 설정")
        settings_layout = QGridLayout(settings_group)
        
        settings_layout.addWidget(QLabel("연도:"), 0, 0)
        self.year_combo = QComboBox()
        self.year_combo.addItems([str(y) for y in range(2020, 2030)])
        self.year_combo.setCurrentText("2025")
        settings_layout.addWidget(self.year_combo, 0, 1)
        
        settings_layout.addWidget(QLabel("분기:"), 1, 0)
        self.quarter_combo = QComboBox()
        self.quarter_combo.addItems(["1", "2", "3", "4"])
        self.quarter_combo.setCurrentText("3")
        settings_layout.addWidget(self.quarter_combo, 1, 1)
        
        layout.addWidget(settings_group)
        
        # 시도 선택
        self.sido_group = SidoCheckboxGroup()
        scroll = QScrollArea()
        scroll.setWidget(self.sido_group)
        scroll.setWidgetResizable(True)
        scroll.setMaximumHeight(250)
        layout.addWidget(scroll)
        
        # 진행률
        progress_group = QGroupBox("📊 진행 상태")
        progress_layout = QVBoxLayout(progress_group)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        self.progress_label = QLabel("대기 중...")
        self.progress_label.setStyleSheet("color: #666;")
        progress_layout.addWidget(self.progress_label)
        
        layout.addWidget(progress_group)
        
        # 버튼
        btn_layout = QHBoxLayout()
        
        self.preview_btn = QPushButton("🔍 미리보기")
        self.preview_btn.clicked.connect(self.on_preview)
        self.preview_btn.setEnabled(False)
        btn_layout.addWidget(self.preview_btn)
        
        self.generate_btn = QPushButton("📥 HWPX 생성")
        self.generate_btn.setObjectName("generateBtn")
        self.generate_btn.clicked.connect(self.on_generate)
        self.generate_btn.setEnabled(False)
        btn_layout.addWidget(self.generate_btn)
        
        layout.addLayout(btn_layout)
        
        # 여백
        layout.addStretch()
        
        return panel
    
    def create_right_panel(self) -> QWidget:
        """우측 패널 생성 (미리보기)"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        
        # 헤더
        header = QHBoxLayout()
        header.addWidget(QLabel("미리보기"))
        
        self.sido_preview_combo = QComboBox()
        self.sido_preview_combo.addItems([
            "서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종",
            "경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"
        ])
        self.sido_preview_combo.currentTextChanged.connect(self.update_preview)
        header.addWidget(self.sido_preview_combo)
        header.addStretch()
        
        layout.addLayout(header)
        
        # 미리보기 영역
        if HAS_WEBENGINE:
            self.preview_view = QWebEngineView()
            self.preview_view.setHtml(self.get_placeholder_html())
        else:
            self.preview_view = QLabel("미리보기를 사용하려면 PyQt6-WebEngine을 설치하세요.")
            self.preview_view.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.preview_view.setStyleSheet("""
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 20px;
                color: #666;
            """)
        
        layout.addWidget(self.preview_view)
        
        return panel
    
    def get_placeholder_html(self) -> str:
        """기본 미리보기 HTML"""
        return """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {
                    font-family: 'Malgun Gothic', sans-serif;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    height: 100vh;
                    margin: 0;
                    background-color: #f5f5f5;
                    color: #999;
                }
                .placeholder {
                    text-align: center;
                }
                .icon {
                    font-size: 48px;
                    margin-bottom: 20px;
                }
            </style>
        </head>
        <body>
            <div class="placeholder">
                <div class="icon">📄</div>
                <p>엑셀 파일을 선택하면<br>미리보기가 표시됩니다</p>
            </div>
        </body>
        </html>
        """
    
    def on_file_selected(self, file_path: str):
        """파일 선택 시"""
        self.excel_path = file_path
        self.preview_btn.setEnabled(True)
        self.generate_btn.setEnabled(True)
        self.statusBar().showMessage(f"파일 로드됨: {os.path.basename(file_path)}")
        
        # 자동으로 연도/분기 감지 시도
        self.detect_year_quarter(file_path)
    
    def detect_year_quarter(self, file_path: str):
        """파일명에서 연도/분기 감지"""
        import re
        filename = os.path.basename(file_path)
        
        # 패턴: 2025년_3분기 또는 2025_3 등
        match = re.search(r'(\d{4})[년_]?\s*(\d)[분기/]?', filename)
        if match:
            year, quarter = match.groups()
            self.year_combo.setCurrentText(year)
            self.quarter_combo.setCurrentText(quarter)
    
    def on_preview(self):
        """미리보기 버튼 클릭"""
        if not self.excel_path:
            return
        
        self.update_preview(self.sido_preview_combo.currentText())
    
    def update_preview(self, sido_name: str):
        """미리보기 업데이트"""
        if not self.excel_path or not HAS_WEBENGINE:
            return
        
        # TODO: 실제 데이터로 HTML 생성
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {{
                    font-family: '휴먼명조', 'Malgun Gothic', serif;
                    padding: 20px;
                    line-height: 1.6;
                }}
                h2 {{
                    text-align: center;
                    color: #333;
                    border-bottom: 2px solid #2196F3;
                    padding-bottom: 10px;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 20px;
                }}
                th, td {{
                    border: 1px solid #ddd;
                    padding: 8px;
                    text-align: center;
                }}
                th {{
                    background-color: #f5f5f5;
                }}
            </style>
        </head>
        <body>
            <h2>《 {sido_name} 주요지표 》</h2>
            <p style="text-align: right; color: #666;">[전년동분기대비, %]</p>
            <table>
                <tr>
                    <th></th>
                    <th>광공업<br>생산</th>
                    <th>서비스업<br>생산</th>
                    <th>소매<br>판매</th>
                    <th>건설<br>수주</th>
                    <th>수출</th>
                    <th>수입</th>
                    <th>소비자<br>물가</th>
                    <th>고용률<br>(%p)</th>
                </tr>
                <tr>
                    <td>'23.2/4</td>
                    <td>-5.5</td>
                    <td>3.2</td>
                    <td>1.5</td>
                    <td>10.2</td>
                    <td>5.5</td>
                    <td>-2.3</td>
                    <td>2.8</td>
                    <td>0.5</td>
                </tr>
                <tr>
                    <td>'24.2/4</td>
                    <td>...</td>
                    <td>...</td>
                    <td>...</td>
                    <td>...</td>
                    <td>...</td>
                    <td>...</td>
                    <td>...</td>
                    <td>...</td>
                </tr>
            </table>
            <p style="color: #999; font-size: 12px; margin-top: 20px;">
                * 실제 데이터는 엑셀 파일에서 추출됩니다.
            </p>
        </body>
        </html>
        """
        
        self.preview_view.setHtml(html)
    
    def on_generate(self):
        """HWPX 생성 버튼 클릭"""
        if not self.excel_path:
            QMessageBox.warning(self, "경고", "엑셀 파일을 먼저 선택하세요.")
            return
        
        selected_sido = self.sido_group.get_selected()
        if not selected_sido:
            QMessageBox.warning(self, "경고", "최소 1개 이상의 시도를 선택하세요.")
            return
        
        # 저장 경로 선택
        year = self.year_combo.currentText()
        quarter = self.quarter_combo.currentText()
        default_name = f"지역경제동향_{year}년_{quarter}분기.hwpx"
        
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "HWPX 파일 저장",
            default_name,
            "HWPX Files (*.hwpx)"
        )
        
        if not output_path:
            return
        
        # 생성 스레드 시작
        self.generate_btn.setEnabled(False)
        self.preview_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        
        self.generator_thread = GeneratorThread(
            self.excel_path,
            output_path,
            int(year),
            int(quarter),
            selected_sido
        )
        self.generator_thread.progress.connect(self.on_progress)
        self.generator_thread.finished.connect(self.on_generation_finished)
        self.generator_thread.start()
    
    def on_progress(self, value: int, message: str):
        """진행률 업데이트"""
        self.progress_bar.setValue(value)
        self.progress_label.setText(message)
        self.statusBar().showMessage(message)
    
    def on_generation_finished(self, success: bool, message: str):
        """생성 완료"""
        self.generate_btn.setEnabled(True)
        self.preview_btn.setEnabled(True)
        
        if success:
            QMessageBox.information(self, "완료", message)
            self.progress_label.setText("생성 완료!")
        else:
            QMessageBox.critical(self, "오류", message)
            self.progress_label.setText("오류 발생")
            self.progress_bar.setValue(0)


def main():
    """앱 실행"""
    app = QApplication(sys.argv)
    
    # 앱 정보 설정
    app.setApplicationName("지역경제동향 보도자료 생성기")
    app.setOrganizationName("국가데이터처")
    app.setApplicationVersion("1.0.0")
    
    # 메인 윈도우 생성 및 표시
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
