# -*- coding: utf-8 -*-
"""
Qt6 메인 윈도우
"""

import sys
from pathlib import Path
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QProgressBar, QMessageBox, QSplitter, QFileDialog
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QClipboard, QMimeData

from qt_ui.controllers import AppController
from qt_ui.widgets.file_upload_widget import FileUploadWidget
from qt_ui.widgets.report_list_widget import ReportListWidget


class ReportGenerationThread(QThread):
    """보도자료 생성 스레드 (백그라운드 처리)"""
    
    progress = pyqtSignal(str)  # 진행 상태 메시지
    finished = pyqtSignal(list)  # 완료 시 페이지 리스트 전달
    error = pyqtSignal(str)  # 오류 메시지
    
    def __init__(self, controller: AppController, selected_reports: list):
        super().__init__()
        self.controller = controller
        self.selected_reports = selected_reports
    
    def run(self):
        """스레드 실행"""
        try:
            self.progress.emit("보도자료 생성 중...")
            
            pages = []
            total = len(self.selected_reports)
            
            for idx, report_config in enumerate(self.selected_reports, 1):
                self.progress.emit(f"보도자료 생성 중... ({idx}/{total}): {report_config['name']}")
                
                try:
                    html_content, error, missing_fields = self.controller._generate_report(report_config)
                    if html_content:
                        pages.append({
                            'id': report_config['id'],
                            'title': report_config['name'],
                            'category': report_config.get('category', 'summary'),
                            'html': html_content
                        })
                except Exception as e:
                    print(f"[오류] {report_config['name']} 생성 실패: {e}")
                    continue
            
            self.progress.emit("보도자료 생성 완료")
            self.finished.emit(pages)
            
        except Exception as e:
            self.error.emit(f"보도자료 생성 오류: {str(e)}")


class MainWindow(QMainWindow):
    """메인 윈도우"""
    
    def __init__(self):
        super().__init__()
        self.controller = AppController()
        self.generation_thread = None
        self.generated_pages = []
        
        self.setup_ui()
        self.setup_connections()
    
    def setup_ui(self):
        """UI 구성"""
        self.setWindowTitle("지역경제동향 보도자료 생성기")
        self.setMinimumSize(900, 700)
        
        # 중앙 위젯
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 메인 레이아웃
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(15, 15, 15, 15)
        
        # 제목
        title_label = QLabel("지역경제동향 보도자료 생성 시스템")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 18pt;
                font-weight: bold;
                color: #0066cc;
                padding: 10px;
            }
        """)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 스플리터 (파일 업로드 | 보도자료 목록)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # 왼쪽: 파일 업로드
        self.file_upload_widget = FileUploadWidget()
        splitter.addWidget(self.file_upload_widget)
        
        # 오른쪽: 보도자료 목록
        self.report_list_widget = ReportListWidget()
        splitter.addWidget(self.report_list_widget)
        
        # 스플리터 비율 설정 (40:60)
        splitter.setSizes([400, 500])
        main_layout.addWidget(splitter)
        
        # 버튼 영역
        button_layout = QHBoxLayout()
        
        # 전체 선택/해제 버튼
        self.select_all_btn = QPushButton("전체 선택")
        self.select_all_btn.setMinimumHeight(35)
        self.select_all_btn.clicked.connect(self.report_list_widget.select_all)
        button_layout.addWidget(self.select_all_btn)
        
        self.deselect_all_btn = QPushButton("전체 해제")
        self.deselect_all_btn.setMinimumHeight(35)
        self.deselect_all_btn.clicked.connect(self.report_list_widget.deselect_all)
        button_layout.addWidget(self.deselect_all_btn)
        
        button_layout.addStretch()
        
        # 생성 버튼
        self.generate_btn = QPushButton("HTML 생성 및 저장")
        self.generate_btn.setMinimumHeight(40)
        self.generate_btn.setMinimumWidth(200)
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                font-size: 12pt;
                font-weight: bold;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #218838;
            }
            QPushButton:pressed {
                background-color: #1e7e34;
            }
            QPushButton:disabled {
                background-color: #ccc;
                color: #666;
            }
        """)
        self.generate_btn.clicked.connect(self.generate_html)
        self.generate_btn.setEnabled(False)
        button_layout.addWidget(self.generate_btn)
        
        main_layout.addLayout(button_layout)
        
        # 진행 상태 표시
        self.progress_label = QLabel("파일을 업로드해주세요")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_label.setStyleSheet("""
            QLabel {
                font-size: 10pt;
                color: #666;
                padding: 5px;
            }
        """)
        main_layout.addWidget(self.progress_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(0)  # 무한 진행 표시
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        central_widget.setLayout(main_layout)
    
    def setup_connections(self):
        """시그널-슬롯 연결"""
        self.file_upload_widget.file_uploaded.connect(self.on_file_uploaded)
    
    def on_file_uploaded(self, filepath: str):
        """파일 업로드 처리"""
        success, message = self.controller.handle_file_upload(filepath)
        
        if success:
            self.file_upload_widget.set_status(message, True)
            self.progress_label.setText(f"파일 업로드 완료: {self.controller.year}년 {self.controller.quarter}분기")
            self.generate_btn.setEnabled(True)
        else:
            self.file_upload_widget.set_status(message, False)
            QMessageBox.warning(self, "업로드 실패", message)
            self.generate_btn.setEnabled(False)
    
    def generate_html(self):
        """HTML 생성 및 저장"""
        if not self.controller.raw_excel_path:
            QMessageBox.warning(self, "오류", "파일을 먼저 업로드해주세요.")
            return
        
        selected_reports = self.report_list_widget.get_selected_reports()
        if not selected_reports:
            QMessageBox.warning(self, "오류", "생성할 보도자료를 선택해주세요.")
            return
        
        # 생성 버튼 비활성화
        self.generate_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_label.setText("보도자료 생성 중...")
        
        # 백그라운드 스레드로 생성
        self.generation_thread = ReportGenerationThread(self.controller, selected_reports)
        self.generation_thread.progress.connect(self.on_generation_progress)
        self.generation_thread.finished.connect(self.on_generation_finished)
        self.generation_thread.error.connect(self.on_generation_error)
        self.generation_thread.start()
    
    def on_generation_progress(self, message: str):
        """생성 진행 상태 업데이트"""
        self.progress_label.setText(message)
    
    def on_generation_finished(self, pages: list):
        """생성 완료"""
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)
        
        if not pages:
            QMessageBox.warning(self, "오류", "생성된 보도자료가 없습니다.")
            return
        
        self.generated_pages = pages
        
        # HTML 생성
        try:
            html_content = self.controller.generate_hwp_html(pages)
            
            # 저장 위치 선택
            default_filename = f"지역경제동향_{self.controller.year}년_{self.controller.quarter}분기_한글불러오기용.html"
            filepath, _ = QFileDialog.getSaveFileName(
                self,
                "HTML 파일 저장",
                default_filename,
                "HTML Files (*.html);;All Files (*)"
            )
            
            if not filepath:
                self.progress_label.setText("저장이 취소되었습니다.")
                return
            
            # 파일 저장
            try:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(html_content)
            except Exception as e:
                QMessageBox.critical(self, "저장 실패", f"파일 저장 중 오류가 발생했습니다:\n{str(e)}")
                return
            
            # 클립보드에 복사
            clipboard = QClipboard()
            mime_data = QMimeData()
            mime_data.setHtml(html_content)
            mime_data.setText(html_content)  # 텍스트 폴백
            clipboard.setMimeData(mime_data)
            
            # 완료 메시지
            QMessageBox.information(
                self,
                "완료",
                f"HTML 파일이 저장되었습니다.\n\n"
                f"파일 위치: {filepath}\n\n"
                f"클립보드에도 복사되었습니다.\n"
                f"한글(HWP)에서 Ctrl+V로 붙여넣으세요."
            )
            
            self.progress_label.setText(f"생성 완료: {len(pages)}개 페이지")
            
        except Exception as e:
            QMessageBox.critical(self, "오류", f"HTML 생성 중 오류가 발생했습니다:\n{str(e)}")
            self.progress_label.setText("오류 발생")
    
    def on_generation_error(self, error_message: str):
        """생성 오류"""
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)
        QMessageBox.critical(self, "오류", error_message)
        self.progress_label.setText("오류 발생")
