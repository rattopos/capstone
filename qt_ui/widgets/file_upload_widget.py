# -*- coding: utf-8 -*-
"""
파일 업로드 위젯
"""

from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent


class FileUploadWidget(QWidget):
    """파일 업로드 위젯 (드래그 앤 드롭 지원)"""
    
    file_uploaded = pyqtSignal(str)  # 파일 경로 전달
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.setAcceptDrops(True)
    
    def setup_ui(self):
        """UI 구성"""
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 제목
        title_label = QLabel("기초자료 수집표 업로드")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 16pt;
                font-weight: bold;
                color: #333;
            }
        """)
        layout.addWidget(title_label)
        
        # 업로드 영역
        upload_label = QLabel("엑셀 파일을 드래그 앤 드롭하거나 버튼을 클릭하세요")
        upload_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        upload_label.setMinimumHeight(150)
        upload_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #999;
                border-radius: 10px;
                background-color: #f5f5f5;
                padding: 20px;
                font-size: 11pt;
                color: #666;
            }
        """)
        upload_label.setAcceptDrops(True)
        self.upload_label = upload_label
        layout.addWidget(upload_label)
        
        # 파일 선택 버튼
        select_btn = QPushButton("📁 파일 선택")
        select_btn.setMinimumHeight(40)
        select_btn.setStyleSheet("""
            QPushButton {
                background-color: #0066cc;
                color: white;
                font-size: 11pt;
                font-weight: bold;
                border-radius: 5px;
                padding: 8px;
            }
            QPushButton:hover {
                background-color: #0052a3;
            }
            QPushButton:pressed {
                background-color: #003d7a;
            }
        """)
        select_btn.clicked.connect(self.select_file)
        layout.addWidget(select_btn)
        
        # 상태 표시
        self.status_label = QLabel("파일을 선택해주세요")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("""
            QLabel {
                font-size: 10pt;
                color: #666;
                padding: 10px;
            }
        """)
        layout.addWidget(self.status_label)
        
        self.setLayout(layout)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """드래그 진입 이벤트"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.upload_label.setStyleSheet("""
                QLabel {
                    border: 2px dashed #0066cc;
                    border-radius: 10px;
                    background-color: #e6f2ff;
                    padding: 20px;
                    font-size: 11pt;
                    color: #0066cc;
                }
            """)
    
    def dragLeaveEvent(self, event):
        """드래그 떠남 이벤트"""
        self.upload_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #999;
                border-radius: 10px;
                background-color: #f5f5f5;
                padding: 20px;
                font-size: 11pt;
                color: #666;
            }
        """)
    
    def dropEvent(self, event: QDropEvent):
        """드롭 이벤트"""
        self.upload_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #999;
                border-radius: 10px;
                background-color: #f5f5f5;
                padding: 20px;
                font-size: 11pt;
                color: #666;
            }
        """)
        
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls:
                filepath = urls[0].toLocalFile()
                if self.validate_file(filepath):
                    self.handle_file(filepath)
                else:
                    QMessageBox.warning(
                        self,
                        "파일 오류",
                        "엑셀 파일(.xlsx, .xls)만 업로드 가능합니다."
                    )
            event.acceptProposedAction()
    
    def select_file(self):
        """파일 선택 다이얼로그"""
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            "기초자료 수집표 선택",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if filepath:
            if self.validate_file(filepath):
                self.handle_file(filepath)
            else:
                QMessageBox.warning(
                    self,
                    "파일 오류",
                    "엑셀 파일(.xlsx, .xls)만 업로드 가능합니다."
                )
    
    def validate_file(self, filepath: str) -> bool:
        """파일 검증"""
        path = Path(filepath)
        return path.exists() and path.suffix.lower() in ['.xlsx', '.xls']
    
    def handle_file(self, filepath: str):
        """파일 처리"""
        self.status_label.setText(f"처리 중: {Path(filepath).name}")
        self.status_label.setStyleSheet("""
            QLabel {
                font-size: 10pt;
                color: #0066cc;
                padding: 10px;
            }
        """)
        self.file_uploaded.emit(filepath)
    
    def set_status(self, message: str, success: bool = True):
        """상태 메시지 설정"""
        self.status_label.setText(message)
        if success:
            self.status_label.setStyleSheet("""
                QLabel {
                    font-size: 10pt;
                    color: #28a745;
                    padding: 10px;
                }
            """)
        else:
            self.status_label.setStyleSheet("""
                QLabel {
                    font-size: 10pt;
                    color: #dc3545;
                    padding: 10px;
                }
            """)
