# -*- coding: utf-8 -*-
"""
보도자료 목록 위젯
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QTreeWidget, QTreeWidgetItem, QHeaderView
)
from PyQt6.QtCore import Qt

from config.reports import REPORT_ORDER, SUMMARY_REPORTS, SECTOR_REPORTS, REGIONAL_REPORTS, STATISTICS_REPORTS


class ReportListWidget(QTreeWidget):
    """보도자료 목록 위젯 (체크박스 지원)"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.load_reports()
    
    def setup_ui(self):
        """UI 구성"""
        self.setHeaderLabel("보도자료 목록")
        self.setRootIsDecorated(True)
        self.setAlternatingRowColors(True)
        self.setSelectionMode(QTreeWidget.SelectionMode.NoSelection)
        
        # 헤더 설정
        header = self.header()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        # 스타일
        self.setStyleSheet("""
            QTreeWidget {
                font-size: 10pt;
                border: 1px solid #ddd;
                border-radius: 5px;
            }
            QTreeWidget::item {
                padding: 5px;
                min-height: 25px;
            }
            QTreeWidget::item:hover {
                background-color: #f0f0f0;
            }
        """)
    
    def load_reports(self):
        """보도자료 목록 로드"""
        self.clear()
        
        # 요약 보도자료
        summary_item = QTreeWidgetItem(self, ["요약"])
        summary_item.setExpanded(True)
        summary_item.setFlags(summary_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
        
        for report in SUMMARY_REPORTS:
            item = QTreeWidgetItem(summary_item, [report['name']])
            item.setCheckState(0, Qt.CheckState.Checked)
            item.setData(0, Qt.ItemDataRole.UserRole, report)
        
        # 부문별 보도자료
        sector_item = QTreeWidgetItem(self, ["부문별"])
        sector_item.setExpanded(True)
        sector_item.setFlags(sector_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
        
        for report in SECTOR_REPORTS:
            item = QTreeWidgetItem(sector_item, [report['name']])
            item.setCheckState(0, Qt.CheckState.Checked)
            item.setData(0, Qt.ItemDataRole.UserRole, report)
        
        # 시도별 보도자료
        regional_item = QTreeWidgetItem(self, ["시도별"])
        regional_item.setExpanded(True)
        regional_item.setFlags(regional_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
        
        for report in REGIONAL_REPORTS:
            item = QTreeWidgetItem(regional_item, [report['name']])
            item.setCheckState(0, Qt.CheckState.Checked)
            item.setData(0, Qt.ItemDataRole.UserRole, report)
        
        # 통계표 보도자료
        stats_item = QTreeWidgetItem(self, ["통계표"])
        stats_item.setExpanded(True)
        stats_item.setFlags(stats_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
        
        for report in STATISTICS_REPORTS:
            item = QTreeWidgetItem(stats_item, [report['name']])
            item.setCheckState(0, Qt.CheckState.Checked)
            item.setData(0, Qt.ItemDataRole.UserRole, report)
    
    def get_selected_reports(self) -> list:
        """선택된 보도자료 리스트 반환"""
        selected = []
        
        def traverse_item(item: QTreeWidgetItem):
            if item.checkState(0) == Qt.CheckState.Checked:
                report_data = item.data(0, Qt.ItemDataRole.UserRole)
                if report_data:
                    selected.append(report_data)
            
            for i in range(item.childCount()):
                traverse_item(item.child(i))
        
        for i in range(self.topLevelItemCount()):
            traverse_item(self.topLevelItem(i))
        
        return selected
    
    def select_all(self):
        """모두 선택"""
        def set_checked(item: QTreeWidgetItem, checked: bool):
            if item.data(0, Qt.ItemDataRole.UserRole):
                item.setCheckState(0, Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
            for i in range(item.childCount()):
                set_checked(item.child(i), checked)
        
        for i in range(self.topLevelItemCount()):
            set_checked(self.topLevelItem(i), True)
    
    def deselect_all(self):
        """모두 해제"""
        def set_checked(item: QTreeWidgetItem, checked: bool):
            if item.data(0, Qt.ItemDataRole.UserRole):
                item.setCheckState(0, Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
            for i in range(item.childCount()):
                set_checked(item.child(i), checked)
        
        for i in range(self.topLevelItemCount()):
            set_checked(self.topLevelItem(i), False)
