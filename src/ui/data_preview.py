"""
Data Preview Panel - displays DataFrame data
"""

from typing import Optional
import pandas as pd

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, 
    QTableWidgetItem, QLabel, QPushButton, QHeaderView,
    QAbstractItemView, QFileDialog
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor


class DataPreviewPanel(QWidget):
    """Panel for previewing DataFrame data"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.df: Optional[pd.DataFrame] = None
        self._setup_ui()
    
    def _setup_ui(self):
        """Set up the UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # Header with info and actions
        header = QHBoxLayout()
        
        self.info_label = QLabel("暂无数据")
        self.info_label.setStyleSheet("color: #888888;")
        header.addWidget(self.info_label)
        
        header.addStretch()
        
        # Export button
        export_btn = QPushButton("📥 导出")
        export_btn.clicked.connect(self._export_data)
        header.addWidget(export_btn)
        
        layout.addLayout(header)
        
        # Table widget
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        
        layout.addWidget(self.table)
    
    def set_data(self, df: pd.DataFrame):
        """Set the DataFrame to display"""
        self.df = df
        self._update_table()
    
    def clear(self):
        """Clear the preview"""
        self.df = None
        self.table.clear()
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.info_label.setText("暂无数据")
    
    def _update_table(self):
        """Update the table with DataFrame data"""
        if self.df is None or self.df.empty:
            self.clear()
            return
        
        df = self.df
        
        # Limit rows for performance
        max_rows = 500
        if len(df) > max_rows:
            df = df.head(max_rows)
            show_truncated = True
        else:
            show_truncated = False
        
        # Update info label
        rows, cols = self.df.shape
        info_text = f"行数: {rows:,} | 列数: {cols}"
        if show_truncated:
            info_text += f" (显示前 {max_rows} 行)"
        self.info_label.setText(info_text)
        
        # Set up table
        self.table.clear()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])
        
        # Populate data
        for row_idx in range(len(df)):
            for col_idx, col in enumerate(df.columns):
                value = df.iloc[row_idx, col_idx]
                item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                
                # Color null values
                if pd.isna(value):
                    item.setBackground(QColor("#3d2d2d"))
                    item.setForeground(QColor("#888888"))
                
                # Color numbers
                elif isinstance(value, (int, float)):
                    item.setForeground(QColor("#4ade80"))
                
                self.table.setItem(row_idx, col_idx, item)
        
        # Auto-resize columns (with limit)
        for i in range(min(10, len(df.columns))):
            self.table.resizeColumnToContents(i)
    
    def _export_data(self):
        """Export current data to file"""
        if self.df is None:
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export Data",
            "data.xlsx",
            "Excel Files (*.xlsx);;CSV Files (*.csv);;All Files (*.*)"
        )
        
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.df.to_csv(file_path, index=False)
                else:
                    self.df.to_excel(file_path, index=False)
            except Exception as e:
                print(f"Export error: {e}")
