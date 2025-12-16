"""
Global Parameters Dialog - UI for managing global workflow parameters
"""

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, 
    QTableWidgetItem, QPushButton, QLabel, QHeaderView,
    QMessageBox, QInputDialog, QFileDialog
)
from PyQt6.QtCore import Qt

class GlobalParamsDialog(QDialog):
    """Dialog to manage global parameters"""
    
    def __init__(self, workflow, parent=None):
        super().__init__(parent)
        self.workflow = workflow
        self.setWindowTitle("全局参数设置")
        self.resize(500, 400)
        self._setup_ui()
        self._load_params()
        
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Description
        desc = QLabel("在此定义全局参数，可以在节点配置中使用 {参数名} 进行引用。\n例如: 定义 base_path = C:/Data，在节点中使用 {base_path}/file.xlsx")
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; margin-bottom: 10px;")
        layout.addWidget(desc)
        
        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["参数名 (Key)", "参数值 (Value)"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table)
        
        # Buttons
        btn_layout = QHBoxLayout()
        
        add_btn = QPushButton("➕ 添加参数")
        add_btn.clicked.connect(self._add_param)
        btn_layout.addWidget(add_btn)
        
        remove_btn = QPushButton("➖ 删除选中")
        remove_btn.clicked.connect(self._remove_param)
        btn_layout.addWidget(remove_btn)
        
        browse_folder_btn = QPushButton("📂 浏览文件夹...")
        browse_folder_btn.setToolTip("将选中参数的值设置为文件夹路径")
        browse_folder_btn.clicked.connect(self._browse_folder)
        btn_layout.addWidget(browse_folder_btn)
        
        browse_file_btn = QPushButton("📄 浏览文件...")
        browse_file_btn.setToolTip("将选中参数的值设置为文件路径")
        browse_file_btn.clicked.connect(self._browse_file)
        btn_layout.addWidget(browse_file_btn)
        
        layout.addLayout(btn_layout)
        
        # Dialog buttons
        dialog_btns = QHBoxLayout()
        dialog_btns.addStretch()
        
        ok_btn = QPushButton("确定")
        ok_btn.clicked.connect(self.accept)
        ok_btn.setDefault(True)
        dialog_btns.addWidget(ok_btn)
        
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        dialog_btns.addWidget(cancel_btn)
        
        layout.addLayout(dialog_btns)
        
    def _load_params(self):
        """Load params from workflow to table"""
        self.table.setRowCount(0)
        for key, value in self.workflow.global_params.items():
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(key))
            self.table.setItem(row, 1, QTableWidgetItem(str(value)))
            
    def _add_param(self):
        """Add a new parameter row"""
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem("new_param"))
        self.table.setItem(row, 1, QTableWidgetItem("value"))
        self.table.editItem(self.table.item(row, 0))
        
    def _remove_param(self):
        """Remove selected parameter"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            self.table.removeRow(current_row)

    def _browse_folder(self):
        """Browse folder and set to current row value"""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "提示", "请先选择一行参数")
            return
            
        folder_path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder_path:
            self.table.setItem(current_row, 1, QTableWidgetItem(folder_path))

    def _browse_file(self):
        """Browse file and set to current row value"""
        current_row = self.table.currentRow()
        if current_row < 0:
            QMessageBox.warning(self, "提示", "请先选择一行参数")
            return
            
        file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "All Files (*.*)")
        if file_path:
            self.table.setItem(current_row, 1, QTableWidgetItem(file_path))
            
    def get_params(self):
        """Get params from table"""
        params = {}
        for row in range(self.table.rowCount()):
            key_item = self.table.item(row, 0)
            val_item = self.table.item(row, 1)
            
            if key_item and val_item:
                key = key_item.text().strip()
                val = val_item.text().strip()
                if key:
                    params[key] = val
        return params
