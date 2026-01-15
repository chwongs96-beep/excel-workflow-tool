"""
Global Parameters Dialog - UI for managing global workflow parameters
"""

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, 
    QTableWidgetItem, QPushButton, QLabel, QHeaderView,
    QMessageBox, QInputDialog, QFileDialog, QFrame, QWidget
)
from PyQt6.QtCore import Qt


class GlobalParamsDialog(QDialog):
    """Dialog to manage global parameters"""
    
    SYSTEM_PLACEHOLDERS = {
        "{YEAR}": "当前年份 (YYYY)",
        "{MONTH}": "当前月份 (MM)",
        "{DAY}": "当前日期 (DD)",
        "{DATE}": "完整日期 (YYYY-MM-DD)",
        "{TIME}": "当前时间 (HH:MM:SS)",
        "{TIMESTAMP}": "时间戳 (YYYYMMDD_HHMMSS)",
        "{UUID}": "唯一标识符",
        "{DESKTOP}": "桌面路径",
        "{DOCUMENTS}": "文档路径",
        "{DOWNLOADS}": "下载路径"
    }

    def __init__(self, workflow, parent=None):
        super().__init__(parent)
        self.workflow = workflow
        self.setWindowTitle("全局参数设置")
        self.resize(700, 500)
        self._setup_ui()
        self._load_params()
        
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Description
        desc_box = QFrame()
        desc_box.setStyleSheet("background-color: #f0f9ff; border: 1px solid #bae6fd; border-radius: 4px; padding: 10px;")
        desc_layout = QVBoxLayout(desc_box)
        desc_layout.setContentsMargins(5, 5, 5, 5)
        
        desc_title = QLabel("📚 关于全局参数")
        desc_title.setStyleSheet("font-weight: bold; color: #0284c7;")
        desc_layout.addWidget(desc_title)
        
        desc = QLabel(
            "在此定义全局参数，可以在节点配置中使用 {参数名} 进行引用。\n"
            "例如: 定义 base_path = C:/Data，在节点中使用 {base_path}/file.xlsx\n"
            "点击“执行”时，您可以临时修改这些值。"
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #334155;")
        desc_layout.addWidget(desc)
        layout.addWidget(desc_box)
        
        # Splitter for Params Table and Placeholder Library
        splitter = QHBoxLayout()
        
        # Left side: Parameters table
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["参数名 (Key)", "参数值 (Value)"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        left_layout.addWidget(self.table)
        
        # Buttons for table
        btn_layout = QHBoxLayout()
        
        add_btn = QPushButton("➕ 添加")
        add_btn.clicked.connect(self._add_param)
        btn_layout.addWidget(add_btn)
        
        remove_btn = QPushButton("➖ 删除")
        remove_btn.clicked.connect(self._remove_param)
        btn_layout.addWidget(remove_btn)
        
        browse_file_btn = QPushButton("📄 文件...")
        browse_file_btn.clicked.connect(self._browse_file)
        btn_layout.addWidget(browse_file_btn)
        
        browse_folder_btn = QPushButton("📂 文件夹...")
        browse_folder_btn.clicked.connect(self._browse_folder)
        btn_layout.addWidget(browse_folder_btn)
        
        left_layout.addLayout(btn_layout)
        
        # Right side: Placeholder library
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        lib_label = QLabel("🧩 占位符库 (双击插入)")
        lib_label.setStyleSheet("font-weight: bold; margin-bottom: 5px;")
        right_layout.addWidget(lib_label)
        
        self.lib_table = QTableWidget()
        self.lib_table.setColumnCount(2)
        self.lib_table.setHorizontalHeaderLabels(["占位符", "说明"])
        self.lib_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.lib_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.lib_table.verticalHeader().setVisible(False)
        self.lib_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.lib_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.lib_table.itemDoubleClicked.connect(self._insert_placeholder)
        
        # Populate library
        self.lib_table.setRowCount(len(self.SYSTEM_PLACEHOLDERS))
        for i, (key, desc) in enumerate(self.SYSTEM_PLACEHOLDERS.items()):
            self.lib_table.setItem(i, 0, QTableWidgetItem(key))
            self.lib_table.setItem(i, 1, QTableWidgetItem(desc))
            
        right_layout.addWidget(self.lib_table)
        
        # Add widget to splitter
        splitter.addWidget(left_widget, stretch=2)
        splitter.addWidget(right_widget, stretch=1)
        
        layout.addLayout(splitter)
        
        # Dialog buttons
        dialog_btns = QHBoxLayout()
        dialog_btns.addStretch()
        
        if self.parent() and getattr(self.parent(), "is_executing", False):
            ok_text = "开始执行"
        else:
            ok_text = "保存配置"
            
        ok_btn = QPushButton(ok_text)
        ok_btn.clicked.connect(self.accept)
        ok_btn.setDefault(True)
        ok_btn.setStyleSheet("background-color: #2563eb; color: white; font-weight: bold; padding: 6px 12px;")
        dialog_btns.addWidget(ok_btn)
        
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(self.reject)
        dialog_btns.addWidget(cancel_btn)
        
        layout.addLayout(dialog_btns)

    def _insert_placeholder(self, item):
        """Insert selected placeholder into the current value cell"""
        row = self.lib_table.row(item)
        placeholder = self.lib_table.item(row, 0).text()
        
        # Check currently selected destination in params table
        current_row = self.table.currentRow()
        if current_row >= 0:
            val_item = self.table.item(current_row, 1)
            current_val = val_item.text()
            self.table.setItem(current_row, 1, QTableWidgetItem(current_val + placeholder))
        else:
            # If no row selected, create new one
            self._add_param()
            row = self.table.rowCount() - 1
            self.table.setItem(row, 0, QTableWidgetItem("date_param"))
            self.table.setItem(row, 1, QTableWidgetItem(placeholder))

        
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
