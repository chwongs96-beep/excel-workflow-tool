"""
Node Configuration Panel - UI for configuring node parameters
"""

from typing import Optional
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLabel, QLineEdit, QTextEdit, QSpinBox, QDoubleSpinBox,
    QComboBox, QCheckBox, QPushButton, QFileDialog,
    QScrollArea, QFrame, QSizePolicy, QMessageBox, QToolButton
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QColor

import sys
from pathlib import Path
import pandas as pd
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from src.nodes.base_node import BaseNode


class NodeConfigPanel(QWidget):
    """Panel for configuring node parameters"""
    
    config_changed = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.node: Optional[BaseNode] = None
        self.workflow = None  # Reference to workflow for upstream traversal
        self.field_widgets = {}
        
        self._setup_ui()
    
    def set_workflow(self, workflow):
        """Set the workflow reference"""
        self.workflow = workflow
    
    def _setup_ui(self):
        """Set up the UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # Header with help button
        header_layout = QHBoxLayout()
        
        self.header_label = QLabel("未选择节点")
        self.header_label.setStyleSheet("font-size: 14px; font-weight: bold; padding: 5px;")
        header_layout.addWidget(self.header_label, 1)
        
        # Help button
        self.help_btn = QToolButton()
        self.help_btn.setText("❓")
        self.help_btn.setToolTip("查看节点使用说明")
        self.help_btn.setStyleSheet("""
            QToolButton {
                background-color: #3d5afe;
                color: white;
                border: none;
                border-radius: 12px;
                font-size: 14px;
                font-weight: bold;
                min-width: 24px;
                min-height: 24px;
                max-width: 24px;
                max-height: 24px;
            }
            QToolButton:hover {
                background-color: #536dfe;
            }
            QToolButton:pressed {
                background-color: #304ffe;
            }
        """)
        self.help_btn.clicked.connect(self._show_help)
        self.help_btn.hide()  # Hide initially
        header_layout.addWidget(self.help_btn)
        
        layout.addLayout(header_layout)
        
        # Description
        self.desc_label = QLabel("")
        self.desc_label.setWordWrap(True)
        self.desc_label.setStyleSheet("color: #888888; padding: 5px;")
        layout.addWidget(self.desc_label)
        
        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("background-color: #3d3d3d;")
        layout.addWidget(line)
        
        # Scroll area for config fields
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        
        self.config_widget = QWidget()
        self.config_layout = QFormLayout(self.config_widget)
        self.config_layout.setContentsMargins(5, 10, 5, 10)
        self.config_layout.setSpacing(10)
        
        scroll.setWidget(self.config_widget)
        layout.addWidget(scroll, 1)
        
        # Placeholder
        self.placeholder = QLabel("请选择一个节点来配置其参数")
        self.placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.placeholder.setStyleSheet("color: #666666;")
        layout.addWidget(self.placeholder)
    
    def set_node(self, node: BaseNode):
        """Set the node to configure"""
        self.node = node
        self._clear_config_fields()
        
        # Update header
        self.header_label.setText(node.node_name)
        color = QColor(node.node_color)
        self.header_label.setStyleSheet(
            f"font-size: 14px; font-weight: bold; padding: 5px; "
            f"background-color: {color.name()}; border-radius: 3px;"
        )
        
        self.desc_label.setText(node.node_description)
        self.placeholder.hide()
        self.help_btn.show()  # Show help button when node is selected
        
        # Create config fields
        schema = node.get_config_ui_schema()
        for field in schema:
            self._create_field(field)
    
    def clear(self):
        """Clear the panel"""
        self.node = None
        self._clear_config_fields()
        self.header_label.setText("未选择节点")
        self.header_label.setStyleSheet("font-size: 14px; font-weight: bold; padding: 5px;")
        self.desc_label.setText("")
        self.placeholder.show()
        self.help_btn.hide()  # Hide help button when no node selected
    
    def _clear_config_fields(self):
        """Clear all config fields"""
        while self.config_layout.count():
            item = self.config_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        self.field_widgets.clear()
    
    def _find_upstream_file_path(self) -> Optional[str]:
        """Find file path from upstream nodes"""
        if not self.workflow or not self.node:
            return None
            
        # Find upstream nodes
        upstream_nodes = []
        for conn in self.workflow.connections:
            if conn.to_node == self.node.node_id:
                if conn.from_node in self.workflow.nodes:
                    upstream_nodes.append(self.workflow.nodes[conn.from_node])
        
        # Check upstream nodes for file paths
        for up_node in upstream_nodes:
            # Check for 'file_path' or 'base_file' or 'output_file'
            # We check the CONFIG of the upstream node
            for key in ['file_path', 'base_file', 'output_file']:
                val = up_node.get_param(key)
                if val and isinstance(val, str) and (val.lower().endswith('.xlsx') or val.lower().endswith('.xls') or val.lower().endswith('.csv')):
                    return val
        return None

    def _create_field(self, field: dict):
        """Create a config field based on schema"""
        key = field["key"]
        label = field.get("label", key)
        field_type = field.get("type", "text")
        default = field.get("default", "")
        current_value = self.node.get_param(key, default)
        
        # Auto-fill file path from upstream if empty
        if not current_value and field_type == "file" and self.workflow:
            upstream_path = self._find_upstream_file_path()
            if upstream_path:
                current_value = upstream_path
                self.node.set_param(key, current_value)
                # We don't emit config_changed here to avoid loop/overhead during init,
                # but the widget creation below will use the new value.
        
        # Create label
        label_widget = QLabel(label + ":")
        if field.get("required"):
            label_widget.setText(label + " *:")
        
        # Create input widget based on type
        if field_type == "text":
            widget = QLineEdit()
            widget.setText(str(current_value))
            widget.setPlaceholderText(field.get("placeholder", ""))
            widget.textChanged.connect(lambda v, k=key: self._on_value_changed(k, v))
        
        elif field_type == "textarea":
            widget = QTextEdit()
            widget.setPlainText(str(current_value))
            widget.setMaximumHeight(100)
            widget.textChanged.connect(lambda k=key: self._on_value_changed(k, widget.toPlainText()))
        
        elif field_type == "number":
            widget = QSpinBox()
            widget.setMinimum(field.get("min", 0))
            widget.setMaximum(field.get("max", 999999))
            widget.setValue(int(current_value) if current_value else 0)
            widget.valueChanged.connect(lambda v, k=key: self._on_value_changed(k, v))
        
        elif field_type == "decimal":
            widget = QDoubleSpinBox()
            widget.setMinimum(field.get("min", 0.0))
            widget.setMaximum(field.get("max", 999999.0))
            widget.setDecimals(field.get("decimals", 2))
            widget.setValue(float(current_value) if current_value else 0.0)
            widget.valueChanged.connect(lambda v, k=key: self._on_value_changed(k, v))
        
        elif field_type == "select":
            widget = QComboBox()
            for option in field.get("options", []):
                widget.addItem(option["label"], option["value"])
            
            # Set current value
            idx = widget.findData(current_value)
            if idx >= 0:
                widget.setCurrentIndex(idx)
            
            widget.currentIndexChanged.connect(
                lambda _, k=key, w=widget: self._on_value_changed(k, w.currentData())
            )
        
        elif field_type == "checkbox":
            widget = QCheckBox()
            widget.setChecked(bool(current_value))
            widget.stateChanged.connect(
                lambda state, k=key: self._on_value_changed(k, state == Qt.CheckState.Checked.value)
            )
        
        elif field_type == "file":
            widget = QWidget()
            hlayout = QHBoxLayout(widget)
            hlayout.setContentsMargins(0, 0, 0, 0)
            
            line_edit = QLineEdit()
            line_edit.setText(str(current_value))
            line_edit.setPlaceholderText(field.get("placeholder", ""))
            line_edit.textChanged.connect(lambda v, k=key: self._on_value_changed(k, v))
            
            browse_btn = QPushButton("...")
            browse_btn.setMaximumWidth(30)
            browse_btn.clicked.connect(
                lambda _, k=key, le=line_edit, f=field: self._browse_file(k, le, f)
            )
            
            hlayout.addWidget(line_edit)
            hlayout.addWidget(browse_btn)
        
        elif field_type == "file_save":
            widget = QWidget()
            hlayout = QHBoxLayout(widget)
            hlayout.setContentsMargins(0, 0, 0, 0)
            
            line_edit = QLineEdit()
            line_edit.setText(str(current_value))
            line_edit.setPlaceholderText(field.get("placeholder", ""))
            line_edit.textChanged.connect(lambda v, k=key: self._on_value_changed(k, v))
            
            browse_btn = QPushButton("...")
            browse_btn.setMaximumWidth(30)
            browse_btn.clicked.connect(
                lambda _, k=key, le=line_edit, f=field: self._browse_file_save(k, le, f)
            )
            
            hlayout.addWidget(line_edit)
            hlayout.addWidget(browse_btn)
        
        elif field_type == "file_multiple":
            widget = QWidget()
            hlayout = QHBoxLayout(widget)
            hlayout.setContentsMargins(0, 0, 0, 0)
            
            text_edit = QTextEdit()
            text_edit.setPlainText(str(current_value))
            text_edit.setMaximumHeight(60)
            text_edit.textChanged.connect(lambda: self._on_value_changed(key, text_edit.toPlainText()))
            
            browse_btn = QPushButton("...")
            browse_btn.setMaximumWidth(30)
            browse_btn.clicked.connect(
                lambda _, k=key, te=text_edit, f=field: self._browse_files(k, te, f)
            )
            
            hlayout.addWidget(text_edit)
            hlayout.addWidget(browse_btn)
        
        elif field_type == "directory":
            widget = QWidget()
            hlayout = QHBoxLayout(widget)
            hlayout.setContentsMargins(0, 0, 0, 0)
            
            line_edit = QLineEdit()
            line_edit.setText(str(current_value))
            line_edit.setPlaceholderText(field.get("placeholder", ""))
            line_edit.textChanged.connect(lambda v, k=key: self._on_value_changed(k, v))
            
            browse_btn = QPushButton("📂")
            browse_btn.setMaximumWidth(30)
            browse_btn.clicked.connect(
                lambda _, k=key, le=line_edit, f=field: self._browse_directory(k, le, f)
            )
            
            hlayout.addWidget(line_edit)
            hlayout.addWidget(browse_btn)
        
        elif field_type == "sheet_selector":
            widget = self._create_sheet_selector_widget(key, field, current_value)

        else:
            widget = QLineEdit()
            widget.setText(str(current_value))
            widget.textChanged.connect(lambda v, k=key: self._on_value_changed(k, v))
        
        self.config_layout.addRow(label_widget, widget)
        self.field_widgets[key] = widget
    
    def _create_sheet_selector_widget(self, key: str, field: dict, current_value):
        """Create a sheet selector widget"""
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        combo.setCurrentText(str(current_value))
        combo.currentTextChanged.connect(lambda t: self._on_value_changed(key, t))
        
        # Set size policy to expand
        combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        
        layout.addWidget(combo)
        
        refresh_btn = QPushButton("🔄")
        refresh_btn.setToolTip("刷新工作表列表")
        refresh_btn.setMaximumWidth(30)
        refresh_btn.clicked.connect(lambda: self._refresh_sheets(combo, field))
        layout.addWidget(refresh_btn)
        
        # Handle dependency
        dependency = field.get("dependency")
        if dependency:
            file_path = None
            
            if dependency == "__upstream__":
                file_path = self._find_upstream_file_path()
            else:
                # Initial population
                # First check if we have a value in the node
                file_path = self.node.get_param(dependency)
                
                # If not, try to find it from upstream (auto-fill logic might have missed it or not run yet)
                if not file_path and self.workflow:
                    upstream_path = self._find_upstream_file_path()
                    if upstream_path:
                        file_path = upstream_path
                        # We don't set param here to avoid side effects during widget creation,
                        # but we use it to populate the list.
            
            if file_path:
                self._populate_sheets(combo, file_path)
            
            # Connect to dependency change (only for local params)
            if dependency != "__upstream__":
                dep_widget = self.field_widgets.get(dependency)
                if dep_widget:
                    # Find QLineEdit in the dependency widget
                    line_edits = dep_widget.findChildren(QLineEdit)
                    if line_edits:
                        le = line_edits[0]
                        le.textChanged.connect(lambda text: self._populate_sheets(combo, text))
        
        return container

    def _refresh_sheets(self, combo: QComboBox, field: dict):
        """Refresh sheet list"""
        dependency = field.get("dependency")
        if dependency:
            file_path = None
            
            if dependency == "__upstream__":
                file_path = self._find_upstream_file_path()
            else:
                file_path = self.node.get_param(dependency)
                
                # If empty, try upstream
                if not file_path and self.workflow:
                    upstream_path = self._find_upstream_file_path()
                    if upstream_path:
                        file_path = upstream_path
                        # Update the node param and the UI widget
                        self.node.set_param(dependency, file_path)
                        if dependency in self.field_widgets:
                            # Find QLineEdit
                            line_edits = self.field_widgets[dependency].findChildren(QLineEdit)
                            if line_edits:
                                line_edits[0].setText(file_path)
            
            if file_path:
                self._populate_sheets(combo, file_path)
    
    def _populate_sheets(self, combo: QComboBox, file_path: str):
        """Populate combo box with sheets from Excel file"""
        if not file_path or not Path(file_path).exists():
            return
            
        try:
            current = combo.currentText()
            sheets = []
            
            # Handle CSV files
            if str(file_path).lower().endswith('.csv'):
                # CSV files don't have sheets, but we provide a default option
                # Use the filename as the sheet name or just "Sheet1"
                sheets = ["Sheet1"]
            else:
                # Read excel file to get sheet names
                # Use openpyxl engine explicitly for xlsx
                try:
                    xl = pd.ExcelFile(file_path)
                    sheets = xl.sheet_names
                except Exception:
                    # Fallback or retry?
                    # If it fails, maybe it's open in another app?
                    # But pandas usually handles read-only fine.
                    pass
            
            if not sheets:
                return

            combo.blockSignals(True)
            combo.clear()
            combo.addItems(sheets)
            
            if current:
                # If user has typed something, keep it (allows for custom target names)
                combo.setCurrentText(current)
            elif sheets:
                # If nothing typed, default to first sheet
                combo.setCurrentIndex(0)
            
            combo.blockSignals(False)
            
            # If we defaulted to first sheet (and nothing was there before), emit change
            if not current and sheets:
                combo.currentTextChanged.emit(sheets[0])
                
        except Exception as e:
            print(f"Error reading sheets: {e}")

    def _on_value_changed(self, key: str, value):
        """Handle value change"""
        if self.node:
            self.node.set_param(key, value)
            self.config_changed.emit()
    
    def _browse_file(self, key: str, line_edit: QLineEdit, field: dict):
        """Open file browser"""
        file_filter = field.get("file_filter", "All Files (*.*)")
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", file_filter)
        if file_path:
            line_edit.setText(file_path)
    
    def _browse_file_save(self, key: str, line_edit: QLineEdit, field: dict):
        """Open file save dialog"""
        file_filter = field.get("file_filter", "All Files (*.*)")
        file_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", file_filter)
        if file_path:
            line_edit.setText(file_path)

    def _browse_files(self, key: str, text_edit: QTextEdit, field: dict):
        """Open file browser for multiple files"""
        file_filter = field.get("file_filter", "All Files (*.*)")
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", file_filter)
        if file_paths:
            text_edit.setPlainText("\n".join(file_paths))

    def _browse_directory(self, key: str, line_edit: QLineEdit, field: dict):
        """Open directory browser"""
        dir_path = QFileDialog.getExistingDirectory(self, "Select Directory", "")
        if dir_path:
            line_edit.setText(dir_path)
    
    def _show_help(self):
        """Show help dialog for the current node"""
        if not self.node:
            return
        
        # Build help content based on node type and config schema
        node_name = self.node.node_name
        node_desc = self.node.node_description
        node_category = self.node.node_category
        
        # Get config schema for detailed help
        schema = self.node.get_config_ui_schema()
        
        help_text = f"""
<h2 style="color: #4fc3f7;">{node_name}</h2>
<p><b>类别:</b> {node_category}</p>
<p><b>功能描述:</b> {node_desc}</p>

<h3 style="color: #81c784;">参数说明:</h3>
<table style="width:100%; border-collapse: collapse;">
"""
        
        for field in schema:
            label = field.get("label", field.get("key", ""))
            field_type = field.get("type", "text")
            required = "必填" if field.get("required") else "可选"
            placeholder = field.get("placeholder", "")
            default = field.get("default", "")
            
            type_desc = {
                "text": "文本输入",
                "textarea": "多行文本",
                "number": "数字",
                "decimal": "小数",
                "select": "下拉选择",
                "checkbox": "复选框",
                "file": "选择文件",
                "file_save": "保存文件"
            }.get(field_type, field_type)
            
            help_text += f"""
<tr style="border-bottom: 1px solid #444;">
    <td style="padding: 8px; font-weight: bold; color: #fff;">{label}</td>
    <td style="padding: 8px; color: #aaa;">{type_desc} ({required})</td>
</tr>
"""
            if placeholder:
                help_text += f"""
<tr>
    <td colspan="2" style="padding: 4px 8px; color: #888; font-size: 12px;">
        💡 提示: {placeholder}
    </td>
</tr>
"""
        
        help_text += """
</table>

<h3 style="color: #ffb74d;">使用步骤:</h3>
<ol>
    <li>将此节点拖拽到画布上</li>
    <li>连接输入端口（如果需要输入数据）</li>
    <li>在右侧配置面板中填写必要参数</li>
    <li>连接输出端口到下一个节点</li>
    <li>点击"执行"按钮运行工作流</li>
</ol>

<p style="color: #90caf9; margin-top: 15px;">
<b>Excel Workflow Tool</b> - Excel 工作流自动化工具
</p>
"""
        
        msg = QMessageBox(self)
        msg.setWindowTitle(f"节点帮助 - {node_name}")
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setText(help_text)
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setStyleSheet("""
            QMessageBox {
                background-color: #2d2d2d;
            }
            QMessageBox QLabel {
                color: #e0e0e0;
                min-width: 400px;
            }
            QPushButton {
                background-color: #3d5afe;
                color: white;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #536dfe;
            }
        """)
        msg.exec()
