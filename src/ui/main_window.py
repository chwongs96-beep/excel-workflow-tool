"""
Main Window - the primary UI for the Excel Workflow Tool
"""

import os
import sys
import re
import json
import copy
import uuid
import platform
import warnings
from pathlib import Path
from typing import Optional, List, Dict, Any
from datetime import datetime

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QSplitter, QMenuBar, QMenu, QToolBar, QStatusBar,
    QFileDialog, QMessageBox, QLabel, QPushButton, QDialog,
    QDockWidget, QListWidget, QListWidgetItem, QFrame,
    QScrollArea, QSizePolicy, QLineEdit, QApplication, QPlainTextEdit,
    QToolButton
)
from PyQt6.QtCore import Qt, QSize, QMimeData, QPoint, QSettings, pyqtSignal
from PyQt6.QtGui import QAction, QIcon, QDrag, QColor, QPalette, QPixmap, QFont, QPainter

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from src.workflow.engine import Workflow
from src.workflow.worker import WorkflowWorker
from src.nodes.node_registry import NodeRegistry
from src.nodes import excel_nodes  # Import to register nodes
from src.nodes import merge_nodes  # Import to register merge nodes
from src.ui.canvas import WorkflowCanvas
from src.ui.node_config import NodeConfigPanel
from src.ui.data_preview import DataPreviewPanel
from src.ui.about_dialog import AboutDialog
from src.ui.global_params import GlobalParamsDialog
from src.utils import get_resource_path


class NodeListItem(QListWidgetItem):
    """Custom list item for node palette"""
    
    def __init__(self, node_class):
        super().__init__(node_class.node_name)
        self.node_class = node_class
        self.node_type = node_class.node_type


class DraggableNodeList(QListWidget):
    """List widget with drag support for nodes"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragEnabled(True)
        self.setDragDropMode(QListWidget.DragDropMode.DragOnly)
    
    def startDrag(self, supportedActions):
        """Start drag operation with node type data"""
        try:
            item = self.currentItem()
            if isinstance(item, NodeListItem):
                drag = QDrag(self)
                mime_data = QMimeData()
                # Store node type in mime data
                mime_data.setText(item.node_type)
                mime_data.setData("application/x-workflow-node", item.node_type.encode())
                drag.setMimeData(mime_data)
                
                # Create drag pixmap
                pixmap = QPixmap(160, 40)
                pixmap.fill(QColor(item.node_class.node_color))
                
                from PyQt6.QtGui import QPainter as QPainterDrag
                painter = QPainterDrag(pixmap)
                painter.setPen(QColor("white"))
                font = QFont("Segoe UI", 10, QFont.Weight.Bold)
                painter.setFont(font)
                painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, item.node_class.node_name)
                painter.end()
                
                drag.setPixmap(pixmap)
                drag.setHotSpot(QPoint(80, 20))
                
                drag.exec(Qt.DropAction.CopyAction)
        except Exception as e:
            print(f"Drag error: {e}")
        

class NodePalette(QDockWidget):
    """Dock widget containing available nodes"""

    recent_add_requested = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__("节点列表", parent)
        self.setAllowedAreas(Qt.DockWidgetArea.LeftDockWidgetArea | Qt.DockWidgetArea.RightDockWidgetArea)
        
        # Store all node items for filtering
        self.all_node_items = []
        
        # Track category collapsed/expanded state
        # Only "灵活合并" is expanded by default, others are collapsed
        self.category_expanded = {}
        
        # Create main widget
        main_widget = QWidget()
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(5, 5, 5, 5)
        
        # Search box
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("🔍 搜索节点...")
        self.search_box.textChanged.connect(self._filter_nodes)
        self.search_box.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 1px solid #3d3d3d;
                border-radius: 4px;
                background-color: #2d2d2d;
                color: #e0e0e0;
            }
            QLineEdit:focus {
                border-color: #db0011;
            }
        """)
        layout.addWidget(self.search_box)

        self.recent_container = QWidget()
        recent_outer = QVBoxLayout(self.recent_container)
        recent_outer.setContentsMargins(0, 4, 0, 0)
        recent_outer.setSpacing(4)
        rl = QLabel("最近使用：")
        rl.setStyleSheet("color: #888888; font-size: 11px;")
        recent_outer.addWidget(rl)
        self.recent_bar_layout = QHBoxLayout()
        self.recent_bar_layout.setContentsMargins(0, 0, 0, 0)
        self.recent_bar_layout.setSpacing(4)
        recent_outer.addLayout(self.recent_bar_layout)
        layout.addWidget(self.recent_container)
        
        # Create draggable list widget for nodes
        self.node_list = DraggableNodeList()
        self.node_list.setSpacing(2)
        self.node_list.itemClicked.connect(self._on_item_clicked)
        
        # Populate with nodes by category
        self._populate_nodes()
        
        layout.addWidget(QLabel("拖拽或双击添加节点:"))
        layout.addWidget(self.node_list)
        
        self.setWidget(main_widget)
        self.setMinimumWidth(200)
        self.refresh_recent_bar()
    
    def refresh_recent_bar(self):
        """从设置中恢复「最近使用」节点快捷按钮。"""
        while self.recent_bar_layout.count():
            item = self.recent_bar_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        settings = QSettings("Excel Workflow Tool", "Settings")
        recent = settings.value("recent_node_types", []) or []
        if not recent:
            self.recent_container.hide()
            return
        added = 0
        for nt in recent:
            node_class = NodeRegistry.get_node_class(nt)
            if not node_class:
                continue
            btn = QToolButton()
            btn.setText(node_class.node_name)
            btn.setToolTip(node_class.node_description)
            btn.setAutoRaise(True)
            btn.clicked.connect(lambda _c=False, t=nt: self.recent_add_requested.emit(t))
            self.recent_bar_layout.addWidget(btn)
            added += 1
        if added == 0:
            self.recent_container.hide()
            return
        self.recent_container.show()
        self.recent_bar_layout.addStretch(1)
    
    def _populate_nodes(self):
        """Populate the node list"""
        self.node_list.clear()
        self.all_node_items = []
        
        categories = NodeRegistry.get_nodes_by_category()
        for category, nodes in sorted(categories.items()):
            # Initialize collapse state: only "灵活合并" is expanded by default
            if category not in self.category_expanded:
                self.category_expanded[category] = (category == "灵活合并")
            
            is_expanded = self.category_expanded[category]
            
            # Add category header with expand/collapse indicator
            expand_icon = "▼" if is_expanded else "▶"
            header = QListWidgetItem(f"{expand_icon} {category}")
            header.setFlags(Qt.ItemFlag.ItemIsEnabled)  # Make it clickable but not draggable
            header.setBackground(QColor("#2d2d2d"))
            header.setForeground(QColor("#db0011"))  # Use brand color for headers
            header.setData(Qt.ItemDataRole.UserRole, f"header:{category}")
            font = header.font()
            font.setBold(True)
            header.setFont(font)
            self.node_list.addItem(header)
            self.all_node_items.append((header, category, None))
            
            # Add nodes in this category (hidden if category is collapsed)
            for node_class in nodes:
                item = NodeListItem(node_class)
                item.setToolTip(node_class.node_description)
                # Set background color based on node color
                color = QColor(node_class.node_color)
                color.setAlpha(50)
                item.setBackground(color)
                item.setHidden(not is_expanded)  # Hide if category is collapsed
                self.node_list.addItem(item)
                self.all_node_items.append((item, category, node_class))
    
    def _on_item_clicked(self, item):
        """Handle item click - toggle category if header is clicked"""
        user_data = item.data(Qt.ItemDataRole.UserRole)
        
        # Check if it's a category header
        if isinstance(user_data, str) and user_data.startswith("header:"):
            category = user_data.replace("header:", "")
            self._toggle_category(category)
    
    def _toggle_category(self, category):
        """Toggle collapse/expand state of a category"""
        # Toggle the state
        self.category_expanded[category] = not self.category_expanded.get(category, False)
        is_expanded = self.category_expanded[category]
        
        # Update header text and node visibility
        for item, item_category, node_class in self.all_node_items:
            if item_category == category:
                if node_class is None:  # It's the header
                    expand_icon = "▼" if is_expanded else "▶"
                    item.setText(f"{expand_icon} {category}")
                else:  # It's a node in this category
                    # Only show if category is expanded AND search filter allows it
                    if self.search_box.text().strip():
                        # If search is active, respect search filter
                        continue
                    else:
                        # No search, just toggle based on category state
                        item.setHidden(not is_expanded)
    
    def _filter_nodes(self, text: str):
        """Filter nodes based on search text"""
        search_text = text.lower().strip()
        
        if not search_text:
            # No search - restore category collapsed/expanded states
            for item, category, node_class in self.all_node_items:
                if node_class is None:  # It's a header
                    item.setHidden(False)
                else:  # It's a node
                    # Show only if category is expanded
                    is_expanded = self.category_expanded.get(category, False)
                    item.setHidden(not is_expanded)
            return
        
        # Track which categories have visible nodes
        visible_categories = set()
        
        # First pass: find matching nodes
        for item, category, node_class in self.all_node_items:
            if node_class is not None:  # It's a node, not a header
                # Search in node name and description
                name_match = search_text in node_class.node_name.lower()
                desc_match = search_text in node_class.node_description.lower()
                type_match = search_text in node_class.node_type.lower()
                
                if name_match or desc_match or type_match:
                    item.setHidden(False)
                    visible_categories.add(category)
                else:
                    item.setHidden(True)
        
        # Second pass: show/hide category headers
        for item, category, node_class in self.all_node_items:
            if node_class is None:  # It's a header
                item.setHidden(category not in visible_categories)


class MainWindow(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        
        # Suppress warnings from pandas and openpyxl
        warnings.filterwarnings('ignore', category=UserWarning)
        warnings.filterwarnings('ignore', category=FutureWarning)
        warnings.filterwarnings('ignore', category=DeprecationWarning)
        
        self.workflow = Workflow()
        self.current_file: Optional[str] = None
        
        # Worker thread for background execution
        self._worker: Optional[WorkflowWorker] = None
        self._is_executing = False
        
        # Undo/Redo history
        self._undo_stack: List[Dict[str, Any]] = []
        self._redo_stack: List[Dict[str, Any]] = []
        self._max_history = 50  # Maximum undo steps
        
        # Theme state (True = dark, False = light) - Default to light theme
        self._is_dark_theme = False
        
        # Auto-save directory
        self._auto_save_dir = Path(__file__).parent.parent.parent / "autosave"
        self._auto_save_dir.mkdir(exist_ok=True)
        
        # Track if workflow has unsaved changes
        self._has_unsaved_changes = False
        
        # Application settings
        self._settings = QSettings("Excel Workflow Tool", "Settings")
        
        self._setup_ui()
        self._setup_menu()
        self._setup_toolbar()
        self._setup_statusbar()
        self._setup_branding()
        
        # Save initial state
        self._save_state()
        
        self.setWindowTitle("Excel 工作流工具")
        self.resize(1400, 900)
        
        # Restore saved settings (geometry, theme, etc.)
        self._restore_settings()
        
        # Apply theme based on restored setting
        if self._is_dark_theme:
            self._apply_dark_theme()
            self.canvas.set_theme(dark=True)
            self._update_brand_style(dark=True)
            self.theme_btn.setText("🌙 浅色")
            self.theme_action.setText("🌙 切换到浅色模式")
        else:
            self._apply_light_theme()
            self.canvas.set_theme(dark=False)
            self._update_brand_style(dark=False)
            self.theme_btn.setText("☀️ 深色")
            self.theme_action.setText("☀️ 切换到深色模式")
    
    def _setup_ui(self):
        """Set up the main UI layout"""
        # Central widget with canvas
        central_widget = QWidget()
        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create workflow canvas
        self.canvas = WorkflowCanvas(self.workflow)
        self.canvas.node_selected.connect(self._on_node_selected)
        self.canvas.node_double_clicked.connect(self._on_node_double_clicked)
        self.canvas.connection_created.connect(self._on_connection_created)
        self.canvas.node_delete_requested.connect(self._on_node_delete_requested)
        self.canvas.node_copy_requested.connect(self._on_node_copy_requested)
        self.canvas.node_execution_requested.connect(self._execute_node)
        self.canvas.workflow_execution_requested.connect(self._execute_workflow)
        self.canvas.node_dropped.connect(self._on_node_dropped)
        central_layout.addWidget(self.canvas)
        
        self.setCentralWidget(central_widget)
        
        # Node palette (left dock)
        self.node_palette = NodePalette(self)
        self.node_palette.node_list.itemDoubleClicked.connect(self._on_palette_item_double_clicked)
        self.node_palette.recent_add_requested.connect(self._on_recent_node_quick_add)
        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.node_palette)
        
        # Node config panel (right dock)
        self.config_dock = QDockWidget("节点配置", self)
        self.config_panel = NodeConfigPanel()
        self.config_panel.set_workflow(self.workflow) # Pass workflow reference
        self.config_panel.config_changed.connect(self._on_config_changed)
        self.config_panel.execution_requested.connect(self._execute_node)
        self.config_dock.setWidget(self.config_panel)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.config_dock)
        
        # Data preview panel (bottom dock)
        self.preview_dock = QDockWidget("数据预览", self)
        self.preview_panel = DataPreviewPanel()
        self.preview_dock.setWidget(self.preview_panel)
        self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, self.preview_dock)

        self.log_dock = QDockWidget("执行日志", self)
        self.log_panel = QPlainTextEdit()
        self.log_panel.setReadOnly(True)
        self.log_panel.setMaximumBlockCount(8000)
        self.log_panel.setPlaceholderText("执行工作流后，此处显示各节点进度与 report_progress 输出…")
        self.log_dock.setWidget(self.log_panel)
        self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, self.log_dock)
        self.tabifyDockWidget(self.preview_dock, self.log_dock)
        self.preview_dock.raise_()
        
        # Set dock sizes
        self.config_dock.setMinimumWidth(280)
        self.preview_dock.setMinimumHeight(200)
        self.log_dock.setMinimumHeight(160)
    
    def _setup_menu(self):
        """Set up menu bar"""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("文件(&F)")
        
        new_action = QAction("新建工作流(&N)", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self._new_workflow)
        file_menu.addAction(new_action)
        
        open_action = QAction("打开工作流(&O)...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self._open_workflow)
        file_menu.addAction(open_action)
        
        save_action = QAction("保存工作流(&S)", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self._save_workflow)
        file_menu.addAction(save_action)
        
        save_as_action = QAction("另存为(&A)...", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self._save_workflow_as)
        file_menu.addAction(save_as_action)
        
        file_menu.addSeparator()
        
        reload_action = QAction("刷新工作流(&R)", self)
        reload_action.setShortcut("Ctrl+R")
        reload_action.triggered.connect(self._reload_workflow)
        file_menu.addAction(reload_action)
        
        file_menu.addSeparator()
        
        # Recent files submenu
        self.recent_menu = QMenu("最近打开(&R)", self)
        self._update_recent_menu()
        file_menu.addMenu(self.recent_menu)
        
        # Templates submenu
        self.templates_menu = QMenu("工作流模板(&T)", self)
        self._setup_templates_menu()
        file_menu.addMenu(self.templates_menu)
        
        file_menu.addSeparator()
        
        # Export/Import submenu
        export_menu = QMenu("导出/导入(&E)", self)
        
        export_workflow_action = QAction("📤 导出工作流...", self)
        export_workflow_action.triggered.connect(self._export_workflow)
        export_menu.addAction(export_workflow_action)
        
        import_workflow_action = QAction("📥 导入工作流...", self)
        import_workflow_action.triggered.connect(self._import_workflow)
        export_menu.addAction(import_workflow_action)
        
        export_menu.addSeparator()
        
        export_image_action = QAction("🖼️ 导出为图片...", self)
        export_image_action.triggered.connect(self._export_as_image)
        export_menu.addAction(export_image_action)
        
        file_menu.addMenu(export_menu)
        
        file_menu.addSeparator()
        
        # Global Params
        params_action = QAction("🌐 全局参数设置...", self)
        params_action.triggered.connect(self._show_global_params)
        file_menu.addAction(params_action)
        
        file_menu.addSeparator()
        
        # Restart app
        restart_action = QAction("重启应用(&T)", self)
        restart_action.setShortcut("Ctrl+Shift+R")
        restart_action.triggered.connect(self._restart_app)
        file_menu.addAction(restart_action)
        
        exit_action = QAction("退出(&X)", self)
        exit_action.setShortcut("Alt+F4")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Edit menu
        edit_menu = menubar.addMenu("编辑(&E)")
        
        self.undo_action = QAction("撤销(&U)", self)
        self.undo_action.setShortcut("Ctrl+Z")
        self.undo_action.triggered.connect(self._undo)
        self.undo_action.setEnabled(False)
        edit_menu.addAction(self.undo_action)
        
        self.redo_action = QAction("重做(&R)", self)
        self.redo_action.setShortcut("Ctrl+Y")
        self.redo_action.triggered.connect(self._redo)
        self.redo_action.setEnabled(False)
        edit_menu.addAction(self.redo_action)
        
        edit_menu.addSeparator()
        
        delete_action = QAction("删除选中(&D)", self)
        delete_action.setShortcut("Delete")
        delete_action.triggered.connect(self._delete_selected)
        edit_menu.addAction(delete_action)
        
        # Run menu
        run_menu = menubar.addMenu("运行(&R)")
        
        run_action = QAction("执行工作流(&E)", self)
        run_action.setShortcut("F6")
        run_action.triggered.connect(self._execute_workflow)
        run_menu.addAction(run_action)
        
        run_menu.addSeparator()
        
        params_run_action = QAction("🌐 全局参数设置...", self)
        params_run_action.triggered.connect(self._show_global_params)
        run_menu.addAction(params_run_action)
        
        # View menu
        view_menu = menubar.addMenu("视图(&V)")
        
        view_menu.addAction(self.node_palette.toggleViewAction())
        view_menu.addAction(self.config_dock.toggleViewAction())
        view_menu.addAction(self.preview_dock.toggleViewAction())
        view_menu.addAction(self.log_dock.toggleViewAction())
        
        view_menu.addSeparator()
        
        # Minimap toggle
        self.minimap_action = QAction("🗺️ 小地图", self)
        self.minimap_action.setCheckable(True)
        self.minimap_action.setChecked(True)
        self.minimap_action.setShortcut("Ctrl+M")
        self.minimap_action.triggered.connect(self._toggle_minimap)
        view_menu.addAction(self.minimap_action)
        
        # Fit to view
        fit_action = QAction("📐 适应窗口", self)
        fit_action.setShortcut("Ctrl+0")
        fit_action.triggered.connect(lambda: self.canvas.fit_to_view())
        view_menu.addAction(fit_action)
        
        view_menu.addSeparator()
        
        # Theme toggle
        self.theme_action = QAction("🌙 切换到浅色模式", self)
        self.theme_action.setShortcut("Ctrl+T")
        self.theme_action.triggered.connect(self._toggle_theme)
        view_menu.addAction(self.theme_action)
        view_menu.addAction(self.theme_action)
        
        # Help menu
        help_menu = menubar.addMenu("帮助(&H)")
        
        shortcuts_action = QAction("快捷键列表(&K)", self)
        shortcuts_action.setShortcut("Ctrl+/")
        shortcuts_action.triggered.connect(self._show_shortcuts)
        help_menu.addAction(shortcuts_action)
        
        help_menu.addSeparator()
        
        about_action = QAction("关于(&A)...", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)
    
    def _setup_toolbar(self):
        """Set up toolbar"""
        toolbar = QToolBar("Main Toolbar")
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(24, 24))
        self.addToolBar(toolbar)
        
        # Add branding with logo
        brand_widget = QWidget()
        brand_layout = QHBoxLayout(brand_widget)
        brand_layout.setContentsMargins(10, 0, 20, 0)
        brand_layout.setSpacing(8)
        
        # Logo
        logo_label = QLabel()
        logo_path = get_resource_path("assets/logo.png")
        if logo_path.exists():
            pixmap = QPixmap(str(logo_path))
            scaled_pixmap = pixmap.scaled(32, 32, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(scaled_pixmap)
        else:
            logo_label.setText("◆")
            logo_label.setStyleSheet("color: #db0011; font-size: 20px;")
        brand_layout.addWidget(logo_label)
        
        # Brand name (save as class attribute for theme switching)
        self.brand_name_label = QLabel("Excel Workflow Tool")
        self._update_brand_style(dark=True)
        brand_layout.addWidget(self.brand_name_label)
        
        toolbar.addWidget(brand_widget)
        toolbar.addSeparator()
        
        # New button
        new_btn = QPushButton("📄 新建")
        new_btn.setToolTip("新建工作流 (Ctrl+N)")
        new_btn.clicked.connect(self._new_workflow)
        toolbar.addWidget(new_btn)
        
        # Open button
        open_btn = QPushButton("📂 打开")
        open_btn.setToolTip("打开工作流 (Ctrl+O)")
        open_btn.clicked.connect(self._open_workflow)
        toolbar.addWidget(open_btn)
        
        # Save button
        save_btn = QPushButton("💾 保存")
        save_btn.setToolTip("保存工作流 (Ctrl+S)")
        save_btn.clicked.connect(self._save_workflow)
        toolbar.addWidget(save_btn)
        
        # Reload button
        reload_btn = QPushButton("🔄 刷新")
        reload_btn.setToolTip("刷新当前工作流 (Ctrl+R)")
        reload_btn.clicked.connect(self._reload_workflow)
        toolbar.addWidget(reload_btn)
        
        toolbar.addSeparator()
        
        # Undo button
        self.undo_btn = QPushButton("↩ 撤销")
        self.undo_btn.setToolTip("撤销 (Ctrl+Z)")
        self.undo_btn.clicked.connect(self._undo)
        self.undo_btn.setEnabled(False)
        toolbar.addWidget(self.undo_btn)
        
        # Redo button
        self.redo_btn = QPushButton("↪ 重做")
        self.redo_btn.setToolTip("重做 (Ctrl+Y)")
        self.redo_btn.clicked.connect(self._redo)
        self.redo_btn.setEnabled(False)
        toolbar.addWidget(self.redo_btn)
        
        toolbar.addSeparator()
        
        # Run button
        run_btn = QPushButton("▶️ 执行")
        run_btn.setStyleSheet("QPushButton { background-color: #22c55e; color: white; font-weight: bold; }")
        run_btn.clicked.connect(self._execute_workflow)
        toolbar.addWidget(run_btn)
        
        toolbar.addSeparator()
        
        # Zoom controls
        zoom_in_btn = QPushButton("🔍 放大")
        zoom_in_btn.clicked.connect(lambda: self.canvas.zoom(1.2))
        toolbar.addWidget(zoom_in_btn)
        
        zoom_out_btn = QPushButton("🔍 缩小")
        zoom_out_btn.clicked.connect(lambda: self.canvas.zoom(0.8))
        toolbar.addWidget(zoom_out_btn)
        
        fit_btn = QPushButton("⊡ 适应")
        fit_btn.clicked.connect(self.canvas.fit_to_view)
        toolbar.addWidget(fit_btn)
        
        toolbar.addSeparator()
        
        # Theme toggle button
        self.theme_btn = QPushButton("🌙 浅色")
        self.theme_btn.setToolTip("切换浅色/深色模式 (Ctrl+T)")
        self.theme_btn.clicked.connect(self._toggle_theme)
        toolbar.addWidget(self.theme_btn)
    
    def _setup_statusbar(self):
        """Set up status bar"""
        self.statusbar = QStatusBar()
        self.setStatusBar(self.statusbar)
        self.statusbar.showMessage("就绪")

    def _append_run_log(self, line: str):
        """追加一行到执行日志（带时间戳）。"""
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_panel.appendPlainText(f"[{ts}] {line}")
        sb = self.log_panel.verticalScrollBar()
        sb.setValue(sb.maximum())

    def _register_recent_node_type(self, node_type: str):
        """记录最近使用的节点类型，供左侧面板快捷选取。"""
        if not node_type:
            return
        recent = self._settings.value("recent_node_types", []) or []
        if not isinstance(recent, list):
            recent = []
        if node_type in recent:
            recent.remove(node_type)
        recent.insert(0, node_type)
        self._settings.setValue("recent_node_types", recent[:12])
        self.node_palette.refresh_recent_bar()
    
    def _apply_dark_theme(self):
        """Apply dark theme to the application"""
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #1e1e1e;
                color: #e0e0e0;
            }
            QMenuBar {
                background-color: #2d2d2d;
                color: #e0e0e0;
            }
            QMenuBar::item:selected {
                background-color: #3d3d3d;
            }
            QMenu {
                background-color: #2d2d2d;
                color: #e0e0e0;
                border: 1px solid #3d3d3d;
            }
            QMenu::item:selected {
                background-color: #4d4d4d;
            }
            QToolBar {
                background-color: #2d2d2d;
                border: none;
                spacing: 5px;
                padding: 5px;
            }
            QPushButton {
                background-color: #3d3d3d;
                color: #e0e0e0;
                border: 1px solid #4d4d4d;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #4d4d4d;
            }
            QPushButton:pressed {
                background-color: #5d5d5d;
            }
            QDockWidget {
                color: #e0e0e0;
                titlebar-close-icon: none;
            }
            QDockWidget::title {
                background-color: #2d2d2d;
                padding: 5px;
            }
            QListWidget {
                background-color: #252525;
                color: #e0e0e0;
                border: 1px solid #3d3d3d;
            }
            QListWidget::item {
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: #4d4d4d;
            }
            QListWidget::item:hover {
                background-color: #3d3d3d;
            }
            QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox, QComboBox {
                background-color: #2d2d2d;
                color: #e0e0e0;
                border: 1px solid #3d3d3d;
                padding: 5px;
            }
            QPlainTextEdit {
                font-family: Consolas, 'Cascadia Mono', monospace;
                font-size: 11px;
            }
            QLabel {
                color: #e0e0e0;
            }
            QStatusBar {
                background-color: #2d2d2d;
                color: #888888;
            }
            QTableWidget {
                background-color: #252525;
                color: #e0e0e0;
                gridline-color: #3d3d3d;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #2d2d2d;
                color: #e0e0e0;
                padding: 5px;
                border: 1px solid #3d3d3d;
            }
            QScrollBar:vertical {
                background-color: #2d2d2d;
                width: 12px;
            }
            QScrollBar::handle:vertical {
                background-color: #4d4d4d;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar:horizontal {
                background-color: #2d2d2d;
                height: 12px;
            }
            QScrollBar::handle:horizontal {
                background-color: #4d4d4d;
                border-radius: 6px;
                min-width: 20px;
            }
            QCheckBox {
                color: #e0e0e0;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
        """)
    
    def _apply_light_theme(self):
        """Apply light theme to the application"""
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: #f5f5f5;
                color: #333333;
            }
            QMenuBar {
                background-color: #ffffff;
                color: #333333;
                border-bottom: 1px solid #e0e0e0;
            }
            QMenuBar::item:selected {
                background-color: #e8e8e8;
            }
            QMenu {
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #d0d0d0;
            }
            QMenu::item:selected {
                background-color: #e8e8e8;
            }
            QToolBar {
                background-color: #ffffff;
                border: none;
                border-bottom: 1px solid #e0e0e0;
                spacing: 5px;
                padding: 5px;
            }
            QPushButton {
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #c0c0c0;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #e8e8e8;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
            QDockWidget {
                color: #333333;
                titlebar-close-icon: none;
            }
            QDockWidget::title {
                background-color: #f0f0f0;
                padding: 5px;
            }
            QListWidget {
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #d0d0d0;
            }
            QListWidget::item {
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: #cce5ff;
            }
            QListWidget::item:hover {
                background-color: #e8f4ff;
            }
            QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox, QComboBox {
                background-color: #ffffff;
                color: #333333;
                border: 1px solid #c0c0c0;
                padding: 5px;
            }
            QPlainTextEdit {
                font-family: Consolas, 'Cascadia Mono', monospace;
                font-size: 11px;
            }
            QLabel {
                color: #333333;
            }
            QStatusBar {
                background-color: #f0f0f0;
                color: #666666;
            }
            QTableWidget {
                background-color: #ffffff;
                color: #333333;
                gridline-color: #e0e0e0;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                color: #333333;
                padding: 5px;
                border: 1px solid #d0d0d0;
            }
            QScrollBar:vertical {
                background-color: #f0f0f0;
                width: 12px;
            }
            QScrollBar::handle:vertical {
                background-color: #c0c0c0;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar:horizontal {
                background-color: #f0f0f0;
                height: 12px;
            }
            QScrollBar::handle:horizontal {
                background-color: #c0c0c0;
                border-radius: 6px;
                min-width: 20px;
            }
            QCheckBox {
                color: #333333;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
        """)
    
    def _toggle_theme(self):
        """Toggle between light and dark theme"""
        self._is_dark_theme = not self._is_dark_theme
        
        if self._is_dark_theme:
            self._apply_dark_theme()
            self.canvas.set_theme(dark=True)
            self._update_brand_style(dark=True)
            self.theme_btn.setText("🌙 浅色")
            self.theme_action.setText("🌙 切换到浅色模式")
            self.statusbar.showMessage("已切换到深色模式")
        else:
            self._apply_light_theme()
            self.canvas.set_theme(dark=False)
            self._update_brand_style(dark=False)
            self.theme_btn.setText("☀️ 深色")
            self.theme_action.setText("☀️ 切换到深色模式")
            self.statusbar.showMessage("已切换到浅色模式")
    
    def _update_brand_style(self, dark: bool = True):
        """Update brand label style based on theme"""
        if dark:
            self.brand_name_label.setStyleSheet("""
                QLabel {
                    color: #ffffff;
                    font-size: 16px;
                    font-weight: bold;
                    font-family: 'Segoe UI', Arial, sans-serif;
                }
            """)
        else:
            self.brand_name_label.setStyleSheet("""
                QLabel {
                    color: #333333;
                    font-size: 16px;
                    font-weight: bold;
                    font-family: 'Segoe UI', Arial, sans-serif;
                }
            """)
    
    def _toggle_minimap(self):
        """Toggle minimap visibility"""
        self.canvas.toggle_minimap()
        self.statusbar.showMessage("小地图: " + ("显示" if self.canvas._show_minimap else "隐藏"))
    
    def _on_recent_node_quick_add(self, node_type: str):
        """从「最近使用」一键在画布中央添加节点。"""
        try:
            self._save_state()
            node = self.workflow.add_node(
                node_type,
                (max(40, self.canvas.width() // 2), max(40, self.canvas.height() // 2)),
            )
            self.canvas.update()
            self.statusbar.showMessage(f"已添加节点: {node.node_name}")
            self._register_recent_node_type(node_type)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"添加节点失败: {e}")

    def _on_palette_item_double_clicked(self, item):
        """Handle double-click on palette item"""
        try:
            if isinstance(item, NodeListItem):
                # Save state before adding node
                self._save_state()
                # Add node to canvas at center
                node = self.workflow.add_node(
                    item.node_type, 
                    (self.canvas.width() // 2, self.canvas.height() // 2)
                )
                self.canvas.update()
                self.statusbar.showMessage(f"已添加节点: {node.node_name}")
                self._register_recent_node_type(item.node_type)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"添加节点失败: {e}")
    
    def _on_node_dropped(self, node_type: str, x: int, y: int):
        """Handle node dropped from palette onto canvas"""
        try:
            # Save state before adding node
            self._save_state()
            # Add node at drop position
            node = self.workflow.add_node(node_type, (x, y))
            self.canvas.update()
            self.statusbar.showMessage(f"已添加节点: {node.node_name}")
            self._register_recent_node_type(node_type)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"添加节点失败: {e}")
    
    def _on_node_selected(self, node_id: str):
        """Handle node selection"""
        if node_id and node_id in self.workflow.nodes:
            node = self.workflow.nodes[node_id]
            self.config_panel.set_node(node)
            self.statusbar.showMessage(f"已选中: {node.node_name}")
        else:
            self.config_panel.clear()
    
    def _on_node_double_clicked(self, node_id: str):
        """Handle node double-click"""
        # Show config panel if hidden
        self.config_dock.show()
    
    def _on_connection_created(self):
        """Handle connection creation - save state for undo"""
        self._save_state()
        self.statusbar.showMessage("已创建连接")
    
    def _on_config_changed(self):
        """Handle configuration change"""
        self.canvas.update()
    
    def _new_workflow(self):
        """Create a new workflow"""
        if self._confirm_discard():
            self.workflow = Workflow()
            self.canvas.set_workflow(self.workflow)
            self.config_panel.set_workflow(self.workflow) # Update workflow ref
            self.config_panel.clear()
            self.preview_panel.clear()
            self.current_file = None
            self.setWindowTitle("Excel 工作流工具 - 新工作流")
            self.statusbar.showMessage("已创建新工作流")
    
    def _open_workflow(self):
        """Open an existing workflow"""
        if not self._confirm_discard():
            return
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, "打开工作流",
            "", "工作流文件 (*.workflow.json);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                self.workflow = Workflow.load(file_path)
                self.canvas.set_workflow(self.workflow)
                self.config_panel.set_workflow(self.workflow) # Update workflow ref
                self.config_panel.clear()
                self.preview_panel.clear()
                self.current_file = file_path
                self.setWindowTitle(f"Excel 工作流工具 - {Path(file_path).name}")
                self.statusbar.showMessage(f"已打开: {file_path}")
                self._add_to_recent_files(file_path)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"打开工作流失败:\n{e}")
    
    def _save_workflow(self):
        """Save the current workflow"""
        if self.current_file:
            try:
                self.workflow.save(self.current_file)
                self.statusbar.showMessage(f"已保存: {self.current_file}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存工作流失败:\n{e}")
        else:
            self._save_workflow_as()
    
    def _save_workflow_as(self):
        """Save the workflow with a new name"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存工作流",
            "workflow.workflow.json", "工作流文件 (*.workflow.json);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                self.workflow.save(file_path)
                self.current_file = file_path
                self.setWindowTitle(f"Excel 工作流工具 - {Path(file_path).name}")
                self.statusbar.showMessage(f"已保存: {file_path}")
                self._add_to_recent_files(file_path)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存工作流失败:\n{e}")
    
    def _reload_workflow(self):
        """Reload the current workflow from file or reset canvas"""
        if self.current_file and Path(self.current_file).exists():
            # Reload from file
            try:
                self.workflow = Workflow.load(self.current_file)
                self.canvas.set_workflow(self.workflow)
                self.config_panel.clear()
                self.preview_panel.clear()
                # Clear undo/redo history
                self._undo_stack.clear()
                self._redo_stack.clear()
                self._update_undo_redo_buttons()
                self.statusbar.showMessage(f"已刷新: {self.current_file}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"刷新工作流失败:\n{e}")
        else:
            # Just refresh the canvas view
            self.canvas.update()
            self.canvas.fit_to_view()
            self.config_panel.clear()
            self.statusbar.showMessage("已刷新画布")
    
    def _delete_selected(self):
        """Delete the selected node"""
        if self.canvas.selected_node:
            # Save state before deleting
            self._save_state()
            self.workflow.remove_node(self.canvas.selected_node)
            self.canvas.selected_node = None
            self.canvas.update()
            self.config_panel.clear()
            self.statusbar.showMessage("节点已删除")
    
    def _on_node_delete_requested(self, node_id: str):
        """Handle node deletion from context menu"""
        if node_id in self.workflow.nodes:
            self._save_state()
            self.workflow.remove_node(node_id)
            self.canvas.selected_node = None
            self.canvas.update()
            self.config_panel.clear()
            self.statusbar.showMessage("节点已删除")
    
    def _on_node_copy_requested(self, node_id: str):
        """Handle node copy from context menu"""
        if node_id in self.workflow.nodes:
            self._save_state()
            source_node = self.workflow.nodes[node_id]
            # Create a new node of the same type
            new_node = self.workflow.add_node(
                source_node.node_type,
                (source_node.x + 50, source_node.y + 50)
            )
            # Copy config
            new_node.config = copy.deepcopy(source_node.config)
            self.canvas.update()
            self.statusbar.showMessage(f"已复制节点: {source_node.node_name}")
    
    def _generate_system_params(self) -> Dict[str, str]:
        """Generate system placeholder parameters"""
        now = datetime.now()
        params = {
            "DATE": now.strftime("%Y-%m-%d"),
            "TIME": now.strftime("%H:%M:%S"),
            "YEAR": now.strftime("%Y"),
            "MONTH": now.strftime("%m"),
            "DAY": now.strftime("%d"),
            "TIMESTAMP": now.strftime("%Y%m%d_%H%M%S"),
            "UUID": str(uuid.uuid4()),
            "DESKTOP": str(Path.home() / "Desktop"),
            "DOCUMENTS": str(Path.home() / "Documents"),
            "DOWNLOADS": str(Path.home() / "Downloads"),
        }
        return params

    def _execute_workflow(self):
        """Execute the workflow in background thread"""
        if not self.workflow.nodes:
            QMessageBox.information(self, "提示", "工作流中没有节点可执行。")
            return
        
        if self._is_executing:
            QMessageBox.warning(self, "警告", "工作流正在执行中，请等待完成。")
            return
        
        # 1. Open Global Params Dialog
        dialog = GlobalParamsDialog(self.workflow, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            self.statusbar.showMessage("执行已取消")
            return
            
        # Update user params
        self.workflow.global_params = dialog.get_params()
        
        self.log_panel.clear()
        self._append_run_log("开始执行工作流…")

        self.statusbar.showMessage("正在执行工作流...")
        
        # Generate system params
        system_params = self._generate_system_params()
        
        # Set all nodes to pending
        for node_id in self.workflow.nodes:
            self.canvas.set_node_status(node_id, 'pending')
        
        # Start animation
        self.canvas.start_animation()
        
        # Mark as executing
        self._is_executing = True
        
        # Create and configure worker thread
        self._worker = WorkflowWorker(self.workflow, external_context=system_params)
        self._worker.progress_updated.connect(self._on_execution_progress)
        self._worker.execution_finished.connect(self._on_execution_finished)
        self._worker.execution_failed.connect(self._on_execution_failed)
        
        # Start execution in background
        self._worker.start()
    
    def _on_execution_progress(self, current: int, total: int, node_name: str, node_id: str, detail_msg: str):
        """Handle progress updates from worker thread"""
        msg = f"正在执行: {node_name} ({current}/{total})"
        if detail_msg:
            msg += f" - {detail_msg}"
        self.statusbar.showMessage(msg)
        log_line = f"{node_name} ({current}/{total})"
        if detail_msg:
            log_line += f" — {detail_msg}"
        self._append_run_log(log_line)
        
        if node_id:
            # Set previous running nodes to success before setting new one to running
            for nid in self.workflow.nodes:
                if self.canvas.node_status.get(nid) == 'running':
                    self.canvas.set_node_status(nid, 'success')
            self.canvas.set_node_status(node_id, 'running')
    
    def _on_execution_finished(self, results: Dict[str, Any]):
        """Handle successful execution completion"""
        self._is_executing = False
        
        # Update node status based on results
        for node_id, result in results.items():
            if result["success"]:
                self.canvas.set_node_status(node_id, 'success')
            else:
                self.canvas.set_node_status(node_id, 'error')
        
        # Show results in preview panel
        last_output = None
        for node_id, result in results.items():
            if result["success"] and result.get("output"):
                for port_name, data in result["output"].items():
                    last_output = data
        
        if last_output is not None:
            self.preview_panel.set_data(last_output)
        
        # Stop animation
        self.canvas.stop_animation()
        
        self.statusbar.showMessage("工作流执行成功！")
        self._append_run_log("工作流执行成功。")
        QMessageBox.information(self, "成功", "工作流执行成功！")
    
    def _on_execution_failed(self, error_msg: str):
        """Handle execution failure"""
        self._is_executing = False
        
        self.canvas.stop_animation()
        self._append_run_log(f"错误: {error_msg}")

        self.canvas.clear_node_status()
        m = re.search(r"Error in node '([^']+)'", error_msg)
        if m:
            failed_name = m.group(1)
            for nid, n in self.workflow.nodes.items():
                if n.node_name == failed_name:
                    self.canvas.set_node_status(nid, "error")
                    break
        
        self.statusbar.showMessage(f"执行失败: {error_msg}")
        QMessageBox.critical(self, "执行错误", error_msg)
    
    def _execute_node(self, target_node_id: str):
        """Execute the workflow up to a specific node"""
        if not target_node_id or target_node_id not in self.workflow.nodes:
            return
            
        target_node = self.workflow.nodes[target_node_id]
        self.statusbar.showMessage(f"正在执行至节点: {target_node.node_name}...")
        
        # Set relevant nodes to pending
        ancestors = self.workflow.get_ancestors(target_node_id)
        nodes_to_reset = ancestors.union({target_node_id})
        
        for node_id in nodes_to_reset:
            self.canvas.set_node_status(node_id, 'pending')
        
        # Start animation
        self.canvas.start_animation()
        QApplication.processEvents()
        
        try:
            def progress(current, total, node_name, node_id=None, detail_msg=None):
                line = f"{node_name} ({current}/{total})"
                if detail_msg:
                    line += f" — {detail_msg}"
                self._append_run_log(line)
                self.statusbar.showMessage(f"正在执行: {node_name} ({current}/{total})")
                if node_id:
                    # Set previous running nodes to success before setting new one to running
                    for nid in self.workflow.nodes:
                        if self.canvas.node_status.get(nid) == 'running':
                            self.canvas.set_node_status(nid, 'success')
                    self.canvas.set_node_status(node_id, 'running')
                QApplication.processEvents()
            
            results = self.workflow.execute_node(target_node_id, progress)
            
            # Update node status based on results
            for node_id, result in results.items():
                if result["success"]:
                    self.canvas.set_node_status(node_id, 'success')
                else:
                    self.canvas.set_node_status(node_id, 'error')
            
            # Show results in preview panel (show output of the target node)
            if target_node_id in results and results[target_node_id]["success"]:
                result = results[target_node_id]
                if result.get("output"):
                    for port_name, data in result["output"].items():
                        self.preview_panel.set_data(data)
                        break # Just show the first output
            
            # Stop animation
            self.canvas.stop_animation()
            
            self.statusbar.showMessage(f"节点执行成功: {target_node.node_name}")
            self._append_run_log(f"单节点执行完成: {target_node.node_name}")
            
        except Exception as e:
            # Stop animation
            self.canvas.stop_animation()
            self._append_run_log(f"单节点执行失败: {e}")
            self.statusbar.showMessage(f"执行失败: {e}")
            QMessageBox.critical(self, "执行错误", str(e))

    def _confirm_discard(self) -> bool:
        """Confirm discarding unsaved changes"""
        # For simplicity, always return True
        # In a real app, you'd track changes and prompt
        return True
    
    def _setup_branding(self):
        """Set up branding elements"""
        # Set window icon if logo exists
        logo_path = get_resource_path("assets/logo.png")
        if logo_path.exists():
            self.setWindowIcon(QIcon(str(logo_path)))
    
    def _save_state(self):
        """Save current workflow state for undo"""
        import copy
        state = {
            'nodes': copy.deepcopy([(n.node_id, n.node_type, n.node_name, n.position[0], n.position[1], copy.deepcopy(n.config)) 
                                    for n in self.workflow.nodes.values()]),
            'connections': copy.deepcopy(list(self.workflow.connections))
        }
        self._undo_stack.append(state)
        if len(self._undo_stack) > self._max_history:
            self._undo_stack.pop(0)
        self._redo_stack.clear()
        self._update_undo_redo_buttons()
    
    def _undo(self):
        """Undo the last action"""
        if not self._undo_stack:
            return
        
        import copy
        # Save current state to redo
        current_state = {
            'nodes': copy.deepcopy([(n.node_id, n.node_type, n.node_name, n.position[0], n.position[1], copy.deepcopy(n.config)) 
                                    for n in self.workflow.nodes.values()]),
            'connections': copy.deepcopy(list(self.workflow.connections))
        }
        self._redo_stack.append(current_state)
        
        # Restore previous state
        state = self._undo_stack.pop()
        self._restore_state(state)
        self._update_undo_redo_buttons()
        self.statusbar.showMessage("已撤销")
    
    def _redo(self):
        """Redo the last undone action"""
        if not self._redo_stack:
            return
        
        import copy
        # Save current state to undo
        current_state = {
            'nodes': copy.deepcopy([(n.node_id, n.node_type, n.node_name, n.position[0], n.position[1], copy.deepcopy(n.config)) 
                                    for n in self.workflow.nodes.values()]),
            'connections': copy.deepcopy(list(self.workflow.connections))
        }
        self._undo_stack.append(current_state)
        
        # Restore redo state
        state = self._redo_stack.pop()
        self._restore_state(state)
        self._update_undo_redo_buttons()
        self.statusbar.showMessage("已重做")
    
    def _restore_state(self, state: dict):
        """Restore workflow to a saved state"""
        # Clear current workflow
        self.workflow.nodes.clear()
        self.workflow.connections.clear()
        
        # Restore nodes
        for node_id, node_type, name, x, y, config in state['nodes']:
            node_class = NodeRegistry.get_node_class(node_type)
            if node_class:
                node = node_class(node_id)
                node.position = (x, y)
                node.config = config
                self.workflow.nodes[node_id] = node
        
        # Restore connections
        for conn in state['connections']:
            self.workflow.connections.append(conn)
        
        # Update canvas
        self.canvas.update()
        self.config_panel.clear()
    
    def _update_undo_redo_buttons(self):
        """Update undo/redo button states"""
        self.undo_action.setEnabled(len(self._undo_stack) > 0)
        self.redo_action.setEnabled(len(self._redo_stack) > 0)
        self.undo_btn.setEnabled(len(self._undo_stack) > 0)
        self.redo_btn.setEnabled(len(self._redo_stack) > 0)
    
    def _update_recent_menu(self):
        """Update the recent files menu"""
        self.recent_menu.clear()
        
        settings = QSettings("ExcelWorkflowTool", "Settings")
        recent_files = settings.value("recent_files", [])
        
        if not recent_files:
            no_recent = QAction("(无最近文件)", self)
            no_recent.setEnabled(False)
            self.recent_menu.addAction(no_recent)
            return
        
        for file_path in recent_files[:10]:  # Show max 10 recent files
            if Path(file_path).exists():
                action = QAction(Path(file_path).name, self)
                action.setToolTip(file_path)
                action.setData(file_path)
                action.triggered.connect(lambda checked, fp=file_path: self._open_recent_file(fp))
                self.recent_menu.addAction(action)
        
        self.recent_menu.addSeparator()
        clear_action = QAction("清除最近文件列表", self)
        clear_action.triggered.connect(self._clear_recent_files)
        self.recent_menu.addAction(clear_action)
    
    def _add_to_recent_files(self, file_path: str):
        """Add a file to the recent files list"""
        settings = QSettings("ExcelWorkflowTool", "Settings")
        recent_files = settings.value("recent_files", [])
        
        # Remove if already exists
        if file_path in recent_files:
            recent_files.remove(file_path)
        
        # Add to front
        recent_files.insert(0, file_path)
        
        # Keep only last 10
        recent_files = recent_files[:10]
        
        settings.setValue("recent_files", recent_files)
        self._update_recent_menu()
    
    def _open_recent_file(self, file_path: str):
        """Open a file from the recent files list"""
        if not Path(file_path).exists():
            QMessageBox.warning(self, "文件不存在", f"文件不存在:\n{file_path}")
            return
        
        if self._confirm_discard():
            try:
                # Load new workflow
                new_workflow = Workflow.load(file_path)
                self.workflow = new_workflow
                self.current_file = file_path
                
                # Update UI components with new workflow
                self.canvas.set_workflow(self.workflow)
                self.config_panel.set_workflow(self.workflow) # Update config panel ref
                self.config_panel.clear()
                
                self.setWindowTitle(f"Excel 工作流工具 - {Path(file_path).name}")
                self.statusbar.showMessage(f"已打开: {file_path}")
                self._add_to_recent_files(file_path)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"打开工作流失败:\n{e}")
    
    def _clear_recent_files(self):
        """Clear the recent files list"""
        settings = QSettings("ExcelWorkflowTool", "Settings")
        settings.setValue("recent_files", [])
        self._update_recent_menu()
        self.statusbar.showMessage("已清除最近文件列表")
    
    def _show_shortcuts(self):
        """Show keyboard shortcuts dialog"""
        shortcuts = """
        <h3>快捷键列表</h3>
        <table style="border-collapse: collapse; width: 100%;">
        <tr><td style="padding: 5px;"><b>Ctrl+N</b></td><td style="padding: 5px;">新建工作流</td></tr>
        <tr><td style="padding: 5px;"><b>Ctrl+O</b></td><td style="padding: 5px;">打开工作流</td></tr>
        <tr><td style="padding: 5px;"><b>Ctrl+S</b></td><td style="padding: 5px;">保存工作流</td></tr>
        <tr><td style="padding: 5px;"><b>Ctrl+Shift+S</b></td><td style="padding: 5px;">另存为</td></tr>
        <tr><td style="padding: 5px;"><b>Ctrl+R</b></td><td style="padding: 5px;">刷新工作流</td></tr>
        <tr><td style="padding: 5px;"><b>Ctrl+Z</b></td><td style="padding: 5px;">撤销</td></tr>
        <tr><td style="padding: 5px;"><b>Ctrl+Y</b></td><td style="padding: 5px;">重做</td></tr>
        <tr><td style="padding: 5px;"><b>Delete</b></td><td style="padding: 5px;">删除选中节点</td></tr>
        <tr><td style="padding: 5px;"><b>Ctrl+T</b></td><td style="padding: 5px;">切换深色/浅色主题</td></tr>
        <tr><td style="padding: 5px;"><b>F6</b></td><td style="padding: 5px;">执行工作流</td></tr>
        <tr><td style="padding: 5px;"><b>Ctrl+/</b></td><td style="padding: 5px;">显示快捷键列表</td></tr>
        <tr><td style="padding: 5px;"><b>Ctrl+Shift+R</b></td><td style="padding: 5px;">重启应用</td></tr>
        </table>
        <br>
        <h4>画布操作</h4>
        <table style="border-collapse: collapse; width: 100%;">
        <tr><td style="padding: 5px;"><b>鼠标滚轮</b></td><td style="padding: 5px;">缩放画布</td></tr>
        <tr><td style="padding: 5px;"><b>中键拖动</b></td><td style="padding: 5px;">平移画布</td></tr>
        <tr><td style="padding: 5px;"><b>双击节点</b></td><td style="padding: 5px;">打开节点配置</td></tr>
        <tr><td style="padding: 5px;"><b>拖动端口</b></td><td style="padding: 5px;">创建连接</td></tr>
        </table>
        """
        QMessageBox.information(self, "快捷键列表", shortcuts)
    
    def _show_about(self):
        """Show about dialog"""
        dialog = AboutDialog(self)
        dialog.exec()
    
    def _show_global_params(self):
        """Show global parameters dialog"""
        dialog = GlobalParamsDialog(self.workflow, self)
        if dialog.exec():
            # Update params
            self.workflow.global_params = dialog.get_params()
            self.statusbar.showMessage("全局参数已更新")
            
    def _restart_app(self):
        """Restart the application"""
        reply = QMessageBox.question(
            self, 
            "确认重启",
            "确定要重启应用程序吗？\n未保存的更改将会丢失。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            import sys
            import os
            
            # Get the path to the main script
            python = sys.executable
            script = os.path.abspath(sys.argv[0])
            
            # Start new instance
            os.execl(python, python, script, *sys.argv[1:])
    
    def closeEvent(self, event):
        """Handle window close event - auto-save workflow and settings"""
        # Save all application settings first
        self._save_settings()
        
        # Check if there are any nodes in the workflow
        if self.workflow.nodes:
            # Auto-save the workflow
            try:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                if self.current_file:
                    # Use current filename as base
                    base_name = Path(self.current_file).stem
                    auto_save_path = self._auto_save_dir / f"{base_name}_autosave_{timestamp}.workflow.json"
                else:
                    auto_save_path = self._auto_save_dir / f"untitled_autosave_{timestamp}.workflow.json"
                
                self.workflow.save(str(auto_save_path))
                
                # Keep only the last 10 auto-save files
                self._cleanup_old_autosaves()
                
                self.statusbar.showMessage(f"自动保存: {auto_save_path}")
            except Exception as e:
                # If auto-save fails, ask user if they want to continue closing
                reply = QMessageBox.warning(
                    self,
                    "自动保存失败",
                    f"无法自动保存工作流:\n{e}\n\n是否仍然关闭应用程序？",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.No:
                    event.ignore()
                    return
        
        event.accept()
    
    def _save_settings(self):
        """Save all application settings"""
        # Window geometry
        self._settings.setValue("geometry", self.saveGeometry())
        self._settings.setValue("windowState", self.saveState())
        
        # Theme
        self._settings.setValue("theme/isDark", self._is_dark_theme)
        
        # Current file
        if self.current_file:
            self._settings.setValue("lastFile", self.current_file)
        
        # Panel visibility
        self._settings.setValue("panels/nodePalette", self.node_palette.isVisible())
        self._settings.setValue("panels/configPanel", self.config_dock.isVisible())
        self._settings.setValue("panels/previewPanel", self.preview_dock.isVisible())
        
        # Canvas zoom level
        self._settings.setValue("canvas/scale", self.canvas.scale)
        
        # Sync settings to disk
        self._settings.sync()
    
    def _restore_settings(self):
        """Restore saved application settings"""
        # Window geometry
        geometry = self._settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)
        
        windowState = self._settings.value("windowState")
        if windowState:
            self.restoreState(windowState)
        
        # Theme - Force light theme as default, ignore previously saved dark theme
        saved_theme = self._settings.value("theme/isDark")
        # Reset to light theme by default
        self._is_dark_theme = False
        # Clear the old setting
        self._settings.setValue("theme/isDark", False)
        
        # Panel visibility
        palette_visible = self._settings.value("panels/nodePalette")
        if palette_visible is not None:
            visible = palette_visible == True or palette_visible == "true"
            self.node_palette.setVisible(visible)
        
        config_visible = self._settings.value("panels/configPanel")
        if config_visible is not None:
            visible = config_visible == True or config_visible == "true"
            self.config_dock.setVisible(visible)
        
        preview_visible = self._settings.value("panels/previewPanel")
        if preview_visible is not None:
            visible = preview_visible == True or preview_visible == "true"
            self.preview_dock.setVisible(visible)
        
        # Canvas scale
        saved_scale = self._settings.value("canvas/scale")
        if saved_scale is not None:
            try:
                self.canvas.scale = float(saved_scale)
            except (ValueError, TypeError):
                pass
        
        # Last opened file (optional: auto-load)
        last_file = self._settings.value("lastFile")
        if last_file and Path(last_file).exists():
            self.current_file = last_file
            # Optionally load the file automatically
            # self._load_workflow_file(last_file)
    
    def _cleanup_old_autosaves(self):
        """Keep only the last 10 auto-save files"""
        try:
            autosave_files = sorted(
                self._auto_save_dir.glob("*_autosave_*.workflow.json"),
                key=lambda f: f.stat().st_mtime,
                reverse=True
            )
            # Remove files beyond the 10 most recent
            for old_file in autosave_files[10:]:
                old_file.unlink()
        except Exception:
            pass  # Ignore cleanup errors
    
    def _setup_templates_menu(self):
        """Set up the templates menu with predefined workflow templates"""
        # Add "Save as Template" option
        save_template_action = QAction("💾 保存为模板...", self)
        save_template_action.triggered.connect(self._save_as_template)
        self.templates_menu.addAction(save_template_action)
        
        self.templates_menu.addSeparator()
        
        # Predefined templates
        templates = [
            ("📊 数据清洗模板", "data_cleaning", "读取Excel → 去重 → 填充空值 → 写入Excel"),
            ("🔗 数据合并模板", "data_merge", "读取多个Excel → 合并数据 → 写入Excel"),
            ("📈 数据分析模板", "data_analysis", "读取Excel → 分组汇总 → 数据透视表 → 写入Excel"),
            ("📁 批量处理模板", "batch_process", "批量读取文件夹 → 数据转换 → 批量写入"),
            ("✅ 数据验证模板", "data_validation", "读取Excel → 数据验证 → 分离有效/无效数据"),
        ]
        
        for name, template_id, description in templates:
            action = QAction(name, self)
            action.setStatusTip(description)
            action.triggered.connect(lambda checked, tid=template_id: self._load_template(tid))
            self.templates_menu.addAction(action)
        
        self.templates_menu.addSeparator()
        
        # User templates submenu
        self.user_templates_menu = QMenu("📂 我的模板", self)
        self._update_user_templates_menu()
        self.templates_menu.addMenu(self.user_templates_menu)
    
    def _update_user_templates_menu(self):
        """Update user templates menu"""
        self.user_templates_menu.clear()
        
        templates_dir = Path(__file__).parent.parent.parent / "templates"
        templates_dir.mkdir(exist_ok=True)
        
        template_files = list(templates_dir.glob("*.template.json"))
        
        if template_files:
            for template_file in sorted(template_files):
                name = template_file.stem.replace(".template", "")
                action = QAction(f"📄 {name}", self)
                action.triggered.connect(lambda checked, f=template_file: self._load_template_file(f))
                self.user_templates_menu.addAction(action)
        else:
            empty_action = QAction("(无保存的模板)", self)
            empty_action.setEnabled(False)
            self.user_templates_menu.addAction(empty_action)
    
    def _save_as_template(self):
        """Save current workflow as a template"""
        from PyQt6.QtWidgets import QInputDialog
        
        name, ok = QInputDialog.getText(self, "保存模板", "模板名称:")
        if ok and name:
            templates_dir = Path(__file__).parent.parent.parent / "templates"
            templates_dir.mkdir(exist_ok=True)
            
            template_file = templates_dir / f"{name}.template.json"
            
            try:
                workflow_data = self.workflow.to_dict()
                workflow_data['template_name'] = name
                workflow_data['template_description'] = f"用户创建的模板: {name}"
                
                with open(template_file, 'w', encoding='utf-8') as f:
                    json.dump(workflow_data, f, ensure_ascii=False, indent=2)
                
                self._update_user_templates_menu()
                self.statusbar.showMessage(f"模板已保存: {name}")
                QMessageBox.information(self, "成功", f"模板 '{name}' 已保存!")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存模板失败: {e}")
    
    def _load_template(self, template_id: str):
        """Load a predefined template"""
        self._save_state()
        
        # Clear current workflow
        self.workflow.nodes.clear()
        self.workflow.connections.clear()
        
        # Create template based on ID
        if template_id == "data_cleaning":
            # Data cleaning template
            read_node = self.workflow.add_node("read_excel", (100, 200))
            dedup_node = self.workflow.add_node("remove_duplicates", (350, 200))
            fill_node = self.workflow.add_node("fill_na", (600, 200))
            write_node = self.workflow.add_node("write_excel", (850, 200))
            
            self.workflow.connect(read_node.node_id, "data", dedup_node.node_id, "data")
            self.workflow.connect(dedup_node.node_id, "data", fill_node.node_id, "data")
            self.workflow.connect(fill_node.node_id, "data", write_node.node_id, "data")
            
        elif template_id == "data_merge":
            # Data merge template
            read1 = self.workflow.add_node("read_excel", (100, 100))
            read2 = self.workflow.add_node("read_excel", (100, 300))
            merge_node = self.workflow.add_node("merge_data", (400, 200))
            write_node = self.workflow.add_node("write_excel", (700, 200))
            
            self.workflow.connect(read1.node_id, "data", merge_node.node_id, "left")
            self.workflow.connect(read2.node_id, "data", merge_node.node_id, "right")
            self.workflow.connect(merge_node.node_id, "data", write_node.node_id, "data")
            
        elif template_id == "data_analysis":
            # Data analysis template
            read_node = self.workflow.add_node("read_excel", (100, 200))
            group_node = self.workflow.add_node("group_by", (350, 200))
            pivot_node = self.workflow.add_node("pivot_table", (600, 200))
            write_node = self.workflow.add_node("write_excel", (850, 200))
            
            self.workflow.connect(read_node.node_id, "data", group_node.node_id, "data")
            self.workflow.connect(group_node.node_id, "data", pivot_node.node_id, "data")
            self.workflow.connect(pivot_node.node_id, "data", write_node.node_id, "data")
            
        elif template_id == "batch_process":
            # Batch processing template
            batch_read = self.workflow.add_node("batch_read_excel", (100, 200))
            filter_node = self.workflow.add_node("filter_rows", (400, 200))
            batch_write = self.workflow.add_node("batch_write_excel", (700, 200))
            
            self.workflow.connect(batch_read.node_id, "data", filter_node.node_id, "data")
            self.workflow.connect(filter_node.node_id, "data", batch_write.node_id, "data")
            
        elif template_id == "data_validation":
            # Data validation template
            read_node = self.workflow.add_node("read_excel", (100, 200))
            validate_node = self.workflow.add_node("data_validation", (400, 200))
            write_valid = self.workflow.add_node("write_excel", (700, 100))
            write_invalid = self.workflow.add_node("write_excel", (700, 300))
            
            self.workflow.connect(read_node.node_id, "data", validate_node.node_id, "data")
            self.workflow.connect(validate_node.node_id, "valid_data", write_valid.node_id, "data")
            self.workflow.connect(validate_node.node_id, "invalid_data", write_invalid.node_id, "data")
        
        self.canvas.update()
        self.canvas.fit_to_view()
        self.statusbar.showMessage(f"已加载模板: {template_id}")
    
    def _load_template_file(self, template_file: Path):
        """Load a user-saved template file"""
        try:
            self._save_state()
            
            with open(template_file, 'r', encoding='utf-8') as f:
                template_data = json.load(f)
            
            self.workflow.from_dict(template_data)
            self.canvas.update()
            self.canvas.fit_to_view()
            
            template_name = template_data.get('template_name', template_file.stem)
            self.statusbar.showMessage(f"已加载模板: {template_name}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载模板失败: {e}")
    
    def _export_workflow(self):
        """Export workflow to a standalone file"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出工作流", "",
            "工作流文件 (*.workflow.json);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                if not file_path.endswith('.workflow.json'):
                    file_path += '.workflow.json'
                
                workflow_data = self.workflow.to_dict()
                workflow_data['export_info'] = {
                    'app_version': '1.0.0',
                    'export_date': datetime.now().isoformat(),
                    'node_count': len(self.workflow.nodes),
                    'connection_count': len(self.workflow.connections)
                }
                
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(workflow_data, f, ensure_ascii=False, indent=2)
                
                self.statusbar.showMessage(f"工作流已导出: {file_path}")
                QMessageBox.information(self, "成功", f"工作流已成功导出到:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {e}")
    
    def _import_workflow(self):
        """Import workflow from a file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "导入工作流", "",
            "工作流文件 (*.workflow.json *.json);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                self._save_state()
                
                with open(file_path, 'r', encoding='utf-8') as f:
                    workflow_data = json.load(f)
                
                self.workflow.from_dict(workflow_data)
                self.canvas.update()
                self.canvas.fit_to_view()
                
                # Show import info if available
                export_info = workflow_data.get('export_info', {})
                node_count = export_info.get('node_count', len(self.workflow.nodes))
                
                self.statusbar.showMessage(f"已导入工作流: {node_count} 个节点")
                QMessageBox.information(self, "成功", f"成功导入工作流!\n节点数: {node_count}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导入失败: {e}")
    
    def _export_as_image(self):
        """Export canvas as image"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出为图片", "workflow",
            "PNG图片 (*.png);;JPEG图片 (*.jpg);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                # Ensure file has extension
                if not any(file_path.endswith(ext) for ext in ['.png', '.jpg', '.jpeg']):
                    file_path += '.png'
                
                # Create a pixmap of the canvas
                pixmap = QPixmap(self.canvas.size())
                self.canvas.render(pixmap)
                
                # Save the pixmap
                pixmap.save(file_path)
                
                self.statusbar.showMessage(f"图片已保存: {file_path}")
                QMessageBox.information(self, "成功", f"工作流图片已保存到:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出图片失败: {e}")

