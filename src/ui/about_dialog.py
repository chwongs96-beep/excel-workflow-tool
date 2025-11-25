"""
About Dialog - Professional about dialog with branding information
"""

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
    QPushButton, QFrame, QWidget
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap, QFont

from pathlib import Path


class AboutDialog(QDialog):
    """Professional about dialog with IBCN Finance branding"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        
        self.setWindowTitle("关于 IBCN Finance Excel 工作流工具")
        self.setFixedSize(450, 400)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)
        
        self._setup_ui()
        self._apply_style()
    
    def _setup_ui(self):
        """Set up the dialog UI"""
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Logo
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_path = Path(__file__).parent.parent.parent / "assets" / "logo.png"
        if logo_path.exists():
            pixmap = QPixmap(str(logo_path))
            scaled = pixmap.scaled(80, 80, Qt.AspectRatioMode.KeepAspectRatio,
                                   Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(scaled)
        else:
            logo_label.setText("◆")
            logo_label.setStyleSheet("font-size: 48px; color: #db0011;")
        layout.addWidget(logo_label)
        
        # App name
        name_label = QLabel("IBCN Finance")
        name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        name_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #ffffff;")
        layout.addWidget(name_label)
        
        # Subtitle
        subtitle_label = QLabel("Excel 工作流自动化工具")
        subtitle_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle_label.setStyleSheet("font-size: 14px; color: #888888;")
        layout.addWidget(subtitle_label)
        
        # Version
        version_label = QLabel("版本 1.0.0")
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version_label.setStyleSheet("font-size: 12px; color: #666666;")
        layout.addWidget(version_label)
        
        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("background-color: #3d3d3d;")
        layout.addWidget(line)
        
        # Description
        desc_label = QLabel(
            "一款专业的可视化Excel数据处理工具。\n"
            "通过拖拽节点构建数据处理工作流，\n"
            "无需编程即可完成复杂的Excel自动化任务。"
        )
        desc_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("font-size: 12px; color: #aaaaaa; line-height: 1.5;")
        layout.addWidget(desc_label)
        
        # Features
        features_label = QLabel(
            "✓ 40+ 数据处理节点\n"
            "✓ 可视化工作流设计\n"
            "✓ 支持多种Excel格式\n"
            "✓ 深色/浅色主题"
        )
        features_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        features_label.setStyleSheet("font-size: 11px; color: #22c55e;")
        layout.addWidget(features_label)
        
        layout.addStretch()
        
        # Copyright
        copyright_label = QLabel("© 2025 IBCN Finance. All rights reserved.")
        copyright_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        copyright_label.setStyleSheet("font-size: 10px; color: #555555;")
        layout.addWidget(copyright_label)
        
        # Close button
        close_btn = QPushButton("确定")
        close_btn.setFixedWidth(100)
        close_btn.clicked.connect(self.accept)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
    
    def _apply_style(self):
        """Apply dialog styling"""
        self.setStyleSheet("""
            QDialog {
                background-color: #1e1e1e;
            }
            QPushButton {
                background-color: #db0011;
                color: white;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #ff1a2d;
            }
            QPushButton:pressed {
                background-color: #b0000e;
            }
        """)
