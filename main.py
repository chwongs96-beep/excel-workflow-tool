"""
Excel Workflow Tool - A simple offline n8n-like tool for Excel processing
No API, no database - just drag-and-drop Excel workflow automation

Excel Workflow Tool - Excel 工作流自动化工具
"""

import sys
import os
from pathlib import Path

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt


def main():
    # Enable High DPI scaling
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setApplicationName("Excel 工作流工具")
    app.setOrganizationName("ExcelWorkflowTool")

    # Set style
    app.setStyle("Fusion")
    
    # Import here to avoid import issues
    from ui.main_window import MainWindow
    from ui.splash_screen import show_splash_and_load
    
    # Show splash screen and load main window
    window = show_splash_and_load(app, MainWindow)
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
