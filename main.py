"""
Excel Workflow Tool - A simple offline n8n-like tool for Excel processing
No API, no database - just drag-and-drop Excel workflow automation

Excel Workflow Tool - Excel 工作流自动化工具
"""

import sys
import os
from pathlib import Path

# Ensure the root directory is in sys.path so we can import 'src'
root_dir = os.path.dirname(os.path.abspath(__file__))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt

# Import using full package path to ensure PyInstaller detects dependencies correctly
try:
    from src.ui.main_window import MainWindow
    from src.ui.splash_screen import show_splash_and_load
except ImportError as e:
    # Fallback for development environment where src might be in path differently
    try:
        sys.path.insert(0, os.path.join(root_dir, 'src'))
        from ui.main_window import MainWindow
        from ui.splash_screen import show_splash_and_load
    except ImportError:
        raise e

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
    
    # Show splash screen and load main window
    window = show_splash_and_load(app, MainWindow)
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
