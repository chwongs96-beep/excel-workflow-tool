"""
Splash Screen - Premium startup screen with GIF animation
"""

from PyQt6.QtWidgets import QWidget, QLabel, QVBoxLayout, QProgressBar, QApplication
from PyQt6.QtCore import Qt, QTimer, QSize
from PyQt6.QtGui import QPixmap, QPainter, QColor, QFont, QMovie, QPainterPath, QBrush, QPen

from pathlib import Path
from src.utils import get_resource_path


class SplashScreen(QWidget):
    """Premium splash screen with branding and GIF animation"""
    
    def __init__(self):
        super().__init__()
        
        self.width = 500
        self.height = 480
        
        self.setWindowFlags(
            Qt.WindowType.WindowStaysOnTopHint | 
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.SplashScreen
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setFixedSize(self.width, self.height)
        
        self._progress = 0
        self._message = "正在初始化..."
        
        self._setup_ui()
        self._center_on_screen()
    
    def _center_on_screen(self):
        """Center the splash screen on the primary screen"""
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - self.width) // 2
        y = (screen.height() - self.height) // 2
        self.move(x, y)
    
    def _setup_ui(self):
        """Set up the UI components"""
        # Main layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 50, 40, 40)
        layout.setSpacing(12)
        
        # GIF animation label
        self.gif_label = QLabel()
        self.gif_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.gif_label.setFixedSize(280, 200)
        self.gif_label.setStyleSheet("background: transparent;")
        
        # Try to load splash image (GIF or PNG)
        assets_path = get_resource_path("assets")
        gif_path = assets_path / "splash.gif"
        png_path = assets_path / "splash.png"
        
        if gif_path.exists():
            # Animated GIF
            self.movie = QMovie(str(gif_path))
            self.movie.setScaledSize(QSize(280, 200))
            self.gif_label.setMovie(self.movie)
            self.movie.start()
        elif png_path.exists():
            # Static PNG splash image
            pixmap = QPixmap(str(png_path)).scaled(
                320, 220, 
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            self.gif_label.setPixmap(pixmap)
        else:
            # Fallback: use logo.png
            logo_path = assets_path / "logo.png"
            if logo_path.exists():
                pixmap = QPixmap(str(logo_path)).scaled(
                    180, 180, 
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation
                )
                self.gif_label.setPixmap(pixmap)
        
        layout.addWidget(self.gif_label, alignment=Qt.AlignmentFlag.AlignCenter)
        
        # Brand name
        self.brand_label = QLabel("Excel Workflow Tool")
        self.brand_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.brand_label.setStyleSheet("""
            QLabel {
                color: #1a1a2e;
                font-size: 28px;
                font-weight: bold;
                font-family: 'Segoe UI', Arial, sans-serif;
                letter-spacing: 2px;
                background: transparent;
            }
        """)
        layout.addWidget(self.brand_label)
        
        # Tagline
        self.tagline_label = QLabel("Excel 工作流自动化工具")
        self.tagline_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.tagline_label.setStyleSheet("""
            QLabel {
                color: #555555;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
                background: transparent;
            }
        """)
        layout.addWidget(self.tagline_label)
        
        # Spacer
        layout.addSpacing(15)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(6)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                background-color: #e0e0e0;
                border: none;
                border-radius: 3px;
            }
            QProgressBar::chunk {
                background-color: #db0011;
                border-radius: 3px;
            }
        """)
        layout.addWidget(self.progress_bar)
        
        # Status message
        self.status_label = QLabel(self._message)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("""
            QLabel {
                color: #666666;
                font-size: 11px;
                font-family: 'Segoe UI', Arial, sans-serif;
                background: transparent;
            }
        """)
        layout.addWidget(self.status_label)
        
        layout.addSpacing(5)
        
        # Copyright
        self.copyright_label = QLabel("© 2025 Excel Workflow Tool. All rights reserved.")
        self.copyright_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.copyright_label.setStyleSheet("""
            QLabel {
                color: #999999;
                font-size: 9px;
                font-family: 'Segoe UI', Arial, sans-serif;
                background: transparent;
            }
        """)
        layout.addWidget(self.copyright_label)
    
    def paintEvent(self, event):
        """Paint the background"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Draw shadow
        for i in range(8, 0, -1):
            alpha = int(25 - i * 3)
            color = QColor(0, 0, 0, alpha)
            painter.setPen(QPen(color, 1))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRoundedRect(
                20 - i, 20 - i,
                self.width - 40 + i * 2, self.height - 40 + i * 2,
                20, 20
            )
        
        # Draw main background
        path = QPainterPath()
        path.addRoundedRect(20, 20, self.width - 40, self.height - 40, 20, 20)
        
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QBrush(QColor("#ffffff")))
        painter.drawPath(path)
        
        # Draw subtle border
        painter.setPen(QPen(QColor(220, 220, 220), 1))
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRoundedRect(20, 20, self.width - 40, self.height - 40, 20, 20)
        
        # Draw red accent line under brand name
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QBrush(QColor("#db0011")))
        line_width = 100
        line_x = (self.width - line_width) // 2
        # Position after GIF (200) + brand label area
        painter.drawRoundedRect(line_x, 310, line_width, 3, 1, 1)
    
    def set_progress(self, value: int, message: str = ""):
        """Update progress bar and message"""
        self._progress = min(100, max(0, value))
        self.progress_bar.setValue(self._progress)
        if message:
            self._message = message
            self.status_label.setText(message)
        QApplication.processEvents()
    
    def finish(self, widget):
        """Finish splash screen (compatibility method)"""
        pass
    
    def close(self):
        """Stop movie when closing"""
        if hasattr(self, 'movie'):
            self.movie.stop()
        super().close()


def show_splash_and_load(app, main_window_class):
    """Show splash screen while loading the application"""
    splash = SplashScreen()
    splash.show()
    app.processEvents()
    
    # Loading steps with more detailed messages
    steps = [
        (5, "正在初始化应用程序..."),
        (15, "正在加载核心模块..."),
        (25, "正在初始化节点注册表..."),
        (40, "正在加载 Excel 处理引擎..."),
        (55, "正在准备用户界面组件..."),
        (70, "正在加载主题和样式..."),
        (85, "正在初始化工作流引擎..."),
        (95, "正在完成最后准备..."),
        (100, "启动完成！"),
    ]
    
    import time
    for progress, message in steps:
        splash.set_progress(progress, message)
        app.processEvents()
        time.sleep(0.15)
    
    # Small delay at 100% to show completion
    time.sleep(0.3)
    
    # Create and show main window
    window = main_window_class()
    window.show()
    
    # Close splash
    splash.close()
    
    return window
