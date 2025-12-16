"""
Workflow Canvas - visual node editor
"""

from typing import Optional, Tuple, Dict, List
from PyQt6.QtWidgets import QWidget, QApplication, QMenu
from PyQt6.QtCore import Qt, QPoint, QPointF, QRectF, pyqtSignal, QTimer
from PyQt6.QtGui import (
    QPainter, QPen, QBrush, QColor, QFont, 
    QMouseEvent, QPaintEvent, QWheelEvent, QFontMetrics,
    QPainterPath, QLinearGradient, QPixmap, QAction
)

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from src.workflow.engine import Workflow
from src.nodes.base_node import BaseNode, PortType


class WorkflowCanvas(QWidget):
    """Widget for visual workflow editing"""
    
    node_selected = pyqtSignal(str)  # Emits node_id
    node_double_clicked = pyqtSignal(str)  # Emits node_id
    connection_created = pyqtSignal()  # Emits when a connection is created
    node_delete_requested = pyqtSignal(str)  # Emits node_id for deletion
    node_copy_requested = pyqtSignal(str)  # Emits node_id for copy
    node_execution_requested = pyqtSignal(str)  # Emits node_id for execution
    workflow_execution_requested = pyqtSignal()  # Emits for full workflow execution
    node_dropped = pyqtSignal(str, int, int)  # Emits (node_type, x, y) when node is dropped
    
    # Node visual constants
    NODE_WIDTH = 180
    NODE_HEIGHT = 80
    NODE_HEADER_HEIGHT = 28
    PORT_RADIUS = 8
    PORT_SPACING = 20
    CORNER_RADIUS = 8
    
    def __init__(self, workflow: Workflow, parent=None):
        super().__init__(parent)
        
        self.workflow = workflow
        self.selected_node: Optional[str] = None
        self.dragging_node: Optional[str] = None
        self.drag_offset = QPoint(0, 0)
        
        # Connection creation state
        self.connecting = False
        self.connection_start_node: Optional[str] = None
        self.connection_start_port: Optional[str] = None
        self.connection_start_is_output = False
        self.connection_end_pos = QPoint(0, 0)
        
        # View state
        self.scale = 1.0
        self.offset = QPoint(0, 0)
        self.panning = False
        self.pan_start = QPoint(0, 0)
        
        # Node execution status
        self.node_status: Dict[str, str] = {}  # node_id -> status (success, error, running, pending)
        
        # Animation state for connections
        self._animation_offset = 0
        self._animation_timer = QTimer(self)
        self._animation_timer.timeout.connect(self._animate_connections)
        self._is_animating = False
        
        # Enable mouse tracking
        self.setMouseTracking(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setMinimumSize(400, 300)
        
        # Enable drag and drop
        self.setAcceptDrops(True)
        
        # Theme state
        self._is_dark_theme = False  # Default to light theme
        self._bg_color_dark = "#1a1a2e"
        self._bg_color_light = "#f5f5f5"
        
        # Node colors for themes
        self._node_body_dark = "#2d2d3d"
        self._node_body_light = "#ffffff"
        self._node_border_dark = "#3d3d5c"
        self._node_border_light = "#cccccc"
        self._port_label_dark = "#aaaaaa"
        self._port_label_light = "#666666"
        self._port_border_dark = "#1a1a2e"
        self._port_border_light = "#888888"
        
        # Minimap settings
        self._show_minimap = False
        self._minimap_size = 150
        self._minimap_margin = 10
        self._minimap_dragging = False
        
        # Background
        self.setAutoFillBackground(True)
        palette = self.palette()
        palette.setColor(self.backgroundRole(), QColor(self._bg_color_light))
        self.setPalette(palette)
    
    def set_workflow(self, workflow: Workflow):
        """Set a new workflow"""
        self.workflow = workflow
        self.selected_node = None
        self.update()
    
    def set_theme(self, dark: bool = True):
        """Set the canvas theme"""
        self._is_dark_theme = dark
        palette = self.palette()
        if dark:
            palette.setColor(self.backgroundRole(), QColor(self._bg_color_dark))
        else:
            palette.setColor(self.backgroundRole(), QColor(self._bg_color_light))
        self.setPalette(palette)
        self.update()
    
    def zoom(self, factor: float):
        """Zoom the canvas"""
        new_scale = self.scale * factor
        if 0.2 <= new_scale <= 3.0:
            self.scale = new_scale
            self.update()
    
    def fit_to_view(self):
        """Fit all nodes in view"""
        if not self.workflow.nodes:
            self.scale = 1.0
            self.offset = QPoint(0, 0)
            self.update()
            return
        
        # Find bounding box
        min_x = min_y = float('inf')
        max_x = max_y = float('-inf')
        
        for node in self.workflow.nodes.values():
            x, y = node.position
            min_x = min(min_x, x)
            min_y = min(min_y, y)
            max_x = max(max_x, x + self.NODE_WIDTH)
            max_y = max(max_y, y + self.NODE_HEIGHT)
        
        # Calculate scale and offset
        bbox_width = max_x - min_x + 100
        bbox_height = max_y - min_y + 100
        
        scale_x = self.width() / bbox_width
        scale_y = self.height() / bbox_height
        self.scale = min(scale_x, scale_y, 1.5)
        
        # Center
        center_x = (min_x + max_x) / 2
        center_y = (min_y + max_y) / 2
        self.offset = QPoint(
            int(self.width() / 2 - center_x * self.scale),
            int(self.height() / 2 - center_y * self.scale)
        )
        self.update()
    
    def screen_to_canvas(self, pos: QPoint) -> QPoint:
        """Convert screen coordinates to canvas coordinates"""
        return QPoint(
            int((pos.x() - self.offset.x()) / self.scale),
            int((pos.y() - self.offset.y()) / self.scale)
        )
    
    def canvas_to_screen(self, pos: Tuple[int, int]) -> QPoint:
        """Convert canvas coordinates to screen coordinates"""
        return QPoint(
            int(pos[0] * self.scale + self.offset.x()),
            int(pos[1] * self.scale + self.offset.y())
        )
    
    def get_node_rect(self, node: BaseNode) -> QRectF:
        """Get the screen rectangle for a node"""
        x, y = node.position
        screen_pos = self.canvas_to_screen((x, y))
        return QRectF(
            screen_pos.x(), screen_pos.y(),
            self.NODE_WIDTH * self.scale,
            self.NODE_HEIGHT * self.scale
        )
    
    def get_port_pos(self, node: BaseNode, port_name: str, is_output: bool) -> QPoint:
        """Get the screen position of a port"""
        x, y = node.position
        
        if is_output:
            ports = node.outputs
            port_x = x + self.NODE_WIDTH
        else:
            ports = node.inputs
            port_x = x
        
        # Find port index
        port_idx = next((i for i, p in enumerate(ports) if p.name == port_name), 0)
        port_y = y + self.NODE_HEADER_HEIGHT + 15 + port_idx * self.PORT_SPACING
        
        return self.canvas_to_screen((port_x, port_y))
    
    def get_node_at(self, pos: QPoint) -> Optional[str]:
        """Get the node ID at the given screen position"""
        for node_id, node in self.workflow.nodes.items():
            rect = self.get_node_rect(node)
            if rect.contains(QPointF(pos)):
                return node_id
        return None
    
    def get_port_at(self, pos: QPoint) -> Optional[Tuple[str, str, bool]]:
        """Get the port at the given screen position"""
        for node_id, node in self.workflow.nodes.items():
            # Check output ports
            for port in node.outputs:
                port_pos = self.get_port_pos(node, port.name, True)
                if (pos - port_pos).manhattanLength() < self.PORT_RADIUS * self.scale * 2:
                    return (node_id, port.name, True)
            
            # Check input ports
            for port in node.inputs:
                port_pos = self.get_port_pos(node, port.name, False)
                if (pos - port_pos).manhattanLength() < self.PORT_RADIUS * self.scale * 2:
                    return (node_id, port.name, False)
        
        return None
    
    def paintEvent(self, event: QPaintEvent):
        """Paint the canvas"""
        try:
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            
            # Draw grid
            self._draw_grid(painter)
            
            # Draw watermark
            self._draw_watermark(painter)
            
            # Draw connections
            self._draw_connections(painter)
            
            # Draw connection being created
            if self.connecting and self.connection_start_node:
                node = self.workflow.nodes.get(self.connection_start_node)
                if node:
                    start = self.get_port_pos(node, self.connection_start_port, 
                                              self.connection_start_is_output)
                    self._draw_connection_line(painter, start, self.connection_end_pos, 
                                              QColor("#fbbf24"))
            
            # Draw nodes
            for node_id, node in list(self.workflow.nodes.items()):
                self._draw_node(painter, node, node_id == self.selected_node)
            
            # Draw minimap
            if self._show_minimap and self.workflow.nodes:
                self._draw_minimap(painter)
        except Exception as e:
            print(f"Paint error: {e}")
    
    def _draw_grid(self, painter: QPainter):
        """Draw background grid"""
        grid_size = int(30 * self.scale)
        if grid_size < 5:
            return
        
        # Use different grid color based on theme
        if self._is_dark_theme:
            pen = QPen(QColor("#252540"))
        else:
            pen = QPen(QColor("#d0d0d0"))
        pen.setWidth(1)
        painter.setPen(pen)
        
        # Calculate grid offset
        offset_x = self.offset.x() % grid_size
        offset_y = self.offset.y() % grid_size
        
        # Draw vertical lines
        x = offset_x
        while x < self.width():
            painter.drawLine(x, 0, x, self.height())
            x += grid_size
        
        # Draw horizontal lines
        y = offset_y
        while y < self.height():
            painter.drawLine(0, y, self.width(), y)
            y += grid_size
    
    def _draw_watermark(self, painter: QPainter):
        """Draw watermark with logo"""
        painter.save()
        
        # Choose text color based on theme
        text_color = QColor("#ffffff") if self._is_dark_theme else QColor("#333333")
        
        # Draw logo in bottom-left corner
        logo_path = Path(__file__).parent.parent.parent / "assets" / "logo.png"
        if logo_path.exists():
            pixmap = QPixmap(str(logo_path))
            # Scale logo to appropriate size
            logo_size = 48
            scaled_pixmap = pixmap.scaled(
                logo_size, logo_size,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            
            # Set opacity for the logo
            painter.setOpacity(0.3)  # 30% opacity for subtle watermark
            
            # Draw logo at bottom-left corner with padding
            logo_x = 20
            logo_y = self.height() - logo_size - 20
            painter.drawPixmap(logo_x, logo_y, scaled_pixmap)
            
            # Draw brand text next to logo
            painter.setOpacity(0.4 if not self._is_dark_theme else 0.25)
            font = QFont("Segoe UI", 14, QFont.Weight.Bold)
            painter.setFont(font)
            painter.setPen(QPen(text_color))
            
            text_x = logo_x + logo_size + 10
            text_y = logo_y + logo_size // 2 + 5
            painter.drawText(text_x, text_y, "Excel Workflow Tool")
            
            # Draw tagline below
            painter.setOpacity(0.35 if not self._is_dark_theme else 0.2)
            font2 = QFont("Segoe UI", 10)
            painter.setFont(font2)
            painter.drawText(text_x, text_y + 18, "Excel 工作流自动化")
        else:
            # Fallback: text-only watermark if logo not found
            painter.setOpacity(0.3 if not self._is_dark_theme else 0.15)
            font = QFont("Segoe UI", 16, QFont.Weight.Bold)
            painter.setFont(font)
            painter.setPen(QPen(text_color))
            painter.drawText(20, self.height() - 30, "Excel Workflow Tool")
        
        painter.restore()
    
    def _draw_node(self, painter: QPainter, node: BaseNode, selected: bool):
        """Draw a single node"""
        rect = self.get_node_rect(node)
        
        # Node colors based on theme
        color = QColor(node.node_color)
        header_color = color
        body_color = QColor(self._node_body_dark if self._is_dark_theme else self._node_body_light)
        border_color = QColor(self._node_border_dark if self._is_dark_theme else self._node_border_light)
        port_label_color = QColor(self._port_label_dark if self._is_dark_theme else self._port_label_light)
        port_border_color = QColor(self._port_border_dark if self._is_dark_theme else self._port_border_light)
        
        # Check node status for glow effect
        node_status = self.node_status.get(node.node_id, None)
        
        # Draw status glow
        if node_status:
            glow_color = None
            if node_status == 'running':
                glow_color = QColor("#fbbf24")  # Yellow
            elif node_status == 'success':
                glow_color = QColor("#22c55e")  # Green
            elif node_status == 'error':
                glow_color = QColor("#ef4444")  # Red
            elif node_status == 'pending':
                glow_color = QColor("#6b7280")  # Gray
            
            if glow_color:
                glow_color.setAlpha(100)
                painter.setBrush(Qt.BrushStyle.NoBrush)
                painter.setPen(QPen(glow_color, 4))
                glow_rect = rect.adjusted(-4, -4, 4, 4)
                painter.drawRoundedRect(glow_rect, 
                                       (self.CORNER_RADIUS + 2) * self.scale,
                                       (self.CORNER_RADIUS + 2) * self.scale)
        
        # Draw shadow if selected
        if selected:
            shadow_rect = rect.adjusted(4, 4, 4, 4)
            painter.setBrush(QBrush(QColor(0, 0, 0, 80)))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(shadow_rect, 
                                   self.CORNER_RADIUS * self.scale,
                                   self.CORNER_RADIUS * self.scale)
        
        # Draw body
        painter.setBrush(QBrush(body_color))
        if selected:
            painter.setPen(QPen(color, 2))
        else:
            painter.setPen(QPen(border_color, 1))
        painter.drawRoundedRect(rect, 
                               self.CORNER_RADIUS * self.scale,
                               self.CORNER_RADIUS * self.scale)
        
        # Draw header
        header_rect = QRectF(rect.x(), rect.y(), 
                            rect.width(), self.NODE_HEADER_HEIGHT * self.scale)
        
        path = QPainterPath()
        path.addRoundedRect(header_rect, 
                           self.CORNER_RADIUS * self.scale,
                           self.CORNER_RADIUS * self.scale)
        
        # Clip bottom corners of header
        clip_rect = QRectF(rect.x(), rect.y() + self.CORNER_RADIUS * self.scale,
                          rect.width(), self.NODE_HEADER_HEIGHT * self.scale)
        path.addRect(clip_rect)
        
        painter.setBrush(QBrush(header_color))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(header_rect,
                               self.CORNER_RADIUS * self.scale,
                               self.CORNER_RADIUS * self.scale)
        painter.drawRect(QRectF(rect.x(), 
                               rect.y() + self.CORNER_RADIUS * self.scale,
                               rect.width(), 
                               (self.NODE_HEADER_HEIGHT - self.CORNER_RADIUS) * self.scale))
        
        # Draw node name
        font = QFont("Segoe UI", int(10 * self.scale), QFont.Weight.Bold)
        painter.setFont(font)
        painter.setPen(QPen(QColor("white")))
        
        # Draw status icon in header
        status_icon = ""
        if node_status == 'running':
            status_icon = "⏳"
        elif node_status == 'success':
            status_icon = "✓"
        elif node_status == 'error':
            status_icon = "✗"
        elif node_status == 'pending':
            status_icon = "○"
        
        name_text = f"{status_icon} {node.node_name}" if status_icon else node.node_name
        
        text_rect = QRectF(rect.x() + 10 * self.scale, rect.y(),
                          rect.width() - 20 * self.scale, 
                          self.NODE_HEADER_HEIGHT * self.scale)
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignVCenter, name_text)
        
        # Draw ports
        port_font = QFont("Segoe UI", int(8 * self.scale))
        painter.setFont(port_font)
        
        # Input ports
        for i, port in enumerate(node.inputs):
            pos = self.get_port_pos(node, port.name, False)
            
            # Port circle
            painter.setBrush(QBrush(QColor("#4ade80")))
            painter.setPen(QPen(port_border_color, 2))
            painter.drawEllipse(pos, 
                              int(self.PORT_RADIUS * self.scale),
                              int(self.PORT_RADIUS * self.scale))
            
            # Port label
            painter.setPen(QPen(port_label_color))
            label_rect = QRectF(pos.x() + 12 * self.scale, 
                               pos.y() - 8 * self.scale,
                               80 * self.scale, 16 * self.scale)
            painter.drawText(label_rect, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, 
                           port.name)
        
        # Output ports
        for i, port in enumerate(node.outputs):
            pos = self.get_port_pos(node, port.name, True)
            
            # Port circle
            painter.setBrush(QBrush(QColor("#f472b6")))
            painter.setPen(QPen(port_border_color, 2))
            painter.drawEllipse(pos,
                              int(self.PORT_RADIUS * self.scale),
                              int(self.PORT_RADIUS * self.scale))
            
            # Port label
            painter.setPen(QPen(port_label_color))
            fm = QFontMetrics(port_font)
            label_width = fm.horizontalAdvance(port.name)
            label_rect = QRectF(pos.x() - label_width - 12 * self.scale,
                               pos.y() - 8 * self.scale,
                               80 * self.scale, 16 * self.scale)
            painter.drawText(label_rect, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
                           port.name)
    
    def _draw_connections(self, painter: QPainter):
        """Draw all connections"""
        for conn in self.workflow.connections:
            from_node = self.workflow.nodes.get(conn.from_node)
            to_node = self.workflow.nodes.get(conn.to_node)
            
            if from_node and to_node:
                start = self.get_port_pos(from_node, conn.from_port, True)
                end = self.get_port_pos(to_node, conn.to_port, False)
                
                # Check if nodes are running for animated connection
                from_status = self.node_status.get(conn.from_node)
                to_status = self.node_status.get(conn.to_node)
                
                if from_status == 'success' and to_status == 'running':
                    self._draw_connection_line(painter, start, end, QColor("#22c55e"), animated=True)
                elif from_status == 'running' or to_status == 'running':
                    self._draw_connection_line(painter, start, end, QColor("#fbbf24"))
                else:
                    self._draw_connection_line(painter, start, end, QColor("#888899"))
    
    def _draw_connection_line(self, painter: QPainter, start: QPoint, end: QPoint, color: QColor, animated: bool = False):
        """Draw a bezier connection line"""
        path = QPainterPath()
        path.moveTo(QPointF(start))
        
        # Calculate control points for smooth curve
        dx = abs(end.x() - start.x())
        ctrl_offset = max(dx * 0.5, 50 * self.scale)
        
        ctrl1 = QPointF(start.x() + ctrl_offset, start.y())
        ctrl2 = QPointF(end.x() - ctrl_offset, end.y())
        
        path.cubicTo(ctrl1, ctrl2, QPointF(end))
        
        pen = QPen(color, 2 * self.scale)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        
        # Animated dashed line for data flow
        if animated and self._is_animating:
            pen.setStyle(Qt.PenStyle.DashLine)
            pen.setDashOffset(self._animation_offset)
            pen.setWidth(int(3 * self.scale))
        
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawPath(path)
    
    def _draw_minimap(self, painter: QPainter):
        """Draw minimap in the corner"""
        if not self.workflow.nodes:
            return
        
        # Calculate minimap position (bottom-right corner)
        minimap_rect = QRectF(
            self.width() - self._minimap_size - self._minimap_margin,
            self.height() - self._minimap_size - self._minimap_margin,
            self._minimap_size,
            self._minimap_size
        )
        
        # Draw minimap background
        painter.save()
        if self._is_dark_theme:
            bg_color = QColor(30, 30, 50, 200)
            border_color = QColor(60, 60, 100)
        else:
            bg_color = QColor(255, 255, 255, 220)
            border_color = QColor(180, 180, 180)
        
        painter.setBrush(QBrush(bg_color))
        painter.setPen(QPen(border_color, 1))
        painter.drawRoundedRect(minimap_rect, 5, 5)
        
        # Calculate bounds of all nodes
        min_x = min_y = float('inf')
        max_x = max_y = float('-inf')
        
        for node in self.workflow.nodes.values():
            x, y = node.position
            min_x = min(min_x, x)
            min_y = min(min_y, y)
            max_x = max(max_x, x + self.NODE_WIDTH)
            max_y = max(max_y, y + self.NODE_HEIGHT)
        
        # Add padding
        padding = 50
        min_x -= padding
        min_y -= padding
        max_x += padding
        max_y += padding
        
        world_width = max_x - min_x
        world_height = max_y - min_y
        
        if world_width <= 0 or world_height <= 0:
            painter.restore()
            return
        
        # Calculate scale to fit in minimap
        minimap_inner = minimap_rect.adjusted(5, 5, -5, -5)
        scale_x = minimap_inner.width() / world_width
        scale_y = minimap_inner.height() / world_height
        minimap_scale = min(scale_x, scale_y)
        
        # Center the content
        scaled_width = world_width * minimap_scale
        scaled_height = world_height * minimap_scale
        offset_x = minimap_inner.x() + (minimap_inner.width() - scaled_width) / 2
        offset_y = minimap_inner.y() + (minimap_inner.height() - scaled_height) / 2
        
        # Draw nodes in minimap
        for node in self.workflow.nodes.values():
            x, y = node.position
            node_x = offset_x + (x - min_x) * minimap_scale
            node_y = offset_y + (y - min_y) * minimap_scale
            node_w = self.NODE_WIDTH * minimap_scale
            node_h = self.NODE_HEIGHT * minimap_scale
            
            node_color = QColor(node.node_color)
            painter.setBrush(QBrush(node_color))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(QRectF(node_x, node_y, node_w, node_h), 2, 2)
        
        # Draw viewport rectangle
        view_x = (-self.offset.x() / self.scale - min_x) * minimap_scale + offset_x
        view_y = (-self.offset.y() / self.scale - min_y) * minimap_scale + offset_y
        view_w = (self.width() / self.scale) * minimap_scale
        view_h = (self.height() / self.scale) * minimap_scale
        
        viewport_color = QColor("#3b82f6") if self._is_dark_theme else QColor("#2563eb")
        viewport_color.setAlpha(100)
        painter.setBrush(QBrush(viewport_color))
        painter.setPen(QPen(QColor("#3b82f6"), 1))
        painter.drawRect(QRectF(view_x, view_y, view_w, view_h))
        
        # Draw minimap label
        painter.setPen(QPen(border_color))
        font = QFont("Segoe UI", 8)
        painter.setFont(font)
        painter.drawText(int(minimap_rect.x() + 5), int(minimap_rect.y() + 12), "小地图")
        
        painter.restore()
    
    def toggle_minimap(self):
        """Toggle minimap visibility"""
        self._show_minimap = not self._show_minimap
        self.update()
    
    def _get_minimap_rect(self) -> QRectF:
        """Get the minimap rectangle"""
        return QRectF(
            self.width() - self._minimap_size - self._minimap_margin,
            self.height() - self._minimap_size - self._minimap_margin,
            self._minimap_size,
            self._minimap_size
        )
    
    def mousePressEvent(self, event: QMouseEvent):
        """Handle mouse press"""
        pos = event.pos()
        
        if event.button() == Qt.MouseButton.LeftButton:
            # Check if clicking on a port
            port_info = self.get_port_at(pos)
            if port_info:
                node_id, port_name, is_output = port_info
                self.connecting = True
                self.connection_start_node = node_id
                self.connection_start_port = port_name
                self.connection_start_is_output = is_output
                self.connection_end_pos = pos
                return
            
            # Check if clicking on a node
            node_id = self.get_node_at(pos)
            if node_id and node_id in self.workflow.nodes:
                self.selected_node = node_id
                self.dragging_node = node_id
                node = self.workflow.nodes[node_id]
                canvas_pos = self.screen_to_canvas(pos)
                self.drag_offset = QPoint(
                    canvas_pos.x() - int(node.position[0]),
                    canvas_pos.y() - int(node.position[1])
                )
                self.node_selected.emit(node_id)
            else:
                self.selected_node = None
                self.node_selected.emit("")
            
            self.update()
        
        elif event.button() == Qt.MouseButton.MiddleButton:
            self.panning = True
            self.pan_start = pos
            self.setCursor(Qt.CursorShape.ClosedHandCursor)
    
    def mouseMoveEvent(self, event: QMouseEvent):
        """Handle mouse move"""
        pos = event.pos()
        
        if self.connecting:
            self.connection_end_pos = pos
            self.update()
        
        elif self.dragging_node:
            if self.dragging_node in self.workflow.nodes:
                canvas_pos = self.screen_to_canvas(pos)
                node = self.workflow.nodes[self.dragging_node]
                # Ensure position is stored as tuple of integers
                new_x = int(canvas_pos.x() - self.drag_offset.x())
                new_y = int(canvas_pos.y() - self.drag_offset.y())
                node.position = (new_x, new_y)
                self.update()
            else:
                # Node was deleted while dragging
                self.dragging_node = None
        
        elif self.panning:
            delta = pos - self.pan_start
            self.offset += delta
            self.pan_start = pos
            self.update()
        
        else:
            # Update cursor
            port_info = self.get_port_at(pos)
            if port_info:
                self.setCursor(Qt.CursorShape.CrossCursor)
            elif self.get_node_at(pos):
                self.setCursor(Qt.CursorShape.SizeAllCursor)
            else:
                self.setCursor(Qt.CursorShape.ArrowCursor)
    
    def mouseReleaseEvent(self, event: QMouseEvent):
        """Handle mouse release"""
        pos = event.pos()
        
        if event.button() == Qt.MouseButton.LeftButton:
            if self.connecting:
                # Check if ending on a port
                port_info = self.get_port_at(pos)
                if port_info:
                    end_node, end_port, end_is_output = port_info
                    
                    # Can only connect output to input
                    if self.connection_start_is_output and not end_is_output:
                        self.workflow.add_connection(
                            self.connection_start_node,
                            self.connection_start_port,
                            end_node,
                            end_port
                        )
                        self.connection_created.emit()
                    elif not self.connection_start_is_output and end_is_output:
                        self.workflow.add_connection(
                            end_node,
                            end_port,
                            self.connection_start_node,
                            self.connection_start_port
                        )
                        self.connection_created.emit()
                
                self.connecting = False
                self.connection_start_node = None
                self.connection_start_port = None
                self.update()
            
            self.dragging_node = None
        
        elif event.button() == Qt.MouseButton.MiddleButton:
            self.panning = False
            self.setCursor(Qt.CursorShape.ArrowCursor)
    
    def mouseDoubleClickEvent(self, event: QMouseEvent):
        """Handle double click"""
        node_id = self.get_node_at(event.pos())
        if node_id:
            self.node_double_clicked.emit(node_id)
    
    def contextMenuEvent(self, event):
        """Handle right-click context menu"""
        node_id = self.get_node_at(event.pos())
        
        menu = QMenu(self)
        
        if node_id:
            # Select the node
            self.selected_node = node_id
            self.node_selected.emit(node_id)
            self.update()
            
            node = self.workflow.nodes.get(node_id)
            if node:
                # Node name as header
                header_action = menu.addAction(f"📦 {node.node_name}")
                header_action.setEnabled(False)
                menu.addSeparator()
                
                # Execute Node
                exec_action = menu.addAction("▶️ 执行节点")
                exec_action.triggered.connect(lambda: self.node_execution_requested.emit(node_id))
                
                menu.addSeparator()
                
                # Copy node
                copy_action = menu.addAction("📋 复制节点")
                copy_action.triggered.connect(lambda: self.node_copy_requested.emit(node_id))
                
                # Delete node
                delete_action = menu.addAction("🗑️ 删除节点")
                delete_action.triggered.connect(lambda: self.node_delete_requested.emit(node_id))
                
                menu.addSeparator()
                
                # View help
                help_action = menu.addAction("❓ 查看帮助")
                help_action.triggered.connect(lambda: self.node_double_clicked.emit(node_id))
        else:
            # Canvas context menu
            run_action = menu.addAction("▶️ 执行工作流")
            run_action.triggered.connect(self.workflow_execution_requested.emit)
            
            menu.addSeparator()
            
            zoom_in_action = menu.addAction("🔍 放大")
            zoom_in_action.triggered.connect(lambda: self.zoom(1.2))
            
            zoom_out_action = menu.addAction("🔍 缩小")
            zoom_out_action.triggered.connect(lambda: self.zoom(0.8))
            
            fit_action = menu.addAction("⊡ 适应视图")
            fit_action.triggered.connect(self.fit_to_view)
        
        menu.exec(event.globalPos())
    
    def dragEnterEvent(self, event):
        """Handle drag enter event"""
        if event.mimeData().hasFormat("application/x-workflow-node"):
            event.acceptProposedAction()
        elif event.mimeData().hasText():
            # Also accept plain text (node type)
            event.acceptProposedAction()
    
    def dragMoveEvent(self, event):
        """Handle drag move event"""
        if event.mimeData().hasFormat("application/x-workflow-node") or event.mimeData().hasText():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        """Handle drop event - create node at drop position"""
        try:
            if event.mimeData().hasFormat("application/x-workflow-node"):
                node_type = bytes(event.mimeData().data("application/x-workflow-node")).decode()
            elif event.mimeData().hasText():
                node_type = event.mimeData().text()
            else:
                return
            
            # Convert drop position to canvas coordinates
            canvas_pos = self.screen_to_canvas(event.position().toPoint())
            
            # Emit signal with node type and position
            self.node_dropped.emit(node_type, canvas_pos.x(), canvas_pos.y())
            
            event.acceptProposedAction()
        except Exception as e:
            print(f"Drop error: {e}")
    
    def wheelEvent(self, event: QWheelEvent):
        """Handle mouse wheel for zooming"""
        delta = event.angleDelta().y()
        
        if delta > 0:
            self.zoom(1.1)
        else:
            self.zoom(0.9)
    
    def keyPressEvent(self, event):
        """Handle key press"""
        if event.key() == Qt.Key.Key_Delete:
            if self.selected_node:
                self.workflow.remove_node(self.selected_node)
                self.selected_node = None
                self.node_selected.emit("")
                self.update()
    
    def set_node_status(self, node_id: str, status: str):
        """Set the execution status of a node
        
        Args:
            node_id: The node ID
            status: One of 'pending', 'running', 'success', 'error'
        """
        self.node_status[node_id] = status
        self.update()
    
    def clear_node_status(self):
        """Clear all node status indicators"""
        self.node_status.clear()
        self.update()
    
    def start_animation(self):
        """Start connection animation"""
        self._is_animating = True
        self._animation_timer.start(50)  # 20 FPS
    
    def stop_animation(self):
        """Stop connection animation"""
        self._is_animating = False
        self._animation_timer.stop()
        self._animation_offset = 0
        self.update()
    
    def _animate_connections(self):
        """Update animation offset"""
        self._animation_offset = (self._animation_offset + 3) % 20
        self.update()
