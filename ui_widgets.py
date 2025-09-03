from PySide6.QtWidgets import QLabel, QScrollArea
from PySide6.QtCore import Qt, Signal, QRect, QPoint
from PySide6.QtGui import QPainter, QPen, QGuiApplication

class ImageLabel(QLabel):
    roi_selected = Signal(QRect)
    def __init__(self, parent=None):
        super().__init__(parent); self.begin, self.end, self.selecting = QPoint(), QPoint(), False
    def mousePressEvent(self, event):
        if self.selecting and event.button() == Qt.MouseButton.LeftButton: self.begin, self.end, _ = event.pos(), event.pos(), self.update()
    def mouseMoveEvent(self, event):
        if self.selecting: self.end, _ = event.pos(), self.update()
    def mouseReleaseEvent(self, event):
        if self.selecting and event.button() == Qt.MouseButton.LeftButton: self.selecting, self.roi_selected.emit(QRect(self.begin, self.end).normalized()), self.update()
    def paintEvent(self, event):
        super().paintEvent(event)
        if self.selecting: painter = QPainter(self); painter.setPen(QPen(Qt.red, 2, Qt.SolidLine)); painter.drawRect(QRect(self.begin, self.end).normalized())

class ZoomableScrollArea(QScrollArea):
    def __init__(self, main_window, parent=None):
        super().__init__(parent); self.main_window = main_window; self.setWidgetResizable(True); self.setAlignment(Qt.AlignCenter)
    def wheelEvent(self, event):
        modifiers = QGuiApplication.keyboardModifiers()
        if modifiers == Qt.KeyboardModifier.ControlModifier:
            angle = event.angleDelta().y(); self.main_window.zoom_image(1.2 if angle > 0 else 0.8); event.accept()
        elif modifiers == Qt.KeyboardModifier.ShiftModifier:
            delta = event.angleDelta().y(); h_bar = self.horizontalScrollBar(); h_bar.setValue(h_bar.value() - delta); event.accept()
        else: super().wheelEvent(event)