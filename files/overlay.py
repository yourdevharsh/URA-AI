import win32gui
from PyQt5.QtWidgets import QWidget
from PyQt5.QtCore import Qt, QRect, QTimer
from PyQt5.QtGui import QPainter, QColor, QPen, QFont

class Overlay(QWidget):
    def __init__(self, target_hwnd=None):
        super().__init__()
        self.target_hwnd = target_hwnd
        self.boxes = []  # [[x1,y1,x2,y2,label,alpha], ...]
        self.img_size = (1, 1)

        # Transparent, topmost, frameless window
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, True)

        # Timer for fading
        self.timer = QTimer()
        self.timer.setInterval(50)  # 50 ms per step
        self.timer.timeout.connect(self.fade_boxes)

        self.hide()

    def set_boxes(self, boxes, img_size):
        """Set detected boxes and start auto-fade after 2 seconds"""
        self.boxes = [[*b, 255] for b in boxes]  # add alpha channel
        self.img_size = img_size
        self.update()

        # Auto-start fade after 2 seconds
        QTimer.singleShot(2000, self.timer.start)

    def reposition_to_target(self):
        if not self.target_hwnd:
            return
        try:
            rect = win32gui.GetWindowRect(self.target_hwnd)
            self.setGeometry(rect[0], rect[1], rect[2]-rect[0], rect[3]-rect[1])
        except Exception as e:
            print("Overlay reposition error:", e)

    def show_forever(self):
        self.reposition_to_target()
        self.show()

    def hide_overlay(self):
        self.hide()
        self.boxes = []
        self.timer.stop()

    def fade_boxes(self):
        """Gradually reduce alpha for fade effect"""
        updated = False
        for b in self.boxes:
            if b[5] > 0:
                b[5] = max(0, b[5]-15)
                updated = True
        if not updated:
            self.timer.stop()
            self.boxes = []
            self.hide()
        self.update()

    def paintEvent(self, event):
        if not self.boxes:
            return
        painter = QPainter(self)
        width, height = self.img_size
        win_w, win_h = self.width(), self.height()
        x_scale = win_w / width
        y_scale = win_h / height

        for b in self.boxes:
            x1, y1, x2, y2, label, alpha = b
            pen = QPen(QColor(255, 0, 0, alpha), 3)
            painter.setPen(pen)
            rect = QRect(int(x1*x_scale), int(y1*y_scale), int((x2-x1)*x_scale), int((y2-y1)*y_scale))
            painter.drawRect(rect)
            painter.setPen(QColor(255, 255, 255, alpha))
            font = QFont()
            font.setPointSize(10)
            painter.setFont(font)
            painter.drawText(rect.topLeft(), label)
