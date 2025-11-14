import sys
import os
import subprocess
import win32gui
import win32con
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QScrollArea, QFrame
)
from PyQt5.QtCore import Qt, QPropertyAnimation
from PyQt5.QtGui import QFont

# -------------------------
# Safe path helper for PyInstaller
# -------------------------
def resource_path(relative_path):
    """
    Get absolute path to resource (works for .py and bundled .exe)
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# -------------------------
# Local imports
# -------------------------
from screen_capture import capture_word_window
from roboflow_detect import detect_objects
from overlay import Overlay
from llm_engine import SmartLLMEngine  # ✅ Use your new LLM engine


# -------------------------
# Word handling
# -------------------------
word_hwnd_cache = None

def get_word_hwnd():
    global word_hwnd_cache
    if word_hwnd_cache and win32gui.IsWindow(word_hwnd_cache):
        return word_hwnd_cache
    hwnd_found = None
    def enum_handler(hwnd, _):
        nonlocal hwnd_found
        title = win32gui.GetWindowText(hwnd)
        if "Word" in title and win32gui.IsWindowVisible(hwnd):
            hwnd_found = hwnd
            return False
        return True
    win32gui.EnumWindows(lambda hwnd, _: enum_handler(hwnd, None), None)
    word_hwnd_cache = hwnd_found
    return hwnd_found

def bring_word_front_and_fullscreen(hwnd):
    if not hwnd:
        return
    try:
        foreground = win32gui.GetForegroundWindow()
        placement = win32gui.GetWindowPlacement(hwnd)
        show_cmd = placement[1]
        if hwnd == foreground and show_cmd == win32con.SW_SHOWMAXIMIZED:
            return
        if show_cmd == win32con.SW_SHOWMINIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(hwnd)
            return
        if hwnd != foreground:
            win32gui.SetForegroundWindow(hwnd)
            if show_cmd != win32con.SW_SHOWMAXIMIZED:
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            return
        if hwnd == foreground and show_cmd == win32con.SW_SHOWNORMAL:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    except Exception as e:
        print("Error bringing Word to front:", e)

# -------------------------
# Initialize Smart LLM Engine
# -------------------------
llm_engine = SmartLLMEngine(model_name="gemma:2b", label_list_path=resource_path("label_mapping.txt"))

# -------------------------
# Chat bubble
# -------------------------
class ChatBubble(QFrame):
    def __init__(self, text, is_user=False):
        super().__init__()
        self.setObjectName("bubble")
        self.setStyleSheet("""
            QFrame#bubble {
                border-radius: 12px;
                padding: 8px;
                margin: 6px;
            }
        """)
        layout = QHBoxLayout()
        label = QLabel(text)
        label.setWordWrap(True)
        label.setFont(QFont("Segoe UI", 10))
        label.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)

        if is_user:
            self.setStyleSheet(self.styleSheet() + "QFrame#bubble { background: #4b6ef6; color: white; margin-left: 80px; }")
            layout.addStretch()
            layout.addWidget(label)
        else:
            self.setStyleSheet(self.styleSheet() + "QFrame#bubble { background: #2e2e2e; color: #eaeaea; margin-right: 80px; }")
            layout.addWidget(label)
            layout.addStretch()

        self.setLayout(layout)

# -------------------------
# Main Chat Window
# -------------------------
class ChatWindow(QWidget):
    def __init__(self):
        super().__init__()
        self._drag_active = False
        self._drag_position = None
        self.overlay = None

        self.setWindowTitle("AI Office Tutor")
        self.resize(400, 650)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)

        self.container = QWidget(self)
        self.container.setStyleSheet("background:#0f1724; color:#e5e7eb; border-radius:12px;")
        self.container.setGeometry(0, 0, self.width(), self.height())
        self.init_ui()

        # Position bottom-right
        screen_geom = QApplication.desktop().availableGeometry()
        self.move(screen_geom.width() - self.width() - 10, screen_geom.height() - self.height() - 70)
        
    def resizeEvent(self, event):
        self.container.setGeometry(0, 0, self.width(), self.height())
        super().resizeEvent(event)
        
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_active = True
            self._drag_position = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if Qt.LeftButton and self._drag_active:
            self.move(event.globalPos() - self._drag_position)
            event.accept()

    def mouseReleaseEvent(self, event):
        self._drag_active = False

    def init_ui(self):
        main_layout = QVBoxLayout(self.container)
        main_layout.setContentsMargins(0,0,0,0)
        main_layout.setSpacing(0)

        title_layout = QHBoxLayout()
        title_layout.setContentsMargins(10,5,10,5)
        title_label = QLabel("URA AI")
        title_label.setFont(QFont("Segoe UI", 12, QFont.Bold))
        title_label.setStyleSheet("color: white;")
        title_layout.addWidget(title_label)
        title_layout.addStretch()

        min_btn = QPushButton("-")
        min_btn.setFixedSize(24,24)
        min_btn.setStyleSheet("color:white; background:#1f2937; border:none;")
        min_btn.clicked.connect(self.showMinimized)
        title_layout.addWidget(min_btn)

        close_btn = QPushButton("×")
        close_btn.setFixedSize(24,24)
        close_btn.setStyleSheet("color:white; background:#1f2937; border:none;")
        close_btn.clicked.connect(self.close)
        title_layout.addWidget(close_btn)

        main_layout.addLayout(title_layout)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_widget)
        self.scroll_layout.addStretch()
        self.scroll.setWidget(self.scroll_widget)
        self.scroll.setStyleSheet("background: transparent; border: none;")
        main_layout.addWidget(self.scroll)

        input_row = QHBoxLayout()
        self.input_box = QLineEdit()
        self.input_box.setPlaceholderText("Ask something like 'how to make text bold'...")
        self.input_box.returnPressed.connect(self.on_send)
        self.input_box.setStyleSheet("padding:10px; border-radius:8px; background:#111827; color:#e5e7eb;")
        input_row.addWidget(self.input_box)

        send_btn = QPushButton("Send")
        send_btn.clicked.connect(self.on_send)
        send_btn.setFixedWidth(80)
        send_btn.setStyleSheet("background:#2563eb; color:white; padding:8px; border-radius:8px;")
        input_row.addWidget(send_btn)
        main_layout.addLayout(input_row)

    def add_bubble(self, text, is_user=False):
        bubble = ChatBubble(text, is_user=is_user)
        self.scroll_layout.insertWidget(self.scroll_layout.count()-1, bubble)
        QApplication.processEvents()
        scroll_bar = self.scroll.verticalScrollBar()
        scroll_bar.setValue(scroll_bar.maximum())
        start_value = scroll_bar.value()
        end_value = scroll_bar.maximum()
        anim = QPropertyAnimation(scroll_bar, b"value", self)
        anim.setDuration(600)
        anim.setStartValue(start_value)
        anim.setEndValue(end_value)
        anim.start()
        self._scroll_anim = anim

    def on_send(self):
        user_query = self.input_box.text().strip()
        if not user_query:
            return
        self.add_bubble(user_query, is_user=True)
        self.input_box.clear()

        hwnd = get_word_hwnd()
        if not hwnd:
            self.add_bubble("⚠️ No Word window found! Please open Word first.", False)
            return

        bring_word_front_and_fullscreen(hwnd)

        response = llm_engine.query(user_query)
        label_name = response.get("label")
        intent = response.get("intent", user_query)
        self.add_bubble(f"🧠 Intent: {intent}\n🔗 Label: {label_name}\n")

        if not label_name:
            self.add_bubble("⚠️ I couldn’t map this to a feature.", False)
            return

        # ✅ Save screenshots safely beside EXE
        screenshot_dir = os.path.join(os.getcwd(), "screenshots")
        os.makedirs(screenshot_dir, exist_ok=True)
        captured_image_path = os.path.join(screenshot_dir, "latest_word.png")

        captured = capture_word_window(hwnd, filename=captured_image_path, exclude_widget=self)
        if not captured:
            self.add_bubble("❌ Failed to capture Word window.", False)
            return

        detections = detect_objects(captured_image_path, save_annotated_path=None, filter_labels=[label_name])
        if not detections:
            self.add_bubble(f"ℹ️ No '{label_name}' detected.", False)
            return

        if self.overlay:
            self.overlay.hide_overlay()
            self.overlay = None

        self.overlay = Overlay(target_hwnd=hwnd)
        boxes = [[*d['box'], d['label']] for d in detections]

        import cv2
        img = cv2.imread(captured_image_path)
        self.overlay.set_boxes(boxes, img_size=(img.shape[1], img.shape[0]))
        self.overlay.show_forever()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = ChatWindow()
    about_message = (
        "👋 **Welcome to URA AI!**\n\n"
        "I'm your intelligent assistant for **Microsoft Word **.\n"
        "You can ask me things like:\n"
        "- 'How to make text bold'\n"
        "- 'Insert a table'\n"
        "- 'Add a chart'\n"
        "- 'Check spelling'\n\n"
        "I'll guide you step-by-step and highlight the exact button on your screen. 🚀"
    )

    w.add_bubble(about_message, is_user=False)

    w.show()
    sys.exit(app.exec_())
