import win32gui
import cv2
from PIL import ImageGrab
import numpy as np
import os
import sys

def resource_path(relative_path):
    """
    Get absolute path to resource.
    Works both for development and for PyInstaller-built EXE.
    """
    try:
        base_path = sys._MEIPASS  # temp folder created by PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_save_dir():
    """
    Create (if needed) and return a safe folder path
    where screenshots will be saved.
    """
    save_dir = os.path.join(os.getcwd(), "screenshots")
    os.makedirs(save_dir, exist_ok=True)
    return save_dir

def capture_word_window(hwnd, filename="latest_word.png", exclude_widget=None):
    if not hwnd:
        return False

    try:
        if exclude_widget:
            exclude_widget.hide()  # hide the bot GUI while capturing

        # Get window rectangle and capture it
        x, y, x1, y1 = win32gui.GetWindowRect(hwnd)
        img = ImageGrab.grab(bbox=(x, y, x1, y1))
        frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)

        # Construct a safe save path
        save_dir = get_save_dir()
        save_path = os.path.join(save_dir, filename)

        # Write the image
        cv2.imwrite(save_path, frame)
        print(f"✅ Screenshot saved at: {save_path}")

        if exclude_widget:
            exclude_widget.show()  # show the bot GUI again
        return True

    except Exception as e:
        print("❌ Error capturing Word window:", e)
        if exclude_widget:
            exclude_widget.show()
        return False
