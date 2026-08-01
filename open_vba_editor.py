# -*- coding: utf-8 -*-
"""激活+最大化 AlphaCAM 窗口，然后发送 Alt+F11 打开 VBA 编辑器。"""
import ctypes
import time

user32 = ctypes.windll.user32

SW_MAXIMIZE = 3


def find_alphacam_window():
    """按窗口标题精确查找 AlphaCAM 主窗口句柄（排除 VSCode 等）。"""
    results = []

    # AlphaCAM 窗口标题特征（排除编辑器/浏览器）
    include_keywords = ["3d 5-", "alphacam", "alpha cam"]
    exclude_keywords = [
        "visual studio code", "code -", "readme", "markdown", ".md",
        "chrome", "edge", "firefox", "explorer", "git", "terminal",
        "cmd.exe", "powershell", "notepad", "预览",
    ]

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def enum_callback(hwnd, lparam):
        length = user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            title = buf.value
            low = title.lower()
            if not user32.IsWindowVisible(hwnd):
                return True
            if any(k in low for k in include_keywords) and \
               not any(e in low for e in exclude_keywords):
                results.append((hwnd, title))
        return True

    user32.EnumWindows(enum_callback, 0)
    # 优先选择含 "3D 5-轴" 的（最可能是主窗口）
    results.sort(key=lambda x: 0 if "3d 5-" in x[1].lower() else 1)
    return results


def activate_and_maximize(hwnd):
    """激活并最大化窗口。"""
    user32.ShowWindow(hwnd, SW_MAXIMIZE)
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.5)
    user32.BringWindowToTop(hwnd)
    user32.SetForegroundWindow(hwnd)


def send_alt_f11():
    """发送 Alt+F11 快捷键。"""
    import subprocess
    ps = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "[System.Windows.Forms.SendKeys]::SendWait('%{F11}')"
    )
    subprocess.run(["powershell", "-Command", ps], check=False, capture_output=True)


if __name__ == "__main__":
    wins = find_alphacam_window()
    if not wins:
        print("ERROR: 未找到 AlphaCAM 窗口（可手动激活后重试）")
    else:
        hwnd, title = wins[0]
        print(f"找到窗口: {title} (hwnd={hwnd})")
        activate_and_maximize(hwnd)
        print("窗口已激活并最大化")
        time.sleep(0.5)
        send_alt_f11()
        print("已发送 Alt+F11，VBA 编辑器应已打开")
