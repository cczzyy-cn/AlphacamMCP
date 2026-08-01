# -*- coding: utf-8 -*-
"""激活+最大化 AlphaCAM 窗口，然后发送 Alt+F11 打开 VBA 编辑器。"""
import ctypes
import time

user32 = ctypes.windll.user32

SW_MAXIMIZE = 3
ALT_F11 = "%{F11}"


def find_alphacam_window():
    """按窗口标题查找 AlphaCAM 主窗口句柄。"""
    # 遍历所有顶层窗口，匹配标题包含 '3D 5-轴' 或 'Alphacam' 或 'AlphaCAM'
    results = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def enum_callback(hwnd, lparam):
        length = user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            title = buf.value
            low = title.lower()
            if ("3d 5-" in low or "alphacam" in low or "alpha cam" in low) and user32.IsWindowVisible(hwnd):
                results.append((hwnd, title))
        return True

    user32.EnumWindows(enum_callback, 0)
    return results


def activate_and_maximize(hwnd):
    """激活并最大化窗口。"""
    # 还原 + 最大化
    user32.ShowWindow(hwnd, SW_MAXIMIZE)
    # 置前
    user32.SetForegroundWindow(hwnd)
    # 确保在前台
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
        print("ERROR: 未找到 AlphaCAM 窗口")
    else:
        hwnd, title = wins[0]
        print(f"找到窗口: {title} (hwnd={hwnd})")
        activate_and_maximize(hwnd)
        print("窗口已激活并最大化")
        time.sleep(0.5)
        send_alt_f11()
        print("已发送 Alt+F11，VBA 编辑器应已打开")
