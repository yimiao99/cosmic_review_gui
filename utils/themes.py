import sys
import os
from extend.matcher_config import MatcherConfig
from utils.themes_dark import get_dark_theme_stylesheet
from utils.themes_light import get_light_theme_stylesheet


def apply_dark_title_bar(window, is_dark=True):
    """
    针对 Windows 10 (1809+) 和 Windows 11 开启原生标题栏深色模式
    """
    if sys.platform != "win32":
        return

    try:
        import ctypes
        from ctypes import wintypes

        # 获取 HWND，PySide6 的 winId() 返回的是个整数
        hwnd = int(window.winId())

        # DWMWA_USE_IMMERSIVE_DARK_MODE:
        # 20 是 Windows 11 及其以后版本
        # 19 是 Windows 10 (1809 - 21H1)
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        DWMWA_USE_IMMERSIVE_DARK_MODE_OLD = 19

        value = ctypes.c_int(1 if is_dark else 0)

        # 加载 dwmapi.dll
        dwmapi = ctypes.windll.dwmapi

        dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value),
            ctypes.sizeof(value),
        )
        dwmapi.DwmSetWindowAttribute(
            hwnd,
            DWMWA_USE_IMMERSIVE_DARK_MODE_OLD,
            ctypes.byref(value),
            ctypes.sizeof(value),
        )
    except Exception as e:
        print(f"Set dark title bar failed: {e}")


def get_theme_stylesheet(is_dark=False):
    """根据主题类型返回对应的样式表"""
    if is_dark:
        return get_dark_theme_stylesheet()
    else:
        return get_light_theme_stylesheet()


LIGHT_THEME = ""  # 为兼容性保留，实际改用 get_theme_stylesheet
DARK_THEME = ""
