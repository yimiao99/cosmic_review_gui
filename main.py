import sys
import os

# ==========================================================
# 0. 适配高分屏 (解决字体模糊问题)
# ==========================================================
if sys.platform == "win32":
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
    os.environ["QT_AUTOSCREENSCALEFACTOR"] = "1"

# ==========================================================
# 1. 核心修复：解决 PySide6 DLL 加载冲突 (必须在任何 PySide6 导入之前)
# ==========================================================
if sys.platform == "win32":
    # 获取基础运行目录
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))

    # 清理环境变量 PATH 中的干扰项
    if "PATH" in os.environ:
        os.environ["PATH"] = os.pathsep.join(
            [
                p
                for p in os.environ["PATH"].split(os.pathsep)
                if "ACE Studio" not in p and "Anaconda" not in p
            ]
        )

    # 确定 PySide6 DLL 的潜在目录（增加对 _internal 的搜寻）
    pyside_dll_dirs = []

    if getattr(sys, "frozen", False):
        # 打包后的环境：可能是 onefile 根目录，也可能是 onedir 的 _internal 目录
        internal_dir = os.path.join(base_dir, "_internal")
        search_roots = (
            [base_dir, internal_dir] if os.path.exists(internal_dir) else [base_dir]
        )

        for root in search_roots:
            pyside_dll_dirs.extend(
                [root, os.path.join(root, "PySide6"), os.path.join(root, "shiboken6")]
            )
    else:
        # 开发环境
        try:
            import importlib.util

            for module_name in ["PySide6", "shiboken6"]:
                spec = importlib.util.find_spec(module_name)
                if spec and spec.origin:
                    pyside_dll_dirs.append(os.path.dirname(spec.origin))
        except ImportError:
            pass

    # 显式添加所有有效的 DLL 搜索路径并置于 PATH 最前
    for d in reversed(pyside_dll_dirs):  # 反向遍历确保最精细的目录在 PATH 最前面
        if os.path.exists(d):
            if hasattr(os, "add_dll_directory"):
                try:
                    os.add_dll_directory(d)
                except Exception:
                    pass
            os.environ["PATH"] = d + os.pathsep + os.environ["PATH"]

# ==========================================================
# 2. 项目路径与依赖导入
# ==========================================================
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from PySide6.QtWidgets import QApplication
    from PySide6.QtGui import QIcon
    from ui.main_window import CosmicMainWindow
    from utils.path_utils import get_resource_path
    from utils.runtime_logger import RuntimeLogger  # ✅ 导入日志器
except ImportError as e:
    # 如果失败，弹出更具体的调试信息
    if sys.platform == "win32":
        import ctypes

        search_info = "\n".join([d for d in pyside_dll_dirs if os.path.exists(d)])
        ctypes.windll.user32.MessageBoxW(
            0,
            f"DLL 加载失败。\n\n错误: {e}\n\n搜寻目录:\n{search_info}",
            "启动失败",
            16,
        )
    sys.exit(1)


def main():
    # 设置高 DPI 缩放策略 (在 QApplication 实例化前)
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QApplication
    from PySide6.QtGui import QIcon

    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # 设置全局默认字体
    from PySide6.QtGui import QFont

    font = QFont("Microsoft YaHei UI", 9)
    app.setFont(font)

    try:
        logo_path = get_resource_path("ui/logo.png")
        if os.path.exists(logo_path):
            app.setWindowIcon(QIcon(logo_path))
    except Exception:
        pass

    window = CosmicMainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    RuntimeLogger.start_session()
    try:
        exit_code = main()
        sys.exit(exit_code)
    finally:
        RuntimeLogger.save_session()
