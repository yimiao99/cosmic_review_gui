# -*- coding: utf-8 -*-
import sys
import os
import traceback

# ==========================================================
# 0. PyInstaller 打包环境兼容性修复 (核心：解决图标/字体不显示)
# ==========================================================
if getattr(sys, 'frozen', False):
    # 如果是打包后的环境
    base_path = sys._MEIPASS

    # 1. 强制指定 Qt 插件路径 (解决图标、下拉箭头、字体渲染等 UI 元素丢失问题)
    qt_plugin_path = os.path.join(base_path, 'PySide6', 'plugins')
    if os.path.exists(qt_plugin_path):
        os.environ['QT_PLUGIN_PATH'] = qt_plugin_path
        # 告诉 QCoreApplication 去哪里找插件
        from PySide6.QtCore import QCoreApplication

        QCoreApplication.addLibraryPath(qt_plugin_path)

    # 2. 将 PySide6 目录加入系统 PATH (解决部分底层 DLL 或字体加载失败的问题)
    pyside6_path = os.path.join(base_path, 'PySide6')
    if os.path.exists(pyside6_path):
        os.environ['PATH'] = pyside6_path + os.pathsep + os.environ.get('PATH', '')
else:
    # 开发环境
    base_path = os.path.abspath(os.path.dirname(__file__))

# 将项目根目录加入 sys.path，确保内部模块导入正常
sys.path.insert(0, base_path)

# ==========================================================
# 1. 核心依赖导入
# ==========================================================
try:
    from PySide6.QtWidgets import QApplication
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QFont, QIcon

    from ui.main_window import CosmicMainWindow
    from utils.path_utils import get_resource_path
    from utils.runtime_logger import RuntimeLogger
    from extend.matcher_config import MatcherConfig
except ImportError as e:
    print(f"核心组件加载失败: {e}")
    if sys.platform == "win32":
        import ctypes

        ctypes.windll.user32.MessageBoxW(
            0,
            f"核心组件加载失败。\n\n错误: {e}\n\n请检查 PySide6 是否安装完整，或依赖是否缺失。",
            "启动失败",
            16,
        )
    sys.exit(1)


def main():
    """主程序入口"""
    # 1. 设置高 DPI 缩放策略 (必须在 QApplication 实例化前)
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    # 2. 创建应用实例
    app = QApplication(sys.argv)

    # 3. 设置全局样式 (Fusion 样式在不同平台上表现更一致)
    app.setStyle("Fusion")

    # 4. 设置全局默认字体 (解决部分 Windows 环境下字体渲染发虚或默认字体不对的问题)
    # 优先使用微软雅黑，如果系统没有则回退到系统默认
    font = QFont("Microsoft YaHei UI", 9)
    if not font.exactMatch():
        font = QFont("Microsoft YaHei", 9)
    app.setFont(font)

    # 5. 设置应用图标 (兼容打包环境，使用 get_resource_path)
    try:
        # 尝试加载 .ico 或 .png 格式的图标
        icon_path = get_resource_path("ui/logo.png")
        if not os.path.exists(icon_path):
            icon_path = get_resource_path("ui/logo.ico")

        if os.path.exists(icon_path):
            app_icon = QIcon(icon_path)
            app.setWindowIcon(app_icon)
        else:
            print(f"警告: 未找到应用图标文件: {icon_path}")
    except Exception as e:
        print(f"设置应用图标时发生异常: {e}")

    # 6. 创建并显示主窗口
    try:
        window = CosmicMainWindow()
        window.show()
    except Exception as e:
        print(f"主窗口初始化失败: {e}")
        traceback.print_exc()
        return 1

    # 7. 启动事件循环
    return app.exec()


if __name__ == "__main__":
    # 启动全局日志会话
    RuntimeLogger.start_session()

    try:
        # 执行主程序
        exit_code = main()
        sys.exit(exit_code)
    except Exception as e:
        # 捕获未处理的顶层异常并记录到日志
        error_msg = f"程序发生严重未捕获错误:\n{str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        try:
            RuntimeLogger.log(error_msg, level="ERROR")
        except:
            pass
        sys.exit(1)
    finally:
        # 确保程序退出时保存日志并释放资源
        try:
            RuntimeLogger.save_session()
            RuntimeLogger.close_file()
        except:
            pass