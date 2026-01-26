import os
import sys


def get_resource_path(relative_path):
    """获取资源的绝对路径，兼容开发环境和 PyInstaller 打包环境"""
    if hasattr(sys, "_MEIPASS"):
        # Onefile 模式：所有东西在 _MEIPASS 根目录（包括子文件夹）
        base_path = sys._MEIPASS
    elif getattr(sys, "frozen", False):
        # Onedir 模式：PyInstaller 6+ 默认把 data 放在 _internal 文件夹下
        exe_dir = os.path.dirname(sys.executable)
        # 优先从 _internal 目录下寻找
        internal_path = os.path.join(exe_dir, "_internal", relative_path)
        if os.path.exists(internal_path):
            return internal_path
        base_path = exe_dir
    else:
        # 开发环境，由于 path_utils.py 在 utils 目录下，其父目录才是项目根目录
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

    return os.path.join(base_path, relative_path)


def clean_project_name(filename):
    """
    净化提取项目名称：去除附件编号前缀及需求规格书等后缀
    保留版本号 (v1.0.0 等)，去除“需求说明书”及其后续文字
    """
    if not filename:
        return "未命名项目"

    import re

    # 1. 基础处理：手动移除扩展名
    name = os.path.basename(filename).strip()
    name = re.sub(r"\.(?:docx?|xlsx?)$", "", name, flags=re.IGNORECASE).strip()

    # 2. 预先提取版本号 (支持 v1.0, V2.1.3 等)
    # 提取最后一个出现的版本号
    versions = re.findall(r"([vV]\d+(?:\.\d+)*)", name)
    version_str = versions[-1] if versions else ""

    # 3. 移除前缀
    name = re.sub(r"^附件\s*\d+\s*[：:.\-\s]*\s*", "", name)
    name = re.sub(r"^\d{1,3}[\.．]\s*", "", name)
    name = name.strip()

    # 4. 强制截断点：找到“需求说明书”等关键后缀词，截断其及之后的所有内容
    truncate_keywords = [
        "产品需求说明书",
        "需求规格说明书",
        "需求规格书",
        "需求说明书",
        "规格说明书",
        "规格书",
        "功能点拆分表",
        "功能拆分表",
        "拆分表",
        "审计方案",
        "测试用例",
        "说明书",
        "文档",
        "需求",
    ]

    for kw in truncate_keywords:
        if kw in name:
            name = name.split(kw)[0]
            break

    # 5. 清理末尾后缀与日期，确保切干净
    while True:
        prev_name = name

        # (a) 去掉日期以及 (1) (2) 这种重复标识
        name = re.sub(r"(?:20)?\d{6,8}\s*$", "", name)
        name = re.sub(r"[\(（]\d+[\)）]\s*$", "", name)
        name = re.sub(r"\s*-\s*副本\s*$", "", name)

        # (b) 去掉末尾的连接词
        name = re.sub(r"(?:的项目|项目|的需求|的产品|产品|的|之)+$", "", name.strip())

        # (c) 去掉悬空的标点
        name = name.strip(" :：-－_|'\"().（）")

        if name == prev_name:
            break

    # 6. 恢复版本号 (如果原名中已无版本号且提取到了，则追加)
    if version_str and version_str not in name:
        name = f"{name} {version_str}"

    return name.strip()


def clear_directory(directory_path):
    """
    彻底清空文件夹下的所有文件和子文件夹
    """
    if not os.path.exists(directory_path):
        return True, "文件夹不存在"

    import shutil

    try:
        for filename in os.listdir(directory_path):
            file_path = os.path.join(directory_path, filename)
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        return True, "清理成功"
    except Exception as e:
        return False, f"清理失败: {str(e)}"


def open_directory(path):
    """
    在操作系统中打开文件夹
    """
    if not os.path.exists(path):
        return

    import platform
    import subprocess

    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":  # macOS
        subprocess.Popen(["open", path])
    else:  # Linux
        subprocess.Popen(["xdg-open", path])
