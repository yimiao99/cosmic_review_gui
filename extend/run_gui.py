#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速启动 GUI 版本
"""

import sys
from pathlib import Path

# 检查依赖
try:
    import ttkbootstrap
    import docx
    import pandas
    import openpyxl
except ImportError as e:
    print("缺少必要的依赖包！")
    print(f"错误: {e}")
    print("\n请运行以下命令安装依赖：")
    print("pip install -r requirements.txt")
    sys.exit(1)

# 启动 GUI

from document_matcher_gui import main

if __name__ == '__main__':
    print("正在启动 Word 目录与 Excel 匹配工具...")
    main()