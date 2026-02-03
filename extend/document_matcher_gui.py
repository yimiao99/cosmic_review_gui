#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 文档目录与 Excel 表格匹配工具 - GUI 版本
支持简单模式和层级模式
作者: yimiao99
日期: 2025-01-15
"""
import os
import tkinter as tk
import time
from datetime import datetime
from tkinter import filedialog, messagebox, ttk, scrolledtext

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DRAG_DROP_AVAILABLE = True
except ImportError:  # 捕获 ImportError 而不是裸的 except
    DRAG_DROP_AVAILABLE = False

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from pathlib import Path
import threading
from hierarchical_matcher import HierarchicalMatcher  # 导入你的 HierarchicalMatcher
import pandas as pd
import traceback  # 用于打印更详细的错误信息
from matcher_config import MatcherConfig


# 从 hierarchical_matcher 导入 WIN32COM_AVAILABLE 标志
from hierarchical_matcher import WIN32COM_AVAILABLE


class DocumentMatcherGUI:
    """文档匹配器图形界面"""

    def __init__(self):
        """初始化 GUI"""
        if DRAG_DROP_AVAILABLE:
            try:
                # 优先使用 ttk.Window 以确保主题环境（包括所有布局和样式）正确初始化
                # 这在打包成 EXE 后比直接用 TkinterDnD.Tk + ttk.Style 更稳健
                self.root = ttk.Window(themename="cosmo")
                # 补充注入拖拽支持
                self.root.TkdndVersion = TkinterDnD._require(self.root)
            except Exception as e:
                print(f"Aesthetic Init Fallback: {e}")
                self.root = ttk.Window(themename="cosmo")
        else:
            self.root = ttk.Window(themename="cosmo")

        self.root.title("Word 目录与 Excel 匹配工具")
        self.root.geometry("1100x850")

        # 启动时最大化
        try:
            self.root.state("zoomed")
        except Exception:
            pass

        # 变量
        self.word_file = tk.StringVar()
        self.excel_file = tk.StringVar()
        self.output_file = tk.StringVar(value="匹配报告.xlsx")
        self.sheet_name = tk.StringVar(value="0")

        # 匹配模式：simple (简单) 或 hierarchical (层级)
        self.match_mode = tk.StringVar(value="simple")

        # 简单模式的列索引
        self.column_index = tk.IntVar(value=0)

        # 层级模式的列索引
        self.level1_col = tk.IntVar(value=1)
        self.level2_col = tk.IntVar(value=2)
        self.level3_col = tk.IntVar(value=3)

        self.fuzzy_match = tk.BooleanVar(value=True)
        self.threshold = tk.DoubleVar(value=0.8)

        # 全文搜索选项（简单模式）
        self.full_text_search = tk.BooleanVar(value=True)

        self.is_busy = False
        self.is_matching = False
        self.current_percent = 0
        self.matcher = None
        self.report = None
        self.excel_info = None

        self.excel_file.trace("w", self.on_excel_file_changed)
        self.match_mode.trace("w", self.on_mode_changed)

        self.setup_ui()

    def setup_ui(self):
        """设置用户界面"""
        # 标题区域 - 增加底部留白
        # 增加顶部 padding (20 -> 40) 让标题整体下移
        title_frame = ttk.Frame(self.root, padding=(20, 30, 20, 10))
        title_frame.pack(fill=X)

        title_label = ttk.Label(
            title_frame,
            text="📄 Word 目录与 Excel 匹配工具",
            font=("Microsoft YaHei UI", 20, "bold"),
            bootstyle="primary",
            anchor="center",
        )
        title_label.pack(fill=X)

        subtitle_label = ttk.Label(
            title_frame,
            text="支持简单匹配和层级匹配两种模式"
            + (" | 支持拖拽上传" if DRAG_DROP_AVAILABLE else ""),
            font=("Microsoft YaHei UI", 10),
            bootstyle="secondary",
            anchor="center",
        )
        subtitle_label.pack(pady=(5, 50), fill=X)  # 副标题与下方内容不再有额外间距

        # 主容器 - 使用两栏布局
        # 增加主容器内边距，使整体内容向内收缩，视觉上“外边框变大”
        main_content = ttk.Frame(self.root, padding=(30, 10, 30, 20))
        main_content.pack(fill=BOTH, expand=YES)

        # 比例分配：左右各占 60% : 40% (3:2)
        main_content.columnconfigure(0, weight=5)
        main_content.columnconfigure(1, weight=2)
        main_content.rowconfigure(0, weight=1)

        # === 左侧面板：配置与设置 ===
        left_panel = ttk.Frame(main_content)
        left_panel.grid(row=0, column=0, sticky=NSEW, padx=(0, 15))  # 右侧间距

        # --- 布局调整：先放置底部按钮，再放置上方滚动区域 ---
        # 这样可以确保底部按钮始终可见，不会被滚动区域挤出屏幕

        # 1. 创建底部按钮区域 (固定在底部)
        left_bottom_frame = ttk.Frame(left_panel)
        left_bottom_frame.pack(side=BOTTOM, fill=X, pady=(10, 0))

        # 2. 创建一个容器用于放滚动区域 (占用剩余空间)
        left_scroll_container = ttk.Frame(left_panel)
        left_scroll_container.pack(side=TOP, fill=BOTH, expand=YES)

        # 3. 滚动区域配置 (放入 left_scroll_container)
        left_scroll_canvas = tk.Canvas(
            left_scroll_container, highlightthickness=0, borderwidth=0
        )
        left_scrollbar = ttk.Scrollbar(
            left_scroll_container, orient=VERTICAL, command=left_scroll_canvas.yview
        )

        # 滚动条布局
        left_scrollbar.pack(side=RIGHT, fill=Y)
        left_scroll_canvas.pack(side=LEFT, fill=BOTH, expand=YES)

        self.left_inner_frame = ttk.Frame(left_scroll_canvas)

        self.left_inner_frame.bind(
            "<Configure>",
            lambda e: left_scroll_canvas.configure(
                scrollregion=left_scroll_canvas.bbox("all")
            ),
        )

        # 鼠标滚轮支持
        def _on_mousewheel(event):
            left_scroll_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        left_scroll_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # 创建窗口并保存 ID
        self.canvas_window_id = left_scroll_canvas.create_window(
            (0, 0), window=self.left_inner_frame, anchor=NW
        )

        # 监听画布大小变化，调整内部 frame 宽度以自适应
        def _configure_canvas(event):
            canvas_width = event.width
            left_scroll_canvas.itemconfig(self.canvas_window_id, width=canvas_width)

        left_scroll_canvas.bind("<Configure>", _configure_canvas)

        left_scroll_canvas.configure(yscrollcommand=left_scrollbar.set)

        # 4. 配置内容填充到 self.left_inner_frame
        left_config_container = ttk.Frame(self.left_inner_frame)
        left_config_container.pack(
            fill=BOTH, expand=YES, pady=5, padx=5
        )  # 增加 padx 防止贴边

        self.setup_file_selection(left_config_container)
        self.setup_mode_selection(left_config_container)

        # 模式设置容器
        self.settings_container = ttk.Frame(left_config_container)
        self.settings_container.pack(fill=X, expand=YES)

        self.setup_simple_options(self.settings_container)
        self.setup_hierarchical_options(self.settings_container)

        self.setup_match_options(left_config_container)

        # 5. 按钮放在固定的底部区域 (left_bottom_frame)
        self.setup_action_buttons(left_bottom_frame)

        # === 右侧面板：结果与反馈 ===
        right_panel = ttk.Frame(main_content)
        right_panel.grid(row=0, column=1, sticky=NSEW)

        self.setup_results_panel(right_panel)

        # 底部状态栏
        self.setup_status_bar()

        # 初始化显示
        self.on_mode_changed()
        self.toggle_threshold()

    def setup_file_selection(self, parent):
        """设置文件选择区域"""
        file_frame = ttk.Labelframe(
            parent,
            text="📁 文件选择" + (" (支持拖拽)" if DRAG_DROP_AVAILABLE else ""),
            padding=35,  # 增加内部填充
            bootstyle="primary",
        )
        file_frame.pack(fill=BOTH, expand=YES, pady=10)

        # 注册整个区域的拖拽 (智能识别)
        if DRAG_DROP_AVAILABLE:
            try:
                file_frame.drop_target_register(DND_FILES)
                file_frame.dnd_bind("<<Drop>>", lambda e: self.on_drop(e, "auto"))
            except Exception:
                pass

        # Word 文件
        word_label = ttk.Label(
            file_frame, text="Word 文档:", font=("Microsoft YaHei UI", 9)
        )
        word_label.grid(row=0, column=0, sticky=W, pady=10)

        self.word_entry = ttk.Entry(file_frame, textvariable=self.word_file)
        self.word_entry.grid(row=0, column=1, padx=5, pady=10, sticky=EW)

        if DRAG_DROP_AVAILABLE:
            try:
                self.word_entry.drop_target_register(DND_FILES)
                self.word_entry.dnd_bind("<<Drop>>", lambda e: self.on_drop(e, "auto"))
            except Exception:
                pass

        word_btn = ttk.Button(
            file_frame,
            text="浏览",
            command=self.browse_word_file,
            bootstyle="info-outline",
            width=6,
        )
        word_btn.grid(row=0, column=2, pady=10)

        # Excel 文件
        excel_label = ttk.Label(
            file_frame, text="Excel 表:", font=("Microsoft YaHei UI", 9)
        )
        excel_label.grid(row=1, column=0, sticky=W, pady=10)

        self.excel_entry = ttk.Entry(file_frame, textvariable=self.excel_file)
        self.excel_entry.grid(row=1, column=1, padx=5, pady=10, sticky=EW)

        if DRAG_DROP_AVAILABLE:
            try:
                self.excel_entry.drop_target_register(DND_FILES)
                self.excel_entry.dnd_bind("<<Drop>>", lambda e: self.on_drop(e, "auto"))
            except Exception:
                pass

        excel_btn = ttk.Button(
            file_frame,
            text="浏览",
            command=self.browse_excel_file,
            bootstyle="info-outline",
            width=6,
        )
        excel_btn.grid(row=1, column=2, pady=10)

        # 输出文件
        output_label = ttk.Label(
            file_frame, text="输出报告:", font=("Microsoft YaHei UI", 9)
        )
        output_label.grid(row=2, column=0, sticky=W, pady=10)

        output_entry = ttk.Entry(file_frame, textvariable=self.output_file)
        output_entry.grid(row=2, column=1, padx=5, pady=10, sticky=EW)

        output_btn = ttk.Button(
            file_frame,
            text="浏览",
            command=self.browse_output_file,
            bootstyle="info-outline",
            width=6,
        )
        output_btn.grid(row=2, column=2, pady=10)

        file_frame.columnconfigure(1, weight=1)

    def setup_mode_selection(self, parent):
        """设置模式选择区域"""
        mode_frame = ttk.Labelframe(
            parent, text="🔧 匹配模式", padding=35, bootstyle="info"  # 增加内部填充
        )
        mode_frame.pack(fill=BOTH, expand=YES, pady=10, padx=5)

        # 简单模式
        simple_radio = ttk.Radiobutton(
            mode_frame,
            text="📝 简单模式（单列匹配）",
            variable=self.match_mode,
            value="simple",
            bootstyle="info",
        )
        simple_radio.pack(anchor=W, pady=(0, 5))  # 减少底部间距，拉近描述

        simple_desc = ttk.Label(
            mode_frame,
            text="  适用于：Excel 中有一列包含所有功能点，直接与 Word 目录匹配",
            font=("Microsoft YaHei UI", 8),
            bootstyle="secondary",
        )
        simple_desc.pack(anchor=W, padx=20)

        # 层级模式
        hierarchical_radio = ttk.Radiobutton(
            mode_frame,
            text="🏗️ 层级模式（多列层级匹配）",
            variable=self.match_mode,
            value="hierarchical",
            bootstyle="success",
        )
        hierarchical_radio.pack(
            anchor=W, pady=(10, 5)
        )  # 增加顶部间距，使其与上方内容分开

        hierarchical_desc = ttk.Label(
            mode_frame,
            text="  适用于：Excel 中有一二级模块列，与 Word 的 4.1/4.1.1/4.1.1.1 层级对应",
            font=("Microsoft YaHei UI", 8),
            bootstyle="secondary",
            wraplength=400,  # 增加自动换行，防止文字过长被截断
        )
        hierarchical_desc.pack(
            anchor=W, padx=20, pady=(0, 10)
        )  # 增加底部间距，防止贴底

    def setup_simple_options(self, parent):
        """设置简单模式选项"""
        self.simple_frame = ttk.Labelframe(
            parent, text="📋 简单模式设置", padding=35, bootstyle="info"  # 增加内部填充
        )
        self.simple_frame.pack(fill=BOTH, expand=YES, pady=10)

        # 工作表选择
        sheet_label = ttk.Label(
            self.simple_frame, text="Excel 工作表:", font=("Microsoft YaHei UI", 10)
        )
        sheet_label.grid(row=0, column=0, sticky=W, pady=10)

        self.simple_sheet_combo = ttk.Combobox(
            self.simple_frame, textvariable=self.sheet_name, width=20, state="readonly"
        )
        self.simple_sheet_combo.grid(row=0, column=1, sticky=W, padx=5, pady=10)
        self.simple_sheet_combo.bind(
            "<<ComboboxSelected>>", self.on_simple_sheet_changed
        )
        # Auto-save when sheet changes
        self.simple_sheet_combo.bind("<<ComboboxSelected>>", lambda e: [self.on_simple_sheet_changed(e), self.save_current_config(e)], add=True)

        # 列选择
        column_label = ttk.Label(
            self.simple_frame, text="功能点列:", font=("Microsoft YaHei UI", 10)
        )
        column_label.grid(row=1, column=0, sticky=W, pady=10)

        self.column_combo = ttk.Combobox(self.simple_frame, width=20, state="readonly", textvariable=self.column_name)
        self.column_combo.grid(row=1, column=1, sticky=W, padx=5, pady=10)
        self.column_combo.bind(
            "<<ComboboxSelected>>", lambda e: self.on_column_changed()
        )
        # Auto-save when column changes
        self.column_combo.bind("<<ComboboxSelected>>", lambda e: [self.on_column_changed(e), self.save_current_config(e)], add=True)

        full_text_check = ttk.Checkbutton(
            self.simple_frame,
            text="🔍 全文搜索（从整个文档搜索，而不仅是目录）",
            variable=self.full_text_search,
            bootstyle="success-round-toggle",
        )
        full_text_check.grid(row=2, column=0, columnspan=2, sticky=W, pady=10)

        # 信息提示
        self.simple_info_label = ttk.Label(
            self.simple_frame,
            text="💡 选择包含功能点的列",
            font=("Microsoft YaHei UI", 8),
            bootstyle="info",
            wraplength=400,
        )
        self.simple_info_label.grid(row=3, column=0, columnspan=2, sticky=W, pady=5)

    def setup_hierarchical_options(self, parent):
        """设置层级模式选项"""
        self.hierarchy_frame = ttk.Labelframe(
            parent,
            text="🏗️ 层级模式设置",
            padding=35,
            bootstyle="success",  # 增加内部填充
        )
        self.hierarchy_frame.pack(fill=BOTH, expand=YES, pady=10)

        # 工作表选择
        sheet_label = ttk.Label(
            self.hierarchy_frame, text="Excel 工作表:", font=("Microsoft YaHei UI", 10)
        )
        sheet_label.grid(row=0, column=0, sticky=W, pady=10)

        self.hier_sheet_combo = ttk.Combobox(
            self.hierarchy_frame,
            textvariable=self.sheet_name,
            width=20,
            state="readonly",
        )
        self.hier_sheet_combo.grid(row=0, column=1, sticky=W, padx=5, pady=10)
        self.hier_sheet_combo.bind("<<ComboboxSelected>>", self.on_hier_sheet_changed)
        # Auto-save when sheet changes
        self.hier_sheet_combo.bind("<<ComboboxSelected>>", lambda e: [self.on_hier_sheet_changed(e), self.save_current_config(e)], add=True)

        # 一级模块列
        level1_label = ttk.Label(
            self.hierarchy_frame, text="一级模块列:", font=("Microsoft YaHei UI", 10)
        )
        level1_label.grid(row=1, column=0, sticky=W, pady=10)

        self.level1_combo = ttk.Combobox(
            self.hierarchy_frame, width=20, state="readonly", textvariable=self.level1_name
        )
        self.level1_combo.grid(row=1, column=1, sticky=W, padx=5, pady=10)
        self.level1_combo.bind(
            "<<ComboboxSelected>>", lambda e: self.on_level_changed(1)
        )
        self.level1_combo.bind("<<ComboboxSelected>>", self.save_current_config, add=True)

        # 二级模块列
        level2_label = ttk.Label(
            self.hierarchy_frame, text="二级模块列:", font=("Microsoft YaHei UI", 10)
        )
        level2_label.grid(row=2, column=0, sticky=W, pady=10)

        self.level2_combo = ttk.Combobox(
            self.hierarchy_frame, width=20, state="readonly", textvariable=self.level2_name
        )
        self.level2_combo.grid(row=2, column=1, sticky=W, padx=5, pady=10)
        self.level2_combo.bind(
            "<<ComboboxSelected>>", lambda e: self.on_level_changed(2)
        )
        self.level2_combo.bind("<<ComboboxSelected>>", self.save_current_config, add=True)

        # 三级模块列
        level3_label = ttk.Label(
            self.hierarchy_frame, text="三级模块列:", font=("Microsoft YaHei UI", 10)
        )
        level3_label.grid(row=3, column=0, sticky=W, pady=10)

        self.level3_combo = ttk.Combobox(
            self.hierarchy_frame, width=20, state="readonly"
        )
        self.level3_combo.grid(row=3, column=1, sticky=W, padx=5, pady=10)
        self.level3_combo.bind(
            "<<ComboboxSelected>>", lambda e: self.on_level_changed(3)
        )

        # 信息提示
        self.hier_info_label = ttk.Label(
            self.hierarchy_frame,
            text="💡 请选择包含一二三级模块的列",
            font=("Microsoft YaHei UI", 8),
            bootstyle="info",
            wraplength=400,
        )
        self.hier_info_label.grid(row=4, column=0, columnspan=2, sticky=W, pady=5)

    def setup_match_options(self, parent):
        """设置匹配选项区域"""
        options_frame = ttk.Labelframe(
            parent, text="⚙️ 匹配选项", padding=35, bootstyle="secondary"  # 增加内部填充
        )
        options_frame.pack(fill=BOTH, expand=YES, pady=10)

        # 模糊匹配 - 增加上下间距
        fuzzy_check = ttk.Checkbutton(
            options_frame,
            text="启用模糊匹配",
            variable=self.fuzzy_match,
            command=self.toggle_threshold,
            bootstyle="success-round-toggle",
        )
        fuzzy_check.grid(
            row=0, column=0, columnspan=3, sticky=W, pady=(5, 15)
        )  # 增加底部间距

        # 相似度阈值 - 增加元素间距
        self.threshold_label = ttk.Label(
            options_frame,
            text="相似度阈值:",
            font=("Microsoft YaHei UI", 10),
            state=DISABLED,
        )
        self.threshold_label.grid(
            row=1, column=0, sticky=W, pady=(0, 5), padx=(0, 10)
        )  # 增加右侧间距

        self.threshold_scale = ttk.Scale(
            options_frame,
            from_=0.5,
            to=1.0,
            variable=self.threshold,
            orient=HORIZONTAL,
            length=200,
            state=DISABLED,
            bootstyle="success",
        )
        self.threshold_scale.grid(
            row=1, column=1, sticky=EW, padx=(0, 10), pady=(0, 5)
        )  # 增加右侧间距

        self.threshold_value_label = ttk.Label(
            options_frame,
            text="0.80",
            font=("Microsoft YaHei UI", 10, "bold"),
            bootstyle="success",
        )
        self.threshold_value_label.grid(row=1, column=2, sticky=W, pady=(0, 5))

        # 配置列权重，让滑块可以扩展
        options_frame.columnconfigure(1, weight=1)

        self.threshold.trace("w", self.update_threshold_label)

    def on_mode_changed(self, *args):
        """当模式改变时"""
        mode = self.match_mode.get()
        if mode == "simple":
            self.hierarchy_frame.pack_forget()
            self.simple_frame.pack(fill=BOTH, expand=YES, pady=10)
            self.output_file.set("匹配报告.xlsx")
        else:  # hierarchical
            self.simple_frame.pack_forget()
            self.hierarchy_frame.pack(fill=BOTH, expand=YES, pady=10)
            self.output_file.set("层级匹配报告.xlsx")

    def on_column_changed(self):
        """当列改变时（简单模式）"""
        selection = self.column_combo.get()
        if selection:
            col_idx = int(selection.split("]")[0].strip("["))
            self.column_index.set(col_idx)

    def on_level_changed(self, level):
        """当层级列改变时（层级模式）"""
        if level == 1:
            selection = self.level1_combo.get()
            if selection:
                col_idx = int(selection.split("]")[0].strip("["))
                self.level1_col.set(col_idx)
        elif level == 2:
            selection = self.level2_combo.get()
            if selection:
                col_idx = int(selection.split("]")[0].strip("["))
                self.level2_col.set(col_idx)
        elif level == 3:
            selection = self.level3_combo.get()
            if selection:
                col_idx = int(selection.split("]")[0].strip("["))
                self.level3_col.set(col_idx)

    def on_drop(self, event, file_type):
        """处理文件拖拽 (支持多文件智能识别)"""
        files = self.root.tk.splitlist(event.data)

        for file_path in files:
            file_path = file_path.strip("{}")
            file_extension = Path(file_path).suffix.lower()

            # 智能识别逻辑
            target_type = file_type
            if file_type == "auto":
                if file_extension in (".docx", ".doc"):
                    target_type = "word"
                elif file_extension in (".xlsx", ".xls"):
                    target_type = "excel"
                else:
                    continue  # 忽略不支持的文件

            if target_type == "word":
                # 允许 .docx 和 .doc 文件
                if file_extension in (".docx", ".doc"):
                    self.word_file.set(file_path)
                    self.update_status(f"已加载 Word 文档: {Path(file_path).name}")
                else:
                    if file_type != "auto":  # 只有明确指定类型时才报错
                        messagebox.showerror(
                            "错误", "请拖拽 .docx 或 .doc 格式的 Word 文档！"
                        )
            elif target_type == "excel":
                if file_extension in (".xlsx", ".xls"):
                    self.excel_file.set(file_path)
                    self.update_status(f"已加载 Excel 文件: {Path(file_path).name}")
                else:
                    if file_type != "auto":
                        messagebox.showerror(
                            "错误", "请拖拽 .xlsx 或 .xls 格式的 Excel 文件！"
                        )

    def on_excel_file_changed(self, *args):
        """当 Excel 文件改变时"""
        excel_path = self.excel_file.get()
        if excel_path and Path(excel_path).exists():
            self.load_excel_info(excel_path)

    def on_simple_sheet_changed(self, event=None):
        """当工作表改变时（简单模式）"""
        self.update_simple_column_list()

    def on_hier_sheet_changed(self, event=None):
        """当工作表改变时（层级模式）"""
        self.update_hier_column_list()

    def load_excel_info(self, excel_path):
        """加载 Excel 文件信息"""
        try:
            self.update_status("正在分析 Excel 文件...")

            xl_file = pd.ExcelFile(excel_path)
            sheet_names = xl_file.sheet_names

            self.excel_info = {"sheets": {}}

            for sheet_name in sheet_names:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
                self.excel_info["sheets"][sheet_name] = {
                    "columns": list(df.columns),
                    "row_count": len(df),
                }

            # 更新两个工作表下拉框
            self.simple_sheet_combo["values"] = sheet_names
            self.hier_sheet_combo["values"] = sheet_names

            if sheet_names:
                # 加载配置
                config = MatcherConfig.load()
                
                # 默认选中配置中的工作表，如果不存在则使用默认逻辑
                s_sheet_idx = config.get("process", {}).get("sheet_name", 2)
                h_sheet_idx = config.get("hierarchy", {}).get("sheet_name", 2)
                
                # Simple mode sheet
                if len(sheet_names) > s_sheet_idx:
                    self.simple_sheet_combo.set(sheet_names[s_sheet_idx])
                    self.sheet_name.set(sheet_names[s_sheet_idx]) # Update global var too? usually handled by mode change
                else:
                    self.simple_sheet_combo.set(sheet_names[0])
                    
                # Hierarchical mode sheet (might be different in config, though UI shares self.sheet_name for logic? Check on_mode_changed)
                # In current UI, on_mode_changed forces update. We should set the one corresponding to current mode or both.
                if len(sheet_names) > h_sheet_idx:
                    self.hier_sheet_combo.set(sheet_names[h_sheet_idx])
                else:
                    self.hier_sheet_combo.set(sheet_names[0])
                
                # 根据当前模式设置 self.sheet_name
                if self.match_mode.get() == "simple":
                    self.sheet_name.set(self.simple_sheet_combo.get())
                else:
                    self.sheet_name.set(self.hier_sheet_combo.get())
                    
                self.update_simple_column_list()
                self.update_hier_column_list()

            self.update_status(f"Excel 文件已加载：{len(sheet_names)} 个工作表")

        except Exception as e:
            self.update_status(f"加载 Excel 失败: {str(e)}")
            messagebox.showerror("错误", f"无法读取 Excel 文件：\n{str(e)}")

    def update_simple_column_list(self):
        """更新列列表（简单模式）"""
        if not self.excel_info:
            return

        sheet_name = self.sheet_name.get()
        if sheet_name not in self.excel_info["sheets"]:
            return

        columns = self.excel_info["sheets"][sheet_name]["columns"]
        row_count = self.excel_info["sheets"][sheet_name]["row_count"]

        column_options = [f"[{i}] {col}" for i, col in enumerate(columns)]
        self.column_combo["values"] = column_options

        if column_options:
            # 加载配置
            config = MatcherConfig.load()
            conf_col_idx = config.get("process", {}).get("column", 6)
            
            # 如果配置的索引有效，优先使用
            if len(column_options) > conf_col_idx:
                default_col_idx = conf_col_idx
            else:
                default_col_idx = 0
                
            self.column_combo.set(column_options[default_col_idx])

            # 提取索引
            col_idx = int(column_options[default_col_idx].split("]")[0].strip("["))
            self.column_index.set(col_idx)

        self.simple_info_label.config(
            text=f"ℹ️ 工作表 '{sheet_name}': {len(columns)} 列, {row_count} 行"
        )

    def update_hier_column_list(self):
        """更新列列表（层级模式）"""
        if not self.excel_info:
            return

        sheet_name = self.sheet_name.get()
        if sheet_name not in self.excel_info["sheets"]:
            return

        columns = self.excel_info["sheets"][sheet_name]["columns"]
        row_count = self.excel_info["sheets"][sheet_name]["row_count"]

        column_options = [f"[{i}] {col}" for i, col in enumerate(columns)]

        self.level1_combo["values"] = column_options
        self.level2_combo["values"] = column_options
        self.level3_combo["values"] = column_options

        # 加载配置
        config = MatcherConfig.load()
        h_conf = config.get("hierarchy", {})
        l1_idx = h_conf.get("level1_col", 1)
        l2_idx = h_conf.get("level2_col", 2)
        l3_idx = h_conf.get("level3_col", 3)
        
        # 简单验证索引是否越界
        if l1_idx >= len(column_options): l1_idx = 0
        if l2_idx >= len(column_options): l2_idx = 1 if len(column_options) > 1 else 0
        if l3_idx >= len(column_options): l3_idx = 2 if len(column_options) > 2 else 0
        
        if column_options:
            self.level1_combo.set(column_options[l1_idx])
            self.level2_combo.set(column_options[l2_idx])
            self.level3_combo.set(column_options[l3_idx])
            self.level1_col.set(l1_idx)
            self.level2_col.set(l2_idx)
            self.level3_col.set(l3_idx)

        self.hier_info_label.config(
            text=f"ℹ️ 工作表 '{sheet_name}': {len(columns)} 列, {row_count} 行"
        )

    def browse_word_file(self):
        """浏览 Word 文件"""
        filename = filedialog.askopenfilename(
            title="选择 Word 文档",
            # 允许选择 .docx 和 .doc 文件
            filetypes=[("Word 文档", "*.docx *.doc"), ("所有文件", "*.*")],
        )
        if filename:
            self.word_file.set(filename)

    def browse_excel_file(self):
        """浏览 Excel 文件"""
        filename = filedialog.askopenfilename(
            title="选择 Excel 表格",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if filename:
            self.excel_file.set(filename)

    def browse_output_file(self):
        """浏览输出文件"""
        filename = filedialog.asksaveasfilename(
            title="保存报告",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if filename:
            self.output_file.set(filename)

    def toggle_threshold(self):
        """切换阈值控件状态"""
        if self.fuzzy_match.get():
            self.threshold_label.config(state=NORMAL)
            self.threshold_scale.config(state=NORMAL)
        else:
            self.threshold_label.config(state=DISABLED)
            self.threshold_scale.config(state=DISABLED)

    def update_threshold_label(self, *args):
        """更新阈值标签"""
        value = self.threshold.get()
        self.threshold_value_label.config(text=f"{value:.2f}")

    def setup_action_buttons(self, parent):
        """设置操作按钮区域"""
        button_frame = ttk.Frame(
            parent, padding=(0, 25, 0, 10)
        )  # 增加顶部间距，从40改为25
        button_frame.pack(fill=X, side=BOTTOM)

        # 使用网格以确保在窄屏幕下也能对齐，增加按钮间距
        self.match_btn = ttk.Button(
            button_frame,
            text="🚀 开始匹配",
            command=self.start_matching,
            bootstyle="success",
            width=15,
        )
        self.match_btn.grid(row=0, column=0, padx=3, pady=5, sticky=EW)  # 增加水平间距

        self.open_report_btn = ttk.Button(
            button_frame,
            text="📊 打开报告",
            command=self.open_report,
            bootstyle="primary",
            width=12,
            state=DISABLED,
        )
        self.open_report_btn.grid(
            row=0, column=1, padx=3, pady=5, sticky=EW
        )  # 增加水平间距

        clear_btn = ttk.Button(
            button_frame,
            text="🗑️ 清除",
            command=self.clear_all,
            bootstyle="warning-outline",
            width=10,
        )
        clear_btn.grid(row=0, column=2, padx=3, pady=5, sticky=EW)  # 增加水平间距

        button_frame.columnconfigure((0, 1, 2), weight=1)

    def setup_results_panel(self, parent):
        """设置结果显示面板"""
        # 使用 Labelframe 添加边框
        # 增加 padding (15 -> 35) 使外边框看起来更大，与左侧保持一致
        results_frame = ttk.Labelframe(
            parent, text="📊 匹配结果", padding=35, bootstyle="info"
        )
        # 调整 pady 以对齐左侧的“文件选择”区域 (左侧容器pady=5 + 内部pady=10 = 15)
        results_frame.pack(fill=BOTH, expand=YES, pady=(15, 10))

        # 使用grid布局确保文本区域能充分扩展
        results_frame.rowconfigure(3, weight=1)  # 文本区域所在行可扩展
        results_frame.columnconfigure(0, weight=1)

        # 统计信息区域 - 紧凑展示，使用2x2网格
        stats_outer = ttk.Frame(results_frame)
        stats_outer.grid(row=0, column=0, sticky=EW, pady=(0, 10))

        self.stat_cards = {}
        stats = [
            ("Excel功能点", "primary"),
            ("精确匹配", "success"),
            ("模糊匹配", "warning"),
            ("缺失项", "danger"),
        ]

        for i, (name, style) in enumerate(stats):
            row_idx = i // 2
            col_idx = i % 2
            card = self.create_stat_card(stats_outer, name, "0", style)
            card.grid(row=row_idx, column=col_idx, padx=2, pady=2, sticky=NSEW)
            self.stat_cards[name] = card

        stats_outer.columnconfigure(0, weight=1)
        stats_outer.columnconfigure(1, weight=1)

        # === 进度信息区域 ===
        progress_container = ttk.Frame(results_frame)
        progress_container.grid(row=1, column=0, sticky=EW, pady=(0, 10))

        # 1. 实时进度 (匹配中)
        self.realtime_progress_frame = ttk.Frame(progress_container)

        realtime_label = ttk.Label(
            self.realtime_progress_frame,
            text="进度:",
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        realtime_label.pack(side=LEFT, padx=(0, 8))

        self.realtime_progress = ttk.Progressbar(
            self.realtime_progress_frame, mode="determinate", bootstyle="info-striped"
        )
        self.realtime_progress.pack(side=LEFT, fill=X, expand=YES, padx=(0, 8))

        self.realtime_progress_label = ttk.Label(
            self.realtime_progress_frame,
            text="0%",
            font=("Microsoft YaHei UI", 9, "bold"),
            bootstyle="info",
        )
        self.realtime_progress_label.pack(side=LEFT)

        # 2. 匹配率 (完成后)
        self.final_progress_frame = ttk.Frame(progress_container)
        self.final_progress_frame.pack(fill=X)

        final_label = ttk.Label(
            self.final_progress_frame,
            text="匹配率:",
            font=("Microsoft YaHei UI", 9, "bold"),
        )
        final_label.pack(side=LEFT, padx=(0, 8))

        self.match_progress = ttk.Progressbar(
            self.final_progress_frame, mode="determinate", bootstyle="success-striped"
        )
        self.match_progress.pack(side=LEFT, fill=X, expand=YES, padx=(0, 8))

        self.match_rate_label = ttk.Label(
            self.final_progress_frame,
            text="0%",
            font=("Microsoft YaHei UI", 10, "bold"),
            bootstyle="success",
        )
        self.match_rate_label.pack(side=LEFT)

        # 3. 时间信息
        # 修改父容器为 progress_container，避免在 results_frame (grid布局) 中使用 pack
        self.time_info_frame = ttk.Frame(progress_container)

        self.elapsed_time_label = ttk.Label(
            self.time_info_frame,
            text="⏱️ 已用 00:00",
            font=("Microsoft YaHei UI", 9),
            bootstyle="secondary",
        )
        self.elapsed_time_label.pack(side=LEFT, padx=(0, 15))
        self.remaining_time_label = ttk.Label(
            self.time_info_frame,
            text="⌛ 剩余 --:--",
            font=("Microsoft YaHei UI", 9),
            bootstyle="secondary",
        )
        self.remaining_time_label.pack(side=LEFT)

        # === 详细文本区域 - 使用grid确保充分扩展 ===
        text_label = ttk.Label(
            results_frame,
            text="匹配详情:",
            font=("Microsoft YaHei UI", 9, "bold"),
            bootstyle="info",
        )
        text_label.grid(row=2, column=0, sticky=W, pady=(10, 5))

        # 文本框容器 - 关键：使用grid并设置row weight
        text_frame = ttk.Frame(results_frame, relief=tk.SUNKEN, borderwidth=1)
        text_frame.grid(row=3, column=0, sticky=NSEW)  # NSEW确保四个方向都扩展

        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=RIGHT, fill=Y)

        self.result_text = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Consolas", 9),
            yscrollcommand=scrollbar.set,
            bg="#ffffff",
            padx=10,
            pady=10,
            relief=tk.FLAT,
            borderwidth=0,
        )
        self.result_text.pack(side=LEFT, fill=BOTH, expand=YES)
        scrollbar.config(command=self.result_text.yview)

        # 标签配置
        self.result_text.tag_config(
            "title", foreground="#2196F3", font=("Microsoft YaHei UI", 10, "bold")
        )
        self.result_text.tag_config(
            "success", foreground="#4CAF50", font=("Consolas", 9, "bold")
        )
        self.result_text.tag_config(
            "warning", foreground="#FF9800", font=("Consolas", 9, "bold")
        )
        self.result_text.tag_config(
            "error", foreground="#F44336", font=("Consolas", 9, "bold")
        )
        self.result_text.tag_config("info", foreground="#607D8B", font=("Consolas", 9))

    def create_stat_card(self, parent, title, value, style):
        """创建统计卡片"""
        card = ttk.Frame(parent, bootstyle=style, relief=RAISED, padding=6)

        title_label = ttk.Label(
            card,
            text=title,
            font=("Microsoft YaHei UI", 9),
            bootstyle=f"{style}-inverse",
        )
        title_label.pack(pady=(1, 0))

        value_label = ttk.Label(
            card,
            text=value,
            font=("Microsoft YaHei UI", 18, "bold"),
            bootstyle=f"{style}-inverse",
        )
        value_label.pack(pady=(0, 1))

        card.value_label = value_label
        return card

    def setup_status_bar(self):
        """设置状态栏"""
        status_frame = ttk.Frame(self.root, relief=SUNKEN, padding=5)
        status_frame.pack(side=BOTTOM, fill=X)

        self.status_label = ttk.Label(
            status_frame,
            text="准备就绪 | 选择匹配模式开始"
            + (" | 支持拖拽上传" if DRAG_DROP_AVAILABLE else ""),
            font=("Microsoft YaHei UI", 9),
            bootstyle="secondary",
        )
        self.status_label.pack(side=LEFT)

        self.progress_bar = ttk.Progressbar(
            status_frame, mode="indeterminate", bootstyle="info", length=200
        )

    def start_matching(self):
        """开始匹配"""
        if not self.word_file.get():
            messagebox.showerror("错误", "请选择 Word 文档！")
            return

        if not self.excel_file.get():
            messagebox.showerror("错误", "请选择 Excel 功能表！")
            return

        word_path = Path(self.word_file.get())
        if not word_path.exists():
            messagebox.showerror("错误", "Word 文档不存在！")
            return

        # 检查 Word 文件类型，如果是 .doc 但 pywin32 不可用，则提前警告
        if word_path.suffix.lower() == ".doc" and not WIN32COM_AVAILABLE:
            messagebox.showerror(
                "错误",
                "检测到 Word 文档为 .doc 格式，但无法进行转换。\n"
                "请确保在 Windows 环境下运行，并已安装 'pywin32' 库 (pip install pywin32)，\n"
                "或手动将 .doc 文件转换为 .docx 格式。",
            )
            return

        if not Path(self.excel_file.get()).exists():
            messagebox.showerror("错误", "Excel 表格不存在！")
            return

        # 显示实时进度条和时间信息，隐藏完成度进度条
        if hasattr(self, "realtime_progress_frame"):
            self.realtime_progress_frame.pack(fill=X, pady=5)
        if hasattr(self, "time_info_frame"):
            self.time_info_frame.pack(fill=X, pady=(0, 5))
        if hasattr(self, "final_progress_frame"):
            self.final_progress_frame.pack_forget()

        # 保存配置
        try:
            config = MatcherConfig.load() # Load existing to preserve other keys if any
            
            # Get current sheet index (helper function or simple lookup)
            # Assuming files are loaded so excel_info is populated
            if self.excel_info:
                current_sheet = self.sheet_name.get()
                # Find index of current sheet
                # Note: simple_sheet_combo and hier_sheet_combo have same values
                sheet_names = self.simple_sheet_combo["values"]
                sheet_idx = sheet_names.index(current_sheet) if current_sheet in sheet_names else 0
                
                config["hierarchy"]["sheet_name"] = sheet_idx
                config["process"]["sheet_name"] = sheet_idx
            
            config["hierarchy"]["level1_col"] = self.level1_col.get()
            config["hierarchy"]["level2_col"] = self.level2_col.get()
            config["hierarchy"]["level3_col"] = self.level3_col.get()
            config["process"]["column"] = self.column_index.get()
            
            MatcherConfig.save(config)
            print("Configuration saved.")
        except Exception as e:
            print(f"Error saving config: {e}")

        # 重置实时进度条
        if hasattr(self, "realtime_progress"):
            self.realtime_progress["value"] = 0
        if hasattr(self, "realtime_progress_label"):
            self.realtime_progress_label.config(text="0%")

        # 重置时间
        self.start_time = time.time()
        if hasattr(self, "elapsed_time_label"):
            self.elapsed_time_label.config(text="已用时间: 00:00")
        if hasattr(self, "remaining_time_label"):
            self.remaining_time_label.config(text="预计剩余: --:--")

        # 重置完成度进度条
        if hasattr(self, "match_progress"):
            self.match_progress["value"] = 0
        if hasattr(self, "match_rate_label"):
            self.match_rate_label.config(text="0%")

        thread = threading.Thread(target=self.perform_matching)
        thread.daemon = True
        thread.start()

    def perform_matching(self):
        """执行匹配操作"""
        try:
            # 禁用按钮和显示进度
            self.root.after(0, lambda: self.set_busy_state(True))
            self.root.after(0, lambda: self.update_status("正在匹配中..."))

            self.is_matching = True
            self.current_percent = 0
            self.root.after(1000, self.update_timer_loop)

            # 创建进度回调函数
            def update_progress(percent, message):
                self.root.after(
                    0, lambda: self.update_matching_progress(percent, message)
                )

            # 创建匹配器
            self.matcher = HierarchicalMatcher(
                fuzzy_match=self.fuzzy_match.get(), threshold=self.threshold.get()
            )

            # 根据模式执行匹配
            mode = self.match_mode.get()

            if mode == "simple":
                # 简单匹配模式
                full_text = self.full_text_search.get()

                self.report = self.matcher.match_documents(
                    self.word_file.get(),
                    self.excel_file.get(),
                    sheet_name=self.sheet_name.get(),
                    column=self.column_index.get(),
                    full_text_search=full_text,
                    progress_callback=update_progress,
                )
            else:
                # 层级匹配模式
                self.report = self.matcher.match_hierarchical_documents(
                    self.word_file.get(),
                    self.excel_file.get(),
                    sheet_name=self.sheet_name.get(),
                    level1_col=self.level1_col.get(),
                    level2_col=self.level2_col.get(),
                    level3_col=self.level3_col.get(),
                    progress_callback=update_progress,
                )

            if self.report is None:
                self.root.after(
                    0, lambda: self.update_status("错误：匹配失败，请检查文件格式")
                )
                self.root.after(0, lambda: self.set_busy_state(False))
                self.root.after(0, self.show_match_failure_message)
                return

            # 保存报告
            self.root.after(0, lambda: self.update_status("正在保存报告..."))
            actual_output_file = self.matcher.save_report(
                self.report, self.output_file.get()
            )
            self.actual_output_file = (
                actual_output_file  # 记录实际保存的文件名（处理副本情况）
            )

            self.root.after(0, lambda: self.update_matching_progress(100, "完成！"))

            # 显示结果
            self.root.after(0, self.display_results)
            self.root.after(
                0,
                lambda: self.update_status(
                    f"✅ 匹配完成！报告已保存：{actual_output_file}"
                ),
            )
            self.root.after(0, lambda: self.set_busy_state(False))

            # 显示完成消息
            self.root.after(0, self.show_success_message)

        except PermissionError as e:
            self.is_matching = False
            error_message = (
                f"无法保存报告文件，文件可能被其他程序占用。\n请关闭 Excel 后重试。"
            )
            self.root.after(
                0, lambda msg=error_message: self.update_status(f"错误: {msg}")
            )
            self.root.after(0, lambda: self.set_busy_state(False))
            self.root.after(
                0, lambda msg=error_message: messagebox.showerror("权限错误", msg)
            )
        except ImportError as e:  # 捕获 HierarchicalMatcher 抛出的 ImportError
            self.is_matching = False
            error_message = (
                f"Word 文档转换失败：{e}\n"
                f"如果您的 Word 文档是 .doc 格式，请确保在 Windows 环境下运行，并已安装 'pywin32' 库 (pip install pywin32)。"
            )
            self.root.after(
                0, lambda msg=error_message: self.update_status(f"错误: {msg}")
            )
            self.root.after(0, lambda: self.set_busy_state(False))
            self.root.after(
                0, lambda msg=error_message: messagebox.showerror("文件转换错误", msg)
            )
        except Exception as e:
            self.is_matching = False
            error_detail = traceback.format_exc()
            error_msg = str(e)
            print(f"错误详情：\n{error_detail}")
            self.root.after(0, lambda msg=error_msg: self.update_status(f"错误: {msg}"))
            self.root.after(0, lambda: self.set_busy_state(False))
            self.root.after(0, lambda msg=error_msg: self.show_error_message(msg))
        finally:
            self.is_matching = False

    def update_matching_progress(self, percent, message):
        """更新匹配进度（实时进度条）"""
        self.current_percent = percent
        # 显示实时进度条和时间信息，隐藏完成度进度条
        if hasattr(self, "realtime_progress_frame") and hasattr(
            self, "final_progress_frame"
        ):
            if percent < 100:
                self.realtime_progress_frame.pack(fill=X)
                if hasattr(self, "time_info_frame"):
                    self.time_info_frame.pack(fill=X, pady=(0, 5))
                self.final_progress_frame.pack_forget()
            else:
                # 匹配完成，隐藏实时进度条和时间信息
                self.realtime_progress_frame.pack_forget()
                if hasattr(self, "time_info_frame"):
                    self.time_info_frame.pack_forget()
                self.final_progress_frame.pack(fill=X)

        # 更新实时进度条
        if hasattr(self, "realtime_progress"):
            self.realtime_progress["value"] = percent

        # 更新实时进度百分比标签
        if hasattr(self, "realtime_progress_label"):
            self.realtime_progress_label.config(text=f"{percent}%")

            # 根据进度改变颜色
            if percent < 30:
                self.realtime_progress_label.configure(bootstyle="info")
                self.realtime_progress.configure(bootstyle="info-striped")
            elif percent < 70:
                self.realtime_progress_label.configure(bootstyle="warning")
                self.realtime_progress.configure(bootstyle="warning-striped")
            elif percent < 100:
                self.realtime_progress_label.configure(bootstyle="primary")
                self.realtime_progress.configure(bootstyle="primary-striped")
            else:
                self.realtime_progress_label.configure(bootstyle="success")
                self.realtime_progress.configure(bootstyle="success-striped")

        # 更新时间显示
        if hasattr(self, "start_time"):
            elapsed = time.time() - self.start_time
            elapsed_str = self.format_time(elapsed)
            if hasattr(self, "elapsed_time_label"):
                self.elapsed_time_label.config(text=f"⏱️ {elapsed_str}")

            if percent > 0 and percent < 100:
                total_est = elapsed / (percent / 100)
                remaining = total_est - elapsed
                remaining_str = self.format_time(remaining)
                if hasattr(self, "remaining_time_label"):
                    self.remaining_time_label.config(text=f"⌛ {remaining_str}")
            elif percent >= 100:
                if hasattr(self, "remaining_time_label"):
                    self.remaining_time_label.config(text="⌛ 00:00")

        # 更新状态信息
        self.update_status(f"[{percent}%] {message}")

    def show_success_message(self):
        """显示成功消息"""
        stats = self.report["statistics"]
        missing = stats["缺失项"]

        if missing == 0:
            message = f"🎉 太棒了！\n\n所有 Excel 功能点都在 Word 中找到了！\n\n完成度：{stats['完成度']}\n精确匹配：{stats['精确匹配']} 项\n模糊匹配：{stats['模糊匹配']} 项\n\n报告已保存: {self.output_file.get()}"
            messagebox.showinfo("匹配完成", message)
        else:
            message = f"⚠️ 匹配完成！\n\n发现 {missing} 个功能点在 Word 中缺失。\n\n完成度：{stats['完成度']}\n精确匹配：{stats['精确匹配']} 项\n模糊匹配：{stats['模糊匹配']} 项\n缺失项：{missing} 项\n\n报告已保存: {self.output_file.get()}"
            messagebox.showwarning("匹配完成", message)

    def show_match_failure_message(self):
        """显示匹配失败消息"""
        messagebox.showerror("错误", "匹配失败，请检查文件格式和内容！")

    def show_error_message(self, error_msg):
        """显示错误消息"""
        messagebox.showerror("错误", f"发生错误：\n{error_msg}")

    def display_results(self):
        """显示匹配结果"""
        if not self.report:
            return

        stats = self.report["statistics"]

        # 更新统计卡片
        self.stat_cards["Excel功能点"].value_label.config(
            text=str(stats["Excel功能点总数"])
        )
        self.stat_cards["精确匹配"].value_label.config(text=str(stats["精确匹配"]))
        self.stat_cards["模糊匹配"].value_label.config(text=str(stats["模糊匹配"]))
        self.stat_cards["缺失项"].value_label.config(text=str(stats["缺失项"]))

        # 确保显示的是最终完成度进度条
        if hasattr(self, "realtime_progress_frame"):
            self.realtime_progress_frame.pack_forget()
        if hasattr(self, "time_info_frame"):
            self.time_info_frame.pack_forget()
        if hasattr(self, "final_progress_frame"):
            self.final_progress_frame.pack(fill=X)

        # 更新匹配率（覆盖进度条的数值）
        match_rate = float(
            stats["匹配率"].replace("%", "")
        )  # 使用"匹配率"而不是"完成度"
        self.match_progress["value"] = match_rate
        self.match_rate_label.config(text=stats["匹配率"])

        # 根据匹配率改变进度条颜色
        if match_rate == 100:
            self.match_progress.configure(bootstyle="success-striped")
            self.match_rate_label.configure(bootstyle="success")
        elif match_rate >= 80:
            self.match_progress.configure(bootstyle="info-striped")
            self.match_rate_label.configure(bootstyle="info")
        elif match_rate >= 60:
            self.match_progress.configure(bootstyle="warning-striped")
            self.match_rate_label.configure(bootstyle="warning")
        else:
            self.match_progress.configure(bootstyle="danger-striped")
            self.match_rate_label.configure(bootstyle="danger")

        # 更新详细结果
        self.result_text.delete(1.0, tk.END)

        # 标题
        self.result_text.insert(tk.END, "=" * 65 + "\n")
        self.result_text.insert(tk.END, "匹配详情\n", "title")
        self.result_text.insert(tk.END, "=" * 65 + "\n\n")

        # 判断是层级模式还是简单模式
        is_hierarchical = self.match_mode.get() == "hierarchical"

        # 1. 缺失项
        if self.report["not_found_in_word"]:
            self.result_text.insert(
                tk.END,
                f"❌ Word 中缺失的功能点 ({len(self.report['not_found_in_word'])} 项):\n",
                "error",
            )
            for i, item in enumerate(self.report["not_found_in_word"][:10], 1):
                if is_hierarchical:
                    # 层级模式：显示层级路径和匹配失败详情
                    level1 = item.get("Excel一级模块", "")
                    level2 = item.get("Excel二级模块", "")
                    level3 = item.get("Excel三级模块", "")

                    excel_path_parts = []
                    if level1:
                        excel_path_parts.append(level1)
                    if level2:
                        excel_path_parts.append(level2)
                    if level3:
                        excel_path_parts.append(level3)
                    excel_path = " > ".join(excel_path_parts)

                    # 显示匹配到的 Word 路径（包含“缺失”标记）
                    w1 = item.get("Word一级标题", "❌ 缺失")
                    w2 = item.get("Word二级标题", "❌ 缺失")
                    w3 = item.get("Word三级标题", "❌ 缺失")
                    word_path = f"{w1} > {w2} > {w3}"

                    info = f"  {i}. Excel: {excel_path}\n"
                    info += f"     Word:  {word_path}\n"
                else:
                    # 简单模式
                    info = f"  {i}. {item.get('Excel功能点', '')}\n"

                self.result_text.insert(tk.END, info, "info")

            if len(self.report["not_found_in_word"]) > 10:
                self.result_text.insert(
                    tk.END,
                    f"  ... 还有 {len(self.report['not_found_in_word']) - 10} 项（详见报告）\n",
                    "info",
                )
            self.result_text.insert(tk.END, "\n")
        else:
            self.result_text.insert(
                tk.END, "🎉 太棒了！所有功能点都已匹配！\n\n", "success"
            )

        # 2. 模糊匹配项
        if self.report["fuzzy_matched"]:
            self.result_text.insert(
                tk.END,
                f"🔍 模糊匹配 ({len(self.report['fuzzy_matched'])} 项):\n",
                "warning",
            )
            for i, item in enumerate(self.report["fuzzy_matched"][:5], 1):
                if is_hierarchical:
                    # 层级模式
                    level1 = item.get("Excel一级模块", "")
                    level2 = item.get("Excel二级模块", "")
                    level3 = item.get("Excel三级模块", "")

                    excel_path_parts = []
                    if level1:
                        excel_path_parts.append(level1)
                    if level2:
                        excel_path_parts.append(level2)
                    if level3:
                        excel_path_parts.append(level3)

                    excel_path = " > ".join(excel_path_parts)

                    word_num = item.get("Word完整编号", "")
                    word_content = item.get(
                        "Word三级标题", ""
                    )  # 使用 Word三级标题，因为它代表了 Excel 三级模块匹配到的 Word 内容
                    similarity = item.get("相似度", "")

                    info = f"  {i}. Excel: {excel_path}\n"
                    info += f"     Word: {word_num} {word_content}\n"
                    info += f"     相似度: {similarity}\n"
                else:
                    # 简单模式
                    excel_item = item.get("Excel功能点", "")
                    word_item = item.get("Word匹配项", "")
                    similarity = item.get("相似度", "")
                    location = item.get("位置", "")

                    info = f"  {i}. Excel: {excel_item}\n"
                    info += f"     Word: {word_item}\n"
                    info += f"     位置: {location}\n" if location else ""
                    info += f"     相似度: {similarity}\n"

                self.result_text.insert(tk.END, info, "info")

            if len(self.report["fuzzy_matched"]) > 5:
                self.result_text.insert(
                    tk.END,
                    f"  ... 还有 {len(self.report['fuzzy_matched']) - 5} 项（详见报告）\n",
                    "info",
                )
            self.result_text.insert(tk.END, "\n")

        # 3. 精确匹配项
        if self.report["exact_matched"]:
            self.result_text.insert(
                tk.END,
                f"✅ 精确匹配 ({len(self.report['exact_matched'])} 项)\n",
                "success",
            )
            for i, item in enumerate(self.report["exact_matched"][:5], 1):
                if is_hierarchical:
                    # 层级模式
                    level1 = item.get("Excel一级模块", "")
                    level2 = item.get("Excel二级模块", "")
                    level3 = item.get("Excel三级模块", "")

                    path_parts = []
                    if level1:
                        path_parts.append(level1)
                    if level2:
                        path_parts.append(level2)
                    if level3:
                        path_parts.append(level3)

                    path = " > ".join(path_parts)
                    word_num = item.get("Word完整编号", "")

                    info = f"  {i}. {path} ← {word_num}\n"
                else:
                    # 简单模式
                    excel_item = item.get("Excel功能点", "")
                    location = item.get("位置", "")

                    info = f"  {i}. {excel_item}"
                    if location:
                        info += f" [{location}]"
                    info += "\n"

                self.result_text.insert(tk.END, info, "info")

            if len(self.report["exact_matched"]) > 5:
                self.result_text.insert(
                    tk.END,
                    f"  ... 还有 {len(self.report['exact_matched']) - 5} 项（详见报告）\n",
                    "info",
                )

        self.open_report_btn.config(state=NORMAL)

    def open_report(self):
        """打开报告文件"""
        import os
        import platform

        # 优先打开实际保存的文件，如果没有则打开设定的文件名
        output_path = getattr(self, "actual_output_file", self.output_file.get())
        if not Path(output_path).exists():
            messagebox.showerror("错误", "报告文件不存在！")
            return

        try:
            if platform.system() == "Windows":
                os.startfile(output_path)
            elif platform.system() == "Darwin":
                os.system(f'open "{output_path}"')
            else:
                os.system(f'xdg-open "{output_path}"')
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件：\n{str(e)}")

    def clear_all(self):
        """清除所有内容"""
        self.word_file.set("")
        self.excel_file.set("")
        self.match_mode.set("simple")
        self.output_file.set("匹配报告.xlsx")
        self.sheet_name.set("0")
        self.column_index.set(0)
        self.level1_col.set(1)
        self.level2_col.set(2)
        self.level3_col.set(3)
        self.fuzzy_match.set(True)
        self.threshold.set(0.8)
        self.full_text_search.set(True)
        self.toggle_threshold()

        if hasattr(self, "elapsed_time_label"):
            self.elapsed_time_label.config(text="已用时间: 00:00")
        if hasattr(self, "remaining_time_label"):
            self.remaining_time_label.config(text="预计剩余: --:--")

        for card in self.stat_cards.values():
            card.value_label.config(text="0")

        self.match_progress["value"] = 0
        self.match_rate_label.config(text="0%")
        self.result_text.delete(1.0, tk.END)
        self.open_report_btn.config(state=DISABLED)

        self.excel_info = None
        self.simple_sheet_combo["values"] = []
        self.hier_sheet_combo["values"] = []
        self.column_combo["values"] = []
        self.level1_combo["values"] = []
        self.level2_combo["values"] = []
        self.level3_combo["values"] = []
        self.simple_info_label.config(text="💡 选择包含功能点的列")
        self.hier_info_label.config(text="💡 请选择包含一二三级模块的列")

        self.report = None
        self.update_status(
            "已清除 | 选择匹配模式开始"
            + (" | 支持拖拽上传" if DRAG_DROP_AVAILABLE else "")
        )

    def set_busy_state(self, busy):
        """设置忙碌状态"""
        state = DISABLED if busy else NORMAL
        self.match_btn.config(state=state)

        if busy:
            self.progress_bar.pack(side=RIGHT, padx=10)
            self.progress_bar.start(10)
        else:
            self.progress_bar.stop()
            self.progress_bar.pack_forget()

    def update_status(self, message):
        """更新状态栏"""
        self.status_label.config(text=message)

    def format_time(self, seconds):
        """格式化秒数为 MM:SS"""
        if seconds < 0:
            seconds = 0
        m, s = divmod(int(seconds), 60)
        return f"{m:02d}:{s:02d}"

    def update_timer_loop(self):
        """每秒更新一次时间的循环"""
        if not self.is_matching:
            return

        if hasattr(self, "start_time"):
            elapsed = time.time() - self.start_time
            elapsed_str = self.format_time(elapsed)
            if hasattr(self, "elapsed_time_label"):
                self.elapsed_time_label.config(text=f"⏱️ {elapsed_str}")

            percent = self.current_percent
            if 0 < percent < 100:
                total_est = elapsed / (percent / 100)
                remaining = max(0, total_est - elapsed)
                remaining_str = self.format_time(remaining)
                if hasattr(self, "remaining_time_label"):
                    self.remaining_time_label.config(text=f"⌛ {remaining_str}")
            elif percent >= 100:
                if hasattr(self, "remaining_time_label"):
                    self.remaining_time_label.config(text="⌛ 00:00")

        # 每一秒调度一次自己
        self.root.after(1000, self.update_timer_loop)

    def save_current_config(self, event=None):
        """Save current configuration to file"""
        if not self.excel_info:
            return
            
        try:
            config = MatcherConfig.load()
            
            # Determine logic based on mode or just save everything visible
            # Sheet Index
            current_sheet = self.sheet_name.get()
            sheet_values = self.simple_sheet_combo["values"]
            if current_sheet in sheet_values:
                sheet_idx = sheet_values.index(current_sheet)
                config["hierarchy"]["sheet_name"] = sheet_idx
                config["process"]["sheet_name"] = sheet_idx
            
            # Hierarchy Columns
            if self.level1_col.get() is not None:
                config["hierarchy"]["level1_col"] = self.level1_col.get()
            if self.level2_col.get() is not None:
                config["hierarchy"]["level2_col"] = self.level2_col.get()
            if self.level3_col.get() is not None:
                config["hierarchy"]["level3_col"] = self.level3_col.get()
                
            # Process Column
            if self.column_index.get() is not None:
                config["process"]["column"] = self.column_index.get()
                
            MatcherConfig.save(config)
            print("Config autosaved.")
            self.update_status("配置已自动保存")
            
        except Exception as e:
            print(f"Autosave failed: {e}")

    def run(self):
        """运行应用"""
        self.root.mainloop()



def main():
    """主函数"""
    app = DocumentMatcherGUI()
    app.run()


if __name__ == "__main__":
    main()
