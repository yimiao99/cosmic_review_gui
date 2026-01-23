from PySide6.QtWidgets import (
    QFrame,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QWidget,
    QProgressBar,
)
from PySide6.QtCore import Qt
import os
import subprocess


class ReReviewTaskCard(QFrame):
    """重评任务卡片"""

    def __init__(self, task_info, parent=None):
        super().__init__(parent)
        self.task_info = task_info
        self.setup_ui()
        self.update_theme_style()  # 初始化样式

    def showEvent(self, event):
        """每次显示时重新检查样式，防止主题切换后旧卡片没变色"""
        super().showEvent(event)
        self.update_theme_style()

    def update_theme_style(self):
        """动态更新主题样式"""
        from PySide6.QtWidgets import QApplication
        from extend.matcher_config import MatcherConfig

        # 从配置中检测主题模式
        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)

        # 备用检测：如果配置未设置，从样式表中检查
        if not is_dark:
            qss = QApplication.instance().styleSheet() or ""
            is_dark = "background-color: #1f2937" in qss

        if is_dark:
            # 深色模式颜色搭配
            card_bg = "#0f172a"
            card_border = "#1e293b"
            text_color = "#f1f5f9"
            meta_color = "#cbd5e1"

            # 按钮样式
            btn_bg = "#1e293b"
            btn_text = "#f1f5f9"
            btn_border = "#334155"
            btn_hover_bg = "#334155"

            # 打开目录按钮
            dir_btn_bg = "#0ea5e9"
            dir_btn_hover = "#0284c7"

            border_color = "#334155"
            progress_bg = "#1e293b"
            progress_chunk = "#0ea5e9"

            # 统计卡片样式 - 深色优化
            stats_bg = "#1e293b"
            stats_border = "#334155"
            divider_color = "#334155"

            # 不同数据的颜色区分
            new_color = "#86efac"  # 绿色
            reuse_color = "#93c5fd"  # 蓝色
            legacy_color = "#fbbf24"  # 黄色/橙色
            total_color = "#f1f5f9"  # 默认浅色
            days_color = "#fca5a5"  # 红色
        else:
            # 浅色模式颜色搭配
            card_bg = "#f8fafc"
            card_border = "#cbd5e1"
            text_color = "#0f172a"
            meta_color = "#475569"

            # 按钮样式
            btn_bg = "#ffffff"
            btn_text = "#0f172a"
            btn_border = "#cbd5e1"
            btn_hover_bg = "#f1f5f9"

            # 打开目录按钮
            dir_btn_bg = "#22d3ee"
            dir_btn_hover = "#06b6d4"

            border_color = "#cbd5e1"
            progress_bg = "#f1f5f9"
            progress_chunk = "#22d3ee"

            # 统计卡片样式 - 浅色优化
            stats_bg = "#f1f5f9"
            stats_border = "#cbd5e1"
            divider_color = "#cbd5e1"

            # 不同数据的颜色区分
            new_color = "#16a34a"  # 绿色
            reuse_color = "#2563eb"  # 蓝色
            legacy_color = "#d97706"  # 橙色
            total_color = "#0f172a"  # 深色
            days_color = "#dc2626"  # 红色

        self.setStyleSheet(
            f"""
            QFrame#ReReviewTaskCard {{
                background-color: {card_bg};
                border: 1px solid {card_border};
                border-radius: 12px;
                padding: 16px;
                margin-bottom: 8px;
            }}
            QLabel#ProjectTitle {{
                font-size: 18px;
                font-weight: bold;
                font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;
                color: {text_color};
                background: transparent;
            }}
            QLabel#TimeLabel {{
                font-size: 13px;
                font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;
                color: {meta_color};
                background: transparent;
            }}
            QLabel#StatusLabel {{
                font-size: 13px;
                font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;
                color: {meta_color};
                background: transparent;
                margin-top: 4px;
            }}
            QFrame#StatsFrame {{
                background-color: {stats_bg};
                border: 1px solid {stats_border};
                border-radius: 8px;
                padding: 12px;
            }}
            QFrame#StatsFrame QLabel {{
                font-size: 13px;
                font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;
                color: {text_color};
                background: transparent;
            }}
            QPushButton.FileBtn {{
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 13px;
                font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;
                background-color: {btn_bg};
                color: {btn_text};
                border: 1px solid {btn_border};
            }}
            QPushButton.FileBtn:hover {{
                background-color: {btn_hover_bg};
                border: 1px solid #60a5fa;
            }}
            QPushButton.FileBtn:pressed {{
                background-color: {btn_hover_bg};
                border: 1px solid #60a5fa;
            }}
            QPushButton#OpenDirBtn {{
                background-color: {dir_btn_bg};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 13px;
                font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;
            }}
            QPushButton#OpenDirBtn:hover {{
                background-color: {dir_btn_hover};
            }}
            QPushButton#OpenDirBtn:pressed {{
                background-color: {dir_btn_hover};
            }}
            QProgressBar {{
                border: 1px solid {border_color};
                border-radius: 6px;
                text-align: center;
                height: 6px;
                background-color: {progress_bg};
            }}
            QProgressBar::chunk {{
                background-color: {progress_chunk};
                border-radius: 6px;
            }}
            QFrame[objectName="DividerFrame"] {{
                background-color: {divider_color};
                border: none;
            }}
        """
        )

        # 为不同的标签设置颜色
        if hasattr(self, "label_new"):
            self.label_new.setStyleSheet(
                f"color: {new_color}; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
            )
        if hasattr(self, "label_new_ratio"):
            self.label_new_ratio.setStyleSheet(
                f"color: {new_color}; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
            )
        if hasattr(self, "label_reuse"):
            self.label_reuse.setStyleSheet(
                f"color: {reuse_color}; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
            )
        if hasattr(self, "label_reuse_ratio"):
            self.label_reuse_ratio.setStyleSheet(
                f"color: {reuse_color}; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
            )
        if hasattr(self, "label_legacy"):
            self.label_legacy.setStyleSheet(
                f"color: {legacy_color}; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
            )
        if hasattr(self, "label_legacy_ratio"):
            self.label_legacy_ratio.setStyleSheet(
                f"color: {legacy_color}; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
            )
        if hasattr(self, "label_total"):
            self.label_total.setStyleSheet(
                f"color: {total_color}; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
            )
        if hasattr(self, "label_total_ratio"):
            self.label_total_ratio.setStyleSheet(
                f"color: {total_color}; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
            )
        if hasattr(self, "label_submission"):
            self.label_submission.setStyleSheet(
                f"color: {days_color}; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
            )
        if hasattr(self, "label_eval"):
            self.label_eval.setStyleSheet(
                f"color: {days_color}; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
            )
        if hasattr(self, "label_reduction"):
            self.label_reduction.setStyleSheet(
                f"color: {days_color}; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
            )

    def setup_ui(self):
        self.setObjectName("ReReviewTaskCard")
        # 移除原有的 hardcoded setStyleSheet，改由 update_theme_style 管理

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        # Header: Title + Time
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(8)

        title_label = QLabel(self.task_info.get("project_name", "未命名项目"))
        title_label.setObjectName("ProjectTitle")
        title_label.setProperty("class", "task-title")

        time_label = QLabel(f"处理时间: {self.task_info.get('time', '--:--:--')}")
        time_label.setObjectName("TimeLabel")
        time_label.setProperty("class", "task-meta")

        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(time_label)
        layout.addLayout(header_layout)

        # Status Log Label
        self.log_label = QLabel(self.task_info.get("status_text", "准备就绪"))
        self.log_label.setObjectName("StatusLabel")
        layout.addWidget(self.log_label)

        # 统计数据卡片区域（初始隐藏）
        self.stats_widget = self._create_stats_widget()
        self.stats_widget.hide()
        layout.addWidget(self.stats_widget)

        layout.addSpacing(2)

        # File Links (Initialy hidden)
        self.files_widget = QWidget()
        files_layout = QHBoxLayout(self.files_widget)
        files_layout.setContentsMargins(0, 0, 0, 0)
        files_layout.setSpacing(10)

        btn_excel1 = QPushButton("评估报告 (公式版)")
        btn_excel1.setProperty("class", "FileBtn")
        btn_excel1.setCursor(Qt.PointingHandCursor)
        self.excel1_btn = btn_excel1  # Save for later update
        btn_excel1.clicked.connect(lambda: self.open_file(self.task_info.get("excel1")))

        btn_excel2 = QPushButton("重评回单 (结果版)")
        btn_excel2.setProperty("class", "FileBtn")
        btn_excel2.setCursor(Qt.PointingHandCursor)
        self.excel2_btn = btn_excel2  # Save for later update
        btn_excel2.clicked.connect(lambda: self.open_file(self.task_info.get("excel2")))

        btn_dir = QPushButton("打开结果目录")
        btn_dir.setObjectName("OpenDirBtn")
        btn_dir.setCursor(Qt.PointingHandCursor)
        btn_dir.clicked.connect(
            lambda: self.open_file(self.task_info.get("output_dir"))
        )

        files_layout.addWidget(btn_excel1)
        files_layout.addWidget(btn_excel2)
        files_layout.addStretch()
        files_layout.addWidget(btn_dir)

        layout.addWidget(self.files_widget)
        self.files_widget.hide()  # Hide until complete

        # Progress Bar - Move to bottom
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(6)  # Compact height
        self.progress_bar.setTextVisible(False)  # Cleaner look
        layout.addWidget(self.progress_bar)

    def _create_stats_widget(self):
        """创建统计数据显示卡片"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # 创建网格显示统计数据
        grid_layout = QVBoxLayout()
        grid_layout.setContentsMargins(12, 12, 12, 12)
        grid_layout.setSpacing(12)

        # 第一组：新增及其占比
        row1 = QHBoxLayout()
        row1.setSpacing(20)
        self.label_new = QLabel("新增: N/A")
        self.label_new_ratio = QLabel("新增占比: 0.0%")
        row1.addWidget(self.label_new)
        row1.addWidget(self.label_new_ratio)
        row1.addStretch()
        grid_layout.addLayout(row1)

        # 第二组：复用及其占比
        row2 = QHBoxLayout()
        row2.setSpacing(20)
        self.label_reuse = QLabel("复用: N/A")
        self.label_reuse_ratio = QLabel("复用占比: 0.0%")
        row2.addWidget(self.label_reuse)
        row2.addWidget(self.label_reuse_ratio)
        row2.addStretch()
        grid_layout.addLayout(row2)

        # 第三组：利旧及其占比
        row3 = QHBoxLayout()
        row3.setSpacing(20)
        self.label_legacy = QLabel("利旧: N/A")
        self.label_legacy_ratio = QLabel("利旧占比: 0.0%")
        row3.addWidget(self.label_legacy)
        row3.addWidget(self.label_legacy_ratio)
        row3.addStretch()
        grid_layout.addLayout(row3)

        # 第四行：合计及其占比
        row4 = QHBoxLayout()
        row4.setSpacing(20)
        self.label_total = QLabel("合计: N/A")
        self.label_total_ratio = QLabel("合计占比: 0.0%")
        row4.addWidget(self.label_total)
        row4.addWidget(self.label_total_ratio)
        row4.addStretch()
        grid_layout.addLayout(row4)

        # 分隔线
        divider = QFrame()
        divider.setFrameShape(QFrame.HLine)
        divider.setLineWidth(1)
        divider.setObjectName("DividerFrame")
        grid_layout.addWidget(divider)

        # 第五行：人天及核减比例
        row5 = QHBoxLayout()
        row5.setSpacing(20)
        self.label_submission = QLabel("送审人天: 0.00")
        self.label_eval = QLabel("核定人天: 0.00")
        self.label_reduction = QLabel("核减比例: 0.00%")
        row5.addWidget(self.label_submission)
        row5.addWidget(self.label_eval)
        row5.addWidget(self.label_reduction)
        row5.addStretch()
        grid_layout.addLayout(row5)

        # 设置卡片背景
        frame = QFrame()
        frame.setLayout(grid_layout)
        frame.setObjectName("StatsFrame")

        layout.addWidget(frame)

        return widget

    def update_stats(self, stats):
        """更新统计数据显示"""
        if not stats:
            return

        # 安全地获取值，处理N/A和转换类型
        def safe_get(val, default="N/A"):
            if val == "N/A" or val is None:
                return default
            return str(val)

        new_fp = safe_get(stats.get("new_fp", "N/A"), "0")
        reuse_fp = safe_get(stats.get("reuse_fp", "N/A"), "0")
        legacy_fp = safe_get(stats.get("legacy_fp", "N/A"), "0")
        total_fp = safe_get(stats.get("total_fp", "N/A"), "0")

        new_ratio = safe_get(stats.get("new_ratio", "0.0%"), "0.0%")
        reuse_ratio = safe_get(stats.get("reuse_ratio", "0.0%"), "0.0%")
        legacy_ratio = safe_get(stats.get("legacy_ratio", "0.0%"), "0.0%")
        total_ratio = safe_get(stats.get("total_ratio", "0.0%"), "0.0%")

        submission_days = safe_get(stats.get("submission_days", "0.00"), "0.00")
        eval_days = safe_get(stats.get("eval_days", "0.00"), "0.00")
        reduction_ratio = safe_get(stats.get("reduction_ratio", "0.00%"), "0.00%")

        # 更新标签
        self.label_new.setText(f"新增: {new_fp}")
        self.label_reuse.setText(f"复用: {reuse_fp}")
        self.label_legacy.setText(f"利旧: {legacy_fp}")
        self.label_total.setText(f"合计: {total_fp}")

        self.label_new_ratio.setText(f"新增占比: {new_ratio}")
        self.label_reuse_ratio.setText(f"复用占比: {reuse_ratio}")
        self.label_legacy_ratio.setText(f"利旧占比: {legacy_ratio}")
        self.label_total_ratio.setText(f"合计占比: {total_ratio}")

        self.label_submission.setText(f"送审人天: {submission_days}")
        self.label_eval.setText(f"核定人天: {eval_days}")
        self.label_reduction.setText(f"核减比例: {reduction_ratio}")

        # 显示统计卡片
        self.stats_widget.show()

    def update_progress(self, value):
        self.progress_bar.setValue(value)
        if value >= 100:
            self.set_completed()

    def set_completed(self, excel1=None, excel2=None):
        self.progress_bar.hide()
        self.log_label.hide()  # 隐藏"准备就绪"状态
        self.files_widget.show()
        if excel1:
            self.task_info["excel1"] = excel1
        if excel2:
            self.task_info["excel2"] = excel2

    def set_error(self, error_msg):
        self.progress_bar.setStyleSheet(
            "QProgressBar::chunk { background-color: #ef4444; }"
        )
        self.progress_bar.setFormat(f"错误: {error_msg}")

    def open_file(self, path):
        if not path or not os.path.exists(path):
            return

        try:
            if os.name == "nt":
                os.startfile(path)
            elif os.name == "posix":
                subprocess.call(["open", path])
            else:
                subprocess.call(["xdg-open", path])
        except Exception as e:
            print(f"无法打开路径: {path}, 错误: {e}")
