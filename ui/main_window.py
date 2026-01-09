import random
from datetime import datetime

from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QFrame,
    QScrollArea,
    QApplication,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QPalette
from .task_card import TaskCard
from .upload_dialog import UploadDialog  # ✅ 使用相对导入
from utils.path_utils import get_resource_path
from utils.themes import LIGHT_THEME, DARK_THEME
from utils.styles import apply_dark_title_bar


class CosmicMainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cosmic 智能审核队列")
        self.setMinimumSize(1200, 800)

        # 记录当前主题状态 (自适应系统默认)
        palette = QApplication.palette()
        window_color = palette.color(QPalette.Window)
        # 如果背景色亮度较低，则判定为深色模式
        self.is_dark_mode = window_color.lightness() < 128

        # 设置窗口图标
        # 如果你有 logo.png 请放在 ui 目录下，这里先写逻辑
        self.setWindowIcon(QIcon(get_resource_path("ui/logo.png")))

        # 主容器
        self.central_widget = QWidget()
        self.central_widget.setObjectName("CentralWidget")
        self.setCentralWidget(self.central_widget)
        layout = QVBoxLayout(self.central_widget)
        layout.setContentsMargins(0, 0, 0, 20)  # 顶部、左右铺满
        layout.setSpacing(20)

        # Header
        self.header = self._create_header()
        self.header.setObjectName("HeaderFrame")
        layout.addWidget(self.header)

        # 任务列表 (滚动区域 - 增加左右边距保持内部美观)
        scroll_container = QWidget()
        scroll_layout = QVBoxLayout(scroll_container)
        scroll_layout.setContentsMargins(20, 0, 20, 0)
        scroll_layout.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_content.setObjectName("ScrollContent")
        self.task_layout = QVBoxLayout(scroll_content)
        self.task_layout.setSpacing(20)

        # 初始不再加载模拟数据，保持界面清爽
        self.task_layout.addStretch()
        scroll.setWidget(scroll_content)
        scroll_layout.addWidget(scroll)
        layout.addWidget(scroll_container)

        # 应用初始主题
        self._apply_theme()

    def _apply_theme(self):
        """应用主题样式"""
        theme = DARK_THEME if self.is_dark_mode else LIGHT_THEME
        QApplication.instance().setStyleSheet(theme)

        # 应用原生标题栏深色模式
        apply_dark_title_bar(self, self.is_dark_mode)

        if self.is_dark_mode:
            if hasattr(self, "theme_btn"):
                self.theme_btn.setText("☀️ 白天模式")
        else:
            if hasattr(self, "theme_btn"):
                self.theme_btn.setText("🌙 深色模式")

    def toggle_theme(self):
        """切换主题"""
        self.is_dark_mode = not self.is_dark_mode
        self._apply_theme()

    def _create_header(self):
        """创建顶部 Header"""
        header = QFrame()
        header.setObjectName("HeaderFrame")
        header.setFixedHeight(80)  # 增加高度
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(30, 0, 30, 0)  # 左右留白让内容居中感更好

        # 左侧信息
        left_info = QVBoxLayout()
        left_info.setSpacing(4)
        left_info.setAlignment(Qt.AlignVCenter)  # 垂直居中

        title = QLabel("🚀 Cosmic 智能审核队列")
        title.setStyleSheet("font-size: 24px; font-weight: bold;")
        left_info.addWidget(title)

        header_layout.addLayout(left_info)
        header_layout.addStretch()

        # 右侧按钮组
        btn_group = QHBoxLayout()
        btn_group.setSpacing(10)
        btn_group.setAlignment(Qt.AlignVCenter)  # 垂直居中

        # 主题切换按钮
        self.theme_btn = QPushButton("🌙 深色模式")
        self.theme_btn.setObjectName("ThemeBtn")
        self.theme_btn.clicked.connect(self.toggle_theme)
        btn_group.addWidget(self.theme_btn)

        upload_btn = QPushButton("新建审核任务")
        upload_btn.setObjectName("UploadBtn")
        # 移除硬编码样式
        upload_btn.clicked.connect(self.show_upload_dialog)
        btn_group.addWidget(upload_btn)

        header_layout.addLayout(btn_group)

        return header

    def show_upload_dialog(self):
        dialog = UploadDialog(self)
        dialog.task_submitted.connect(self.add_new_task)  # 关键连接
        dialog.exec()

    def add_new_task(self, task_info):
        """动态添加新任务卡片，随机取色作为边框"""
        # 确定类型
        type_label = "结算"

        # 随机取色 (选取一些好看的 UI 色系)
        colors = [
            "#2563eb",
            "#10b981",
            "#f59e0b",
            "#a855f7",
            "#ef4444",
            "#06b6d4",
            "#ec4899",
        ]
        border_color = random.choice(colors)

        # 日志（1~6）
        timestamp = datetime.now().strftime("%H:%M:%S")

        # 处理初始日志，第一步现在交给 TaskCard 异步处理
        step1_status = "pending"
        step1_log = "等待比对任务开始..."
        if task_info.get("validation_results"):
            # 如果已经有结果（比如重载），则保持原有逻辑
            step1_status = "done"  # 示例简化

        # ✅ 6 个步骤，初始状态大部分为 done (模拟)
        steps = [
            ("模板校验", "pending"),
            ("空值检查", "pending"),
            ("送审比例", "pending"),
            ("附加值因子", "pending"),
            ("层级匹配", "pending"),
            ("功能过程", "pending"),
        ]

        logs = {
            1: step1_log,
            2: "等待执行...",
            3: "等待执行...",
            4: "等待执行...",
            5: "等待执行...",
            6: "等待执行...",
        }

        # 默认日志（显示在卡片底部）
        default_log = f"✅ 任务于 {timestamp} 提交，正在处理..."

        # 构造 task_data，完全匹配 mock_data 结构
        task_data = {
            "filename": task_info["filename"],
            "type_label": type_label,
            "days": task_info["days"],  # 字符串也可，UI 只做拼接
            "border_color": border_color,
            "steps": steps,  # ✅ 6 个元组，名称/状态对齐 mock
            "default_log": default_log,
            "logs": logs,  # ✅ 1~6 键完整
            "report_sections": [],  # 新任务暂无报告，但字段存在（兼容 ReportDialog）
            "raw_task_info": task_info,  # 保留原始数据
        }

        # 创建卡片并插入顶部
        card = TaskCard(task_data)
        self.task_layout.insertWidget(0, card)
