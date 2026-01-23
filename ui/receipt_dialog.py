import os
import pandas as pd
from datetime import datetime
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QFrame,
    QFileDialog,
    QMessageBox,
    QWidget,
    QRadioButton,
    QButtonGroup,
    QScrollArea,
    QDateEdit,
    QGroupBox,
)
from PySide6.QtCore import Qt, Signal, QDate, QByteArray, QSize
from PySide6.QtGui import QIcon, QImage, QPixmap
from utils.path_utils import get_resource_path
from utils.themes import apply_dark_title_bar
from extend.matcher_config import MatcherConfig
from .upload_dialog import UploadAreaWidget, ElidedLabel

class ReceiptFileItem(QFrame):
    """回单文件项 - 模仿图1样式"""

    remove_clicked = Signal(str)

    def __init__(self, file_path, file_type, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.filename = os.path.basename(file_path)
        self.setProperty("class", "FileQueueItem")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 8, 12, 8)

        # 文件图标
        icon_label = QLabel("X")
        icon_label.setFixedSize(28, 28)
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet(
            f"""
            background: #217346;
            color: white; border-radius: 4px;
            font-size: 11px; font-weight: bold;
        """
        )
        layout.addWidget(icon_label)

        # 文件名
        name_label = ElidedLabel(self.filename)
        name_label.setStyleSheet("font-weight: 600; font-size: 12px; border: none; background: transparent;")
        layout.addWidget(name_label, stretch=1)

        # 角色标签
        self.role_label = QLabel(file_type)
        self.role_label.setProperty("class", "task-meta")
        self.role_label.setFixedWidth(80)  # 给一个固定宽度，防止压缩
        self.role_label.setAlignment(Qt.AlignCenter)
        self.role_label.setStyleSheet("""
            font-size: 11px; 
            font-weight: 600; 
            color: #10b981; 
            background: rgba(16, 185, 129, 0.1); 
            border-radius: 4px;
            padding: 2px 4px;
        """)
        layout.addWidget(self.role_label)

        # 删除按钮
        self.remove_btn = QPushButton()
        self.remove_btn.setFixedSize(28, 28)
        self.remove_btn.setCursor(Qt.PointingHandCursor)
        
        trash_icon_svg = """
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#ef4444" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
        </svg>
        """
        trash_qicon = QIcon(QPixmap.fromImage(QImage.fromData(QByteArray(trash_icon_svg.encode()))))
        self.remove_btn.setIcon(trash_qicon)
        self.remove_btn.setIconSize(QSize(16, 16))
        self.remove_btn.setStyleSheet("""
            QPushButton { background: #fee2e2; border-radius: 14px; border: none; }
            QPushButton:hover { background: #ef4444; }
        """)
        self.remove_btn.clicked.connect(lambda: self.remove_clicked.emit(self.file_path))
        layout.addWidget(self.remove_btn)

class ReceiptFileDialog(QDialog):
    """回单配置对话框 - 模仿新建任务样式"""
    
    generation_requested = Signal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("评估确认单配置")
        self.setMinimumSize(800, 700)
        self.setWindowIcon(QIcon(get_resource_path("ui/logo.png")))
        self.setObjectName("ReceiptFileDialog")
        
        self.file_queue = {} # path -> role
        
        self._init_ui()
        
        # 应用深色标题栏
        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)
        apply_dark_title_bar(self, is_dark)
        
    def _init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # 滚动区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        
        content_widget = QWidget()
        content_widget.setObjectName("ScrollContent")
        layout = QVBoxLayout(content_widget)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(25)
        
        # 1. 基础配置
        base_group = QGroupBox("⚙️ 基础配置")
        base_layout = QVBoxLayout(base_group)
        base_layout.setContentsMargins(25, 20, 25, 25)
        base_layout.setSpacing(15)
        
        # 项目名称 (强制标红提示)
        base_layout.addWidget(QLabel("项目名称 (必填):"))
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("请输入项目名称 (必填)")
        base_layout.addWidget(self.name_input)
        
        # 项目编号
        base_layout.addWidget(QLabel("项目编号:"))
        self.id_input = QLineEdit()
        self.id_input.setPlaceholderText("请输入项目编号")
        base_layout.addWidget(self.id_input)

        # 送审单位
        base_layout.addWidget(QLabel("送审单位:"))
        self.unit_input = QLineEdit()
        self.unit_input.setPlaceholderText("请输入送审单位")
        base_layout.addWidget(self.unit_input)

        # 送 审 人
        base_layout.addWidget(QLabel("送 审 人:"))
        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("请输入送审人")
        base_layout.addWidget(self.user_input)

        # 送审时间
        self.date_label = QLabel("送审时间:")
        base_layout.addWidget(self.date_label)
        self.date_input = QDateEdit()
        self.date_input.setCalendarPopup(True)
        self.date_input.setDate(QDate.currentDate())
        self.date_input.setFixedHeight(35)
        self.date_input.setObjectName("ReceiptDateInput")
        base_layout.addWidget(self.date_input)
        
        # 送审方式
        row_mode = QHBoxLayout()
        row_mode.addWidget(QLabel("送审方式:"))
        self.mode_group = QButtonGroup(self)
        self.radio_online = QRadioButton("线上")
        self.radio_email = QRadioButton("邮件")
        self.radio_online.setChecked(True)
        self.mode_group.addButton(self.radio_online, 0)
        self.mode_group.addButton(self.radio_email, 1)
        row_mode.addWidget(self.radio_online)
        row_mode.addWidget(self.radio_email)
        row_mode.addStretch()
        base_layout.addLayout(row_mode)
        
        layout.addWidget(base_group)
        
        # 2. 文件选择
        file_group = QGroupBox("📂 文件选择 (支持拖拽)")
        file_layout = QVBoxLayout(file_group)
        file_layout.setContentsMargins(25, 20, 25, 25)
        file_layout.setSpacing(15)
        
        self.upload_area = UploadAreaWidget()
        self.upload_area.files_dropped.connect(self._handle_files)
        file_layout.addWidget(self.upload_area)
        
        # 已上传列表
        list_header = QHBoxLayout()
        list_header.addWidget(QLabel("📋 已上传文件:"))
        list_header.addStretch()
        file_layout.addLayout(list_header)
        
        self.file_list_container = QWidget()
        self.file_list_layout = QVBoxLayout(self.file_list_container)
        self.file_list_layout.setContentsMargins(0, 0, 0, 0)
        self.file_list_layout.setSpacing(8)
        
        file_scroll = QScrollArea()
        file_scroll.setWidgetResizable(True)
        file_scroll.setFixedHeight(180)
        file_scroll.setWidget(self.file_list_container)
        file_scroll.setStyleSheet("QScrollArea { border: 1px solid #374151; border-radius: 8px; }")
        file_layout.addWidget(file_scroll)
        
        layout.addWidget(file_group)
        
        scroll.setWidget(content_widget)
        main_layout.addWidget(scroll)
        
        # 3. 底部按钮
        btn_layout = QHBoxLayout()
        btn_layout.setContentsMargins(20, 15, 20, 15)
        btn_layout.setSpacing(15)  # 按钮之间的间距
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.setFixedSize(120, 40)
        self.cancel_btn.clicked.connect(self.reject)
        
        self.start_btn = QPushButton("开始生成")
        self.start_btn.setFixedSize(120, 40)
        self.start_btn.setObjectName("UploadBtn") # 复用全局蓝色按钮样式
        self.start_btn.clicked.connect(self._on_start)
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.cancel_btn)
        btn_layout.addWidget(self.start_btn)
        main_layout.addLayout(btn_layout)

    def _handle_files(self, paths):
        for path in paths:
            if not path.endswith(".xlsx"):
                continue
            if path in self.file_queue:
                continue
                
            # 简单猜测角色
            fname = os.path.basename(path)
            role = "评估报告"
            if "认同" in fname or "确认" in fname or "结论" in fname:
                role = "评估认同表"
            
            # 如果已经有一个认同表了，这个设为报告，反之亦然
            roles_in_queue = list(self.file_queue.values())
            if role in roles_in_queue:
                role = "评估认同表" if role == "评估报告" else "评估报告"

            self.file_queue[path] = role
            item = ReceiptFileItem(path, role)
            item.remove_clicked.connect(self._remove_file)
            self.file_list_layout.insertWidget(0, item)

    def _remove_file(self, path):
        if path in self.file_queue:
            del self.file_queue[path]
            # 找到对应的项目并删除
            for i in range(self.file_list_layout.count()):
                item_widget = self.file_list_layout.itemAt(i).widget()
                if isinstance(item_widget, ReceiptFileItem) and item_widget.file_path == path:
                    item_widget.deleteLater()
                    break

    def _on_start(self):
        # 获取报告和认同表路径
        report_path = ""
        consent_path = ""
        for p, r in self.file_queue.items():
            if r == "评估报告": report_path = p
            elif r == "评估认同表": consent_path = p
            
        # 校验：项目名称和附件均为必填
        if not self.name_input.text().strip():
            QMessageBox.warning(self, "警告", "请输入项目名称")
            return
        if not report_path:
            QMessageBox.warning(self, "警告", "请上传评估报告 (Excel)")
            return
        if not consent_path:
            QMessageBox.warning(self, "警告", "请上传评估认同表 (Excel)")
            return
            
        data = {
            "project_name": self.name_input.text().strip(),
            "project_id": self.id_input.text().strip() or "N/A",
            "submission_unit": self.unit_input.text().strip() or "N/A",
            "submitter": self.user_input.text().strip() or "N/A",
            "submission_time": self.date_input.date().toString("yyyy-MM-dd"),
            "submission_mode": "线上" if self.radio_online.isChecked() else "邮件",
            "eval_report_path": report_path,
            "eval_consent_path": consent_path
        }
        self.generation_requested.emit(data)
        self.accept()
