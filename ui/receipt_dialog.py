import os
import re

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
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QImage, QPixmap
from utils.path_utils import get_resource_path
from utils.themes import apply_dark_title_bar
from utils.archive_utils import ArchiveUtils  # ✅ 导入压缩包处理工具
from extend.matcher_config import MatcherConfig
from .upload_dialog import UploadAreaWidget, ElidedLabel, DragOverlay


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
        name_label.setStyleSheet(
            "font-weight: 600; font-size: 12px; border: none; background: transparent;"
        )
        layout.addWidget(name_label, stretch=1)

        # 角色标签
        self.role_label = QLabel(file_type)
        self.role_label.setProperty("class", "task-meta")
        self.role_label.setFixedWidth(80)  # 给一个固定宽度，防止压缩
        self.role_label.setAlignment(Qt.AlignCenter)
        self.role_label.setStyleSheet(
            """
            font-size: 11px;
            font-weight: 600;
            color: #10b981;
            background: rgba(16, 185, 129, 0.1);
            border-radius: 4px;
            padding: 2px 4px;
        """
        )
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
        trash_qicon = QIcon(
            QPixmap.fromImage(QImage.fromData(QByteArray(trash_icon_svg.encode())))
        )
        self.remove_btn.setIcon(trash_qicon)
        self.remove_btn.setIconSize(QSize(16, 16))
        self.remove_btn.setStyleSheet(
            """
            QPushButton { background: #fee2e2; border-radius: 14px; border: none; }
            QPushButton:hover { background: #ef4444; }
        """
        )
        self.remove_btn.clicked.connect(
            lambda: self.remove_clicked.emit(self.file_path)
        )
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
        self.setAcceptDrops(True)  # 支持全局拖拽

        self.file_queue = {}  # path -> role

        # 初始化拖拽覆盖层
        self.drag_overlay = DragOverlay(self)

        self._init_ui()

        # 应用深色标题栏
        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)
        apply_dark_title_bar(self, is_dark)

    def dragEnterEvent(self, event: QDragEnterEvent):
        """支持全局拖拽进入"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drag_overlay.show_overlay()

    def dragLeaveEvent(self, event):
        """拖拽离开"""
        self.drag_overlay.hide()

    def dropEvent(self, event: QDropEvent):
        """支持全局拖拽放下"""
        self.drag_overlay.hide()
        if event.mimeData().hasUrls():
            files = [url.toLocalFile() for url in event.mimeData().urls()]
            self._handle_files(files)
            event.acceptProposedAction()

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

        # 送审方式与合并选项
        row_options = QHBoxLayout()

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
        row_options.addLayout(row_mode)

        row_options.addSpacing(40)

        row_merge = QHBoxLayout()
        row_merge.addWidget(QLabel("合并选项:"))
        self.merge_group = QButtonGroup(self)
        self.radio_single = QRadioButton("单个生成")
        self.radio_merge = QRadioButton("合并生成")
        self.radio_single.setChecked(True)
        self.merge_group.addButton(self.radio_single, 0)
        self.merge_group.addButton(self.radio_merge, 1)
        row_merge.addWidget(self.radio_single)
        row_merge.addWidget(self.radio_merge)
        row_options.addLayout(row_merge)

        row_options.addStretch()
        base_layout.addLayout(row_options)

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
        file_scroll.setStyleSheet(
            "QScrollArea { border: 1px solid #374151; border-radius: 8px; }"
        )
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
        self.start_btn.setObjectName("UploadBtn")  # 复用全局蓝色按钮样式
        self.start_btn.clicked.connect(self._on_start)

        btn_layout.addStretch()
        btn_layout.addWidget(self.cancel_btn)
        btn_layout.addWidget(self.start_btn)
        main_layout.addLayout(btn_layout)

    def _handle_files(self, paths):
        # 预处理压缩包
        expanded_paths = []
        for p in paths:
            if ArchiveUtils.is_archive(p):
                extracted = ArchiveUtils.extract_archive(p)
                expanded_paths.extend(extracted)
            else:
                expanded_paths.append(p)

        for path in expanded_paths:
            if not path.endswith(".xlsx"):
                continue
            if path in self.file_queue:
                continue

            # 简单猜测角色
            fname = os.path.basename(path).lower()
            if any(k in fname for k in ["认同", "确认", "结论"]):
                role = "评估认同表"
            elif any(k in fname for k in ["报告", "评估", "拆分"]):
                role = "评估报告"
            else:
                role = "评估报告"  # 默认角色

            # [修复] 移除“平衡逻辑”，该逻辑在上传多个同名项目时会导致角色乱跳
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
                if (
                    isinstance(item_widget, ReceiptFileItem)
                    and item_widget.file_path == path
                ):
                    item_widget.deleteLater()
                    break

    def _on_start(self):
        # 1. 基础校验
        if not self.name_input.text().strip():
            QMessageBox.warning(self, "警告", "请输入项目名称")
            return

        reports = [p for p, r in self.file_queue.items() if r == "评估报告"]
        consents = [p for p, r in self.file_queue.items() if r == "评估认同表"]

        if not reports or not consents:
            QMessageBox.warning(
                self, "警告", "请确保至少各上传了一个评估报告(Excel)和评估认同表(Excel)"
            )
            return

        # 2. 自动配对 (基于文件名最长公共子串或关键词提取)
        file_groups = {}
        processed_consents = set()

        project_name = self.name_input.text().strip()

        # 【优化】定义统一的清理关键词，且移除常见日期噪音，确保配对时的对称性
        def get_core_name(fname):
            core = fname
            # 移除日期噪音 (如 20260202, 2025-01-01)
            core = re.sub(r"202[4-9][-_]?\d{2}[-_]?\d{2}", "", core)
            # 统一清理关键词 (长项在前)
            for k in [
                "评估报告",
                "结论认同表",
                "认同表",
                "确认单",
                "核定表",
                "报告",
                "评估",
                "认同",
                "确认",
                "结论",
                "副本",
                ".xlsx",
                ".xls",
            ]:
                core = core.replace(k, "")
            return core.strip(" -_—")

        for r_path in reports:
            r_name = os.path.basename(r_path)
            r_core = get_core_name(r_name)

            # 寻找最匹配的 consent
            best_match = None
            best_score = -1

            for c_path in consents:
                if c_path in processed_consents:
                    continue
                c_name = os.path.basename(c_path)
                c_core = get_core_name(c_name)

                # 比较核心名称的相似度（简单包含或相等）
                # 增加了非空校验，避免空字符串匹配导致乱对
                if (
                    r_core
                    and c_core
                    and (r_core == c_core or r_core in c_core or c_core in r_core)
                ) or (not r_core and not c_core):
                    score = len(os.path.commonprefix([r_core, c_core]))
                    if score > best_score:
                        best_score = score
                        best_match = c_path

            if best_match:
                # 【修复】使用文件全路径作为 key，避免多子项目重名导致冲突
                group_key = r_path
                file_groups[group_key] = {
                    "name": r_core if r_core else os.path.basename(r_path),
                    "eval_report": r_path,
                    "eval_consent": best_match,
                }
                processed_consents.add(best_match)

        if not file_groups:
            QMessageBox.warning(
                self, "警告", "无法自动配对评估报告和认同表，请检查文件名是否对应"
            )
            return

        is_merge = self.radio_merge.isChecked()

        # 3. 构造任务包
        # 如果是合并模式，发送一个包含所有组的任务
        if is_merge:
            data = {
                "project_name": project_name,
                "project_id": self.id_input.text().strip() or "N/A",
                "submission_unit": self.unit_input.text().strip() or "N/A",
                "submitter": self.user_input.text().strip() or "N/A",
                "submission_time": self.date_input.date().toString("yyyy-MM-dd"),
                "submission_mode": "线上" if self.radio_online.isChecked() else "邮件",
                "is_merge": True,
                "file_groups": file_groups,
            }
            self.generation_requested.emit(data)
        else:
            # 如果是非合并模式，发送多个任务（每个子项目一个卡片）
            # 如果是非合并模式，打包发送所有任务
            batch_tasks = []
            for group_name, files in file_groups.items():
                # 使用文件名作为基础名称
                simple_name = (
                    os.path.basename(files["eval_report"])
                    .replace(".xlsx", "")
                    .replace(".xls", "")
                )

                # 【优化：清理文件名中的冗余信息，优先使用用户输入的项目名称】
                final_task_name = simple_name
                if project_name:
                    # 如果文件名中包含用户输入的项目名，提取剩余的差异部分（子项目名）
                    if project_name in simple_name:
                        sub_part = simple_name.replace(project_name, "").strip(" -_—")
                        # 进一步清理日期、关键词等噪音
                        for kw in [
                            r"评估报告",
                            r"结论认同表",
                            r"认同表",
                            r"核定表",
                            r"确认单",
                            r"202[4-6][-_]?\d{2,4}",
                            r"[-_—]?\d{3,5}$",
                            r"副本",
                            r"V\d+",
                        ]:
                            sub_part = re.sub(
                                kw, "", sub_part, flags=re.IGNORECASE
                            ).strip(" -_—")

                        if sub_part:
                            final_task_name = f"{project_name}-{sub_part}"
                        else:
                            final_task_name = project_name
                    else:
                        # 如果不包含，则保留原名，但尝试清理噪音
                        cleaned_simple = simple_name
                        for kw in [
                            r"评估报告",
                            r"结论认同表",
                            r"认同表",
                            r"核定表",
                            r"确认单",
                            r"副本",
                            r"[-_—]?\d{3,5}$",
                        ]:
                            cleaned_simple = re.sub(
                                kw, "", cleaned_simple, flags=re.IGNORECASE
                            ).strip(" -_—")
                        final_task_name = cleaned_simple

                sub_data = {
                    "project_name": final_task_name,
                    "project_id": self.id_input.text().strip() or "N/A",
                    "submission_unit": self.unit_input.text().strip() or "N/A",
                    "submitter": self.user_input.text().strip() or "N/A",
                    "submission_time": self.date_input.date().toString("yyyy-MM-dd"),
                    "submission_mode": (
                        "线上" if self.radio_online.isChecked() else "邮件"
                    ),
                    "is_merge": False,
                    "eval_report_path": files["eval_report"],
                    "eval_consent_path": files["eval_consent"],
                }
                batch_tasks.append(sub_data)

            # 发送批量任务包
            self.generation_requested.emit(
                {"is_batch": True, "tasks": batch_tasks, "project_name": project_name}
            )

        self.accept()
