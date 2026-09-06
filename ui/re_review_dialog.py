import os
import re
import pandas as pd
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
    QGroupBox,
    QScrollArea,
)
from PySide6.QtCore import Qt, Signal, QByteArray, QSize
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QImage, QPixmap
from utils.path_utils import get_resource_path
from .upload_dialog import UploadAreaWidget, ElidedLabel, DragOverlay
from extend.matcher_config import MatcherConfig
from utils.archive_utils import ArchiveUtils


class ReReviewFileItem(QFrame):
    """重评文件项 - 模仿图1样式"""

    remove_clicked = Signal(str)

    def __init__(self, file_path, is_eval_report, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.filename = os.path.basename(file_path)
        self.setProperty("class", "FileQueueItem")

        self.setStyleSheet(
            """
            QFrame[class="FileQueueItem"] {
                border-radius: 8px;
                padding: 10px;
                margin: 2px 0;
            }
        """
        )

        layout = QHBoxLayout(self)

        # 文件图标
        icon_label = QLabel("X")
        icon_label.setFixedSize(28, 28)
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet(
            f"""
            background: {'#217346' if is_eval_report else '#2b579a'};
            color: white; border-radius: 4px;
            font-size: 11px; font-weight: bold;
        """
        )
        layout.addWidget(icon_label)

        # 文件名
        name_label = ElidedLabel(self.filename)
        name_label.setStyleSheet("font-weight: 600; font-size: 12px;")
        layout.addWidget(name_label, stretch=1)

        # 状态
        status_label = QLabel("评估报告" if is_eval_report else "待重评文件")
        status_label.setProperty("class", "task-meta")
        status_label.setFixedWidth(80)  # 固定宽度，防止压缩
        status_label.setAlignment(Qt.AlignCenter)
        status_label.setStyleSheet(
            f"""
            font-size: 11px; font-weight: 600;
            color: {'#10b981' if is_eval_report else '#3b82f6'};
            background: {'rgba(16, 185, 129, 0.1)' if is_eval_report else 'rgba(59, 130, 246, 0.1)'};
            border-radius: 4px; padding: 2px 4px;
        """
        )
        layout.addWidget(status_label)

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
        # 不再硬编码按钮颜色，让全局样式 class="FileQueueItem" 中的 QPushButton 处理
        self.remove_btn.clicked.connect(
            lambda: self.remove_clicked.emit(self.file_path)
        )
        layout.addWidget(self.remove_btn)


class ReReviewUploadDialog(QDialog):
    """重评文件上传对话框 - 复用新建任务样式"""

    task_started = Signal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("新建重评任务")
        self.setMinimumSize(700, 600)
        self.setWindowIcon(QIcon(get_resource_path("ui/logo.ico")))
        self.setAcceptDrops(True)  # 支持全局拖拽

        # 应用原生标题栏深色模式
        from PySide6.QtWidgets import QApplication
        from utils.styles import apply_dark_title_bar

        is_dark = "background-color: #1f2937" in (
            QApplication.instance().styleSheet() or ""
        )
        apply_dark_title_bar(self, is_dark)

        self.excel1_path = ""
        self.excel2_path = ""

        # 初始化拖拽覆盖层
        self.drag_overlay = DragOverlay(self)

        self.setup_ui()

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
            self.handle_files(files)
            event.acceptProposedAction()

    def setup_ui(self):
        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("border: none; background: transparent;")

        scroll_content = QWidget()
        scroll_content.setObjectName("ScrollContent")
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(25)

        title = QLabel("vlookup 重评自动标注项目")
        title.setStyleSheet("font-size: 24px; font-weight: bold; margin-bottom: 5px;")
        layout.addWidget(title)

        # ========== 1. 文件选择区域 (模拟图4样式的深色蓝框) ==========
        file_section = QGroupBox("📂 文件选择 (支持拖拽 Excel)")
        file_section.setStyleSheet(
            """
            QGroupBox {
                border: 2px solid #38bdf8;
                border-radius: 12px;
                margin-top: 15px;
                padding-top: 15px;
                font-weight: bold;
                color: #38bdf8;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 15px;
                padding: 0 5px;
            }
        """
        )
        file_layout = QVBoxLayout(file_section)
        file_layout.setContentsMargins(20, 20, 20, 20)
        file_layout.setSpacing(15)

        # 上传区域 (复用 UploadAreaWidget)
        self.upload_area = UploadAreaWidget()
        self.upload_area.files_dropped.connect(self.handle_files)
        file_layout.addWidget(self.upload_area)

        # ========== 已上传文件列表 (匹配图3样式) ==========
        queue_label = QLabel("📋 已上传文件:")
        queue_label.setStyleSheet(
            "font-size: 13px; font-weight: bold; margin-top: 5px;"
        )
        file_layout.addWidget(queue_label)

        self.files_scroll = QScrollArea()
        self.files_scroll.setWidgetResizable(True)
        self.files_scroll.setFixedHeight(180)
        self.files_scroll.setStyleSheet(
            """
            QScrollArea {
                border: 1px solid #4b5563;
                border-radius: 8px;
                background-color: rgba(31, 41, 55, 0.5);
            }
        """
        )
        self.queue_content = QWidget()
        self.queue_content.setStyleSheet("background: transparent;")
        self.queue_layout = QVBoxLayout(self.queue_content)
        self.queue_layout.setSpacing(8)
        self.queue_layout.addStretch()
        self.files_scroll.setWidget(self.queue_content)
        file_layout.addWidget(self.files_scroll)

        layout.addWidget(file_section)

        # 提示信息
        hint_label = QLabel(
            "💡 提示：您可以直接将两个 Excel 文件拖入上方区域，系统将自动按顺序分配。"
        )
        hint_label.setProperty("class", "task-meta")
        layout.addWidget(hint_label)

        layout.addStretch()
        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)

        # ========== 2. 底部按钮 ==========
        footer = QFrame()
        footer.setObjectName("DialogFooter")
        footer.setFixedHeight(80)
        footer.setStyleSheet(
            """
            QFrame#DialogFooter {
                background-color: transparent;
                border-top: 1px solid #374151;
            }
        """
        )
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(30, 0, 30, 0)

        self.btn_run = QPushButton("开始处理")
        self.btn_run.setFixedHeight(45)
        self.btn_run.setFixedWidth(240)
        self.btn_run.setObjectName("UploadBtn")
        self.btn_run.clicked.connect(self.run_process)
        footer_layout.addStretch()
        footer_layout.addWidget(self.btn_run)
        footer_layout.addStretch()

        main_layout.addWidget(footer)

    def handle_files(self, files):
        """处理拖拽进出的文件"""
        # 预处理压缩包
        expanded_files = []
        for f in files:
            if ArchiveUtils.is_archive(f):
                extracted = ArchiveUtils.extract_archive(f)
                expanded_files.extend(extracted)
            else:
                expanded_files.append(f)

        excel_files = [f for f in expanded_files if f.lower().endswith(".xlsx")]
        for f in excel_files:
            filename = os.path.basename(f).replace(".xlsx", "")
            is_eval = bool(re.search(r"\d+$", filename))
            self.add_file_to_ui(f, is_eval)

    def add_file_to_ui(self, file_path, is_eval):
        """统一管理文件添加后的 UI 更新"""
        if is_eval:
            if self.excel1_path:
                return
            self.excel1_path = file_path
        else:
            if self.excel2_path:
                return
            self.excel2_path = file_path

        # 添加图形化列表项
        item = ReReviewFileItem(file_path, is_eval)
        item.remove_clicked.connect(self.remove_file)
        self.queue_layout.insertWidget(self.queue_layout.count() - 1, item)

    def remove_file(self, file_path):
        """移除文件逻辑"""
        if self.excel1_path == file_path:
            self.excel1_path = ""
        elif self.excel2_path == file_path:
            self.excel2_path = ""

        for i in range(self.queue_layout.count()):
            w = self.queue_layout.itemAt(i).widget()
            if isinstance(w, ReReviewFileItem) and w.file_path == file_path:
                w.deleteLater()
                break

    def run_process(self):
        if not self.excel1_path or not self.excel2_path:
            QMessageBox.warning(self, "错误", "请选择两个 Excel 文件")
            return

        import time

        timestamp = time.strftime("%Y%m%d_%H%M%S")

        # 从配置中获取重评文件存放位置
        config = MatcherConfig.load()
        base_output_dir = config.get("storage", {}).get("re_review")
        if not base_output_dir:
            base_output_dir = os.path.dirname(self.excel2_path)

        output_dir = os.path.join(base_output_dir, f"重评结果_{timestamp}")

        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        try:
            # 获取项目名称逻辑：去掉评估报告后的数字和文字
            e1_full_name = os.path.basename(self.excel1_path).replace(".xlsx", "")
            # 去掉末尾数字
            project_name = re.sub(r"\d+$", "", e1_full_name)
            # 去掉 "评估报告" 字样 (如果有)
            project_name = project_name.replace("评估报告", "")
            # 去掉首尾的横杠及空格
            project_name = project_name.strip("-").strip()

            # 如果去干净了变空了，回退到原始名
            if not project_name:
                project_name = e1_full_name

            task_info = {
                "project_name": project_name,
                "time": time.strftime("%H:%M:%S"),
                "excel1": self.excel1_path,
                "excel2": self.excel2_path,
                "output_dir": output_dir,
            }

            self.task_started.emit(task_info)
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
