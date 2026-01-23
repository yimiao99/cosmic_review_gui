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
    QMessageBox,
    QWidget,
    QComboBox,
    QRadioButton,
    QButtonGroup,
    QFileDialog,
    QGroupBox,
    QCheckBox,
    QSlider,
    QScrollArea,
    QGridLayout,
)
from PySide6.QtCore import Qt, Signal, QByteArray, QSize
from PySide6.QtGui import (
    QDragEnterEvent,
    QDropEvent,
    QIcon,
    QFontMetrics,
    QPainter,
    QImage,
    QPixmap,
)

from ui.task_card import TaskCard
from utils.document_processor import DocumentProcessor
from utils.similarity_checker import SimilarityChecker
from utils.styles import apply_dark_title_bar
from utils.path_utils import get_resource_path


class DownOnlyComboBox(QComboBox):
    """强制向下展开并屏蔽滑轮滚动的下拉框"""

    def wheelEvent(self, event):
        """屏蔽鼠标滑轮滚动切换选项，防止误触"""
        event.ignore()

    def showPopup(self):
        """重写弹出方法，强制向下展开并设置最大高度"""
        super().showPopup()
        # 获取下拉列表窗口
        popup = self.view().window()
        if popup:
            # 计算位置：在控件正下方
            pos = self.mapToGlobal(self.rect().bottomLeft())
            popup.move(pos)

            # 强制设置下拉列表的最大高度（例如，5个选项的高度）
            # 假设每个选项高度为30px，那么5个就是150px
            max_height = 150  # 可以根据需要调整
            popup.setFixedHeight(max_height)


class ElidedLabel(QLabel):
    """自动省略超长文本的标签"""

    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self._full_text = text

    def setText(self, text):
        self._full_text = text
        self._update_elided_text()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_elided_text()

    def _update_elided_text(self):
        metrics = QFontMetrics(self.font())
        # 使用当前宽度减去一些余量，防止计算精度问题导致溢出
        width = max(0, self.width() - 5)
        elided = metrics.elidedText(self._full_text, Qt.ElideRight, width)
        super().setText(elided)

    def sizeHint(self):
        # 提供一个较小的 sizeHint，允许布局将其压缩
        hint = super().sizeHint()
        hint.setWidth(100)
        return hint

    def minimumSizeHint(self):
        hint = super().minimumSizeHint()
        hint.setWidth(50)
        return hint


class FileQueueItem(QFrame):
    """文件队列项"""

    remove_clicked = Signal(str)

    def __init__(self, display_name, cleaned_name, has_word, has_excel, parent=None):
        super().__init__(parent)
        self.setProperty("class", "FileQueueItem")
        self.display_name = display_name
        self.cleaned_name = cleaned_name

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)

        # 左侧：文件信息
        left_widget = QWidget()
        left_widget.setStyleSheet("background: transparent; border: none;")
        from PySide6.QtWidgets import QSizePolicy

        left_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(6)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # 文件名
        name_label = ElidedLabel(display_name)
        name_label.setStyleSheet(
            "font-weight: 600; font-size: 12px; border: none; background: transparent;"
        )
        name_label.setToolTip(display_name)  # 长文件名悬浮显示
        left_layout.addWidget(name_label)

        # 文件图标行
        icon_row = QHBoxLayout()
        icon_row.setSpacing(8)

        # Word 图标
        word_icon = QLabel("W")
        word_icon.setFixedSize(28, 28)
        word_icon.setAlignment(Qt.AlignCenter)
        word_icon.setStyleSheet(
            """
            background: #2b579a; color: white; border-radius: 4px;
            font-size: 11px; font-weight: bold;
        """
            if has_word
            else """
            background: transparent; color: palette(mid); border: 2px dashed palette(mid);
            border-radius: 4px; font-size: 11px; font-weight: bold;
        """
        )
        icon_row.addWidget(word_icon)

        # Excel 图标
        excel_icon = QLabel("X")
        excel_icon.setFixedSize(28, 28)
        excel_icon.setAlignment(Qt.AlignCenter)
        excel_icon.setStyleSheet(
            """
            background: #217346; color: white; border-radius: 4px;
            font-size: 11px; font-weight: bold;
        """
            if has_excel
            else """
            background: transparent; color: palette(mid); border: 2px dashed palette(mid);
            border-radius: 4px; font-size: 11px; font-weight: bold;
        """
        )
        icon_row.addWidget(excel_icon)

        # 状态标签
        status_label = QLabel("配对成功" if (has_word and has_excel) else "无法开始")
        status_label.setStyleSheet(
            f"""
            font-size: 11px;
            color: {'#10b981' if (has_word and has_excel) else '#ef4444'};
            font-weight: 600;
        """
        )
        icon_row.addWidget(status_label)
        icon_row.addStretch()

        left_layout.addLayout(icon_row)
        layout.addWidget(left_widget, stretch=1)

        # 右侧：删除按钮
        self.remove_btn = QPushButton()
        self.remove_btn.setFixedSize(32, 32)
        self.remove_btn.setCursor(Qt.PointingHandCursor)
        self.remove_btn.setProperty("class", "DeleteBtn")
        self.remove_btn.setToolTip("移除此文件")

        # 使用自定义的 SVG 回收站图标
        trash_icon_svg = """
        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#ef4444" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            <line x1="10" y1="11" x2="10" y2="17"></line>
            <line x1="14" y1="11" x2="14" y2="17"></line>
        </svg>
        """
        trash_qicon = QIcon(
            QPixmap.fromImage(QImage.fromData(QByteArray(trash_icon_svg.encode())))
        )
        self.remove_btn.setIcon(trash_qicon)
        self.remove_btn.setIconSize(QSize(18, 18))

        self.remove_btn.setStyleSheet(
            """
            QPushButton {
                background: #fee2e2;
                border: 1px solid #fca5a5;
                border-radius: 16px;
                padding: 0px;
                color: transparent;
            }
            QPushButton:hover {
                background: #ef4444;
                border-color: #dc2626;
            }
            QPushButton:hover QIcon {
                /* 注意：QSS 无法直接改变 SVG 颜色，所以我们通过设置按钮背景色来强化视觉 */
            }
        """
        )
        # 为深色模式特殊控制图标颜色（如果需要）
        # 我们可以通过在 hover 时切换图标来实现更好的效果，但目前通过背景色切换已经很明显了。
        # 使用闭包捕获当时的 cleaned_name
        current_name = self.cleaned_name
        self.remove_btn.clicked.connect(lambda: self.remove_clicked.emit(current_name))
        layout.addWidget(self.remove_btn, 0, Qt.AlignRight | Qt.AlignVCenter)


class UploadAreaWidget(QFrame):
    """上传区域组件"""

    files_dropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setFixedHeight(150)
        self.setObjectName("UploadArea")

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(10)

        # 图标
        icon = QLabel("📂")
        icon.setStyleSheet("font-size: 48px; background: transparent;")
        icon.setAlignment(Qt.AlignCenter)
        layout.addWidget(icon)

        # 主文本
        main_text = QLabel("点击或拖拽文件到此处")
        main_text.setStyleSheet("font-weight: bold; font-size: 12px;")
        main_text.setAlignment(Qt.AlignCenter)
        layout.addWidget(main_text)

        # 提示文本
        hint_text = QLabel("支持 .docx 和 .xlsx 自动配对")
        hint_text.setProperty("class", "task-meta")
        hint_text.setStyleSheet("font-size: 11px;")
        hint_text.setAlignment(Qt.AlignCenter)
        layout.addWidget(hint_text)

    def mousePressEvent(self, event):
        """点击打开文件选择对话框"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择文件",
            "",
            # 修改这里：添加 .doc 格式支持
            "文档文件 (*.doc *.docx *.xlsx);;Word文档 (*.doc *.docx);;Excel表格 (*.xlsx)",
        )
        if files:
            self.files_dropped.emit(files)

    def dragEnterEvent(self, event: QDragEnterEvent):
        """拖拽进入"""
        if event.mimeData().hasUrls():
            event.accept()
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        """拖拽放下"""
        if event.mimeData().hasUrls():
            event.accept()
            event.acceptProposedAction()
            files = [url.toLocalFile() for url in event.mimeData().urls()]
            self.files_dropped.emit(files)


class UploadDialog(QDialog):
    task_submitted = Signal(dict)  # ✅ 新增信号
    """上传任务弹窗 - 新UI外观 + 完整功能"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("新建审核任务")
        self.setWindowIcon(QIcon(get_resource_path("ui/logo.png")))
        self.setMinimumSize(900, 800)

        # 应用原生标题栏深色模式
        from PySide6.QtWidgets import QApplication

        is_dark = "background-color: #1f2937" in (
            QApplication.instance().styleSheet() or ""
        )
        apply_dark_title_bar(self, is_dark)

        self.file_queue = {}
        self.excel_info = {}  # 存储Excel文件信息

        # 主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        # 移除本地背景设置，遵循全局主题
        scroll_area.setFrameShape(QFrame.NoFrame)

        scroll_content = QWidget()
        scroll_content.setObjectName("ScrollContent")
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        # ========== 1. 人天输入 + 并发数 ==========
        config_section = self._create_config_section()
        layout.addWidget(config_section)

        # ========== 2. 文件上传区域 ==========
        file_section = self._create_file_upload_section()
        layout.addWidget(file_section)

        # ========== 3. 节点选择 ==========
        node_section = self._create_node_selection_section()
        layout.addWidget(node_section)

        # ========== 4. 简单模式设置 ==========
        self.simple_section = self._create_simple_section()
        layout.addWidget(self.simple_section)
        # 根据 checkbox 状态初始化，默认为 True 所以这里应该显示
        self.simple_section.setVisible(self.simple_checkbox.isChecked())

        # ========== 5. 层级模式设置 ==========
        self.hierarchy_section = self._create_hierarchy_section()
        layout.addWidget(self.hierarchy_section)

        # ========== 6. 匹配选项 ==========
        options_section = self._create_options_section()
        layout.addWidget(options_section)

        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area)

        # ========== 7. 底部按钮 ==========
        footer = self._create_footer()
        main_layout.addWidget(footer)

    def _create_config_section(self):
        """创建配置区域（人天 + 并发数）"""
        section = QGroupBox("⚙️ 基础配置")
        layout = QVBoxLayout(section)
        layout.setSpacing(15)
        layout.setContentsMargins(25, 10, 25, 20)

        # 人天输入
        row = QHBoxLayout()
        row.setAlignment(Qt.AlignVCenter)
        label = QLabel("线上送审人天:")
        label.setFixedWidth(110)
        label.setStyleSheet("font-size: 13px; font-weight: 600;")

        self.days_input = QLineEdit()
        self.days_input.setPlaceholderText("请输入数值，例如: 125.5")
        self.days_input.setFixedHeight(38)

        row.addWidget(label)
        row.addWidget(self.days_input, stretch=1)
        layout.addLayout(row)

        return section

    def _create_file_upload_section(self):
        """创建文件上传区域"""
        section = QGroupBox("📂 文件选择 (支持拖拽)")
        # 移除硬编码样式

        layout = QVBoxLayout(section)
        layout.setSpacing(15)
        layout.setContentsMargins(25, 12, 25, 20)

        # 上传区域
        self.upload_area = UploadAreaWidget()
        self.upload_area.files_dropped.connect(self.handle_files)
        layout.addWidget(self.upload_area)

        # 文件队列
        queue_row = QHBoxLayout()
        queue_label = QLabel("📋 已上传文件:")
        queue_label.setFixedWidth(110)
        queue_label.setStyleSheet("font-size: 13px; font-weight: 600;")
        queue_row.addWidget(queue_label)
        queue_row.addStretch(1)
        layout.addLayout(queue_row)

        # 滚动区域
        files_scroll = QScrollArea()
        files_scroll.setWidgetResizable(True)
        files_scroll.setFixedHeight(180)
        files_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        files_scroll.setStyleSheet(
            """
            QScrollArea {
                border: 1px solid palette(mid);
                border-radius: 8px;
                background-color: transparent;
            }
        """
        )

        files_content = QWidget()
        files_content.setStyleSheet("background: transparent;")
        self.queue_container = QVBoxLayout(files_content)
        self.queue_container.setSpacing(8)
        self.queue_container.setContentsMargins(10, 10, 10, 10)
        files_scroll.setWidget(files_content)

        layout.addWidget(files_scroll)

        return section

    def _create_node_selection_section(self):
        """创建节点选择区域(取代原匹配模式选择)"""
        section = QGroupBox("🛡 审核节点选择 (可多选)")
        layout = QVBoxLayout(section)
        layout.setSpacing(15)
        layout.setContentsMargins(25, 12, 25, 20)

        # 节点 1-4 (基础校验)
        base_grid = QHBoxLayout()
        base_grid.setAlignment(Qt.AlignVCenter)
        base_grid.setContentsMargins(0, 5, 0, 5)

        self.check_template = QCheckBox("1. 模板校验")
        self.check_template.setChecked(True)

        self.check_empty = QCheckBox("2. 空值检查")
        self.check_empty.setChecked(True)

        self.check_ratio = QCheckBox("3. 送审比例")
        self.check_ratio.setChecked(True)

        self.check_factors = QCheckBox("4. 附加值因子")
        self.check_factors.setChecked(True)

        base_grid.addWidget(self.check_template)
        base_grid.addSpacing(25)
        base_grid.addWidget(self.check_empty)
        base_grid.addSpacing(25)
        base_grid.addWidget(self.check_ratio)
        base_grid.addSpacing(25)
        base_grid.addWidget(self.check_factors)
        base_grid.addStretch(1)
        layout.addLayout(base_grid)

        # 间隔线或间距
        layout.addSpacing(10)

        # 节点 5-6 (匹配逻辑) - 带详细说明
        match_content = QVBoxLayout()
        match_content.setSpacing(15)

        # 5. 层级模式
        h_container = QWidget()
        h_container.setStyleSheet("background: transparent;")
        h_layout = QVBoxLayout(h_container)
        h_layout.setContentsMargins(0, 0, 0, 0)
        h_layout.setSpacing(4)

        self.hierarchy_checkbox = QCheckBox("5. 层级匹配（三级模块匹配）")
        self.hierarchy_checkbox.setChecked(True)
        h_desc = QLabel("适用于：Excel 三级模块列 与 Word 4.1.1.x 分级对应")
        h_desc.setProperty("class", "task-meta")
        h_desc.setStyleSheet("margin-left: 28px; font-size: 11px;")

        h_layout.addWidget(self.hierarchy_checkbox)
        h_layout.addWidget(h_desc)
        match_content.addWidget(h_container)

        # 6. 简单模式
        s_container = QWidget()
        s_container.setStyleSheet("background: transparent;")
        s_layout = QVBoxLayout(s_container)
        s_layout.setContentsMargins(0, 0, 0, 0)
        s_layout.setSpacing(4)

        self.simple_checkbox = QCheckBox("6. 功能过程（单列单行匹配）")
        self.simple_checkbox.setChecked(True)
        s_desc = QLabel("适用于：Excel 单列直接展示所有功能过程")
        s_desc.setProperty("class", "task-meta")
        s_desc.setStyleSheet("margin-left: 28px; font-size: 11px;")

        s_layout.addWidget(self.simple_checkbox)
        s_layout.addWidget(s_desc)
        match_content.addWidget(s_container)

        # 7. 数据移动类型
        dm_container = QWidget()
        dm_container.setStyleSheet("background: transparent;")
        dm_layout = QVBoxLayout(dm_container)
        dm_layout.setContentsMargins(0, 0, 0, 0)
        dm_layout.setSpacing(4)

        self.dm_checkbox = QCheckBox("7. 功能过程数据移动类型（E开头W/X结束）")
        self.dm_checkbox.setChecked(True)
        dm_desc = QLabel("核对功能过程子项是否符合以“E”开头，“W”或“X”结束的规则")
        dm_desc.setProperty("class", "task-meta")
        dm_desc.setStyleSheet("margin-left: 28px; font-size: 11px;")

        dm_layout.addWidget(self.dm_checkbox)
        dm_layout.addWidget(dm_desc)
        match_content.addWidget(dm_container)

        layout.addLayout(match_content)

        # 绑定信号，联动控制下方设置区域显隐
        self.simple_checkbox.stateChanged.connect(self.on_mode_checkbox_changed)
        self.hierarchy_checkbox.stateChanged.connect(self.on_mode_checkbox_changed)

        return section

    def on_mode_checkbox_changed(self, state):
        """复选框状态改变，联动控制设置区域显示"""
        self.simple_section.setVisible(self.simple_checkbox.isChecked())
        self.hierarchy_section.setVisible(self.hierarchy_checkbox.isChecked())

    def _create_simple_section(self):
        """创建简单模式设置"""
        section = QGroupBox("📋 功能过程设置")
        layout = QVBoxLayout(section)
        layout.setSpacing(12)
        layout.setContentsMargins(25, 12, 25, 20)

        # Excel 工作表
        excel_row = QHBoxLayout()
        excel_row.setAlignment(Qt.AlignVCenter)
        excel_label = QLabel("Excel 工作表:")
        excel_label.setFixedWidth(110)
        excel_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        excel_label.setStyleSheet("font-size: 12px; font-weight: 600;")

        self.simple_excel_combo = DownOnlyComboBox()
        self.simple_excel_combo.setFixedHeight(36)
        self.simple_excel_combo.setMaxVisibleItems(5)
        self.simple_excel_combo.currentIndexChanged.connect(
            self.on_simple_sheet_changed
        )

        excel_row.addWidget(excel_label)
        excel_row.addWidget(self.simple_excel_combo, stretch=1)
        layout.addLayout(excel_row)

        # 功能点列
        func_row = QHBoxLayout()
        func_label = QLabel("功能点列:")
        func_label.setFixedWidth(110)
        func_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        func_label.setStyleSheet("font-size: 12px; font-weight: 600;")

        self.func_combo = DownOnlyComboBox()
        self.func_combo.setFixedHeight(36)
        self.func_combo.setMaxVisibleItems(5)

        func_row.addWidget(func_label)
        func_row.addWidget(self.func_combo, stretch=1)
        layout.addLayout(func_row)

        # 提示
        hint_row = QHBoxLayout()
        hint = QLabel("💡 选择包含功能点的列")
        hint.setProperty("class", "task-meta")
        hint.setStyleSheet("font-size: 11px;")

        hint_row.addWidget(hint, stretch=1)
        layout.addLayout(hint_row)

        return section

    def _create_hierarchy_section(self):
        """创建层级模式设置"""
        section = QGroupBox("📊 层级模式设置")
        layout = QVBoxLayout(section)
        layout.setSpacing(12)
        layout.setContentsMargins(25, 12, 25, 20)

        # Excel 工作表
        excel_row = QHBoxLayout()
        excel_row.setAlignment(Qt.AlignVCenter)
        excel_label = QLabel("Excel 工作表:")
        excel_label.setFixedWidth(110)
        excel_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        excel_label.setStyleSheet("font-size: 12px; font-weight: 600;")

        self.excel_sheet_combo = DownOnlyComboBox()
        self.excel_sheet_combo.setFixedHeight(36)
        self.excel_sheet_combo.setMaxVisibleItems(5)
        self.excel_sheet_combo.currentIndexChanged.connect(self.on_sheet_changed)

        excel_row.addWidget(excel_label)
        excel_row.addWidget(self.excel_sheet_combo, stretch=1)
        layout.addLayout(excel_row)

        # 一级模块列
        level1_row = QHBoxLayout()
        level1_label = QLabel("一级模块列:")
        level1_label.setFixedWidth(110)
        level1_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        level1_label.setStyleSheet("font-size: 12px; font-weight: 600;")

        self.level1_combo = DownOnlyComboBox()
        self.level1_combo.setFixedHeight(36)
        self.level1_combo.setMaxVisibleItems(5)

        level1_row.addWidget(level1_label)
        level1_row.addWidget(self.level1_combo, stretch=1)
        layout.addLayout(level1_row)

        # 二级模块列
        level2_row = QHBoxLayout()
        level2_label = QLabel("二级模块列:")
        level2_label.setFixedWidth(110)
        level2_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        level2_label.setStyleSheet("font-size: 12px; font-weight: 600;")

        self.level2_combo = DownOnlyComboBox()
        self.level2_combo.setFixedHeight(36)
        self.level2_combo.setMaxVisibleItems(5)

        level2_row.addWidget(level2_label)
        level2_row.addWidget(self.level2_combo, stretch=1)
        layout.addLayout(level2_row)

        # 三级模块列
        level3_row = QHBoxLayout()
        level3_label = QLabel("三级模块列:")
        level3_label.setFixedWidth(110)
        level3_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        level3_label.setStyleSheet("font-size: 12px; font-weight: 600;")

        self.level3_combo = DownOnlyComboBox()
        self.level3_combo.setFixedHeight(36)
        self.level3_combo.setMaxVisibleItems(5)

        level3_row.addWidget(level3_label)
        level3_row.addWidget(self.level3_combo, stretch=1)
        layout.addLayout(level3_row)

        # 提示
        hint_row = QHBoxLayout()
        hint = QLabel("💡 请选择包含一二三级模块的列")
        hint.setProperty("class", "task-meta")
        hint.setStyleSheet("font-size: 11px;")

        hint_row.addWidget(hint, stretch=1)
        layout.addLayout(hint_row)

        return section

    def on_sheet_changed(self, index):
        """层级模式：工作表变更时更新列下拉框和列信息显示"""
        if index < 0:
            return

        # 获取选中的工作表名称
        sheet_name = self.excel_sheet_combo.itemData(index)
        if not sheet_name:
            sheet_name = self.excel_sheet_combo.itemText(index)
            # 从显示文本中提取工作表名称
            if "、" in sheet_name:
                sheet_name = sheet_name.split("、", 1)[1]

        # 获取第一个有效的Excel文件信息
        excel_key = None
        for key, info in self.file_queue.items():
            if info["has_excel"] and key in self.excel_info:
                excel_key = key
                break

        if not excel_key:
            return

        # 获取该工作表的列信息
        excel_data = self.excel_info.get(excel_key)
        if not excel_data:
            return

        columns_info = excel_data.get("columns_info", {}).get(sheet_name, [])

        print(f"\n工作表 '{sheet_name}' 变更，更新列下拉框")
        print(f"列信息: {[(col['letter'], col['name']) for col in columns_info]}")

        # 更新所有列下拉框
        self.update_column_combo(self.level1_combo, columns_info, "一级模块列")
        self.update_column_combo(self.level2_combo, columns_info, "二级模块列")
        self.update_column_combo(self.level3_combo, columns_info, "三级模块列")

    def _create_options_section(self):
        """创建匹配选项"""
        section = QGroupBox("⚙️ 匹配选项")
        layout = QVBoxLayout(section)
        layout.setSpacing(12)
        layout.setContentsMargins(25, 12, 25, 20)

        # 启用模糊匹配
        fuzzy_row = QHBoxLayout()
        fuzzy_row.setAlignment(Qt.AlignVCenter)
        fuzzy_label = QLabel("启用模糊匹配:")
        fuzzy_label.setFixedWidth(110)
        fuzzy_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        fuzzy_label.setStyleSheet("font-size: 12px; font-weight: 600;")

        self.fuzzy_check = QCheckBox()
        self.fuzzy_check.setChecked(True)

        fuzzy_row.addWidget(fuzzy_label)
        fuzzy_row.addWidget(self.fuzzy_check, stretch=1)
        layout.addLayout(fuzzy_row)

        # 相似度阈值
        threshold_row = QHBoxLayout()
        threshold_row.setAlignment(Qt.AlignVCenter)
        threshold_label = QLabel("相似度阈值:")
        threshold_label.setFixedWidth(110)
        threshold_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        threshold_label.setStyleSheet("font-size: 12px; font-weight: 600;")

        self.threshold_slider = QSlider(Qt.Horizontal)
        self.threshold_slider.setMinimum(50)
        self.threshold_slider.setMaximum(100)
        self.threshold_slider.setValue(80)
        self.threshold_slider.setFixedHeight(28)

        self.threshold_value = QLabel("0.80")
        self.threshold_value.setStyleSheet(
            "font-size: 12px; font-weight: bold; min-width: 50px;"
        )
        self.threshold_slider.valueChanged.connect(
            lambda v: self.threshold_value.setText(f"{v / 100:.2f}")
        )

        threshold_row.addWidget(threshold_label)
        threshold_row.addWidget(self.threshold_slider, stretch=1)
        threshold_row.addWidget(self.threshold_value)
        layout.addLayout(threshold_row)

        return section

    def _create_footer(self):
        """创建底部按钮"""
        footer = QFrame()
        # 移除硬编码样式
        footer.setStyleSheet("border-top: 1px solid palette(mid); padding: 12px 20px;")

        layout = QHBoxLayout(footer)
        layout.setContentsMargins(0, 0, 0, 0)

        layout.addStretch()

        # 清除按钮
        clear_btn = QPushButton("🗑️ 清除")
        clear_btn.setFixedSize(90, 40)
        clear_btn.setStyleSheet(
            """
            QPushButton {
                color: #ef4444;
                border: 2px solid palette(mid);
                border-radius: 8px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background: palette(alternate-base);
                border-color: #ef4444;
            }
        """
        )
        clear_btn.clicked.connect(self.clear_all)
        layout.addWidget(clear_btn)

        # 开始审核按钮
        start_btn = QPushButton("开始审核")
        start_btn.setFixedSize(135, 40)
        start_btn.setObjectName("UploadBtn")
        start_btn.clicked.connect(self.start_review)
        layout.addWidget(start_btn)

        return footer

    def handle_files(self, files):
        """处理上传的文件"""
        try:
            print("=" * 50)
            print("开始处理文件:")

            for file_path in files:
                filename = os.path.basename(file_path)
                base_name = os.path.splitext(filename)[0]
                ext = os.path.splitext(filename)[1].lower()

                print(f"\n原始文件名: {filename}")
                print(f"基础名称: {base_name}")
                print(f"扩展名: {ext}")

                # 清理文件名：去掉前后缀
                cleaned_name = self.clean_filename(base_name)
                print(f"清理后名称: '{cleaned_name}'")

                # 强化匹配策略：如果找不到完全一致的 Key，尝试搜寻是否有“高度相似”的 Key（连续 6 个字符相同）
                target_key = cleaned_name
                if cleaned_name not in self.file_queue:
                    for existing_key in self.file_queue.keys():
                        if self._is_fuzzy_match(cleaned_name, existing_key):
                            target_key = existing_key
                            print(
                                f"检测到模糊匹配: '{cleaned_name}' 与现有项目 '{existing_key}' 自动合并"
                            )
                            break

                if target_key not in self.file_queue:
                    print(f"新建条目: {target_key}")
                    self.file_queue[target_key] = {
                        "has_word": False,
                        "has_excel": False,
                        "original_names": {"word": None, "excel": None},
                        "file_paths": {"word": None, "excel": None},
                    }
                else:
                    print(f"匹配到现有条目: {target_key}")

                # 修改这里：同时支持 .doc 和 .docx
                if ext == ".doc" or ext == ".docx":
                    print("-> 标记为 Word 文件")
                    self.file_queue[target_key]["has_word"] = True
                    self.file_queue[target_key]["original_names"]["word"] = filename
                    self.file_queue[target_key]["file_paths"]["word"] = file_path
                elif ext == ".xlsx":
                    print("-> 标记为 Excel 文件")
                    self.file_queue[target_key]["has_excel"] = True
                    self.file_queue[target_key]["original_names"]["excel"] = filename
                    self.file_queue[target_key]["file_paths"]["excel"] = file_path
                    # 解析Excel文件信息
                    self.parse_excel_file(target_key, file_path)

            print("\n当前文件队列:")
            for key, value in self.file_queue.items():
                print(f"  '{key}': Word={value['has_word']}, Excel={value['has_excel']}")
            print("=" * 50)

            # 当有Excel文件被添加时，更新下拉框选项
            if any(v["has_excel"] for v in self.file_queue.values()):
                self.update_excel_combos()

            self.update_queue_display()
        except Exception as e:
            from PySide6.QtWidgets import QMessageBox
            import traceback
            error_msg = f"处理文件时发生意外错误:\n{str(e)}\n\n{traceback.format_exc()}"
            print(error_msg)
            QMessageBox.critical(self, "错误", error_msg)

    def parse_excel_file(self, key, file_path):
        """解析Excel文件，获取工作表名称和列信息"""
        try:
            print(f"\n解析Excel文件: {file_path}")

            # 读取Excel文件的所有工作表
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names

            print(f"工作表列表: {sheet_names}")

            # 存储到excel_info字典中
            self.excel_info[key] = {
                "file_path": file_path,
                "sheet_names": sheet_names,
                "columns_info": {},
            }

            # 读取每个工作表
            for sheet_name in sheet_names:
                try:
                    # 读取前10行，确保能获取第3、4行数据
                    df = pd.read_excel(
                        file_path, sheet_name=sheet_name, nrows=10, header=None
                    )

                    print(f"\n工作表 '{sheet_name}' 的原始数据:")
                    print(df.head())

                    column_info = []

                    # 读取所有列
                    for i in range(len(df.columns)):
                        # 🔥 获取第4行数据作为首选表头（行索引是3，因为从0开始）
                        header_value = None
                        row4_value = None
                        if len(df) > 3:  # 至少有4行数据
                            row4_value = df.iloc[3, i]  # 第4行数据

                        # 🔥 获取第3行数据作为备选表头（行索引是2）
                        row3_value = None
                        if len(df) > 2:  # 至少有3行数据
                            row3_value = df.iloc[2, i]  # 第3行数据

                        # 🔥 智能选择表头：优先用第4行，如果为空则用第3行
                        detected_header_row = 0
                        if (
                            row4_value is not None
                            and not pd.isna(row4_value)
                            and str(row4_value).strip() != ""
                        ):
                            header_value = row4_value
                            source = "第4行"
                            detected_header_row = 3  # 0-indexed
                        elif (
                            row3_value is not None
                            and not pd.isna(row3_value)
                            and str(row3_value).strip() != ""
                        ):
                            header_value = row3_value
                            source = "第3行"
                            detected_header_row = 2  # 0-indexed
                        else:
                            header_value = None
                            source = "无表头"

                        print(
                            f"[DEBUG] Col {i}: header_row={detected_header_row}, val='{header_value}', source='{source}'"
                        )

                        # 🔥 获取第一行数据作为样本参考
                        sample_data = None
                        if len(df) > 0:
                            sample_data = df.iloc[0, i]

                        # 🔥 判断是否为空列
                        is_empty = True
                        if header_value is not None:
                            is_empty = False
                        elif (
                            sample_data is not None
                            and not pd.isna(sample_data)
                            and str(sample_data).strip() != ""
                        ):
                            is_empty = False
                        # 检查整列是否有任何非空数据
                        elif not df.iloc[:, i].dropna().empty:
                            is_empty = False

                        column_info.append(
                            {
                                "index": i,
                                "name": header_value,  # 智能选择的表头
                                "header_row": detected_header_row,  # 0-indexed row number of the header
                                "row3_value": row3_value,  # 第3行数据
                                "row4_value": row4_value,  # 第4行数据
                                "source": source,  # 表头来源
                                "first_row_data": sample_data,  # 第一行数据
                                "letter": self.index_to_excel_column(i),
                                "sample_data": sample_data,  # 第一行作为示例
                                "is_empty": is_empty,
                            }
                        )

                    self.excel_info[key]["columns_info"][sheet_name] = column_info

                    print(f"\n工作表 '{sheet_name}' 的列信息（智能表头选择）:")
                    for col in column_info:
                        if col["name"] is not None and not pd.isna(col["name"]):
                            header_display = f"{col['source']}表头: {col['name']}"
                        else:
                            header_display = "无表头"

                        sample_display = (
                            f" (第1行示例: {col['sample_data']})"
                            if col["sample_data"] is not None
                            else ""
                        )
                        print(f"  {col['letter']}列: {header_display}{sample_display}")

                except Exception as e:
                    print(f"  读取工作表 '{sheet_name}' 时出错: {str(e)}")
                    import traceback

                    traceback.print_exc()

        except Exception as e:
            print(f"解析Excel文件失败: {str(e)}")
            import traceback

            traceback.print_exc()
            # 如果解析失败，设置默认信息
            self.excel_info[key] = {
                "file_path": file_path,
                "sheet_names": ["Sheet1"],
                "columns_info": {"Sheet1": []},
            }

    def index_to_excel_column(self, index):
        """将列索引转换为Excel列字母"""
        result = ""
        while index >= 0:
            remainder = index % 26
            result = chr(65 + remainder) + result
            index = index // 26 - 1
        return result

    def update_excel_combos(self):
        """更新所有Excel相关下拉框的选项"""
        try:
            # 获取第一个有效的Excel文件信息
            excel_key = None
            for key, info in self.file_queue.items():
                if info["has_excel"] and key in self.excel_info:
                    excel_key = key
                    break

            if not excel_key:
                return

            excel_data = self.excel_info.get(excel_key)
            if not excel_data:
                return

            # 获取工作表列表
            sheet_names = excel_data.get("sheet_names", [])

            print(f"\n更新下拉框选项，工作表: {sheet_names}")

            # 清空并重新填充工作表下拉框
            self.excel_sheet_combo.clear()
            self.simple_excel_combo.clear()

            for i, sheet_name in enumerate(sheet_names):
                display_text = f"{i + 1}、{sheet_name}"
                self.excel_sheet_combo.addItem(display_text, sheet_name)
                self.simple_excel_combo.addItem(display_text, sheet_name)

            # 智能默认选择：优先寻找包含“功能点”或“拆分”的工作表
            default_sheet_idx = 0  # 兜底选第一个
            for i, name in enumerate(sheet_names):
                if any(
                    keyword in name for keyword in ["功能点", "拆分", "清单", "审核"]
                ):
                    default_sheet_idx = i
                    print(f"找到默认工作表: {name} (索引 {i})")
                    break

            # 如果没找到关键词且工作表够多，可以维持原本尝试选第3个的逻辑(针对特定模板)
            if default_sheet_idx == 0 and len(sheet_names) >= 3:
                # 检查一下逻辑，如果确实想优先选第3个（可能有些模板前两个是封面和说明）
                default_sheet_idx = 2

            # Hierarchy sheet
            if self.excel_sheet_combo.count() > default_sheet_idx:
                self.excel_sheet_combo.setCurrentIndex(default_sheet_idx)
                self.on_sheet_changed(default_sheet_idx)
            elif self.excel_sheet_combo.count() > 0:
                self.excel_sheet_combo.setCurrentIndex(0)
                self.on_sheet_changed(0)

            # Simple sheet
            if self.simple_excel_combo.count() > default_sheet_idx:
                self.simple_excel_combo.setCurrentIndex(default_sheet_idx)
                self.on_simple_sheet_changed(default_sheet_idx)
            elif self.simple_excel_combo.count() > 0:
                self.simple_excel_combo.setCurrentIndex(0)
                self.on_simple_sheet_changed(0)

        except Exception as e:
            print(f"更新下拉框时出错: {str(e)}")

    def on_sheet_changed(self, index):
        """层级模式：工作表变更时更新列下拉框和列信息显示"""
        if index < 0:
            return

        # 获取选中的工作表名称
        sheet_name = self.excel_sheet_combo.itemData(index)
        if not sheet_name:
            sheet_name = self.excel_sheet_combo.itemText(index)
            # 从显示文本中提取工作表名称
            if "、" in sheet_name:
                sheet_name = sheet_name.split("、", 1)[1]

        # 获取第一个有效的Excel文件信息
        excel_key = None
        for key, info in self.file_queue.items():
            if info["has_excel"] and key in self.excel_info:
                excel_key = key
                break

        if not excel_key:
            return

        # 获取该工作表的列信息
        excel_data = self.excel_info.get(excel_key)
        if not excel_data:
            return

        columns_info = excel_data.get("columns_info", {}).get(sheet_name, [])

        print(f"\n工作表 '{sheet_name}' 变更，更新列下拉框")
        print(f"列信息: {[(col['letter'], col['name']) for col in columns_info]}")

        # 更新所有列下拉框
        self.update_column_combo(self.level1_combo, columns_info, "一级模块列")
        self.update_column_combo(self.level2_combo, columns_info, "二级模块列")
        self.update_column_combo(self.level3_combo, columns_info, "三级模块列")

    def on_simple_sheet_changed(self, index):
        """简单模式：工作表变更时更新列下拉框和列信息显示"""
        if index < 0:
            return

        # 获取选中的工作表名称
        sheet_name = self.simple_excel_combo.itemData(index)
        if not sheet_name:
            sheet_name = self.simple_excel_combo.itemText(index)
            # 从显示文本中提取工作表名称
            if "、" in sheet_name:
                sheet_name = sheet_name.split("、", 1)[1]

        # 获取第一个有效的Excel文件信息
        excel_key = None
        for key, info in self.file_queue.items():
            if info["has_excel"] and key in self.excel_info:
                excel_key = key
                break

        if not excel_key:
            return

        # 获取该工作表的列信息
        excel_data = self.excel_info.get(excel_key)
        if not excel_data:
            return

        columns_info = excel_data.get("columns_info", {}).get(sheet_name, [])

        print(f"\n简单模式：工作表 '{sheet_name}' 变更，更新功能点列下拉框")
        print(f"列信息: {[(col['letter'], col['name']) for col in columns_info]}")

        # 更新功能点列下拉框
        self.update_column_combo(self.func_combo, columns_info, "功能点列")

    def update_column_combo(self, combo_box, columns_info, combo_name):
        """更新列下拉框选项"""
        combo_box.clear()

        # 添加空选项
        combo_box.addItem(f"请选择{combo_name}", None)

        # 添加列选项
        for col_info in columns_info:
            if col_info["is_empty"]:
                continue

            letter = col_info["letter"]
            name = col_info["name"]
            sample = col_info["sample_data"]

            # 🔥 统一使用第4行表头名称，格式化为图片中的格式
            if name and not pd.isna(name) and str(name).strip() != "":
                # 如果是中文数字，显示为"A. 一级模块"的格式
                display_text = f"{letter}. {name}"
            else:
                display_text = f"{letter}. 空列"

            # 🔥 限制显示文本长度，防止过长
            if len(display_text) > 25:
                display_text = display_text[:22] + "..."

            combo_box.addItem(display_text, col_info)

            # 🔥 设置工具提示显示完整信息
            tooltip_parts = [f"列{letter}:"]

            # 显示第4行表头
            if not pd.isna(name) and str(name).strip() != "":
                tooltip_parts.append(f"表头(第4行): {name}")

            # 显示第一行示例数据
            if sample and not pd.isna(sample):
                sample_str = str(sample)
                tooltip_parts.append(f"示例(第1行): {sample_str}")

            if len(tooltip_parts) > 1:
                combo_box.setItemData(
                    combo_box.count() - 1, "\n".join(tooltip_parts), Qt.ToolTipRole
                )
            else:
                combo_box.setItemData(
                    combo_box.count() - 1, f"列{letter}: 无数据", Qt.ToolTipRole
                )

        print(f"{combo_name} 下拉框已更新，共 {combo_box.count() - 1} 个选项")

        # 应用默认选中项
        if combo_box.count() > 1:  # Index 0 is "Please select..."

            target_idx = -1
            target_keywords = []

            if "功能点" in combo_name:
                target_idx = 6  # Default to 7th column (index 6, fallback)
                target_keywords = ["子过程描述", "功能过程", "功能描述", "功能点"]
            elif "一级" in combo_name:
                target_idx = 1  # Default to 2nd column (index 1)
                target_keywords = ["一级模块", "一级功能", "模块一", "功能一"]
            elif "二级" in combo_name:
                target_idx = 2  # Default to 3rd column (index 2)
                target_keywords = ["二级模块", "二级功能", "模块二", "功能二"]
            elif "三级" in combo_name:
                target_idx = 3  # Default to 4th column (index 3)
                target_keywords = ["三级模块", "三级功能", "模块三", "功能三"]

            print(
                f"[DEBUG] auto-selecting for {combo_name}, keywords={target_keywords}, fallback_idx={target_idx}"
            )

            found = False

            # Strategy 1: Match by Name (Priority)
            if target_keywords:
                for i in range(1, combo_box.count()):
                    data = combo_box.itemData(i)
                    if data and isinstance(data, dict):
                        col_name = str(data.get("name", "")).strip()
                        if any(kw in col_name for kw in target_keywords):
                            print(f"[DEBUG] Name match found at item {i} ({col_name})")
                            combo_box.setCurrentIndex(i)
                            found = True
                            break

            # Strategy 2: Match by Index (Fallback)
            if not found and target_idx != -1:
                for i in range(1, combo_box.count()):
                    data = combo_box.itemData(i)
                    if data and isinstance(data, dict):
                        idx = data.get("index")
                        if idx == target_idx:
                            print(f"[DEBUG] Index match found at item {i}")
                            combo_box.setCurrentIndex(i)
                            found = True
                            break

            # Strategy 3: Default to First Available (Last Resort)
            if not found:
                print("[DEBUG] No match found, falling back to index 1")
                combo_box.setCurrentIndex(1)

    def on_simple_sheet_changed(self, index):
        """简单模式：工作表变更时更新列下拉框和列信息显示"""
        if index < 0:
            return

        # 获取选中的工作表名称
        sheet_name = self.simple_excel_combo.itemData(index)
        if not sheet_name:
            sheet_name = self.simple_excel_combo.itemText(index)
            # 从显示文本中提取工作表名称
            if "、" in sheet_name:
                sheet_name = sheet_name.split("、", 1)[1]

        # 获取第一个有效的Excel文件信息
        excel_key = None
        for key, info in self.file_queue.items():
            if info["has_excel"] and key in self.excel_info:
                excel_key = key
                break

        if not excel_key:
            return

        # 获取该工作表的列信息
        excel_data = self.excel_info.get(excel_key)
        if not excel_data:
            return

        columns_info = excel_data.get("columns_info", {}).get(sheet_name, [])

        print(f"\n简单模式：工作表 '{sheet_name}' 变更，更新功能点列下拉框")
        print(f"列信息: {[(col['letter'], col['name']) for col in columns_info]}")

        # 更新功能点列下拉框
        self.update_column_combo(self.func_combo, columns_info, "功能点列")

    def clean_filename(self, filename):
        """
        全量清理：去掉前缀附件序号、末尾冗余词（项目、说明书等），提取纯净项目名
        """
        # 1. 基础处理：取文件名，先移除扩展名 (.docx, .doc, .xlsx)
        name = os.path.basename(filename).strip()
        for ext in [".docx", ".doc", ".xlsx", ".XLSX", ".DOCX"]:
            if name.lower().endswith(ext.lower()):
                name = name[: -len(ext)].strip()
                break

        # 2. 精准去掉序号前缀（如 1. 或 附件1：）
        # 先处理附件前缀
        name = re.sub(r"^附件\s*\d+\s*[：:.\-\s]*\s*", "", name)
        # 再处理纯数字点前缀 (如 1. 或 1．)
        # 匹配 1-3位数字 + 点 + 可选空格。1-3位是为了避开 2024. 这种年份开头
        name = re.sub(r"^\d{1,3}[\.．]\s*", "", name)
        name = name.strip()

        # 3. 循环清理末尾后缀，确保切干净
        while True:
            prev_name = name

            # (a) 精确匹配末尾的冗余词
            suffixes = [
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
            for s in suffixes:
                if name.endswith(s):
                    # 只有当剥离后的长度仍然合理时才剥离
                    new_n = name[: -len(s)].strip()
                    if len(new_n) >= 2:
                        name = new_n

            # (b) 去掉版本号、日期、以及 (1) (2) 这种重复标识
            name = re.sub(r"[vV]\d+(?:\.\d+)*\s*$", "", name)
            # 只有当日期在末尾时才去掉
            name = re.sub(r"(?:20)?\d{6,8}\s*$", "", name)
            # 处理副本标识，如 (1), （1）, - 副本
            name = re.sub(r"[\(（]\d+[\)）]\s*$", "", name)
            name = re.sub(r"\s*-\s*副本\s*$", "", name)

            # (c) 去掉末尾的悬空连接词
            name = re.sub(
                r"(?:的项目|项目|的需求|的产品|产品|的|之)+$", "", name
            ).strip()

            # (d) 去掉悬空的标点
            name = name.strip(" :：-－_|'\"().（）")

            if name == prev_name:
                break

        return name

    def _is_fuzzy_match(self, name1, name2):
        """
        核心匹配算法：滑动窗口检查。
        只要两个名字包含连续 6 个字符的相同片段，即视为匹配。
        """
        n1 = name1.strip()
        n2 = name2.strip()
        if not n1 or not n2:
            return False

        # 如果一方包含另一方，直接成功
        if n1 in n2 or n2 in n1:
            return True

        # 滑动窗口查找 6 位连续匹配
        min_len = 6
        if len(n1) < min_len or len(n2) < min_len:
            return n1 == n2  # 太短了就硬匹配

        for i in range(len(n1) - min_len + 1):
            window = n1[i : i + min_len]
            if window in n2:
                print(f"[匹配成功] 发现共同片段: '{window}'")
                return True
        return False

    def update_queue_display(self):
        """更新文件队列显示"""
        # 清空现有显示
        while self.queue_container.count():
            item = self.queue_container.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # 添加文件项
        valid_count = 0
        for cleaned_name, status in self.file_queue.items():
            # 为显示构建一个合适的文件名
            display_name = cleaned_name

            # 如果有原始文件名，使用原始文件名的组合
            if status["original_names"]["word"] and status["original_names"]["excel"]:
                # 显示两个原始文件名
                word_name = os.path.splitext(status["original_names"]["word"])[0]
                excel_name = os.path.splitext(status["original_names"]["excel"])[0]
                display_name = f"{word_name} | {excel_name}"
            elif status["original_names"]["word"]:
                display_name = os.path.splitext(status["original_names"]["word"])[0]
            elif status["original_names"]["excel"]:
                display_name = os.path.splitext(status["original_names"]["excel"])[0]

            item = FileQueueItem(
                display_name, cleaned_name, status["has_word"], status["has_excel"]
            )
            item.remove_clicked.connect(self.remove_file)
            self.queue_container.addWidget(item)

            if status["has_word"] and status["has_excel"]:
                valid_count += 1

        self.queue_container.addStretch()

    def remove_file(self, cleaned_name):
        """删除文件（使用清理后的文件名）"""
        if cleaned_name in self.file_queue:
            del self.file_queue[cleaned_name]
            # 同时删除Excel信息
            if cleaned_name in self.excel_info:
                del self.excel_info[cleaned_name]
            self.update_queue_display()

    def clear_all(self):
        """清除所有"""
        self.file_queue.clear()
        self.excel_info.clear()
        self.update_queue_display()
        self.days_input.clear()

        # 清空下拉框
        self.excel_sheet_combo.clear()
        self.simple_excel_combo.clear()
        self.level1_combo.clear()
        self.level2_combo.clear()
        self.level3_combo.clear()
        self.func_combo.clear()

        # 恢复默认选项
        self.excel_sheet_combo.addItems(["0", "1", "2"])
        self.simple_excel_combo.addItems(["0", "1", "2"])
        self.level1_combo.addItems(["", "A", "B", "C", "D", "E"])
        self.level2_combo.addItems(["", "A", "B", "C", "D", "E"])
        self.level3_combo.addItems(["", "A", "B", "C", "D", "E"])
        self.func_combo.addItems(["", "A", "B", "C", "D", "E", "F"])

        # 清空列信息显示
        if hasattr(self, "column_info_text"):
            self.column_info_text.setText("选择工作表后，这里会显示Excel的列信息...")
        if hasattr(self, "simple_column_info_text"):
            self.simple_column_info_text.setText(
                "选择工作表后，这里会显示Excel的列信息..."
            )

    def add_new_task(self, task_info):
        """动态添加新上传的任务卡片到主页顶部"""
        from datetime import datetime

        # 判断类型标签（示例：可根据文件名或逻辑判断“结算”或“预算”）
        # 这里简单用“结算”作为默认，你可后续扩展
        type_label = "结算"  # 或根据 task_info['filename'] 判断

        # 设置边框颜色（与 mock 一致）
        border_color = "#2563eb"  # 蓝色，代表“进行中”

        # 构造步骤状态（初始时只有第1步完成）
        steps = [
            {"num": 1, "name": "文件解析", "status": "done"},
            {"num": 2, "name": "层级匹配", "status": "pending"},
            {"num": 3, "name": "模糊比对", "status": "pending"},
            {"num": 4, "name": "人天计算", "status": "pending"},
            {"num": 5, "name": "生成报告", "status": "pending"},
        ]

        # 初始日志
        timestamp = datetime.now().strftime("%H:%M:%S")
        default_log = f"✅ 任务于 {timestamp} 提交，开始解析文件..."
        logs = {
            1: f"✅ 文件解析成功（共 {len(task_info['file_pairs'])} 对文件）",
            2: "等待匹配...",
            3: "等待比对...",
            4: "等待人天计算...",
            5: "等待生成最终报告...",
        }

        # 构造 TaskCard 所需的完整 task_data
        task_data = {
            "filename": task_info["filename"],
            "type_label": type_label,
            "days": task_info["days"],
            "border_color": border_color,
            "steps": steps,
            "default_log": default_log,
            "logs": logs,
            # 保留原始信息，供 ReportDialog 使用
            "raw_task_info": task_info,
        }

        # 创建任务卡片
        card = TaskCard(task_data)

        # 插入到滚动区域的顶部
        # 假设你在 __init__ 中已保存 scroll_layout 为实例变量
        # 如果没有，请先做这一步（见下方“重要提示”）

        self.task_layout.insertWidget(0, card)  # 插入到最上方

    def get_combo_data(self, combo):
        """安全获取下拉框选中的数据（列信息）"""
        index = combo.currentIndex()
        if index <= 0:
            return None
        return combo.itemData(index)

    def start_review(self):
        """开始审核"""
        if not self.days_input.text():
            QMessageBox.warning(
                self, "警告", "必须输入【线上送审人天】才能进行第2步计算！"
            )
            self.days_input.setStyleSheet(
                """
                QLineEdit {
                    padding: 8px 12px;
                    border: 3px solid #ef4444;
                    border-radius: 6px;
                    font-size: 12px;
                    background: palette(window);
                    color: #ef4444;
                }
            """
            )
            return

        valid_files = [
            k for k, v in self.file_queue.items() if v["has_word"] and v["has_excel"]
        ]
        if not valid_files:
            QMessageBox.warning(self, "警告", "没有有效的配对文件！")
            return

        # 使用第一个配对文件的词条作为显示名，并经过深度清洗
        display_name = valid_files[0]
        if valid_files[0] in self.file_queue:
            word_name = self.file_queue[valid_files[0]]["original_names"]["word"]
            if word_name:
                display_name = self.clean_filename(word_name)
            else:
                # 如果没Word，清洗索引Key
                display_name = self.clean_filename(valid_files[0])

        # 获取选中列的表头行信息
        func_col_info = self.func_combo.currentData()
        level1_col_info = self.level1_combo.currentData()

        func_header_row = 0
        if isinstance(func_col_info, dict):
            func_header_row = func_col_info.get("header_row", 0)

        hier_header_row = 0
        if isinstance(level1_col_info, dict):
            hier_header_row = level1_col_info.get("header_row", 0)

        # 获取选中列的索引信息
        func_col_idx = 6  # Default
        if isinstance(func_col_info, dict):
            func_col_idx = func_col_info.get("index", 6)

        l1_col_idx = 1  # Default
        if isinstance(level1_col_info, dict):
            l1_col_idx = level1_col_info.get("index", 1)

        l2_col_info = self.level2_combo.currentData()
        l2_col_idx = 2  # Default
        if isinstance(l2_col_info, dict):
            l2_col_idx = l2_col_info.get("index", 2)

        l3_col_info = self.level3_combo.currentData()
        l3_col_idx = 3  # Default
        if isinstance(l3_col_info, dict):
            l3_col_idx = l3_col_info.get("index", 3)

        # 获取选中的工作表名称 (分别针对两种模式)
        hierarchy_sheet = self.excel_sheet_combo.currentData()
        simple_sheet = self.simple_excel_combo.currentData()

        print(
            f"[DEBUG] start_review: func_col_idx={func_col_idx}, func_header_row={func_header_row}, simple_sheet={simple_sheet}, hierarchy_sheet={hierarchy_sheet}"
        )

        task_info = {
            "filename": display_name,
            "days": self.days_input.text(),
            "file_pairs": [
                {
                    "word": info["file_paths"]["word"],
                    "excel": info["file_paths"]["excel"],
                    "excel_sheets": info.get("worksheets", []),
                }
                for key, info in self.file_queue.items()
                if info["has_word"] and info["has_excel"]
            ],
            "excel_columns": {
                "level1": self.level1_combo.currentText(),
                "level2": self.level2_combo.currentText(),
                "level3": self.level3_combo.currentText(),
                "point": self.func_combo.currentText(),
            },
            "selected_sheet": (
                hierarchy_sheet if self.hierarchy_checkbox.isChecked() else simple_sheet
            ),  # 兼容旧逻辑
            "hierarchy_sheet": hierarchy_sheet,
            "simple_sheet": simple_sheet,
            "functional_header_row": func_header_row,
            "hierarchy_header_row": hier_header_row,
            "functional_column_index": func_col_idx,  # Pass explicit index
            "level1_column_index": l1_col_idx,  # Pass explicit index
            "level2_column_index": l2_col_idx,  # Pass explicit index
            "level3_column_index": l3_col_idx,  # Pass explicit index
            "run_template": self.check_template.isChecked(),
            "run_empty": self.check_empty.isChecked(),
            "run_ratio": self.check_ratio.isChecked(),
            "run_factors": self.check_factors.isChecked(),
            "run_hierarchy": self.hierarchy_checkbox.isChecked(),
            "run_simple": self.simple_checkbox.isChecked(),
            "run_move": self.dm_checkbox.isChecked(),
            "mode": (
                "both"
                if (
                    self.hierarchy_checkbox.isChecked()
                    and self.simple_checkbox.isChecked()
                )
                else ("hierarchy" if self.hierarchy_checkbox.isChecked() else "simple")
            ),
            "fuzzy": self.fuzzy_check.isChecked(),
            "threshold": self.threshold_slider.value() / 100.0,
            "validation_results": [],
        }

        # 发射信号
        self.task_submitted.emit(task_info)

        # 记录日志，但不弹窗阻碍流程，直接关闭即可
        print(f"[OK] 已添加任务：{display_name}")

        self.accept()  # 关闭对话框，返回主界面查看任务进度
