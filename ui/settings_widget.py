import os
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QFrame,
    QScrollArea,
    QSpacerItem,
    QSizePolicy,
    QComboBox,
    QSpinBox,
    QColorDialog,
)
from PySide6.QtCore import Qt, Signal
from extend.matcher_config import MatcherConfig
from utils.path_utils import clear_directory, open_directory


class SettingsWidget(QWidget):
    """设置页面组件"""

    config_updated = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.config = MatcherConfig.load()
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)

        # 标题
        title_label = QLabel("通用设置")
        title_label.setStyleSheet(
            "font-size: 24px; font-weight: bold; margin-bottom: 10px;"
        )
        layout.addWidget(title_label)

        # 滚动区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.setSpacing(30)
        scroll_layout.setAlignment(Qt.AlignTop)

        # 1. 存储路径设置
        storage_group = self._create_section(
            "存储位置设置", "配置生成的报告和回单存放的文件夹路径"
        )

        # 初评路径
        self.initial_path_edit = self._add_path_setting(
            storage_group.layout(),
            "初评文件存放位置",
            self.config.get("storage", {}).get("initial_review", ""),
        )

        # 重评路径
        self.re_review_path_edit = self._add_path_setting(
            storage_group.layout(),
            "重评文件存放位置",
            self.config.get("storage", {}).get("re_review", ""),
        )

        # 回单路径
        self.receipt_path_edit = self._add_path_setting(
            storage_group.layout(),
            "回单文件存放位置",
            self.config.get("storage", {}).get("receipt", ""),
        )

        # 日志路径
        self.logs_path_edit = self._add_path_setting(
            storage_group.layout(),
            "日志文件存放位置",
            self.config.get("storage", {}).get("logs", ""),
        )

        scroll_layout.addWidget(storage_group)

        # 2. 外观主题设置
        theme_group = self._create_section("颜色主题设置", "选择界面的显示模式")

        theme_btn_layout = QHBoxLayout()
        self.light_mode_btn = QPushButton("☀️ 浅色模式")
        self.light_mode_btn.setObjectName("ThemeLightBtn")
        self.light_mode_btn.setCheckable(True)
        self.light_mode_btn.setFixedHeight(45)

        self.dark_mode_btn = QPushButton("🌙 深色模式")
        self.dark_mode_btn.setObjectName("ThemeDarkBtn")
        self.dark_mode_btn.setCheckable(True)
        self.dark_mode_btn.setFixedHeight(45)

        # 初始化选中状态
        is_dark = self.config.get("theme", {}).get("is_dark", False)
        self.light_mode_btn.setChecked(not is_dark)
        self.dark_mode_btn.setChecked(is_dark)

        # 互斥
        theme_btn_layout.addWidget(self.light_mode_btn)
        theme_btn_layout.addWidget(self.dark_mode_btn)

        # 这种逻辑最好在主窗口处理，这里只做标记
        self.light_mode_btn.clicked.connect(lambda: self._toggle_theme_btns(False))
        self.dark_mode_btn.clicked.connect(lambda: self._toggle_theme_btns(True))

        theme_group.layout().addLayout(theme_btn_layout)
        scroll_layout.addWidget(theme_group)

        # 3. 自动化设置
        interaction_group = self._create_section("自动化设置", "配置程序的自动化行为")
        self.auto_open_check = QPushButton("完成后自动打开文件夹")
        self.auto_open_check.setCheckable(True)
        self.auto_open_check.setChecked(
            self.config.get("automation", {}).get("auto_open", True)
        )
        self.auto_open_check.setFixedWidth(200)
        interaction_group.layout().addWidget(self.auto_open_check)

        scroll_layout.addWidget(interaction_group)

        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)

        # 底部按钮
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(0, 10, 0, 0)

        self.save_btn = QPushButton("保存配置")
        self.save_btn.setFixedSize(140, 45)
        self.save_btn.setObjectName("PrimaryBtn")
        self.save_btn.setStyleSheet(
            """
            QPushButton#PrimaryBtn {
                background-color: #2563eb;
                color: white;
                font-weight: bold;
                font-size: 15px;
                border: none;
            }
            QPushButton#PrimaryBtn:hover {
                background-color: #1d4ed8;
            }
        """
        )
        self.save_btn.clicked.connect(self.save_settings)

        self.reset_btn = QPushButton("恢复默认值")
        self.reset_btn.setFixedSize(120, 45)
        self.reset_btn.setStyleSheet("color: #666;")
        self.reset_btn.clicked.connect(self.reset_defaults)

        button_layout.addStretch()
        button_layout.addWidget(self.reset_btn)
        button_layout.addWidget(self.save_btn)
        layout.addLayout(button_layout)

    def _toggle_theme_btns(self, is_dark):
        self.light_mode_btn.setChecked(not is_dark)
        self.dark_mode_btn.setChecked(is_dark)
        # 这里只是 UI 反馈，保存时才真正生效，或者通过信号通知
        # 为了体验直观，可以直接通知主窗口切换
        # 但设置页面通常是点击保存才生效，不过主题可以例外

    def _create_section(self, title, description):
        group = QFrame()
        group.setProperty("class", "SettingsSection")
        group_layout = QVBoxLayout(group)
        group_layout.setContentsMargins(0, 15, 0, 15)  # 减少内边距因为没有边框了
        group_layout.setSpacing(10)

        section_title = QLabel(title)
        section_title.setStyleSheet("font-size: 18px; font-weight: 600;")
        group_layout.addWidget(section_title)

        section_desc = QLabel(description)
        section_desc.setObjectName("SectionDesc")
        section_desc.setStyleSheet("font-size: 13px; margin-bottom: 5px;")
        group_layout.addWidget(section_desc)

        return group

    def _add_path_setting(self, layout, label_text, path_value):
        label = QLabel(label_text)
        layout.addWidget(label)

        row = QHBoxLayout()
        # 统一路径格式
        if path_value:
            path_value = os.path.normpath(path_value)
        edit = QLineEdit(path_value)
        edit.setFixedHeight(35)

        browse_btn = QPushButton("浏览...")
        browse_btn.setFixedSize(70, 35)
        browse_btn.clicked.connect(lambda: self.browse_path(edit))

        open_btn = QPushButton("打开")
        open_btn.setFixedSize(60, 35)
        open_btn.clicked.connect(lambda: open_directory(edit.text()))

        clear_btn = QPushButton("清空")
        clear_btn.setFixedSize(60, 35)
        clear_btn.setStyleSheet(
            "background-color: #fee2e2; color: #b91c1c; border: 1px solid #fecaca;"
        )
        clear_btn.clicked.connect(lambda: self.clear_target_dir(edit.text()))

        row.addWidget(edit)
        row.addWidget(browse_btn)
        row.addWidget(open_btn)
        row.addWidget(clear_btn)
        layout.addLayout(row)

        return edit

    def choose_primary_color(self):
        color = QColorDialog.getColor(self.current_primary_color, self, "选择主题色")
        if color.isValid():
            self.current_primary_color = color.name()
            self.primary_color_btn.setStyleSheet(
                f"background-color: {self.current_primary_color}; border: 1px solid #ccc;"
            )

    def choose_text_light(self):
        color = QColorDialog.getColor(
            self.current_text_light, self, "选择亮模式文字颜色"
        )
        if color.isValid():
            self.current_text_light = color.name()
            self.text_light_btn.setStyleSheet(
                f"background-color: {self.current_text_light}; border: 1px solid #ccc;"
            )

    def choose_text_dark(self):
        color = QColorDialog.getColor(
            self.current_text_dark, self, "选择暗模式文字颜色"
        )
        if color.isValid():
            self.current_text_dark = color.name()
            self.text_dark_btn.setStyleSheet(
                f"background-color: {self.current_text_dark}; border: 1px solid #ccc;"
            )

    def browse_path(self, line_edit):
        path = QFileDialog.getExistingDirectory(
            self, "选择文件夹", line_edit.text() or os.path.expanduser("~")
        )
        if path:
            line_edit.setText(os.path.normpath(path))

    def clear_target_dir(self, path):
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "警告", "文件夹路径非法或不存在")
            return

        reply = QMessageBox.question(
            self,
            "确认清空",
            f"确定要彻底清空文件夹吗？\n{path}\n此操作不可逆！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )

        if reply == QMessageBox.Yes:
            success, msg = clear_directory(path)
            if success:
                QMessageBox.information(self, "成功", "文件夹已清空")
            else:
                QMessageBox.critical(self, "出错", msg)

    def save_settings(self):
        new_storage = {
            "initial_review": (
                os.path.normpath(self.initial_path_edit.text())
                if self.initial_path_edit.text()
                else ""
            ),
            "re_review": (
                os.path.normpath(self.re_review_path_edit.text())
                if self.re_review_path_edit.text()
                else ""
            ),
            "receipt": (
                os.path.normpath(self.receipt_path_edit.text())
                if self.receipt_path_edit.text()
                else ""
            ),
            "logs": (
                os.path.normpath(self.logs_path_edit.text())
                if self.logs_path_edit.text()
                else ""
            ),
        }

        # 确保目录存在
        for k, p in new_storage.items():
            if p and not os.path.exists(p):
                try:
                    os.makedirs(p, exist_ok=True)
                except:
                    pass

        self.config["storage"] = new_storage
        self.config["automation"] = {"auto_open": self.auto_open_check.isChecked()}

        # 主题逻辑现在更简单
        is_dark = self.dark_mode_btn.isChecked()
        # 我们可以暂存这个状态，主窗口可以读取
        self.config["theme"]["is_dark"] = is_dark

        MatcherConfig.save(self.config)
        QMessageBox.information(self, "成功", "设置已保存")
        self.config_updated.emit()

    def reset_defaults(self):
        defaults = MatcherConfig.get_defaults()
        self.initial_path_edit.setText(defaults["storage"]["initial_review"])
        self.re_review_path_edit.setText(defaults["storage"]["re_review"])
        self.receipt_path_edit.setText(defaults["storage"]["receipt"])
        self.logs_path_edit.setText(defaults["storage"]["logs"])
        self.auto_open_check.setChecked(True)
        self.dark_mode_btn.setChecked(False)
        self.light_mode_btn.setChecked(True)
