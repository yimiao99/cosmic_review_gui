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
    QSlider,
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
        # 解压路径
        self.extraction_path_edit = self._add_path_setting(
            storage_group.layout(),
            "解压临时存放位置",
            self.config.get("storage", {}).get("extraction", ""),
        )
        # 默认评估路径
        self.evaluation_folder_edit = self._add_path_setting(
            storage_group.layout(),
            "默认评估选择路径",
            self.config.get("evaluation_folder", ""),
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
        theme_group.layout().addLayout(theme_btn_layout)
        scroll_layout.addWidget(theme_group)

        # 3. 自动化设置
        interaction_group = self._create_section("自动化设置", "配置程序的自动化行为")
        auto_open_layout = QHBoxLayout()
        self.auto_open_check = QPushButton("完成后自动打开文件夹")
        self.auto_open_check.setCheckable(True)
        self.auto_open_check.setChecked(
            self.config.get("automation", {}).get("auto_open", True)
        )
        self.auto_open_check.setFixedWidth(200)
        auto_open_layout.addWidget(self.auto_open_check)
        auto_open_layout.addStretch()
        interaction_group.layout().addLayout(auto_open_layout)

        threshold_layout = QHBoxLayout()
        threshold_label = QLabel("全篇复用模糊匹配阈值:")
        self.threshold_slider = QSlider(Qt.Horizontal)
        self.threshold_slider.setRange(0, 100)
        self.threshold_slider.setValue(self.config.get("automation", {}).get("reuse_threshold", 70))
        self.threshold_val_label = QLabel(f"{self.threshold_slider.value()}%")
        self.threshold_slider.valueChanged.connect(lambda v: self.threshold_val_label.setText(f"{v}%"))
        threshold_layout.addWidget(threshold_label)
        threshold_layout.addWidget(self.threshold_slider)
        threshold_layout.addWidget(self.threshold_val_label)
        interaction_group.layout().addLayout(threshold_layout)
        scroll_layout.addWidget(interaction_group)

        # ================= 【新增】 4. 日志与调试设置 =================
        log_group = self._create_section(
            "日志与调试设置",
            "控制程序运行日志的详细程度。\n⚠️ 注意：选择“详细调试”会产生大量日志并显著降低匹配速度，仅在排查问题时使用。"
        )
        log_level_layout = QHBoxLayout()
        log_level_label = QLabel("日志输出级别:")
        log_level_label.setFixedWidth(120)

        self.log_level_combo = QComboBox()
        self.log_level_combo.setFixedHeight(35)
        # 添加选项，第二个参数为实际存储的值 (Data)
        self.log_level_combo.addItem("关闭 (不记录日志文件)", "NONE")
        self.log_level_combo.addItem("仅错误 (ERROR)", "ERROR")
        self.log_level_combo.addItem("警告 (WARN)", "WARN")
        self.log_level_combo.addItem("常规信息 (INFO) - 推荐", "INFO")
        self.log_level_combo.addItem("详细调试 (DEBUG) - 极慢", "DEBUG")

        # 初始化默认选中项
        current_log_level = self.config.get("log_level", "INFO")
        index = self.log_level_combo.findData(current_log_level)
        if index >= 0:
            self.log_level_combo.setCurrentIndex(index)

        log_level_layout.addWidget(log_level_label)
        log_level_layout.addWidget(self.log_level_combo)
        log_level_layout.addStretch()
        log_group.layout().addLayout(log_level_layout)
        scroll_layout.addWidget(log_group)
        # ============================================================

        # ================= 【新增】 4. 标题清理规则设置 =================
        cleanup_group = self._create_section(
            "标题清理规则",
            "配置在匹配时需要从 Word 标题中去除的内容（逗号分隔）。\n"
            "例如：去除“（一级功能模块）”等无意义后缀，提高匹配准确率。"
        )

        self.title_cleanup_keywords_edit = self._add_keyword_setting(
            cleanup_group.layout(), "标题后缀清理关键字",
            ", ".join(self.config.get("title_cleanup", {}).get("suffix_keywords", [
                "一级功能模块", "二级功能模块", "三级功能模块", "功能模块", "模块"
            ]))
        )

        self.title_cleanup_patterns_edit = self._add_keyword_setting(
            cleanup_group.layout(), "标题清理正则模式",
            ", ".join(self.config.get("title_cleanup", {}).get("regex_patterns", [
                r"[（(][一二三级123]+[级]?功能模块[)）]",
                r"[（(][一二三级123]+级[)）]"
            ]))
        )
        scroll_layout.addWidget(cleanup_group)
        # ============================================================

        # 5. COSMIC 评估规则关键字设置
        rules_group = self._create_section("COSMIC 识别规则", "配置 COSMIC 识别时使用的关键字 (逗号分隔)")
        rules_data = self.config.get("evaluation_rules", {})
        self.write_keywords_edit = self._add_keyword_setting(
            rules_group.layout(), "写入 (W) 关键字",
            ", ".join(rules_data.get("write_keywords", []))
        )
        self.read_keywords_edit = self._add_keyword_setting(
            rules_group.layout(), "读取 (R) 关键字",
            ", ".join(rules_data.get("read_keywords", []))
        )
        self.exit_keywords_edit = self._add_keyword_setting(
            rules_group.layout(), "输出 (X) 关键字",
            ", ".join(rules_data.get("exit_keywords", []))
        )
        self.entry_keywords_edit = self._add_keyword_setting(
            rules_group.layout(), "输入 (E) 关键字",
            ", ".join(rules_data.get("entry_keywords", []))
        )
        self.ew_group_keywords_edit = self._add_keyword_setting(
            rules_group.layout(), "EW类功能(过程)关键字",
            ", ".join(rules_data.get("ew_group_keywords", []))
        )
        self.erx_group_keywords_edit = self._add_keyword_setting(
            rules_group.layout(), "ERX类功能(过程)关键字",
            ", ".join(rules_data.get("erx_group_keywords", []))
        )
        self.ew_reuse_keywords_edit = self._add_keyword_setting(
            rules_group.layout(), "设置EW复用(功能)关键词",
            ", ".join(rules_data.get("ew_reuse_keywords", []))
        )
        self.erx_reuse_keywords_edit = self._add_keyword_setting(
            rules_group.layout(), "设置ERX复用(功能)关键词",
            ", ".join(rules_data.get("erx_reuse_keywords", []))
        )
        self.legacy_keywords_edit = self._add_keyword_setting(
            rules_group.layout(), "利旧/非功能关键字",
            ", ".join(rules_data.get("legacy_keywords", []))
        )
        scroll_layout.addWidget(rules_group)

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

    def _create_section(self, title, description):
        group = QFrame()
        group.setProperty("class", "SettingsSection")
        group_layout = QVBoxLayout(group)
        group_layout.setContentsMargins(0, 15, 0, 15)
        group_layout.setSpacing(10)
        section_title = QLabel(title)
        section_title.setStyleSheet("font-size: 18px; font-weight: 600;")
        group_layout.addWidget(section_title)
        section_desc = QLabel(description)
        section_desc.setObjectName("SectionDesc")
        section_desc.setStyleSheet("font-size: 13px; margin-bottom: 5px; color: #64748b;")
        group_layout.addWidget(section_desc)
        return group

    def _add_path_setting(self, layout, label_text, path_value):
        label = QLabel(label_text)
        layout.addWidget(label)
        row = QHBoxLayout()
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

    def _add_keyword_setting(self, layout, label_text, value):
        label = QLabel(label_text)
        layout.addWidget(label)
        edit = QLineEdit(value)
        edit.setFixedHeight(35)
        edit.setPlaceholderText("例如: 查询, 读取, 获取")
        layout.addWidget(edit)
        return edit

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
            "extraction": (
                os.path.normpath(self.extraction_path_edit.text())
                if self.extraction_path_edit.text()
                else ""
            ),
        }
        for k, p in new_storage.items():
            if p and not os.path.exists(p):
                try:
                    os.makedirs(p, exist_ok=True)
                except:
                    pass
        self.config["storage"] = new_storage

        self.config["automation"] = {
            "auto_open": self.auto_open_check.isChecked(),
            "reuse_threshold": self.threshold_slider.value()
        }
        self.config["evaluation_folder"] = self.evaluation_folder_edit.text().strip()

        # ================= 【新增】 保存日志级别配置 =================
        self.config["log_level"] = self.log_level_combo.currentData()
        # ============================================================

        is_dark = self.dark_mode_btn.isChecked()
        self.config["theme"]["is_dark"] = is_dark

        def split_keywords(text):
            if not text: return []
            import re
            tokens = re.split(r'[,，;；\n\r\t\s、/|\\/]+', text)
            result = []
            for t in tokens:
                clean_t = t.strip()
                if clean_t:
                    result.append(clean_t)
            return result

        rules = {
            "write_keywords": split_keywords(self.write_keywords_edit.text()),
            "read_keywords": split_keywords(self.read_keywords_edit.text()),
            "exit_keywords": split_keywords(self.exit_keywords_edit.text()),
            "entry_keywords": split_keywords(self.entry_keywords_edit.text()),
            "ew_group_keywords": split_keywords(self.ew_group_keywords_edit.text()),
            "erx_group_keywords": split_keywords(self.erx_group_keywords_edit.text()),
            "ew_reuse_keywords": split_keywords(self.ew_reuse_keywords_edit.text()),
            "erx_reuse_keywords": split_keywords(self.erx_reuse_keywords_edit.text()),
            "legacy_keywords": split_keywords(self.legacy_keywords_edit.text())
        }
        self.config["evaluation_rules"] = rules

        # ================= 【新增】 保存标题清理规则 =================
        cleanup_rules = {
            "suffix_keywords": split_keywords(self.title_cleanup_keywords_edit.text()),
            "regex_patterns": split_keywords(self.title_cleanup_patterns_edit.text())
        }
        self.config["title_cleanup"] = cleanup_rules
        # ============================================================

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
        self.threshold_slider.setValue(70)
        self.dark_mode_btn.setChecked(False)
        self.light_mode_btn.setChecked(True)

        # ================= 【新增】 重置日志级别为默认 =================
        self.log_level_combo.setCurrentIndex(self.log_level_combo.findData("INFO"))
        # ============================================================

        rules = defaults.get("evaluation_rules", {})
        self.write_keywords_edit.setText(", ".join(rules.get("write_keywords", [])))
        self.read_keywords_edit.setText(", ".join(rules.get("read_keywords", [])))
        self.exit_keywords_edit.setText(", ".join(rules.get("exit_keywords", [])))
        self.entry_keywords_edit.setText(", ".join(rules.get("entry_keywords", [])))
        self.ew_group_keywords_edit.setText(", ".join(rules.get("ew_group_keywords", [])))
        self.erx_group_keywords_edit.setText(", ".join(rules.get("erx_group_keywords", [])))
        self.ew_reuse_keywords_edit.setText(", ".join(rules.get("ew_reuse_keywords", [])))
        self.erx_reuse_keywords_edit.setText(", ".join(rules.get("erx_reuse_keywords", [])))
        self.legacy_keywords_edit.setText(", ".join(rules.get("legacy_keywords", [])))

        # ================= 【新增】 重置标题清理规则 =================
        self.title_cleanup_keywords_edit.setText(", ".join([
            "一级功能模块", "二级功能模块", "三级功能模块", "功能模块"
        ]))
        self.title_cleanup_patterns_edit.setText(", ".join([
            r"[（(][一二三级123]+[级]?功能模块[)）]",
            r"[（(][一二三级123]+级[)）]"
        ]))
        # ============================================================