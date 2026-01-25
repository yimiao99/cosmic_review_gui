from PySide6.QtWidgets import (
    QFrame,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QWidget,
    QMenu,
    QApplication,
    QProgressBar,
)
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QThread, Signal, QTimer
import time
import os
from .steps_widget import StepsWidget
from .report_dialog import ReportDialog, SummaryDialog
from utils.document_processor import DocumentProcessor
from utils.similarity_checker import SimilarityChecker
from utils.report_generator import ReportGenerator
from utils.path_utils import get_resource_path, clean_project_name, open_directory
from utils.runtime_logger import RuntimeLogger
from extend.matcher_config import MatcherConfig


class ValidationWorker(QThread):
    """后台校验线程，支持进度反馈"""

    progress = Signal(int, str)  # 升级：进度值, 子步骤描述
    step_result = Signal(int, dict)  # step_num, result_data
    finished = Signal(dict)

    def __init__(self, task_data):
        super().__init__()
        self.task_data = task_data
        self._is_running = True

    def stop(self):
        """停止线程"""
        self._is_running = False

    def run(self):
        # 设置项目名称用于日志
        project_name = self.task_data.get("filename", "UnknownProject")
        RuntimeLogger.set_project(project_name)

        RuntimeLogger.log(f"开始后台校验任务...")
        try:
            # 清理之前的缓存，确保使用的是当前任务的文件
            DocumentProcessor.clear_cache()

            # 优化初始进度显示，从很小的值开始
            self.progress.emit(1, "正在初始化校验环境...")

            raw_info = self.task_data.get("raw_task_info", {})
            file_pairs = raw_info.get("file_pairs", [])
            RuntimeLogger.log(f"获取到文件配对数量: {len(file_pairs)}")
            if not file_pairs:
                RuntimeLogger.log(
                    "⚠️ 错误: file_pairs 列表为空，无法继续。", level="ERROR"
                )
                self.finished.emit({})
                return

            # 1. 启动并行加载架构 (核心加速点：多线程并发解析)
            import concurrent.futures

            pair = file_pairs[0]
            template_path = get_resource_path("folder/附件1XX项目需求说明书V1.0.0.docx")

            RuntimeLogger.log(
                f"🚀 [Step 0] 环境准备 - 启动并行架构，并发解析多个文档源..."
            )
            self.progress.emit(2, "正在并发加载 Word/Excel 文件...")

            with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
                # 定义三个核心预加载任务
                future_tpl = executor.submit(
                    DocumentProcessor.load_word_document, template_path
                )
                future_target = executor.submit(
                    DocumentProcessor.load_word_document, pair["word"]
                )
                future_wb = executor.submit(
                    DocumentProcessor.load_excel_workbook, pair["excel"], True
                )
                future_excel_info = executor.submit(
                    DocumentProcessor.extract_excel_info, pair["excel"]
                )

                # 等待加载基础对象
                tpl_doc = future_tpl.result()
                target_doc = future_target.result()
                target_wb = future_wb.result()

                self.progress.emit(8, "正在提取文档深层结构 (目录/层级)...")

                # 开始并发提取详细结构 (利用已加载的对象)
                RuntimeLogger.log(f"🔎 [Step 0] 正在预提取 Word 目录树... (深度解析中)")
                future_tpl_struct = executor.submit(
                    DocumentProcessor.extract_word_structure, tpl_doc
                )
                future_target_struct = executor.submit(
                    DocumentProcessor.extract_word_structure, target_doc
                )
                excel_info = future_excel_info.result()  # 这个通常很快，因为内部有缓存

                template_sections = future_tpl_struct.result()
                target_sections = future_target_struct.result()
                RuntimeLogger.log(
                    f"✅ [Step 0] 预提取完成，共获取 {len(target_sections)} 个三级及以上章节项"
                )

            if not tpl_doc:
                RuntimeLogger.log("⚠️ 模板文件加载失败", level="WARN")
            if not target_doc:
                RuntimeLogger.log("❌ 目标 Word 加载失败", level="ERROR")
                self.finished.emit({"error": "无法加载 Word 文件"})
                return
            if not target_wb:
                RuntimeLogger.log("❌ 目标 Excel 加载失败", level="ERROR")
                self.finished.emit({"error": "无法加载 Excel 文件"})
                return

            self.progress.emit(15, "正在扫描 Excel 工作表列表...")
            RuntimeLogger.log(
                f"✅ [Step 0] 预解析完成 (模板: {len(template_sections)}项, 目标: {len(target_sections)}项)"
            )

            if not self._is_running:
                return

            if not self._is_running:
                return

            v_res = {}  # 保证 v_res 始终存在
            self.progress.emit(15, "正在扫描 Excel 工作表列表...")
            # 自动识别要比对的工作表
            check_sheet = (
                raw_info.get("hierarchy_sheet")
                if raw_info.get("run_hierarchy")
                else raw_info.get("simple_sheet")
            )

            # 4. 节点 1：模板合规性校验
            step1_start = time.time()
            if raw_info.get("run_template", True):
                RuntimeLogger.log(f"正在启动 [Step 1] 模板合规性比对校验...")
                self.progress.emit(20, "正在比对 Word 章节与标准模板...")
                res_s1 = SimilarityChecker.validate_template(
                    target_sections, excel_info, template_sections
                )
                v_res.update(res_s1)
            else:
                v_res.update({"is_valid": True, "skipped": True})

            if not self._is_running:
                return
            v_res["duration"] = time.time() - step1_start
            self.step_result.emit(1, v_res)
            # [UI 对齐] 立即跃迁至第 2 步起始进度 (35%)
            self.progress.emit(35, "正在准备 Excel 空值扫描...")

            # 5. 节点 2：Excel 空值校验
            step2_start = time.time()
            if raw_info.get("run_empty", True):
                RuntimeLogger.log(
                    f"正在启动 [Step 2] Excel 关键列空值扫描 (Sheet: {check_sheet})..."
                )
                self.progress.emit(38, f"正在扫描 Excel ({check_sheet}) 空值行...")
                excel_check_res = DocumentProcessor.check_excel_empty_cells(
                    pair["excel"], sheet_name=check_sheet
                )
                v_res["excel_check"] = excel_check_res
            else:
                v_res["excel_check"] = {"is_ok": True, "skipped": True}

            if not self._is_running:
                return
            v_res["excel_check"]["duration"] = time.time() - step2_start
            self.step_result.emit(2, {"excel_check": v_res["excel_check"]})
            # [UI 对齐] 立即跃迁至第 3 步起始进度 (55%)
            self.progress.emit(55, "正在准备送审比例计算...")

            # 6. 功能匹配校验 (辅助数据)
            RuntimeLogger.log(
                f"正在进行 [辅助步骤] Word 与 Excel 模块名称匹配度计算..."
            )
            # 传入已加载的 WB
            excel_modules = DocumentProcessor.get_excel_modules(
                target_wb, sheet_name=check_sheet
            )
            func_match = SimilarityChecker.validate_function_matching(
                target_sections, excel_modules
            )
            v_res["func_match"] = func_match

            # 7. 节点 3：送审比例校验
            step3_start = time.time()
            if raw_info.get("run_ratio", True):
                RuntimeLogger.log(f"正在进行 [Step 3] 送审比例计算...")
                self.progress.emit(60, "当前正在计算送审功能点与人天比例...")
                fp_count = DocumentProcessor.get_functional_points_count(
                    target_wb, sheet_name=check_sheet
                )
                mandays = float(self.task_data.get("days", 0))
                ratio = fp_count / mandays if mandays > 0 else 0
                r_range = "0.8 ~ 2.0" if mandays <= 1000 else "0.8 ~ 1.5"
                v_res["ratio_check"] = {
                    "fp_count": fp_count,
                    "mandays": mandays,
                    "ratio": round(ratio, 2),
                    "is_ok": (
                        (0.8 <= ratio <= 2.0)
                        if mandays <= 1000
                        else (0.8 <= ratio <= 1.5)
                    ),
                    "range": r_range,
                }
            else:
                v_res["ratio_check"] = {"is_ok": True, "skipped": True}

            if not self._is_running:
                return
            v_res["ratio_check"]["duration"] = time.time() - step3_start
            self.step_result.emit(3, {"ratio_check": v_res["ratio_check"]})
            # [UI 对齐] 立即跃迁至第 4 步起始进度 (72%)
            self.progress.emit(72, "正在准备附加值调整因子提取...")

            # 8. 节点 4：附加值调整因子校验
            step4_start = time.time()
            if raw_info.get("run_factors", True):
                RuntimeLogger.log(f"正在进行 [Step 4] Word 附加值调整因子提取...")
                self.progress.emit(75, "正在扫描文档中的因子表与描述文字...")
                factors = DocumentProcessor.check_adjustment_factors_in_word(target_doc)
                v_res["factor_check"] = factors
            else:
                v_res["factor_check"] = {"skipped": True}

            if not self._is_running:
                return
            v_res["factor_check"]["duration"] = time.time() - step4_start
            self.step_result.emit(4, {"factor_check": v_res["factor_check"]})
            # [UI 对齐] 立即跃迁至第 5 步起始进度 (83%)
            self.progress.emit(83, "正在启动核心层级匹配引擎...")

            # 9. 层级匹配校验 (Node 5)
            step5_start = time.time()
            run_hierarchy = raw_info.get("run_hierarchy", True)
            fuzzy = raw_info.get("fuzzy", True)
            threshold = raw_info.get("threshold", 0.8)

            if run_hierarchy:
                RuntimeLogger.log(f"正在启动 [Step 5] 核心层级匹配校验...")

                def hierarchy_progress_proxy(p, msg):
                    if not self._is_running:
                        return
                    # 映射 83-92 的区间
                    mapped_progress = 83 + int(p * 0.09)
                    self.progress.emit(mapped_progress, f"层级匹配: {msg}")
                    if msg and (
                        "开始" in msg or "完成" in msg or "1/" in msg or "/100" in msg
                    ):
                        RuntimeLogger.log(f"[Step 5] {msg}")

                h_header_row = raw_info.get("hierarchy_header_row", 0)
                l1_col = raw_info.get("level1_column_index", 1)
                l2_col = raw_info.get("level2_column_index", 2)
                l3_col = raw_info.get("level3_column_index", 3)
                hier_sheet = raw_info.get("hierarchy_sheet")

                hierarchy_res = DocumentProcessor.validate_hierarchy_matching(
                    target_doc,
                    pair["excel"],
                    header_row=h_header_row,
                    level1_col=l1_col,
                    level2_col=l2_col,
                    level3_col=l3_col,
                    sheet_name=hier_sheet,
                    fuzzy_match=fuzzy,
                    threshold=threshold,
                    progress_callback=hierarchy_progress_proxy,
                    word_sections_preloaded=target_sections,  # [CORE] 数据透传
                    project_name=self.task_data["filename"],
                )
                v_res["hierarchy_res"] = hierarchy_res
                RuntimeLogger.log(
                    f"层级匹配完成: {hierarchy_res.get('statistics', {})}"
                )
            else:
                v_res["hierarchy_res"] = {"is_valid": True, "skipped": True}

            if not self._is_running:
                return
            v_res["hierarchy_res"]["duration"] = time.time() - step5_start
            self.step_result.emit(5, {"hierarchy_res": v_res["hierarchy_res"]})
            # [UI 对齐] 完成第 5 步后，立即将 UI 文字推进至第 6 步区位 (92%)
            self.progress.emit(92, "正在启动功能过程内容匹配...")

            # 10. 功能过程校验 (Node 6)
            step6_start = time.time()
            run_simple = raw_info.get("run_simple", False)
            simple_sheet = raw_info.get("simple_sheet")
            if run_simple:
                RuntimeLogger.log(
                    f"正在进行功能过程匹配校验 (Sheet: {simple_sheet})..."
                )
                f_header_row = raw_info.get("functional_header_row", 0)
                f_col = raw_info.get("functional_column_index", 6)

                def process_progress_proxy(p, msg):
                    if not self._is_running:
                        return
                    # 映射 92%-96% 的区间
                    mapped_progress = 92 + int(p * 0.04)
                    self.progress.emit(mapped_progress, f"过程匹配: {msg}")
                    if msg and (
                        "开始" in msg or "完成" in msg or "1/" in msg or "/100" in msg
                    ):
                        RuntimeLogger.log(f"[Step 6] {msg}")

                process_res = DocumentProcessor.validate_functional_process(
                    target_doc,
                    pair["excel"],
                    header_row=f_header_row,
                    func_col=f_col,
                    sheet_name=simple_sheet,
                    fuzzy_match=fuzzy,
                    threshold=threshold,
                    progress_callback=process_progress_proxy,
                    word_sections_preloaded=target_sections,  # [CORE] 数据透传
                    project_name=self.task_data["filename"],
                )
                v_res["process_res"] = process_res
            else:
                v_res["process_res"] = {"is_valid": True, "skipped": True}

            if not self._is_running:
                return
            v_res["process_res"]["duration"] = time.time() - step6_start
            self.step_result.emit(6, {"process_res": v_res["process_res"]})
            # [UI 对齐] 完成第 6 步后，立即将 UI 文字推进至第 7 步 (96%)
            self.progress.emit(96, "正在启动数据移动类型校验...")

            # 11. 功能过程数据移动类型校验 (Node 7)
            step7_start = time.time()
            run_move = raw_info.get("run_move", True)
            if run_move:
                RuntimeLogger.log(f"正在进行 [Step 7] 数据移动类型合规性校验...")
                # 传入已加载的 WB 路径路径（此时已缓存）
                f_header_row = raw_info.get("functional_header_row", 0)
                f_col = raw_info.get("functional_column_index", 6)
                move_res = DocumentProcessor.validate_data_movement_types(
                    pair["excel"],  # 使用路径是因为该方法内部做了路径缓存优化
                    sheet_name=simple_sheet,
                    header_row=f_header_row,
                    func_col=f_col,
                    move_col=f_col + 2,
                    project_name=self.task_data["filename"],
                )
                v_res["move_res"] = move_res
            else:
                v_res["move_res"] = {"is_valid": True, "skipped": True}

            if not self._is_running:
                return
            v_res["move_res"]["duration"] = time.time() - step7_start
            self.step_result.emit(7, {"move_res": v_res["move_res"]})
            self.progress.emit(98, "正在进行最后的评估报告汇总...")

            # 12. 自动生成报表
            RuntimeLogger.log(f"正在汇总结果并生成评估报告...")
            report_path = ReportGenerator.generate_validation_report(
                self.task_data["filename"], v_res
            )
            v_res["auto_report_path"] = report_path

            # 13. 保存运行日志
            log_path = RuntimeLogger.save_to_file(self.task_data["filename"])
            v_res["runtime_log_path"] = log_path

            self.progress.emit(100, "所有校验任务已完成")
            self.finished.emit(v_res)

        except Exception as e:
            import traceback

            RuntimeLogger.log(f"❌ 校验任务中断: {str(e)}", level="ERROR")
            RuntimeLogger.log(traceback.format_exc(), level="DEBUG")
            self.finished.emit({"is_valid": False, "error": str(e)})
        finally:
            # 无论成功失败，确保保存一次会话日志
            project_name = self.task_data.get("filename", "Unknown")
            RuntimeLogger.save_to_file(project_name)
            # 任务完成后清理缓存
            DocumentProcessor.clear_cache()


class TaskCard(QFrame):
    """任务卡片，带实时计时和平滑进度条"""

    def __init__(self, task_data, parent=None):
        super().__init__(parent)
        self.task_data = task_data

        # 计时器相关
        self.start_time = time.time()
        self.est_total = 45.0  # 预计初始耗时 45 秒 (更符合实际情况)
        self.current_display_progress = 0
        self.target_backend_progress = 0
        self.current_step_num = 1
        self.is_running = True

        self.setFrameShape(QFrame.StyledPanel)
        self.setProperty("class", "TaskCard")
        self._init_ui()
        self.update_style()

    def update_style(self):
        """根据主题更新边框颜色等特定样式"""
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
            card_bg = "transparent"
            card_border = "#374151"
            text_color = "#f3f4f6"
        else:
            card_bg = "#e5e7eb"
            card_border = "#cbd5e1"
            text_color = "#000000"

        self.setStyleSheet(
            f"""
            QFrame[class="TaskCard"] {{
                background-color: {card_bg};
                border: 2px solid {card_border};
                border-radius: 12px;
                border-left: 6px solid {self.task_data['border_color']};
                padding: 20px;
                margin-bottom: 12px;
            }}
            QLabel {{ color: {text_color}; }}
        """
        )

        # 同步更新子组件样式
        if hasattr(self, "steps_widget"):
            self.steps_widget.update_theme_style()

        # 刷新进度面板背景和边框
        if hasattr(self, "log_panel"):
            # 保持透明背景，仅更新左侧边框颜色
            pass

        # 更新 Badge 样式
        if hasattr(self, "badge"):
            if is_dark:
                badge_style = {
                    "结算": "background: #022c22; color: #10b981; border: 1px solid #064e3b;",
                    "预算": "background: #451a03; color: #f59e0b; border: 1px solid #78350f;",
                }
            else:
                badge_style = {
                    "结算": "background: #ecfdf5; color: #047857; border: 1px solid #6ee7b7;",
                    "预算": "background: #fffbeb; color: #b45309; border: 1px solid #fcd34d;",
                }
            self.badge.setStyleSheet(
                f"{badge_style.get(self.task_data['type_label'], '')} padding: 4px 10px; border-radius: 4px; font-size: 12px; font-family: 'Microsoft YaHei UI';"
            )

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)  # 紧凑一点
        layout.setContentsMargins(16, 16, 16, 16)  # 与re_review_card统一

        # Header
        header = self._create_header()
        layout.addWidget(header)

        # 正在执行状态提示 (新增)
        self.status_bar_layout = QHBoxLayout()
        self.status_icon = QLabel("🚀")
        self.status_label = QLabel("正在准备校验环境...")
        self.status_label.setStyleSheet(
            "font-weight: bold; color: #10b981; font-size: 14px;"
        )
        self.status_bar_layout.addWidget(self.status_icon)
        self.status_bar_layout.addWidget(self.status_label)
        self.status_bar_layout.addStretch()
        layout.addLayout(self.status_bar_layout)

        # 计时面板
        self.time_label = QLabel("⏱️ 预计完成时间: 计算中... | 实际耗时: 00:00")
        self.time_label.setProperty("class", "task-meta")
        self.time_label.setStyleSheet(
            "font-size: 14px; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif; margin-left: 2px;"
        )
        layout.addWidget(self.time_label)

        # 步骤进度条
        self.steps_widget = StepsWidget(self.task_data["steps"])
        self.steps_widget.node_clicked.connect(self.on_node_clicked)
        self.steps_widget.label_clicked.connect(self.on_label_clicked)
        layout.addWidget(self.steps_widget)

        # 全局辅助进度条 (新增：线性反馈万级匹配进度)
        self.full_progress_bar = QProgressBar()
        self.full_progress_bar.setFixedHeight(8)
        self.full_progress_bar.setTextVisible(False)
        self.full_progress_bar.setValue(0)
        self.full_progress_bar.setStyleSheet(
            """
            QProgressBar {
                background-color: #374151;
                border: none;
                border_radius: 4px;
            }
            QProgressBar::chunk {
                background-color: #10b981;
                border-radius: 4px;
            }
        """
        )
        layout.addWidget(self.full_progress_bar)

        # 日志面板
        self.log_panel = QLabel(self.task_data["default_log"])
        self.log_panel.setObjectName("LogPanel")
        self.log_panel.setWordWrap(True)
        layout.addWidget(self.log_panel)

        # 计时器驱动
        self.ui_timer = QTimer(self)
        self.ui_timer.timeout.connect(self.on_timer_tick)
        self.ui_timer.start(100)  # 每0.1秒更新一次视觉，极其丝滑

        # 启动异步校验
        if not self.task_data.get("validation_results"):
            self.worker = ValidationWorker(self.task_data)
            self.worker.progress.connect(self.on_validation_progress)
            self.worker.step_result.connect(self.on_step_finished)
            self.worker.finished.connect(self.on_validation_finished)
            self.worker.start()
        else:
            self.is_running = False
            self.time_label.setText("⏱️ 校验已完成")
            self.stop_btn.setEnabled(False)
            self.stop_btn.setText("已完成")

            # 处理已有结果显示送审功能点
            results = self.task_data["validation_results"][0]
            ratio_res = results.get("ratio_check", {})
            if ratio_res and not ratio_res.get("skipped"):
                fp_count = ratio_res.get("fp_count")
                if fp_count is not None:
                    self.meta_label.setText(
                        f"送审人天: {self.task_data['days']}   送审功能点：{fp_count}"
                    )

    def _create_header(self):
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)

        left_info = QWidget()
        left_layout = QVBoxLayout(left_info)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(5)

        title_row = QHBoxLayout()
        display_name = clean_project_name(self.task_data["filename"])
        title = QLabel(f"📄 {display_name}")
        title.setToolTip(self.task_data["filename"])  # 悬停显示完整名称
        title.setProperty("class", "task-title")
        title.setStyleSheet(
            "font-size: 20px; font-weight: bold; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
        )
        title_row.addWidget(title)

        self.badge = QLabel(self.task_data["type_label"])
        # 检测深色模式
        is_dark = "background-color: #1f2937" in (
            QApplication.instance().styleSheet() or ""
        )

        if is_dark:
            badge_style = {
                "结算": "background: #022c22; color: #10b981; border: 1px solid #064e3b;",
                "预算": "background: #451a03; color: #f59e0b; border: 1px solid #78350f;",
            }
        else:
            badge_style = {
                "结算": "background: #ecfdf5; color: #047857; border: 1px solid #6ee7b7;",
                "预算": "background: #fffbeb; color: #b45309; border: 1px solid #fcd34d;",
            }

        self.badge.setStyleSheet(
            f"{badge_style.get(self.task_data['type_label'], 'background: palette(midlight); color: palette(text); bord-er: 1px solid palette(mid);')} padding: 4px 10px; border-radius: 4px; font-size: 13px; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
        )
        title_row.addWidget(self.badge)
        title_row.addStretch()
        left_layout.addLayout(title_row)

        self.meta_label = QLabel(f"送审人天: {self.task_data['days']} ")
        self.meta_label.setProperty("class", "task-meta")
        self.meta_label.setStyleSheet(
            "font-size: 15px; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif;"
        )
        left_layout.addWidget(self.meta_label)

        header_layout.addWidget(left_info, stretch=1)

        btn_container = QWidget()
        btn_layout = QHBoxLayout(btn_container)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.setSpacing(8)

        # 预览按钮 (打开对话框)
        self.report_btn = QPushButton("📋 结果情况")
        self.report_btn.setStyleSheet(
            """
            QPushButton { border: 1px solid palette(mid); padding: 6px 12px; border-radius: 6px; font-size: 13px; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif; }
            QPushButton:hover { border-color: #3b82f6; color: #3b82f6; background: palette(alternate-base); }
        """
        )
        self.report_btn.clicked.connect(self.show_summary)
        btn_layout.addWidget(self.report_btn)

        # 停止按钮
        self.stop_btn = QPushButton("🛑 停止")
        self.stop_btn.setStyleSheet(
            """
            QPushButton { border: 1px solid palette(mid); color: #ef4444; padding: 6px 12px; border-radius: 6px; font-size: 13px; font-family: 'Microsoft YaHei UI', 'Microsoft YaHei', SimHei, sans-serif; }
            QPushButton:hover { background: palette(alternate-base); border-color: #ef4444; }
            QPushButton:disabled { color: palette(disabled); border-color: palette(mid); opacity: 0.5; }
        """
        )
        self.stop_btn.clicked.connect(self.stop_task)
        btn_layout.addWidget(self.stop_btn)

        header_layout.addWidget(btn_container)
        return header

    def stop_task(self):
        """停止当前校验任务"""
        if hasattr(self, "worker") and self.worker.isRunning():
            self.worker.stop()
            self.is_running = False
            self.stop_btn.setEnabled(False)
            self.stop_btn.setText("已停止")
            self.time_label.setText("⏱️ 校验已手动停止")
            self.update_log(self.current_step_num, "🛑 用户手动停止了校验任务。")

            # 将当前步骤设为警告状态
            self.steps_widget.set_step_status(self.current_step_num, "warn")

            # 停止 UI 计时器
            self.ui_timer.stop()

    def on_timer_tick(self):
        """每秒平滑增长进度并更新计时器"""
        if not self.is_running:
            return

        elapsed = time.time() - self.start_time

        # 1. 更新计时文字 (采用动态评估)
        if self.current_display_progress > 5 and self.current_display_progress < 100:
            # 基于当前进度动态修正总耗时
            real_est_total = elapsed / (self.current_display_progress / 100.0)
            # 平滑切换：预计时间不应突变
            self.est_total = self.est_total * 0.95 + real_est_total * 0.05

        rem = max(1, self.est_total - elapsed)
        # 如果即将完成但进度还没到，稍微等等
        if self.current_display_progress < 95 and rem < 3:
            rem = 3

        self.time_label.setText(
            f"⏱️ 预计剩余: {int(rem//60):02d}:{int(rem%60):02d} | 实际已耗时: {int(elapsed//60):02d}:{int(elapsed%60):02d}"
        )

        # 2. 进度平滑“捕捉”后台真实值
        if self.current_display_progress < self.target_backend_progress:
            # 向目标进度靠近，由原本的由于预跑导致 90% 的策略改为直接追赶
            diff = self.target_backend_progress - self.current_display_progress
            if diff > 5:
                # 较大的差距快速追赶
                self.current_display_progress += 1.0
            else:
                # 小差距平滑逼近
                self.current_display_progress += 0.2
        elif self.current_display_progress < 99.8:
            # 极慢速的前进，表示系统活跃
            self.current_display_progress += 0.01

        # 3. 计算当前步骤内的相对进度 (让圆圈填充效果更精确)
        # [NEW] 线性步进：当前的百分比在步骤区间内的位置
        ranges = [
            (0, 10),
            (10, 25),
            (25, 45),
            (45, 65),
            (65, 78),
            (78, 88),
            (88, 95),
            (95, 100),
        ]
        idx = max(0, min(len(ranges) - 1, self.current_step_num - 1))
        curr_range = ranges[idx]

        span = curr_range[1] - curr_range[0]
        if span <= 0:
            span = 1

        # 视觉进度映射到 0-100% 的渲染值，增加对起始位置的精准控制
        relative = (self.current_display_progress - curr_range[0]) / span
        step_progress = int(max(0, min(100, relative * 100)))

        self.steps_widget.set_step_progress(self.current_step_num, step_progress)

    def on_validation_progress(self, value, sub_step_text=""):
        """后台报告真实进度点"""
        if not self.is_running:
            return

        old_step = self.current_step_num

        # 1. 更新目标进度
        self.target_backend_progress = float(value)
        self.full_progress_bar.setValue(int(value))

        # 2. 定位当前所处的语义阶段 (8个子任务)
        step_names = [
            "环境解析与预加载",
            "模板合规性校验",
            "Excel 空值扫描",
            "送审比例计算",
            "附加值因子提取",
            "层级关系校验",
            "简单过程匹配",
            "报告生成与汇总",
        ]

        # 匹配 ValidationWorker.run 中的 emit 点
        if value < 15:
            current_idx = 0
        elif value < 35:
            current_idx = 1
        elif value < 55:
            current_idx = 2
        elif value < 72:
            current_idx = 3
        elif value < 83:
            current_idx = 4
        elif value < 92:
            current_idx = 5
        elif value < 96:
            current_idx = 6
        else:
            current_idx = 7

        # UI 只有 7 个气泡，将前两步（环境解析+合规校验）合并为 UI 第 1 步
        self.current_step_num = min(7, current_idx if current_idx > 0 else 1)

        # 3. 更新界面状态文字
        if hasattr(self, "status_label"):
            disp_step = self.current_step_num
            step_text = step_names[current_idx]
            # 强化描述：如果有子步骤文字则展示，否则展示大标题
            display_text = sub_step_text if sub_step_text else step_text
            self.status_label.setText(
                f"正在进行: 第 {disp_step} 步 - {display_text}..."
            )

            if value >= 100:
                self.status_icon.setText("✅")
                self.status_label.setText("所有校验任务已完成")
            else:
                self.status_icon.setText("🚀")

        # [NEW] 同步更新日志面板与内存日志池，确保即使点击气泡也能看到最新动态
        # 仅在进度未完成且该步骤尚未有最终结论时，更新中间状态日志，避免覆盖已生成的详情报告
        if value < 100:
            current_log = self.task_data["logs"].get(self.current_step_num, "")
            # 如果已有 ✅ 或 ❌ 标志，说明该步骤已经完成并输出了正式结论，不再更新中间进度描述
            if not (
                current_log.startswith("✅")
                or current_log.startswith("❌")
                or "⚠️" in current_log
            ):
                if sub_step_text:
                    log_msg = f"🔍 {sub_step_text}"
                    self.task_data["logs"][self.current_step_num] = log_msg
                    self.update_log(self.current_step_num, log_msg)
                else:
                    # 基础进度汇报
                    self.update_log(
                        self.current_step_num, f"🔍 正在执行校验逻辑... ({int(value)}%)"
                    )

        # 4. 如果步骤跨越了，处理视觉流转
        if self.current_step_num > old_step:
            for s in range(1, self.current_step_num):
                curr_status = self.steps_widget.step_nodes[s - 1].status
                if curr_status not in ["done", "fail", "warn", "finished"]:
                    self.steps_widget.set_step_status(s, "finished")
            self.steps_widget.set_step_progress(self.current_step_num, 5)

    def on_step_finished(self, step_num, result_part):
        """当单个节点完成时，立即更新该节点的状态和日志"""
        if not self.task_data.get("validation_results"):
            self.task_data["validation_results"] = [{}]

        results = self.task_data["validation_results"][0]
        results.update(result_part)

        # 实时更新该步骤的 UI
        self._update_specific_step_ui(step_num, results)

    def _update_specific_step_ui(self, step_num, results):
        """更新特定步骤的 UI 状态和日志"""
        if step_num == 1:
            if results.get("skipped"):
                status = "skipped"
                log = "⚪ 模板校验已跳过。"
            else:
                suspect_count = len(results.get("suspect_modules", []))
                if results.get("is_valid") and suspect_count == 0:
                    status = "done"
                    log = "✅ 模板校验通过：全层级正文已填充。"
                elif suspect_count > 0:
                    status = "warn"
                    log = f"⚠️ 发现 {suspect_count} 处正文与模板高度相似，疑似未填写！"
                else:
                    status = "fail"
                    log = "❌ 模板校验发现缺失或严重不符。"

            # 加入耗时记录 (Node Logic)
            dur = results.get("duration", 0)
            log += f" (耗时: {dur:.1f}s)"

            self.steps_widget.set_step_status(1, status)
            self.task_data["logs"][1] = log
            self.update_log(1, log)

        elif step_num == 2:
            excel_res = results.get("excel_check", {})
            if excel_res.get("skipped"):
                status = "skipped"
                log = "⚪ 空值检查已跳过。"
            else:
                status = "done"
                log = "✅ Excel 关键列空值检查通过。"
                if not excel_res.get("is_ok"):
                    errors = excel_res.get("errors", [])
                    if errors:
                        status = "fail"
                        log = "❌ Excel 发现空值行：\n- " + "\n- ".join(errors[:5])
                        if len(errors) > 5:
                            log += f"\n...等共 {len(errors)} 项异常"

            # 加入耗时记录
            dur = excel_res.get("duration", 0)
            log += f" (耗时: {dur:.1f}s)"

            self.steps_widget.set_step_status(2, status)
            self.task_data["logs"][2] = log
            self.update_log(2, log)

        elif step_num == 3:
            ratio_res = results.get("ratio_check", {})
            if ratio_res:
                if ratio_res.get("skipped"):
                    status = "skipped"
                    log = "⚪ 送审比例校验已跳过。"
                else:
                    status = "done" if ratio_res.get("is_ok") else "fail"
                    mandays = ratio_res.get("mandays", 0)
                    upper_limit = "2.0" if mandays <= 1000 else "1.5"

                    if ratio_res.get("is_ok"):
                        log = f"✅送审比例正常，当前送审比例为【{ratio_res.get('ratio', 0)}】,送审功能点'：{ratio_res['fp_count']}，送审人天：{ratio_res['mandays']}"
                    else:
                        ratio_val = ratio_res.get("ratio", 0)
                        desc = "过多" if ratio_val >= float(upper_limit) else "过少"
                        log = f"❌送审比例{desc}，当前送审比例为【{ratio_val}】，送审功能点：{ratio_res['fp_count']}，送审人天：{ratio_res['mandays']}"

                    if ratio_res.get("is_ok"):
                        log += f"\n- 送审比例范围在 0.8 ~ {upper_limit}"

                # 加入耗时记录
                dur = ratio_res.get("duration", 0)
                log += f" (耗时: {dur:.1f}s)"

                # 更新送审功能点显示
                fp_count = ratio_res.get("fp_count")
                if fp_count is not None:
                    self.meta_label.setText(
                        f"送审人天: {self.task_data['days']}   送审功能点：{fp_count}"
                    )

                self.steps_widget.set_step_status(3, status)
                self.task_data["logs"][3] = log
                self.update_log(3, log)

        elif step_num == 4:
            factors = results.get("factor_check", {})
            if factors:
                if factors.get("skipped"):
                    status = "skipped"
                    log = "⚪ 附加值因子校验已跳过。"
                else:
                    factor_errors = []
                    scale = factors.get("scale", {})
                    scale_val = scale.get("value")
                    if scale_val in ["结算", "预算"]:
                        self.badge.setText(scale_val)
                        badge_style = {
                            "结算": "background: #ecfdf5; color: #047857; border: 1px solid #6ee7b7;",
                            "预算": "background: #fffbeb; color: #b45309; border: 1px solid #fcd34d;",
                        }
                        self.badge.setStyleSheet(
                            f"{badge_style.get(scale_val)} padding: 2px 8px; border-radius: 4px; font-size: 11px;"
                        )
                    if not scale_val:
                        factor_errors.append("❌ 需求变更规模因子缺失或未识别。")
                    elif scale_val != "结算":
                        factor_errors.append(
                            f"❌ 需求变更规模因子异常：当前为【{scale_val}】，标准应为【结算】。"
                        )

                    valid_v = []
                    text_missing = []
                    table_missing = []
                    abnormal = []

                    for key in [
                        "distributed",
                        "performance",
                        "reliability",
                        "multiple_sites",
                    ]:
                        f = factors.get(key, {})
                        val = f.get("value")
                        text_val = f.get("text_value")
                        table_val = f.get("table_value")
                        name = f.get("name")

                        # 1. 检查文字描述是否存在
                        # 如果 text_value 为 None 或 "缺失"，视为文字部分缺失
                        if not text_val or text_val == "缺失":
                            text_missing.append(name)

                        # 2. 检查总结表格是否存在
                        if not table_val or table_val == "缺失":
                            table_missing.append(name)

                        # 3. 检查数值异常 (核心判定逻辑)
                        str_val = str(val).strip()
                        if str_val in ["-1", "空", "无"]:
                            # 只有明确的负面描述或缺失才计入异常
                            abnormal.append(f"{name}({str_val})")
                        elif str_val == "1" or (str_val.isdigit() and int(str_val) > 0):
                            valid_v.append(f"{name}: {str_val}")
                        elif val and val != "缺失":
                            valid_v.append(f"{name}: 有描述/勾选")
                        else:
                            # 确实没有找到有效值
                            abnormal.append(f"{name}(未识别)")

                    # 4. 汇总错误信息
                    if text_missing:
                        safe_text_missing = [str(x) for x in text_missing if x]
                        factor_errors.append(
                            f"❌ 文字描述因子项缺失：{', '.join(safe_text_missing)}。"
                        )
                    if table_missing:
                        factor_errors.append(
                            f"❌ 总结表格因子项缺失：{', '.join(table_missing)}。"
                        )

                    # 检查一致性 (文字有的表格也得有，文字没的表格也不能有)
                    text_set = set(
                        key
                        for key in [
                            "distributed",
                            "performance",
                            "reliability",
                            "multiple_sites",
                        ]
                        if factors.get(key, {}).get("text_value")
                        and factors.get(key, {}).get("text_value") != "缺失"
                    )
                    table_set = set(
                        key
                        for key in [
                            "distributed",
                            "performance",
                            "reliability",
                            "multiple_sites",
                        ]
                        if factors.get(key, {}).get("table_value")
                        and factors.get(key, {}).get("table_value") != "缺失"
                    )

                    if text_set != table_set:
                        factor_errors.append(
                            "❌ 一致性异常：文字描述的因子集合与总结表格不一致。"
                        )

                    if abnormal:
                        factor_errors.append(
                            f"❌ 因子数值异常或缺失：{', '.join(abnormal)}。"
                        )

                    if not factor_errors:
                        status = "done"
                        v_str = " | ".join(valid_v)
                        log = f"✅ 附加值调整因子校验通过。\n• 规模: {scale_val}\n• 质量及特征因子: {v_str}"
                    else:
                        status = "fail"
                        log = "❌ 附加值因子校验异常：\n" + "\n".join(factor_errors)
                        if valid_v:
                            log += "\n\n部分正常项：\n- " + "\n- ".join(valid_v)

                # 加入耗时记录
                dur = factors.get("duration", 0)
                log += f" (耗时: {dur:.1f}s)"

                self.steps_widget.set_step_status(4, status)
                self.task_data["logs"][4] = log
                self.update_log(4, log)

        elif step_num == 5:
            hierarchy_res = results.get("hierarchy_res", {})
            if hierarchy_res:
                if hierarchy_res.get("skipped"):
                    status = "skipped"
                    log = "⚪ 层级匹配校验已跳过。"
                else:
                    stats = hierarchy_res.get("statistics", {})
                    miss_count = stats.get("缺失项", 0)
                    mismatch_count = stats.get("层级不匹配", 0)
                    match_rate_str = stats.get("匹配率", "0%")

                    try:
                        match_rate_val = float(match_rate_str.strip("%")) / 100.0
                    except:
                        match_rate_val = 0

                    total_excel = stats.get("Excel功能点总数", 0)
                    if total_excel == 0:
                        status = "warn"
                        log = "⚠️ 层级匹配未执行：Excel 中未找到有效的三级模块数据。"
                    elif miss_count == 0 and mismatch_count == 0:
                        status = "done"
                        log = f"✅层级匹配通过（通过率{match_rate_str}）：所有Excel模块均在大纲中找到。"
                    else:
                        # 只要有缺失或者不匹配，标记为 fail
                        status = "fail"

                        not_found_items = hierarchy_res.get("not_found_in_word", [])
                        mismatched_items = hierarchy_res.get("hierarchy_mismatched", [])
                        all_failed_items = not_found_items + mismatched_items

                        issue_details = []

                        # 1. 汇总缺失项
                        miss_list = []
                        seen_miss_prefixes = set()
                        for item in all_failed_items:
                            desc = item.get("缺失简略描述")
                            if not desc or desc == "-":
                                if item.get("匹配状态") == "缺失":
                                    desc = item.get("简略描述")

                            if desc and desc != "-":
                                prefix = desc.split("：")[0] if "：" in desc else desc
                                if prefix not in seen_miss_prefixes:
                                    miss_list.append(desc)
                                    seen_miss_prefixes.add(prefix)

                        if miss_list:
                            issue_details.append(
                                f"• [缺失项] (Excel在Word未体现): \n  - "
                                + "\n  - ".join(miss_list[:10])
                            )

                        # 2. 汇总层级不匹配项
                        match_list = []
                        seen_match_prefixes = set()
                        for item in all_failed_items:
                            desc = item.get("层级不匹配简略描述")
                            if not desc or desc == "-":
                                if item.get("匹配状态") == "层级不匹配":
                                    d_gen = item.get("简略描述", "")
                                    if "不匹配" in d_gen:
                                        desc = d_gen

                            if desc and desc != "-":
                                prefix = desc.split("：")[0] if "：" in desc else desc
                                if prefix not in seen_match_prefixes:
                                    match_list.append(desc)
                                    seen_match_prefixes.add(prefix)

                        if match_list:
                            issue_details.append(
                                f"• [层级不匹配] (对应关系错误): \n  - "
                                + "\n  - ".join(match_list[:10])
                            )

                        total_issues = len(not_found_items) + len(mismatched_items)
                        log = (
                            f"❌层级匹配存在异常（通过率{match_rate_str}，共{total_issues}处问题）：\n"
                            + "\n".join(issue_details)
                        )

                        log += "\n\n📂 [提示]：点击上方圆圈图标可直接打开详细的 Excel 匹配报告。"

                # 加入耗时记录
                dur = hierarchy_res.get("duration", 0)
                log += f" (耗时: {dur:.1f}s)"

                self.steps_widget.set_step_status(5, status)
                self.task_data["logs"][5] = log
                self.update_log(5, log)

        elif step_num == 6:
            process_res = results.get("process_res", {})
            if process_res:
                if process_res.get("skipped"):
                    status = "skipped"
                    log = "⚪ 功能过程校验已跳过。"
                else:
                    stats = process_res.get("statistics", {})
                    miss_count = stats.get("缺失项", 0)
                    match_rate = stats.get("匹配率", "0%")
                    if miss_count == 0:
                        status = "done"
                        log = f"✅功能过程校验通过（匹配率{match_rate}）：所有功能过程描述均在正文中找到。"
                    else:
                        status = "warn"
                        not_found = process_res.get("not_found_in_word", [])
                        names = [
                            f"【{item.get('Excel功能点', '未知')}】"
                            for item in not_found[:2]
                        ]
                        names_str = "、".join(names)
                        suffix = "等" if len(not_found) > 2 else ""
                        log = f"⚠️功能过程校验不通过（匹配率：{match_rate}）\n{names_str}{suffix}功能过程在需求规格书未体现"

                    log += "\n\n📂 [提示]：点击上方圆圈图标可直接打开详细的 Excel 功能过程匹配报告。"

                # 加入耗时记录
                dur = process_res.get("duration", 0)
                log += f" (耗时: {dur:.1f}s)"

                self.steps_widget.set_step_status(6, status)
                self.task_data["logs"][6] = log
                self.update_log(6, log)

        elif step_num == 7:
            move_res = results.get("move_res", {})
            if move_res:
                if move_res.get("skipped"):
                    status = "skipped"
                    log = "⚪ 数据移动类型校验已跳过。"
                else:
                    stats = move_res.get("statistics", {})
                    failed_count = stats.get("不合规", 0)
                    total_count = stats.get("总数", 0)
                    passed_count = stats.get("合规", 0)

                    if failed_count == 0 and total_count > 0:
                        status = "done"
                        log = f"✅数据移动类型校验通过（合规率 100%）：监测到 {total_count} 个功能过程，全部符合 E 开头、W/X 结束的规则。"
                    elif total_count == 0:
                        status = "warn"
                        log = "⚠️ 未发现有效的功能过程数据移动类型数据。"
                    else:
                        status = "fail"
                        # 获取所有不合规项
                        failed_list = [
                            r
                            for r in move_res.get("items", [])
                            if r.get("result") != "合规"
                        ]

                        # 按错误类型分组
                        err_types = {}
                        for item in failed_list:
                            err = item.get("result", "不合规")
                            if err not in err_types:
                                err_types[err] = []
                            err_types[err].append(item.get("process", "未知"))

                        err_details = []
                        for err, procs in err_types.items():
                            sample = "、".join(procs[:3]) + (
                                "等" if len(procs) > 3 else ""
                            )
                            err_details.append(
                                f"  - [{err}]: {sample} ({len(procs)}处)"
                            )

                        log = (
                            f"❌数据移动类型异常：共监测 {total_count} 处，其中 {failed_count} 处不合规。\n"
                            + "\n".join(err_details)
                            + "\n\n📂 [提示]：请点击上方圆圈图标查看详细的 Excel 数据移动校验报告。"
                        )

                # 加入耗时记录
                dur = move_res.get("duration", 0)
                log += f" (耗时: {dur:.1f}s)"

                self.steps_widget.set_step_status(7, status)
                self.task_data["logs"][7] = log
                self.update_log(7, log)

    def on_validation_finished(self, results):
        self.is_running = False
        self.ui_timer.stop()
        if hasattr(self, "stop_btn"):
            self.stop_btn.setEnabled(False)

        # 清除所有正在转动的进度环
        self.steps_widget.clear_all_progress()

        elapsed = time.time() - self.start_time
        self.time_label.setText(
            f"⏱️ 校验完成 | 总耗时: {int(elapsed//60):02d}:{int(elapsed%60):02d}"
        )

        if not results:
            return
        # 确保最终结果完整
        if not self.task_data.get("validation_results"):
            self.task_data["validation_results"] = [results]
        else:
            self.task_data["validation_results"][0].update(results)

        self.current_display_progress = 100

        # 最终回顾：如果某个步骤还没更新状态（比如因为出错跳过了），标记为 fail
        for i in range(1, 8):
            if i not in self.task_data["logs"]:
                self.steps_widget.set_step_status(i, "fail")
                self.task_data["logs"][i] = "❌ 该步骤未正常完成或被跳过。"

        # 最终展示优先级：报错优先，无错则显示最后一步结果
        final_show_step = 7

        res_dict = self.task_data["validation_results"][0]
        if res_dict.get("excel_check", {}).get("is_ok") == False:
            final_show_step = 2
        elif res_dict.get("ratio_check", {}).get("is_ok") == False:
            final_show_step = 3
        elif (
            res_dict.get("factor_check")
            and self.steps_widget.step_nodes[3].status != "done"
        ):
            final_show_step = 4
        elif (
            res_dict.get("hierarchy_res", {}).get("statistics", {}).get("缺失项", 0) > 0
            or res_dict.get("hierarchy_res", {})
            .get("statistics", {})
            .get("层级不匹配", 0)
            > 0
        ):
            final_show_step = 5
        elif res_dict.get("process_res", {}).get("statistics", {}).get("缺失项", 0) > 0:
            final_show_step = 6

        # 兜底：如果选中的步骤没有任何日志（可能被跳过），则按倒序找第一个有日志的
        if final_show_step not in self.task_data["logs"]:
            for i in range(7, 0, -1):
                if i in self.task_data["logs"]:
                    final_show_step = i
                    break

        final_log = self.task_data["logs"].get(final_show_step, "")
        if res_dict.get("auto_report_path"):
            final_log += f"\n\n📂 完整评估报告已自动保存至：{os.path.dirname(res_dict.get('auto_report_path'))}"

        self.update_log(final_show_step, final_log)

        # 自动化：完成后自动打开文件夹
        config = MatcherConfig.load()
        if config.get("automation", {}).get("auto_open", True):
            report_path = res_dict.get("auto_report_path")
            if report_path:
                open_directory(os.path.dirname(report_path))

    def update_log(self, step_num, text):
        color = "#2563eb" if "🔍" in text else "#10b981"
        if "⚠️" in text or "❌" in text:
            color = "#ef4444"

        # 针对 border-left 依然保持动态设置，同时确保背景透明且文字颜色正确
        self.log_panel.setStyleSheet(
            f"border-left: 4px solid {color}; background-color: transparent;"
        )

        # 转换换行符为 HTML 换行
        html_text = text.replace("\n", "<br/>")

        # 使用更灵活的字体控制
        content_style = (
            "font-family: 'Consolas', 'Microsoft YaHei UI'; font-size: 15px;"
        )

        self.log_panel.setText(
            f"<b style='color:{color}; font-family:\"Microsoft YaHei UI\"; font-size:16px;'>第 {step_num} 步:</b><br/><span style='{content_style}'>{html_text}</span>"
        )

    def on_label_clicked(self, step_num):
        """点击文字：仅展示日志"""
        log_text = self.task_data["logs"].get(step_num, "⏳ 等待处理...")
        self.update_log(step_num, log_text)

    def on_node_clicked(self, step_num):
        """点击圆圈：完成看报告，未完成看日志"""
        if step_num == 1:
            if not self.is_running and self.task_data.get("validation_results"):
                self.show_report()
            else:
                self.on_label_clicked(step_num)
        elif step_num == 2:
            # 尝试打开Excel空值检查报告
            results = self.task_data.get("validation_results", [{}])[0]
            path = results.get("excel_check", {}).get("report_path")
            if path and os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception as e:
                    self.update_log(2, f"❌ 无法打开报告文件: {e}")
            else:
                self.on_label_clicked(step_num)
        elif step_num == 5:
            # 尝试打开层级匹配报告
            results = self.task_data.get("validation_results", [{}])[0]
            path = results.get("hierarchy_res", {}).get("report_path")
            if path and os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception as e:
                    self.update_log(5, f"❌ 无法打开报告文件: {e}")
            else:
                self.on_label_clicked(step_num)
        elif step_num == 6:
            # 尝试打开功能过程报告
            results = self.task_data.get("validation_results", [{}])[0]
            path = results.get("process_res", {}).get("report_path")
            if path and os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception as e:
                    self.update_log(6, f"❌ 无法打开报告文件: {e}")
            else:
                self.on_label_clicked(step_num)
        elif step_num == 7:
            # 尝试打开数据移动类型报告
            results = self.task_data.get("validation_results", [{}])[0]
            path = results.get("move_res", {}).get("report_path")
            if path and os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception as e:
                    self.update_log(7, f"❌ 无法打开报告文件: {e}")
            else:
                self.on_label_clicked(step_num)
        else:
            self.on_label_clicked(step_num)

    def on_step_clicked(self, step_num):
        # 兼容旧代码调用
        self.on_node_clicked(step_num)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        view_report_action = QAction("📄 查看详细报告", self)
        view_report_action.triggered.connect(self.show_report)
        menu.addAction(view_report_action)

        # 添加打开日志选项
        open_log_action = QAction("📂 打开运行日志", self)
        open_log_action.triggered.connect(self.open_runtime_log)
        menu.addAction(open_log_action)

        copy_name_action = QAction("📋 复制项目名称", self)
        copy_name_action.triggered.connect(
            lambda: QApplication.clipboard().setText(self.task_data["filename"])
        )
        menu.addAction(copy_name_action)
        menu.exec(event.globalPos())

    def open_runtime_log(self):
        results = self.task_data.get("validation_results", [{}])[0]
        log_path = results.get("runtime_log_path")
        if log_path and os.path.exists(log_path):
            try:
                os.startfile(log_path)
            except Exception as e:
                self.update_log(1, f"❌ 无法打开日志文件: {e}")
        else:
            self.update_log(1, "⏳ 日志文件尚不存在或任务尚未完成。")

    def show_report(self):
        if not self.task_data.get("validation_results"):
            return
        dialog = ReportDialog(self.task_data, self)
        dialog.exec()

    def show_summary(self):
        dialog = SummaryDialog(self.task_data, self)
        dialog.exec()
