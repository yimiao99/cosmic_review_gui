from PySide6.QtWidgets import (
    QFrame,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QWidget,
    QMenu,
    QApplication,
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
from utils.path_utils import get_resource_path, clean_project_name
from utils.runtime_logger import RuntimeLogger


class ValidationWorker(QThread):
    """后台校验线程，支持进度反馈"""

    progress = Signal(int)
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
        RuntimeLogger.clear()
        RuntimeLogger.log(f"开始后台校验任务...")
        try:
            # 模拟初始准备负载
            self.progress.emit(5)
            
            if not self._is_running: return

            raw_info = self.task_data.get("raw_task_info", {})
            file_pairs = raw_info.get("file_pairs", [])
            RuntimeLogger.log(f"获取到文件配对数量: {len(file_pairs)}")
            if not file_pairs:
                RuntimeLogger.log("⚠️ 错误: file_pairs 列表为空，无法继续。", level="ERROR")
                self.finished.emit({})
                return

            # 1. 提取模板
            self.progress.emit(5)
            if not self._is_running: return

            template_path = get_resource_path("folder/附件1XX项目需求说明书V1.0.0.docx")
            RuntimeLogger.log(f"正在检测模板文件: {template_path}")
            template_sections = []
            if os.path.exists(template_path):
                RuntimeLogger.log(f"模板存在，开始提取结构...")
                template_sections = DocumentProcessor.extract_word_structure(
                    template_path
                )
                RuntimeLogger.log(f"模板提取完成，章节数: {len(template_sections)}")
            else:
                RuntimeLogger.log(f"❌ 警告: 模板文件不存在！", level="WARN")

            if not self._is_running: return

            # 2. 提取目标文档
            pair = file_pairs[0]
            RuntimeLogger.log(f"正在解析目标 Word: {os.path.basename(pair['word'])}")
            target_sections = DocumentProcessor.extract_word_structure(pair["word"])
            RuntimeLogger.log(f"目标 Word 提取完成，章节数: {len(target_sections)}")
            self.progress.emit(15)

            if not self._is_running: return

            # 3. 提取 Excel
            RuntimeLogger.log(f"正在解析关联 Excel: {os.path.basename(pair['excel'])}")
            excel_info = DocumentProcessor.extract_excel_info(pair["excel"])
            RuntimeLogger.log(f"Excel 工作表和列信息提取完成。")
            self.progress.emit(20)

            if not self._is_running: return

            # 4. 深度校验 (第一步：Word 模板比对)
            run_template = raw_info.get("run_template", True)
            if run_template:
                RuntimeLogger.log(f"正在进入模型对齐与相似度计算子流程...")
                v_res = SimilarityChecker.validate_template(
                    target_sections, excel_info, template_sections
                )
            else:
                RuntimeLogger.log(f"跳过模板校验节点")
                v_res = {"is_valid": True, "skipped": True}
            
            if not self._is_running: return
            self.step_result.emit(1, v_res)
            self.progress.emit(25)

            if not self._is_running: return

            # 5. Excel 空值校验 (第二步)
            run_empty = raw_info.get("run_empty", True)
            if run_empty:
                RuntimeLogger.log(f"正在进行 Excel 关键列空值扫描...")
                check_sheet = (
                    raw_info.get("hierarchy_sheet")
                    if raw_info.get("run_hierarchy")
                    else raw_info.get("simple_sheet")
                )
                excel_check = DocumentProcessor.check_excel_empty_cells(
                    pair["excel"], sheet_name=check_sheet
                )
                v_res["excel_check"] = excel_check
            else:
                RuntimeLogger.log(f"跳过空值校验节点")
                v_res["excel_check"] = {"is_ok": True, "skipped": True}
            
            if not self._is_running: return
            self.step_result.emit(2, {"excel_check": v_res["excel_check"]})
            self.progress.emit(45)

            if not self._is_running: return

            # 6. 功能匹配校验 (辅助数据)
            # 始终提取用于后续计算
            check_sheet = (
                raw_info.get("hierarchy_sheet")
                if raw_info.get("run_hierarchy")
                else raw_info.get("simple_sheet")
            )
            RuntimeLogger.log(f"正在提取 Excel 模块用于功能匹配...")
            excel_modules = DocumentProcessor.get_excel_modules(
                pair["excel"], sheet_name=check_sheet
            )
            func_match = SimilarityChecker.validate_function_matching(
                target_sections, excel_modules
            )
            v_res["func_match"] = func_match

            if not self._is_running: return

            # 7. 送审比例校验 (第三步)
            run_ratio = raw_info.get("run_ratio", True)
            if run_ratio:
                RuntimeLogger.log(f"正在计算送审比例 (功能点/人天)...")
                fp_count = DocumentProcessor.get_functional_points_count(
                    pair["excel"], sheet_name=check_sheet
                )

                mandays_input = self.task_data.get("days", "0")
                mandays = float(mandays_input)
                RuntimeLogger.log(f"【输入数据检查】线上送审人天原始值: {mandays_input}, 数字化结果: {mandays}")
                RuntimeLogger.log(f"【计算要素】送审功能点: {fp_count}, 送审人天: {mandays}")
                
                ratio = fp_count / mandays if mandays > 0 else 0
                RuntimeLogger.log(f"【计算过程】{fp_count} / {mandays} = {ratio}")
                
                # 判定标准
                if mandays <= 1000:
                    is_ratio_ok = 0.8 <= ratio <= 2.0
                    ratio_range = "0.8 ~ 2.0"
                else:
                    is_ratio_ok = 0.8 <= ratio <= 1.5
                    ratio_range = "0.8 ~ 1.5"
                
                RuntimeLogger.log(f"【判定结果】比例: {round(ratio, 2)}, 合规范围: {ratio_range}, 是否合格: {is_ratio_ok}")

                ratio_res = {
                    "fp_count": fp_count,
                    "mandays": mandays,
                    "ratio": round(ratio, 2),
                    "is_ok": is_ratio_ok,
                    "range": ratio_range,
                }
                v_res["ratio_check"] = ratio_res
            else:
                RuntimeLogger.log(f"跳过送审比例校验节点")
                v_res["ratio_check"] = {"is_ok": True, "skipped": True}

            if not self._is_running: return
            self.step_result.emit(3, {"ratio_check": v_res["ratio_check"]})
            self.progress.emit(65)

            if not self._is_running: return

            # 8. 附加值调整因子校验 (Node 4: Word 因子提取)
            run_factors = raw_info.get("run_factors", True)
            if run_factors:
                RuntimeLogger.log(f"正在提取 Word 附加值调整因子...")
                factors = DocumentProcessor.check_adjustment_factors_in_word(pair["word"])
                v_res["factor_check"] = factors
            else:
                RuntimeLogger.log(f"跳过附加值因子校验节点")
                v_res["factor_check"] = {"skipped": True}

            if not self._is_running: return
            self.step_result.emit(4, {"factor_check": v_res["factor_check"]})
            self.progress.emit(78)

            if not self._is_running: return

            # 9. 层级匹配校验 (Node 5)
            run_hierarchy = raw_info.get("run_hierarchy", True)
            if run_hierarchy:
                RuntimeLogger.log(f"正在进行层级匹配校验...")
                h_header_row = raw_info.get("hierarchy_header_row", 0)
                l1_col = raw_info.get("level1_column_index", 1)
                l2_col = raw_info.get("level2_column_index", 2)
                l3_col = raw_info.get("level3_column_index", 3)
                hier_sheet = raw_info.get("hierarchy_sheet")
                fuzzy = raw_info.get("fuzzy", True)
                threshold = raw_info.get("threshold", 0.8)

                hierarchy_res = DocumentProcessor.validate_hierarchy_matching(
                    pair["word"],
                    pair["excel"],
                    header_row=h_header_row,
                    level1_col=l1_col,
                    level3_col=l3_col,
                    sheet_name=hier_sheet,
                    fuzzy_match=fuzzy,
                    threshold=threshold,
                )
                v_res["hierarchy_res"] = hierarchy_res
                RuntimeLogger.log(f"层级匹配完成: {hierarchy_res.get('statistics', {})}")
            else:
                v_res["hierarchy_res"] = {"is_valid": True, "skipped": True}

            if not self._is_running: return
            self.step_result.emit(5, {"hierarchy_res": v_res["hierarchy_res"]})
            self.progress.emit(88)

            if not self._is_running: return

            # 10. 功能过程校验 (Node 6)
            run_simple = raw_info.get("run_simple", False)
            simple_sheet = raw_info.get("simple_sheet")
            if run_simple:
                RuntimeLogger.log(f"正在进行功能过程校验 (Sheet: {simple_sheet})...")
                f_header_row = raw_info.get("functional_header_row", 0)
                f_col = raw_info.get("functional_column_index", 6)

                process_res = DocumentProcessor.validate_functional_process(
                    pair["word"],
                    pair["excel"],
                    header_row=f_header_row,
                    func_col=f_col,
                    sheet_name=simple_sheet,
                    fuzzy_match=fuzzy,
                    threshold=threshold,
                )
                v_res["process_res"] = process_res
                RuntimeLogger.log(f"功能过程校验完成: {process_res.get('statistics', {})}")
            else:
                v_res["process_res"] = {"is_valid": True, "skipped": True}

            if not self._is_running: return
            self.step_result.emit(6, {"process_res": v_res["process_res"]})
            self.progress.emit(98)

            if not self._is_running: return

            # 11. 自动生成报表
            RuntimeLogger.log(f"正在生成自动评估报告...")
            report_path = ReportGenerator.generate_validation_report(
                self.task_data["filename"], v_res
            )
            v_res["auto_report_path"] = report_path
            RuntimeLogger.log(f"报表生成成功: {report_path}")

            if not self._is_running: return

            # 12. 保存完整运行日志
            log_path = RuntimeLogger.save_to_file(self.task_data["filename"])
            v_res["runtime_log_path"] = log_path
            RuntimeLogger.log(f"完整运行日志已保存: {log_path}")

            self.progress.emit(100)
            self.finished.emit(v_res)
        except Exception as e:
            import traceback
            RuntimeLogger.log(f"❌ 致命错误: {str(e)}", level="ERROR")
            RuntimeLogger.log(traceback.format_exc(), level="DEBUG")
            
            # 即使出错也尝试保存已有日志
            try:
                log_path = RuntimeLogger.save_to_file(self.task_data.get("filename", "ErrorTask"))
                self.task_data["runtime_log_path"] = log_path
            except: pass

            self.finished.emit(
                {
                    "is_valid": False,
                    "error": str(e),
                    "runtime_log_path": getattr(self, "runtime_log_path", None)
                }
            )


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
        self.update_style()
        self._init_ui()

    def update_style(self):
        """根据主题更新边框颜色等特定样式"""
        self.setStyleSheet(
            f"""
            QFrame[class="TaskCard"] {{
                border-radius: 10px;
                border: 1px solid palette(mid);
                border-left: 6px solid {self.task_data['border_color']};
            }}
        """
        )

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)  # 紧凑一点
        layout.setContentsMargins(25, 15, 25, 15)

        # Header
        header = self._create_header()
        layout.addWidget(header)

        # 计时面板
        self.time_label = QLabel("⏱️ 预计完成时间: 计算中... | 实际耗时: 00:00")
        self.time_label.setProperty("class", "task-meta")
        self.time_label.setStyleSheet(
            "font-size: 11px; margin-left: 2px;"
        )
        layout.addWidget(self.time_label)

        # 步骤进度条
        self.steps_widget = StepsWidget(self.task_data["steps"])
        self.steps_widget.node_clicked.connect(self.on_node_clicked)
        self.steps_widget.label_clicked.connect(self.on_label_clicked)
        layout.addWidget(self.steps_widget)

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
                    self.meta_label.setText(f"送审人天: {self.task_data['days']}   送审功能点：{fp_count}")

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
        title.setStyleSheet("font-size: 18px; font-weight: bold; font-family: 'Microsoft YaHei UI';")
        title_row.addWidget(title)

        self.badge = QLabel(self.task_data["type_label"])
        # 检测深色模式
        is_dark = "background-color: #1f2937" in (QApplication.instance().styleSheet() or "")
        
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
            f"{badge_style.get(self.task_data['type_label'], 'background: palette(midlight); color: palette(text); bord-er: 1px solid palette(mid);')} padding: 2px 8px; border-radius: 4px; font-size: 11px;"
        )
        title_row.addWidget(self.badge)
        title_row.addStretch()
        left_layout.addLayout(title_row)

        self.meta_label = QLabel(f"送审人天: {self.task_data['days']} ")
        self.meta_label.setProperty("class", "task-meta")
        self.meta_label.setStyleSheet("font-size: 13px; font-family: 'Microsoft YaHei UI';")
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
            QPushButton { border: 1px solid palette(mid); padding: 6px 12px; border-radius: 6px; font-size: 12px; }
            QPushButton:hover { border-color: #3b82f6; color: #3b82f6; background: palette(alternate-base); }
        """
        )
        self.report_btn.clicked.connect(self.show_summary)
        btn_layout.addWidget(self.report_btn)

        # 停止按钮
        self.stop_btn = QPushButton("🛑 停止")
        self.stop_btn.setStyleSheet(
            """
            QPushButton { border: 1px solid palette(mid); color: #ef4444; padding: 6px 12px; border-radius: 6px; font-size: 12px; }
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

        # 2. 进度平滑“蠕动”
        if self.current_display_progress < self.target_backend_progress:
            # 如果落后太多（跨阶段了），加速追赶
            diff = self.target_backend_progress - self.current_display_progress
            if diff > 10:
                self.current_display_progress += 2.0
            else:
                self.current_display_progress += 0.5
        elif self.current_display_progress < 99.5:
            self.current_display_progress += 0.05

        # 3. 计算当前步骤内的相对进度 (让每个圆圈都有 0-100% 的视觉效果)
        ranges = [(0, 70), (70, 85), (85, 92), (92, 95), (95, 98), (98, 100)]
        idx = max(0, min(len(ranges) - 1, self.current_step_num - 1))
        curr_range = ranges[idx]

        span = curr_range[1] - curr_range[0]
        if span <= 0:
            span = 1
        relative = (self.current_display_progress - curr_range[0]) / span
        step_progress = int(
            max(5, min(95, relative * 100))
        )  # 保持在 5-95 之间蠕动，不跳 100

        self.steps_widget.set_step_progress(self.current_step_num, step_progress)

    def on_validation_progress(self, value):
        """后台报告真实进度点"""
        old_step = self.current_step_num

        # 切换阶段 (根据 ValidationWorker.run 中的 emit 调整)
        if value < 25:
            self.current_step_num = 1
        elif value < 45:
            self.current_step_num = 2
        elif value < 65:
            self.current_step_num = 3
        elif value < 78:
            self.current_step_num = 4
        elif value < 88:
            self.current_step_num = 5
        else:
            self.current_step_num = 6

        # 如果步骤跨越了，前一步骤如果还没显示状态，设为已处理状态 (视觉上表示已跳过/完成)
        if self.current_step_num > old_step:
            for s in range(1, self.current_step_num):
                # 这里使用 "finished" 蓝勾，表示该环节已通过，但最终评价（绿/橙/红）待全部完成后汇总
                # 这样可以解决“未完成前全是绿色”的误导，也可以防止用户以为能点开结果
                curr_status = self.steps_widget.step_nodes[s - 1].status
                if curr_status not in ["done", "fail", "warn", "finished"]:
                    self.steps_widget.set_step_status(s, "finished")

            # 同时也清除之前所有步骤的进度环，确保唯一性
            self.steps_widget.set_step_progress(self.current_step_num, 5)

        self.target_backend_progress = value
        self.update_log(
            self.current_step_num, f"🔍 正在进行比对与校验... (后台进度: {value}%)"
        )

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
                    mandays = ratio_res.get('mandays', 0)
                    upper_limit = "2.0" if mandays <= 1000 else "1.5"
                    
                    if ratio_res.get("is_ok"):
                        log = f"✅送审比例正常，当前送审比例为【{ratio_res.get('ratio', 0)}】,送审功能点：{ratio_res['fp_count']}，送审人天：{ratio_res['mandays']}"
                    else:
                        ratio_val = ratio_res.get('ratio', 0)
                        desc = "过多" if ratio_val >= float(upper_limit) else "过少"
                        log = f"❌送审比例{desc}，当前送审比例为【{ratio_val}】，送审功能点：{ratio_res['fp_count']}，送审人天：{ratio_res['mandays']}"
                    
                    if ratio_res.get("is_ok"):
                        log += f"\n- 送审比例范围在 0.8 ~ {upper_limit}"
                
                # 更新送审功能点显示
                fp_count = ratio_res.get("fp_count")
                if fp_count is not None:
                    self.meta_label.setText(f"送审人天: {self.task_data['days']}   送审功能点：{fp_count}")
                
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
                    missing = []
                    abnormal = []
                    for key in [
                        "distributed",
                        "performance",
                        "reliability",
                        "multiple_sites",
                    ]:
                        f = factors.get(key, {})
                        val = f.get("value")
                        source = f.get("source")

                        # 判定逻辑 (根据用户最新要求：只有特定否定文字才算-1，其他文字算正常)：
                        # 1. 明确没找到关键字 -> 缺失 (异常)
                        # 2. 找到了，值是 "-1" -> 明确的否定文字 (异常)
                        # 3. 找到了，值是 "空"/"无" -> 缺失 (异常)
                        # 4. 找到了，值是 "1" 或其他正数 -> 正常
                        # 5. 找到了，是其他文字描述 -> 属于“有内容”，视作正常

                        str_val = str(val).strip() if val is not None else ""

                        if not f.get("found_in_text"):
                            missing.append(f.get("name"))
                        elif str_val in ["-1", "空", "无"]:
                            # 只有特定负面词(变成-1)或显式的“空/无”才判定为异常
                            abnormal.append(f"{f.get('name')}({str_val})")
                        elif str_val == "1" or (str_val.isdigit() and int(str_val) > 0):
                            valid_v.append(f"{f.get('name')}: {str_val}")
                        elif val:
                            # 只要有任何其他非否定描述，都算正常
                            valid_v.append(f"{f.get('name')}: 有详细描述")
                        else:
                            # 找到了关键字但没提取出值且在表格中，通常视作存在该特性
                            if source == "table":
                                valid_v.append(f"{f.get('name')}: 勾选")
                            else:
                                abnormal.append(f"{f.get('name')}(无有效内容)")

                    if missing:
                        factor_errors.append(
                            f"❌ 缺失质量特性因子项：{', '.join(missing)}。"
                        )
                    if abnormal:
                        factor_errors.append(
                            f"❌ 质量特性因子数值异常：{', '.join(abnormal)}。"
                        )

                    if not factor_errors:
                        status = "done"
                        v_str = " | ".join(valid_v)
                        log = f"✅ 附加值调整因子校验通过。\n• 规模: {scale_val}\n• 质量及特征因子: {v_str}"
                    else:
                        status = "fail"
                        log = "❌ 附加值因子校验异常：\n" + "\n".join(factor_errors)
                        if valid_v:
                            log += "\n\n正常项：\n- " + "\n- ".join(valid_v)
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
                    match_rate_str = stats.get("匹配率", "0%")
                    
                    try:
                        match_rate_val = float(match_rate_str.strip('%')) / 100.0
                    except:
                        match_rate_val = 0

                    total_excel = stats.get("Excel功能点总数", 0)
                    if total_excel == 0:
                        status = "warn"
                        log = "⚠️ 层级匹配未执行：Excel 中未找到有效的三级模块数据。"
                    elif miss_count == 0:
                        status = "done"
                        log = f"✅层级匹配通过（匹配率{match_rate_str}）：所有Excel模块均在大纲中找到。"
                    else:
                        status = "warn"
                        missing_items = hierarchy_res.get("not_found_in_word", [])
                        
                        # 用户要求：同类型的异常（前缀相同）只保留一个示例
                        grouped_logs = {}
                        for item in missing_items:
                            desc = item.get("简略描述", "未知项")
                            if not desc: continue
                            # 按照“：”分割，提取大类
                            prefix = desc.split("：")[0] if "：" in desc else desc
                            if prefix not in grouped_logs:
                                grouped_logs[prefix] = desc
                        
                        unique_logs = list(grouped_logs.values())
                        top_failed = unique_logs[:20]
                        # 确保每一项前面都有换行和圆点
                        log = f"⚠️层级匹配存在异常（匹配率{match_rate_str}）：\n• " + "\n• ".join(top_failed)
                        if len(unique_logs) > 20:
                            log += f"\n• ...等共 {len(unique_logs)} 项不匹配"

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
                        not_found = process_res.get('not_found_in_word', [])
                        names = [f"【{item.get('Excel功能点', '未知')}】" for item in not_found[:2]]
                        names_str = "、".join(names)
                        suffix = "等" if len(not_found) > 2 else ""
                        log = f"⚠️功能过程校验不通过（匹配率：{match_rate}）\n{names_str}{suffix}功能过程在需求规格书未体现"
                self.steps_widget.set_step_status(6, status)
                self.task_data["logs"][6] = log
                self.update_log(6, log)

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
        for i in range(1, 7):
            if i not in self.task_data["logs"]:
                self.steps_widget.set_step_status(i, "fail")
                self.task_data["logs"][i] = "❌ 该步骤未正常完成或被跳过。"

        # 最终展示优先级：报错优先
        final_show_step = 1
        final_log = self.task_data["logs"].get(1, "")

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
        ):
            final_show_step = 5
        elif res_dict.get("process_res", {}).get("statistics", {}).get("缺失项", 0) > 0:
            final_show_step = 6

        final_log = self.task_data["logs"].get(final_show_step, "")
        if res_dict.get("auto_report_path"):
            final_log += f"\n\n📂 完整评估报告已自动保存至根目录。"

        self.update_log(final_show_step, final_log)

    def update_log(self, step_num, text):
        color = "#2563eb" if "🔍" in text else "#10b981"
        if "⚠️" in text or "❌" in text:
            color = "#ef4444"
            
        # 针对 border-left 依然保持动态设置，但移除 background 和 color 等基础样式映射
        self.log_panel.setStyleSheet(f"border-left: 4px solid {color};")
        
        # 转换换行符为 HTML 换行
        html_text = text.replace("\n", "<br/>")
        
        # 使用更灵活的字体控制
        content_style = "font-family: 'Consolas', 'Microsoft YaHei UI'; font-size: 13px;"
        
        self.log_panel.setText(f"<b style='color:{color}; font-family:\"Microsoft YaHei UI\"; font-size:14px;'>第 {step_num} 步:</b><br/><span style='{content_style}'>{html_text}</span>")

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
