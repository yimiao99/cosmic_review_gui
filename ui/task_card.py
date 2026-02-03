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
    QStackedWidget,
    QTextBrowser,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QButtonGroup,
    QScrollArea,
    QGridLayout,
)
from PySide6.QtGui import QAction, QColor
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
from extend.hierarchical_matcher import HierarchicalMatcher


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
                f"🚀 [Step 0] 环境准备 - 启动并行架构，Word文档处理前置到此阶段完成..."
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

                # === 新增：数据预处理阶段 ===
                # 为功能过程匹配预处理数据，避免在Step 6时重复读取
                run_simple = raw_info.get("run_simple", False)
                word_preprocessing_success = False  # [FIX] 标记预处理是否成功

                # [核心优化] 将Word文档打开和大纲提取完全移到Step 0中完成
                if run_simple and pair["word"] and pair["excel"]:
                    RuntimeLogger.log(
                        f"🔄 [Step 0.5] 开始预处理Word和Excel数据以优化后续匹配..."
                    )
                    self.progress.emit(10, "Word文档打开与大纲提取...")

                    try:
                        # [关键改进] 使用word_outline_extractor直接提取完整大纲（1170项）
                        # 而不是prepare_word_data_async（只获取4项）
                        def word_prep_progress(p, msg):
                            mapped_p = 10 + int(p * 0.12)  # 10%-22%
                            self.progress.emit(mapped_p, f"Word预处理: {msg}")

                        RuntimeLogger.log(
                            f"  [INFO] 🚀 [Step 0] 开始完整Word大纲及内容预处理..."
                        )

                        if pair["word"] and os.path.exists(pair["word"]):
                            try:
                                # [FIX] 使用 HierarchicalMatcher.prepare_word_data_async
                                # 它会使用 win32 稳定提取大纲并在后台读取全文正文
                                matcher = HierarchicalMatcher()

                                word_prep_progress(
                                    10, "正在启动 Word 结构及全文预提取..."
                                )

                                # 执行完整预处理（包含大纲、全文、Excel内容）
                                word_cache = matcher.prepare_word_data_async(
                                    pair["word"]
                                )

                                if word_cache and word_cache.get("items"):
                                    word_items = word_cache["items"]
                                    full_text_content = word_cache.get(
                                        "full_text_content", []
                                    )

                                    word_prep_progress(
                                        60,
                                        f"已提取 {len(word_items)} 个章节，正在集成数据...",
                                    )

                                    # [NEW] 提前在Step 0.5提取Excel内容，确保一致性
                                    word_prep_progress(
                                        70, "正在预加载 Excel 功能点数据..."
                                    )
                                    try:
                                        excel_config = raw_info
                                        excel_sheet = excel_config.get(
                                            "simple_sheet", 2
                                        )
                                        excel_header = excel_config.get(
                                            "functional_header_row", 0
                                        )
                                        excel_col = excel_config.get(
                                            "functional_column_index", 6
                                        )

                                        excel_data = matcher.extract_excel_content(
                                            pair["excel"],
                                            mode="flat",
                                            sheet_name=excel_sheet,
                                            header=excel_header,
                                            column=excel_col,
                                            level1_col=1,
                                            level2_col=2,
                                            level3_col=3,
                                        )
                                        word_cache["excel_data"] = excel_data
                                        RuntimeLogger.log(
                                            f"  [INFO] Excel内容预加载完成: {len(excel_data) if excel_data else 0} 项"
                                        )
                                    except Exception as e:
                                        RuntimeLogger.log(
                                            f"  [WARN] Excel预加载失败: {str(e)[:80]}",
                                            level="WARN",
                                        )
                                        word_cache["excel_data"] = None

                                    if not hasattr(self, "data_cache"):
                                        self.data_cache = {}
                                    self.data_cache["word_data"] = word_cache

                                    word_prep_progress(
                                        100,
                                        f"完成: {len(word_items)} 项已缓存，包含全文内容",
                                    )
                                    RuntimeLogger.log(
                                        f"✅ [Step 0] Word完整数据就绪: {len(word_items)} 章节, 全文 {len(full_text_content)} 项"
                                    )
                                    word_preprocessing_success = True
                                else:
                                    RuntimeLogger.log(
                                        f"⚠️ Word预处理返回空结果", level="WARN"
                                    )

                            except Exception as e:
                                RuntimeLogger.log(
                                    f"⚠️ Word数据预处理异常: {str(e)[:150]}",
                                    level="ERROR",
                                )
                        else:
                            RuntimeLogger.log(f"⚠️ Word文件不存在", level="WARN")

                    except Exception as e:
                        RuntimeLogger.log(f"⚠️ Word数据预处理异常: {e}")
                        RuntimeLogger.log(f"  [INFO] 将使用标准Word处理流程")

                self.progress.emit(22, "正在继续提取文档结构...")

                # 开始并发提取详细结构 (利用已加载的对象)
                # 【优先使用稳定提取】使用 HierarchicalMatcher 替代 DocumentProcessor

                # [优化] 只有在需要 Word 相关内容时才执行昂贵的结构提取
                run_template = raw_info.get("run_template", True)
                run_factors = raw_info.get("run_factors", True)
                run_hierarchy = raw_info.get("run_hierarchy", True)

                need_target_structure = any(
                    [run_template, run_factors, run_hierarchy, run_simple]
                )
                # 只有 Step 1 模板校验真正需要解析模板文档结构
                need_tpl_structure = run_template

                template_sections = []
                target_sections = []

                # [DEBUG] 调试标志状态
                RuntimeLogger.log(
                    f"[DEBUG] word_preprocessing_success={word_preprocessing_success}, need_target_structure={need_target_structure}"
                )
                RuntimeLogger.log(
                    f"[DEBUG] run_template={run_template}, run_factors={run_factors}, run_hierarchy={run_hierarchy}, run_simple={run_simple}"
                )

                # [FIX] 如果Word预处理已成功，则跳过重复的Word结构提取
                if need_target_structure and not word_preprocessing_success:
                    RuntimeLogger.log(
                        f"🔎 [Step 0] 正在预提取 Word 目录树... (优先使用稳定大纲提取)"
                    )

                    matcher = HierarchicalMatcher()

                    # 修改为接受路径 and 对象，优先使用 COM 接口（获取准确编号 and 过滤正文）
                    def extract_word_hierarchy(path, doc_obj):
                        """包装函数：提取 Word 内容结构，包含章节及正文"""

                        def extraction_progress_proxy(p, msg):
                            """转换提取进度为 UI 友好的进度点"""
                            mapped_p = 8 + int(p * 0.1)
                            self.progress.emit(mapped_p, msg)

                        try:
                            # ===== 核心方案：使用 HierarchicalMatcher 提供的稳健提取 =====
                            if path and os.path.exists(path):
                                RuntimeLogger.log(
                                    f"  [INFO] 🌟 使用 HierarchicalMatcher 提取内容结构: {os.path.basename(path)}"
                                )
                                extraction_progress_proxy(
                                    0, "正在通过 Win32 接口提取 Word 结构及内容..."
                                )

                                matcher = HierarchicalMatcher()
                                result = matcher.extract_word_outline_as_hierarchy(
                                    path, include_content=True
                                )

                                if result:
                                    # [FIX] extract_word_outline_as_hierarchy 返回的是字典，需要提取 all_items 列表
                                    items = (
                                        result
                                        if isinstance(result, list)
                                        else result.get("all_items", [])
                                    )
                                    extraction_progress_proxy(100, "Word 结构提取完成")
                                    RuntimeLogger.log(
                                        f"  [OK] ✅ 内容结构提取成功: {len(items)} 项 (含正文内容)"
                                    )
                                    return items

                            # ===== 降级方案 A：DocumentProcessor 稳定模式 =====
                            RuntimeLogger.log(
                                f"  [WARN] ⚠️ 降级使用 DocumentProcessor.extract_word_structure(use_stable=True)..."
                            )
                            result = DocumentProcessor.extract_word_structure(
                                path,
                                use_stable=True,
                                progress_callback=extraction_progress_proxy,
                            )
                            if result and len(result) > 0:
                                RuntimeLogger.log(
                                    f"  [OK] DocumentProcessor 提取成功: {len(result)} 项"
                                )
                                return result

                            # ===== 降级方案 B：HierarchicalMatcher 从 docx 对象提取 =====
                            RuntimeLogger.log(
                                f"  [WARN] ⚠️ 降级使用 HierarchicalMatcher.extract_outline_from_docx_object..."
                            )
                            result = matcher.extract_outline_from_docx_object(doc_obj)
                            if result.get("all_items") and len(result["all_items"]) > 0:
                                RuntimeLogger.log(
                                    f"  [OK] HierarchicalMatcher 提取成功: {len(result['all_items'])} 项"
                                )
                                return result["all_items"]

                            raise Exception("无法通过任何方式提取文档结构")

                        except Exception as e:
                            import traceback

                            RuntimeLogger.log(
                                f"[ERROR] extract_word_hierarchy 异常: {e}",
                                level="ERROR",
                            )
                            RuntimeLogger.log(traceback.format_exc(), level="DEBUG")

                            # ===== 最终兜底方案 =====
                            RuntimeLogger.log(
                                f"[WARN] 🛟 最终降级: DocumentProcessor.extract_word_structure(无 use_stable)..."
                            )
                            try:
                                fallback = DocumentProcessor.extract_word_structure(
                                    doc_obj
                                )
                                RuntimeLogger.log(
                                    f"  [OK] 🛟 最终降级成功: {len(fallback) if fallback else 0} 项"
                                )
                                return fallback or []
                            except Exception as e2:
                                RuntimeLogger.log(
                                    f"[ERROR] 所有提取方案均失败: {e2}", level="ERROR"
                                )
                                return []

                    # 并发执行
                    futures = {}
                    if need_tpl_structure:
                        futures["tpl"] = executor.submit(
                            extract_word_hierarchy, template_path, tpl_doc
                        )

                    futures["target"] = executor.submit(
                        extract_word_hierarchy, pair["word"], target_doc
                    )

                    excel_info = future_excel_info.result()  # 已有缓存

                    if "tpl" in futures:
                        template_sections = futures["tpl"].result()

                    target_sections = futures["target"].result()

                    RuntimeLogger.log(
                        f"✅ [Step 0] 预提取完成，共获取 {len(target_sections)} 个章节/内容项"
                    )
                elif word_preprocessing_success:
                    # [FIX] 如果Word预处理成功，直接使用预处理的数据，跳过重复提取
                    RuntimeLogger.log("🔎 [Step 0] 使用预处理的Word数据，跳过重复提取")
                    excel_info = future_excel_info.result()
                    target_sections = (
                        getattr(self, "data_cache", {})
                        .get("word_data", {})
                        .get("items", [])
                    )

                    # [FIX] 即使已预处理目标文档，如果需要模板校验，也必须确保 template_sections 被加载
                    if need_tpl_structure and not template_sections:
                        RuntimeLogger.log(
                            "🔎 [Step 0] 预处理模式：同步加载模板文档结构 (含正文)..."
                        )
                        try:
                            matcher = HierarchicalMatcher()
                            tpl_res = matcher.extract_word_outline_as_hierarchy(
                                template_path, include_content=True
                            )
                            template_sections = tpl_res.get("all_items", [])
                            RuntimeLogger.log(
                                f"✅ [Step 0] 模板结构加载完成: {len(template_sections)} 项"
                            )
                        except Exception as e:
                            RuntimeLogger.log(
                                f"⚠️ [Step 0] 模板结构加载失败: {e}", level="WARN"
                            )

                    RuntimeLogger.log(
                        f"✅ [Step 0] 使用缓存数据: {len(target_sections)} 个章节/内容项"
                    )
                else:
                    RuntimeLogger.log(
                        "🔎 [Step 0] 跳过 Word 结构提取 (未选择任何 Word 校验节点)"
                    )
                    excel_info = future_excel_info.result()

            if not tpl_doc and need_tpl_structure:
                RuntimeLogger.log("⚠️ 模板文件加载失败", level="WARN")
            if not target_doc and need_target_structure:
                RuntimeLogger.log("❌ 目标 Word 加载失败", level="ERROR")
                self.finished.emit({"error": "无法加载 Word 文件"})
                return
            if not target_wb:
                RuntimeLogger.log("❌ 目标 Excel 加载失败", level="ERROR")
                self.finished.emit({"error": "无法加载 Excel 文件"})
                return

            self.progress.emit(15, "正在扫描 Excel 工作表列表...")
            RuntimeLogger.log(
                f"✅ [Step 0] 基础对象并发解析完成 (目标项: {len(target_sections)})"
            )

            if not self._is_running:
                return

            # [NEW] 提交 Step 0 结果
            self.step_result.emit(
                0, {"env_check": {"is_ok": True, "sections": len(target_sections)}}
            )

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
            # [UI 对齐] 立即跃迁至第 2 步起始进度 (30%)
            self.progress.emit(30, "正在准备 Excel 空值扫描...")

            # 5. 节点 2：Excel 空值校验
            step2_start = time.time()
            if raw_info.get("run_empty", True):
                RuntimeLogger.log(
                    f"正在启动 [Step 2] Excel 关键列空值扫描 (Sheet: {check_sheet})..."
                )
                self.progress.emit(31, f"正在扫描 Excel ({check_sheet}) 空值行...")
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
            # [UI 对齐] 立即跃迁至第 3 步起始进度 (33%)
            self.progress.emit(33, "正在准备送审比例计算...")

            # 6. 功能匹配校验 (辅助数据)
            if target_sections and (
                raw_info.get("run_hierarchy") or raw_info.get("run_simple")
            ):
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
            else:
                v_res["func_match"] = {"skipped": True}

            # 7. 节点 3：送审比例校验
            step3_start = time.time()
            if raw_info.get("run_ratio", True):
                RuntimeLogger.log(f"正在进行 [Step 3] 送审比例计算...")
                self.progress.emit(34, "当前正在计算送审功能点与人天比例...")
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
            # [UI 对齐] 立即跃迁至第 4 步起始进度 (36%)
            self.progress.emit(36, "正在准备附加值调整因子提取...")

            # 8. 节点 4：附加值调整因子校验
            step4_start = time.time()
            if raw_info.get("run_factors", True):
                RuntimeLogger.log(f"正在进行 [Step 4] Word 附加值调整因子提取...")
                self.progress.emit(37, "正在扫描文档中的因子表与描述文字...")
                # [NEW] 传入已预提取好的 target_sections 以便进行范围限定扫描
                factors = DocumentProcessor.check_adjustment_factors_in_word(
                    target_doc, target_sections=target_sections
                )
                v_res["factor_check"] = factors
            else:
                v_res["factor_check"] = {"skipped": True}

            if not self._is_running:
                return
            v_res["factor_check"]["duration"] = time.time() - step4_start
            self.step_result.emit(4, {"factor_check": v_res["factor_check"]})
            # [UI 对齐] 立即跃迁至第 5 步起始进度 (39%)
            self.progress.emit(39, "正在启动核心层级匹配引擎...")

            # 9. 层级匹配校验 (Node 5)
            step5_start = time.time()
            run_hierarchy = raw_info.get("run_hierarchy", True)
            fuzzy = raw_info.get("fuzzy", True)
            threshold = raw_info.get("threshold", 0.8)

            hierarchy_mapping = None  # [NEW]

            if run_hierarchy:
                RuntimeLogger.log(f"正在启动 [Step 5] 核心层级匹配校验...")

                def hierarchy_progress_proxy(p, msg):
                    if not self._is_running:
                        return
                    # 映射 39-70 的区间 (耗时最长步骤之一，权重增加)
                    mapped_progress = 39 + int(p * 0.31)
                    self.progress.emit(mapped_progress, f"层级匹配: {msg}")
                    if msg and (
                        "开始" in msg
                        or "完成" in msg
                        or "1/" in msg
                        or "/100" in msg
                        or "[PROCESS]" in msg
                        or "阶段" in msg
                        or "匹配" in msg
                    ):
                        RuntimeLogger.log(f"[Step 5] {msg}")

                h_header_row = raw_info.get("hierarchy_header_row", 0)
                l1_col = raw_info.get("level1_column_index", 1)
                l2_col = raw_info.get("level2_column_index", 2)
                l3_col = raw_info.get("level3_column_index", 3)
                hier_sheet = raw_info.get("hierarchy_sheet")

                hierarchy_res = DocumentProcessor.validate_hierarchy_matching(
                    pair[
                        "word"
                    ],  # [FIX] 传递路径而不是 Document 对象，避免 os.path.basename 失败
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

                # [Optimization] 提取层级映射结果用于辅助功能过程匹配
                try:
                    hierarchy_mapping = {}
                    # 合并精确匹配和模糊匹配的结果
                    matched_items = hierarchy_res.get(
                        "exact_matched", []
                    ) + hierarchy_res.get("fuzzy_matched", [])
                    for item in matched_items:
                        # 构造 Excel 层级 key
                        e_l1 = str(item.get("Excel一级模块", "")).strip()
                        e_l2 = str(item.get("Excel二级模块", "")).strip()
                        e_l3 = str(item.get("Excel三级模块", "")).strip()
                        key = (e_l1, e_l2, e_l3)

                        # 寻找 Word 匹配项 (优先使用三级标题，其次二级，其次一级)
                        w_title = (
                            item.get("Word三级标题")
                            or item.get("Word二级标题")
                            or item.get("Word一级标题")
                        )
                        if w_title:
                            hierarchy_mapping[key] = w_title

                    RuntimeLogger.log(
                        f"已成功提取 {len(hierarchy_mapping)} 个层级映射锚点用于加速后续匹配"
                    )
                except Exception as e:
                    RuntimeLogger.log(f"提取层级映射失败: {e}", level="WARN")

                RuntimeLogger.log(
                    f"层级匹配完成: {hierarchy_res.get('statistics', {})}"
                )
            else:
                v_res["hierarchy_res"] = {"is_valid": True, "skipped": True}

            if not self._is_running:
                return
            v_res["hierarchy_res"]["duration"] = time.time() - step5_start
            self.step_result.emit(5, {"hierarchy_res": v_res["hierarchy_res"]})
            # [UI 对齐] 完成第 5 步后，立即将 UI 文字推进至第 6 步区位 (70%)
            self.progress.emit(70, "正在启动功能过程内容匹配...")

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
                    # 映射 70%-95% 的区间 (耗时最长步骤之一，权重增加)
                    mapped_progress = 70 + int(p * 0.25)
                    self.progress.emit(mapped_progress, f"过程匹配: {msg}")
                    # 扩展日志白名单，确保功能点匹配的各个阶段进度也能记录到日志文件
                    if msg and (
                        "开始" in msg
                        or "完成" in msg
                        or "1/" in msg
                        or "/100" in msg
                        or "[PROCESS]" in msg
                        or "阶段" in msg
                        or "搜索" in msg
                        or "进度" in msg
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
                    hierarchy_mapping=hierarchy_mapping,  # [NEW] 传入已有的映射结果
                    preloaded_word_data=getattr(self, "data_cache", {}).get(
                        "word_data"
                    ),  # [NEW] 预处理数据
                )
                v_res["process_res"] = process_res
            else:
                v_res["process_res"] = {"is_valid": True, "skipped": True}

            if not self._is_running:
                return
            v_res["process_res"]["duration"] = time.time() - step6_start
            self.step_result.emit(6, {"process_res": v_res["process_res"]})
            # [UI 对齐] 完成第 6 步后，立即将 UI 文字推进至第 7 步 (95%)
            self.progress.emit(95, "正在启动数据移动类型校验...")

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
        self.current_step_num = 0  # 从节点 0 开始
        self.is_running = True
        self.current_view_step = 0  # 初始化当前查看的步骤 (用于主题切换时重新应用样式)

        self.setFrameShape(QFrame.StyledPanel)
        self.setProperty("class", "TaskCard")
        self._init_ui()
        self.update_style()

    def update_style(self):
        """全面刷新卡片样式，确保无底色残留且对比度达标"""
        from PySide6.QtWidgets import QApplication
        from extend.matcher_config import MatcherConfig

        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)

        # 获取任务初始色 (作为左侧垂直条颜色)
        left_bar_color = self.task_data.get("border_color", "#3b82f6")

        if is_dark:
            # 深色模式精选色板
            card_bg = "transparent"
            card_border = "#334155"
            text_color = "#94a3b8"
            title_color = "#f8fafc"
            detail_card_bg = "transparent"
            val_text_color = "#f8fafc"
            lbl_text_color = "#94a3b8"
            btn_bg = "rgba(255, 255, 255, 0.05)"
            btn_border = "#334155"
            btn_text = "#f1f5f9"
            sep_color = "#334155"
        else:
            # 浅色模式精选色板 (高对比度)
            card_bg = "#ffffff"
            card_border = "#cbd5e1"
            text_color = "#334155"
            title_color = "#0f172a"
            detail_card_bg = "#f1f5f9"
            val_text_color = "#0f172a"
            lbl_text_color = "#64748b"
            btn_bg = "#ffffff"
            btn_border = "#94a3b8"
            btn_text = "#0f172a"
            sep_color = "#cbd5e1"

        # 核心 QSS：一次性解决所有子组件的底色和边距问题
        self.setStyleSheet(
            f"""
            QFrame[class="TaskCard"] {{
                background: {card_bg};
                border: 1px solid {card_border};
                border-left: 4px solid {left_bar_color};
                border-radius: 16px;
            }}
            QWidget#MainContainer {{
                background: transparent;
            }}
            QLabel {{
                background: transparent;
                color: {text_color};
                font-family: 'Segoe UI', 'Microsoft YaHei UI';
            }}
            QLabel[class="task-title"] {{
                color: {title_color};
                font-size: 21px;
                font-weight: 700;
            }}
            QLabel[class="detail-title"] {{
                color: {title_color};
                font-size: 16px;
                font-weight: 800;
            }}

            /* 解决顽固黑盒：强制 QTextBrowser 透明化 */
            QFrame#DetailCard {{
                background: transparent;
                border: 1px solid {card_border};
                border-radius: 12px;
            }}
            QTextBrowser {{
                background: transparent;
                border: none;
                color: {text_color};
            }}

            /* 按钮统一样式 */
            QPushButton {{
                padding: 6px 16px; border-radius: 8px; font-size: 13px; font-weight: 600;
                border: 1px solid {btn_border}; background-color: {btn_bg}; color: {btn_text};
            }}
            QPushButton:hover {{
                background-color: {"rgba(255, 255, 255, 0.1)" if is_dark else "#f1f5f9"};
            }}
        """
        )

        # 针对特定状态的手动补齐
        if hasattr(self, "stop_btn"):
            stop_color = "#ef4444"
            self.stop_btn.setStyleSheet(
                f"color: {stop_color}; border-color: {stop_color}66;"
            )

        # 刷新所有动态文字颜色
        if hasattr(self, "_meta_labels"):
            for lbl in self._meta_labels:
                lbl.setStyleSheet(
                    f"color: {lbl_text_color}; font-size: 13px; font-weight: 500;"
                )

        if hasattr(self, "_meta_values"):
            for val, orig_color in self._meta_values:
                # 如果是 DYNAMIC，随主题变；如果是固定的（如 FP 的绿色），保持原样
                target_color = val_text_color if orig_color == "DYNAMIC" else orig_color
                val.setStyleSheet(
                    f"color: {target_color}; font-size: 18px; font-weight: 600;"
                )

        if hasattr(self, "time_lbl"):
            self.time_lbl.setStyleSheet(
                f"color: {lbl_text_color}; font-size: 13px; font-weight: 500;"
            )
        if hasattr(self, "time_label"):
            self.time_label.setStyleSheet(
                f"color: {val_text_color}; font-size: 13px; font-weight: 600;"
            )

        if hasattr(self, "_separators"):
            for s in self._separators:
                s.setStyleSheet(f"color: {sep_color}; font-size: 13px; margin: 0 5px;")

        # 刷新徽章
        if hasattr(self, "badge"):
            if is_dark:
                self.badge.setStyleSheet(
                    "background: rgba(16, 185, 129, 0.1); color: #10b981; border-radius: 4px; padding: 2px 8px; font-weight: 700; font-size: 11px; border: 1px solid rgba(16, 185, 129, 0.2);"
                )
            else:
                self.badge.setStyleSheet(
                    "background: #f0fdf4; color: #16a34a; border-radius: 4px; padding: 2px 8px; font-weight: 700; font-size: 11px; border: 1px solid #bbfcce;"
                )

        # 刷新子组件主题
        if hasattr(self, "steps_widget"):
            self.steps_widget.update_theme_style()

        # [FIX] 主题切换后，如果当前有选中的步骤，重新应用其样式
        # 这修复了点击节点后切换夜间/白天模式导致背景色不协调的问题
        if hasattr(self, "current_view_step") and self.current_view_step > 0:
            log_text = self.task_data["logs"].get(self.current_view_step, "")
            if log_text:
                # 重新应用当前步骤的样式 (颜色、背景等)
                self.update_log(self.current_view_step, log_text)

    def _create_metric_card(self, label, value, color="#3b82f6"):
        card = QFrame()
        card.setProperty("class", "MetricCard")
        # 添加轻微阴影
        card.setGraphicsEffect(None)  # 先清理原有的
        layout = QVBoxLayout(card)
        layout.setSpacing(8)
        layout.setAlignment(Qt.AlignCenter)

        lbl = QLabel(label)
        lbl.setProperty("class", "MetricLabel")
        lbl.setAlignment(Qt.AlignCenter)
        val = QLabel(str(value))
        val.setProperty("class", "MetricValue")
        val.setStyleSheet(f"color: {color};")
        val.setAlignment(Qt.AlignCenter)

        layout.addWidget(lbl)
        layout.addWidget(val)
        return card, val

    def _init_ui(self):
        # 恢复垂直布局，简化结构，消除多重边框
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        self.container_widget = QWidget()
        self.container_widget.setObjectName("MainContainer")
        layout = QVBoxLayout(self.container_widget)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 16, 24, 16)

        main_layout.addWidget(self.container_widget)

        # 1. Header (Line 1: Title & Buttons)
        header = self._create_header()
        layout.addWidget(header)

        # 2. Metrics & Time (Line 2: 送审人天： | 送审工作量： | 总耗时：)
        metrics_line = QHBoxLayout()
        metrics_line.setSpacing(15)

        def create_meta_item(label, value, color=None):
            # 颜色逻辑移入 update_style 处理 DYNAMIC 情况
            container = QWidget()
            l = QHBoxLayout(container)
            l.setContentsMargins(0, 0, 0, 0)
            l.setSpacing(6)
            lbl = QLabel(f"{label}:")
            lbl.setProperty("class", "meta-label")
            val = QLabel(str(value))
            val.setProperty("class", "meta-value")
            l.addWidget(lbl)
            l.addWidget(val)
            # 保存引用
            if not hasattr(self, "_meta_labels"):
                self._meta_labels = []
            self._meta_labels.append(lbl)
            if not hasattr(self, "_meta_values"):
                self._meta_values = []
            self._meta_values.append((val, color if color else "DYNAMIC"))
            return container, val

        mandays_box, self.mandays_val = create_meta_item(
            "送审人天", self.task_data.get("days", "0")
        )
        fp_box, self.fp_val = create_meta_item("送审工作量", "0", "#10b981")

        self.time_label = QLabel("00:00")
        time_container = QWidget()
        time_layout = QHBoxLayout(time_container)
        time_layout.setContentsMargins(0, 0, 0, 0)
        time_layout.setSpacing(6)
        self.time_lbl = QLabel("总耗时:")
        self.time_lbl.setProperty("class", "meta-label")
        time_layout.addWidget(self.time_lbl)
        time_layout.addWidget(self.time_label)

        def create_sep():
            s = QLabel("|")
            # 颜色由 update_style 统一控制
            s.setProperty("class", "metric-sep")
            if not hasattr(self, "_separators"):
                self._separators = []
            self._separators.append(s)
            return s

        metrics_line.addWidget(mandays_box)
        metrics_line.addWidget(create_sep())
        metrics_line.addWidget(fp_box)
        metrics_line.addWidget(create_sep())
        metrics_line.addWidget(time_container)
        metrics_line.addStretch()
        layout.addLayout(metrics_line)

        # 3. Status (Line 3: Status Message)
        status_line = QHBoxLayout()
        status_line.setSpacing(8)

        self.status_icon = QLabel("🚀")
        self.status_icon.setStyleSheet("font-size: 14px;")
        self.status_label = QLabel("正在核查...")
        self.status_label.setStyleSheet(
            "color: #3b82f6; font-weight: 600; font-size: 14px;"
        )
        status_line.addWidget(self.status_icon)
        status_line.addWidget(self.status_label)
        status_line.addStretch()
        layout.addLayout(status_line)

        # 3. 步骤进度条
        self.steps_widget = StepsWidget(self.task_data["steps"])
        self.steps_widget.node_clicked.connect(self.on_node_clicked)
        self.steps_widget.label_clicked.connect(self.on_label_clicked)
        # 减小 StepsWidget 的高度
        self.steps_widget.container.setFixedHeight(95)
        layout.addWidget(self.steps_widget)

        # 4. 详情卡片
        self.detail_card = QFrame()
        self.detail_card.setObjectName("DetailCard")
        detail_layout = QVBoxLayout(self.detail_card)
        detail_layout.setContentsMargins(20, 15, 20, 15)
        detail_layout.setSpacing(10)

        detail_header = QHBoxLayout()
        self.step_badge = QLabel("STEP 01")
        self.step_badge.setStyleSheet(
            """
            background: #3b82f6; color: white; border-radius: 4px;
            padding: 2px 8px; font-weight: 800; font-size: 10px;
        """
        )
        self.detail_title = QLabel("核查明细")
        self.detail_title.setProperty("class", "detail-title")
        # 移至 update_style 统一处理文本颜色

        # [NEW] 视图切换按钮 (暂时注释掉切换模式)
        # self.btn_view_text = QPushButton("文字模式")
        # self.btn_view_table = QPushButton("表格模式")
        # for btn in [self.btn_view_text, self.btn_view_table]:
        #     btn.setCheckable(True)
        #     btn.setCursor(Qt.PointingHandCursor)
        #     btn.setStyleSheet("""
        #         QPushButton {
        #             border: 1px solid palette(mid); border-radius: 4px; padding: 2px 10px; font-size: 11px;
        #             background: palette(button); color: palette(text);
        #         }
        #         QPushButton:checked {
        #             background: #3b82f6; color: white; border: none; font-weight: bold;
        #         }
        #     """)

        # self.view_group = QButtonGroup(self)
        # self.view_group.addButton(self.btn_view_text)
        # self.view_group.addButton(self.btn_view_table)
        # self.btn_view_text.setChecked(True)
        # self.view_group.idClicked.connect(self.on_view_toggle)

        # btn_container = QHBoxLayout()
        # btn_container.setSpacing(5)
        # btn_container.addWidget(self.btn_view_text)
        # btn_container.addWidget(self.btn_view_table)

        detail_header.addWidget(self.step_badge)
        detail_header.addWidget(self.detail_title)
        detail_header.addStretch()
        # detail_header.addLayout(btn_container)
        detail_layout.addLayout(detail_header)

        # 视图堆栈
        self.detail_stack = QStackedWidget()

        self.detail_content = QTextBrowser()  # 改用 QTextBrowser 以支持更好的 HTML 渲染
        self.detail_content.setOpenExternalLinks(True)
        # 细节样式统一移至 update_style

        # 表格预览：暂时下线
        # self.detail_table_scroll = QScrollArea()
        # self.detail_table_scroll.setWidgetResizable(True)
        # self.detail_table_scroll.setStyleSheet("background: transparent; border: none;")

        # self.table_container = QWidget()
        # self.table_container.setObjectName("TableContainer")
        # self.table_layout = QVBoxLayout(self.table_container)
        # self.table_layout.setContentsMargins(0, 0, 0, 0)
        # self.table_layout.setSpacing(15)
        # self.detail_table_scroll.setWidget(self.table_container)

        self.detail_stack.addWidget(self.detail_content)
        # self.detail_stack.addWidget(self.detail_table_scroll)
        detail_layout.addWidget(self.detail_stack)

        layout.addWidget(self.detail_card)

        # 计时器
        self.ui_timer = QTimer(self)
        self.ui_timer.timeout.connect(self.on_timer_tick)
        self.ui_timer.start(100)

        # 延迟启动异步校验，让卡片先显示出来（避免卡顿）
        self.worker = None
        if not self.task_data.get("validation_results"):
            # 使用 QTimer 在下一个事件循环中启动 worker
            QTimer.singleShot(50, self.start_validation_worker)
        else:
            self.on_validation_finished(self.task_data["validation_results"][0])

    def start_validation_worker(self):
        """延迟启动校验 worker，让 UI 先显示"""
        if self.worker is None:
            self.worker = ValidationWorker(self.task_data)
            self.worker.progress.connect(self.on_validation_progress)
            self.worker.step_result.connect(self.on_step_finished)
            self.worker.finished.connect(self.on_validation_finished)
            self.worker.start()

    def _create_header(self):
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(15)

        # 1. 标题图标与文字 (对齐图2)
        display_name = clean_project_name(self.task_data["filename"])
        title_icon = QLabel("📄")
        title_icon.setStyleSheet("font-size: 18px;")

        title_text = QLabel(f"{display_name}")
        title_text.setToolTip(self.task_data["filename"])
        title_text.setProperty("class", "task-title")

        # 徽章样式：紧凑、淡绿色 (模拟结算徽章)
        self.badge = QLabel("结算")
        self.badge.setStyleSheet(
            """
            background: rgba(16, 185, 129, 0.1);
            color: #10b981;
            border-radius: 4px;
            padding: 2px 8px;
            font-weight: 700;
            font-size: 11px;
            border: 1px solid rgba(16, 185, 129, 0.2);
        """
        )

        header_layout.addWidget(title_icon)
        header_layout.addWidget(title_text)
        header_layout.addWidget(self.badge)
        header_layout.addStretch()

        # 2. 按钮组 (结果汇总、停止)
        # 初始样式设为空，由 update_style 统一管理颜色
        self.report_btn = QPushButton("结果汇总")
        self.report_btn.setCursor(Qt.PointingHandCursor)
        self.report_btn.clicked.connect(self.show_summary)

        self.stop_btn = QPushButton("停止")
        self.stop_btn.setCursor(Qt.PointingHandCursor)
        self.stop_btn.clicked.connect(self.stop_task)

        header_layout.addWidget(self.report_btn)
        header_layout.addWidget(self.stop_btn)

        return header

    def stop_task(self):
        """停止当前校验任务"""
        if (
            hasattr(self, "worker")
            and self.worker is not None
            and self.worker.isRunning()
        ):
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
            f"⏱️ 预计剩余 {int(rem//60):02d}:{int(rem%60):02d} | 已耗时 {int(elapsed//60):02d}:{int(elapsed%60):02d}"
        )

        # 2. 进度平滑“捕捉”后台真实值
        if self.current_display_progress < self.target_backend_progress:
            # 向目标进度靠近，由原本的由于预跑导致 90% 的策略改为直接追赶
            diff = self.target_backend_progress - self.current_display_progress
            if diff > 15:
                # 较大的差距快速追赶
                self.current_display_progress += 3.0
            elif diff > 5:
                # 中等差距加速
                self.current_display_progress += 1.0
            else:
                # 小差距平滑逼近
                self.current_display_progress += 0.2
        elif self.current_display_progress < 99.8:
            # 极慢速的前进，表示系统活跃
            self.current_display_progress += 0.01

        # 3. 计算当前步骤内的相对进度 (让圆圈填充效果更精确)
        # 更新全局进度条 (集成在 StepsWidget 的线条中)
        self.steps_widget.set_total_progress(self.current_display_progress)

        # [NEW] 线性步进：当前的百分比在步骤区间内的位置
        # 同步 ValidationWorker 中的 emit 节点：0, 15, 30, 33, 36, 39, 70, 95, 100
        ranges = [
            (0, 15),  # Step 0: 环境准备
            (15, 30),  # Step 1: 模板
            (30, 33),  # Step 2: 空值
            (33, 36),  # Step 3: 比例
            (36, 39),  # Step 4: 因子
            (39, 70),  # Step 5: 层级
            (70, 95),  # Step 6: 过程
            (95, 100),  # Step 7: 移动
        ]
        idx = max(0, min(len(ranges) - 1, self.current_step_num))
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

        # 2. 定位当前所处的语义阶段 (同步 ValidationWorker.run 中的 emit 进度点与 UI 气泡)
        # 步骤列表索引：0:环境, 1:模板, 2:空值, 3:比例, 4:因子, 5:层级, 6:过程, 7:报告
        step_names = [
            "环境解析与预加载",
            "模板合规性校验",
            "Excel 空值扫描",
            "送审比例计算",
            "附加值因子提取",
            "核心层级关系校验",
            "简单过程内容匹配",
            "数据移动校验与报告汇总",
        ]

        # 匹配 ValidationWorker.run 中的 emit 点：[0, 15, 30, 33, 36, 39, 70, 95]
        if value < 15:
            current_idx = 0
            self.current_step_num = 0
        elif value < 30:
            current_idx = 1
            self.current_step_num = 1
        elif value < 33:
            current_idx = 2
            self.current_step_num = 2
        elif value < 36:
            current_idx = 3
            self.current_step_num = 3
        elif value < 39:
            current_idx = 4
            self.current_step_num = 4
        elif value < 70:
            current_idx = 5
            self.current_step_num = 5
        elif value < 95:
            current_idx = 6
            self.current_step_num = 6
        else:
            current_idx = 7
            self.current_step_num = 7

        # 3. 更新界面状态文字
        if hasattr(self, "status_label"):
            # [UI FIX] 根据用户反馈，Step 数字与 Node ID 保持一致 (环境准备=0, 功能过程=6)
            disp_step = self.current_step_num
            step_text = step_names[current_idx]

            # [USER UPDATE] 按照用户要求格式化：正在进行第x步 - 正在xxx
            # 强化描述：如果有子步骤文字则展示，否则展示大标题
            display_text = sub_step_text if sub_step_text else step_text

            if "[PROCESS]" in display_text:
                # [Optimization] 针对功能过程匹配的特殊进度格式化
                # 去掉多余的阶段前缀，保留核心进度
                clean_detail = (
                    display_text.replace("[PROCESS]", "")
                    .replace("过程匹配:", "")
                    .strip()
                )
                if self.current_step_num == 6:
                    self.status_label.setText(
                        f"正在进行第 {disp_step} 步:功能过程匹配 {clean_detail}"
                    )
                else:
                    self.status_label.setText(
                        f"正在进行第 {disp_step} 步 - {step_text} {clean_detail}"
                    )
            else:
                self.status_label.setText(
                    f"正在进行第 {disp_step} 步 - {display_text}..."
                )
            self.status_icon.setText("🔄")

            if value >= 100:
                self.status_label.setText("所有校验任务已完成")
                self.status_icon.setText("✅")
            else:
                pass

        # [NEW] 同步更新日志面板与内存日志池
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
            for s in range(0, self.current_step_num):
                curr_status = self.steps_widget.step_nodes[s].status
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
        if step_num == 0:
            env_res = results.get("env_check", {})
            status = "done"
            log = f"✅ 环境准备完成，成功解析文档对象并预提取了 {env_res.get('sections', 0)} 个章节内容。"
            self.steps_widget.set_step_status(0, status)
            self.task_data["logs"][0] = log
            self.update_log(0, log)

        elif step_num == 1:
            if results.get("skipped"):
                status = "skipped"
                log = "⚪ 模板校验已跳过。"
            else:
                suspect_count = len(results.get("suspect_modules", []))
                if results.get("is_valid") and suspect_count == 0:
                    status = "done"
                    log = "✅ 模板校验通过：全层级正文已填充。"
                elif suspect_count > 0:
                    status = "fail"
                    log = f"❌ 发现 {suspect_count} 处正文与模板高度相似，疑似未填写！"
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
                    self.fp_val.setText(str(fp_count))

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
                    report_lines = []

                    # 1. 规模因子行
                    scale = factors.get("scale", {})
                    scale_val = scale.get("value")
                    if scale_val == "结算":
                        report_lines.append(f"✅ 需求变更规模因子: {scale_val}")
                    else:
                        report_lines.append(
                            f"❌ 需求变更规模因子: {scale_val if scale_val else '无'} (一般为结算, 请确认)"
                        )

                    # 更新徽章展示
                    if scale_val in ["结算", "预算"]:
                        self.badge.setText(scale_val)
                        badge_style = {
                            "结算": "background: #ecfdf5; color: #047857; border: 1px solid #6ee7b7;",
                            "预算": "background: #fffbeb; color: #b45309; border: 1px solid #fcd34d;",
                        }
                        self.badge.setStyleSheet(
                            f"{badge_style.get(scale_val)} padding: 2px 8px; border-radius: 4px; font-size: 11px;"
                        )

                    # 2. 质量特性解析
                    quality_keys = [
                        "distributed",
                        "performance",
                        "reliability",
                        "multiple_sites",
                    ]
                    name_map = {
                        "distributed": "分布式处理",
                        "performance": "性能",
                        "reliability": "可靠性",
                        "multiple_sites": "多重站点",
                    }

                    text_missing = []
                    text_ok_names = []
                    table_vals = []
                    table_missing_names = []
                    has_consistency_issue = False

                    for key in quality_keys:
                        f = factors.get(key, {})
                        name = name_map.get(key, key)

                        # 文字描述收集
                        if f.get("found_in_text"):
                            text_ok_names.append(name)
                        else:
                            text_missing.append(name)

                        # 表格描述收集
                        t_val = str(f.get("table_value") or "").strip()
                        if f.get("found_in_table") and t_val not in ["缺失", "-1"]:
                            table_vals.append(
                                f"{name}({t_val if t_val in ['0', '1'] else '有描述'})"
                            )
                        else:
                            table_missing_names.append(name)

                        # 一致性标记
                        if f.get("consistency_warn"):
                            has_consistency_issue = True

                    # 3. 文字行
                    if not text_missing:
                        report_lines.append(
                            f"✅ 质量及特性文字描述: { '、'.join(text_ok_names) }"
                        )
                    else:
                        report_lines.append(
                            f"❌ 质量及特性文字描述异常: { '、'.join(text_missing) } 缺少"
                        )

                    # 4. 表格行
                    if not table_missing_names:
                        report_lines.append(
                            f"✅ 质量及特性表格描述: { '、'.join(table_vals) }"
                        )
                    else:
                        # 汇总显示表格项，对于缺失的明确标注
                        all_t_parts = table_vals + [
                            f"{m}(缺失)" for m in table_missing_names
                        ]
                        report_lines.append(
                            f"❌ 质量及特性表格描述异常: { '、'.join(all_t_parts) }"
                        )

                    # 5. 最终一致性校验
                    # [USER UPDATE] 理由不写吧 -> 简洁化提示
                    if (
                        not has_consistency_issue
                        and not text_missing
                        and not table_missing_names
                    ):
                        report_lines.append(f"✅ 质量及特性描述一致性: 正常")
                    else:
                        report_lines.append(f"❌ 质量及特性一致性校验不匹配")

                    # 6. 设置显示日志和状态
                    log = "\n".join(report_lines)

                    # 判定整体状态：任何红叉出现即为 fail
                    status = (
                        "fail"
                        if any(line.startswith("❌") for line in report_lines)
                        else "done"
                    )

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
                        status = "fail"
                        log = "❌ 层级匹配未执行：Excel 中未找到有效的三级模块数据。"
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

                        def _group_by_prefix(desc_list):
                            groups = {}
                            order = 0
                            for raw in desc_list:
                                if not raw or raw == "-":
                                    continue
                                desc = " ".join(str(raw).split())
                                parts = desc.split("：", 1)
                                prefix = parts[0].strip()
                                body = parts[1].strip() if len(parts) > 1 else ""

                                if prefix not in groups:
                                    groups[prefix] = {"order": order, "bodies": []}
                                    order += 1
                                if body and body not in groups[prefix]["bodies"]:
                                    groups[prefix]["bodies"].append(body)
                                elif not body and body not in groups[prefix]["bodies"]:
                                    groups[prefix]["bodies"].append(body)
                            return groups

                        # 1) 缺失：按类型前缀分组（最多展示 3 种）
                        miss_descs = []
                        for item in all_failed_items:
                            d = item.get("缺失简略描述")
                            if not d or d == "-":
                                d_gen = item.get("简略描述")
                                if d_gen and "未体现" in str(d_gen):
                                    d = d_gen
                            if d and d != "-":
                                miss_descs.append(d)

                        miss_groups = _group_by_prefix(miss_descs)
                        if miss_groups:
                            lines = []
                            for prefix, info in sorted(
                                miss_groups.items(), key=lambda kv: kv[1]["order"]
                            ):
                                bodies = info["bodies"]
                                body0 = bodies[0] if bodies else ""
                                count = len([b for b in bodies if b]) or 1
                                if body0:
                                    lines.append(
                                        f"{prefix}：{body0}（已合并{count}条）"
                                    )
                                else:
                                    lines.append(f"{prefix}（已合并{count}条）")
                            issue_details.append(
                                f"• [缺失项] (Excel在Word未体现): \n  - "
                                + "\n  - ".join(lines[:3])
                            )

                        # 2) 层级不匹配：按类型前缀分组（最多展示 7 种）
                        mismatch_descs = []
                        for item in all_failed_items:
                            d = item.get("层级不匹配简略描述")
                            if not d or d == "-":
                                d_gen = item.get("简略描述", "")
                                if d_gen and "不匹配" in str(d_gen):
                                    d = d_gen
                            if d and d != "-":
                                mismatch_descs.append(d)

                        mismatch_groups = _group_by_prefix(mismatch_descs)
                        if mismatch_groups:
                            lines = []
                            for prefix, info in sorted(
                                mismatch_groups.items(), key=lambda kv: kv[1]["order"]
                            ):
                                bodies = info["bodies"]
                                body0 = bodies[0] if bodies else ""
                                count = len([b for b in bodies if b]) or 1
                                if body0:
                                    lines.append(
                                        f"{prefix}：{body0}（已合并{count}条）"
                                    )
                                else:
                                    lines.append(f"{prefix}（已合并{count}条）")
                            issue_details.append(
                                f"• [层级不匹配] (对应关系错误): \n  - "
                                + "\n  - ".join(lines[:7])
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
                        status = "fail"
                        not_found = process_res.get("not_found_in_word", [])
                        names = [
                            f"【{item.get('Excel功能点', '未知')}】"
                            for item in not_found[:2]
                        ]
                        names_str = "、".join(names)
                        suffix = "等" if len(not_found) > 2 else ""
                        log = f"❌功能过程校验不通过（匹配率：{match_rate}）\n{names_str}{suffix}功能过程在需求规格书未体现"

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
                        status = "fail"
                        log = "❌ 未发现有效的功能过程数据移动类型数据。"
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

        # 实时刷新整体异常状态指示 (解决用户提到的“图中圈起来的地方变红”)
        has_any_fail = False
        for node in getattr(self.steps_widget, "step_nodes", []):
            if node.status in ["fail", "error", "warn"]:
                has_any_fail = True
                break

        if has_any_fail:
            # 文字同步变红
            if hasattr(self, "status_label"):
                self.status_label.setStyleSheet(
                    "font-weight: 800; font-size: 14px; color: #ef4444;"
                )
        else:
            # 正常状态：文字蓝色 (或保持原始)
            if hasattr(self, "status_label"):
                self.status_label.setStyleSheet(
                    "font-weight: 600; font-size: 14px; color: #3b82f6;"
                )

    def on_validation_finished(self, results):
        """校验完全结束"""
        self.is_running = False
        self.ui_timer.stop()
        if hasattr(self, "stop_btn"):
            self.stop_btn.setEnabled(False)
            self.stop_btn.setText("已完成")

        # 1. 强行拉满所有视觉进度
        self.current_display_progress = 100
        if hasattr(self, "steps_widget"):
            # 填满所有连接线
            self.steps_widget.set_total_progress(100)
            # 清除所有旋转动画，转为静态图标
            self.steps_widget.clear_all_progress()

            # 确保所有之前的步骤如果是 pending/processing，都标记为已完成
            for i in range(1, 8):
                node = self.steps_widget.step_nodes[i - 1]
                if node.status in ["pending", "processing"]:
                    self.steps_widget.set_step_status(i, "finished")

        # 2. 状态文字更新 (简化颜色：仅红/绿)
        # 不再依赖字符串搜索，而是直接检查步骤节点状态
        has_issue = False
        if hasattr(self, "steps_widget"):
            for node in self.steps_widget.step_nodes:
                if node.status in ["fail", "error", "warn"]:
                    has_issue = True
                    break

        status_text = "所有任务校验完成"
        if has_issue:
            status_text += " (存在异常)"

        self.status_label.setText(status_text)
        self.status_icon.setText("❌" if has_issue else "✅")
        # 只要有异常就显示红色 #ef4444，否则显示绿色 #10b981
        self.status_label.setStyleSheet(
            f"font-weight: 800; font-size: 14px; color: {'#ef4444' if has_issue else '#10b981'};"
        )

        elapsed = time.time() - self.start_time
        self.time_label.setText(f"{int(elapsed//60):02d}:{int(elapsed%60):02d}")

        if not results:
            return

        # 3. 结果合并与展示
        if not self.task_data.get("validation_results"):
            self.task_data["validation_results"] = [results]
        else:
            self.task_data["validation_results"][0].update(results)

        # 展示逻辑：优先展示有问题的步骤
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
        ):
            final_show_step = 5
        elif res_dict.get("process_res", {}).get("statistics", {}).get("缺失项", 0) > 0:
            final_show_step = 6

        # 触发最终详情页更新
        for i in range(1, 8):
            if i not in self.task_data["logs"]:
                self.task_data["logs"][i] = "✅ 校验通过，未发现异常。"

        final_log = self.task_data["logs"].get(final_show_step, "")
        self.update_log(final_show_step, final_log)

        # 自动化：完成后自动打开文件夹
        config = MatcherConfig.load()
        if config.get("automation", {}).get("auto_open", True):
            report_path = res_dict.get("auto_report_path")
            if report_path:
                open_directory(os.path.dirname(report_path))
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

    def on_view_toggle(self, btn_id):
        """切换文字/表格视图 (已暂时注释)"""
        pass
        # if self.btn_view_text.isChecked():
        #     self.detail_stack.setCurrentIndex(0)
        # else:
        #     self.detail_stack.setCurrentIndex(1)
        #     # 切换时如果当前步骤有数据，触发重绘表格
        #     step_num = getattr(self, "current_view_step", 1)
        #     self._update_table_data(step_num)

    def _update_table_data(self, step_num):
        """根据当前步骤结果渲染高仿原型图的表格布局 (已暂时注释)"""
        pass
        # # 1. 清理旧组件
        # while self.table_layout.count():
        #     item = self.table_layout.takeAt(0)
        #     if item.widget():
        #         item.widget().deleteLater()
        # ... (rest of the code below should also be commented if I could, but I'll just pass)

        # 获取数据
        results = {}
        if self.task_data.get("validation_results"):
            results = self.task_data["validation_results"][0]

        from extend.matcher_config import MatcherConfig

        is_dark = MatcherConfig.load().get("theme", {}).get("is_dark", False)

        # 样式辅助
        header_qss = f"background: {'#21262d' if is_dark else '#e2e8f0'}; color: {'#c9d1d9' if is_dark else '#1e293b'}; border: 1px solid {'#30363d' if is_dark else '#cbd5e1'}; padding: 8px; font-weight: bold; font-size: 13px;"
        cell_qss = f"border: 1px solid {'#30363d' if is_dark else '#cbd5e1'}; padding: 8px; font-size: 13px; color: {'#8b949e' if is_dark else '#475569'};"
        summary_qss = f"background: {'#0d1117' if is_dark else '#f1f5f9'}; border-radius: 4px; border: 1px solid {'#30363d' if is_dark else '#e2e8f0'}; padding: 10px;"

        def create_mock_table(headers, data_rows, stretch_cols=None):
            w = QTableWidget(len(data_rows), len(headers))
            w.setHorizontalHeaderLabels(headers)
            w.verticalHeader().setVisible(False)
            w.setShowGrid(True)
            w.setEditTriggers(QTableWidget.NoEditTriggers)
            w.setFocusPolicy(Qt.NoFocus)

            # 简约样式
            table_style = f"""
                QTableWidget {{
                    background: transparent;
                    gridline-color: {'#30363d' if is_dark else '#cbd5e1'};
                    border: 1px solid {'#30363d' if is_dark else '#cbd5e1'};
                    color: {'#c9d1d9' if is_dark else '#1e293b'};
                }}
                QHeaderView::section {{
                    background: {'#21262d' if is_dark else '#e2e8f0'};
                    color: {'#c9d1d9' if is_dark else '#1e293b'};
                    padding: 8px;
                    border: 1px solid {'#30363d' if is_dark else '#cbd5e1'};
                    font-weight: bold;
                }}
            """
            w.setStyleSheet(table_style)

            for r, row_data in enumerate(data_rows):
                for c, val in enumerate(row_data):
                    it = QTableWidgetItem(str(val))
                    if "❌" in str(val) or "缺失" in str(val) or "不匹配" in str(val):
                        it.setForeground(QColor("#ef4444"))
                    elif "✅" in str(val) or "正常" in str(val):
                        it.setForeground(QColor("#10b981"))
                    w.setItem(r, c, it)

            w.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            w.setFixedHeight(min(400, 45 + 35 * len(data_rows)))
            return w

        if step_num == 3:  # 送审比例 (高仿 2x2 结构，并去掉边框)
            res = results.get("ratio_check", {})
            table_widget = QWidget()
            g_layout = QGridLayout(table_widget)
            g_layout.setSpacing(10)  # 保持间距但去掉物理边框
            g_layout.setContentsMargins(0, 5, 0, 5)

            items = [
                ("送审功能点", str(res.get("fp_count", 0)), 0, 0),
                ("送审人天", str(res.get("mandays", 0)), 0, 1),
                (
                    "送审比例",
                    f"<span style='color: {'#ef4444' if not res.get('is_ok') else '#10b981'}; font-size: 18px; font-weight: bold;'>{res.get('ratio', 0)}</span>",
                    1,
                    0,
                ),
                (
                    "校验结论",
                    f"<span style='color: {'#ef4444' if not res.get('is_ok') else '#10b981'}; font-weight: bold;'>{'❌ 送审功能点过多/少' if not res.get('is_ok') else '✅ 比例正常'}</span>",
                    1,
                    1,
                ),
            ]

            for title, val, r, c in items:
                container = QFrame()
                # 核心要求：彻底去掉所有外边框和内边框，使文字模式具象化为表格排版
                container.setStyleSheet("border: none; background: transparent;")
                vbox = QVBoxLayout(container)
                vbox.setContentsMargins(0, 5, 0, 5)

                t_lbl = QLabel(title)
                t_lbl.setStyleSheet(
                    f"color: {'#8b949e' if is_dark else '#64748b'}; font-size: 13px; font-weight: bold;"
                )

                v_lbl = QLabel(val)
                v_lbl.setTextFormat(Qt.RichText)
                v_lbl.setStyleSheet(
                    f"color: {'#c9d1d9' if is_dark else '#1e293b'}; font-size: 16px; font-weight: 800; margin-top: 2px;"
                )

                vbox.addWidget(t_lbl)
                vbox.addWidget(v_lbl)
                g_layout.addWidget(container, r, c)

            self.table_layout.addWidget(table_widget)

        elif step_num == 4:  # 附加值因子
            res = results.get("factors", {})
            # 1. 规模因子 Header
            header = QLabel("需求变更规模因子")
            header.setStyleSheet(header_qss)
            self.table_layout.addWidget(header)

            content = QLabel(
                f"预算 / 结算 / <span style='color: #ef4444;'>{res.get('scale_val', '缺少')}</span> / 未识别"
            )
            content.setTextFormat(Qt.RichText)
            content.setStyleSheet(cell_qss + "border-top: none;")
            self.table_layout.addWidget(content)

            # 2. 质量特征 Table
            rows = []
            factor_errors = res.get("errors", [])
            for err in factor_errors:
                rows.append(
                    [
                        "文字描述缺少",
                        "总结描述表格缺少",
                        "否" if "不一致" in err else "是",
                    ]
                )

            if not rows:
                rows = [["正常", "正常", "是"]]
            self.table_layout.addWidget(
                create_mock_table(["质量及特征因子", "总结描述表格", "是否一致"], rows)
            )

            footer = QLabel(
                f"<span style='color: #ef4444; font-weight: bold;'>结论：{'文字描述集合和总结表格不一致' if factor_errors else '各因子校验一致'}</span>"
            )
            footer.setTextFormat(Qt.RichText)
            self.table_layout.addWidget(footer)

        elif step_num == 5:  # 层级匹配
            res = results.get("hierarchy_res", {})
            stats = res.get("statistics", {})

            summary = QHBoxLayout()
            l1 = QLabel(
                f"通过率: <span style='color: #ef4444;'>{stats.get('匹配率', '0%')}</span>"
            )
            l2 = QLabel(
                f"问题项: <span style='color: #ef4444;'>{stats.get('缺失项', 0) + stats.get('层级不匹配', 0)}个</span>"
            )
            for l in [l1, l2]:
                l.setTextFormat(Qt.RichText)
                l.setStyleSheet(summary_qss)
                summary.addWidget(l)
            self.table_layout.addLayout(summary)

            # 不匹配层级 Table
            mismatched = []
            for item in res.get("hierarchy_mismatched", [])[:10]:
                mismatched.append(
                    [item.get("excel_path", "-"), item.get("word_path", "-"), "不匹配"]
                )
            if mismatched:
                self.table_layout.addWidget(
                    create_mock_table(["拆分表", "规格书", "不匹配层级"], mismatched)
                )

            # 缺少项 Table
            missing = []
            for item in res.get("not_found_in_word", [])[:10]:
                missing.append([item.get("Excel功能点", "-"), "规格书中未体现"])
            if missing:
                self.table_layout.addWidget(
                    create_mock_table(["拆分表", "缺少项"], missing)
                )

        elif step_num == 6:  # 功能过程
            res = results.get("process_res", {})
            stats = res.get("statistics", {})

            summary = QHBoxLayout()
            l1 = QLabel(
                f"通过率: <span style='color: #f59e0b;'>{stats.get('匹配率', '0%')}</span>"
            )
            l2 = QLabel(
                f"问题项: <span style='color: #f59e0b;'>{stats.get('缺失项', 0)}个</span>"
            )
            for l in [l1, l2]:
                l.setTextFormat(Qt.RichText)
                l.setStyleSheet(summary_qss)
                summary.addWidget(l)
            self.table_layout.addLayout(summary)

            rows = []
            for item in res.get("not_found_in_word", [])[:10]:
                rows.append([f"【{item.get('Excel功能点', '未知')}】", "缺失"])

            if rows:
                self.table_layout.addWidget(
                    create_mock_table(["功能过程", "状态"], rows)
                )

            footer = QLabel("建议：功能过程应逐一核对并与规格书逐字匹配。")
            footer.setStyleSheet("color: #d97706; font-style: italic; font-size: 12px;")
            self.table_layout.addWidget(footer)

        elif step_num == 7:  # 数据移动
            res = results.get("move_res", {})
            stats = res.get("statistics", {})

            summary = QHBoxLayout()
            l1 = QLabel(
                f"通过率: <span style='color: #ef4444;'>{int(stats.get('合规', 0)/max(1, stats.get('总数', 1))*100)}%</span>"
            )
            l2 = QLabel(
                f"问题项: <span style='color: #ef4444;'>{stats.get('不合规', 0)}处不合规</span>"
            )
            for l in [l1, l2]:
                l.setTextFormat(Qt.RichText)
                l.setStyleSheet(summary_qss)
                summary.addWidget(l)
            self.table_layout.addLayout(summary)

            # 模拟原型图中的三个并列方块 (针对主要错误类型)
            grid = QGridLayout()
            grid.setSpacing(10)

            # 逻辑简化：根据结果汇总
            types = {"缺少 X": [], "缺少 e": [], "缺少 w": []}
            for item in res.get("items", []):
                err = item.get("result", "")
                if "缺少 X" in err:
                    types["缺少 X"].append(item.get("process", ""))
                elif "缺少 e" in err:
                    types["缺少 e"].append(item.get("process", ""))
                elif "缺少 w" in err:
                    types["缺少 w"].append(item.get("process", ""))

            titles = [
                ("不合规处", "缺少 X", 0, 0),
                ("缺少 e", "", 0, 1),
                ("缺少 w", "", 1, 0),
            ]
            for main_t, sub_t, r, c in titles:
                box = QFrame()
                box_bg = (
                    "rgba(239, 68, 68, 0.1)"
                    if "缺少 X" in main_t or sub_t == "缺少 X"
                    else "rgba(148, 163, 184, 0.05)"
                )
                box.setStyleSheet(
                    f"background: {box_bg}; border: 1px solid {'#30363d' if is_dark else '#cbd5e1'}; border-radius: 4px; padding: 10px;"
                )
                bl = QVBoxLayout(box)
                tl = QLabel(main_t)
                tl.setStyleSheet(
                    f"font-weight: bold; font-size: 12px; color: {'#c9d1d9' if is_dark else '#1e293b'};"
                )
                bl.addWidget(tl)

                content_str = "、".join(types.get(sub_t or main_t, [])[:2]) or "暂无"
                cl = (
                    QLabel(
                        f"<span style='color: #ef4444;'>{sub_t}</span><br/>{content_str} 等"
                    )
                    if sub_t
                    else QLabel(content_str)
                )
                cl.setTextFormat(Qt.RichText)
                cl.setStyleSheet(
                    f"font-size: 11px; color: {'#8b949e' if is_dark else '#475569'};"
                )
                bl.addWidget(cl)
                grid.addWidget(box, r, c)

            self.table_layout.addLayout(grid)

            footer = QLabel("标准：一个完整功能过程应以 E 开始，并以 W 或 X 结束。")
            footer.setStyleSheet(
                f"color: {'#8b949e' if is_dark else '#94a3b8'}; font-size: 11px; margin-top: 10px;"
            )
            self.table_layout.addWidget(footer)

        else:  # 默认表格
            self.table_layout.addWidget(
                create_mock_table(
                    ["详情", "内容"], [["该步骤暂无表格视图", "请参考文字报告"]]
                )
            )

        self.table_layout.addStretch()

    def update_log(self, step_num, text):
        self.current_view_step = step_num

        # 统一颜色方案 (简化红色/绿色)
        color_map = {
            "processing": "#3b82f6",  # Blue
            "success": "#10b981",  # Green
            "error": "#ef4444",  # Red
        }

        # [FIX] 颜色优先级判定：只要含有错误标识，优先显示红色
        if "❌" in text or "⚠️" in text:
            current_color = color_map["error"]
        elif "✅" in text:
            current_color = color_map["success"]
        else:
            current_color = color_map["processing"]

        from extend.matcher_config import MatcherConfig

        is_dark = MatcherConfig.load().get("theme", {}).get("is_dark", False)

        # 1. 更新卡片侧边条颜色 (优化 Light Mode 边框)
        border_col = "#334155" if is_dark else "#e2e8f0"

        self.detail_card.setStyleSheet(
            f"""
            QFrame#DetailCard {{
                background: transparent;
                border: 1px solid {border_col};
                border-left: 4px solid {current_color};
                border-radius: 10px;
            }}
        """
        )

        # 2. 更新徽章
        # [UI FIX] 徽标与节点 0 对应，显示 STEP 00
        self.step_badge.setText(f"STEP {step_num:02d}")
        self.step_badge.setStyleSheet(
            f"""
            background: {current_color}; color: white; border-radius: 4px;
            padding: 2px 8px; font-weight: 800; font-size: 10px;
        """
        )

        # 3. 更新标题
        # [UI FIX] 修复 0 索引偏移，使得 STEP 00 对应 "环境准备"
        step_name = "详情"
        if 0 <= step_num < len(self.task_data["steps"]):
            step_name = self.task_data["steps"][step_num][0]

        self.detail_title.setText(f"{step_name}校验报告")
        self.detail_title.setStyleSheet(
            f"font-weight: 800; font-size: 16px; color: {current_color};"
        )

        # 4. 格式化正文 (优化 Light Mode 字体对比度)
        formatted_text = text.replace("\n", "<br/>")
        text_col = "#cbd5e1" if is_dark else "#334155"
        style = f"color: {text_col}; background: transparent; font-family: 'Segoe UI', 'Microsoft YaHei UI'; font-size: 14px; line-height: 1.6;"
        self.detail_content.setHtml(f"<div style='{style}'>{formatted_text}</div>")

        # 5. 更新表格数据 (如果当前在表格视图)
        if self.detail_stack.currentIndex() == 1:
            self._update_table_data(step_num)

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
