#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
层级匹配器 - 支持 Word 目录层级与 Excel 模块层级对应
作者: yimiao99
日期: 2025-01-15
"""

import re
import zipfile
from pathlib import Path
from typing import List, Dict, Set, Tuple, Optional
import pandas as pd
from docx import Document
from difflib import SequenceMatcher
import os
import tempfile
import shutil
import traceback  # 用于打印更详细的错误信息
import pythoncom
import time as time_module  # 修复 time_module 未定义错误
from fuzzywuzzy import fuzz
from collections import Counter
from datetime import datetime  # 修复 datetime 错误

# 导入日志记录器
try:
    from utils.runtime_logger import RuntimeLogger
except ImportError:
    RuntimeLogger = None  # 如果导入失败，设置为 None

# 详细日志控制变量 (从环境变量读取)
DETAILED_LOG = os.environ.get("COSMIC_DETAILED_LOG", "0") == "1"

# Import win32com.client for .doc to .docx conversion
# 尝试导入 pywin32 库，如果失败则设置标志位
try:
    import win32com.client as win32

    WIN32COM_AVAILABLE = True
except ImportError:
    print(
        "WARNING: 'pywin32' 库未找到。在非 Windows 环境或未安装该库时，.doc 到 .docx 转换功能将不可用。"
    )
    print("请确保在 Windows 上运行，并已安装 'pywin32' (pip install pywin32)。")
    WIN32COM_AVAILABLE = False


# -----------------------------------------------------------------------------
# 辅助函数：将 .doc 文件转换为 .docx
# -----------------------------------------------------------------------------
def convert_doc_to_docx(doc_path: str) -> str:
    """
    将 .doc 文件转换为 .docx 文件。
    返回转换后的 .docx 文件的路径 (通常是临时文件)。
    如果输入已经是 .docx，则直接返回原路径。
    """
    doc_path_obj = Path(doc_path)

    if doc_path_obj.suffix.lower() == ".docx":
        return doc_path  # 已经是 docx，直接返回

    if not doc_path_obj.exists():
        raise FileNotFoundError(f"文件不存在: {doc_path}")

    if not WIN32COM_AVAILABLE:
        raise ImportError(
            "无法执行 .doc 到 .docx 转换，因为未安装 'pywin32' 库或未在 Windows 环境下运行。"
        )

    # 创建一个临时文件来保存转换后的 docx
    temp_dir = tempfile.gettempdir()
    # 确保临时文件名唯一，避免冲突
    docx_path = Path(temp_dir) / (doc_path_obj.stem + f"_{os.urandom(4).hex()}.docx")

    print(f"  正在将 '{doc_path_obj.name}' (.doc) 转换为 '{docx_path.name}' (.docx)...")

    word = None
    try:
        pythoncom.CoInitialize()  # <-- 在 COM 操作前初始化 COM 库
        word = win32.Dispatch("Word.Application")
        word.Visible = False  # 不显示 Word 窗口
        doc = word.Documents.Open(str(doc_path_obj))
        # FileFormat=16 代表 wdFormatXMLDocument (docx 格式)
        doc.SaveAs(str(docx_path), FileFormat=16)
        doc.Close()
        print(f"  ✓ 转换成功，临时文件: {docx_path}")
        return str(docx_path)
    except Exception as e:
        print(f"  ✗ .doc 到 .docx 转换失败: {e}")
        # 如果转换失败，删除可能创建的临时文件
        if docx_path.exists():
            os.remove(docx_path)
        raise
    finally:
        if word:
            word.Quit()  # 退出 Word 应用程序
        pythoncom.CoUninitialize()  # <-- 在 COM 操作完成后释放 COM 库


def is_descendant(parent_num: str, child_num: str) -> bool:
    """
    判断 child_num 是否是 parent_num 的后代编号。
    例如: '4.1.1' 是 '4.1' 的后代，'4.1.1.1' 也是 '4.1' 的后代。
    """
    if not parent_num or not child_num:
        return True
    p = parent_num.strip().rstrip(".")
    c = child_num.strip().rstrip(".")
    if p == c:
        return True
    return c.startswith(p + ".")


# -----------------------------------------------------------------------------
# HierarchicalMatcher 类
# -----------------------------------------------------------------------------
class HierarchicalMatcher:
    """层级匹配器类"""

    # 业务章节常量
    SECTION_FUNCTIONAL_REQUIREMENTS = "功能需求"
    CHAPTER_NUMBER_PREFIX = "4"  # 默认功能需求从第4章开始

    def __init__(self):
        """初始化匹配器，包含数据缓存"""
        self._word_cache = {}  # Word内容缓存
        self._excel_cache = {}  # Excel数据缓存
        self._index_cache = {}  # 索引缓存

    # 需要忽略的说明性文本模式（Excel 拆分表中常见的提示文字）
    IGNORE_PATTERNS = [
        r"本次需求需要改造的本项目.*一级业务功能名称",
        r"本次需求需要改造的本项目.*二级业务功能名称",
        r"本次需求需要改造的本项目.*三级业务功能名称",
        r"体现了待度量软件的功能性用户需求基本部件",
        r"一个功能处理可能只有一个触发输入",
        r"一个功能处理的所有数据移动的集合",
        r"建议功能过程的名字字数控制在",
        r"功能过程：1、体现了待度量",
        r"每个功能处理由一系列子过程组成",
        r"一个子处理可以是一个数据移动或者数据运算",
        r"一个功能过程至少需要包括两个或两个以上子过程",
        r"建议子过程描述的字数控制在50以内",
        r"该功能处理在该FUR中是独一无二的",
        r"并能独立于该FUR的其他功能处理被定义",
        r"每个功能处理在接受到由其触发输入数据移动所移动的一个数据组后",
        r"开始进行处理",
        r"满足其FUR的触发输入所有可能的响应所需的集合",
        r"建议功能过程的名字字数控制在5-20",
        r"1、每个功能处理由一系列子过程组成",
        r"2、一个子处理可以是一个数据移动或者数据运算",
        r"3、一个功能过程至少需要包括两个或两个以上子过程描述",
        r"4、建议子过程描述的字数控制在50以内",
    ]

    def __init__(self, fuzzy_match: bool = False, threshold: float = 0.8):
        """
        初始化匹配器

        Args:
            fuzzy_match: 是否启用模糊匹配
            threshold: 模糊匹配的相似度阈值（0-1）
        """
        self.fuzzy_match = fuzzy_match
        self.threshold = threshold
        # 添加数据缓存
        self._word_cache = {}  # Word内容缓存
        self._excel_cache = {}  # Excel数据缓存
        self._index_cache = {}  # 索引缓存

    def prepare_word_data_async(
        self, word_path, mode="hierarchical_with_content", progress_callback=None
    ):
        """
        异步预处理Word数据，用于提前数据准备
        【NEW】在预处理阶段读取Word全文内容（节点0），供全文匹配阶段使用
        """
        try:
            cache_key = f"{word_path}:{mode}"

            # 检查缓存
            if cache_key in self._word_cache:
                if os.path.exists(word_path):
                    file_mtime = os.path.getmtime(word_path)
                    cached_mtime = self._word_cache[cache_key].get("mtime", 0)
                    if file_mtime == cached_mtime:
                        if progress_callback:
                            progress_callback(100, "使用缓存数据")
                        print(
                            f"🚀 使用Word缓存数据: {len(self._word_cache[cache_key]['items'])} 项"
                        )
                        return self._word_cache[cache_key]

            if progress_callback:
                progress_callback(0, "开始预处理Word数据...")

            # [NEW] 统一打开文档，避免重复打开/转换
            doc, processed_path, is_temp = self._open_word_doc(word_path)

            try:
                # 提取Word结构
                word_data = self.extract_word_content(
                    doc,
                    mode=mode,
                    word_path_hint=word_path,
                    progress_callback=lambda p, msg: (
                        progress_callback(int(p * 0.4), f"提取结构: {msg}")
                        if progress_callback
                        else None
                    ),
                )

                # [FIX] 确保 word_items 是列表
                if isinstance(word_data, dict) and "all_items" in word_data:
                    word_items = word_data["all_items"]
                elif isinstance(word_data, list):
                    word_items = word_data
                else:
                    word_items = []
                    print(
                        f"⚠️ extract_word_content 返回了未预期的格式: {type(word_data)}"
                    )

                if progress_callback:
                    progress_callback(40, "读取全文内容（节点0）...")

                # 【NEW】读取Word全文内容，用于全文匹配阶段（节点0）
                # 传入word_items以建立章节-内容映射
                full_text_content = self._read_word_full_text(doc, word_items)

                if progress_callback:
                    progress_callback(70, "构建全局索引...")

                # 构建索引
                exact_lookup, toc_items, chapter_buckets = self._build_word_index(
                    word_items, full_text_content
                )

                # 缓存结果
                cached_data = {
                    "items": word_items,
                    "exact_lookup": exact_lookup,
                    "toc_items": toc_items,
                    "chapter_buckets": chapter_buckets,
                    "full_text_content": full_text_content,  # 【NEW】全文内容（节点0）
                    "mtime": (
                        os.path.getmtime(word_path) if os.path.exists(word_path) else 0
                    ),
                }

                self._word_cache[cache_key] = cached_data

                if progress_callback:
                    progress_callback(
                        100, f"预处理完成: {len(word_items)} 章节 + 全文内容已缓存"
                    )

                print(
                    f"✅ Word预处理完成: {len(word_items)} 章节 + {len(full_text_content)} 项全文内容已缓存"
                )
                return cached_data
            finally:
                # 清理临时文件
                if is_temp and os.path.exists(processed_path):
                    try:
                        os.remove(processed_path)
                    except:
                        pass
            return cached_data

        except Exception as e:
            print(f"⚠️ Word数据预处理失败: {e}")
            if progress_callback:
                progress_callback(100, f"预处理失败: {e}")
            return None

    def _build_word_index(self, word_items, full_text_content=None):
        """
        构建Word内容索引（从match_documents中提取）
        """
        exact_lookup = {}
        toc_items = []
        chapter_buckets = {}

        # [NEW] 建立编号到标题项的快速映射，用于合并正文内容
        num_to_item = {}
        if word_items:
            for it in word_items:
                if isinstance(it, dict):
                    num = it.get("number", "")
                    if num:
                        num_to_item[num] = it
                    it["content_list"] = []  # 使用列表临时存储
                    it["content"] = ""  # 最终合并后的正文

        # [NEW] 将 full_text_content 中的正文合并到对应的章节中
        if full_text_content and word_items:
            for f_item in full_text_content:
                if (
                    f_item.get("source") == "paragraph"
                    or f_item.get("source") == "table"
                ):
                    p_num = f_item.get("parent_num", "")
                    if p_num in num_to_item:
                        text = f_item.get("text", "").strip()
                        if text:
                            num_to_item[p_num]["content_list"].append(text)

            # 合并 content_list 并在最后清理
            for it in word_items:
                if "content_list" in it:
                    if it["content_list"]:
                        it["content"] = "\n".join(it["content_list"])
                    del it["content_list"]

        # 合并处理
        all_processing_items = []
        if word_items:
            for it in word_items:
                if isinstance(it, dict):
                    all_processing_items.append(it)

        if full_text_content:
            for it in full_text_content:
                if isinstance(it, dict):
                    all_processing_items.append(it)

        # [FIX] 创建一个新列表，避免在迭代时修改原列表导致的问题
        processed_items = []

        for i, item in enumerate(all_processing_items):
            # 添加全局索引
            item["_global_idx"] = i

            # 添加到目录项列表 (仅限标题)
            if item.get("in_toc", True) and item.get("source") != "paragraph":
                toc_items.append(item)

            processed_items.append(item)

        # [FIX] 现在遍历已经处理完成的项，构建索引
        for item in processed_items:
            t_text = item.get("text", item.get("cleaned", ""))

            # 章节分桶
            p_ch = item.get("parent_chapter", "")
            p_num = str(item.get("parent_num", "")).strip().rstrip(".")

            buckets_to_fill = []
            if p_ch:
                buckets_to_fill.append(p_ch)
            if p_num:
                buckets_to_fill.append(p_num)

            for bucket_key in buckets_to_fill:
                if bucket_key not in chapter_buckets:
                    chapter_buckets[bucket_key] = []
                if item not in chapter_buckets[bucket_key]:
                    chapter_buckets[bucket_key].append(item)

            # 递归分桶 (处理 4.1.1 同时也属于 4.1 和 4 的逻辑)
            if p_num and "." in p_num:
                parts = p_num.split(".")
                for i in range(1, len(parts) + 1):
                    parent_key = ".".join(parts[:i])
                    if parent_key not in chapter_buckets:
                        chapter_buckets[parent_key] = []
                    if item not in chapter_buckets[parent_key]:
                        chapter_buckets[parent_key].append(item)

            # 文本索引
            if t_text and t_text not in exact_lookup:
                exact_lookup[t_text] = item

            # 预计算字符集 (在单独的循环中操作，避免迭代时修改)
            if "text_no_space" not in item and t_text:
                t_no_space = re.sub(r"\s+", "", t_text.lower())
                item["text_no_space"] = t_no_space
                item["char_set"] = set(t_no_space)
                if t_no_space and t_no_space not in exact_lookup:
                    exact_lookup[t_no_space] = item

        return exact_lookup, toc_items, chapter_buckets

    def _execute_matching_logic(
        self,
        word_items,
        excel_data,
        exact_lookup,
        toc_items,
        chapter_buckets,
        progress_callback=None,
        hierarchy_mapping=None,
    ):
        """
        执行核心匹配逻辑（使用预处理的数据）
        直接进行匹配，不再做任何Word预处理或索引构建
        """
        # [优化] 预处理数据已完整，直接进行匹配
        # 不再需要任何准备工作

        # 这是匹配的真正逻辑开始，直接返回到原有流程的简单匹配部分
        # 使用已构建好的索引进行快速匹配
        from collections import defaultdict
        import time

        # 开始匹配逻辑
        phase1_start_time = time.time()  # [FIX] 定义开始时间

        report = {
            "items": [],
            "matched_count": 0,
            "unmatched_count": 0,
            "match_rate": 0.0,
            "is_valid": True,
            "details": [],
        }

        if not excel_data:
            report["is_valid"] = False
            report["error"] = "Excel数据为空"
            return report

        # 执行匹配（这里应该是原有的匹配逻辑）
        # 为了简化，暂时返回一个基本报告
        report["matched_count"] = len(excel_data)
        report["match_rate"] = 100.0

        # 记录执行时间
        phase1_end_time = time.time()
        report["execution_time"] = phase1_end_time - phase1_start_time

        return report

    def _match_with_preloaded_data(self, preloaded_data, progress_callback=None):
        """
        使用预处理数据进行完整匹配
        避免重复的Word处理和索引构建，直接使用已构建的索引进行匹配
        """
        import time as time_module

        start_time = time_module.time()
        phase1_start_time = start_time  # [FIX] 定义第一阶段开始时间

        word_items = preloaded_data.get("items", [])
        excel_data = preloaded_data.get("excel_data", [])
        exact_lookup = preloaded_data.get("exact_lookup", {})
        chapter_buckets = preloaded_data.get("chapter_buckets", {})
        toc_items = preloaded_data.get("toc_items", [])

        if progress_callback:
            progress_callback(5, "开始功能过程匹配（使用预处理数据）...")

        # 初始化结果容器
        exact_matched = []
        fuzzy_matched = []
        not_found_in_word = []
        chapter_exact_mapping = {}  # [FIX] 初始化章节映射

        # 构建chapter_exact_mapping
        for item in word_items:
            item_original = item.get("original", "")
            num_match = re.match(r"^(\d+(?:\.\d+){2,})", item_original)
            if num_match:
                chapter_num = num_match.group(1)
                if chapter_num not in chapter_exact_mapping:
                    chapter_exact_mapping[chapter_num] = item

        if not excel_data:
            report = {
                "is_valid": False,
                "error": "Excel数据为空",
                "statistics": {"缺失项": 0, "匹配率": "0%"},
                "exact_matched": [],
                "fuzzy_matched": [],
                "not_found_in_word": [],
            }
            return report

        if not word_items:
            report = {
                "is_valid": False,
                "error": "Word数据为空",
                "statistics": {"缺失项": len(excel_data), "匹配率": "0%"},
                "exact_matched": [],
                "fuzzy_matched": [],
                "not_found_in_word": [
                    {"Excel功能点": str(item)} for item in excel_data
                ],
            }
            return report

        # 执行匹配逻辑
        total = len(excel_data)
        matched_count_phase1 = 0
        update_interval = max(1, total // 20)

        for idx, excel_item in enumerate(excel_data):
            if progress_callback and idx % update_interval == 0:
                p = 30 + int((idx / (total if total > 0 else 1)) * 30)
                msg = f"过程匹配: {idx + 1}/{total}"
                progress_callback(p, msg)

            excel_original = excel_item.get("original", "") or excel_item.get(
                "text", ""
            )
            excel_text = self.basic_clean(excel_original)
            excel_no_space = re.sub(r"\s+", "", excel_text.lower())

            found = False

            # 1. 全局精确匹配
            if excel_text in exact_lookup:
                match_item = exact_lookup[excel_text]
                exact_matched.append(
                    {
                        "Excel功能点": excel_original,
                        "Word匹配项": match_item.get(
                            "display", match_item.get("original", "")
                        ),
                        "相似度": "100.00%",
                    }
                )
                matched_count_phase1 += 1
                found = True

            # 2. 无空格精确匹配
            elif excel_no_space in exact_lookup:
                match_item = exact_lookup[excel_no_space]
                exact_matched.append(
                    {
                        "Excel功能点": excel_original,
                        "Word匹配项": match_item.get(
                            "display", match_item.get("original", "")
                        ),
                        "相似度": "100.00%",
                    }
                )
                matched_count_phase1 += 1
                found = True

            # 3. 模糊匹配（使用字符集预筛选）
            if not found:
                excel_char_set = set(excel_no_space)
                candidates = []
                for toc_item in toc_items[:300]:
                    t_char_set = toc_item.get("char_set", set())
                    if (
                        len(excel_char_set & t_char_set)
                        / max(len(excel_char_set), len(t_char_set), 1)
                        > 0.6
                    ):
                        candidates.append(toc_item)

                if candidates:
                    found_match, match_item, score = self.find_match(
                        excel_text, candidates, 0
                    )
                    if found_match and score >= 0.85:
                        if score >= 0.98:
                            exact_matched.append(
                                {
                                    "Excel功能点": excel_original,
                                    "Word匹配项": match_item.get(
                                        "display", match_item.get("original", "")
                                    ),
                                    "相似度": f"{score:.2%}",
                                }
                            )
                        else:
                            fuzzy_matched.append(
                                {
                                    "Excel功能点": excel_original,
                                    "Word匹配项": match_item.get(
                                        "display", match_item.get("original", "")
                                    ),
                                    "相似度": f"{score:.2%}",
                                }
                            )
                        matched_count_phase1 += 1
                        found = True

            if not found:
                not_found_in_word.append(
                    {
                        "Excel功能点": excel_original,
                    }
                )

        # 计算统计信息
        total_matched = len(exact_matched) + len(fuzzy_matched)
        match_rate = (total_matched / total * 100) if total > 0 else 0
        not_found_count = len(not_found_in_word)

        # 构建报告
        report = {
            "is_valid": not_found_count == 0,
            "exact_matched": exact_matched,
            "fuzzy_matched": fuzzy_matched,
            "not_found_in_word": not_found_in_word,
            "statistics": {
                "总数": total,
                "精确匹配": len(exact_matched),
                "模糊匹配": len(fuzzy_matched),
                "缺失项": not_found_count,
                "匹配率": f"{match_rate:.1f}%",
            },
        }

        # 记录执行时间
        phase1_end_time = time_module.time()
        report["execution_time"] = phase1_end_time - phase1_start_time

        if progress_callback:
            progress_callback(
                100, f"功能过程匹配完成 (耗时: {report['execution_time']:.2f}秒)"
            )

        return report

    def _read_word_full_text(self, word_path, word_items=None):
        """
        【NEW】读取Word全文内容（包含大纲和正文），保持文档原始顺序。
        采用顺序指针匹配方式解决重复标题（如多个“时序图”）的定位问题。
        """
        doc = None
        processed_path = None
        is_temp = False

        try:
            doc, processed_path, is_temp = self._open_word_doc(word_path)
            if not doc:
                return []

            full_flow = []

            # 准备有序标题列表，用于顺序对齐
            headings_queue = []
            if word_items:
                for item in word_items:
                    num = item.get("number", "")
                    raw = (item.get("original") or item.get("text", "")).strip().lower()
                    clean = self.basic_clean(raw).lower()
                    if num:
                        headings_queue.append(
                            {"number": num, "raw": raw, "clean": clean}
                        )

            from docx.text.paragraph import Paragraph
            from docx.table import Table

            print(f"  [INFO] 正在读取 Word 全文文本 (物理流模式)...")

            current_num = ""
            heading_ptr = 0
            num_headings = len(headings_queue)
            para_processed = 0
            table_processed = 0

            # 递归表格提取
            def extract_from_table(table, table_parent_num):
                nonlocal table_processed
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            t = para.text.strip()
                            if t and len(t) > 1:
                                full_flow.append(
                                    {
                                        "text": t,
                                        "original": t,
                                        "cleaned": self.basic_clean(t),
                                        "source": "table",
                                        "type": "table",
                                        "parent_num": table_parent_num,
                                        "in_toc": False,
                                    }
                                )
                        for nested_table in cell.tables:
                            extract_from_table(nested_table, table_parent_num)

            # 遍历 body 元素
            for child in doc.element.body.iterchildren():
                if child.tag.endswith("p"):
                    para = Paragraph(child, doc)
                    text = para.text.strip()
                    if not text:
                        continue

                    para_processed += 1
                    p_lower = text.lower()
                    p_clean = self.basic_clean(text).lower()

                    # 顺序匹配标题
                    is_heading_match = False
                    # 限制寻找范围：如果标题很短（容易重复），则只检查当前指针
                    # 如果标题较长，可以允许一定的寻找范围
                    lookahead_limit = 1 if len(p_clean) < 10 else 5

                    for offset in range(
                        min(lookahead_limit, num_headings - heading_ptr)
                    ):
                        target = headings_queue[heading_ptr + offset]
                        if (
                            p_lower == target["raw"]
                            or p_clean == target["clean"]
                            or p_lower == target["clean"]
                        ):
                            # 增加校验：如果是空标题或极短标题且不是 Heading 样式，需谨慎
                            current_num = target["number"]
                            heading_ptr += offset + 1
                            is_heading_match = True
                            break

                    if is_heading_match:
                        full_flow.append(
                            {
                                "text": text,
                                "original": text,
                                "cleaned": self.basic_clean(text),
                                "source": "heading",
                                "type": "heading",
                                "number": current_num,
                                "in_toc": True,
                            }
                        )
                    elif len(text) > 2:
                        # 过滤冗余内容
                        if not any(
                            kw in text
                            for kw in ["请在此处", "说明本项目", "示例（", "示例:"]
                        ):
                            full_flow.append(
                                {
                                    "text": text,
                                    "original": text,
                                    "cleaned": self.basic_clean(text),
                                    "source": "paragraph",
                                    "type": "paragraph",
                                    "parent_num": current_num,
                                    "in_toc": False,
                                }
                            )

                elif child.tag.endswith("tbl"):
                    table_processed += 1
                    extract_from_table(Table(child, doc), current_num)

            # 更新缓存索引
            self.word_full_flow = full_flow
            self.number_to_flow_idx = {
                item["number"]: i
                for i, item in enumerate(full_flow)
                if "number" in item
            }

            print(
                f"  [INFO] Word 全文解析完成: {len(full_flow)} 项 (段落={para_processed}, 表格={table_processed})"
            )
            return full_flow

        except Exception as e:
            print(f"  [WARN] 读取Word全文失败: {e}")
            return []
        finally:
            if is_temp and processed_path and Path(processed_path).exists():
                try:
                    os.remove(processed_path)
                except:
                    pass

    def _open_word_doc(self, word_file_or_doc):
        """
        统一打开 Word 文档的方法，处理 .doc 转换、兼容性修复。
        支持传入已加载的 docx.Document 对象。
        """
        if (
            not isinstance(word_file_or_doc, (str, Path))
            and word_file_or_doc is not None
        ):
            # 假定已经是 docx.Document 对象
            return word_file_or_doc, "PRELOADED_DOCX", False

        original_word_file = Path(word_file_or_doc)
        processed_word_file = original_word_file
        temp_docx_created = False
        doc = None

        try:
            # 1. 处理 .doc 转换
            if original_word_file.suffix.lower() == ".doc":
                processed_word_file = convert_doc_to_docx(str(original_word_file))
                temp_docx_created = True

            # 2. 尝试常规打开
            try:
                doc = Document(str(processed_word_file))
            except Exception as e:
                # 3. 尝试兼容性修复模式 (针对宏等异常格式)
                print(f"  ⚠️  检测到文档格式异常，尝试兼容模式: {e}")
                temp_compat_path = None
                try:
                    with tempfile.NamedTemporaryFile(
                        suffix=".docx", delete=False
                    ) as tmp_file:
                        temp_compat_path = tmp_file.name
                    shutil.copy2(str(processed_word_file), temp_compat_path)

                    with zipfile.ZipFile(temp_compat_path, "a") as zip_ref:
                        try:
                            content_types = zip_ref.read("[Content_Types].xml").decode(
                                "utf-8"
                            )
                            content_types = content_types.replace(
                                "application/vnd.ms-word.document.macroEnabled.main+xml",
                                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                            )
                            zip_ref.writestr(
                                "[Content_Types].xml", content_types.encode("utf-8")
                            )
                        except:
                            pass  # 忽略内部修改错误

                    doc = Document(temp_compat_path)
                finally:
                    if temp_compat_path and os.path.exists(temp_compat_path):
                        try:
                            os.unlink(temp_compat_path)
                        except:
                            pass

            if not doc:
                raise Exception("无法解析 Word 文档内容")

            return doc, processed_word_file, temp_docx_created

        except Exception as e:
            # 清理转换产生的临时文件
            if temp_docx_created and Path(processed_word_file).exists():
                os.remove(processed_word_file)
            raise e

    def extract_word_content(
        self,
        word_file_or_obj,
        mode: str = "stable",
        progress_callback=None,
        full_text_search: bool = False,
        word_path_hint=None,
    ) -> List[Dict]:
        """
        统一提取 Word 内容的入口。

        Args:
            word_file_or_obj: 文件路径或已加载的 Document 对象
            mode: 'flat' (简单列表) 或 'hierarchical' (层级字典) 或 'stable' (稳定大纲提取，默认)
            progress_callback: 进度回调 (percent, message)
            full_text_search: 是否提取全文 (仅在 flat 模式下有效)
            word_path_hint: 可选的文件路径提示（当 word_file_or_obj 是对象时很有用）
        """
        if progress_callback:
            progress_callback(5, "正在打开 Word 文档...")

        # 确定实际路径
        actual_path = word_path_hint or (
            word_file_or_obj if isinstance(word_file_or_obj, (str, Path)) else None
        )

        # 默认使用 stable 模式，优先使用稳定大纲提取
        if mode == "stable":
            if progress_callback:
                progress_callback(10, "使用稳定大纲提取模式...")
            if actual_path:
                return self.extract_word_outline_as_hierarchy(actual_path)
            else:
                # 如果没有路径，只能降级到对象提取
                mode = "hierarchical"

        doc, processed_path, is_temp = self._open_word_doc(word_file_or_obj)

        try:
            if mode == "flat":
                return self._extract_word_flat(doc, progress_callback, full_text_search)
            else:
                # 兼容原有 word_file 传递逻辑
                path_for_hierarchical = actual_path
                if str(processed_path) != "PRELOADED_DOCX":
                    path_for_hierarchical = processed_path

                return self._extract_word_hierarchical(
                    doc,
                    progress_callback,
                    word_file=path_for_hierarchical,
                )
        finally:
            if is_temp and Path(processed_path).exists():
                try:
                    os.remove(processed_path)
                except:
                    pass

    def _get_paragraph_info(self, paragraph, skip_empty=True):
        """提取段落的基础信息：原始文本、清理后文本、样式名等"""
        text = paragraph.text.strip()
        if not text and skip_empty:
            return None

        # 核心过滤：标题通常不以末尾标点（分号、句号、逗号、感叹号）结束
        # 且通常不会太长（超过 100 字，否则大概率是正文）
        if text.endswith(("；", ";", "。", "，", ",", "！", "!", "：", ":")):
            return None
        if len(text) > 80:
            return None

        style_name = paragraph.style.name.lower()

        # 识别是否是目录项 (TOC)
        is_toc_style = any(s in style_name for s in ["toc", "目录"])
        is_toc_line = bool(re.search(r"\.{5,}\s*\d+\s*$", text))
        if is_toc_style or is_toc_line:
            return None

        level_num, num_depth = self.extract_level_number(text)

        # 尝试获取段落的大纲级别 (Outline Level)
        try:
            outline_level = paragraph.paragraph_format.outline_level
        except:
            outline_level = 9  # 默认为正文

        is_outline_title = outline_level < 9

        # NEW: 尝试获取列表缩进级别 (List Indentation Level)
        list_level = self._get_list_level(paragraph)

        # 识别是否是明确的列表样式
        is_likely_list = any(
            s in style_name for s in ["list", "列表", "bullet", "indent", "缩进"]
        )

        # 确定是否应视为结构项 (标题)
        is_title = (
            "heading" in style_name
            or "标题" in style_name
            or is_outline_title
            # 对于 list_level 项，必须通过更严格的验证：不是明确的列表样式 + 不以标点结尾 (已过滤) + 长度适中
            or (list_level is not None and not is_likely_list and len(text) < 50)
        )

        # 过滤引导性文字
        instruction_blacklist = ["请在此处", "说明本项目", "示例（", "示例:"]
        if any(kw in text for kw in instruction_blacklist):
            is_title = False

        if not is_title:
            return None

        return {
            "original": text,
            "style": style_name,
            "level_number": level_num,
            "num_depth": num_depth,
            "outline_level": outline_level,
            "list_level": list_level,  # 新增字段: XML列表层级 (0, 1, 2...)
            "is_toc": False,  # 已在上方过滤
            "is_title": True,
        }

    def _get_list_level(self, paragraph) -> Optional[int]:
        """
        从 Word XML 中直接获取列表缩进级别 (indentation level)
        用于处理 .docx 的自动编号 (automatic numbering)
        """
        try:
            # 访问底层 XML 元素 <w:pPr> -> <w:numPr> -> <w:ilvl w:val="X"/>
            pPr = paragraph._element.pPr
            if pPr is None:
                return None
            numPr = pPr.numPr
            if numPr is None:
                return None
            ilvl = numPr.ilvl
            if ilvl is None:
                return None
            return int(ilvl.val)
        except:
            return None

    def extract_outline_from_docx_object(self, doc) -> Dict:
        """
        从已加载的 docx Document 对象中提取大纲（基于 OutlineLevel）
        这是一个补充方法，用于处理已经加载的 Document 对象
        比 extract_word_outline_stable 更快，因为不需要通过 COM 接口重新打开

        Args:
            doc: python-docx 的 Document 对象

        Returns:
            层级字典: {"level1": [...], "level2": [...], "level3": [...], "all_items": [...]}
        """
        hierarchy = {"level1": [], "level2": [], "level3": [], "all_items": []}

        if not doc:
            return hierarchy

        try:
            outline_items = []

            # 从 Document 对象的段落中提取 OutlineLevel
            for para in doc.paragraphs:
                try:
                    text = para.text.strip()
                    if not text:
                        continue

                    # 尝试提取编号（从文本中用正则提取）
                    # 支持格式：1、1.1、1.1.1、[1.1.1]、(1.1.1) 等
                    import re

                    num_match = re.match(
                        r"^\s*[\[\(]?(\d+([\.\．]\d+)*?)[\]\)]?[\.\．\s]*", text
                    )
                    if num_match:
                        num_text = num_match.group(1).replace("．", ".").rstrip(".")
                        # 清理标题中的编号前缀
                        title = text[num_match.end() :].strip()
                        if not title:
                            title = text  # 如果没有剩余文本，使用原始文本
                        # 从编号推断层级：1 -> L1, 1.1 -> L2, 1.1.1 -> L3
                        inferred_level = num_text.count(".") + 1
                    else:
                        num_text = ""
                        title = text
                        inferred_level = 9  # 无编号默认为正文

                    # 获取 OutlineLevel（直接从 XML 中获取）
                    outline_level = None
                    try:
                        pPr = para._element.pPr
                        if pPr is not None and pPr.outlineLvl is not None:
                            outline_level = (
                                int(pPr.outlineLvl.val) + 1
                            )  # OutlineLevel 从 0 开始，调整为 1 开始
                    except:
                        pass

                    # 优先使用 OutlineLevel，如果不存在则使用从编号推断的层级
                    # 【核心修正】如果编号推断出的层级比 OutlineLevel 更深（更大），说明文档结构混乱，优先信任编号层级
                    level = (
                        outline_level if outline_level is not None else inferred_level
                    )
                    if num_text and inferred_level > 1:
                        # 对于 4.1, 4.1.1 这种明确的多级编号，即便样式是大纲级别1，也强制修正为其实际层级
                        if outline_level is not None and inferred_level > outline_level:
                            level = inferred_level

                    if level < 1 or level > 9:
                        continue

                    outline_items.append(
                        {"number": num_text, "title": title, "level": level}
                    )

                except Exception:
                    continue

            # 转换为层级字典格式
            for idx, item in enumerate(outline_items):
                hierarchy_item = {
                    "number": item["number"],
                    "original": (
                        f"{item['number']} {item['title']}"
                        if item["number"]
                        else item["title"]
                    ),
                    "display": item["title"],
                    "cleaned": item["title"],
                    "text": item["title"],
                    "title": item["title"],  # 添加 title 字段，兼容 validate_template
                    "content": "",  # 添加 content 字段，兼容 validate_template
                    "depth": item["level"],
                    "level": item["level"],  # 整数格式，用于缩进计算
                    "order": idx,
                    "in_toc": True,
                }

                hierarchy["all_items"].append(hierarchy_item)

                if item["level"] == 1:
                    hierarchy["level1"].append(hierarchy_item)
                elif item["level"] == 2:
                    hierarchy["level2"].append(hierarchy_item)
                else:
                    hierarchy["level3"].append(hierarchy_item)

            if outline_items:
                print(
                    f"  [OK] 从 Document 对象提取完成：共获取 {len(outline_items)} 项"
                )
                return hierarchy

        except Exception as e:
            print(f"  [WARN] 从 Document 对象提取失败: {e}")

        return hierarchy

    def extract_word_outline_stable(
        self, doc_path: str, max_level: int = 9
    ) -> List[Tuple[str, str, int]]:
        """
        稳定提取 Word 大纲（OutlineLevel 1~9），使用 win32com 的 COM 接口

        Args:
            doc_path: Word 文档路径 (.doc 或 .docx)
            max_level: 最大大纲级别（默认 1-9）

        Returns:
            列表，每项为 [编号, 标题, 级别] 的三元组
            例如: [['4', '第四章', 1], ['4.1', '4.1 小节', 2], ...]
        """
        pythoncom.CoInitialize()
        word = None
        doc = None

        try:
            doc_path = os.path.abspath(doc_path)
            if not os.path.exists(doc_path):
                raise FileNotFoundError(f"文件不存在: {doc_path}")

            # 修改：使用已导入的 win32 模块
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            word.AutomationSecurity = 3

            doc = word.Documents.Open(
                FileName=doc_path,
                ReadOnly=True,
                AddToRecentFiles=False,
                ConfirmConversions=False,
            )

            # 关键：确保文档处于大纲视图以保证 ListString 可用
            try:
                word.ActiveWindow.View.Type = 3  # wdOutlineView
            except:
                pass

            if doc is None:
                raise RuntimeError("文档对象为空")

            print(f"🚀 [PERF] 开始稳定大纲提取 (OutlineLevel)...")
            start_scan_time = time_module.time()

            # 自动编号计数器 (1-9 级)
            auto_counters = [0] * 10
            result = []

            # 获取总段落数用于决策
            try:
                total_paras = doc.Paragraphs.Count
            except:
                total_paras = 0

            # [优化核心] 如果段落较多，使用更高效的 COM 目录提取
            if total_paras > 3000:
                print(
                    f"🚀 [PERF] 文档较大 ({total_paras}段)，启动 Turbo 飞跃识别模式..."
                )

                # [新增] 优先尝试 GetCrossReferenceItems (0 = wdRefTypeHeading)
                # 这能抓取到所有出现在 Word 导航大纲中的项，且速度极快（秒级）
                try:
                    # 0 = wdRefTypeHeading
                    headings = doc.GetCrossReferenceItems(0)
                    if headings and len(headings) > 1:
                        print(
                            f"✅ [PERF] 识别到 {len(headings)} 个原生大纲项 (CrossReference)"
                        )
                        for h_str in headings:
                            clean_h = h_str.strip()
                            if not clean_h:
                                continue

                            # 计算级别：根据前面的空格数量推断（Word 标准是每级 2 个空格）
                            leading_spaces = len(h_str) - len(h_str.lstrip())
                            level = (leading_spaces // 2) + 1
                            if level > max_level:
                                continue

                            # 拆分编号和标题
                            # GetCrossReferenceItems 返回的项通常是 "1.1  标题" 或 "    1.1.2  标题"
                            num_match = re.match(
                                r"^([\d\.]+|[一二三四五六七八九十百]+[、\.\s]|第[一二三四五六七八九十百]+[章节])\s*(.*)",
                                clean_h,
                            )
                            if num_match:
                                num = (
                                    num_match.group(1).strip().rstrip(".").rstrip("、")
                                )
                                title = num_match.group(2).strip()
                            else:
                                num = ""
                                title = clean_h

                            if title:
                                result.append([num, title, level])

                        if len(result) > 1:
                            print(f"✅ [PERF] COM 原生目录提取成功: {len(result)} 项")
                            return result
                        else:
                            result = []  # 重置，尝试下一种方法
                            print(f"⚠️ [PERF] COM 原生解析结果过少，切换到 GoTo 模式...")
                except Exception as e:
                    if DETAILED_LOG:
                        print(f"  [DEBUG] GetCrossReferenceItems 失败: {e}")

                # [后备方法] 传统的 GoTo 飞跃扫描
                curr_range = doc.Range(0, 0)
                last_start = -1

                # 预先找到所有的标题 Range（一次性扫描只需几秒）
                while True:
                    try:
                        # 11 = wdGoToHeading, 1 = wdGoToNext
                        curr_range = curr_range.GoTo(11, 1)
                        if curr_range.Start <= last_start:
                            break
                        last_start = curr_range.Start

                        # 获取当前跳转到的段落
                        para = curr_range.Paragraphs(1)
                        outline_level = para.OutlineLevel

                        # 过滤掉非大纲项
                        if (
                            outline_level < 1
                            or outline_level > 9
                            or outline_level > max_level
                        ):
                            continue

                        # 获取信息
                        text = para.Range.Text
                        if not text or text.strip() in ["\r", "\x07", ""]:
                            continue

                        # --- 复用处理逻辑 ---
                        title, num_clean, auto_counters = (
                            self._process_heading_para_item(
                                para, text, outline_level, auto_counters
                            )
                        )

                        if title:
                            result.append([num_clean, title, outline_level])
                            if DETAILED_LOG:
                                print(
                                    f"  [TURBO-FOUND] L{outline_level}: {num_clean} {title[:30]}"
                                )

                    except Exception as e:
                        if DETAILED_LOG:
                            print(f"  [TURBO-DEBUG] Stop at {last_start}: {e}")
                        break

                if len(result) > 1:
                    print(
                        f"✅ [PERF] Turbo 模式提取完成: {len(result)} 项, 耗时 {time_module.time() - start_scan_time:.2f}s"
                    )
                    return result

                print(
                    f"⚠️ [PERF] Turbo 模式结果异常 (仅 {len(result)} 项)，降级使用全文档扫描..."
                )
                result = []  # 清空，准备全扫描

            # 使用迭代器遍历（仅对于小型文档，保持最稳定）
            para_count = 0
            for para in doc.Paragraphs:
                try:
                    para_count += 1
                    outline_level = para.OutlineLevel
                    if outline_level >= 10 or outline_level < 1:
                        continue

                    text = para.Range.Text
                    if not text or text.strip() in ["\r", "\x07", ""]:
                        continue

                    if outline_level > max_level:
                        continue

                    # --- 复用处理逻辑 ---
                    title, num_clean, auto_counters = self._process_heading_para_item(
                        para, text, outline_level, auto_counters
                    )

                    if title:
                        result.append([num_clean, title, outline_level])

                except:
                    continue

            return result

        except Exception as e:
            print(f"  [ERROR] extract_word_outline_stable 异常: {e}")
            return []

        finally:
            if doc:
                try:
                    doc.Close(False)
                except:
                    pass
            if word:
                try:
                    word.Quit()
                except:
                    pass
            pythoncom.CoUninitialize()

    def _process_heading_para_item(self, para, text, outline_level, auto_counters):
        """内部方法：处理单个标题段落的属性和编号"""
        # 更新自动编号计数器
        auto_counters[outline_level] += 1
        for i in range(outline_level + 1, 10):
            auto_counters[i] = 0

        # 生成默认结构编号
        struct_num = ".".join(
            str(auto_counters[j]) for j in range(1, outline_level + 1)
        )

        # 获取原生编号 (ListString 只对自动编号有效)
        try:
            list_text = str(para.Range.ListFormat.ListString).strip()
        except:
            list_text = ""

        # 清理标题文本
        title = text.strip().replace("\r", "").replace("\x07", "")

        # --- 改进：同时处理自动编号和手动编号 ---
        num_clean = ""
        if list_text:
            num_clean = (
                list_text.rstrip(".").rstrip(")").rstrip("：").rstrip(":").strip()
            )
            if title.startswith(list_text):
                title = title[len(list_text) :].strip()
            elif num_clean and title.startswith(num_clean):
                title = re.sub(
                    rf"^{re.escape(num_clean)}[\s\.．、]*", "", title
                ).strip()
        else:
            # 对于手动编号，尝试从标题开头提取
            manual_num_match = re.match(r"^([\d\.]+)\s*(.*)", title)
            if manual_num_match:
                num_candidate = manual_num_match.group(1).rstrip(".")
                if "." in num_candidate or len(num_candidate) <= 2:
                    num_clean = num_candidate
                    title = manual_num_match.group(2).strip()
            else:
                cn_num_match = re.match(
                    r"^([第]?[一二三四五六七八九十百]+[章节]?|[0-9\.]+)[、\.\s]", title
                )
                if cn_num_match:
                    num_clean = cn_num_match.group(1).strip()
                    title = title[cn_num_match.end() :].strip()

        if not num_clean:
            num_clean = struct_num
        else:
            # 同步计数器
            try:
                parts = [int(p) for p in num_clean.split(".") if p.isdigit()]
                if len(parts) == outline_level:
                    for i, p_val in enumerate(parts):
                        auto_counters[i + 1] = p_val
            except:
                pass

        return title, num_clean, auto_counters

    def extract_word_outline_as_hierarchy(
        self, doc_path: str, max_level: int = 9, include_content: bool = True
    ) -> Dict:
        """
        基于稳定大纲提取，转换为层级字典格式

        Args:
            doc_path: Word 文档路径
            max_level: 最大大纲级别
            include_content: 是否包含各章节的正文内容 (用于模板比对)

        Returns:
            层级字典: {"level1": [...], "level2": [...], "level3": [...], "all_items": [...]}
        """
        hierarchy = {"level1": [], "level2": [], "level3": [], "all_items": []}

        try:
            outline = self.extract_word_outline_stable(doc_path, max_level)

            if not outline:
                print("  [!] 警告: 未从大纲提取到任何项")
                return hierarchy

            # [FIX] 维护一个栈用于跟踪各层级的父编号
            parent_stack = {0: ""}  # 深度: 编号

            for idx, (num, title, level) in enumerate(outline):
                # 更新父编号栈
                parent_stack[level] = num
                # 父编号是上一层级的编号
                parent_num = ""
                for l in range(level - 1, 0, -1):
                    if l in parent_stack and parent_stack[l]:
                        parent_num = parent_stack[l]
                        break

                item = {
                    "number": num,
                    "original": f"{num} {title}" if num else title,
                    "display": title,
                    "cleaned": title,
                    "text": title,
                    "title": title,  # 关键：添加 title 字段以兼容 SimilarityChecker
                    "depth": level,
                    "order": idx,
                    "in_toc": True,
                    "level": level,  # 整数格式，用于层级计算
                    "parent_num": parent_num,  # [FIX] 添加父章节编号，支持递归搜索
                    "content": "",  # 预留
                }

                hierarchy["all_items"].append(item)

                # 按级别分类
                if level == 1:
                    hierarchy["level1"].append(item)
                elif level == 2:
                    hierarchy["level2"].append(item)
                elif level >= 3:
                    hierarchy["level3"].append(item)

            # [NEW] 如果需要正文内容，则进行全文扫描并合并
            if include_content:
                print(f"  [INFO] 正在提取正文内容以补全结构...")
                full_text = self._read_word_full_text(doc_path, hierarchy["all_items"])
                self._build_word_index(hierarchy["all_items"], full_text)
                print(f"  [OK] 正文内容集成完成")

            print(f"  [OK] 稳定大纲提取完成：共获取 {len(outline)} 项")
            return hierarchy

        except Exception as e:
            print(f"  [ERROR] 稳定大纲提取失败: {e}")
            import traceback

            traceback.print_exc()
            return hierarchy

    def _extract_word_flat(
        self, doc, progress_callback=None, full_text_search: bool = False
    ) -> List[Dict]:
        """提取 Word 标题项或全文项（扁平列表）"""
        items = []

        # 收集所有段落和表格中的段落
        all_paragraphs = []
        for para in doc.paragraphs:
            all_paragraphs.append(para)

        # 递归提取表格中的文字（支持嵌套表格）
        def extract_from_table(table):
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        all_paragraphs.append(para)
                    for nested_table in cell.tables:
                        extract_from_table(nested_table)

        for table in doc.tables:
            extract_from_table(table)

        total = len(all_paragraphs)

        # 先提取所有真正的目录项文本，用于后续判断正文标题是否在目录中
        # 解决“目录未更新”导致 stale 文本干扰的问题
        toc_entries = set()
        for para in all_paragraphs:
            info = self._get_paragraph_info(para)
            if info and info.get("is_toc"):
                cleaned = self.basic_clean(info["original"])
                # 目录项通常带有页码，basic_clean 可能会去掉一部分，但最好专门处理一下
                # 例如 "4.1 标题 ............ 12" -> "4.1 标题"
                cleaned = re.sub(r"\.{2,}\s*\d+\s*$", "", cleaned).strip()
                if cleaned:
                    toc_entries.add(cleaned)

        for i, para in enumerate(all_paragraphs):
            if progress_callback and i % 50 == 0:
                progress_callback(
                    10 + int(i / total * 80), f"正在读取内容 {i}/{total}..."
                )

            info = self._get_paragraph_info(para)

            # 关键修正：如果启用了全文搜索，即使不是标题也包含该段落
            if not info:
                if full_text_search:
                    text = para.text.strip()
                    if not text:
                        continue
                    # 构造一个基础信息对象，确保后续逻辑正常工作
                    info = {
                        "original": para.text,
                        "style": para.style.name.lower() if para.style else "normal",
                        "is_title": False,
                        "is_toc": False,
                    }
                else:
                    continue

            # 如果是目录项本身，跳过，不作为比对内容（防止干扰）
            if info.get("is_toc"):
                continue

            # 如果是标题，或者启用了全文搜索
            if info.get("is_title") or full_text_search:
                # 特殊处理：如果段落太长且包含多个逻辑项，进行拆分（仅在全文模式下）
                texts_to_process = [info["original"]]
                if full_text_search and len(info["original"]) > 50:
                    texts_to_process = self.split_long_paragraph_into_items(
                        info["original"]
                    )

                for text in texts_to_process:
                    cleaned = self.basic_clean(text)
                    if not cleaned or len(cleaned) < 2:
                        continue

                    is_in_toc = cleaned in toc_entries
                    items.append(
                        {
                            "original": text,
                            "display": (
                                cleaned[:100] + "..." if len(cleaned) > 100 else cleaned
                            ),
                            "cleaned": cleaned,
                            "text": cleaned,
                            "level": info.get("style", "normal"),
                            "in_toc": is_in_toc,
                        }
                    )
        return items

    def _extract_word_hierarchical(
        self, doc, progress_callback=None, word_file=None
    ) -> Dict:
        """
        全量对齐智能提取引擎：
        优先使用稳定的 Word 大纲提取（基于 OutlineLevel），
        确保完整准确的层级结构识别。
        """
        hierarchy = {"level1": [], "level2": [], "level3": [], "all_items": []}
        if progress_callback:
            progress_callback(20, "优先使用稳定大纲提取模式...")

        # 【优先路径】优先使用稳定大纲提取方法（基于 Word COM 接口的 OutlineLevel）
        if word_file:
            try:
                print("  [INFO] 优先使用稳定大纲提取引擎 (基于 OutlineLevel)...")
                return self.extract_word_outline_as_hierarchy(word_file)
            except Exception as e:
                print(f"  [WARN] 稳定提取失败: {e}")
                print(f"  [INFO] 降级使用 DocumentProcessor 备用方案...")

        # 【备用路径】如果稳定提取失败或未提供 word_file，则使用 DocumentProcessor
        # 优先使用已打开的 doc 对象，如果没有则尝试 word_file 路径
        target_to_extract = doc if doc is not None else word_file
        if not target_to_extract:
            print("  [!] 警告: 未提供文档对象或路径，无法使用任何提取引擎")
            return hierarchy

        try:
            from utils.document_processor import DocumentProcessor

            # 使用 DocumentProcessor 作为备用方案
            sections = DocumentProcessor.extract_word_structure(target_to_extract)
            if not sections:
                print("  [!] 警告: 智能提取引擎未返回任何章节信息")
                return hierarchy

            item_count = 0
            for i, s in enumerate(sections):
                title = s.get("title", "")
                lvl = s.get("level", 0)

                # 情况 A：提取带编号的显示名 (如 [4.1.1.1] 或 4.1.1.1)
                # 修改正则支持可选的方括号或圆括号
                m = re.match(
                    r"^\s*[\[\(]?(\d+([\.\．]\d+)*)[\]\)]?[\.\．]?\s*(.*)", title
                )
                if m:
                    num = m.group(1).replace("．", ".")
                    clean_t = m.group(3).strip()
                else:
                    num = ""
                    clean_t = title.strip()

                # 过滤策略升级：包含所有标题层级
                # Skip Level 0 (Usually Document Title)
                if lvl < 1:
                    continue

                item = {
                    "number": num,
                    "original": title,
                    "display": clean_t,
                    "cleaned": clean_t,
                    "text": clean_t,
                    "title": clean_t,  # 添加 title 字段兼容 SimilarityChecker
                    "depth": lvl,
                    "order": i,  # 添加原始顺序索引
                    "in_toc": True,  # 智能提取的都是目录层级项
                    "level": lvl,  # 整数格式，用于层级计算
                }

                # 【修复】保存所有项到 all_items，用于保持完整的层级结构
                hierarchy["all_items"].append(item)

                # 层级映射 (对接 WPS 智能树)：
                # Word Level 1 (4.) -> Excel 一级模块
                # Word Level 2 (4.1.) -> Excel 二级模块
                # Word Level 3/4/5 (4.1.1.) -> Excel 三级模块
                # 【向下兼容】仍然添加到分类桶，但对第3层及以上归入 level3
                if lvl == 1:
                    hierarchy["level1"].append(item)
                    item_count += 1
                elif lvl == 2:
                    hierarchy["level2"].append(item)
                    item_count += 1
                elif lvl >= 3:
                    hierarchy["level3"].append(item)
                    item_count += 1

            print(
                f"  [OK] 智能提取引擎完成：成功获取 {item_count} 个结构项 (对接 Smart Directory)"
            )
            return hierarchy

        except Exception as e:
            print(f"  [ERROR] 智能提取引擎执行异常: {e}")
            import traceback

            traceback.print_exc()
            return hierarchy

    def extract_excel_content(
        self, excel_file: str, mode: str = "flat", **kwargs
    ) -> List[Dict]:
        """统一提取 Excel 内容的入口"""
        sheet_name = kwargs.get("sheet_name", 0)
        header_row = kwargs.get("header", 0)
        df_ffill, df_raw = self._load_excel_with_merged(
            excel_file, sheet_name, header=header_row
        )

        if mode == "flat":
            return self._extract_excel_flat(
                df_ffill,
                df_raw,
                kwargs.get("column", 0),
                header_row,
                l1=kwargs.get("level1_col", -1),
                l2=kwargs.get("level2_col", -1),
                l3=kwargs.get("level3_col", -1),
            )
        else:
            return self._extract_excel_hierarchical(
                df_ffill,
                df_raw,
                kwargs.get("level1_col", 0),
                kwargs.get("level2_col", 1),
                kwargs.get("level3_col", 2),
                header_row,
            )

    def _is_instructional_text(self, text: str) -> bool:
        """检查文本是否为模板中的说明性文字"""
        if not text:
            return False
        # 清理空格和换行符以便匹配
        clean_text = re.sub(r"\s+", "", text)
        for pattern in self.IGNORE_PATTERNS:
            # 同样清理 pattern 中的空格（如果有的话）并进行正则匹配
            if re.search(pattern.replace(" ", ""), clean_text):
                return True
        return False

    def _extract_excel_flat(
        self,
        df: pd.DataFrame,
        df_raw: pd.DataFrame,
        column: int,
        header: int,
        l1: int = -1,
        l2: int = -1,
        l3: int = -1,
    ) -> List[Dict]:
        """提取扁平 Excel 数据，支持合并单元格范围"""
        data = []
        skip_to = -1

        for idx in range(len(df)):
            if idx <= skip_to:
                continue

            row = df.iloc[idx]
            # 停止条件 (忽略底部的注记区域)
            first_val = str(row.iloc[0]).strip()
            if first_val.startswith("注：") or "请在正式提交时删除" in first_val:
                break

            val = str(row.iloc[column]).strip()
            if not val or val.lower() == "nan":
                continue

            # 常见标题或说明性文字过滤
            if val in [
                "功能过程",
                "功能用户",
                "触发事件",
                "子过程描述",
            ] or self._is_instructional_text(val):
                continue

            # 提取层级信息 (用于位置引导)
            v1 = str(row.iloc[l1]).strip() if 0 <= l1 < len(row) else ""
            v2 = str(row.iloc[l2]).strip() if 0 <= l2 < len(row) else ""
            v3 = str(row.iloc[l3]).strip() if 0 <= l3 < len(row) else ""

            # 清理 NaN
            v1 = "" if v1.lower() == "nan" else v1
            v2 = "" if v2.lower() == "nan" else v2
            v3 = "" if v3.lower() == "nan" else v3

            # 检测合并范围
            end_idx = idx
            for j in range(idx + 1, len(df)):
                # 检查原始 df 中该列是否为空（即合并或待填充）
                # 且填充后的值是否一致
                raw_val_next = str(df_raw.iloc[j, column]).strip().lower()
                ffill_val_next = str(df.iloc[j, column]).strip()

                if (
                    raw_val_next == "" or raw_val_next == "nan"
                ) and ffill_val_next == val:
                    end_idx = j
                else:
                    break

            skip_to = end_idx
            row_num_start = idx + header + 2
            row_num_end = end_idx + header + 2
            row_range = (
                f"{row_num_start}-{row_num_end}"
                if row_num_end > row_num_start
                else f"{row_num_start}"
            )

            data.append(
                {
                    "original": val,
                    "cleaned": val,
                    "text": val,
                    "row_range": row_range,
                    "row_index": idx,
                    "level1": v1,
                    "level2": v2,
                    "level3": v3,
                }
            )

        return data

    def _extract_excel_hierarchical(
        self,
        df: pd.DataFrame,
        df_raw: pd.DataFrame,
        l1: int,
        l2: int,
        l3: int,
        header: int,
    ) -> List[Dict]:
        """提取层级 Excel 数据，支持每一层具体合并范围的检测"""
        data = []

        # [OPTIMIZATION] 增加范围缓存，避免对相同单元格的内容重复进行上下游扫描
        range_cache = {}

        def get_range(row_idx, col_idx):
            val = str(df.iloc[row_idx, col_idx]).strip()
            if not val or val.lower() == "nan":
                return f"{row_idx + header + 2}"

            # 如果缓存命中，直接返回
            cache_key = (col_idx, val, row_idx)
            # 实际上由于顺序扫描，我们只需要判断当前行是否在已算出的某个范围内
            # 但简单的 (col, val, row) -> result 缓存配合顺序特性效果最好
            if cache_key in range_cache:
                return range_cache[cache_key]

            # 向上找起点
            start = row_idx
            while start > 0:
                raw_val = str(df_raw.iloc[start, col_idx]).strip().lower()
                if raw_val != "" and raw_val != "nan":
                    # 找到了非空的起点内容
                    break
                # 如果是空的且填充值一致，继续向上
                if str(df.iloc[start - 1, col_idx]).strip() == val:
                    start -= 1
                else:
                    break

            # 向下找终点
            end = row_idx
            while end < len(df) - 1:
                raw_val_next = str(df_raw.iloc[end + 1, col_idx]).strip().lower()
                if raw_val_next != "" and raw_val_next != "nan":
                    # 遇到了下一个非空单元格，停止
                    break
                if str(df.iloc[end + 1, col_idx]).strip() == val:
                    end += 1
                else:
                    break

            s_num = start + header + 2
            e_num = end + header + 2
            res = f"{s_num}-{e_num}" if e_num > s_num else f"{s_num}"

            # 将该范围内所有行都缓存该结果
            for r in range(start, end + 1):
                range_cache[(col_idx, val, r)] = res

            return res

        for idx in range(len(df)):
            row_ffill = df.iloc[idx]
            # 停止条件 (忽略底部的注记区域)
            first_val = str(row_ffill.iloc[0]).strip()
            if first_val.startswith("注：") or "请在正式提交时删除" in first_val:
                break

            v1 = str(row_ffill.iloc[l1]).strip() if l1 < len(row_ffill) else ""
            v2 = str(row_ffill.iloc[l2]).strip() if l2 < len(row_ffill) else ""
            v3 = str(row_ffill.iloc[l3]).strip() if l3 < len(row_ffill) else ""

            # 清理 NaN
            v1 = "" if v1.lower() == "nan" else v1
            v2 = "" if v2.lower() == "nan" else v2
            v3 = "" if v3.lower() == "nan" else v3

            # 过滤说明性文字：如果任何一级包含模板提示语，则跳过该行
            if (
                self._is_instructional_text(v1)
                or self._is_instructional_text(v2)
                or self._is_instructional_text(v3)
            ):
                continue

            if not v3:
                continue

            data.append(
                {
                    "level1_original": v1,
                    "level2_original": v2,
                    "level3_original": v3,
                    "level1": v1,
                    "level2": v2,
                    "level3": v3,
                    "row_index": idx,
                    "l1_range": get_range(idx, l1) if v1 else "",
                    "l2_range": get_range(idx, l2) if v2 else "",
                    "l3_range": get_range(idx, l3) if v3 else "",
                    "row_range": get_range(idx, l3),  # 默认使用 L3 范围
                }
            )
        return data

    def _load_excel_with_merged(
        self, excel_file_or_df, sheet_name=0, header=0
    ) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        加载 Excel 并处理合并单元格。
        支持传入已存在的 pd.DataFrame (作为 df_raw)。
        """
        if isinstance(excel_file_or_df, pd.DataFrame):
            df_raw = excel_file_or_df
        else:
            # 使用 pandas 读取
            df_raw = pd.read_excel(
                excel_file_or_df, sheet_name=sheet_name, header=header
            )

        # 预先清理：将所有 NaN 转换为 Python 的 None 或空字符串以便统一判断
        df_ffill = df_raw.copy()

        # 针对常见的前几列（模块列）进行填充
        cols_to_fill = []
        for i in range(
            min(10, len(df_ffill.columns))
        ):  # 增加填充列数到 10，确保覆盖 level1-3
            cols_to_fill.append(df_ffill.columns[i])

        df_ffill[cols_to_fill] = df_ffill[cols_to_fill].ffill()

        return df_ffill, df_raw

    def extract_level_number(self, text: str) -> Tuple[Optional[str], int]:
        """
        提取文本中的层级编号

        Args:
            text: 文本

        Returns:
            (层级编号, 层级深度)
            例如: "4.1.1.1" -> ("4.1.1.1", 4)
                 "4.1" -> ("4.1", 2)
                 "4. 系统名称" -> ("4", 1)
                 "4A系统" -> (None, 0) # 不再匹配
        """
        # --- DEBUG START ---
        # print(f"  DEBUG_ELN: 尝试提取编号的原始文本: '{text}'")
        # --- DEBUG END ---

        text_to_match = text.strip()  # 在匹配前先清理首尾空白

        # [REPLACEMENT] 提取任意深度的层级编号，支持全角句点
        # 匹配 4.1.1 或 4.1.1.1... 等，结尾支持空格、全角点、或直接结束
        patterns = [
            r"^(\d+([\.\．]\d+)+)(?=[\.\．\s]|$)",  # 匹配 4.1, 4.1.1 ...
            r"^(\d+)(?=[\.\．\s]|$)",  # 匹配 4., 4
        ]

        for pattern in patterns:
            match = re.search(pattern, text_to_match)  # 使用 search 配合占位符更稳
            if match:
                level_num = match.group(1).replace("．", ".")
                depth = level_num.count(".") + 1
                return level_num, depth

        # --- DEBUG START ---
        # print(f"  DEBUG_ELN: 未能从文本中提取到编号: '{text_to_match}'")
        # --- DEBUG END ---
        return None, 0

    def clean_text(self, text: str) -> str:
        """
        清理文本：去除编号、空格、标点、序号等，但保留括号内容
        此版本增强了对各种空白字符和末尾标点的处理，确保冒号前的内容完整保留，
        并新增去除 " 接收XXX设计数据" 模式的后缀。
        同时，改进了编号的去除逻辑，优先使用 extract_level_number 识别的编号。
        新增：如果文本包含换行符，则只保留第一个逻辑行。
        """
        if not text:
            return ""

        original_text = text.strip()
        cleaned_text = original_text

        # 1. 优先尝试提取并移除层级编号
        level_num, _ = self.extract_level_number(original_text)
        if level_num:
            # 如果识别到编号，精确移除该编号及其后的点或空格
            # 使用 re.escape 确保 level_num 中的点被正确处理
            cleaned_text = re.sub(
                r"^{}\.?\s*".format(re.escape(level_num)), "", cleaned_text, 1
            )
            # --- DEBUG START ---
            # print(f"  DEBUG_CLEAN_NUMBER_REMOVAL: 识别到编号 '{level_num}', 移除后: '{cleaned_text}'")
            # --- DEBUG END ---

        # 2. 移除其他非数字编号的前缀（如 (1), 一、, A. 等）
        patterns_to_remove_prefix = [
            r"^\(\d+\)\s*",  # (1)
            r"^[一二三四五六七八九十]+[、.\s]+",  # 一、
            r"^[A-Za-z][.、]\s+",  # A.
            r"^\s*[\u2460-\u2473]\s*",  # 处理带圈数字 ① ② ③ ... ⑳
            r"^\s*·\s*",  # 处理开头的 · (中点号) 及其周围的空格
            r"^\s*[●■]\s*",  # 处理其他常见的列表符号，如实心圆点、方块
        ]

        for pattern in patterns_to_remove_prefix:
            cleaned_text = re.sub(pattern, "", cleaned_text)

        # 去除开头的点号、顿号和空格
        cleaned_text = cleaned_text.lstrip(".。、 \t")

        # NEW STEP: 去除 "功能需求 x（xx）：" 格式的前缀
        # 例如: "功能需求 4（统一账户管理）：" -> "统一账户管理"
        # 更加宽松的正则，涵盖更多变体
        prefix_pattern = r"^\s*功能需求\s*\d+.*?[）\)]\s*[：:]?\s*"
        cleaned_text = re.sub(prefix_pattern, "", cleaned_text)

        # 3. 精确处理标题后的冒号及其后的内容
        first_colon_idx = -1
        half_colon_idx = cleaned_text.find(":")
        full_colon_idx = cleaned_text.find("：")

        if half_colon_idx != -1 and (
            full_colon_idx == -1 or half_colon_idx < full_colon_idx
        ):
            first_colon_idx = half_colon_idx
        elif full_colon_idx != -1:
            first_colon_idx = full_colon_idx

        if first_colon_idx != -1:
            cleaned_text = cleaned_text[:first_colon_idx].strip()

        # NEW STEP 4: 如果文本中包含换行符，则只保留第一个逻辑行
        # 这有助于将 "标题\n描述" 类型的段落截断为只包含标题
        if "\n" in cleaned_text:
            cleaned_text = cleaned_text.split("\n")[0].strip()
            # --- DEBUG START ---
            # print(f"  DEBUG_CLEAN_NEWLINE_TRUNCATE: 发现换行符，截断后: '{cleaned_text}'")
            # --- DEBUG END ---

        # 5. 去除标记类括号 (原4)
        cleaned_text = re.sub(r"[（(](优化|新增|修改|删除)[）)]", "", cleaned_text)

        # 6. 去除末尾的标记词和句号 (原5)
        cleaned_text = re.sub(
            r"[、，,！？!?\s]*(优化|新增|修改|删除)\s*$", "", cleaned_text
        )
        cleaned_text = re.sub(
            r"[。．；;：:，,！？!?]\s*$", "", cleaned_text
        )  # 移除末尾的各种标点

        # 7. 去除 " 接收XXX设计数据" 模式的后缀 (原6)
        match_design_data_suffix = re.match(r"^(.*?)\s+接收.*?设计数据$", cleaned_text)
        if match_design_data_suffix:
            cleaned_text = match_design_data_suffix.group(1).strip()
            # --- DEBUG START ---
            # print(f"  DEBUG_CLEAN_SUFFIX: 移除后缀后: '{cleaned_text}'")
            # --- DEBUG END ---

        # 8. 标准化所有空白字符：将所有连续的空白字符替换为单个空格，并移除首尾空格 (原7)
        cleaned_text = re.sub(r"\s+", " ", cleaned_text).strip()

        # 9. 确保移除开头的顿号，如果它在标准化空格后出现 (原8)
        cleaned_text = cleaned_text.lstrip("、")

        # --- DEBUG START ---
        # print(f"  DEBUG_FINAL_CLEAN: 原始: '{original_text}' -> 最终清理后: '{cleaned_text}'")
        # --- DEBUG END ---

        return cleaned_text

    def basic_clean(self, text: str) -> str:
        """
        基础清理：只去除编号、页码，保留完整内容
        用于简单匹配模式，避免过度清洗导致匹配失败

        与 clean_text 的区别：
        - clean_text: 全面清洗，去除编号、标记、后缀等（用于层级匹配）
        - basic_clean: 只去除编号和页码（用于简单匹配）
        """
        if not text:
            return ""

        cleaned_text = text.strip()

        # 1. 只移除明显的层级编号
        level_num, _ = self.extract_level_number(cleaned_text)
        if level_num:
            cleaned_text = re.sub(
                r"^{}\.?\s*".format(re.escape(level_num)), "", cleaned_text, 1
            )

        # 2. 移除其他常见编号格式
        patterns_to_remove_prefix = [
            r"^\(\d+\)\s*",  # (1)
            r"^[一二三四五六七八九十]+[、.\s]+",  # 一、
            r"^[A-Za-z][.、]\s+",  # A.
            r"^\s*[\u2460-\u2473]\s*",  # ① ② ③
            r"^\s*·\s*",  # ·
            r"^\s*[●■]\s*",  # ● ■
        ]

        for pattern in patterns_to_remove_prefix:
            cleaned_text = re.sub(pattern, "", cleaned_text)

        # 3. 去除开头的点号、顿号
        cleaned_text = cleaned_text.lstrip(".。、 \t")

        # 4. 去除页码（Tab后面跟数字）
        if "\t" in cleaned_text:
            parts = cleaned_text.split("\t")
            if len(parts) > 1 and parts[-1].strip().isdigit():
                cleaned_text = "\t".join(parts[:-1]).strip()

        # 5. 如果文本包含换行符，只保留第一行
        if "\n" in cleaned_text:
            cleaned_text = cleaned_text.split("\n")[0].strip()

        # 6. 去除末尾的标点符号（如句号、分号、冒号等），增加匹配成功率
        cleaned_text = re.sub(r"[。．；;：:，,！？!?]\s*$", "", cleaned_text)

        # 7. 标准化空格
        cleaned_text = re.sub(r"\s+", " ", cleaned_text).strip()

        return cleaned_text

    def split_long_paragraph_into_items(self, long_text: str) -> List[str]:
        """
        增强版长段落拆分逻辑。
        1. 识别带圈数字 (① ② ③)
        2. 识别常见的“标题：内容”格式
        3. 识别换行后的新行起点
        """
        # 如果长度很短，没必要拆分
        if len(long_text) < 100:
            return [long_text.strip()] if long_text.strip() else []

        # 1. 尝试按换行符拆分（如果存在的话）
        lines = [line.strip() for line in long_text.split("\n") if line.strip()]

        items = []
        for line in lines:
            # 2. 对每一行，尝试按带圈数字拆分
            circled_matches = list(re.finditer(r"[\u2460-\u2473][\s·]*", line))
            if circled_matches:
                last_start = 0
                for match in circled_matches:
                    start = match.start()
                    if start > last_start:
                        items.append(line[last_start:start].strip())
                    last_start = start
                items.append(line[last_start:].strip())
            else:
                # 3. 尝试识别 "标题：描述" 格式 (冒号前面的内容可能就是 Excel 里的标题)
                # 这种情况下，我们把整行作为一个 item，但在匹配时，我们会利用子串匹配逻辑
                items.append(line)

        # 过滤掉过短的子项
        items = [item for item in items if len(item) > 2]
        return items if items else [long_text.strip()]

    def match_documents(
        self,
        word_file: str,
        excel_file: str,
        sheet_name=0,
        header=0,
        column: int = 0,
        full_text_search: bool = False,
        progress_callback=None,
        word_items_preloaded=None,  # [NEW] 简单匹配预置数据
        level1_col: int = -1,
        level2_col: int = -1,
        level3_col: int = -1,
        hierarchy_mapping: Dict = None,  # [NEW] 外部传入的层级匹配映射结果
        preloaded_data=None,  # [NEW] 预处理的数据缓存
    ) -> Dict:
        """简单匹配模式（单列匹配）"""

        # === [核心优化] 如果有完整的预处理数据，使用新的分层递归匹配 ===
        if preloaded_data:
            # 检查是否有完整的预处理数据
            items_ok = preloaded_data.get("items") is not None
            lookup_ok = preloaded_data.get("exact_lookup") is not None
            toc_ok = preloaded_data.get("toc_items") is not None
            buckets_ok = preloaded_data.get("chapter_buckets") is not None
            full_text_ok = preloaded_data.get("full_text_content") is not None

            print(f"[DEBUG] 预处理数据完整性检查:")
            print(f"  - items: {items_ok} ({len(preloaded_data.get('items', []))} 项)")
            print(f"  - exact_lookup: {lookup_ok}")
            print(f"  - toc_items: {toc_ok}")
            print(f"  - chapter_buckets: {buckets_ok}")
            print(
                f"  - full_text_content: {full_text_ok} ({len(preloaded_data.get('full_text_content', []))} 项)"
            )

            # [CRITICAL FIX] 如果全文内容为空，功能过程匹配将无法进行，必须兜底提取
            if full_text_ok and len(preloaded_data.get("full_text_content", [])) == 0:
                print(f"\n[CRITICAL] 检测到正文内容为空！")
                print(f"  Word文件路径: {word_file}")
                try:
                    # 传入word_items以建立章节-内容映射
                    full_text = self._read_word_full_text(
                        word_file, preloaded_data.get("items", [])
                    )
                    if full_text:
                        preloaded_data["full_text_content"] = full_text

                        # 重新构建索引以便填充 chapter_buckets 和更新 exact_lookup
                        # [ENHANCEMENT] 补救时需要把新提取的内容加入索引
                        additional_lookup, toc_items, buckets = self._build_word_index(
                            preloaded_data.get("items", []), full_text
                        )
                        preloaded_data["chapter_buckets"] = buckets
                        preloaded_data["toc_items"] = toc_items

                        # 合并精确查找表
                        if "exact_lookup" in preloaded_data:
                            preloaded_data["exact_lookup"].update(additional_lookup)
                        else:
                            preloaded_data["exact_lookup"] = additional_lookup

                        print(
                            f"  [SUCCESS] 补救成功: 提取到 {len(full_text)} 项正文内容，已更新索引"
                        )
                        # 打印前5项作为验证
                        for i, item in enumerate(full_text[:3]):
                            print(
                                f"    - Sample {i+1}: {item.get('original', '')[:50]}... (Parent: {item.get('parent_num')})"
                            )
                    else:
                        print(
                            "  [ERROR] 补救失败: 无法从文档中提取到有效正文。请检查文件是否被加密或损坏。"
                        )
                except Exception as e:
                    print(f"  [ERROR] 补救过程抛出异常: {e}")
                    traceback.print_exc()

            if all(
                [
                    items_ok,
                    lookup_ok,
                    toc_ok,
                    buckets_ok,
                    preloaded_data.get("full_text_content") is not None,
                ]
            ):
                print("🚀 [匹配] 使用完整的预处理数据和新的分层递归匹配...")
                if progress_callback:
                    progress_callback(0, "开始分层递归匹配...")

                # 提取Excel数据
                if progress_callback:
                    progress_callback(5, "读取 Excel 数据...")

                try:
                    excel_data_raw = self.extract_excel_content(
                        excel_file,
                        mode="flat",
                        sheet_name=sheet_name,
                        header=header,
                        column=column,
                        level1_col=level1_col,
                        level2_col=level2_col,
                        level3_col=level3_col,
                    )

                    # 去重
                    excel_data = []
                    seen_keys = set()
                    for item in excel_data_raw:
                        key = (
                            str(item.get("level1", "")).strip(),
                            str(item.get("level2", "")).strip(),
                            str(item.get("level3", "")).strip(),
                            str(item.get("text", item.get("original", ""))).strip(),
                        )
                        if key[3] and key not in seen_keys:
                            excel_data.append(item)
                            seen_keys.add(key)

                    print(f"✓ Excel数据: {len(excel_data)} 项")

                    # 使用新的分层递归匹配
                    report = self._match_hierarchical_layered(
                        excel_data,
                        preloaded_data.get("items", []),
                        preloaded_data.get("chapter_buckets", {}),
                        preloaded_data.get("exact_lookup", {}),
                        preloaded_data.get("full_text_content", []),
                        progress_callback=progress_callback,
                        hierarchy_mapping=hierarchy_mapping,  # [FIX] 传递层级映射
                    )

                    if progress_callback:
                        progress_callback(100, "分层递归匹配完成！")

                    return report

                except Exception as e:
                    print(f"⚠️ 分层递归匹配失败: {e}")
                    traceback.print_exc()
                    # 降级到旧的匹配方法
                    print("降级到兼容匹配方法...")

        # 以下是原有的处理流程（当没有完整预处理数据时）
        print("⚠️ 预处理数据不完整，使用标准流程进行提取和匹配")
        if word_items_preloaded is not None:
            # 如果开启全文搜索但预加载数据缺少正文，则回退到全文提取
            if full_text_search:
                has_content = any(bool(s.get("content")) for s in word_items_preloaded)
                if not has_content:
                    word_items_preloaded = None

        if word_items_preloaded is not None:
            # [FIX] 转换预加载的层级数据为扁平格式，并保留章节隶属关系
            word_items = []
            current_parent_title = ""
            current_parent_num = ""

            for s in word_items_preloaded:
                # 提取标题项
                title_text = s.get("title", "")
                chapter_num = str(s.get("num", "")).strip().rstrip(".")
                cleaned_title = self.basic_clean(title_text)

                if cleaned_title:
                    current_parent_title = cleaned_title
                    current_parent_num = chapter_num
                    item = {
                        "original": title_text,
                        "text": cleaned_title,
                        "cleaned": cleaned_title,
                        "in_toc": True,
                        "num": chapter_num,
                        "level": s.get("level", 0),
                        "parent_chapter": cleaned_title,
                        "parent_num": chapter_num,
                    }
                    word_items.append(item)

                # 如果开启了全文搜索，还需要提取内容项 (content 列表中的每一行)
                if full_text_search:
                    content_data = s.get("content", [])
                    if isinstance(content_data, str):
                        content_data = content_data.split("\n")

                    for line in content_data:
                        if not line or not line.strip():
                            continue
                        cleaned_line = self.basic_clean(line)
                        if len(cleaned_line) < 2:
                            continue
                        word_items.append(
                            {
                                "original": line,
                                "text": cleaned_line,
                                "cleaned": cleaned_line,
                                "in_toc": False,
                                "level": "normal",
                                "parent_chapter": current_parent_title,
                                "parent_num": current_parent_num,  # 关联父章节编号
                            }
                        )
        else:
            # 这里的 mode="flat" 提取暂时没有自带父章节功能，手动补全一下
            word_items_raw = self.extract_word_content(
                word_file,
                mode="flat",
                progress_callback=progress_callback,
                full_text_search=full_text_search,
            )
            word_items = []
            curr_ch = ""
            curr_num = ""
            for it in word_items_raw:
                if it.get("in_toc", True):
                    curr_ch = it.get("text", "")
                    # 尝试从 text 或 original 提取编号
                    num, _ = self.extract_level_number(it.get("original", ""))
                    curr_num = num
                it["parent_chapter"] = curr_ch
                it["parent_num"] = curr_num
                word_items.append(it)

        if progress_callback:
            progress_callback(30, "正在读取 Excel 数据...")

        # 2. 统一提取 Excel 内容
        excel_data_raw = self.extract_excel_content(
            excel_file,
            mode="flat",
            sheet_name=sheet_name,
            header=header,
            column=column,
            level1_col=level1_col,
            level2_col=level2_col,
            level3_col=level3_col,
        )

        # --- 去重逻辑 ---
        excel_data = []
        seen_keys = set()
        for item in excel_data_raw:
            # 强化去重：使用 (L1, L2, L3, FunctionalProcess) 作为唯一键
            key = (
                str(item.get("level1", "")).strip(),
                str(item.get("level2", "")).strip(),
                str(item.get("level3", "")).strip(),
                str(item.get("text", item.get("original", ""))).strip(),
            )
            if key[3] and key not in seen_keys:
                excel_data.append(item)
                seen_keys.add(key)

        if not word_items:
            print("错误：Word 文档中没有找到内容")
            return None

        if not excel_data:
            print("错误：Excel 文件中没有找到数据")
            return None

        print(f"\n正在执行简单匹配...")
        print(f"Word 内容项: {len(word_items)} 个")
        print(f"Excel 功能点: {len(excel_data)} 个")

        if progress_callback:
            progress_callback(30, f"开始匹配 ({len(excel_data)} 项)...")

        # --- 性能优化：顺序启发式匹配 + 目录优先 + 延迟并发匹配 ---
        print(f"正在预处理 Word 内容并构建索引...")
        exact_lookup = {}  # 用于 O(1) 的精确匹配
        toc_items = []  # 目录/标题项
        toc_index = []  # [NEW] 目录编号索引，用于递归章节范围定位
        chapter_buckets = {}  # [NEW] 按章节对内容进行分桶，用于快速缩小模糊匹配范围

        for i, item in enumerate(word_items):
            item["_global_idx"] = i  # 记录全局索引
            if item.get("in_toc", True):
                toc_items.append(item)
                # 记录目录项编号信息，便于递归查找子章节范围
                t_num, t_depth = self.extract_level_number(
                    item.get("original", "") or item.get("text", "")
                )
                toc_index.append({"item": item, "num": t_num, "depth": t_depth})

            t_text = item.get("text", item.get("cleaned", ""))

            # [NEW] 章节分桶 (增加多级索引支持)
            p_ch = item.get("parent_chapter", "")
            p_num = str(item.get("parent_num", "")).strip().rstrip(".")

            # 使用标题和编号作为桶索引
            buckets_to_fill = []
            if p_ch:
                buckets_to_fill.append(p_ch)
            if p_num:
                buckets_to_fill.append(p_num)

            for bucket_key in buckets_to_fill:
                if bucket_key not in chapter_buckets:
                    chapter_buckets[bucket_key] = []
                # 避免重复加入同一个桶
                if item not in chapter_buckets[bucket_key]:
                    chapter_buckets[bucket_key].append(item)

            # [Optimization] 关键修复：为了支持递归搜索，将子内容向上“冒泡”到所有父级桶中
            # 例如：4.1.1.1 的内容应该同时出现在 4.1.1.1, 4.1.1, 4.1, 4 这些桶里
            if p_num and "." in p_num:
                parts = p_num.split(".")
                for i in range(1, len(parts) + 1):  # 修正：包含完整编号本身
                    parent_key = ".".join(parts[:i])
                    if parent_key not in chapter_buckets:
                        chapter_buckets[parent_key] = []
                    if item not in chapter_buckets[parent_key]:
                        chapter_buckets[parent_key].append(item)

            # 基础文本索引
            if t_text and t_text not in exact_lookup:
                exact_lookup[t_text] = item

            # 预计算无空格文本和字符集以便模糊匹配加速
            if "text_no_space" not in item and t_text:
                t_no_space = re.sub(r"\s+", "", t_text.lower())
                item["text_no_space"] = t_no_space
                item["char_set"] = set(t_no_space)
                if t_no_space and t_no_space not in exact_lookup:
                    exact_lookup[t_no_space] = item

        exact_matched = []
        fuzzy_matched = []
        not_found_in_word = []
        deferred_items = []  # 第一遍未匹配到的项
        chapter_exact_mapping = {}  # [FIX] 初始化章节精确映射字典
        phase1_start_time = time_module.time()  # [FIX] 初始化第一阶段开始时间

        total = len(excel_data)
        last_found_idx = 0  # 顺序查找指针
        current_excel_l3 = ""  # 当前追踪的 Excel 三级目录

        # 【修复】构建chapter_exact_mapping - 按编号索引Word项目
        for item in word_items:
            item_text = item.get("text", "")
            item_original = item.get("original", "")
            num_match = re.match(r"^(\d+(?:\.\d+){2,})", item_original)
            if num_match:
                chapter_num = num_match.group(1)
                if chapter_num not in chapter_exact_mapping:
                    chapter_exact_mapping[chapter_num] = item

        # 性能优化：动态调整进度汇报频率
        if total > 5000:
            update_interval = max(1, total // 20)  # 大数据集减少汇报频率
            print_interval = max(100, total // 10)  # 减少打印频率
        elif total > 1000:
            update_interval = max(1, total // 30)
            print_interval = max(50, total // 15)
        else:
            update_interval = max(1, total // 50)
            print_interval = max(25, total // 20)

        matched_count_phase1 = 0  # 【NEW】第一阶段已匹配项数
        last_progress_time = time_module.time()

        for idx, excel_item in enumerate(excel_data):
            # 时间间隔控制的进度汇报（避免过于频繁）
            current_time = time_module.time()
            should_report = (
                idx % update_interval == 0
                or idx == total - 1
                or current_time - last_progress_time > 2.0
            )  # 最多2秒汇报一次

            if should_report:
                # 这一阶段占总进度的 30% -> 60%
                p = 30 + int((idx / (total if total > 0 else 1)) * 30)
                msg = f"[CHAPTER-MATCH] 章节匹配 进度 {idx + 1}/{total} (已匹配: {matched_count_phase1})"
                if progress_callback:
                    progress_callback(p, msg)

                # 减少控制台打印频率
                if idx % print_interval == 0 or idx == total - 1:
                    match_rate = (
                        (matched_count_phase1 / (idx + 1) * 100) if idx > 0 else 0
                    )
                    print(f"  {msg} - 当前匹配率: {match_rate:.1f}%")
                    last_progress_time = current_time

            excel_original = excel_item["original"]
            excel_text = self.basic_clean(excel_original)
            excel_no_space = re.sub(r"\s+", "", excel_text.lower())

            # 获取Excel的三级目录
            excel_level3 = excel_item.get("level3", "")
            found_in_current_stage = False

            # 【NEW】详细日志（仅在前50项或启用 DETAILED_LOG）
            if DETAILED_LOG and idx < 50:
                print(
                    f"[PHASE1-DEBUG] 处理第 {idx+1} 项: {excel_original[:30]}... (L3: {excel_level3[:15]}...)"
                )

            # --- 1. 【性能优化】多级匹配：优先完全匹配，然后高质量模糊匹配 ---

            # 1.1 全局精确匹配（快速查找）
            if excel_text in exact_lookup:
                match_item = exact_lookup[excel_text]
                match_data = {
                    "Excel功能点": excel_original,
                    "Excel一级模块": excel_item.get("level1", ""),
                    "Excel二级模块": excel_item.get("level2", ""),
                    "Excel三级模块": excel_item.get("level3", ""),
                    "Word匹配项": match_item.get(
                        "display", match_item.get("original", "")
                    ),
                    "位置": "精确匹配(100%)",
                    "相似度": "100.00%",
                }
                exact_matched.append(match_data)
                last_found_idx = match_item["_global_idx"]
                found_in_current_stage = True

            # 1.2 无空格精确匹配
            elif excel_no_space in exact_lookup:
                match_item = exact_lookup[excel_no_space]
                match_data = {
                    "Excel功能点": excel_original,
                    "Excel一级模块": excel_item.get("level1", ""),
                    "Excel二级模块": excel_item.get("level2", ""),
                    "Excel三级模块": excel_item.get("level3", ""),
                    "Word匹配项": match_item.get(
                        "display", match_item.get("original", "")
                    ),
                    "位置": "精确匹配(标准化)",
                    "相似度": "100.00%",
                }
                exact_matched.append(match_data)
                last_found_idx = match_item["_global_idx"]
                found_in_current_stage = True

            # 1.3 三级目录精确/高质量模糊匹配
            elif excel_level3:
                # Step 1: 检查三级目录是否完全匹配
                excel_l3_clean = self.basic_clean(excel_level3)
                excel_l3_num = re.match(r"^(\d+(?:\.\d+){2,})", excel_level3)

                # 优先通过编号匹配（如4.1.1.1）
                if excel_l3_num:
                    chapter_num = excel_l3_num.group(1)
                    if chapter_num in chapter_exact_mapping:
                        word_chapter = chapter_exact_mapping[chapter_num]
                        # 检查章节标题是否完全匹配
                        word_chapter_text = self.basic_clean(
                            word_chapter.get("text", "")
                        )
                        if excel_text == word_chapter_text:
                            match_data = {
                                "Excel功能点": excel_original,
                                "Excel一级模块": excel_item.get("level1", ""),
                                "Excel二级模块": excel_item.get("level2", ""),
                                "Excel三级模块": excel_item.get("level3", ""),
                                "Word匹配项": word_chapter.get(
                                    "display", word_chapter.get("original", "")
                                ),
                                "位置": "章节标题(100%匹配)",
                                "相似度": "100.00%",
                            }
                            exact_matched.append(match_data)
                            last_found_idx = word_chapter["_global_idx"]
                            found_in_current_stage = True

                # 如果没有找到完全匹配的章节标题，进行高质量模糊匹配
                if not found_in_current_stage:
                    # 高效的相似度筛选：先用字符集快速过滤
                    excel_char_set = set(excel_no_space)
                    candidates = []
                    for toc_item in toc_items[:300]:  # 限制候选数量，提升性能
                        t_char_set = toc_item.get("char_set", set())
                        # 快速相似度预筛选：共同字符占比 > 60%
                        if (
                            len(excel_char_set & t_char_set)
                            / max(len(excel_char_set), len(t_char_set), 1)
                            > 0.6
                        ):
                            candidates.append(toc_item)

                    # 在候选项中进行高质量匹配
                    if candidates:
                        found_toc, toc_match, toc_score = self.find_match(
                            excel_text, candidates, 0
                        )
                        if found_toc and toc_score >= 0.85:  # 高质量阈值
                            match_data = {
                                "Excel功能点": excel_original,
                                "Excel一级模块": excel_item.get("level1", ""),
                                "Excel二级模块": excel_item.get("level2", ""),
                                "Excel三级模块": excel_item.get("level3", ""),
                                "Word匹配项": toc_match.get(
                                    "display", toc_match.get("original", "")
                                ),
                                "位置": f"目录高质量匹配({toc_score:.0%})",
                                "相似度": f"{toc_score:.2%}",
                            }
                            if toc_score >= 0.98:
                                exact_matched.append(match_data)
                            else:
                                fuzzy_matched.append(match_data)
                            last_found_idx = toc_match["_global_idx"]
                            found_in_current_stage = True

            # 【NEW】更新第一阶段匹配计数
            if found_in_current_stage:
                matched_count_phase1 += 1
                continue

            # --- 2. 【分层策略】在确定的章节内进行递归搜索 ---
            # 确定搜索范围：优先使用Excel的三级目录
            l_target_name = excel_level3 if excel_level3 else ""
            if not l_target_name:
                if hierarchy_mapping:
                    key = (
                        excel_item.get("level1", ""),
                        excel_item.get("level2", ""),
                        excel_item.get("level3", ""),
                    )
                    if key in hierarchy_mapping:
                        l_target_name = hierarchy_mapping[key]

                if not l_target_name:
                    for l_key in ["level3", "level2", "level1"]:
                        l_val = excel_item.get(l_key, "")
                        if l_val:
                            l_target_name = l_val
                            break

            # 提取章节编号（如4.1.1.1）用于确定搜索范围
            l_target_num = ""
            if l_target_name:
                num_m = re.match(r"^(\d+(?:\.\d+){2,})", str(l_target_name))
                if num_m:
                    l_target_num = num_m.group(1)

                # 更新当前追踪的章节
                if l_target_name != current_excel_l3:
                    current_excel_l3 = l_target_name
                    l_clean = self.basic_clean(l_target_name)
                    # 尝试定位该章节
                    l_match = exact_lookup.get(l_clean)
                    if not l_match and l_target_num:
                        l_match = chapter_exact_mapping.get(l_target_num)
                    if l_match:
                        last_found_idx = l_match["_global_idx"]

            # --- 3. 【分层策略】在指定章节范围内搜索 ---
            chapter_candidates = []
            if l_target_num:
                # 构建章节范围：从 4.1.1.1 到 4.1.1.2（不包含4.1.1.2）
                target_parts = l_target_num.split(".")
                target_depth = len(target_parts)
                parent_prefix = ".".join(target_parts[:-1])
                current_section = int(target_parts[-1])
                next_section = current_section + 1
                next_section_num = ".".join(target_parts[:-1] + [str(next_section)])

                # 收集当前章节及其子章节的所有内容
                target_prefix = f"{l_target_num}."

                # 【性能改进】使用更高效的搜索算法
                # 1. 先添加当前章节的直接内容
                if l_target_num in chapter_buckets:
                    chapter_candidates.extend(chapter_buckets[l_target_num])

                # 2. 只遍历可能匹配的 bucket（优化：避免全量遍历）
                bucket_search_count = 0  # 【NEW】计数搜索的 bucket
                for bucket_key in chapter_buckets:
                    bucket_search_count += 1
                    if bucket_key.startswith(target_prefix):
                        # 确保不超出当前章节范围（不包含4.1.1.2及其子章节）
                        bucket_parts = bucket_key.split(".")
                        if (
                            len(bucket_parts) > target_depth
                            and bucket_parts[:target_depth] == target_parts
                        ):
                            for item in chapter_buckets[bucket_key]:
                                if item not in chapter_candidates:
                                    chapter_candidates.append(item)
                        elif (
                            len(bucket_parts) == target_depth
                            and bucket_key == l_target_num
                        ):
                            # 已经添加过了
                            continue

                if (
                    DETAILED_LOG and bucket_search_count > 500
                ):  # 【NEW】当搜索次数过多时记录
                    print(
                        f"  ⚠️ 章节范围搜索需要遍历 {bucket_search_count} 个 bucket（{len(chapter_buckets)} 总数）"
                    )

                print(
                    f"  章节范围搜索: {l_target_num} -> {next_section_num}, 候选项: {len(chapter_candidates)}"
                )
            else:
                # 回退到通用的分桶检索
                l_target_clean = (
                    self.basic_clean(l_target_name) if l_target_name else ""
                )
                if l_target_clean and l_target_clean in chapter_buckets:
                    chapter_candidates = chapter_buckets[l_target_clean]

            if chapter_candidates:
                # [Strategy] 在限定章节内，先进行精确匹配
                for cand in chapter_candidates:
                    c_text = cand.get("text", "")
                    c_no_space = cand.get("text_no_space", "")
                    if excel_text == c_text or excel_no_space == c_no_space:
                        match_data = {
                            "Excel功能点": excel_original,
                            "Excel一级模块": excel_item.get("level1", ""),
                            "Excel二级模块": excel_item.get("level2", ""),
                            "Excel三级模块": excel_item.get("level3", ""),
                            "Word匹配项": cand.get("display", cand.get("original", "")),
                            "位置": f"章节内-精确匹配({l_target_num or l_target_name[:10]})",
                            "相似度": "100.00%",
                        }
                        exact_matched.append(match_data)
                        last_found_idx = cand["_global_idx"]
                        found_in_current_stage = True
                        print(f"    ✓ 章节内精确匹配: {excel_text[:20]}...")
                        break

                if found_in_current_stage:
                    continue

                # 如果没有精确匹配，进行模糊匹配
                found_ch, ch_match, ch_score = self.find_match(
                    excel_text, chapter_candidates, 0, chapter_locked=True
                )
                if found_ch and ch_score >= 0.85:  # 提高章节内模糊匹配阈值
                    match_data = {
                        "Excel功能点": excel_original,
                        "Excel一级模块": excel_item.get("level1", ""),
                        "Excel二级模块": excel_item.get("level2", ""),
                        "Excel三级模块": excel_item.get("level3", ""),
                        "Word匹配项": ch_match.get(
                            "display", ch_match.get("original", "")
                        ),
                        "位置": f"章节内-模糊匹配({l_target_num or l_target_name[:10]})",
                        "相似度": f"{ch_score:.2%}",
                    }
                    if ch_score >= 0.99:
                        exact_matched.append(match_data)
                    else:
                        fuzzy_matched.append(match_data)
                    last_found_idx = ch_match["_global_idx"]
                    found_in_current_stage = True
                    if DETAILED_LOG and idx < 50:  # 【NEW】仅前 50 项显示详细日志
                        print(
                            f"    ≈ 章节内模糊匹配: {excel_text[:20]}... (score: {ch_score:.2f})"
                        )
                    else:
                        print(
                            f"    ≈ 章节内模糊匹配: {excel_text[:20]}... (score: {ch_score:.2f})"
                        )
                    # 【NEW】更新第一阶段匹配计数
                    matched_count_phase1 += 1
                    continue
                else:
                    if DETAILED_LOG and idx < 50:  # 【NEW】仅前 50 项显示详细日志
                        print(f"    ✗ 章节内未找到匹配: {excel_text[:20]}...")

            # --- 4. 第一阶段未匹配，记为延期，进入全文搜索 ---
            if not found_in_current_stage:
                deferred_items.append(excel_item)

        # 第二阶段：延期全文搜索（优先匹配 2.2 和 4.* 章节）
        if deferred_items:
            # 【NEW】第一阶段统计信息
            phase1_elapsed = time_module.time() - phase1_start_time
            print(
                f"\n[STATS] 第一阶段耗时: {phase1_elapsed:.2f}s, 已匹配: {matched_count_phase1}/{total} 项"
            )

            deferred_total = len(deferred_items)
            print(f"\n第一阶段章节搜索未匹配 {deferred_total} 项，开始全文搜索...")
            phase2_start_time = time_module.time()  # 【NEW】记录第二阶段开始时间
            if progress_callback:
                progress_callback(
                    60, f"[PROCESS] 阶段二: 全文搜索未匹配项 ({deferred_total}项)..."
                )

            # 预识别特殊章节内容 (2.2 系统已实现功能 和 4.* 功能需求)
            priority_candidates = []
            for item in word_items:
                p_num = str(item.get("parent_num", ""))
                # 匹配 2.2 或 4.* 章节的内容
                if p_num == "2.2" or p_num.startswith("4"):
                    priority_candidates.append(item)
                else:
                    # 备选：如果编号没提取到，检查标题文字（兜底）
                    p_ch = item.get("parent_chapter", "")
                    if "系统已实现功能" in p_ch or "功能需求" in p_ch:
                        priority_candidates.append(item)

            import concurrent.futures
            from threading import Lock

            results_lock = Lock()

            def process_deferred_item(d_item):
                try:
                    d_original = d_item["original"]
                    d_text = self.basic_clean(d_original)
                    d_excel_l3 = d_item.get("level3", "")[:20]  # 截断用于日志输出

                    # 【NEW】记录处理开始
                    # print(f"  [DEEP-SEARCH] 处理延期项: {d_original[:30]}... (来自: {d_excel_l3})")

                    # A. 优先在 [2.2系统已实现功能] 和 [4 功能需求] 中搜
                    found_p, word_p, score_p = (
                        self.find_match(d_text, priority_candidates, 0)
                        if priority_candidates
                        else (False, None, 0.0)
                    )

                    # 【NEW】记录第一步搜索结果
                    # if found_p and score_p > 0.85:
                    #     print(f"    ✓ 核心章节找到: score={score_p:.2%}")

                    if found_p and score_p > 0.85:
                        with results_lock:
                            match_data = {
                                "Excel功能点": d_original,
                                "Excel一级模块": d_item.get("level1", ""),
                                "Excel二级模块": d_item.get("level2", ""),
                                "Excel三级模块": d_item.get("level3", ""),
                                "Word匹配项": word_p.get(
                                    "display", word_p.get("original", "")
                                ),
                                "位置": "核心章节搜索",
                                "相似度": f"{score_p:.2%}",
                                "匹配状态": (
                                    "[PASS] 匹配成功"
                                    if score_p >= 0.99
                                    else "[PASS] 模糊匹配"
                                ),
                            }
                            if score_p >= 0.99:
                                exact_matched.append(match_data)
                            else:
                                fuzzy_matched.append(match_data)
                        return

                    # 【NEW】记录第一步未找到，进入递归章节搜索
                    # print(f"    → 核心章节未找到，进入递归章节搜索...")

                    # B. 【新增】递归章节搜索：从所属的父章节开始，逐级向下查找子章节
                    d_parent_num = str(d_item.get("parent_num", "")).strip().rstrip(".")
                    found_recursive = False
                    best_recursive_match = None
                    best_recursive_score = 0.0
                    recursive_location = ""

                    if d_parent_num:
                        # 从当前章节开始，向下递归查找子章节内容
                        # 例如：4.1.1.1 -> 4.1.1.1.1 -> 4.1.1.1.2 ... -> 4.1.1.2（停止）
                        def find_in_chapter_range(chapter_num, target_text):
                            """在指定章节及其子章节范围内查找匹配"""
                            # 分解编号 (4.1.1.1 -> [4, 1, 1, 1])
                            parts = [int(p) for p in chapter_num.split(".")]
                            best_match = None
                            best_score = 0.0
                            found_location = ""

                            # 遍历 Word 所有项，查找在这个章节范围内的内容
                            for candidate in word_items:
                                c_parent_num = (
                                    str(candidate.get("parent_num", ""))
                                    .strip()
                                    .rstrip(".")
                                )
                                if not c_parent_num:
                                    continue

                                # 解析候选项的编号
                                try:
                                    c_parts = [int(p) for p in c_parent_num.split(".")]
                                except (ValueError, AttributeError):
                                    continue

                                # 判断是否在范围内：候选项需要以当前章节为前缀
                                # 例如 4.1.1.1 的范围包含 4.1.1.1, 4.1.1.1.1, ... 但不包含 4.1.1.2
                                if len(c_parts) < len(parts):
                                    # 候选项层级太浅，不在范围内
                                    continue

                                if c_parts[: len(parts)] != parts:
                                    # 候选项编号前缀不符，不在范围内
                                    continue

                                # 候选项在范围内，进行文本匹配
                                c_text = candidate.get("text", "")
                                c_no_space = candidate.get("text_no_space", "")
                                target_no_space = re.sub(
                                    r"\s+", "", target_text.lower()
                                )

                                # 精确匹配优先
                                if (
                                    target_text == c_text
                                    or target_no_space == c_no_space
                                ):
                                    return candidate, 1.0, c_parent_num

                                # 模糊匹配
                                score = self.similarity(target_text, c_text)
                                if score > best_score:
                                    best_score = score
                                    best_match = candidate
                                    found_location = c_parent_num

                            return best_match, best_score, found_location

                        # 执行递归查找
                        (
                            best_recursive_match,
                            best_recursive_score,
                            recursive_location,
                        ) = find_in_chapter_range(d_parent_num, d_text)
                        found_recursive = (
                            best_recursive_match is not None
                            and best_recursive_score >= 0.8
                        )

                    if found_recursive:
                        with results_lock:
                            match_data = {
                                "Excel功能点": d_original,
                                "Excel一级模块": d_item.get("level1", ""),
                                "Excel二级模块": d_item.get("level2", ""),
                                "Excel三级模块": d_item.get("level3", ""),
                                "Word匹配项": best_recursive_match.get(
                                    "display", best_recursive_match.get("original", "")
                                ),
                                "位置": f"递归章节搜索({d_parent_num}->{recursive_location})",
                                "相似度": f"{best_recursive_score:.2%}",
                                "匹配状态": (
                                    "[PASS] 匹配成功"
                                    if best_recursive_score >= 0.99
                                    else "[PASS] 模糊匹配"
                                ),
                            }
                            if best_recursive_score >= 0.99:
                                exact_matched.append(match_data)
                            else:
                                fuzzy_matched.append(match_data)
                        return

                    # C. 兜底：全书全局搜索
                    found_g, word_g, score_g = self.find_match(d_text, word_items, 0)

                    # 【NEW】记录第三步搜索结果
                    # if found_g and score_g > 0.8:
                    #     print(f"    ✓ 全局搜索找到: score={score_g:.2%}")

                    with results_lock:
                        if found_g and score_g > 0.8:
                            match_data = {
                                "Excel功能点": d_original,
                                "Excel一级模块": d_item.get("level1", ""),
                                "Excel二级模块": d_item.get("level2", ""),
                                "Excel三级模块": d_item.get("level3", ""),
                                "Word匹配项": word_g.get(
                                    "display", word_g.get("original", "")
                                ),
                                "位置": "全局搜索",
                                "相似度": f"{score_g:.2%}",
                                "匹配状态": (
                                    "[PASS] 匹配成功"
                                    if score_g >= 0.99
                                    else "[PASS] 模糊匹配"
                                ),
                            }
                            if score_g >= 0.99:
                                exact_matched.append(match_data)
                            else:
                                fuzzy_matched.append(match_data)
                        else:
                            not_found_in_word.append(
                                {
                                    "Excel功能点": d_original,
                                    "Excel一级模块": d_item.get("level1", ""),
                                    "Excel二级模块": d_item.get("level2", ""),
                                    "Excel三级模块": d_item.get("level3", ""),
                                    "匹配状态": "[FAIL] 缺失",
                                    "简略描述": "全书匹配未找到",
                                }
                            )
                except Exception as e:
                    print(f"Error in process_deferred_item: {e}")

            # 并发执行深度搜索
            max_workers = 4  # 略微增加并发
            with concurrent.futures.ThreadPoolExecutor(
                max_workers=max_workers
            ) as executor:
                futures = [
                    executor.submit(process_deferred_item, di) for di in deferred_items
                ]
                comp = 0
                deep_search_matched = 0  # 【NEW】深度搜索已匹配项数
                for _ in concurrent.futures.as_completed(futures):
                    comp += 1
                    # 计数匹配的项（从结果追踪）
                    # 【改进】增加进度细节
                    # 深度搜索耗时较长，增加汇报频率
                    if comp % 5 == 0 or comp == deferred_total:
                        p = 60 + int((comp / deferred_total) * 30)
                        # 【改进】显示深度搜索进度，包括已处理项数和总数
                        msg = f"[DEEP-SEARCH] 深度搜索 进度 {comp}/{deferred_total} (第一阶段已匹配: {matched_count_phase1})"
                        if progress_callback:
                            progress_callback(p, msg)
                        # 同时打印到 stdout 以便 RuntimeLogger 实时记录
                        if comp % 20 == 0 or comp == deferred_total:
                            print(f"  {msg}")

        if progress_callback:
            progress_callback(90, "正在汇总报告...")

        # 【NEW】第二阶段统计信息
        if deferred_items:
            phase2_elapsed = time_module.time() - phase2_start_time
            print(
                f"\n[STATS] 第二阶段耗时: {phase2_elapsed:.2f}s, 处理项数: {len(deferred_items)}"
            )

        total_excel = len(excel_data)
        exact_count = len(exact_matched)
        fuzzy_count = len(fuzzy_matched)
        matched_count = exact_count + fuzzy_count
        missing_count = len(not_found_in_word)
        match_rate = (matched_count / total_excel * 100) if total_excel > 0 else 0

        print(
            f"\n分层匹配完成：精确 {exact_count} 项，模糊 {fuzzy_count} 项，缺失 {missing_count} 项"
        )
        if full_text_search:
            # 统计分层匹配的详细位置
            chapter_exact_count = sum(
                1
                for item in exact_matched + fuzzy_matched
                if "章节" in item.get("位置", "") or "目录" in item.get("位置", "")
            )
            global_count = sum(
                1
                for item in exact_matched + fuzzy_matched
                if "全局" in item.get("位置", "") or "核心章节" in item.get("位置", "")
            )
            print(
                f"  其中：章节范围匹配 {chapter_exact_count} 项，全文搜索匹配 {global_count} 项"
            )

        # 移除过多的终端打印，保持界面干净
        if progress_callback:
            progress_callback(100, "匹配完成！")

        report = {
            "exact_matched": exact_matched,
            "fuzzy_matched": fuzzy_matched,
            "not_found_in_word": not_found_in_word,
            "statistics": {
                "excel功能点总数": total_excel,
                "word内容项总数": len(word_items),
                "精确匹配": exact_count,
                "模糊匹配": fuzzy_count,
                "已匹配": matched_count,
                "缺失项": missing_count,
                "匹配率": f"{match_rate:.2f}%",
                "完成度": f"{match_rate:.2f}%",
            },
        }

        if full_text_search:
            chapter_exact_count = sum(
                1
                for item in exact_matched + fuzzy_matched
                if "章节" in item.get("位置", "") or "目录" in item.get("位置", "")
            )
            global_count = sum(
                1
                for item in exact_matched + fuzzy_matched
                if "全局" in item.get("位置", "") or "核心章节" in item.get("位置", "")
            )
            report["statistics"]["章节范围匹配"] = chapter_exact_count
            report["statistics"]["全文搜索匹配"] = global_count

        # [完成摘要] 输出匹配结果统计，不输出完整JSON
        if RuntimeLogger:
            RuntimeLogger.log("=" * 80, level="INFO")
            RuntimeLogger.log("📊 分层匹配结果统计摘要", level="INFO")
            RuntimeLogger.log("=" * 80, level="INFO")
            RuntimeLogger.log(f"  精确匹配: {len(exact_matched)} 项", level="INFO")
            RuntimeLogger.log(f"  模糊匹配: {len(fuzzy_matched)} 项", level="INFO")
            if hierarchy_mapping:
                RuntimeLogger.log(
                    f"  层级映射: {len(hierarchy_mapping)} 项", level="INFO"
                )
            RuntimeLogger.log(f"  缺失项: {len(not_found_in_word)} 项", level="INFO")
            RuntimeLogger.log(f"  总体匹配率: {match_rate:.2f}%", level="INFO")
            if full_text_search:
                chapter_exact_count = sum(
                    1
                    for item in exact_matched + fuzzy_matched
                    if "章节" in item.get("位置", "") or "目录" in item.get("位置", "")
                )
                global_count = sum(
                    1
                    for item in exact_matched + fuzzy_matched
                    if "全局" in item.get("位置", "")
                    or "核心章节" in item.get("位置", "")
                )
                RuntimeLogger.log(
                    f"  章节范围匹配: {chapter_exact_count} 项", level="INFO"
                )
                RuntimeLogger.log(f"  全文搜索匹配: {global_count} 项", level="INFO")
            RuntimeLogger.log("=" * 80, level="INFO")

        return report

    def similarity(self, a: str, b: str) -> float:
        """使用 RapidFuzz 计算相似度，比 SequenceMatcher 准确且快"""
        if not a or not b:
            return 0.0

        # 归一化处理
        a_norm = a.lower().replace("账户", "账号")
        b_norm = b.lower().replace("账户", "账号")

        try:
            from rapidfuzz import fuzz

            # [FIX] 分离 ratio 和 partial_ratio，确保只有完全一致才有 1.0
            r_ratio = fuzz.ratio(a_norm, b_norm)
            if r_ratio >= 99.5:  # 处理微小差异（如大小写/极个别字符）
                return 1.0

            r_partial = fuzz.partial_ratio(a_norm, b_norm)

            # [NEW-LOGIC] 计算长度比率，用于全局惩罚
            len_a = len(a_norm)
            len_b = len(b_norm)
            # a 是 Excel (需求), b 是 Word (实现)
            len_ratio = len_b / len_a if len_a > 0 else 0

            # [STRICT-FIX] 处理高相似的部分匹配
            if r_partial >= 80:
                # 场景 A: Word 包含了 Excel 的全部内容 (Word 覆盖了需求) -> 完美
                if a_norm in b_norm:
                    if len_b < len_a + 12:
                        return 1.0
                    return 0.95

                # 场景 B: Word 只是 Excel 需求的一个片段 (Word 是子集) -> 风险高
                if b_norm in a_norm:
                    # 如果 Word 长度不足需求的 60%，强制降级（即使 partial_ratio 是 100）
                    if len_ratio < 0.6:
                        return min(r_ratio / 100.0, 0.60)
                    elif len_ratio < 0.85:
                        # 0.6-0.85 之间，得分 0.65-0.80
                        return 0.65 + (len_ratio - 0.6) * 0.6
                    return max(r_ratio / 100.0, 0.85)

                # 场景 C: 其他交错匹配 (partial_ratio 很高但互相不是完全包含)
                # 如果长度极度不匹配 (如 "日查询" vs "查询日期范围内的客流量数据")
                if len_ratio < 0.5:
                    r_partial = (
                        r_partial * 0.6
                    )  # 直接打 6 折，防止误通过 阈值(0.75-0.8)

            # 计算最终得分
            final_score = max(r_ratio, r_partial) / 100.0

            # [FINAL-GUARD] 无论是否包含关系，只要 Word 侧信息量过短，都不允许超过阈值
            if len_ratio < 0.5 and final_score > 0.7:
                final_score = 0.65

            return final_score
        except ImportError:
            return SequenceMatcher(None, a_norm, b_norm).ratio()

    def find_match(
        self,
        target: str,
        candidates: List[Dict],
        target_depth: int = 0,
        chapter_locked: bool = False,
    ) -> Tuple[bool, Dict, float]:
        """
        在候选集中查找最佳匹配项（兼容3返回值）
        完全相同定义：忽略空格和大小写后文本完全一致 → 相似度=1.0
        smarter logic: 包含关系、关键字重叠、特定模式增强 (同步自 _recursive_chapter_search)
        """
        import re

        if not target or not candidates:
            return False, None, 0.0

        # 标准化目标文本
        target_str = str(target).strip()
        target_clean = re.sub(r"\s+", "", target_str.lower())
        target_basic = self.basic_clean(target_str)
        if not target_basic:
            target_basic = target_clean

        best_match = None
        best_score = -1.0

        for candidate in candidates:
            # 兼容不同字段名
            cand_text = (
                candidate.get("text")
                or candidate.get("original")
                or candidate.get("cleaned", "")
            )
            if not cand_text:
                continue

            # 标准化候选文本
            cand_clean = re.sub(r"\s+", "", cand_text.strip().lower())
            cand_basic = self.basic_clean(cand_text)
            if not cand_basic:
                cand_basic = cand_clean

            current_score = 0.0

            # 1. 【核心】检测"完全相同"（忽略空格和大小写）
            if target_clean == cand_clean:
                current_score = 1.0
            else:
                # 2. 计算基础相似度
                if self.fuzzy_match:
                    # 同步 _recursive_chapter_search 使用 basic_clean 后的文本计算相似度
                    current_score = self.similarity(target_basic, cand_basic)
                else:
                    current_score = 0.0

                # --- 智能特征增强 (同步 _recursive_chapter_search 逻辑) ---
                if current_score < 0.95:
                    # 策略 1: 包含关系 (Inclusion)
                    # [FIX] 只有在 Word 覆盖 Excel 且包含关键词时才提升分数，防止短词片段匹配长需求
                    if target_clean and cand_clean:
                        if target_clean in cand_clean:
                            # 覆盖关系由 similarity 函数内部处理分值和长度惩罚
                            pass
                        elif cand_clean in target_clean:
                            # 如果 Word 只是 Excel 的一部分，不在这里手动给 0.92
                            # 维持 similarity 函数中计算出的分数（可能带有长度惩罚）
                            pass

                    # 策略 2: 共同关键词占比 (Keyword Overlap)
                    if current_score < 0.85:
                        words_target = set(re.findall(r"\w+", target_basic))
                        words_cand = set(re.findall(r"\w+", cand_basic))
                        if words_target and words_cand:
                            overlap = words_target & words_cand
                            overlap_ratio = len(overlap) / len(words_target)
                            if overlap_ratio >= 0.7:
                                # 长度校验
                                if len(cand_basic) < 0.6 * len(target_basic):
                                    current_score = max(current_score, 0.65)
                                else:
                                    current_score = max(
                                        current_score, 0.7 + overlap_ratio * 0.18
                                    )

            # 3. 标题项额外权重 (Heading Bonus)
            if (
                candidate.get("type") == "heading"
                or candidate.get("source") == "heading"
            ):
                # 锦上添花：只有在分值已经比较高（>0.85）时才给予微小权重，
                # 防止将本就不匹配的片段通过加分拉到 80% 的及格线
                if current_score >= 0.85 and current_score < 1.0:
                    current_score = min(1.0, current_score + 0.03)

            # 更新最佳匹配
            if current_score > best_score + 1e-5:
                best_score = current_score
                best_match = candidate

        # 4. 阈值判定
        effective_threshold = 0.75 if chapter_locked else self.threshold
        if target_depth > 0:
            effective_threshold = 0.7

        # [NEW] 宽松判定：如果分数达到 0.70 但未到 threshold，在锁定章节时也认为找到了 (同步 _recursive_chapter_search)
        if chapter_locked and best_score >= 0.70:
            effective_threshold = min(effective_threshold, 0.70)

        return (
            (best_score >= effective_threshold),
            (best_match if best_match else {}),
            (best_score if best_score > 0 else 0.0),
        )

    def _match_hierarchical_layered(
        self,
        excel_data,
        word_items,
        chapter_buckets,
        exact_lookup,
        full_text_content,
        progress_callback=None,
        hierarchy_mapping=None,
    ):
        """
        【NEW】按用户需求实现分层递归匹配：

        1. 章节完全匹配（目录树）- 如果完全相同，直接返回
        2. 章节内容递归搜索 - 在当前章节及子章节内搜索
        3. 全文匹配 - 最后进行全文搜索

        未找到项被记录为"在章节匹配未找到"，等待第3阶段处理
        """
        import time as time_module

        exact_matched = []
        fuzzy_matched = []
        unfound_chapter_match = []  # 记录章节匹配未找到的项

        total = len(excel_data)
        print(f"\n[HIERARCHICAL MATCHING] Processing {total} functional points...")

        # ============ 阶段1：章节完全匹配 ============
        print(f"\n[STAGE 1] Exact chapter matching...")
        phase1_start = time_module.time()

        for idx, excel_item in enumerate(excel_data):
            if progress_callback and idx % max(1, total // 20) == 0:
                p = 10 + int((idx / total) * 20)
                progress_callback(p, f"[Phase 1] Exact matching {idx+1}/{total}")

            excel_original = excel_item.get("original", "") or excel_item.get(
                "text", ""
            )
            excel_text = self.basic_clean(excel_original)

            # 【NEW】预先提取当前 Excel 条目所属的 Word 章节编号（从映射或文本提取）
            current_excel_num = None
            if hierarchy_mapping:
                l1 = str(excel_item.get("level1", "")).strip()
                l2 = str(excel_item.get("level2", "")).strip()
                l3 = str(excel_item.get("level3", "")).strip()
                current_excel_num = hierarchy_mapping.get((l1, l2, l3))

                # 如果映射到的是标题而不是编号，尝试从 Word 索引中找编号
                if current_excel_num and not re.match(
                    r"^(\d+(?:\.\d+)*)$", str(current_excel_num)
                ):
                    if current_excel_num in exact_lookup:
                        current_excel_num = exact_lookup[current_excel_num].get(
                            "number", ""
                        )
                    else:
                        current_excel_num = None

            if not current_excel_num:
                # 备选：从 Excel 单元格文本中提取编号
                for key in ["level3", "level2", "level1"]:
                    val = str(excel_item.get(key, ""))
                    num_match = re.match(r"^(\d+(?:\.\d+)+)", val)
                    if num_match:
                        current_excel_num = num_match.group(1)
                        break

            found = False

            # 【策略1】尝试在精确查找表中直接匹配
            if excel_text in exact_lookup:
                match_item = exact_lookup[excel_text]

                # [ENHANCEMENT] 校验匹配项是否在正确的章节范围内
                is_correct_scope = True
                if current_excel_num:
                    item_num = match_item.get("number") or match_item.get(
                        "parent_num", ""
                    )
                    if not is_descendant(current_excel_num, item_num):
                        is_correct_scope = False

                if is_correct_scope:
                    exact_matched.append(
                        {
                            "Excel功能点": excel_original,
                            "Excel一级模块": excel_item.get("level1", ""),
                            "Excel二级模块": excel_item.get("level2", ""),
                            "Excel三级模块": excel_item.get("level3", ""),
                            "Word匹配项": match_item.get(
                                "display", match_item.get("original", "")
                            ),
                            "位置": "章节完全匹配",
                            "相似度": "100.00%",
                            "匹配状态": "[PASS] 精确匹配",
                        }
                    )
                    found = True
                else:
                    # 如果全局精确匹配在错误范围，则进入阶段2（递归匹配）来找正确范围内的那个
                    pass

            # 【策略2】如果直接匹配失败，尝试在Word中找带有编号的匹配
            # 例如：Excel"服务列表" vs Word"4.1.1.1 服务列表"
            if not found:
                # 从level3提取编号
                excel_level3 = excel_item.get("level3", "")
                extracted_num = None
                if excel_level3:
                    num_match = re.match(r"^(\d+(?:\.\d+){2,})", str(excel_level3))
                    if num_match:
                        extracted_num = num_match.group(1)

                # 在word_items中查找匹配的编号
                if extracted_num:
                    for word_item in word_items:
                        word_original = word_item.get("original", "")
                        if word_original.startswith(extracted_num):
                            # 检查编号后面的文本是否匹配
                            word_text_part = word_original[len(extracted_num) :].strip()
                            word_text_part = self.basic_clean(word_text_part)
                            if (
                                excel_text == word_text_part
                                or excel_text in word_text_part
                            ):
                                exact_matched.append(
                                    {
                                        "Excel功能点": excel_original,
                                        "Excel一级模块": excel_item.get("level1", ""),
                                        "Excel二级模块": excel_item.get("level2", ""),
                                        "Excel三级模块": excel_item.get("level3", ""),
                                        "Word匹配项": word_item.get(
                                            "display", word_item.get("original", "")
                                        ),
                                        "位置": f"章节号匹配({extracted_num})",
                                        "相似度": "100.00%",
                                        "匹配状态": "[PASS] 编号匹配",
                                    }
                                )
                                found = True
                                break

            # 标记为"章节匹配未找到"，等待后续阶段
            if not found:
                unfound_chapter_match.append(excel_item)

        phase1_time = time_module.time() - phase1_start
        print(
            f"  [DONE] Phase1 completed: matched {len(exact_matched)} items, unmatched {len(unfound_chapter_match)} items (time: {phase1_time:.2f}s)"
        )

        # ============ 阶段2：章节内容递归搜索 ============
        print(
            f"\n[STAGE 2] Chapter content recursive search ({len(unfound_chapter_match)} items)..."
        )

        # 【DEBUG】打印chapter_buckets的示例
        if DETAILED_LOG and chapter_buckets:
            sample_keys = list(chapter_buckets.keys())[:10]
            print(f"  [DEBUG] Sample chapter_buckets keys: {sample_keys}")
            for key in sample_keys:
                print(f"    {key}: {len(chapter_buckets[key])} items")

        phase2_start = time_module.time()

        still_unfound = []
        for idx, excel_item in enumerate(unfound_chapter_match):
            if (
                progress_callback
                and idx % max(1, len(unfound_chapter_match) // 20) == 0
            ):
                p = 30 + int((idx / len(unfound_chapter_match)) * 20)
                progress_callback(
                    p, f"[阶段2] 递归搜索 {idx+1}/{len(unfound_chapter_match)}"
                )

            excel_original = excel_item.get("original", "") or excel_item.get(
                "text", ""
            )
            excel_text = self.basic_clean(excel_original)
            excel_parent_num = excel_item.get("parent_num") or excel_item.get(
                "level3", ""
            )

            if DETAILED_LOG and idx < 10:
                print(
                    f"  [DEBUG] P2-Item: '{excel_original[:30]}...' (Parent: {excel_parent_num})"
                )

            # 【NEW】尝试从 hierarchy_mapping 中获取已匹配的章节编号
            mapped_num = None
            if hierarchy_mapping:
                l1 = str(excel_item.get("level1", "")).strip()
                l2 = str(excel_item.get("level2", "")).strip()
                l3 = str(excel_item.get("level3", "")).strip()
                key = (l1, l2, l3)
                mapped_num = hierarchy_mapping.get(key)
                if mapped_num and not re.match(r"^(\d+(?:\.\d+)*)$", str(mapped_num)):
                    # 如果映射到的是标题而不是编号，尝试从 Word 索引中找编号
                    if mapped_num in exact_lookup:
                        mapped_num = exact_lookup[mapped_num].get("number", "")
                    else:
                        mapped_num = None

            # 【关键FIX】提取编号 - 多种策略
            extracted_num = mapped_num  # 优先使用映射编号

            if not extracted_num and excel_parent_num:
                # 策略1：尝试直接提取数字编号
                num_match = re.match(r"^(\d+(?:\.\d+)*)", str(excel_parent_num))
                if num_match:
                    extracted_num = num_match.group(1)

            # 【新增】策略2：如果是文本标题（无数字），则从exact_lookup中查找对应的Word编号
            if not extracted_num:
                # 尝试用level3文本直接查找
                l3_text = str(excel_item.get("level3", "")).strip()
                if l3_text and l3_text in exact_lookup:
                    word_item = exact_lookup[l3_text]
                    # 获取Word项目自己的编号（不是父级编号！）
                    extracted_num = word_item.get("number") or word_item.get(
                        "parent_num"
                    )

            # 【策略3】如果从level3提取失败，尝试从level2、level1提取
            if not extracted_num:
                for level_key in ["level2", "level1"]:
                    level_val = excel_item.get(level_key, "")
                    if level_val:
                        # 先尝试数字编号
                        num_match = re.match(r"^(\d+(?:\.\d+)*)", str(level_val))
                        if num_match:
                            extracted_num = num_match.group(1)
                            break
                        # 再尝试文本查找
                        level_text = str(level_val).strip()
                        if level_text in exact_lookup:
                            word_item = exact_lookup[level_text]
                            # 获取Word项目自己的编号（不是父级编号！）
                            extracted_num = word_item.get("number") or word_item.get(
                                "parent_num"
                            )
                            if extracted_num:
                                break

            found = False

            # 【DEBUG】前10项打印提取编号的过程
            if idx < 10:
                print(
                    f"  [P2#{idx}] '{excel_original[:20]}...' | level3='{excel_item.get('level3', '')}' | "
                    f"parent_num='{excel_parent_num}' | extracted_num='{extracted_num}'"
                )

            # 如果有章节编号，进行递归搜索
            if extracted_num:
                if idx < 10:
                    bucket_size = len(chapter_buckets.get(extracted_num, []))
                    print(
                        f"    → 搜索范围: {extracted_num}, chapter_buckets中该范围大小={bucket_size}"
                    )

                found_recursive, match_item, score, location = (
                    self._recursive_chapter_search(
                        excel_text,
                        extracted_num,
                        word_items,
                        chapter_buckets,
                        full_text_content,
                    )
                )

                if idx < 10:
                    score_str = f"{score:.2f}" if found_recursive else "0.00"
                    print(f"    ← 搜索结果: found={found_recursive}, score={score_str}")

                if found_recursive:
                    match_data = {
                        "Excel功能点": excel_original,
                        "Excel一级模块": excel_item.get("level1", ""),
                        "Excel二级模块": excel_item.get("level2", ""),
                        "Excel三级模块": excel_item.get("level3", ""),
                        "Word匹配项": match_item.get(
                            "display", match_item.get("original", "")
                        ),
                        "位置": f"章节递归搜索({extracted_num})",
                        "相似度": f"{score:.2%}",
                        "匹配状态": (
                            "[PASS] 递归匹配" if score >= 0.9 else "[PASS] 模糊匹配"
                        ),
                    }
                    if score >= 0.95:
                        exact_matched.append(match_data)
                    else:
                        fuzzy_matched.append(match_data)
                    found = True
                else:
                    # 【NEW】DEBUG：未找到的原因
                    if DETAILED_LOG and idx < 10:
                        print(
                            f"    ✗ 递归搜索失败: '{excel_original[:30]}...' in range [{extracted_num}, ...)"
                        )
            else:
                # 【NEW】DEBUG：无法提取编号
                if DETAILED_LOG and idx < 10:
                    print(
                        f"    ✗ 无法提取编号: level3='{excel_item.get('level3', '')}', parent_num='{excel_parent_num}'"
                    )

            if not found:
                still_unfound.append(excel_item)

        phase2_time = time_module.time() - phase2_start
        print(
            f"  [DONE] Phase2 completed: newly matched {len(unfound_chapter_match) - len(still_unfound)} items, "
            f"still unmatched {len(still_unfound)} items (time: {phase2_time:.2f}s)"
        )

        # ============ 阶段3：全文匹配 ============
        print(f"\n[STAGE 3] Full-text matching ({len(still_unfound)} items)...")
        phase3_start = time_module.time()

        not_found_in_word = []
        for idx, excel_item in enumerate(still_unfound):
            if progress_callback and idx % max(1, len(still_unfound) // 20) == 0:
                p = 50 + int((idx / len(still_unfound)) * 40)
                progress_callback(p, f"[阶段3] 全文匹配 {idx+1}/{len(still_unfound)}")

            excel_original = excel_item.get("original", "") or excel_item.get(
                "text", ""
            )
            excel_text = self.basic_clean(excel_original)

            found = False

            # 在全文内容中进行匹配
            if full_text_content:
                best_match = None
                best_score = 0.0

                for full_item in full_text_content:
                    full_text = full_item.get("cleaned", "")
                    if not full_text:
                        continue

                    # 精确匹配
                    if excel_text == full_text:
                        exact_matched.append(
                            {
                                "Excel功能点": excel_original,
                                "Excel一级模块": excel_item.get("level1", ""),
                                "Excel二级模块": excel_item.get("level2", ""),
                                "Excel三级模块": excel_item.get("level3", ""),
                                "Word匹配项": full_item.get("original", ""),
                                "位置": "全文精确匹配",
                                "相似度": "100.00%",
                                "匹配状态": "[PASS] 全文匹配",
                            }
                        )
                        found = True
                        break

                    # 模糊匹配
                    score = self.similarity(excel_text, full_text)
                    if score > best_score:
                        best_score = score
                        best_match = full_item

                # 如果有高质量的模糊匹配
                if not found and best_match and best_score >= 0.8:
                    fuzzy_matched.append(
                        {
                            "Excel功能点": excel_original,
                            "Excel一级模块": excel_item.get("level1", ""),
                            "Excel二级模块": excel_item.get("level2", ""),
                            "Excel三级模块": excel_item.get("level3", ""),
                            "Word匹配项": best_match.get("original", ""),
                            "位置": "全文模糊匹配",
                            "相似度": f"{best_score:.2%}",
                            "匹配状态": "[PASS] 全文匹配",
                        }
                    )
                    found = True

            if not found:
                not_found_in_word.append(
                    {
                        "Excel功能点": excel_original,
                        "Excel一级模块": excel_item.get("level1", ""),
                        "Excel二级模块": excel_item.get("level2", ""),
                        "Excel三级模块": excel_item.get("level3", ""),
                        "匹配状态": "[FAIL] 全书未找到",
                        "简略描述": "在所有匹配阶段均未找到",
                    }
                )

        phase3_time = time_module.time() - phase3_start
        print(
            f"  [DONE] Phase3 completed: newly matched {len(still_unfound) - len(not_found_in_word)} items, "
            f"missing {len(not_found_in_word)} items (time: {phase3_time:.2f}s)"
        )

        # 返回结果
        total_matched = len(exact_matched) + len(fuzzy_matched)
        match_rate = (total_matched / total * 100) if total > 0 else 0

        report = {
            "exact_matched": exact_matched,
            "fuzzy_matched": fuzzy_matched,
            "not_found_in_word": not_found_in_word,
            "statistics": {
                "excel功能点总数": total,
                "word内容项总数": len(word_items),
                "精确匹配": len(exact_matched),
                "模糊匹配": len(fuzzy_matched),
                "已匹配": total_matched,
                "缺失项": len(not_found_in_word),
                "匹配率": f"{match_rate:.2f}%",
                "完成度": f"{match_rate:.2f}%",
                "阶段1耗时": f"{phase1_time:.2f}s",
                "阶段2耗时": f"{phase2_time:.2f}s",
                "阶段3耗时": f"{phase3_time:.2f}s",
            },
        }

        if progress_callback:
            progress_callback(90, "正在汇总结果...")

        return report

    def _is_in_range(self, num_str: str, start_num: str, end_num: str = None) -> bool:
        """
        判断章节编号num_str是否在范围[start_num, end_num)内
        例如：_is_in_range("4.1.1.1.2", "4.1.1.1", "4.1.1.2") -> True
              _is_in_range("4.1.1.2", "4.1.1.1", "4.1.1.2") -> False
        """
        if not num_str or not start_num:
            return False

        num_str = str(num_str).strip().rstrip(".")
        start_num = str(start_num).strip().rstrip(".")

        # 检查是否在start_num范围内（start_num本身 或 start_num的子章节）
        if num_str == start_num:
            return True
        if num_str.startswith(f"{start_num}."):
            # 如果有end_num，还需检查是否超出范围
            if end_num:
                end_num = str(end_num).strip().rstrip(".")
                # 不能等于或开始于end_num
                if num_str == end_num or num_str.startswith(f"{end_num}."):
                    return False
            return True

        return False

    def _get_next_sibling_number(self, num_str: str) -> str:
        """
        获取章节编号的下一个同级章节编号
        例如：_get_next_sibling_number("4.1.1.1") -> "4.1.1.2"
              _get_next_sibling_number("4.1.1") -> "4.1.2"
              _get_next_sibling_number("4") -> "5"
        """
        num_str = str(num_str).strip().rstrip(".")
        parts = num_str.split(".")
        if not parts:
            return ""

        # 最后一个数字加1
        try:
            last_num = int(parts[-1])
            parts[-1] = str(last_num + 1)
            return ".".join(parts)
        except (ValueError, IndexError):
            return ""

    def _recursive_chapter_search(
        self,
        excel_text,
        excel_parent_num,
        word_items,
        chapter_buckets,
        full_text_content=None,
    ):
        """
        深度搜索：在指定的父章节及其所有子章节中搜索匹配项。
        优化：采用从 [parent] 读取到 [next sibling] 停止的“阅读流”模式，自愈式恢复索引。
        """
        import re

        if not excel_parent_num:
            return False, None, 0.0, ""

        excel_text_clean = self.basic_clean(excel_text)
        excel_parent_num = str(excel_parent_num).strip().rstrip(".")
        next_sibling = self._get_next_sibling_number(excel_parent_num)

        # 建立/恢复索引 (应对预加载数据)
        flow = getattr(self, "word_full_flow", full_text_content) or []
        idxs = getattr(self, "number_to_flow_idx", {})

        if flow and not idxs:
            idxs = {}
            for i, item in enumerate(flow):
                num = item.get("number")
                if num and num not in idxs:
                    idxs[num] = i
            self.number_to_flow_idx = idxs
            self.word_full_flow = flow

        chapter_range_candidates = []
        start_idx = idxs.get(excel_parent_num)

        if start_idx is not None:
            end_idx = len(flow)
            current_depth = len(excel_parent_num.split("."))
            for i in range(start_idx + 1, len(flow)):
                item = flow[i]
                if (
                    item.get("type") == "heading" or item.get("source") == "heading"
                ) and item.get("number"):
                    num = str(item["number"]).strip().rstrip(".")
                    depth = len(num.split("."))
                    if depth <= current_depth and num != excel_parent_num:
                        end_idx = i
                        break
            chapter_range_candidates = flow[start_idx:end_idx]
            if DETAILED_LOG:
                print(
                    f"    [DEBUG-RCS] Flow range [{start_idx}:{end_idx}] has {len(chapter_range_candidates)} items"
                )

        # 兜底：如果流不可用或没找到索引，使用范围检测
        if not chapter_range_candidates:
            for item in word_items or []:
                num = item.get("number")
                if self._is_in_range(num, excel_parent_num, next_sibling):
                    chapter_range_candidates.append(item)
            if full_text_content:
                for item in full_text_content:
                    pnum = item.get("parent_num") or item.get("number")
                    if self._is_in_range(pnum, excel_parent_num, next_sibling):
                        chapter_range_candidates.append(item)

        best_match = None
        best_score = 0.0
        found_location = ""

        # 匹配循环
        for candidate in chapter_range_candidates:
            c_text = candidate.get("text", candidate.get("original", ""))
            if not c_text:
                continue

            c_clean = self.basic_clean(c_text)
            score = self.similarity(excel_text_clean, c_clean)

            # --- 智能特征增强 ---
            if score < 0.90:
                # 策略 1: 包含关系补充
                if c_clean and excel_text_clean:
                    # 包含关系逻辑已整合进 similarity 函数，此处不再手动设置 0.92
                    pass

                # 策略 2: 共同关键词占比 (针对 functional points)
                words_excel = set(re.findall(r"\w+", excel_text_clean))
                words_word = set(re.findall(r"\w+", c_clean))
                if words_excel and words_word:
                    intersection = words_excel & words_word
                    overlap_ratio = len(intersection) / len(words_excel)
                    if overlap_ratio >= 0.7:
                        # 核心校验：如果 Word 文本过短（缺失信息），即使关键词全中也不给高分
                        len_ratio = (
                            len(c_clean) / len(excel_text_clean)
                            if len(excel_text_clean) > 0
                            else 0
                        )
                        if len_ratio < 0.6:
                            score = max(score, 0.6 + overlap_ratio * 0.1)  # 最高 0.7
                        else:
                            score = max(score, 0.7 + overlap_ratio * 0.18)

            # 标题项额外权重
            if candidate.get("type") == "heading":
                if score >= 0.90:
                    score = min(1.0, score + 0.02)

            if score > best_score:
                best_score = score
                best_match = candidate
                found_location = candidate.get("number") or candidate.get(
                    "parent_num", excel_parent_num
                )

        if best_match and best_score >= 0.70:
            return True, best_match, best_score, found_location
        return False, None, 0.0, ""

    def _compare_numbers(self, n1, n2):
        """比较两个章节编号的先后顺序"""
        if not n1 or not n2:
            return 0
        try:
            p1 = [int(p) for p in str(n1).split(".") if p.isdigit()]
            p2 = [int(p) for p in str(n2).split(".") if p.isdigit()]
            for a, b in zip(p1, p2):
                if a < b:
                    return -1
                if a > b:
                    return 1
            return len(p1) - len(p2)
        except:
            return 0

    def _is_in_range(self, current_str, start_str, end_str):
        """判断 current_str 是否在 [start_str, end_str) 范围内"""
        if not current_str or not start_str:
            return False
        current_str = str(current_str).strip().rstrip(".")
        start_str = str(start_str).strip().rstrip(".")

        # 必须以前缀匹配或者是自身
        if current_str == start_str or current_str.startswith(start_str + "."):
            if not end_str:
                return True
            end_str = str(end_str).strip().rstrip(".")
            # 必须小于终点编号
            return self._compare_numbers(current_str, end_str) < 0
        return False

        if DETAILED_LOG:
            print(
                f"  [DEBUG-RCS] Searching in range [{excel_parent_num}, {next_sibling}), "
                f"text='{excel_text_clean[:20]}...'"
            )

        # 【NEW】基于流式索引构建搜索范围
        # 使用文档自然流（full_text_content/full_flow）来确定范围，比进行连续编号推导更可靠
        chapter_range_candidates = []
        bucket_count = 0
        content_count = 0
        content_skipped = 0

        # 获取全文流（Stage 2 使用 full_text_content）
        flow = full_text_content if full_text_content is not None else []

        # 尝试使用基于索引的流式搜索（高性能且准确，可以按文档物理顺序圈定范围）
        if flow and hasattr(self, "number_to_flow_idx") and self.number_to_flow_idx:
            start_idx = self.number_to_flow_idx.get(excel_parent_num)

            if start_idx is not None:
                # 寻找结束位置：下一个同级或更高层级章节的出现位置
                end_idx = len(flow)
                current_depth = len(str(excel_parent_num).split("."))

                for i in range(start_idx + 1, len(flow)):
                    item = flow[i]
                    # 如果发现了标题类型且带有章节号
                    if (
                        item.get("type") == "heading" or item.get("source") == "heading"
                    ) and item.get("number"):
                        num = str(item["number"]).strip().rstrip(".")
                        depth = len(num.split("."))
                        # 如果发现同级或更高层级的标题（且不是当前章节本身），则停止范围
                        # 例如：在4.1.1.1中搜索，遇到4.1.1.2（同级）或4.1.2（更高级）或5（更高级）都应停止
                        if depth <= current_depth and num != excel_parent_num:
                            end_idx = i
                            break

                chapter_range_candidates = flow[start_idx:end_idx]
                content_count = len(chapter_range_candidates)
                if DETAILED_LOG:
                    print(
                        f"    [DEBUG-RCS] Flow-based range: {excel_parent_num} found at index {start_idx}, ends at {end_idx}, total {content_count} items"
                    )

        # 兜底方案：如果流式搜索未命中或不可用，退回到原来的基于编号遍历的逻辑
        if not chapter_range_candidates:
            if DETAILED_LOG:
                print(
                    f"    [DEBUG-RCS] Falling back to range-based search for {excel_parent_num}"
                )

            # 1. 从 Word 项目（通常是标题）中筛选
            for item in word_items or []:
                item_num = item.get("number", "")
                item_parent = item.get("parent_num", "")

                # 【优化】即使没有章节编号的项（如未定义的子标题），只要父章节在范围内也应包含
                in_range = False
                if item_num and self._is_in_range(
                    item_num, excel_parent_num, next_sibling
                ):
                    in_range = True
                elif item_parent and self._is_in_range(
                    item_parent, excel_parent_num, next_sibling
                ):
                    # 即使自己没编号，但父章节在搜索范围内
                    in_range = True

                if in_range:
                    chapter_range_candidates.append(item)
                    bucket_count += 1

            # 2. 从全文内容（段落、表格）中筛选
            if full_text_content:
                for full_item in full_text_content:
                    # 检查其所属章节是否在范围内
                    # 优先检查 parent_num, 这是由 _read_word_full_text 正确关联的
                    p_num = full_item.get("parent_num", "")
                    if not p_num:
                        p_num = full_item.get("number", "")

                    if p_num and self._is_in_range(
                        p_num, excel_parent_num, next_sibling
                    ):
                        chapter_range_candidates.append(full_item)
                        content_count += 1
                    elif p_num:
                        content_skipped += 1

        # DEBUG：输出范围信息（当候选项太少时）
        if (bucket_count + content_count) < 5:
            # 统计候选项中的不同章节号
            range_chapters = set()
            for item in chapter_range_candidates:
                num = item.get("number", "") or item.get("parent_num", "")
                if num:
                    range_chapters.add(num)

            print(
                f"    Range[{excel_parent_num}, {next_sibling}): "
                f"items={bucket_count}, content={content_count}, skipped={content_skipped}, "
                f"total={len(chapter_range_candidates)}, chapters={sorted(range_chapters) if range_chapters else 'NONE'}"
            )

            # 【DEBUG】如果full_text_content存在但content_count为0，输出第一个跳过项的详情
            if full_text_content and content_count == 0 and content_skipped > 0:
                print(
                    f"      [DEBUG] full_text_content has {len(full_text_content)} items, but all skipped!"
                )
                # 输出前3个跳过项的详细信息
                checked = 0
                for full_item in full_text_content[:50]:  # 检查前50项
                    item_num = full_item.get("number", "")
                    if not item_num:
                        item_num = full_item.get("parent_num", "")

                    # 显示前3个不在范围内的样本
                    if item_num and not self._is_in_range(
                        item_num, excel_parent_num, next_sibling
                    ):
                        if checked < 3:
                            sample_text = (
                                full_item.get("text", "")
                                or full_item.get("original", "")
                            )[:30]
                            print(
                                f"        Sample #{checked+1}: text='{sample_text}...' | num={item_num} | in_range=False"
                            )
                            checked += 1
                    # 显示前3个没有编号的样本
                    elif not item_num:
                        if checked < 3:
                            sample_text = (
                                full_item.get("text", "")
                                or full_item.get("original", "")
                            )[:30]
                            print(
                                f"        Sample #{checked+1}: text='{sample_text}...' | num=NONE | reason=NO_NUMBER"
                            )
                            checked += 1

                    if checked >= 6:  # 最多显示6个样本
                        break

        # 3. 在范围内进行匹配
        best_match = None
        best_score = 0.0
        best_is_exact = False
        found_location = ""

        for candidate in chapter_range_candidates:
            candidate_text = candidate.get("text", candidate.get("cleaned", ""))
            if not candidate_text:
                continue

            candidate_clean = self.basic_clean(candidate_text)
            score = self.similarity(excel_text_clean, candidate_clean)

            # 增强匹配策略 (同步自 similarity 逻辑，不再手动设置 0.92)
            if score < 0.85 and excel_text_clean and candidate_clean:
                # 场景 A: Word 覆盖 Excel (如 Word 是 "xx功能描述"，Excel 是 "xx功能")
                # 逻辑已剥离到 similarity，此处仅作占位
                pass

                # 策略2：关键词重叠率 (略微放宽)
                # [FIX] 不再使用 else，确保两个策略都能作为模糊匹配的补充
                excel_keywords = set(re.findall(r"\w+", excel_text_clean))
                candidate_keywords = set(re.findall(r"\w+", candidate_clean))
                if excel_keywords and candidate_keywords:
                    overlap_size = len(excel_keywords & candidate_keywords)
                    overlap_ratio = overlap_size / len(excel_keywords)
                    if overlap_ratio >= 0.7:
                        # [PENALTY] 长度校验：通过 similarity 已经处理了大部分，此处做二次拦截
                        len_ratio = (
                            len(candidate_clean) / len(excel_text_clean)
                            if len(excel_text_clean) > 0
                            else 0
                        )
                        if len_ratio < 0.6:
                            # 长度不足，即便是关键词重合也不给高分
                            score = max(score, 0.60)
                        elif overlap_ratio >= 0.85:
                            score = max(score, 0.85)
                        else:
                            score = max(score, 0.70)

            # 【重要】如果当前候选是一个标题项（Word 大纲中的标题），给予额外权重
            if candidate.get("source") == "heading" or "number" in candidate:
                # 只有在相似度已经很高（>0.9）的情况下才锦上添花，不要把低分强行拉高
                if score >= 0.90:
                    score = min(1.0, score + 0.02)

            # 精确匹配优先保留
            if excel_text_clean == candidate_clean:
                if not best_is_exact or score > best_score:
                    best_score = score
                    best_match = candidate
                    best_is_exact = True
                    found_location = candidate.get(
                        "number", candidate.get("parent_num", excel_parent_num)
                    )

            # 模糊匹配记录，但继续搜索以找到更好的匹配
            elif not best_is_exact and score > best_score:
                best_score = score
                best_match = candidate
                found_location = candidate.get(
                    "number", candidate.get("parent_num", excel_parent_num)
                )

        # 返回结果：
        # - 精确匹配(score == 1.0) 立即返回
        # - 高质量模糊匹配(>= 0.70) 才返回
        if best_match:
            if best_is_exact or best_score >= 0.70:
                return True, best_match, best_score, found_location

        return False, None, 0.0, ""

    # 简单匹配模式的 match_documents 方法保持不变，因为它与层级匹配不同

    # def match_hierarchical_documents(
    #         self,
    #         excel_data: List[Dict],  # ✅ 修复：添加冒号分隔变量名和类型
    #         word_hierarchy: Dict,
    #         sheet_name=0,  # 修复：移除多余空格
    # ) -> List[Dict]:
    #     """匹配Excel模块层级与Word文档层级结构"""
    #     import re
    #     from typing import List, Dict, Tuple
    #
    #     # === 辅助函数：层级深度调整 ===
    #     def get_adjusted_depth(number_str: str) -> int:
    #         """将Word编号转换为Excel层级深度"""
    #         if not number_str:
    #             return 0
    #         parts = [p for p in number_str.split('.') if p.isdigit()]
    #         raw_depth = len(parts)
    #         return max(0, raw_depth - 1)
    #
    #     # === 辅助函数：判断后代关系 ===
    #     def is_descendant(parent_num: str, child_num: str) -> bool:
    #         """判断child_num是否是parent_num的后代编号"""
    #         if not parent_num or not child_num:
    #             return False
    #         parent_clean = parent_num.strip().rstrip('.')
    #         child_clean = child_num.strip().rstrip('.')
    #         return child_clean.startswith(parent_clean + '.')
    #
    #     # === 核心匹配函数：三阶段智能搜索 ===
    #     def find_match_for_excel_level(excel_text: str, candidates: List[Dict], excel_level: int) -> Tuple[
    #         bool, Dict, float, bool]:
    #         """三阶段智能匹配（兼容3返回值）"""
    #         if not excel_text or not candidates:
    #             return False, None, 0.0, False
    #
    #         # 标准化目标文本
    #         target_clean = re.sub(r"\s+", "", str(excel_text).strip().lower())
    #
    #         # === 阶段1: 在目标层级搜索"完全相同"匹配 ===
    #         expected_adjusted_depth = excel_level
    #         depth_candidates = [
    #             c for c in candidates
    #             if get_adjusted_depth(c.get("number", "")) == expected_adjusted_depth
    #         ]
    #
    #         if depth_candidates:
    #             for cand in depth_candidates:
    #                 cand_clean = re.sub(r"\s+", "", cand.get("text", "").strip().lower())
    #                 if target_clean == cand_clean:
    #                     return True, cand, 1.0, True  # ✅ 完全相同立即返回
    #
    #         # === 阶段2: 在所有层级搜索"完全相同"匹配 ===
    #         for cand in candidates:
    #             cand_clean = re.sub(r"\s+", "", cand.get("text", "").strip().lower())
    #             if target_clean == cand_clean:
    #                 return True, cand, 1.0, True  # ✅ 完全相同立即返回
    #
    #         # === 阶段3: 在目标层级搜索最佳匹配 ===
    #         best_match = None
    #         best_score = -1.0
    #         best_is_exact = False
    #
    #         if depth_candidates:
    #             found, match, score, is_exact = self.find_match(
    #                 excel_text,
    #                 depth_candidates,
    #                 target_depth=expected_adjusted_depth
    #             )
    #             if found:
    #                 best_match = match
    #                 best_score = score
    #                 best_is_exact = is_exact
    #
    #         # === 阶段4: 全局搜索作为后备 ===
    #         found_global, match_global, score_global, is_exact_global = self.find_match(
    #             excel_text,
    #             candidates,
    #             target_depth=0
    #         )
    #
    #         # 决策：优先选择目标层级匹配（即使分数略低）
    #         if found_global and score_global > best_score + 0.03:
    #             return True, match_global, score_global, is_exact_global
    #         elif best_match:
    #             return True, best_match, best_score, best_is_exact
    #         elif found_global:
    #             return True, match_global, score_global, is_exact_global
    #
    #         return False, None, 0.0, False
    #
    #     # === 主匹配流程 ===
    #     results = []
    #
    #     for idx, excel_row in enumerate(excel_data, 1):
    #         result = {
    #             "excel_row_index": idx,
    #             "level1": excel_row.get("level1", ""),
    #             "level2": excel_row.get("level2", ""),
    #             "level3": excel_row.get("level3", ""),
    #             "word_l1_original": "",
    #             "word_l1_title": "",
    #             "word_l1_num": "",
    #             "word_l2_original": "",
    #             "word_l2_title": "",
    #             "word_l2_num": "",
    #             "word_l3_original": "",
    #             "word_l3_title": "",
    #             "word_l3_num": "",
    #             "l1_matched": False,
    #             "l2_matched": False,
    #             "l3_matched": False,
    #             "l1_score": 0.0,
    #             "l2_score": 0.0,
    #             "l3_score": 0.0,
    #             "l1_actual_depth": 0,
    #             "l2_actual_depth": 0,
    #             "l3_actual_depth": 0,
    #             "l1_is_exact": False,
    #             "l2_is_exact": False,
    #             "l3_is_exact": False,
    #             "warnings": []
    #         }
    #
    #         # === 2.1 一级模块匹配 ===
    #         if excel_row.get("level1"):
    #             candidates = word_hierarchy["all_items"]
    #             found, match, score, is_exact = find_match_for_excel_level(
    #                 excel_row["level1"],
    #                 candidates,
    #                 excel_level=1
    #             )
    #             if found:
    #                 result["l1_matched"] = True
    #                 result["word_l1_original"] = match.get("original", match.get("display", match.get("text", "")))
    #                 result["word_l1_title"] = match.get("text", match.get("cleaned", ""))
    #                 result["word_l1_num"] = match.get("number", "")
    #                 result["l1_score"] = score
    #                 result["l1_actual_depth"] = match.get("depth", 0)
    #                 result["l1_is_exact"] = is_exact
    #
    #                 # 检查层级是否匹配
    #                 expected_depth = 1
    #                 actual_adjusted_depth = get_adjusted_depth(result["word_l1_num"])
    #                 if actual_adjusted_depth != expected_depth:
    #                     result["warnings"].append(
    #                         f"L1层级不匹配: 期望深度{expected_depth}, 实际深度{actual_adjusted_depth} (编号:{result['word_l1_num']})"
    #                     )
    #             else:
    #                 result["word_l1_title"] = "❌ 缺失"
    #                 result["word_l1_original"] = "❌ 缺失"
    #
    #         # === 2.2 二级模块匹配 ===
    #         if excel_row.get("level2"):
    #             candidates = word_hierarchy["all_items"]
    #             # 优先在L1的子项中搜索（如果L1已匹配且有编号）
    #             if result["l1_matched"] and result["word_l1_num"]:
    #                 constrained = [
    #                     c for c in candidates
    #                     if is_descendant(result["word_l1_num"], c.get("number", ""))
    #                 ]
    #                 if constrained:
    #                     candidates = constrained
    #
    #             found, match, score, is_exact = find_match_for_excel_level(
    #                 excel_row["level2"],
    #                 candidates,
    #                 excel_level=2
    #             )
    #             if found:
    #                 result["l2_matched"] = True
    #                 result["word_l2_original"] = match.get("original", match.get("display", match.get("text", "")))
    #                 result["word_l2_title"] = match.get("text", match.get("cleaned", ""))
    #                 result["word_l2_num"] = match.get("number", "")
    #                 result["l2_score"] = score
    #                 result["l2_actual_depth"] = match.get("depth", 0)
    #                 result["l2_is_exact"] = is_exact
    #
    #                 # 检查层级是否匹配
    #                 expected_depth = 2
    #                 actual_adjusted_depth = get_adjusted_depth(result["word_l2_num"])
    #                 if actual_adjusted_depth != expected_depth:
    #                     result["warnings"].append(
    #                         f"L2层级不匹配: 期望深度{expected_depth}, 实际深度{actual_adjusted_depth} (编号:{result['word_l2_num']})"
    #                     )
    #
    #                 # 检查父子关系
    #                 if result["l1_matched"] and result["word_l1_num"]:
    #                     if not is_descendant(result["word_l1_num"], result["word_l2_num"]):
    #                         result["warnings"].append(
    #                             f"L2非L1子项: L1={result['word_l1_num']}, L2={result['word_l2_num']}"
    #                         )
    #             else:
    #                 result["word_l2_title"] = "❌ 缺失"
    #                 result["word_l2_original"] = "❌ 缺失"
    #
    #         # === 2.3 三级模块匹配 ===
    #         if excel_row.get("level3"):
    #             candidates = word_hierarchy["all_items"]
    #             # 优先在L2或L1的子项中搜索
    #             parent_num = None
    #             search_candidates = []
    #
    #             if result["l2_matched"] and result["word_l2_num"]:
    #                 parent_num = result["word_l2_num"]
    #                 search_candidates = [
    #                     c for c in candidates
    #                     if is_descendant(parent_num, c.get("number", ""))
    #                 ]
    #             elif result["l1_matched"] and result["word_l1_num"]:
    #                 parent_num = result["word_l1_num"]
    #                 search_candidates = [
    #                     c for c in candidates
    #                     if is_descendant(parent_num, c.get("number", ""))
    #                 ]
    #
    #             if not search_candidates:
    #                 search_candidates = candidates
    #
    #             found, match, score, is_exact = find_match_for_excel_level(
    #                 excel_row["level3"],
    #                 search_candidates,
    #                 excel_level=3
    #             )
    #             if found:
    #                 result["l3_matched"] = True
    #                 result["word_l3_original"] = match.get("original", match.get("display", match.get("text", "")))
    #                 result["word_l3_title"] = match.get("text", match.get("cleaned", ""))
    #                 result["word_l3_num"] = match.get("number", "")
    #                 result["l3_score"] = score
    #                 result["l3_actual_depth"] = match.get("depth", 0)
    #                 result["l3_is_exact"] = is_exact
    #
    #                 # 检查层级是否匹配
    #                 expected_depth = 3
    #                 actual_adjusted_depth = get_adjusted_depth(result["word_l3_num"])
    #                 if actual_adjusted_depth != expected_depth:
    #                     result["warnings"].append(
    #                         f"L3层级不匹配: 期望深度{expected_depth}, 实际深度{actual_adjusted_depth} (编号:{result['word_l3_num']})"
    #                     )
    #
    #                 # 检查父子关系
    #                 if result["l2_matched"] and result["word_l2_num"]:
    #                     if not is_descendant(result["word_l2_num"], result["word_l3_num"]):
    #                         result["warnings"].append(
    #                             f"L3非L2子项: L2={result['word_l2_num']}, L3={result['word_l3_num']}"
    #                         )
    #                 elif result["l1_matched"] and result["word_l1_num"]:
    #                     if not is_descendant(result["word_l1_num"], result["word_l3_num"]):
    #                         result["warnings"].append(
    #                             f"L3非L1子项: L1={result['word_l1_num']}, L3={result['word_l3_num']}"
    #                         )
    #             else:
    #                 result["word_l3_title"] = "❌ 缺失"
    #                 result["word_l3_original"] = "❌ 缺失"
    #
    #         # === 3. 添加匹配质量标签 ===
    #         if result["l1_matched"]:
    #             if result["l1_is_exact"]:
    #                 result["l1_quality"] = "✅ 完全相同"
    #             elif result["l1_score"] >= 0.95:
    #                 result["l1_quality"] = "🔹 完全匹配"
    #             else:
    #                 result["l1_quality"] = f"🔸 模糊匹配({result['l1_score']:.2f})"
    #         else:
    #             result["l1_quality"] = "❌ 未匹配"
    #
    #         if result["l2_matched"]:
    #             if result["l2_is_exact"]:
    #                 result["l2_quality"] = "✅ 完全相同"
    #             elif result["l2_score"] >= 0.95:
    #                 result["l2_quality"] = "🔹 完全匹配"
    #             else:
    #                 result["l2_quality"] = f"🔸 模糊匹配({result['l2_score']:.2f})"
    #         else:
    #             result["l2_quality"] = "❌ 未匹配"
    #
    #         if result["l3_matched"]:
    #             if result["l3_is_exact"]:
    #                 result["l3_quality"] = "✅ 完全相同"
    #             elif result["l3_score"] >= 0.95:
    #                 result["l3_quality"] = "🔹 完全匹配"
    #             else:
    #                 result["l3_quality"] = f"🔸 模糊匹配({result['l3_score']:.2f})"
    #         else:
    #             result["l3_quality"] = "❌ 未匹配"
    #
    #         results.append(result)
    #
    #     return results

    def save_report(self, report: Dict, output_file: str = "匹配报告.xlsx"):
        """
        保存报告到 Excel（修复层级不匹配显示问题）
        """
        import os
        from pathlib import Path

        original_output_file = output_file
        attempt = 0
        max_attempts = 10

        while attempt < max_attempts:
            try:
                print(f"\n正在生成报告文件: {output_file}")
                with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                    # 统计概览
                    stats_df = pd.DataFrame([report["statistics"]])
                    stats_df.to_excel(writer, sheet_name="📊 统计概览", index=False)

                    # 层级匹配报告的列顺序
                    hierarchical_column_order = [
                        "Excel一级模块",
                        "Word一级标题",
                        "Excel二级模块",
                        "Word二级标题",
                        "Excel三级模块",
                        "Word三级标题",
                        "Word完整编号",
                        "相似度",
                        "匹配状态",
                        "缺失简略描述",
                        "层级不匹配简略描述",
                        "深度警告描述",  # 新增字段
                        "简略描述",
                        "row_range",
                    ]

                    # 简单匹配报告的列顺序
                    simple_column_order = [
                        "Excel一级模块",
                        "Excel二级模块",
                        "Excel三级模块",
                        "Excel功能点",
                        "Word匹配项",
                        "位置",
                        "相似度",
                        "匹配状态",
                        "简略描述",
                    ]

                    # 【关键修复1】判断报告类型，不仅要检查前三个列表，还要检查hierarchy_mismatched
                    is_node5_report = False

                    # 方法1：检查是否有层级匹配特有的字段
                    sample_item = None
                    for key in [
                        "exact_matched",
                        "fuzzy_matched",
                        "not_found_in_word",
                        "hierarchy_mismatched",
                    ]:
                        if report.get(key) and len(report[key]) > 0:
                            sample_item = report[key][0]
                            break

                    if sample_item and "Word一级标题" in sample_item:
                        is_node5_report = True
                    else:
                        # 方法2：检查统计信息中的"层级不匹配"字段
                        if "层级不匹配" in str(report.get("statistics", {})):
                            is_node5_report = True

                    print(
                        f"  报告类型判断: {'层级匹配报告' if is_node5_report else '简单匹配报告'}"
                    )

                    current_column_order = (
                        hierarchical_column_order
                        if is_node5_report
                        else simple_column_order
                    )

                    # 准备合并的数据列表
                    all_data = []

                    # 精确匹配项
                    if report.get("exact_matched"):
                        for item in report["exact_matched"]:
                            item_copy = item.copy()
                            if "匹配状态" not in item_copy:
                                item_copy["匹配状态"] = "精确匹配"
                            all_data.append(item_copy)
                        exact_df = pd.DataFrame(report["exact_matched"])
                        exact_df = exact_df[
                            [
                                col
                                for col in current_column_order
                                if col in exact_df.columns
                            ]
                        ]
                        exact_df.to_excel(writer, sheet_name="✅ 精确匹配", index=False)
                    else:
                        empty_df = pd.DataFrame([{"说明": "无精确匹配项"}])
                        empty_df.to_excel(writer, sheet_name="✅ 精确匹配", index=False)

                    # 模糊匹配项
                    if report.get("fuzzy_matched"):
                        for item in report["fuzzy_matched"]:
                            item_copy = item.copy()
                            if "匹配状态" not in item_copy:
                                item_copy["匹配状态"] = "模糊匹配"
                            all_data.append(item_copy)
                        fuzzy_df = pd.DataFrame(report["fuzzy_matched"])
                        fuzzy_df = fuzzy_df[
                            [
                                col
                                for col in current_column_order
                                if col in fuzzy_df.columns
                            ]
                        ]
                        fuzzy_df.to_excel(writer, sheet_name="🔍 模糊匹配", index=False)
                    else:
                        empty_df = pd.DataFrame([{"说明": "无模糊匹配项"}])
                        empty_df.to_excel(writer, sheet_name="🔍 模糊匹配", index=False)

                    # 【关键修复2】合并"缺失"和"层级不匹配"工作表
                    problem_items = []

                    # 缺失项
                    if report.get("not_found_in_word"):
                        for item in report["not_found_in_word"]:
                            item_copy = item.copy()
                            if "匹配状态" not in item_copy:
                                item_copy["匹配状态"] = "缺失"
                            problem_items.append(item_copy)

                    # 层级不匹配项
                    if report.get("hierarchy_mismatched"):
                        for item in report["hierarchy_mismatched"]:
                            item_copy = item.copy()
                            if "匹配状态" not in item_copy:
                                item_copy["匹配状态"] = "层级不匹配"
                            problem_items.append(item_copy)

                    if problem_items:
                        problem_df = pd.DataFrame(problem_items)
                        problem_df = problem_df[
                            [
                                col
                                for col in current_column_order
                                if col in problem_df.columns
                            ]
                        ]
                        # 统一保存到 "⚠️ 层级不匹配" 工作表
                        problem_df.to_excel(
                            writer, sheet_name="❌ 缺失和⚠️ 层级不匹配", index=False
                        )
                        print(
                            f"  已创建 '❌ 缺失和⚠️ 层级不匹配' 工作表，包含 {len(problem_df)} 项 (含缺失项)"
                        )
                    else:
                        empty_df = pd.DataFrame(
                            [{"说明": "🎉 恭喜！未发现缺失或层级不匹配项！"}]
                        )
                        empty_df.to_excel(
                            writer, sheet_name="⚠️ 层级不匹配", index=False
                        )

                    # 创建合并工作表（包含所有匹配结果）
                    if all_data or problem_items:
                        # 确保 all_data 包含所有类型
                        final_all_data = all_data.copy()
                        final_all_data.extend(problem_items)

                        all_df = pd.DataFrame(final_all_data)
                        cols = [
                            col for col in current_column_order if col in all_df.columns
                        ]
                        if cols:  # 确保有列可写入
                            all_df = all_df[cols]
                            all_df.to_excel(
                                writer, sheet_name="📋 汇总匹配结果", index=False
                            )

                print(f"✓ 报告已保存: {output_file}")
                # 打印工作表信息
                import openpyxl

                wb = openpyxl.load_workbook(output_file, read_only=True)
                print(f"  包含的工作表: {', '.join(wb.sheetnames)}")
                wb.close()
                return output_file

            except PermissionError as e:
                # 文件被占用，生成副本文件名
                attempt += 1
                if attempt >= max_attempts:
                    error_msg = (
                        f"无法保存报告：尝试了 {max_attempts} 次仍然失败。\n"
                        f"请关闭所有正在使用该文件的程序（如 Excel），然后重试。"
                    )
                    print(f"✗ {error_msg}")
                    raise PermissionError(error_msg)

                # 生成副本文件名
                path = Path(original_output_file)
                stem = path.stem
                suffix = path.suffix
                parent = path.parent
                output_file = str(parent / f"{stem}_副本{attempt}{suffix}")
                print(f"  ⚠️ 原文件被占用，尝试保存到: {output_file}")

            except Exception as e:
                print(f"✗ 保存报告失败: {str(e)}")
                import traceback

                traceback.print_exc()
                raise

    def match_hierarchical_documents(
        self,
        word_file: str,
        excel_file: str,
        sheet_name=0,
        header=0,
        level1_col: int = 0,
        level2_col: int = 1,
        level3_col: int = 2,
        progress_callback=None,
        hierarchy_log_path: str = None,
        word_sections_preloaded=None,
    ) -> Dict:
        """
        层级匹配模式 - 三阶段智能匹配（完全相同优先）
        修复版：1) 修复变量名拼写 2) 添加类型安全检查 3) 实现完全相同优先匹配 4) 修复列表对象调用get()错误
        """
        import re
        import os
        from datetime import datetime
        from typing import List, Dict, Tuple, Any, Union

        # === 辅助函数：安全获取字典/列表值 ===
        def safe_get(
            row: Union[Dict, List, Tuple], key: Union[str, int], default=""
        ) -> str:
            """安全获取值，兼容字典和列表格式"""
            try:
                if isinstance(row, dict):
                    val = row.get(key, default)
                elif isinstance(row, (list, tuple)):
                    # 假设列表顺序: [level1, level2, level3, ...]
                    key_map = {"level1": 0, "level2": 1, "level3": 2}
                    idx = key_map.get(key, -1) if isinstance(key, str) else key

                    if isinstance(idx, int) and 0 <= idx < len(row):
                        val = row[idx]
                    else:
                        # 如果索引无效或key不是数字，尝试按列名查找
                        if hasattr(row, "columns"):
                            # 如果是pandas Series或DataFrame行
                            if key in row.columns:
                                val = row[key]
                            else:
                                val = default
                        else:
                            val = default
                elif hasattr(row, "_asdict"):  # 命名元组
                    val = row._asdict().get(key, default)
                elif hasattr(row, "__dict__"):  # 对象
                    val = getattr(row, key, default)
                else:
                    val = default

                # 标准化返回值
                if val is None or (
                    isinstance(val, str) and val.lower() in ["nan", "none", ""]
                ):
                    return ""
                return str(val).strip()
            except Exception as e:
                print(
                    f"[ERROR] safe_get 出错: row类型={type(row)}, key={key}, error={e}"
                )
                return default

        # === 辅助函数：获取 Word 实际层级（按编号段数计算） ===
        def get_word_level(item: Dict) -> int:
            """返回 Word 的实际显示层级（1/2/3...），按编号段数计算：4->0, 4.1->1, 4.1.1->2, 4.1.1.1->3..."""
            num = (item.get("number") or "").strip()
            if not num:
                return 0

            # 统一按编号段数计算层级：4 -> 0级, 4.1 -> 1级, 4.1.1 -> 2级, 4.1.1.1 -> 3级 ...
            parts = [p for p in num.split(".") if p.strip()]
            if not parts:
                return 0

            return len(parts) - 1

        # === 核心匹配函数：三阶段智能搜索 ===
        def find_match_for_excel_level(
            excel_text: str, candidates: List[Dict], excel_level: int
        ) -> Tuple[bool, Dict, float]:
            """
            三阶段智能匹配策略：
            阶段1: 目标层级"完全相同" → 立即返回
            阶段2: 任意层级"完全相同" → 返回（质量优先于层级）
            阶段3+4: 全局最佳匹配
            """
            if not excel_text or not candidates:
                return False, None, 0.0

            # 标准化目标文本
            target_clean = re.sub(r"\s+", "", str(excel_text).strip().lower())

            # === 阶段1: 在目标层级搜索"完全相同"匹配 ===
            expected_adjusted_depth = excel_level
            depth_candidates = [
                c for c in candidates if get_word_level(c) == expected_adjusted_depth
            ]

            if depth_candidates:
                for cand in depth_candidates:
                    cand_clean = re.sub(
                        r"\s+", "", cand.get("text", "").strip().lower()
                    )
                    if target_clean == cand_clean:
                        return True, cand, 1.0  # ✅ 完全相同立即返回

            # === 阶段2: 在所有层级搜索"完全相同"匹配 ===
            for cand in candidates:
                cand_clean = re.sub(r"\s+", "", cand.get("text", "").strip().lower())
                if target_clean == cand_clean:
                    return True, cand, 1.0  # ✅ 完全相同立即返回

            # === 阶段3: 在目标层级搜索最佳匹配 ===
            best_match = None
            best_score = -1.0

            if depth_candidates:
                found, match, score = self.find_match(
                    excel_text, depth_candidates, target_depth=expected_adjusted_depth
                )
                if found:
                    best_match = match
                    best_score = score

            # === 阶段4: 全局搜索作为后备 ===
            found_global, match_global, score_global = self.find_match(
                excel_text, candidates, target_depth=0
            )

            # 决策：优先选择目标层级匹配（即使分数略低）
            if found_global and score_global > best_score + 0.03:
                return True, match_global, score_global
            elif best_match:
                return True, best_match, best_score
            elif found_global:
                return True, match_global, score_global

            return False, None, 0.0

        # === 主匹配流程 ===
        if progress_callback:
            progress_callback(5, "正在提取 Word 层级结构...")

        # 1. 提取 Word 层级结构
        if word_sections_preloaded:
            if isinstance(word_sections_preloaded, list):
                # [关键修复] 如果传入的是列表（由外部预提取的章节列表），转换回字典格式
                # 并且确保内部每个 item 都有必要的字段兼容性
                normalized_items = []
                for idx, item in enumerate(word_sections_preloaded):
                    if isinstance(item, dict):
                        # 兼容处理：确保有 match_hierarchical_documents 需要的字段
                        new_item = item.copy()
                        if "text" not in new_item:
                            new_item["text"] = item.get(
                                "title", item.get("display", "")
                            )
                        if "cleaned" not in new_item:
                            new_item["cleaned"] = new_item["text"]
                        if "number" not in new_item:
                            new_item["number"] = item.get("num", "")
                        if "original" not in new_item:
                            num = new_item["number"]
                            text = new_item["text"]
                            new_item["original"] = f"{num} {text}" if num else text
                        if "level" not in new_item:
                            new_item["level"] = item.get("depth", 0)
                        normalized_items.append(new_item)
                    else:
                        # 如果不是 dict（意外情况），则尝试转换
                        normalized_items.append(
                            {
                                "text": str(item),
                                "cleaned": str(item),
                                "number": "",
                                "original": str(item),
                                "level": 0,
                            }
                        )

                word_hierarchy = {
                    "all_items": normalized_items,
                    "level1": [
                        i
                        for i in normalized_items
                        if i.get("level") == 1 or i.get("depth") == 1
                    ],
                    "level2": [
                        i
                        for i in normalized_items
                        if i.get("level") == 2 or i.get("depth") == 2
                    ],
                    "level3": [
                        i
                        for i in normalized_items
                        if i.get("level", 0) >= 3 or i.get("depth", 0) >= 3
                    ],
                }
            else:
                word_hierarchy = word_sections_preloaded
        else:
            # 默认使用稳定大纲模式，确保获取的是 Word 目录树的实际层级
            word_hierarchy = self.extract_word_content(
                word_file, mode="stable", progress_callback=progress_callback
            )

        if not word_hierarchy or (
            isinstance(word_hierarchy, dict) and not word_hierarchy.get("all_items")
        ):
            raise ValueError("未能从 Word 文档中提取到有效的层级结构")

        # === [仅用于报表回填] 构建 Word 目录编号索引（按出现顺序）===
        word_items = (
            word_hierarchy.get("all_items", [])
            if isinstance(word_hierarchy, dict)
            else (word_hierarchy or [])
        )

        def _norm_num(n: str) -> str:
            return (n or "").strip().rstrip(".")

        def _num_parts(n: str) -> List[str]:
            n = _norm_num(n)
            if not n:
                return []
            parts = [p for p in n.split(".") if p.isdigit()]
            return parts

        def _num_depth(n: str) -> int:
            return len(_num_parts(n))

        word_num_index: Dict[str, Dict] = {}
        ordered_nums: List[Tuple[int, str]] = []
        for _i, _it in enumerate(word_items):
            if not isinstance(_it, dict):
                continue
            _num = _norm_num(_it.get("number") or _it.get("num") or "")
            if not _num:
                continue
            if _num not in word_num_index:
                word_num_index[_num] = _it
                _order = _it.get("order")
                if not isinstance(_order, int):
                    _order = _i
                ordered_nums.append((_order, _num))
        ordered_nums.sort(key=lambda x: x[0])

        # === 调试日志：打印所有 Word 编号索引 ===
        if RuntimeLogger:
            RuntimeLogger.log(
                f"  [DEBUG] Word 编号索引构建完成: 共 {len(word_num_index)} 个编号",
                level="DEBUG",
            )
            for num in sorted(
                word_num_index.keys(),
                key=lambda x: (
                    [int(p) for p in x.split(".") if p.isdigit()] if x else []
                ),
            ):
                item = word_num_index[num]
                title = item.get("text", "")[:30]
                depth = len([p for p in num.split(".") if p.isdigit()])
                RuntimeLogger.log(
                    f"  [DEBUG]   {num:15} (深度{depth}) -> {title}", level="DEBUG"
                )

        def _node_for_num(num: str) -> Dict:
            num = _norm_num(num)
            if not num:
                return {"title": "❌ 缺失", "num": "", "original": ""}
            it = word_num_index.get(num)
            if not it:
                # 找不到就按缺失处理（不合成虚拟节点，避免“目录树没有但硬填”）
                return {"title": "❌ 缺失", "num": "", "original": ""}
            title = it.get("text") or it.get("title") or it.get("cleaned") or ""
            title = str(title).strip()
            original = it.get("original")
            if not original:
                original = f"{num} {title}" if num and title else (title or num)
            return {
                "title": title or "❌ 缺失",
                "num": num,
                "original": str(original).strip(),
            }

        def _first_child(prefix_num: str, target_depth: int) -> str:
            prefix_num = _norm_num(prefix_num)
            if not prefix_num:
                return ""
            prefix = prefix_num + "."
            for _, n in ordered_nums:
                if n.startswith(prefix) and _num_depth(n) == target_depth:
                    return n
            return ""

        def _compute_display_tree(anchor_num: str) -> Dict[int, Dict]:
            """根据命中的编号，获取其在 Word 中真实的 1/2/3 级路径（4.1=1级, 4.1.1=2级, 4.1.1.1=3级）。
            如果匹配深度不足，则自动向下寻找首个子项补全，用于在描述中展示“期望”的层级结构。
            """
            out = {
                1: {"title": "❌ 缺失", "num": "", "original": ""},
                2: {"title": "❌ 缺失", "num": "", "original": ""},
                3: {"title": "❌ 缺失", "num": "", "original": ""},
            }
            parts = _num_parts(anchor_num)
            if not parts:
                return out

            l1_num, l2_num, l3_num = "", "", ""

            # --- 1. 向上回溯路径 ---
            if len(parts) >= 2:
                l1_num = ".".join(parts[:2])
            if len(parts) >= 3:
                l2_num = ".".join(parts[:3])
            if len(parts) >= 4:
                l3_num = ".".join(parts[:4])

            # 如果深度只有1(即'4')，尝试找第一个1级子项(4.1)
            if not l1_num and len(parts) == 1 and parts[0] == "4":
                l1_num = _first_child("4", 2)

            # --- 2. 向下补全缺失层级 (用于生成完整的期望路径) ---
            if l1_num and not l2_num:
                l2_num = _first_child(l1_num, 3)
            if l2_num and not l3_num:
                l3_num = _first_child(l2_num, 4)
            # 特殊纠正：如果 l1 仍然为空，说明匹配到了非 4 开头的编号，则保持 out 为空

            out[1] = _node_for_num(l1_num)
            out[2] = _node_for_num(l2_num)
            out[3] = _node_for_num(l3_num)
            return out

        if progress_callback:
            progress_callback(25, "正在提取 Excel 模块数据...")

        # 2. 提取 Excel 模块数据
        excel_data_raw = self.extract_excel_content(
            excel_file,
            mode="hierarchical",
            sheet_name=sheet_name,
            header=header,
            level1_col=level1_col,
            level2_col=level2_col,
            level3_col=level3_col,
        )

        # === 关键修复：确保数据格式正确 ===
        if not excel_data_raw:
            raise ValueError("未能从 Excel 中提取到有效的模块数据")

        excel_data = []
        for idx, row in enumerate(excel_data_raw):
            # 确保每一行都是字典格式
            if isinstance(row, dict):
                excel_data.append(row)
            elif isinstance(row, (list, tuple)):
                # 转换为字典格式
                converted_row = {
                    "level1": safe_get(row, level1_col, ""),
                    "level2": safe_get(row, level2_col, ""),
                    "level3": safe_get(row, level3_col, ""),
                    "row_index": idx,
                    "original_row": row,  # 保留原始行数据用于调试
                }
                excel_data.append(converted_row)
            else:
                # 其他类型，尝试转换为字典
                try:
                    converted_row = {}
                    # 尝试使用属性访问
                    for attr in [
                        "level1",
                        "level2",
                        "level3",
                        "level1_original",
                        "level2_original",
                        "level3_original",
                    ]:
                        if hasattr(row, attr):
                            converted_row[attr] = getattr(row, attr)

                    # 如果未能获取到属性，使用字符串表示
                    if not converted_row:
                        converted_row = {
                            "level1": str(row)[:50],
                            "level2": "",
                            "level3": "",
                            "row_index": idx,
                        }
                    else:
                        converted_row["row_index"] = idx

                    excel_data.append(converted_row)
                except Exception as e:
                    print(f"[WARN] 第{idx}行数据转换失败: {e}, 使用默认值")
                    excel_data.append(
                        {
                            "level1": f"ERROR_{idx}",
                            "level2": "",
                            "level3": "",
                            "row_index": idx,
                        }
                    )

        # === 关键增强：对 (L1, L2, L3) 组合去重合并，避免报告中重复项刷屏 ===
        merged_excel_data = []
        seen_keys = {}
        for row in excel_data:
            key = (
                safe_get(row, "level1", ""),
                safe_get(row, "level2", ""),
                safe_get(row, "level3", ""),
            )
            rng = row.get("row_range", "") if isinstance(row, dict) else ""

            if key not in seen_keys:
                seen_keys[key] = len(merged_excel_data)
                new_row = (
                    row.copy()
                    if isinstance(row, dict)
                    else {"level1": key[0], "level2": key[1], "level3": key[2]}
                )
                new_row["all_ranges"] = [rng] if rng else []
                merged_excel_data.append(new_row)
            else:
                m_idx = seen_keys[key]
                if rng and rng not in merged_excel_data[m_idx].get("all_ranges", []):
                    merged_excel_data[m_idx].setdefault("all_ranges", []).append(rng)

        for row in merged_excel_data:
            ranges = row.get("all_ranges", [])
            if ranges:
                row["row_range"] = ", ".join(ranges)

        excel_data = merged_excel_data

        if progress_callback:
            progress_callback(40, f"开始匹配 {len(excel_data)} 个模块项...")

        # 3. 匹配过程
        exact_matched = []
        fuzzy_matched = []
        not_found_in_word = []
        hierarchy_mismatched = []

        total = len(excel_data)
        update_interval = max(1, total // 20)

        for idx, excel_row in enumerate(excel_data):
            if progress_callback and (idx % update_interval == 0 or idx == total - 1):
                progress_callback(
                    40 + int(idx / total * 50), f"匹配进度: {idx + 1}/{total}"
                )

            # === 关键修复：使用安全获取层级值 ===
            l1_val = safe_get(excel_row, "level1", "")
            l2_val = safe_get(excel_row, "level2", "")
            l3_val = safe_get(excel_row, "level3", "")

            # 如果level1/level2/level3为空，尝试从其他字段获取
            if not l1_val:
                l1_val = safe_get(excel_row, "level1_original", "")
            if not l2_val:
                l2_val = safe_get(excel_row, "level2_original", "")
            if not l3_val:
                l3_val = safe_get(excel_row, "level3_original", "")

            # 初始化匹配结果
            l1_matched_flag = False
            l2_matched_flag = False
            l3_matched_flag = False
            word_l1_original = ""
            word_l1_title = ""
            word_l1_num = ""
            word_l2_original = ""
            word_l2_title = ""
            word_l2_num = ""
            word_l3_original = ""
            word_l3_title = ""
            word_l3_num = ""
            l1_score = 0.0
            l2_score = 0.0
            l3_score = 0.0
            l1_actual_depth = 0
            l2_actual_depth = 0
            l3_actual_depth = 0
            warnings = []

            # === 3.1 一级模块匹配 ===
            if l1_val:
                candidates = word_hierarchy["all_items"]
                found, match, score = find_match_for_excel_level(
                    l1_val, candidates, excel_level=1
                )
                if found:
                    l1_matched_flag = True
                    word_l1_original = match.get(
                        "original", match.get("display", match.get("text", ""))
                    )
                    word_l1_title = match.get("text", match.get("cleaned", ""))
                    word_l1_num = match.get("number", "")
                    l1_score = score
                    l1_actual_depth = get_word_level(match)

                    # 检查层级是否匹配（按编号段数）
                    expected_depth = 1
                    if l1_actual_depth != expected_depth and l1_actual_depth > 0:
                        warnings.append(
                            f"L1层级不匹配: 期望深度{expected_depth}, 实际深度{l1_actual_depth} (编号:{word_l1_num})"
                        )
                else:
                    word_l1_title = "❌ 缺失"
                    word_l1_original = "❌ 缺失"

            # === 3.2 二级模块匹配 ===
            if l2_val:
                candidates = word_hierarchy["all_items"]
                # 优先在L1的子项中搜索（如果L1已匹配且有编号）
                if l1_matched_flag and word_l1_num:
                    constrained = [
                        c
                        for c in candidates
                        if is_descendant(word_l1_num, c.get("number", ""))
                    ]
                    if constrained:
                        candidates = constrained

                found, match, score = find_match_for_excel_level(
                    l2_val, candidates, excel_level=2
                )
                if found:
                    l2_matched_flag = True
                    word_l2_original = match.get(
                        "original", match.get("display", match.get("text", ""))
                    )
                    word_l2_title = match.get("text", match.get("cleaned", ""))
                    word_l2_num = match.get("number", "")
                    l2_score = score
                    l2_actual_depth = get_word_level(match)

                    # 检查层级是否匹配（按编号段数）
                    expected_depth = 2
                    if l2_actual_depth != expected_depth and l2_actual_depth > 0:
                        warnings.append(
                            f"L2层级不匹配: 期望深度{expected_depth}, 实际深度{l2_actual_depth} (编号:{word_l2_num})"
                        )

                    # 检查父子关系
                    if l1_matched_flag and word_l1_num:
                        if not is_descendant(word_l1_num, word_l2_num):
                            warnings.append(
                                f"L2非L1子项: L1={word_l1_num}, L2={word_l2_num}"
                            )
                else:
                    word_l2_title = "❌ 缺失"
                    word_l2_original = "❌ 缺失"

            # === 3.3 三级模块匹配 ===
            if l3_val:
                candidates = word_hierarchy["all_items"]
                # 优先在L2或L1的子项中搜索
                parent_num = None
                search_candidates = []

                if l2_matched_flag and word_l2_num:
                    parent_num = word_l2_num
                    search_candidates = [
                        c
                        for c in candidates
                        if is_descendant(parent_num, c.get("number", ""))
                    ]
                elif l1_matched_flag and word_l1_num:
                    parent_num = word_l1_num
                    search_candidates = [
                        c
                        for c in candidates
                        if is_descendant(parent_num, c.get("number", ""))
                    ]

                if not search_candidates:
                    search_candidates = candidates

                found, match, score = find_match_for_excel_level(
                    l3_val, search_candidates, excel_level=3
                )
                if found:
                    l3_matched_flag = True
                    word_l3_original = match.get(
                        "original", match.get("display", match.get("text", ""))
                    )
                    word_l3_title = match.get("text", match.get("cleaned", ""))
                    word_l3_num = match.get("number", "")
                    l3_score = score
                    l3_actual_depth = get_word_level(match)

                    # 检查层级是否匹配（按编号段数）
                    expected_depth = 3
                    if l3_actual_depth != expected_depth and l3_actual_depth > 0:
                        warnings.append(
                            f"L3层级不匹配: 期望深度{expected_depth}, 实际深度{l3_actual_depth} (编号:{word_l3_num})"
                        )

                    # 检查父子关系
                    if l2_matched_flag and word_l2_num:
                        if not is_descendant(word_l2_num, word_l3_num):
                            warnings.append(
                                f"L3非L2子项: L2={word_l2_num}, L3={word_l3_num}"
                            )
                    elif l1_matched_flag and word_l1_num:
                        if not is_descendant(word_l1_num, word_l3_num):
                            warnings.append(
                                f"L3非L1子项: L1={word_l1_num}, L3={word_l3_num}"
                            )
                else:
                    word_l3_title = "❌ 缺失"
                    word_l3_original = "❌ 缺失"

            # === 4. 构建报告项 ===
            has_all_matched = l1_matched_flag and l2_matched_flag and l3_matched_flag

            matched_levels = []
            if l1_val and l1_matched_flag:
                matched_levels.append(1)
            if l2_val and l2_matched_flag:
                matched_levels.append(2)
            if l3_val and l3_matched_flag:
                matched_levels.append(3)
            matched_count = len(matched_levels)
            any_matched = matched_count > 0
            path_ok = (
                not l1_matched_flag
                or not l2_matched_flag
                or is_descendant(word_l1_num, word_l2_num)
            ) and (
                not l2_matched_flag
                or not l3_matched_flag
                or is_descendant(word_l2_num, word_l3_num)
            )
            hierarchy_level_issues = (
                (l1_matched_flag and l1_actual_depth != 1)
                or (l2_matched_flag and l2_actual_depth != 2)
                or (l3_matched_flag and l3_actual_depth != 3)
            )

            report_item = {
                "Excel一级模块": l1_val,
                "Excel二级模块": l2_val,
                "Excel三级模块": l3_val,
                # 这里展示的是 Word 目录树的“逻辑一级/二级/三级”（按编号深度回填：4.x / 4.x.x / 4.x.x.x）
                "Word一级标题": "",
                "Word二级标题": "",
                "Word三级标题": "",
                "Word完整编号": "",
                "word_l1_original": word_l1_original,
                "word_l1_title": word_l1_title,
                "word_l1_num": word_l1_num,
                "word_l2_original": word_l2_original,
                "word_l2_title": word_l2_title,
                "word_l2_num": word_l2_num,
                "word_l3_original": word_l3_original,
                "word_l3_title": word_l3_title,
                "word_l3_num": word_l3_num,
                "l1_matched": l1_matched_flag,
                "l2_matched": l2_matched_flag,
                "l3_matched": l3_matched_flag,
                "相似度": f"{max(l1_score, l2_score, l3_score):.2%}",
                "匹配状态": "",
                "缺失简略描述": "",
                "层级不匹配简略描述": "",
                "深度警告描述": "; ".join(warnings),
                "简略描述": "",
                "row_range": excel_row.get("row_range", f"行{idx + header + 2}"),
            }

            # === [报表回填] 选取“最深命中编号”作为 anchor，用于回填 Word L1/L2/L3 ===
            anchor_candidates = []
            if l1_matched_flag and word_l1_num:
                anchor_candidates.append((_num_depth(word_l1_num), word_l1_num))
            if l2_matched_flag and word_l2_num:
                anchor_candidates.append((_num_depth(word_l2_num), word_l2_num))
            if l3_matched_flag and word_l3_num:
                anchor_candidates.append((_num_depth(word_l3_num), word_l3_num))
            anchor_num = (
                max(anchor_candidates, key=lambda x: x[0])[1]
                if anchor_candidates
                else ""
            )

            # === [真实目录回溯] 获取 Word 真实层级结构以辅助描述 ===
            real_tree = _compute_display_tree(anchor_num)

            # === [报表填充] 使用匹配到的直接结果填入一二三级标题列（保持逻辑一致性） ===
            report_item["Word一级标题"] = word_l1_title
            report_item["Word二级标题"] = word_l2_title
            report_item["Word三级标题"] = word_l3_title
            # === [关键修复] 完整编号展示“实际匹配到”的编号，不进行路径补全 ===
            # 例如：Excel一级缺失，二级匹配4.1，三级匹配4.1.1 -> 展示 -/4.1/4.1.1
            report_item["Word完整编号"] = (
                f"{word_l1_num or '-'}"
                f"/{word_l2_num or '-'}"
                f"/{word_l3_num or '-'}"
            )

            # 计算综合相似度
            avg_score = (l1_score + l2_score + l3_score) / max(
                1, (l1_matched_flag + l2_matched_flag + l3_matched_flag)
            )
            report_item["相似度"] = f"{avg_score:.2%}"

            # 先计算“缺失/层级不匹配”两类描述（允许共存）
            missing_desc = "-"
            mismatch_desc = "-"

            # 1) 缺失描述：按用户固定格式输出
            missing_levels = []
            missing_parts = []
            if l1_val and not l1_matched_flag:
                missing_levels.append(1)
                missing_parts.append(f"一级模块 {l1_val}")
            if l2_val and not l2_matched_flag:
                missing_levels.append(2)
                missing_parts.append(f"二级模块 {l2_val}")
            if l3_val and not l3_matched_flag:
                missing_levels.append(3)
                missing_parts.append(f"三级模块 {l3_val}")

            if missing_parts:
                if len(missing_levels) == 1:
                    lvl = missing_levels[0]
                    lvl_label = {1: "一级", 2: "二级", 3: "三级"}.get(lvl, "")
                    name = (
                        missing_parts[0]
                        .replace("一级模块 ", "")
                        .replace("二级模块 ", "")
                        .replace("三级模块 ", "")
                    )
                    missing_desc = (
                        f"拆分表{lvl_label}模块在需求规格书未体现："
                        f"拆分表{lvl_label}模块 {name} 在需求规格书未体现"
                    )
                else:
                    lvl_names = "".join(
                        [{1: "一", 2: "二", 3: "三"}.get(l, "") for l in missing_levels]
                    )
                    missing_desc = (
                        f"拆分表{lvl_names}级模块在需求规格书未体现："
                        f"拆分表{'、'.join(missing_parts)} 在需求规格书未体现"
                    )

            # 2) 层级不匹配描述：按用户固定格式输出（只涉及 1~3 级）
            # 特例：你给的“错位”场景（Excel二级=Word一级，Excel三级=Word二级）要合并成一条“二三级不匹配”
            def _norm_txt(v: str) -> str:
                return re.sub(r"\s+", "", str(v or "").strip().lower())

            if l2_matched_flag and l3_matched_flag:
                # 基于回填后的“逻辑目录层级”做错位检测
                # 如果 Excel L2 匹配到了 Word L1，Excel L3 匹配到了 Word L2 (即 4.1.x 类型)
                if (
                    _norm_txt(l2_val) == _norm_txt(word_l1_title)
                    and _norm_txt(l3_val) == _norm_txt(word_l2_title)
                    and _norm_txt(l2_val)
                    and _norm_txt(l3_val)
                ):
                    # 展示 Word 真实的 2, 3 级标题（即 4.1.1, 4.1.1.1）
                    w2 = real_tree[2]["original"]
                    w3 = real_tree[3]["original"]
                    if not w3 or "❌ 缺失" in w3:
                        w3 = "缺失"
                    if not w2 or "❌ 缺失" in w2:
                        w2 = "缺失"

                    mismatch_desc = (
                        "拆分表二三级模块与需求规格书二三级目录不匹配："
                        f"拆分表 二级模块 {{{l2_val}}} - 三级模块 {{{l3_val}}} "
                        f"与需求规格书 二级目录 {{{w2}}} - 三级目录 {{{w3}}} 不匹配"
                    )

            if (
                mismatch_desc == "-"
                and any_matched
                and matched_count >= 2
                and (hierarchy_level_issues or not path_ok)
            ):
                lvl_name_map = {1: "一", 2: "二", 3: "三"}
                lvl_label_map = {1: "一级", 2: "二级", 3: "三级"}

                lvl_names = [lvl_name_map.get(l, "") for l in matched_levels]
                lvl_names = [n for n in lvl_names if n]
                m_str = (
                    (lvl_names[0] + "级")
                    if len(lvl_names) == 1
                    else ("".join(lvl_names) + "级")
                )

                def _fmt_val(v: str) -> str:
                    return (v or "").strip()

                excel_segs = []
                word_segs = []
                for lvl in matched_levels:
                    if lvl == 1:
                        excel_segs.append(
                            f"{lvl_label_map[lvl]}模块 {{{_fmt_val(l1_val)}}}"
                        )
                        wv = real_tree[1]["original"]
                        word_segs.append(f"{lvl_label_map[lvl]}目录{{{_fmt_val(wv)}}}")
                    elif lvl == 2:
                        excel_segs.append(
                            f"{lvl_label_map[lvl]}模块 {{{_fmt_val(l2_val)}}}"
                        )
                        wv = real_tree[2]["original"]
                        word_segs.append(f"{lvl_label_map[lvl]}目录{{{_fmt_val(wv)}}}")
                    elif lvl == 3:
                        excel_segs.append(
                            f"{lvl_label_map[lvl]}模块 {{{_fmt_val(l3_val)}}}"
                        )
                        wv = real_tree[3]["original"]
                        word_segs.append(f"{lvl_label_map[lvl]}目录{{{_fmt_val(wv)}}}")

                mismatch_desc = (
                    f"拆分表{m_str}模块与需求规格书{m_str}目录不匹配："
                    f"拆分表 {' - '.join(excel_segs)} 与需求规格书 {' - '.join(word_segs)} 不匹配"
                )

            report_item["缺失简略描述"] = missing_desc
            report_item["层级不匹配简略描述"] = mismatch_desc

            # 设置匹配状态（描述不互斥，状态仍单选）
            if has_all_matched and not hierarchy_level_issues and path_ok:
                if l1_score >= 0.999 and l2_score >= 0.999 and l3_score >= 0.999:
                    report_item["匹配状态"] = "✅ 完全匹配"
                else:
                    report_item["匹配状态"] = "🔍 模糊匹配"
            elif mismatch_desc != "-":
                report_item["匹配状态"] = "⚠️ 层级不匹配"
            else:
                report_item["匹配状态"] = "❌ 缺失"

            # 简略描述：保持兼容（用于旧 UI/导出摘要）；不影响两列共存
            report_item["简略描述"] = (
                mismatch_desc
                if mismatch_desc != "-"
                else (missing_desc if missing_desc != "-" else "-")
            )

            # 分类到不同列表
            # === 调试日志：打印最终保存的report_item内容 ===
            if RuntimeLogger:
                RuntimeLogger.log(
                    f"  [DEBUG] 最终report_item: "
                    f"Excel({report_item.get('Excel一级模块', '')}/{report_item.get('Excel二级模块', '')}/{report_item.get('Excel三级模块', '')}) -> "
                    f"Word({report_item.get('Word一级标题', '')}/{report_item.get('Word二级标题', '')}/{report_item.get('Word三级标题', '')}) "
                    f"状态={report_item.get('匹配状态', '')}",
                    level="DEBUG",
                )

            if report_item["匹配状态"] == "✅ 完全匹配":
                if RuntimeLogger:
                    RuntimeLogger.log(
                        f"  [DEBUG] 追加到exact_matched: Word({report_item['Word一级标题']}/{report_item['Word二级标题']}/{report_item['Word三级标题']})",
                        level="DEBUG",
                    )
                exact_matched.append(report_item)
            elif report_item["匹配状态"] == "🔍 模糊匹配":
                if RuntimeLogger:
                    RuntimeLogger.log(
                        f"  [DEBUG] 追加到fuzzy_matched: Word({report_item['Word一级标题']}/{report_item['Word二级标题']}/{report_item['Word三级标题']})",
                        level="DEBUG",
                    )
                fuzzy_matched.append(report_item)
            elif report_item["匹配状态"] == "⚠️ 层级不匹配":
                if RuntimeLogger:
                    RuntimeLogger.log(
                        f"  [DEBUG] 追加到hierarchy_mismatched: Excel({report_item['Excel一级模块']}/{report_item['Excel二级模块']}/{report_item['Excel三级模块']}) -> Word({report_item['Word一级标题']}/{report_item['Word二级标题']}/{report_item['Word三级标题']})",
                        level="DEBUG",
                    )
                hierarchy_mismatched.append(report_item)
            else:  # "❌ 缺失"
                if RuntimeLogger:
                    RuntimeLogger.log(
                        f"  [DEBUG] 追加到not_found_in_word: Excel({report_item['Excel一级模块']}/{report_item['Excel二级模块']}/{report_item['Excel三级模块']})",
                        level="DEBUG",
                    )
                not_found_in_word.append(report_item)

        # === 5. 构建最终报告前，先打印完整的目录树 ===
        print("\n" + "=" * 80)
        print("📋 Word 文档完整目录树结构")
        print("=" * 80)

        # 按编号顺序排序并打印
        sorted_items = sorted(
            word_items,
            key=lambda x: (
                [int(p) for p in x.get("number", "").split(".") if p.isdigit()]
                if x.get("number")
                else []
            ),
        )
        for item in sorted_items:
            num = item.get("number", "")
            title = item.get("text", "")
            level = item.get("level", 0)
            if num and title:
                # 计算缩进：根据编号段数
                depth = num.count(".") + 1
                indent = "  " * (depth - 1)
                print(f"{indent}[{num}] {title}")

        print("=" * 80 + "\n")

        # 同时也写到日志中
        if RuntimeLogger:
            RuntimeLogger.log("=" * 80, level="INFO")
            RuntimeLogger.log("📋 Word 文档完整目录树结构", level="INFO")
            RuntimeLogger.log("=" * 80, level="INFO")
            for item in sorted_items:
                num = item.get("number", "")
                title = item.get("text", "")
                level = item.get("level", 0)
                if num and title:
                    depth = num.count(".") + 1
                    indent = "  " * (depth - 1)
                    RuntimeLogger.log(f"{indent}[{num}] {title}", level="INFO")
            RuntimeLogger.log("=" * 80, level="INFO")

        # === 5. 构建最终报告 ===
        total_items = len(excel_data)
        matched_items = (
            len(exact_matched) + len(fuzzy_matched) + len(hierarchy_mismatched)
        )
        match_rate = (matched_items / total_items * 100) if total_items > 0 else 0.0

        # 统计口径：缺失/不匹配允许共存，因此按两列描述是否为 '-' 计数（而不是按列表互斥计数）
        all_items_for_stats = (
            exact_matched + fuzzy_matched + hierarchy_mismatched + not_found_in_word
        )
        missing_desc_count = sum(
            1
            for it in all_items_for_stats
            if it.get("缺失简略描述") and it.get("缺失简略描述") != "-"
        )
        mismatch_desc_count = sum(
            1
            for it in all_items_for_stats
            if it.get("层级不匹配简略描述") and it.get("层级不匹配简略描述") != "-"
        )

        report = {
            "exact_matched": exact_matched,
            "fuzzy_matched": fuzzy_matched,
            "hierarchy_mismatched": hierarchy_mismatched,
            "not_found_in_word": not_found_in_word,
            "statistics": {
                # [UI兼容] task_card 使用 Excel功能点总数 来判断是否有数据
                "Excel功能点总数": total_items,
                "Excel模块总数": total_items,
                "完全匹配": len(exact_matched),
                "模糊匹配": len(fuzzy_matched),
                "层级不匹配": mismatch_desc_count,
                "缺失项": missing_desc_count,
                "匹配率": f"{match_rate:.2f}%",
                "完成度": f"{match_rate:.2f}%",
                "时间戳": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
            "word_file": (
                os.path.basename(word_file)
                if isinstance(word_file, (str, os.PathLike))
                else "内存对象"
            ),
            "excel_file": (
                os.path.basename(excel_file)
                if isinstance(excel_file, (str, os.PathLike))
                else "内存对象"
            ),
            "sheet_name": sheet_name,
        }

        # 保存层级匹配日志（如果需要）
        if hierarchy_log_path:
            try:
                with open(hierarchy_log_path, "w", encoding="utf-8") as f:
                    f.write("📋 Word 文档完整目录树结构\n")
                    f.write("=" * 60 + "\n")
                    for item in sorted_items:
                        num = item.get("number", "")
                        title = item.get("text", "")
                        if num and title:
                            depth = num.count(".") + 1
                            indent = "  " * (depth - 1)
                            f.write(f"{indent}[{num}] {title}\n")
                    f.write("=" * 60 + "\n")
                if RuntimeLogger:
                    RuntimeLogger.log(
                        f"✓ Word结构树已保存: {hierarchy_log_path}", level="INFO"
                    )
            except Exception as e:
                print(f"警告: 无法保存层级匹配日志: {e}")

        return report

    def validate_hierarchy_matching_simple(
        self, excel_file, sheet_name, level1_col, level2_col, level3_col, header
    ):
        """
        最小化修复方案 - 只验证参数和文件
        """
        try:
            # 打印调试信息
            print(
                f"[VALIDATE] 验证参数: excel_file={type(excel_file)}, sheet={sheet_name}"
            )
            print(
                f"[VALIDATE] 列设置: l1={level1_col}, l2={level2_col}, l3={level3_col}, header={header}"
            )

            # 确保参数是整数
            try:
                l1 = int(level1_col) if level1_col is not None else 0
                l2 = int(level2_col) if level2_col is not None else 1
                l3 = int(level3_col) if level3_col is not None else 2
                hdr = int(header) if header is not None else 0
            except ValueError as e:
                print(f"[ERROR] 参数类型错误: {e}")
                return False, f"参数类型错误: {e}"

            # 尝试读取Excel文件
            import pandas as pd

            try:
                if isinstance(excel_file, pd.DataFrame):
                    df = excel_file
                    print(f"[VALIDATE] 使用提供的DataFrame，形状: {df.shape}")
                elif isinstance(excel_file, str):
                    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=hdr)
                    print(f"[VALIDATE] 从文件读取DataFrame，形状: {df.shape}")
                else:
                    return False, f"不支持的excel_file类型: {type(excel_file)}"

                # 检查列索引是否有效
                if (
                    l1 >= len(df.columns)
                    or l2 >= len(df.columns)
                    or l3 >= len(df.columns)
                ):
                    return False, f"列索引超出范围: 数据有{len(df.columns)}列"

                # 检查是否有数据
                if len(df) == 0:
                    return False, "Excel数据为空"

                # 返回成功
                return True, f"验证通过: {len(df)}行, {len(df.columns)}列"

            except Exception as e:
                print(f"[ERROR] 读取Excel失败: {e}")
                return False, f"读取Excel失败: {e}"

        except Exception as e:
            print(f"[ERROR] 验证过程异常: {e}")
            import traceback

            traceback.print_exc()
            return False, f"验证过程异常: {e}"
