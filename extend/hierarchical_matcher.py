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
from fuzzywuzzy import fuzz

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


# -----------------------------------------------------------------------------
# HierarchicalMatcher 类
# -----------------------------------------------------------------------------
class HierarchicalMatcher:
    """层级匹配器类"""

    # 业务章节常量
    SECTION_FUNCTIONAL_REQUIREMENTS = "功能需求"
    CHAPTER_NUMBER_PREFIX = "4"  # 默认功能需求从第4章开始

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
        word_file: str,
        mode: str = "flat",
        progress_callback=None,
        full_text_search: bool = False,
    ) -> List[Dict]:
        """
        统一提取 Word 内容的入口。

        Args:
            word_file: 文件路径
            mode: 'flat' (简单列表) 或 'hierarchical' (层级字典)
            progress_callback: 进度回调 (percent, message)
            full_text_search: 是否提取全文 (仅在 flat 模式下有效)
        """
        if progress_callback:
            progress_callback(5, "正在打开 Word 文档...")

        doc, processed_path, is_temp = self._open_word_doc(word_file)

        try:
            if mode == "flat":
                return self._extract_word_flat(doc, progress_callback, full_text_search)
            else:
                # 兼容原有 word_file 传递逻辑
                if str(processed_path) == "PRELOADED_DOCX":
                    # 如果是预加载的，word_file 实际上没法通过 doc 获取路径，
                    # 除非 doc 对象带了路径。这里我们还是希望调用者传 word_file 为原路径
                    pass

                return self._extract_word_hierarchical(
                    doc,
                    progress_callback,
                    word_file=word_file if isinstance(word_file, (str, Path)) else None,
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
        直接利用 DocumentProcessor 提取的目录树结构及其自动识别的编号层级。
        不再使用传统的大纲计数策略或显式匹配策略。
        """
        hierarchy = {"level1": [], "level2": [], "level3": []}

        # 优先使用已打开的 doc 对象，如果没有则尝试 word_file 路径
        target_to_extract = doc if doc is not None else word_file

        if not target_to_extract:
            print("  [!] 警告: 未提供文档对象或路径，无法使用智能提取引擎")
            return hierarchy

        try:
            from utils.document_processor import DocumentProcessor

            # 获取已经过 DocumentProcessor（模拟 WPS 智能目录）识别处理后的完整结构
            # DocumentProcessor.extract_word_structure 已支持传入路径或 doc 对象
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
                    "depth": lvl,
                    "order": i,  # 添加原始顺序索引
                    "in_toc": True,  # 智能提取的都是目录层级项
                    "level": "heading",  # 标记为标题样式
                }

                # 层级映射 (对接 WPS 智能树)：
                # Word Level 1 (4.) -> Excel 一级模块
                # Word Level 2 (4.1.) -> Excel 二级模块
                # Word Level 3/4/5 (4.1.1.) -> Excel 三级模块
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

        # 移除原有的 _extract_word_hierarchical 的其余冗余实现（Strategy 1 & 2）

        # --- DEBUG ---
        print("\n[DEBUG] 最终提取结果统计:")
        print(f"  Level 1 (Word H2): {len(final_hierarchy['level1'])}")
        if final_hierarchy["level1"]:
            print(f"    示例: {[i['display'] for i in final_hierarchy['level1'][:2]]}")
        print(f"  Level 2 (Word H3): {len(final_hierarchy['level2'])}")
        if final_hierarchy["level2"]:
            print(f"    示例: {[i['display'] for i in final_hierarchy['level2'][:2]]}")
        print(f"  Level 3 (Word H4/H5): {len(final_hierarchy['level3'])}")
        if final_hierarchy["level3"]:
            print(f"    示例: {[i['display'] for i in final_hierarchy['level3'][:2]]}")

        return final_hierarchy

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
                df_ffill, df_raw, kwargs.get("column", 0), header_row
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
        self, df: pd.DataFrame, df_raw: pd.DataFrame, column: int, header: int
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

        # 调整模式，使单数字匹配更严格，避免匹配“4A系统”
        # 同时支持全角点 ．
        patterns = [
            r"^(\d+[\.\．]\d+[\.\．]\d+[\.\．]\d+)",  # 4.1.1.1
            r"^(\d+[\.\．]\d+[\.\．]\d+)",  # 4.1.1
            r"^(\d+[\.\．]\d+)",  # 4.1
            r"^(\d+)(?=[\.\．]|\s|$)",  # 4. 或 4
        ]

        for pattern in patterns:
            match = re.match(pattern, text_to_match)
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
    ) -> Dict:
        """简单匹配模式（单列匹配）"""

        # 1. 统一提取 Word 内容
        if word_items_preloaded is not None:
            # [FIX] 转换预加载的层级数据为扁平格式
            word_items = []
            for s in word_items_preloaded:
                # 提取标题项
                title_text = s.get("title", "")
                if title_text:
                    cleaned = self.basic_clean(title_text)
                    word_items.append(
                        {
                            "original": title_text,
                            "text": cleaned,
                            "cleaned": cleaned,
                            "in_toc": True,
                            "level": s.get("level", 0),
                        }
                    )

                # 如果开启了全文搜索，还需要提取内容项 (content 列表中的每一行)
                if full_text_search:
                    for line in s.get("content", []):
                        if not line.strip():
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
                            }
                        )
        else:
            word_items = self.extract_word_content(
                word_file,
                mode="flat",
                progress_callback=progress_callback,
                full_text_search=full_text_search,
            )

        if progress_callback:
            progress_callback(30, "正在读取 Excel 数据...")

        # 2. 统一提取 Excel 内容
        excel_data_raw = self.extract_excel_content(
            excel_file, mode="flat", sheet_name=sheet_name, header=header, column=column
        )

        # --- 去重逻辑 ---
        excel_data = []
        seen_texts = set()
        for item in excel_data_raw:
            text = item.get("text", item.get("original", ""))
            if text and text not in seen_texts:
                excel_data.append(item)
                seen_texts.add(text)

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

        exact_matched = []
        fuzzy_matched = []
        not_found_in_word = []

        total = len(excel_data)
        update_interval = max(1, total // 20)

        import concurrent.futures

        for idx, excel_item in enumerate(excel_data):
            if progress_callback and (idx % update_interval == 0 or idx == total - 1):
                progress = 30 + int((idx / (total if total > 0 else 1)) * 60)
                progress_callback(progress, f"匹配中... {idx + 1}/{total}")

            excel_text_for_matching = self.basic_clean(excel_item["original"])
            found, word_match, similarity_score = self.find_match(
                excel_text_for_matching, word_items, 0
            )

            if found:
                is_in_toc = word_match.get("in_toc", True)
                location = "目录" if is_in_toc else "正文"

                match_data = {
                    "Excel功能点": excel_item["original"],
                    "Word匹配项": word_match.get(
                        "display", word_match.get("original", "")
                    ),
                    "位置": location,
                    "相似度": f"{similarity_score:.2%}",
                }

                if similarity_score >= 0.99:
                    exact_matched.append(match_data)
                else:
                    fuzzy_matched.append(match_data)
            else:
                match_data = {
                    "Excel功能点": excel_item["original"],
                    "状态": "[FAIL] 缺失",
                    "简略描述": f"拆分表{excel_item['original']}在需求规格书未体现",
                }
                not_found_in_word.append(match_data)

        if progress_callback:
            progress_callback(90, "生成统计报告...")

        total_excel = len(excel_data)
        exact_count = len(exact_matched)
        fuzzy_count = len(fuzzy_matched)
        matched_count = exact_count + fuzzy_count
        missing_count = len(not_found_in_word)
        match_rate = (matched_count / total_excel * 100) if total_excel > 0 else 0

        print(
            f"\n匹配完成：精确 {exact_count} 项，模糊 {fuzzy_count} 项，缺失 {missing_count} 项"
        )
        if full_text_search:
            toc_count = sum(
                1
                for item in exact_matched + fuzzy_matched
                if item.get("位置") == "目录"
            )
            body_count = sum(
                1
                for item in exact_matched + fuzzy_matched
                if item.get("位置") == "正文"
            )
            print(f"  其中：目录中 {toc_count} 项，正文中 {body_count} 项")

        # 移除过多的终端打印，保持界面干净
        if progress_callback:
            progress_callback(100, "匹配完成！")

        report = {
            "exact_matched": exact_matched,
            "fuzzy_matched": fuzzy_matched,
            "not_found_in_word": not_found_in_word,
            "statistics": {
                "Excel功能点总数": total_excel,
                "Word内容项总数": len(word_items),
                "精确匹配": exact_count,
                "模糊匹配": fuzzy_count,
                "已匹配": matched_count,
                "缺失项": missing_count,
                "匹配率": f"{match_rate:.2f}%",
                "完成度": f"{match_rate:.2f}%",
            },
        }

        if full_text_search:
            report["statistics"]["目录中找到"] = toc_count
            report["statistics"]["正文中找到"] = body_count

        return report

    def similarity(self, a: str, b: str) -> float:
        """计算两个字符串的相似度"""
        a_norm = a.lower()
        b_norm = b.lower()

        # 如果开启了模糊匹配，对常见近义词进行归一化处理（仅用于相似度衡量）
        if self.fuzzy_match:
            a_norm = a_norm.replace("账户", "账号")
            b_norm = b_norm.replace("账户", "账号")

        return SequenceMatcher(None, a_norm, b_norm).ratio()

    def find_match(
        self, target: str, candidates: List[Dict], target_depth: int = 0
    ) -> Tuple[bool, Dict, float]:
        """
        在候选集中查找最佳匹配项。
        不仅比较相似度，还在得分相同时优先选择目录项或标题项。
        """
        best_match = None
        best_score = -1.0
        best_quality = -1  # 用于打破得分平局的质量分

        # 预处理：去除目标字符串的所有空格（用于比较）
        if target is None:
            target = ""
        target_no_space = re.sub(r"\s+", "", str(target).lower())

        # 预处理：目标字符集（用于快筛）
        target_chars = set(target_no_space)
        target_len = len(target_no_space)

        for candidate in candidates:
            # 优化 1：缓存无空格文本及字符集
            candidate_no_space = candidate.get("text_no_space")
            candidate_text = candidate.get("text", candidate.get("cleaned", ""))

            if candidate_no_space is None:
                candidate_no_space = re.sub(r"\s+", "", candidate_text.lower())
                candidate["text_no_space"] = candidate_no_space
                candidate["char_set"] = set(candidate_no_space)

            candidate_chars = candidate.get("char_set")
            if candidate_chars is None:
                candidate_chars = set(candidate_no_space)
                candidate["char_set"] = candidate_chars

            # 优化 2：【重点】字集快筛 (Fast Jaccard-like filter)
            # 如果目标较长，且两者字符重合度极低，直接跳过昂贵的相似度计算
            if target_len > 5:
                overlap = len(target_chars.intersection(candidate_chars))
                # 经验阈值：字符重合率低于 30% 且不包含全字，通常相似度无法达到 0.7
                if (
                    overlap / target_len < 0.3
                    and target_no_space not in candidate_no_space
                ):
                    continue

            # 1. 精确匹配（带空格）
            if target == candidate_text:
                current_score = 1.0
            # 2. 精确匹配（忽略空格）
            elif target_no_space == candidate_no_space:
                current_score = 0.995
            # 3. 模糊匹配
            elif self.fuzzy_match:
                score_with_space = self.similarity(target, candidate_text)
                score_no_space = self.similarity(target_no_space, candidate_no_space)
                current_score = max(score_with_space, score_no_space)

            # 4. 包含关系匹配 (子串/前缀奖励)
            if len(target_no_space) >= 3:
                # 情况 A: Target 是 Candidate 的前缀
                if candidate_no_space.startswith(target_no_space):
                    current_score = max(current_score, 0.985)
                # 情况 B: Candidate 以 Target 开头，但中间可能有微小字符差异
                elif (
                    self.similarity(
                        target_no_space, candidate_no_space[: len(target_no_space)]
                    )
                    > 0.9
                ):
                    current_score = max(current_score, 0.95)
                # 情况 C: Target 包含在 Candidate 中
                elif target_no_space in candidate_no_space:
                    overlap_ratio = len(target_no_space) / len(candidate_no_space)
                    current_score = max(current_score, 0.85, overlap_ratio)

            # 只有达到阈值才考虑
            if current_score >= self.threshold:
                # 计算质量分用于打破平局
                # 质量分权重：在目录中 +10, 是标题 +5, 长度匹配度 (0.0~1.0)
                quality = 0
                if candidate.get("in_toc"):
                    quality += 10

                style = str(candidate.get("level", "")).lower()
                if "heading" in style or "标题" in style:
                    quality += 5

                # 长度匹配度：目标长度与候选长度越接近越好 (避免长段落包含短词导致的误匹配)
                len_ratio = min(len(target_no_space), len(candidate_no_space)) / max(
                    len(target_no_space), len(candidate_no_space)
                )

                # 总比较逻辑：优先看相似度，相似度极近时看质量分
                # 使用一个小增量 epsilon 来判断相似度是否“基本相等”
                if current_score > best_score + 1e-5:
                    best_score = current_score
                    best_match = candidate
                    best_quality = quality + len_ratio
                elif abs(current_score - best_score) <= 1e-5:
                    if (quality + len_ratio) > best_quality:
                        best_score = current_score
                        best_match = candidate
                        best_quality = quality + len_ratio

        if best_match and best_score >= self.threshold:
            return True, best_match, best_score

        return False, {}, 0.0

    # 简单匹配模式的 match_documents 方法保持不变，因为它与层级匹配不同
    # 它的逻辑已经包含在原始代码中，这里不再重复

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
        word_sections_preloaded=None,  # [NEW] 层级匹配预置数据
    ) -> Dict:
        """层级匹配模式 (增强版)"""
        # 1. 统一提取 Word 内容
        if word_sections_preloaded is not None:
            # 这里的 word_sections_preloaded 是 DocumentProcessor.extract_word_structure 返回的格式
            # 需要将其转换为 HierarchicalMatcher 内部使用的三层层级字典格式
            word_hierarchy = {"level1": [], "level2": [], "level3": []}
            for idx, s in enumerate(word_sections_preloaded):
                lvl = s.get("level", 0)
                item = {
                    "number": s.get("num", ""),
                    "text": s.get("title", ""),
                    "display": (
                        f"[{s.get('num', '')}] {s.get('title', '')}"
                        if s.get("num")
                        else s.get("title", "")
                    ),
                    "depth": lvl,
                    "order": idx,
                    "in_toc": True,  # 预解析数据默认按标题类处理
                }
                if lvl == 1:
                    word_hierarchy["level1"].append(item)
                elif lvl == 2:
                    word_hierarchy["level2"].append(item)
                else:
                    word_hierarchy["level3"].append(item)
        else:
            word_hierarchy = self.extract_word_content(
                word_file, mode="hierarchical", progress_callback=progress_callback
            )

        if progress_callback:
            progress_callback(25, "正在读取 Excel 数据...")

        # 2. 统一提取 Excel 内容
        excel_data_raw = self.extract_excel_content(
            excel_file,
            mode="hierarchical",
            sheet_name=sheet_name,
            header=header,
            level1_col=level1_col,
            level2_col=level2_col,
            level3_col=level3_col,
        )

        # --- 频次统计与去重逻辑 ---
        # 1. 统计各层级（L1, L2, L3）在整个 Excel 中出现的总行数 (用于判断是否显示行号)
        l1_counts = {}
        l2_counts = {}
        l3_counts = {}
        for row in excel_data_raw:
            v1, v2, v3 = row["level1"], row["level2"], row["level3"]
            if v1:
                l1_counts[v1] = l1_counts.get(v1, 0) + 1
            if v2:
                l2_counts[v2] = l2_counts.get(v2, 0) + 1
            if v3:
                l3_counts[v3] = l3_counts.get(v3, 0) + 1

        # 2. 对原始 Excel 数据进行去重合并，同一个 (L1, L2, L3) 组合合并为一项，并汇总所有行号范围
        excel_data = []
        seen_modules = {}  # key -> index in excel_data

        for row in excel_data_raw:
            key = (row["level1"], row["level2"], row["level3"])
            if key not in seen_modules:
                seen_modules[key] = len(excel_data)
                # 初始化合并后的记录
                merged_row = row.copy()
                merged_row["all_ranges"] = [row["row_range"]]
                # 记录该模块各级文本在全表的频次，供后续 row_info 使用
                merged_row["l1_total_count"] = l1_counts.get(row["level1"], 0)
                merged_row["l2_total_count"] = l2_counts.get(row["level2"], 0)
                merged_row["l3_total_count"] = l3_counts.get(row["level3"], 0)
                excel_data.append(merged_row)
            else:
                idx = seen_modules[key]
                rng = row["row_range"]
                if rng not in excel_data[idx]["all_ranges"]:
                    excel_data[idx]["all_ranges"].append(rng)

        # 格式化最终的汇总行号字符串
        for row in excel_data:
            unique_ranges = []
            for r in row["all_ranges"]:
                if r not in unique_ranges:
                    unique_ranges.append(r)
            row["final_row_info"] = ", ".join(unique_ranges)

        if not excel_data:
            print("错误：Excel 中没有找到数据")
            return {
                "exact_matched": [],
                "fuzzy_matched": [],
                "not_found_in_word": [],
                "statistics": {
                    "Excel功能点总数": 0,
                    "缺失项": 0,
                    "匹配率": "0%",
                    "错误": "Excel 无有效数据",
                },
            }

        # 统计各层级数量
        total_l1 = len(word_hierarchy["level1"])
        total_l2 = len(word_hierarchy["level2"])
        total_l3 = len(word_hierarchy["level3"])

        # 合并所有层级并按原始顺序排序
        all_items = []
        for item in word_hierarchy["level1"]:
            all_items.append(item)
        for item in word_hierarchy["level2"]:
            all_items.append(item)
        for item in word_hierarchy["level3"]:
            all_items.append(item)

        # 按原始文档顺序排序（使用order字段）
        all_items.sort(key=lambda x: x.get("order", 0))

        # 保存 Word 结构树到外部文件（避免 GUI 卡死）
        tree_lines = []
        tree_lines.append(f"{'='*80}")
        tree_lines.append(f"Word文档层级结构树")
        tree_lines.append(f"源文件: {word_file}")
        tree_lines.append(f"提取时间: {pd.Timestamp.now()}")
        tree_lines.append(f"{'='*80}\n")

        for item in all_items:
            depth = item["depth"]
            number = item["number"]
            text = item["text"]
            indent = "  " * (depth - 1)
            tree_lines.append(f"{indent}[{number}] {text}")

        if hierarchy_log_path:
            try:
                with open(hierarchy_log_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(tree_lines))
                print(f"  [OK] Word 结构树已保存至: {hierarchy_log_path}")
            except Exception as e:
                print(f"  [!] 无法保存 Word 结构树文件: {e}")

        print(f"\n{'='*80}")
        print(
            f"层级统计: Level 1={total_l1}个 (Word L1), Level 2={total_l2}个, Level 3+={total_l3}个"
        )
        print(f"{'='*80}\n")

        print(f"正在执行层级匹配...")
        print(f"Excel 功能点 (仅包含三级模块): {len(excel_data)} 个")

        # 显示Excel数据示例
        if excel_data:
            print(f"\nExcel数据示例（前3行）:")
            for i, row in enumerate(excel_data[:3]):
                print(
                    f"  [{i+1}] L1: '{row['level1']}' | L2: '{row['level2']}' | L3: '{row['level3']}'"
                )

        print(f"\nWord Level 1 标题: {total_l1} 个")
        print(f"Word Level 2 标题: {total_l2} 个")
        print(f"Word Level 3+ 标题: {total_l3} 个")

        if progress_callback:
            progress_callback(30, f"开始匹配 ({len(excel_data)} 项)...")

        exact_matched = []
        fuzzy_matched = []
        hierarchy_mismatched = []  # 新增：记录匹配但层级错位的项
        not_found_in_word = []

        import concurrent.futures

        # 预加载搜索空间以减少线程内过滤开销
        l1_space = (
            word_hierarchy["level1"]
            + word_hierarchy["level2"]
            + word_hierarchy["level3"]
        )
        l2_space = (
            word_hierarchy["level2"]
            + word_hierarchy["level1"]
            + word_hierarchy["level3"]
        )
        l3_space = (
            word_hierarchy["level3"]
            + word_hierarchy["level2"]
            + word_hierarchy["level1"]
        )

        total = len(excel_data)
        u_interval = max(1, total // 20)

        for idx, excel_row in enumerate(excel_data):
            if progress_callback and (idx % u_interval == 0 or idx == total - 1):
                progress = 30 + int((idx / (total if total > 0 else 1)) * 60)
                progress_callback(progress, f"层级匹配中... {idx + 1}/{total}")

            # --- 内部辅助函数保持逻辑一致 ---
            def is_descendant(parent_num, child_num):
                if not parent_num or not child_num:
                    return True
                if parent_num == child_num:
                    return True
                return child_num.startswith(parent_num + ".")

            def get_formatted_row_info(total_count, range_str):
                if not range_str:
                    return ""
                return f"（{range_str}行）"

            # 1. 初始化状态
            l1_matched_flag = not bool(excel_row["level1"])
            l2_matched_flag = not bool(excel_row["level2"])
            l3_matched_flag = not bool(excel_row["level3"])

            l1_score = 1.0 if l1_matched_flag else 0.0
            l2_score = 1.0 if l2_matched_flag else 0.0
            l1_matched_depth = 2 if not excel_row["level1"] else 0
            l2_matched_depth = 3 if not excel_row["level2"] else 0
            l3_matched_depth = 4 if not excel_row["level3"] else 0

            word_l1_num, word_l1_title = "", ""
            word_l2_num, word_l2_title = "", ""
            word_l3_num, word_l3_title = "", ""
            l3_best_match_item = None
            l3_best_score = 0.0

            # 2. 逐级比对 (L1 -> L2 -> L3)
            if excel_row["level1"]:
                found, match, score = self.find_match(excel_row["level1"], l1_space)
                if found and score >= self.threshold:
                    l1_matched_flag, l1_score = True, score
                    word_l1_title, word_l1_num, l1_matched_depth = (
                        match.get("display", "❌ 缺失"),
                        match.get("number", ""),
                        match.get("depth", 2),
                    )
                else:
                    word_l1_title = "❌ 缺失"

            if excel_row["level2"]:
                found, match, score = self.find_match(excel_row["level2"], l2_space)
                if found and score >= self.threshold:
                    l2_matched_flag, l2_score = True, score
                    word_l2_title, word_l2_num, l2_matched_depth = (
                        match.get("display", "❌ 缺失"),
                        match.get("number", ""),
                        match.get("depth", 3),
                    )
                else:
                    word_l2_title = "❌ 缺失"

            if excel_row["level3"]:
                found, match, score = self.find_match(excel_row["level3"], l3_space)
                if found and score >= self.threshold:
                    l3_matched_flag, l3_best_match_item, l3_best_score = (
                        True,
                        match,
                        score,
                    )
                    word_l3_title, word_l3_num, l3_matched_depth = (
                        match.get("display", "❌ 缺失"),
                        match.get("number", ""),
                        match.get("depth", 4),
                    )
            else:
                word_l3_title = ""

            l3_matched_flag = l3_best_match_item is not None

            # 3. 各种规则判定 (层级递增、关系闭环)
            n1, n2, n3 = (
                (word_l1_num if l1_matched_flag else None),
                (word_l2_num if l2_matched_flag else None),
                (word_l3_num if l3_matched_flag else None),
            )
            l1_l2_ok = is_descendant(n1, n2) if (n1 and n2) else True
            l2_l3_ok = is_descendant(n2, n3) if (n2 and n3) else True
            l1_l3_ok = is_descendant(n1, n3) if (n1 and n3) else True
            path_broken = (
                (
                    l1_matched_flag
                    and l2_matched_flag
                    and l3_matched_flag
                    and (not l1_l2_ok or not l2_l3_ok)
                )
                or (l2_matched_flag and l3_matched_flag and not l2_l3_ok)
                or (l1_matched_flag and l3_matched_flag and not l1_l3_ok)
            )

            # 深度校验
            depth_errs = []
            depth_seq = []
            if excel_row["level1"] and l1_matched_flag:
                depth_seq.append(("一级", l1_matched_depth))
            if excel_row["level2"] and l2_matched_flag:
                depth_seq.append(("二级", l2_matched_depth))
            if excel_row["level3"] and l3_matched_flag:
                depth_seq.append(("三级", l3_matched_depth))
            if len(depth_seq) >= 2:
                for i in range(len(depth_seq) - 1):
                    if depth_seq[i + 1][1] <= depth_seq[i][1]:
                        depth_errs.append(f"{depth_seq[i+1][0]}深度不足")
                    elif depth_seq[i + 1][1] - depth_seq[i][1] > 2:
                        depth_errs.append("层级跳跃过大")
            if excel_row["level1"] and l1_matched_flag and l1_matched_depth > 3:
                depth_errs.append("一级匹配过深")

            # 4. 构建描述与状态
            l1_missing = bool(excel_row["level1"]) and not l1_matched_flag
            l2_missing = bool(excel_row["level2"]) and not l2_matched_flag
            l3_missing = bool(excel_row["level3"]) and not l3_matched_flag
            has_any = (
                (l1_matched_flag and excel_row["level1"])
                or (l2_matched_flag and excel_row["level2"])
                or (l3_matched_flag and excel_row["level3"])
            )

            missing_desc = "-"
            if l1_missing or l2_missing or l3_missing:
                m_parts = []
                if l1_missing:
                    m_parts.append(f"一级模块 [{excel_row['level1_original']}]")
                if l2_missing:
                    m_parts.append(f"二级模块 [{excel_row['level2_original']}]")
                if l3_missing:
                    m_parts.append(f"三级模块 [{excel_row['level3_original']}]")

                count_missing = len(m_parts)
                if count_missing == 1:
                    prefix = (
                        "拆分表一级模块在需求规格书未体现："
                        if l1_missing
                        else (
                            "拆分表二级模块在需求规格书未体现："
                            if l2_missing
                            else "拆分表三级模块在需求规格书未体现："
                        )
                    )
                    missing_desc = f"{prefix}拆分表{m_parts[0]} 在需求规格书未体现"
                else:
                    level_names = "".join(
                        [
                            "一" if l1_missing else "",
                            "二" if l2_missing else "",
                            "三" if l3_missing else "",
                        ]
                    )
                    prefix = f"拆分表{level_names}级模块在需求规格书未体现："
                    missing_desc = (
                        f"{prefix}拆分表{'、'.join(m_parts)} 在需求规格书未体现"
                    )

            mismatch_desc = "-"
            # 改进判定：只要 matched 且 (有路径错误 或 有深度错误 或 有层级丢失)
            if has_any and (
                depth_errs or path_broken or l1_missing or l2_missing or l3_missing
            ):
                # 只有当至少有两个层级被匹配到时，才进行路径不匹配描述（排除掉缺失的部分）
                matched_levels = []
                if excel_row["level1"] and l1_matched_flag:
                    matched_levels.append(1)
                if excel_row["level2"] and l2_matched_flag:
                    matched_levels.append(2)
                if excel_row["level3"] and l3_matched_flag:
                    matched_levels.append(3)

                if (
                    path_broken
                    or depth_errs
                    or any([l1_missing, l2_missing, l3_missing])
                ) and len(matched_levels) >= 2:
                    e_segments = []
                    w_segments = []
                    m_names = []
                    for lvl in matched_levels:
                        if lvl == 1:
                            e_segments.append(
                                f"一级模块{{{excel_row['level1_original']}}}"
                            )
                            w_segments.append(f"一级目录{{{word_l1_title}}}")
                            m_names.append("一")
                        elif lvl == 2:
                            e_segments.append(
                                f"二级模块{{{excel_row['level2_original']}}}"
                            )
                            w_segments.append(f"二级目录{{{word_l2_title}}}")
                            m_names.append("二")
                        elif lvl == 3:
                            e_segments.append(
                                f"三级模块{{{excel_row['level3_original']}}}"
                            )
                            w_segments.append(f"三级目录{{{word_l3_title}}}")
                            m_names.append("三")

                    m_str = "".join(m_names)
                    # 严格按照用户要求的格式：拆分表{m_str}级模块与需求规格书不匹配：拆分表 {' - '.join(e_segments)} 与需求规格书 {' - '.join(w_segments)}不匹配
                    mismatch_desc = f"拆分表{m_str}级模块与需求规格书不匹配：拆分表 {' - '.join(e_segments)} 与需求规格书 {' - '.join(w_segments)}不匹配"

                    if depth_errs:
                        mismatch_desc += f" (原因: {', '.join(depth_errs)})"
                elif depth_errs and not mismatch_desc.startswith("拆分表"):
                    mismatch_desc = f"层级深度不匹配：{', '.join(depth_errs)}"

            # 最终整合：优先展示缺失，其次展示层级不匹配
            final_brief = (
                missing_desc
                if missing_desc != "-"
                else (mismatch_desc if mismatch_desc != "-" else "-")
            )

            # 最终整合
            word_full_num = (
                f"{word_l1_num or '-'}/{word_l2_num or '-'}/{word_l3_num or '-'}"
            )
            sim_score = (
                l3_best_score
                if l3_matched_flag
                else (l2_score if l2_matched_flag else l1_score)
            )
            match_status = (
                "缺失"
                if not has_any
                else (
                    "层级不匹配"
                    if (
                        l1_missing
                        or l2_missing
                        or l3_missing
                        or path_broken
                        or depth_errs
                    )
                    else ("精确匹配" if sim_score >= 0.99 else "模糊匹配")
                )
            )

            data = {
                "Excel一级模块": excel_row["level1_original"],
                "Word一级标题": word_l1_title,
                "Excel二级模块": excel_row["level2_original"],
                "Word二级标题": word_l2_title,
                "Excel三级模块": excel_row["level3_original"],
                "Word三级标题": word_l3_title,
                "Word完整编号": word_full_num,
                "相似度": f"{sim_score:.2%}" if sim_score > 0 else "",
                "匹配状态": match_status,
                "缺失简略描述": missing_desc,
                "层级不匹配简略描述": mismatch_desc,
                "简略描述": final_brief,
                "row_range": excel_row.get("final_row_info", ""),
            }

            if match_status == "精确匹配":
                exact_matched.append(data)
            elif match_status == "模糊匹配":
                fuzzy_matched.append(data)
            elif match_status == "层级不匹配":
                hierarchy_mismatched.append(data)
            else:
                not_found_in_word.append(data)

        if progress_callback:
            progress_callback(90, "生成统计报告...")

        total_excel = len(excel_data)
        exact_count = len(exact_matched)
        fuzzy_count = len(fuzzy_matched)
        mismatch_count = len(hierarchy_mismatched)

        # 核心判定：只有精确匹配和模糊匹配才算“成功已匹配”
        # 层级不匹配虽然在 Word 中找到了文本，但因为结构问题不应作为“通过”
        pass_matched_count = exact_count + fuzzy_count
        all_found_count = exact_count + fuzzy_count + mismatch_count
        missing_count = len(not_found_in_word)

        # 匹配率展示通过率
        match_rate = (pass_matched_count / total_excel * 100) if total_excel > 0 else 0

        print(
            f"\n匹配完成：精确 {exact_count} 项，模糊 {fuzzy_count} 项，层级不匹配 {mismatch_count} 项，缺失 {missing_count} 项"
        )
        print(f"  最终通过率 (精确+模糊): {match_rate:.2f}%")

        if progress_callback:
            progress_callback(100, "匹配完成！")

        report = {
            "exact_matched": exact_matched,
            "fuzzy_matched": fuzzy_matched,
            "hierarchy_mismatched": hierarchy_mismatched,
            "not_found_in_word": not_found_in_word,
            "failed_matches": not_found_in_word
            + hierarchy_mismatched,  # 合并缺失与层级不匹配
            "statistics": {
                "Excel功能点总数": total_excel,
                "Word内容项总数": len(word_hierarchy["level1"])
                + len(word_hierarchy["level2"])
                + len(word_hierarchy["level3"]),
                "精确匹配": exact_count,
                "模糊匹配": fuzzy_count,
                "层级不匹配": mismatch_count,
                "已匹配": pass_matched_count,
                "总计找到": all_found_count,
                "缺失项": missing_count,
                "匹配率": f"{match_rate:.2f}%",
                "完成度": f"{match_rate:.2f}%",
            },
        }

        return report

    def save_report(self, report: Dict, output_file: str = "匹配报告.xlsx"):
        """
        保存报告到 Excel（如果文件被占用，自动生成副本）

        Args:
            report: 匹配报告
            output_file: 输出文件路径
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

                    # 定义层级匹配报告的列顺序
                    hierarchical_column_order = [
                        "Excel一级模块",
                        "Word一级标题",
                        "Excel二级模块",
                        "Word二级标题",
                        "Excel三级模块",
                        "Word三级标题",
                        "Word完整编号",
                        "相似度",
                        "匹配层级",
                        "匹配状态",
                        "缺失简略描述",
                        "层级不匹配简略描述",
                        "简略描述",
                    ]

                    # 定义简单匹配报告的列顺序
                    simple_column_order = [
                        "Excel功能点",
                        "Word匹配项",
                        "位置",
                        "相似度",
                        "匹配状态",
                        "简略描述",
                    ]

                    # 检查报告类型以确定列顺序
                    is_hierarchical_report = (
                        bool(
                            report["exact_matched"]
                            and "Excel一级模块" in report["exact_matched"][0]
                        )
                        or bool(
                            report["fuzzy_matched"]
                            and "Excel一级模块" in report["fuzzy_matched"][0]
                        )
                        or bool(
                            report.get("hierarchy_mismatched")
                            and "Excel一级模块" in report["hierarchy_mismatched"][0]
                        )
                        or bool(
                            report["not_found_in_word"]
                            and "Excel一级模块" in report["not_found_in_word"][0]
                        )
                    )

                    current_column_order = (
                        hierarchical_column_order
                        if is_hierarchical_report
                        else simple_column_order
                    )

                    # 准备合并的数据列表
                    all_data = []

                    # 精确匹配项
                    if report["exact_matched"]:
                        for item in report["exact_matched"]:
                            item["匹配状态"] = "精确匹配"
                            all_data.append(item)
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
                    if report["fuzzy_matched"]:
                        for item in report["fuzzy_matched"]:
                            item["匹配状态"] = "模糊匹配"
                            all_data.append(item)
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

                    # 合并层级不匹配与缺失项
                    failed_hierarchical_data = []
                    if report.get("hierarchy_mismatched"):
                        for item in report["hierarchy_mismatched"]:
                            if "匹配状态" not in item:
                                item["匹配状态"] = "层级不匹配"
                            failed_hierarchical_data.append(item)

                    if report.get("not_found_in_word") and is_hierarchical_report:
                        for item in report["not_found_in_word"]:
                            if "匹配状态" not in item:
                                item["匹配状态"] = "缺失"
                            failed_hierarchical_data.append(item)

                    if failed_hierarchical_data:
                        failed_df = pd.DataFrame(failed_hierarchical_data)
                        failed_df = failed_df[
                            [
                                col
                                for col in current_column_order
                                if col in failed_df.columns
                            ]
                        ]
                        failed_df.to_excel(
                            writer, sheet_name="❌ 缺失及层级不匹配", index=False
                        )
                    elif is_hierarchical_report:
                        empty_df = pd.DataFrame(
                            [{"说明": "🎉 恭喜！未发现缺失或层级不匹配项！"}]
                        )
                        empty_df.to_excel(
                            writer, sheet_name="❌ 缺失及层级不匹配", index=False
                        )

                    # 简单模式下的缺失处理（非层级报告）
                    if report["not_found_in_word"] and not is_hierarchical_report:
                        for item in report["not_found_in_word"]:
                            if "匹配状态" not in item:
                                item["匹配状态"] = "缺失"
                            all_data.append(item)

                        missing_df = pd.DataFrame(report["not_found_in_word"])
                        missing_df = missing_df[
                            [
                                col
                                for col in current_column_order
                                if col in missing_df.columns
                            ]
                        ]
                        missing_df.to_excel(writer, sheet_name="❌ 缺失", index=False)
                    elif not is_hierarchical_report:
                        empty_df = pd.DataFrame([{"说明": "🎉 恭喜！未发现缺失模块！"}])
                        empty_df.to_excel(writer, sheet_name="❌ 缺失", index=False)

                    # 新增：总表
                    if all_data:
                        all_df = pd.DataFrame(all_data)
                        # 确保列顺序
                        cols = [
                            col for col in current_column_order if col in all_df.columns
                        ]
                        all_df = all_df[cols]
                        all_df.to_excel(
                            writer, sheet_name="📋 汇总匹配结果", index=False
                        )

                print(f"✓ 报告已保存: {output_file}")
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

                # 生成新文件名：原文件名_副本1. xlsx, 原文件名_副本2.xlsx, ...
                output_file = str(parent / f"{stem}_副本{attempt}{suffix}")

                print(f"  ⚠️ 原文件被占用，尝试保存到: {output_file}")

            except Exception as e:
                print(f"✗ 保存报告失败: {str(e)}")
                import traceback

                traceback.print_exc()
                raise
