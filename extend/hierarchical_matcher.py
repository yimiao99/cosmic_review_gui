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

    def __init__(self, fuzzy_match: bool = False, threshold: float = 0.8):
        """
        初始化匹配器

        Args:
            fuzzy_match: 是否启用模糊匹配
            threshold: 模糊匹配的相似度阈值（0-1）
        """
        self.fuzzy_match = fuzzy_match
        self.threshold = threshold

    def _open_word_doc(self, word_file: str):
        """
        统一打开 Word 文档的方法，处理 .doc 转换、兼容性修复。
        返回 docx.Document 对象。
        """
        original_word_file = Path(word_file)
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
                    if temp_compat_path and os.exists(temp_compat_path):
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
                return self._extract_word_hierarchical(doc, progress_callback)
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

        style_name = paragraph.style.name.lower()
        level_num, num_depth = self.extract_level_number(text)

        # 新增：尝试获取段落的大纲级别 (Outline Level)
        # Word 内部 0=Level 1, 1=Level 2, ..., 8=Level 9, 9=Body Text
        try:
            outline_level = paragraph.paragraph_format.outline_level
        except:
            outline_level = 9  # 默认为正文

        is_outline_title = outline_level < 9

        # NEW: 尝试获取列表缩进级别 (List Indentation Level)
        # 解决 .docx 自动编号读不到文本的问题
        list_level = self._get_list_level(paragraph)

        return {
            "original": text,
            "style": style_name,
            "level_number": level_num,
            "num_depth": num_depth,
            "outline_level": outline_level,
            "list_level": list_level,  # 新增字段: XML列表层级 (0, 1, 2...)
            "is_title": "heading" in style_name
            or style_name.startswith("toc ")
            or "标题" in style_name
            or is_outline_title
            or list_level is not None,  # 如果是列表项，也视为潜在的结构项
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
        total = len(doc.paragraphs)

        # 先提取所有标题，用于后续判断是否在目录中
        toc_texts = set()
        for para in doc.paragraphs:
            info = self._get_paragraph_info(para)
            if info and info["is_title"]:
                cleaned = self.basic_clean(info["original"])
                if cleaned:
                    toc_texts.add(cleaned)

        for i, para in enumerate(doc.paragraphs):
            if progress_callback and i % 50 == 0:
                progress_callback(
                    10 + int(i / total * 80), f"正在读取段落 {i}/{total}..."
                )

            info = self._get_paragraph_info(para)
            if not info:
                continue

            # 如果是标题，或者启用了全文搜索
            if info["is_title"] or full_text_search:
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

                    is_in_toc = cleaned in toc_texts
                    items.append(
                        {
                            "original": text,
                            "display": (
                                cleaned[:100] + "..." if len(cleaned) > 100 else cleaned
                            ),
                            "cleaned": cleaned,
                            "text": cleaned,
                            "level": info["style"],
                            "in_toc": is_in_toc,
                        }
                    )
        return items

    def _extract_word_hierarchical(self, doc, progress_callback=None) -> Dict:
        """
        混合策略提取：
        1. 显式编号策略：尝试在“功能需求”章节寻找文本中包含 "4.1" 等显式编号的项。
        2. 大纲结构策略：扫描全文大纲级别，构建层级树，并自动生成编号。

        决策：
        - 如果显式编号策略找到足够多的项，优先使用它（最准确）。
        - 否则，使用大纲结构策略，并尝试定位目标章节，将其编号前缀修正为 "4"。
        """

        # --- 策略 1 数据结构: 显式编号 ---
        explicit_hierarchy = {"level1": [], "level2": [], "level3": []}
        explicit_in_section = False

        # --- 策略 2 数据结构: 大纲结构 ---
        chapters = []
        current_chapter = None
        outline_counts = [0, 0, 0, 0, 0]
        last_heading_depth = 0  # 追踪上一个标题的深度，用于推断列表项的深度

        total = len(doc.paragraphs)
        for i, para in enumerate(doc.paragraphs):
            if progress_callback and i % 100 == 0:
                progress_callback(
                    10 + int(i / total * 80), f"正在解析文档结构 {i}/{total}..."
                )

            info = self._get_paragraph_info(para)
            if not info:
                continue

            cleaned = self.basic_clean(info["original"])

            # ====== 策略 1: 显式编号逻辑 ======
            # 检查章节进入
            is_target_section = (
                (self.SECTION_FUNCTIONAL_REQUIREMENTS in cleaned)
                or ("功能" in cleaned and "需求" in cleaned)
                or ("业务" in cleaned and "功能" in cleaned)
            )
            # 检查是否是标题 (Heading 1 或 Depth 1)
            is_h1 = (
                "heading 1" in info["style"]
                or "标题 1" in info["style"]
                or info["outline_level"] == 0
            )

            if is_target_section and (info["num_depth"] == 1 or is_h1):
                explicit_in_section = True
            elif (
                explicit_in_section
                and (info["num_depth"] == 1 or is_h1)
                and not is_target_section
            ):
                explicit_in_section = False

            # 收集显式编号项
            if explicit_in_section and info["level_number"]:
                # 必须以 "4." 开头 (或者配置的前缀)
                if info["level_number"].startswith(self.CHAPTER_NUMBER_PREFIX + "."):
                    item = {
                        "number": info["level_number"],
                        "original": info["original"],
                        "display": cleaned,
                        "cleaned": cleaned,
                        "text": cleaned,
                        "depth": info["num_depth"],
                    }
                    mapping = {2: "level1", 3: "level2", 4: "level3"}
                    if info["num_depth"] in mapping:
                        explicit_hierarchy[mapping[info["num_depth"]]].append(item)
                        print(f"  [DEBUG-显式] 添加到 {mapping[info['num_depth']]}: {item['number']} {item['display'][:20]}")
                    else:
                        print(f"  [DEBUG-显式] 忽略深度 {info['num_depth']}: {info['level_number']} {info['display'][:20]}")

            # ====== 策略 2: 大纲结构逻辑 ======
            # 仅处理有大纲级别的段落
            if not info["is_title"]:
                continue

            # 计算深度
            depth = 0
            if info["outline_level"] < 9:
                depth = info["outline_level"] + 1
            else:
                style = info["style"]
                if "heading 1" in style or "标题 1" in style:
                    depth = 1
                elif "heading 2" in style or "标题 2" in style:
                    depth = 2
                elif "heading 3" in style or "标题 3" in style:
                    depth = 3
                elif "heading 4" in style or "标题 4" in style:
                    depth = 4
                elif "heading 5" in style or "标题 5" in style:
                    depth = 5  # 新增支持 H5
            
            # 记录标题深度
            if depth > 0:
                last_heading_depth = depth

            # 如果没有从样式/大纲获取到深度，尝试从列表级别推断
            # 使用上下文感知逻辑：基于上一个标题的深度 + 列表层级 + 1
            if depth == 0 and info["list_level"] is not None:
                if last_heading_depth > 0:
                    # 逻辑: H3 (Depth 3) 下面的 List Level 0 应该是 Depth 4 (Level 3 Item)
                    # H2 (Depth 2) 下面的 List Level 0 应该是 Depth 3 (Level 2 Item)
                    calculated_depth = last_heading_depth + 1 + info["list_level"]
                    depth = min(calculated_depth, 5) # 限制最大深度
                    print(f"  [DEBUG-列表] 上级Depth {last_heading_depth} + ListLvl {info['list_level']} -> Depth {depth}: {cleaned[:10]}...")
            
            if depth == 0:
                continue

            # 更新计数器
            if depth <= len(outline_counts):
                outline_counts[depth - 1] += 1
                for j in range(depth, len(outline_counts)):
                    outline_counts[j] = 0

            # 生成编号 (优先用提取的，没有则用计数器)
            # 注意: 对于自动编号列表，info["level_number"] 通常为空，所以会进入 else 分支使用计数器
            if info["level_number"]:
                final_num = info["level_number"]
            else:
                parts = [str(c) for c in outline_counts[:depth] if c > 0]
                final_num = ".".join(parts)

            item_outline = {
                "number": final_num,
                "original": info["original"],
                "display": cleaned,
                "cleaned": cleaned,
                "text": cleaned,
                "depth": depth,
                "outline_counts": list(
                    outline_counts[:depth]
                ),  # 保存当时的计数状态，方便后续修正前缀
            }

            # 归类到章节
            if depth == 1:
                current_chapter = {
                    "title": cleaned,
                    "original": info["original"],
                    "items": {"level1": [], "level2": [], "level3": []},
                }
                chapters.append(current_chapter)
            elif depth in [2, 3, 4, 5]:  # 允许处理 depth 5
                if current_chapter is None:
                    current_chapter = {
                        "title": "文档开头",
                        "original": "Start",
                        "items": {"level1": [], "level2": [], "level3": []},
                    }
                    chapters.append(current_chapter)

                # 动态映射：
                # 如果文档整体下沉 (H2-H3-H4-H5)，则 H3->L1, H4->L2, H5->L3
                # 但为了兼容标准文档 (H1-H2-H3-H4)，我们采用宽容策略：
                # H2 -> level1
                # H3 -> level2
                # H4 -> level3 (标准) 或 level2 (下沉)
                # H5 -> level3 (下沉)

                target_list = None
                if depth == 2:
                    target_list = "level1"
                elif depth == 3:
                    target_list = "level2"
                elif depth == 4:
                    target_list = "level3"  # 标准 H4 是 L3
                elif depth == 5:
                    target_list = "level3"  # 下沉 H5 也是 L3

                if target_list:
                    current_chapter["items"][target_list].append(item_outline)
                    print(f"  [DEBUG-大纲] Depth {depth} -> {target_list}: {item_outline['number']} {item_outline['display'][:20]}")

                    # 特殊处理：如果是 H4，它也可能是下沉文档的 L2，所以也存一份到 level2 (标记一下)
                    if depth == 4:
                        item_copy = item_outline.copy()
                        current_chapter["items"]["level2"].append(item_copy)
                else:
                    print(f"  [DEBUG-大纲] Depth {depth} 未映射到任何 list: {item_outline['number']} {cleaned[:20]}")

        # ====== 决策阶段 ======
        explicit_count = (
            len(explicit_hierarchy["level1"])
            + len(explicit_hierarchy["level2"])
            + len(explicit_hierarchy["level3"])
        )
        
        # 统计大纲策略找到的项目数 (仅仅为了比较)
        outline_total_l3 = 0
        for ch in chapters:
            outline_total_l3 += len(ch["items"]["level3"])

        print(f"\n[DEBUG] 策略1(显式编号)总数: {explicit_count} (L3: {len(explicit_hierarchy['level3'])})")
        print(f"[DEBUG] 策略2(大纲结构)潜在L3总数: {outline_total_l3}")

        # 逻辑修正: 
        # 如果显式编号找到了一些 (>=5)，通常我们会采纳。
        # 但是，如果显式编号完全没找到 L3 (L3=0)，而大纲策略找到了 L3 (>0)，
        # 这说明显式编号可能失效了 (例如 docx 自动编号)，此时应强制切换到大纲策略。
        
        use_explicit = False
        if explicit_count >= 5:
            use_explicit = True
            if len(explicit_hierarchy["level3"]) == 0 and outline_total_l3 > 0:
                print("  ⚠️ 虽然显式编号策略找到了一些项，但缺少L3，而大纲策略找到了L3。")
                print("  -> 切换到大纲策略 (Strategy 2)")
                use_explicit = False
        
        if use_explicit:
            print("  ✓ 采用显式编号策略结果")
            return explicit_hierarchy

        print("  -> 采用策略2(大纲结构)...")

        # 在大纲章节中寻找目标
        target_chapter = None
        keywords = ["功能需求", "业务需求", "系统功能", "功能列表"]

        # 1. 标题匹配
        for chapter in chapters:
            if any(kw in chapter["title"] for kw in keywords):
                target_chapter = chapter
                print(f"  ✓ [大纲策略] 定位到目标章节: {chapter['title']}")
                break

        # 2. 内容最多匹配
        if not target_chapter and chapters:
            best_chap = max(
                chapters,
                key=lambda c: len(c["items"]["level1"])
                + len(c["items"]["level2"])
                + len(c["items"]["level3"]),
            )
            if len(best_chap["items"]["level1"]) > 0:
                target_chapter = best_chap
                print(f"  ⚠️ [大纲策略] 自动选择内容最多的章节: {best_chap['title']}")

        final_hierarchy = {"level1": [], "level2": [], "level3": []}

        if target_chapter:
            # 修正编号前缀！
            # 如果目标章节的计数是 "1" (因为它是文档第一个H1)，但我们需要 "4"
            # 我们遍历该章节所有条目，把编号的第一位修正为 CHAPTER_NUMBER_PREFIX

            prefix_to_use = self.CHAPTER_NUMBER_PREFIX  # "4"

            for lvl in ["level1", "level2", "level3"]:
                for item in target_chapter["items"][lvl]:
                    # item['number'] 可能是 "1.1", "1.1.1"
                    # 我们想把它变成 "4.1", "4.1.1"

                    # 只有当它是自动生成的编号(或者不符合4开头)时才修正
                    if not item["number"].startswith(prefix_to_use + "."):
                        parts = item["number"].split(".")
                        if parts and parts[0].isdigit():
                            parts[0] = prefix_to_use
                            item["number"] = ".".join(parts)

                    final_hierarchy[lvl].append(item)
        else:
            # 合并所有
            print("  ⚠️ [大纲策略] 未识别到章节，合并全文。")
            for chapter in chapters:
                for lvl in ["level1", "level2", "level3"]:
                    final_hierarchy[lvl].extend(chapter["items"][lvl])

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
        df_ffill, df_raw = self._load_excel_with_merged(excel_file, sheet_name, header=header_row)

        if mode == "flat":
            return self._extract_excel_flat(df_ffill, df_raw, kwargs.get("column", 0), header_row)
        else:
            return self._extract_excel_hierarchical(
                df_ffill,
                df_raw,
                kwargs.get("level1_col", 0),
                kwargs.get("level2_col", 1),
                kwargs.get("level3_col", 2),
                header_row
            )

    def _extract_excel_flat(self, df: pd.DataFrame, df_raw: pd.DataFrame, column: int, header: int) -> List[Dict]:
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

            # 常见标题过滤
            if val in ["功能过程", "功能用户", "触发事件", "子过程描述"]:
                continue

            # 检测合并范围
            end_idx = idx
            for j in range(idx + 1, len(df)):
                # 检查原始 df 中该列是否为空（即合并或待填充）
                # 且填充后的值是否一致
                raw_val_next = str(df_raw.iloc[j, column]).strip().lower()
                ffill_val_next = str(df.iloc[j, column]).strip()
                
                if (raw_val_next == "" or raw_val_next == "nan") and ffill_val_next == val:
                    end_idx = j
                else:
                    break
            
            skip_to = end_idx
            row_num_start = idx + header + 2
            row_num_end = end_idx + header + 2
            row_range = f"{row_num_start}-{row_num_end}" if row_num_end > row_num_start else f"{row_num_start}"

            data.append({
                "original": val, 
                "cleaned": val, 
                "text": val, 
                "row_range": row_range,
                "row_index": idx
            })
            
        return data

    def _extract_excel_hierarchical(
        self, df: pd.DataFrame, df_raw: pd.DataFrame, l1: int, l2: int, l3: int, header: int
    ) -> List[Dict]:
        """提取层级 Excel 数据，支持每一层具体合并范围的检测"""
        data = []
        
        # 预计算每一层级（L1, L2, L3）在每一行的具体范围
        # 即使 L3 是最细粒度的，我们也需要知道这一行所属的 L1 到底跨越了哪些行
        
        def get_range(row_idx, col_idx):
            val = str(df.iloc[row_idx, col_idx]).strip()
            if not val or val.lower() == "nan":
                return f"{row_idx + header + 2}"
            
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
            return f"{s_num}-{e_num}" if e_num > s_num else f"{s_num}"

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
                    "row_range": get_range(idx, l3), # 默认使用 L3 范围
                }
            )
        return data

    def _load_excel_with_merged(self, excel_file: str, sheet_name=0, header=0) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        加载 Excel 并处理合并单元格。
        返回 (填充后的 DataFrame, 原始 DataFrame)。
        """
        # 使用 pandas 读取
        df_raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=header)
        
        # 预先清理：将所有 NaN 转换为 Python 的 None 或空字符串以便统一判断
        df_ffill = df_raw.copy()

        # 针对常见的前几列（模块列）进行填充
        cols_to_fill = []
        for i in range(min(10, len(df_ffill.columns))): # 增加填充列数到 10，确保覆盖 level1-3
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
        patterns = [
            r"^(\d+\.\d+\.\d+\.\d+)",  # 4.1.1.1
            r"^(\d+\.\d+\.\d+)",  # 4.1.1
            r"^(\d+\.\d+)",  # 4.1
            r"^(\d+)(?=\.|\s|$)",  # NEW: 4. 或 4 (后面跟空格或字符串结束)，避免匹配“4A系统”
        ]

        for pattern in patterns:
            match = re.match(pattern, text_to_match)
            if match:
                level_num = match.group(1)
                depth = level_num.count(".") + 1
                # --- DEBUG START ---
                # print(f"  DEBUG_ELN: 成功匹配模式 '{pattern}', 编号: '{level_num}', 深度: {depth}")
                # --- DEBUG END ---
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
        cleaned_text = re.sub(r"[、，,\s]*(优化|新增|修改|删除)\s*$", "", cleaned_text)
        cleaned_text = re.sub(r"[。.]\s*$", "", cleaned_text)  # 移除末尾的句号

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

        # 6.  标准化空格
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
    ) -> Dict:
        """简单匹配模式（单列匹配）"""

        # 1. 统一提取 Word 内容
        # 注意：simple 模式下的 full_text_search 暂不通过统一 API 处理特殊过滤，直接使用其原有逻辑
        # 但我们为了对齐，可以使用统一 API 提取
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
        # 提高进度更新频率，每一项都更新（如果数据量极大则每10项更新一次）
        update_interval = 1 if total < 500 else 10

        for idx, excel_item in enumerate(excel_data):
            if progress_callback:
                if idx % update_interval == 0 or idx == total - 1:
                    progress = 30 + int((idx / (total if total > 0 else 1)) * 60)
                    progress_callback(progress, f"匹配中... {idx + 1}/{total}")

            # ✅ 关键修复：对 Excel 原始数据也做基础清洗，用于匹配
            excel_text_for_matching = self.basic_clean(excel_item["original"])

            # ====== 添加调试信息（仅前5条）======
            if idx < 5:
                print(f"\n【调试-匹配 {idx + 1}】")
                print(f"  Excel原始: '{excel_item['original']}'")
                print(f"  Excel清洗后用于匹配: '{excel_text_for_matching}'")
            # ====== 调试信息结束 ======

            found, word_match, similarity_score = self.find_match(
                excel_text_for_matching, word_items, 0
            )

            if found:
                is_in_toc = word_match.get("in_toc", True)
                location = "📑 目录" if is_in_toc else "📄 正文"

                match_data = {
                    "Excel功能点": excel_item["original"],  # ✅ 显示原始文本
                    "Word匹配项": word_match.get(
                        "display", word_match.get("original", "")
                    ),
                    "位置": location,
                    "相似度": f"{similarity_score:.2%}",  # ✅ 修复：去掉空格
                }

                # ====== 添加调试信息（仅前5条）======
                if idx < 5:
                    print(f"  ✅ 找到匹配!")
                    print(f"  Word display: '{word_match.get('display', '')}'")
                    print(f"相似度: {similarity_score:.2%}")  # ✅ 修复：去掉空格
                    print(
                        f"  【关键】match_data['Excel功能点']: '{match_data['Excel功能点']}'"
                    )
                    print(
                        f"  【关键】match_data['Word匹配项']: '{match_data['Word匹配项']}'"
                    )
                # ====== 调试信息结束 ======

                if similarity_score >= 0.99:
                    exact_matched.append(match_data)
                else:
                    fuzzy_matched.append(match_data)
            else:
                match_data = {
                    "Excel功能点": excel_item["original"], 
                    "状态": "❌ 缺失",
                    "简略描述": f"拆分表模块 {excel_item['original']} 行号 {excel_item.get('row_range', idx + 2)} 在需求规格书未体现"
                }
                not_found_in_word.append(match_data)

                # ====== 添加调试信息（仅前5条）======
                if idx < 5:
                    print(f"  ❌ 未找到匹配")
                    print(
                        f"  【关键】not_found Excel功能点: '{excel_item['original']}'"
                    )
                # ====== 调试信息结束 ======

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
                if item.get("位置") == "📑 目录"
            )
            body_count = sum(
                1
                for item in exact_matched + fuzzy_matched
                if item.get("位置") == "📄 正文"
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
        target_no_space = re.sub(r"\s+", "", target.lower())

        for candidate in candidates:
            candidate_text = candidate.get("text", candidate.get("cleaned", ""))
            candidate_no_space = re.sub(r"\s+", "", candidate_text.lower())

            current_score = 0.0

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
            if current_score < self.threshold and len(target_no_space) >= 3:
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
    ) -> Dict:
        """层级匹配模式 (增强版)"""
        # 1. 统一提取 Word 内容
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
            if v1: l1_counts[v1] = l1_counts.get(v1, 0) + 1
            if v2: l2_counts[v2] = l2_counts.get(v2, 0) + 1
            if v3: l3_counts[v3] = l3_counts.get(v3, 0) + 1

        # 2. 对原始 Excel 数据进行去重合并，同一个 (L1, L2, L3) 组合合并为一项，并汇总所有行号范围
        excel_data = []
        seen_modules = {} # key -> index in excel_data
        
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
                if r not in unique_ranges: unique_ranges.append(r)
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
                    "错误": "Excel 无有效数据"
                }
            }

        print(f"\n正在执行层级匹配...")
        print(f"Excel 功能点 (仅包含三级模块): {len(excel_data)} 个")
        print(f"Word 二级编号 (4.x): {len(word_hierarchy['level1'])} 个")
        print(f"Word 三级编号 (4.x.x): {len(word_hierarchy['level2'])} 个")
        print(f"Word 四级编号 (4.x.x.x): {len(word_hierarchy['level3'])} 个")

        if progress_callback:
            progress_callback(30, f"开始匹配 ({len(excel_data)} 项)...")

        exact_matched = []
        fuzzy_matched = []
        not_found_in_word = []

        total = len(excel_data)
        # 提高进度更新频率
        update_interval = 1 if total < 500 else 10

        for idx, excel_row in enumerate(excel_data):
            if progress_callback:
                if idx % update_interval == 0 or idx == total - 1:
                    progress = 30 + int((idx / (total if total > 0 else 1)) * 60)
                    progress_callback(progress, f"匹配中... {idx + 1}/{total}")

            # 初始化当前 Excel 行的匹配结果和状态
            # 如果 Excel 某级为空，则默认该级“已匹配”，以便不影响整体 all_levels_found 的判定
            l1_matched_flag = not bool(excel_row["level1"])
            l2_matched_flag = not bool(excel_row["level2"])
            l3_matched_flag = not bool(excel_row["level3"])

            l1_score = 1.0 if l1_matched_flag else 0.0
            l2_score = 1.0 if l2_matched_flag else 0.0

            # --- DEBUG START ---
            if idx < 3:  # 仅调试前3行
                print(f"\n[DEBUG] 正在匹配第 {idx+1} 行 Excel 数据:")
                print(f"  L1: '{excel_row['level1']}'")
                print(f"  L2: '{excel_row['level2']}'")
                print(f"  L3: '{excel_row['level3']}'")
            # --- DEBUG END ---

            # --- 尝试匹配 Excel L1 和 L2 ---
            word_l1_num = ""
            word_l2_num = ""

            # 1. 匹配 Excel 一级模块 (cleaned) 与 Word 二级编号 (depth 2, e.g., 4.1)
            if excel_row["level1"]:
                # 扩大搜索范围：不仅找 level1，也尝试找 level2 (防止 Word 层级错位)
                candidates_l1 = word_hierarchy["level1"] + word_hierarchy["level2"]
                found_l1, match_l1_item, score_l1 = self.find_match(
                    excel_row["level1"], candidates_l1
                )
                if found_l1 and score_l1 >= self.threshold:
                    l1_matched_flag = True
                    l1_score = score_l1
                    word_l1_title = match_l1_item.get("display", "❌ 缺失")
                    word_l1_num = match_l1_item.get("number", "")
                else:
                    word_l1_title = "❌ 缺失"
            else:
                word_l1_title = ""  # Excel 为空，显示为空白

            # --- DEBUG START ---
            if excel_row["level1"] and not l1_matched_flag and idx < 10:
                print(f"  [DEBUG-FAIL] L1 未找到匹配: '{excel_row['level1']}'")
            # --- DEBUG END ---

            # 2. 匹配 Excel 二级模块 (cleaned) 与 Word 三级编号 (depth 3, e.g., 4.1.1)
            if excel_row["level2"]:
                # 扩大搜索范围：level2 + level3 + level1
                candidates_l2 = (
                    word_hierarchy["level2"]
                    + word_hierarchy["level3"]
                    + word_hierarchy["level1"]
                )
                found_l2, match_l2_item, score_l2 = self.find_match(
                    excel_row["level2"], candidates_l2
                )
                if found_l2 and score_l2 >= self.threshold:
                    l2_matched_flag = True
                    l2_score = score_l2
                    word_l2_title = match_l2_item.get("display", "❌ 缺失")
                    word_l2_num = match_l2_item.get("number", "")
                else:
                    word_l2_title = "❌ 缺失"
            else:
                word_l2_title = ""  # Excel 为空，显示为空白

            # --- DEBUG START ---
            if excel_row["level2"] and not l2_matched_flag and idx < 10:
                print(f"  [DEBUG-FAIL] L2 未找到匹配: '{excel_row['level2']}'")
            # --- DEBUG END ---

            # --- 尝试匹配 Excel L3 (灵活匹配 Word L4, L3, L2) ---
            l3_best_match_item = None
            l3_best_score = 0.0
            l3_matched_word_depth = 0  # 记录Excel L3实际匹配到的Word标题的深度

            if excel_row["level3"]:
                # 尝试 1: 匹配 Excel L3 到 Word L4/L5 (深度 4/5)
                # word_hierarchy['level3'] 现在包含了 H4 和 H5
                found_l3_d4, match_l3_d4_item, score_l3_d4 = self.find_match(
                    excel_row["level3"],
                    word_hierarchy["level3"],  # 包含 Word 深度 4 和 5 的项目
                )
                if found_l3_d4 and score_l3_d4 >= self.threshold:
                    l3_best_match_item = match_l3_d4_item
                    l3_best_score = score_l3_d4
                    # 动态获取匹配项的深度
                    l3_matched_word_depth = match_l3_d4_item.get("depth", 4)

                # 尝试 2: 匹配 Excel L3 到 Word L3 (深度 3)，如果 L4/L5 匹配不理想
                if l3_best_match_item is None or l3_best_score < self.threshold:
                    found_l3_d3, match_l3_d3_item, score_l3_d3 = self.find_match(
                        excel_row["level3"],
                        word_hierarchy["level2"],  # 包含 Word 深度 3 的项目
                    )
                    # 只有当找到更好的匹配时才更新
                    if (
                        found_l3_d3
                        and score_l3_d3 >= self.threshold
                        and score_l3_d3 > l3_best_score
                    ):
                        l3_best_match_item = match_l3_d3_item
                        l3_best_score = score_l3_d3
                        l3_matched_word_depth = 3

                # 尝试 3: 匹配 Excel L3 到 Word L2 (深度 2)，如果 L4/L3 匹配不理想
                if l3_best_match_item is None or l3_best_score < self.threshold:
                    found_l3_d2, match_l3_d2_item, score_l3_d2 = self.find_match(
                        excel_row["level3"],
                        word_hierarchy["level1"],  # 包含 Word 深度 2 的项目
                    )
                    # 只有当找到更好的匹配时才更新
                    if (
                        found_l3_d2
                        and score_l3_d2 >= self.threshold
                        and score_l3_d2 > l3_best_score
                    ):
                        l3_best_match_item = match_l3_d2_item
                        l3_best_score = score_l3_d2
                        l3_matched_word_depth = 2

            # 根据最佳匹配结果设置 L3 匹配标志和报告字段
            l3_matched_flag = l3_best_match_item is not None

            if excel_row["level3"]:
                word_l3_title = (
                    l3_best_match_item.get("display", "❌ 缺失")
                    if l3_best_match_item
                    else "❌ 缺失"
                )
                word_l3_num = (
                    l3_best_match_item.get("number", "") if l3_best_match_item else ""
                )
            else:
                word_l3_title = ""
                word_l3_num = ""

            # --- DEBUG START ---
            if excel_row["level3"] and not l3_matched_flag and idx < 10:
                print(f"  [DEBUG-FAIL] L3 未找到匹配: '{excel_row['level3']}'")
            # --- DEBUG END ---

            # 构建完整的编号路径 (始终显示三段，缺失项用 - 表示)
            number_parts = []
            number_parts.append(word_l1_num if word_l1_num else "-")
            number_parts.append(word_l2_num if word_l2_num else "-")
            number_parts.append(word_l3_num if word_l3_num else "-")
            word_full_number = " / ".join(number_parts)

            # 构建 '匹配层级' 字符串
            matched_levels_str = []
            if l1_matched_flag:
                matched_levels_str.append("一级")
            if l2_matched_flag:
                matched_levels_str.append("二级")
            if l3_matched_flag:
                matched_levels_str.append(
                    f"三级(Word L{l3_matched_word_depth})"
                )  # 指明 Excel L3 实际匹配到的 Word 深度

            matching_level_text = (
                "、".join(matched_levels_str) + "匹配"
                if matched_levels_str
                else "无匹配"
            )

            # 确定最终的相似度分数（主要基于 L3 匹配，如果 L3 未匹配则基于 L2/L1）
            overall_similarity_score = 0.0
            if l3_matched_flag:
                overall_similarity_score = l3_best_score
            elif l2_matched_flag:
                overall_similarity_score = l2_score
            elif l1_matched_flag:
                overall_similarity_score = l1_score

            # --- 生成简略描述 (修订版 V11 - 严格按照 7 种情况格式化) ---
            l1_missing = bool(excel_row["level1"]) and not l1_matched_flag
            l2_missing = bool(excel_row["level2"]) and not l2_matched_flag
            l3_missing = bool(excel_row["level3"]) and not l3_matched_flag

            def is_descendant(parent_num, child_num):
                if not parent_num or not child_num: return True
                if parent_num == child_num: return True
                return child_num.startswith(parent_num + ".")

            n1 = word_l1_num if l1_matched_flag else None
            n2 = word_l2_num if l2_matched_flag else None
            n3 = word_l3_num if l3_matched_flag else None
            
            l1_l2_ok = is_descendant(n1, n2) if (n1 and n2) else True
            l2_l3_ok = is_descendant(n2, n3) if (n2 and n3) else True
            l1_l3_ok = is_descendant(n1, n3) if (n1 and n3) else True
            path_broken_3_level = not l1_l2_ok or not l2_l3_ok

            l1_text = excel_row["level1_original"]
            l2_text = excel_row["level2_original"]
            l3_text = excel_row["level3_original"]
            
            def get_formatted_row_info(total_count, range_str):
                if not range_str: return ""
                if total_count > 1 or "-" in range_str:
                    return f"（{range_str}行）"
                return ""

            missing_description = "-"
            
            # --- 逻辑分支判定 ---
            
            # 1. 存在性缺失分支 (中括号 [])
            missing_parts = []
            if l1_missing: missing_parts.append(f"一级模块 [{l1_text}]")
            if l2_missing: missing_parts.append(f"二级模块 [{l2_text}]")
            if l3_missing: missing_parts.append(f"三级模块 [{l3_text}]")

            if missing_parts:
                count_missing = len(missing_parts)
                # 根据缺失数量生成前缀
                if count_missing == 1:
                    prefix = "拆分表一级模块在需求规格书未体现：" if l1_missing else \
                             "拆分表二级模块在需求规格书未体现：" if l2_missing else \
                             "拆分表三级模块在需求规格书未体现："
                    # 单级缺失处理行号
                    active_info = get_formatted_row_info(excel_row["l1_total_count"] if l1_missing else (excel_row["l2_total_count"] if l2_missing else excel_row["l3_total_count"]), 
                                                        excel_row["l1_range"] if l1_missing else (excel_row["l2_range"] if l2_missing else excel_row["l3_range"]))
                    missing_description = f"{prefix}拆分表{missing_parts[0]}{active_info} 在需求规格书未体现"
                else:
                    # 多级缺失 (Case 3, 4, 5, 6 in user example)
                    level_names = []
                    if l1_missing: level_names.append("一")
                    if l2_missing: level_names.append("二")
                    if l3_missing: level_names.append("三")
                    prefix = f"拆分表{''.join(level_names)}级模块在需求规格书未体现："
                    missing_description = f"{prefix}拆分表{'、'.join(missing_parts)} 在需求规格书未体现"
            
            # 2. 路径不匹配分支 (大括号 {}, 使用破折号 - 分隔)
            elif l1_matched_flag and l2_matched_flag and l3_matched_flag and path_broken_3_level:
                # 情况 4: 三级路径断裂
                missing_description = f"拆分表一二三级模块 与需求规格书一二三级目录不匹配：拆分表一级模块 {{{l1_text}}} - 二级模块 {{{l2_text}}} - 三级模块 {{{l3_text}}} 与一二三级目录不匹配"
            
            elif l1_matched_flag and l3_matched_flag and not l1_l3_ok:
                # 情况 5-7: 一三不匹配
                missing_description = f"拆分表一三级模块 与需求规格书一三级目录不匹配：拆分表一级模块 {{{l1_text}}} - 三级模块 {{{l3_text}}} 与一三级目录不匹配"
            elif l2_matched_flag and l3_matched_flag and not l2_l3_ok:
                # 情况 5-7: 二三不匹配
                missing_description = f"拆分表二三级模块 与需求规格书二三级目录不匹配：拆分表二级模块 {{{l2_text}}} - 三级模块 {{{l3_text}}} 与二三级目录不匹配"
            elif l1_matched_flag and l2_matched_flag and not l1_l2_ok:
                # 情况 5-7: 一二不匹配
                missing_description = f"拆分表一二级模块 与需求规格书一二级目录不匹配：拆分表一级模块 {{{l1_text}}} - 二级模块 {{{l2_text}}} 与一二级目录不匹配"

            # --- 汇总缺失层级文本 (供报表明示) ---
            missing_levels_list = []
            if l1_missing: missing_levels_list.append("一级")
            if l2_missing: missing_levels_list.append("二级")
            if l3_missing: missing_levels_list.append("三级")
            if "目录不匹配" in missing_description:
                if "一二三级" in missing_description: missing_levels_list.append("一二三级目录不匹配")
                elif "一二" in missing_description: missing_levels_list.append("一二目录不匹配")
                elif "一三" in missing_description: missing_levels_list.append("一三目录不匹配")
                elif "二三" in missing_description: missing_levels_list.append("二三目录不匹配")

            missing_levels_text = "、".join(missing_levels_list) if missing_levels_list else "-"

            # 构建报告中的数据行
            match_data = {
                "Excel一级模块": l1_text,
                "Word一级标题": word_l1_title,
                "Excel二级模块": l2_text,
                "Word二级标题": word_l2_title,
                "Excel三级模块": l3_text,
                "Word三级标题": word_l3_title,
                "Word完整编号": word_full_number,
                "相似度": (
                    f"{overall_similarity_score:.2%}"
                    if overall_similarity_score > 0
                    else ""
                ),
                "匹配层级": matching_level_text,
                "缺失层级": missing_levels_text,
                "匹配状态": "缺失" if (l1_missing or l2_missing or l3_missing or path_broken_3_level) else "匹配通过",
                "简略描述": missing_description,
                "row_range": excel_row.get("final_row_info", ""),
            }

            if match_data["匹配状态"] == "匹配通过":
                # 进一步区分精确还是模糊
                # 精确要求：所有存在的层级都匹配，且分数都很高
                l1_ok = (not excel_row["level1"]) or (
                    l1_matched_flag and l1_score >= 0.99
                )
                l2_ok = (not excel_row["level2"]) or (
                    l2_matched_flag and l2_score >= 0.99
                )
                l3_ok = (not excel_row["level3"]) or (
                    l3_matched_flag and l3_best_score >= 0.99
                )

                # 还要检查层级深度是否对应 (Word L4 对应 Excel L3)
                depth_ok = True
                if excel_row["level3"] and l3_matched_word_depth not in [4, 5]:
                    depth_ok = False 

                if l1_ok and l2_ok and l3_ok and depth_ok:
                    exact_matched.append(match_data)
                else:
                    reasons = []
                    if excel_row["level1"] and l1_score < 0.99: reasons.append("一级模糊")
                    if excel_row["level2"] and l2_score < 0.99: reasons.append("二级模糊")
                    if excel_row["level3"] and l3_best_score < 0.99: reasons.append("三级模糊")
                    if not depth_ok: reasons.append(f"层级错位(Word L{l3_matched_word_depth})")
                    if reasons: match_data["匹配层级"] += f" ({'; '.join(reasons)})"
                    fuzzy_matched.append(match_data)
            else:
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

        if progress_callback:
            progress_callback(100, "匹配完成！")

        report = {
            "exact_matched": exact_matched,
            "fuzzy_matched": fuzzy_matched,
            "not_found_in_word": not_found_in_word,
            "statistics": {
                "Excel功能点总数": total_excel,
                "Word四级编号总数": len(
                    word_hierarchy["level3"]
                ),  # 这里的统计可以更全面，但保持原样
                "精确匹配": exact_count,
                "模糊匹配": fuzzy_count,
                "已匹配": matched_count,
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
                        "简略描述",  # 新增简略描述
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

                    # 缺失项
                    if report["not_found_in_word"]:
                        for item in report["not_found_in_word"]:
                            item["匹配状态"] = "缺失"
                            all_data.append(item)
                        missing_df = pd.DataFrame(report["not_found_in_word"])
                        missing_cols_order = (
                            [
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
                                "简略描述",
                            ]
                            if is_hierarchical_report
                            else ["Excel功能点", "状态", "匹配状态"]
                        )

                        missing_df = missing_df[
                            [
                                col
                                for col in missing_cols_order
                                if col in missing_df.columns
                            ]
                        ]
                        missing_df.to_excel(writer, sheet_name="❌ 缺失项", index=False)
                    else:
                        empty_df = pd.DataFrame(
                            [{"说明": "🎉 恭喜！所有 Excel 功能点都在 Word 中找到了！"}]
                        )
                        empty_df.to_excel(writer, sheet_name="❌ 缺失项", index=False)

                    # 新增：总表
                    if all_data:
                        all_df = pd.DataFrame(all_data)
                        # 确保列顺序
                        cols = [
                            col for col in current_column_order if col in all_df.columns
                        ]
                        all_df = all_df[cols]
                        all_df.to_excel(
                            writer, sheet_name="📋 匹配详情总表", index=False
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
