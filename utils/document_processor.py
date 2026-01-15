import os
import re
import pandas as pd
from docx import Document
import tempfile
import shutil
import sys

# 确保可以导入 extend 目录下的模块
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
if project_root not in sys.path:
    sys.path.append(project_root)

from extend.hierarchical_matcher import HierarchicalMatcher
from extend.matcher_config import MatcherConfig
import json


class DocumentProcessor:
    """文档处理类，用于提取 Word 和 Excel 的内容结构"""

    @staticmethod
    def _convert_doc_to_docx(doc_path):
        """将 .doc 转换为临时 .docx 文件 (仅限 Windows)"""
        try:
            import win32com.client as win32

            word = win32.gencache.EnsureDispatch("Word.Application")
            word.Visible = False

            # 使用绝对路径
            abs_doc_path = os.path.abspath(doc_path)
            doc = word.Documents.Open(abs_doc_path)

            # 创建临时文件
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                tmp_path = tmp.name

            doc.SaveAs(tmp_path, FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
            doc.Close()
            # word.Quit() # 建议在外部统一管理或保持开启以提高性能
            return tmp_path
        except Exception as e:
            print(f".doc 转换失败: {e}")
            return None

    @staticmethod
    def extract_word_structure(file_path):
        """
        提取 Word 文档的全层级标题及正文
        """
        temp_docx = None
        try:
            if file_path.lower().endswith(".doc"):
                temp_docx = DocumentProcessor._convert_doc_to_docx(file_path)
                if not temp_docx:
                    return []
                doc_to_read = temp_docx
            else:
                doc_to_read = file_path

            doc = Document(doc_to_read)
            sections = []

            # 当前状态：跟踪最近的 1, 2, 3 级标题
            current_titles = {1: "", 2: "", 3: "", 4: ""}
            current_section = {
                "level": 0,
                "title": "前言/未归类",
                "content": [],
                "full_path": "前言",
            }

            toc_styles = ["toc", "目录", "TOC"]

            # 常见核心章节关键字
            core_keywords = [
                "需求说明",
                "总体描述",
                "建设目标",
                "建设必要性",
                "功能架构图",
                "功能需求",
                "关键时序图",
                "功能描述",
                "需求功能清单",
                "附加值",
                "规模因子",
                "应用领域",
                "质量及特性",
                "开发语言",
                "团队背景",
                "完整性",
            ]

            # 引导性文字黑名单 (防止误判为标题)
            instruction_blacklist = [
                "请在此处",
                "说明本项目",
                "示例（",
                "示例:",
                "图3.",
                "图4.",
                "本工程应",
                "建议进行",
            ]

            def get_level_enhanced(text, style_name, p_obj):
                if any(kw in text for kw in instruction_blacklist):
                    return None

                # 1. 样式判定
                style_lower = style_name.lower()
                if "heading 1" in style_lower or "标题 1" in style_lower:
                    return 1
                if "heading 2" in style_lower or "标题 2" in style_lower:
                    return 2
                if "heading 3" in style_lower or "标题 3" in style_lower:
                    return 3
                if "heading 4" in style_lower or "标题 4" in style_lower:
                    return 4

                # 2. 编号判定 (支持 1. 1.1 1、 (1) 一、)
                # 标准数字点: 1.1.1
                num_dot_match = re.match(r"^(\d+(\.\d+)*)\.?\s+", text)
                if num_dot_match:
                    return min(4, len(num_dot_match.group(1).split(".")))

                # 中文顿号: 1、 2、
                num_zh_match = re.match(r"^(\d+)[、]\s*", text)
                if num_zh_match:
                    return 1

                # 括号数字: (1) (2)
                num_paren_match = re.match(r"^\((\d+)\)\s*", text)
                if num_paren_match:
                    return 3

                # 3. 启发式判定: 加粗且短，或者是核心关键字
                is_bold = False
                if p_obj.runs:
                    # 只要第一个 run 加粗了就认为可能是标题
                    is_bold = any(
                        run.bold for run in p_obj.runs[:2] if run.text.strip()
                    )

                clean_t = re.sub(r"[^\u4e00-\u9fa5a-zA-Z]", "", text)
                is_core = any(kw in clean_t for kw in core_keywords)

                if (is_bold or is_core) and len(text) < 50:
                    # 核心词通常是一级或二级
                    if is_core and len(text) < 15:
                        return 2
                    return 3

                return None

            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                if not text:
                    continue

                style_name = paragraph.style.name
                if any(s in style_name.lower() for s in toc_styles):
                    continue
                if re.search(r"\.{5,}\s*\d+$", text):
                    continue  # 过滤目录页码行

                level = get_level_enhanced(text, style_name, paragraph)

                if level:
                    # 保存前一个章节
                    if (
                        current_section["title"] != "前言/未归类"
                        or current_section["content"]
                    ):
                        current_section["content"] = "\n".join(
                            current_section["content"]
                        ).strip()
                        sections.append(current_section)

                    # 更新路径跟踪
                    clean_title = text.strip()
                    current_titles[level] = clean_title
                    for l in range(level + 1, 5):
                        current_titles[l] = ""

                    path_parts = [
                        current_titles[l]
                        for l in range(1, level + 1)
                        if current_titles[l]
                    ]

                    current_section = {
                        "level": level,
                        "title": clean_title,
                        "content": [],
                        "full_path": " > ".join(path_parts),
                    }
                else:
                    current_section["content"].append(text)

            # 最后一个章节
            if current_section and (
                current_section["content"] or current_section["level"] > 0
            ):
                current_section["content"] = "\n".join(
                    current_section["content"]
                ).strip()
                sections.append(current_section)

            return sections
        except Exception as e:
            print(f"提取 Word 结构失败: {e}")
            return []
        finally:
            if temp_docx and os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except:
                    pass

    @staticmethod
    def get_functional_points_count(file_path, sheet_name=None):
        """
        获取 Excel 中有效功能点总数 (即数据行数)
        - 仅计算子功能过程
        - 跳过空行 (现由于用户反馈建议：改为遇到空行停止)
        - 遇到“注：”等提示文字停止遍历
        """
        try:
            import openpyxl
            from utils.runtime_logger import RuntimeLogger

            # 增加 read_only=True 提高大型/复杂 XML 文件的加载成功率
            # 这种模式比普通加载占用内存更少，且对某些特殊格式（如超长行）更鲁棒
            try:
                wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            except Exception as e:
                RuntimeLogger.log(f"Excel 初始加载失败: {e}，尝试标准模式加载...", level="WARN")
                wb = openpyxl.load_workbook(file_path, data_only=True)

            target_sheet_name = sheet_name

            if not target_sheet_name:
                for name in wb.sheetnames:
                    if "功能点拆分" in name:
                        target_sheet_name = name
                        break

            if not target_sheet_name or target_sheet_name not in wb.sheetnames:
                # 兜底：如果找不到且未指定，尝试第一个工作表
                if not sheet_name and wb.sheetnames:
                    target_sheet_name = wb.sheetnames[0]
                else:
                    return 0

            ws = wb[target_sheet_name]

            # 1. 探测表头区域以确定数据起点
            header_end = -1
            target_col = -1
            # read_only 模式下使用 iter_rows 效率最高
            probe_rows = list(ws.iter_rows(min_row=1, max_row=40, min_col=1, max_col=20))
            for row_idx, row_cells in enumerate(probe_rows, 1):
                row_vals = [str(cell.value) for cell in row_cells]
                row_str = " ".join([v for v in row_vals if v != "None"])
                
                # 常见的关键字段标记表头结束
                if any(
                    kw in row_str
                    for kw in ["客户需求", "一级模块", "功能过程", "子过程", "功能描述"]
                ):
                    header_end = row_idx
                    # 确定“子过程”所在的列
                    for col_idx, cell in enumerate(row_cells, 1):
                        val = str(cell.value or "")
                        if "子过程" in val:
                            target_col = col_idx
                            break
                    if target_col == -1:
                        for col_idx, cell in enumerate(row_cells, 1):
                            val = str(cell.value or "")
                            if "功能过程" in val:
                                target_col = col_idx
                                break
                    break

            if header_end == -1:
                header_end = 0  # 没有探测到则从头开始

            # 2. 遍历计数
            count = 0
            # 使用 iter_rows 从 header_end+1 开始遍历
            for r_idx, row in enumerate(ws.iter_rows(min_row=header_end + 1), header_end + 1):
                # 检查是否触底 (图 3 中的注记文字)
                row_note_found = False
                instructional_row = False
                
                # 预提取前几个单元格的字符串，用于检测说明文字
                row_head_texts = []
                for i in range(min(5, len(row))):
                    val = str(row[i].value or "").strip()
                    row_head_texts.append(val)
                    
                    if val.startswith("注：") or val.startswith("注:") or "请在正式提交时删除" in val:
                        row_note_found = True
                        break
                    
                    # 检查是否是模板说明文字 (Instructional Text)
                    # 这里的逻辑与 HierarchicalMatcher._is_instructional_text 保持一致
                    if val:
                        clean_v = re.sub(r"\s+", "", val)
                        for pattern in HierarchicalMatcher.IGNORE_PATTERNS:
                            if re.search(pattern.replace(" ", ""), clean_v):
                                instructional_row = True
                                break
                    if instructional_row: break

                if row_note_found:
                    break
                
                # 如果是说明提示行，跳过本行不计数
                if instructional_row:
                    continue
                
                # 检查整行是否为空 (探测前15列)
                row_has_data = False
                for i in range(min(15, len(row))):
                    if row[i].value is not None:
                        row_has_data = True
                        break
                
                if not row_has_data:
                    # 根据用户建议：整行为空就可以停止计数
                    break
                
                # 如果找到了目标列，则该列不为空才计数
                if target_col != -1 and target_col <= len(row):
                    if row[target_col-1].value is not None:
                        count += 1
                else:
                    count += 1

            if hasattr(wb, 'close'):
                wb.close()
            return count
        except Exception as e:
            try:
                from utils.runtime_logger import RuntimeLogger
                RuntimeLogger.log(f"获取功能点数失败: {e}", level="ERROR")
            except:
                pass
            print(f"获取功能点数失败: {e}")
            return 0

    @staticmethod
    def check_adjustment_factors_in_word(file_path):
        """
        深度扫描 Word 正文段落及表格中的附加值调整因子 (Node 4)
        """
        temp_docx = None
        try:
            from docx import Document
            import re

            # .doc 转换处理
            if file_path.lower().endswith(".doc"):
                temp_docx = DocumentProcessor._convert_doc_to_docx(file_path)
                if not temp_docx:
                    return {}
                doc_to_read = temp_docx
            else:
                doc_to_read = file_path

            doc = Document(doc_to_read)

            # 因子初始化
            factors = {
                "scale": {
                    "name": "需求变更规模因子",
                    "value": None,
                    "found_in_text": False,
                    "source": None,
                },
                "distributed": {
                    "name": "分布式处理",
                    "value": None,
                    "found_in_text": False,
                    "source": None,
                },
                "performance": {
                    "name": "性能",
                    "value": None,
                    "found_in_text": False,
                    "source": None,
                },
                "reliability": {
                    "name": "可靠性",
                    "value": None,
                    "found_in_text": False,
                    "source": None,
                },
                "multiple_sites": {
                    "name": "多重站点",
                    "value": None,
                    "found_in_text": False,
                    "source": None,
                },
            }

            # 1. 扫描段落 (正文文本查找)
            # 兼容带引号(中文/英文)和不带引号的情况
            scale_regex = re.compile(r'本项目目前属于\s*[“""]?(.+?)[”""]?\s*阶段')

            for para in doc.paragraphs:
                text = para.text.strip()
                if not text:
                    continue

                # A. 检查规模因子 (特定句式优先)
                # 句式: "本项目目前属于“结算”阶段" 或 "本项目目前属于预算阶段"
                m_scale = scale_regex.search(text)
                if m_scale:
                    factors["scale"]["found_in_text"] = True
                    factors["scale"]["value"] = m_scale.group(1).strip()
                    factors["scale"]["source"] = "text"

                # 备用：传统模糊匹配
                elif factors["scale"]["name"] in text and not factors["scale"]["value"]:
                    factors["scale"]["found_in_text"] = True
                    # 匹配常见的类型词
                    m = re.search(r"(结算|预算|概算|匡算)", text)
                    if m:
                        factors["scale"]["value"] = m.group(1)
                        factors["scale"]["source"] = "text"

                # B. 检查四项质量特性 (段落形式: "分布式处理: -1")
                # 特殊情况处理：如果整段话就是“无”，且没有具体因子关键字，可能意味着所有因子均为默认
                if text.strip() == "无" and "质量" not in text:
                    # 标记一个全局标识，稍后如果发现因子缺失，可以用这个补全
                    factors["_global_default"] = True

                for key in [
                    "distributed",
                    "performance",
                    "reliability",
                    "multiple_sites",
                ]:
                    if factors[key]["value"] is not None:
                        continue  # 已找到则跳过

                    kw = factors[key]["name"]
                    # 优化：限制匹配深度，避免在长段落的描述文字中误匹配到关键词
                    # 要求关键词出现在行首（允许有编号）或者紧随冒号
                    is_key_match = False
                    if text.startswith(kw) or re.search(rf"^[\d\s\.、\(\)（）]*{kw}", text):
                        # 如果匹配到了，确保它看起来像是一个标题或项的开始（后面有冒号或文本较短）
                        if re.search(rf"{kw}[:：\s]", text) or len(text) < 20:
                            is_key_match = True

                    if is_key_match:
                        # 排除目录行
                        if text.count(".") > 5:
                            continue

                        factors[key]["found_in_text"] = True
                        factors[key]["source"] = "text"

                        # 1. 查找明确的数字 (优先)
                        m_num = re.search(rf"{kw}.*?(-?\d+(\.\d+)?)", text)
                        if m_num:
                            factors[key]["value"] = m_num.group(1)
                            continue

                        # 2. 查找明确的“空”字 (视为异常)
                        # 例如: "分布式处理: 空"
                        if re.search(rf"{kw}.*?[:：]?\s*空", text):
                            factors[key]["value"] = "空"  # 显式赋值“空”字符串
                            continue

                        # 3. 查找特定的“负向/默认”语句 (只有特定文字才视为 -1)
                        negative_phrases = {
                            "distributed": ["没有明示对分布式处理的需求事项"],
                            "performance": ["没有明示对性能的特别需求事项或仅需提供基本性能"],
                            "reliability": ["没有明示对可靠性的特别需求事项或仅需提供基本的可靠性"],
                            "multiple_sites": ["在相同用途的硬件或软件环境下运行"],
                        }
                        
                        # 4. 查找特定的“正常/中性”语句 (视为 0)
                        neutral_phrases = {
                            "distributed": ["通过网络进行客户端/服务器及网络基础应用分布式处理和传输"],
                            "performance": ["应答时间或处理率对高峰时间或所有业务时间来说都很重要", "存在对连动系统结束处理时间的限制"],
                            "reliability": ["发生故障时带来较多不便或经济损失"],
                            "multiple_sites": ["在用途类似的硬件或软件环境下运行"],
                        }
                        
                        # 5. 查找特定的“增强/高要求”语句 (视为 1)
                        positive_phrases = {
                            "distributed": ["通过特别的设计保证在多个服务器及处理器上同时相互执行应用中的处理功能"],
                            "performance": ["要求设计阶段开始进行性能分析", "在设计、开发阶段使用分析工具"],
                            "reliability": ["发生故障时造成重大经济损失或有生命危害"],
                            "multiple_sites": ["在不同用途的硬件或软件环境下运行"],
                        }

                        found_val = None
                        for p in negative_phrases.get(key, []):
                            if p in text: found_val = "-1"; break
                        if not found_val:
                            for p in neutral_phrases.get(key, []):
                                if p in text: found_val = "0"; break
                        if not found_val:
                            for p in positive_phrases.get(key, []):
                                if p in text: found_val = "1"; break
                        
                        if found_val:
                            factors[key]["value"] = found_val
                            continue

                        # 6. 兜底：如果前面都没匹配到，但有冒号加内容，提取内容
                        # 例如: "分布式处理：高要求"
                        # 避免提取到空字符串
                        m_gen = re.search(rf"{kw}.*?[:：]\s*(\S+)", text)
                        if m_gen:
                            val_str = m_gen.group(1).strip()
                            # 再次过滤一下，如果提取出来的是“空”或“无”，修正一下
                            if val_str in ["空", "无"]:
                                factors[key]["value"] = "空"
                            else:
                                factors[key]["value"] = val_str

            # 如果检测到了 global "无"，且因子未提取到值，则设为默认
            if factors.get("_global_default"):
                for key in [
                    "distributed",
                    "performance",
                    "reliability",
                    "multiple_sites",
                ]:
                    if not factors[key]["found_in_text"]:
                        factors[key]["found_in_text"] = True
                        factors[key]["value"] = "无"  # 全局无
                        factors[key]["source"] = "text"  # 视为文本中找到

            # 2. 扫描表格 (如果段落没找全，或者表格更准确，优先以表格为准)
            # 很多文档会将质量特性放在表格中: [特性名] [描述] [分值]
            # 同时也检查表格中的“需求变更规模因子”
            quality_keys = [
                "distributed",
                "performance",
                "reliability",
                "multiple_sites",
            ]

            for table in doc.tables:
                # 检查表头是否包含“需求变更规模因子”
                is_scale_table = False
                try:
                    if "需求变更规模因子" in "".join(
                        [c.text for c in table.rows[0].cells]
                    ):
                        is_scale_table = True
                except:
                    pass

                for row in table.rows:
                    # 获取行内所有文本
                    cells_text = [cell.text.strip() for cell in row.cells]
                    row_content = " ".join(cells_text)

                    # 优先检查规模因子表格
                    if is_scale_table and not factors["scale"]["value"]:
                        # 扫描行内容寻找关键词
                        m = re.search(r"(结算|预算|概算|匡算)", row_content)
                        if m:
                            factors["scale"]["found_in_text"] = True
                            factors["scale"]["value"] = m.group(1)
                            factors["scale"]["source"] = "table"

                    for key in quality_keys:
                        # 注意：此处允许表格覆盖段落提取的结果 (表格通常比正文提及的数字更准确)
                        # 如果表格中有明确的数字分值，无论正文提取到什么，都应该以表格分值为准
                        kw = factors[key]["name"]
                        
                        # 优化：仅在行前两个单元格中搜索关键字，避免匹配到后方的说明文字（如可靠性说明中包含“性能”字样）
                        if any(kw in str(c) for c in cells_text[:2]):
                            factors[key]["found_in_text"] = True  # 标记为已找到

                            # 尝试获取分值单元格
                            # 规则：寻找行内最后一个能匹配数字的单元格
                            possible_val = None
                            for cell_text in reversed(cells_text):
                                clean_cell = (
                                    cell_text.strip()
                                    .replace(" ", "")
                                    .replace("\n", "")
                                    .replace("\r", "")
                                )
                                if clean_cell and re.match(
                                    r"^-?\d+(\.\d+)?$", clean_cell
                                ):
                                    possible_val = clean_cell
                                    break

                            if possible_val:
                                factors[key]["value"] = possible_val
                                factors[key]["source"] = "table"
                            else:
                                # 如果分值单元格为空，尝试通过描述（判断标准列）自动计算
                                # 判断标准通常在第二列 cell_text[1] 或者 row_content 中
                                standards = {
                                    "distributed": {
                                        "-1": "没有明示对分布式处理的需求事项",
                                        "0": "通过网络进行客户端/服务器及网络基础应用分布式处理和传输",
                                        "1": "通过特别的设计保证在多个服务器及处理器上同时相互执行应用中的处理功能"
                                    },
                                    "performance": {
                                        "-1": "没有明示对性能的特别需求事项或仅需提供基本性能",
                                        "0": "存在对连动系统结束处理时间的限制",
                                        "1": "要求设计阶段开始进行性能分析"
                                    },
                                    "reliability": {
                                        "-1": "没有明示对可靠性的特别需求事项或仅需提供基本的可靠性",
                                        "0": "发生故障时带来较多不便或经济损失",
                                        "1": "发生故障时造成重大经济损失或有生命危害"
                                    },
                                    "multiple_sites": {
                                        "-1": "在相同用途的硬件或软件环境下运行",
                                        "0": "在用途类似的硬件或软件环境下运行",
                                        "1": "在不同用途的硬件或软件环境下运行"
                                    }
                                }
                                
                                normalized_row = row_content.replace(" ", "").replace("\n", "").replace("\r", "")
                                for val, desc in standards.get(key, {}).items():
                                    if desc.replace(" ", "") in normalized_row:
                                        factors[key]["value"] = val
                                        factors[key]["source"] = "table_auto"
                                        break
                                
                                # 如果还是没找到值，且没有正文值，才设为 None
                                if not factors[key]["value"]:
                                    factors[key]["value"] = None
                                    factors[key]["source"] = "table"

            return factors
        except Exception as e:
            print(f"提取调整因子失败: {e}")
            return {}
        finally:
            if temp_docx and os.path.exists(temp_docx):
                try:
                    os.remove(temp_docx)
                except:
                    pass

    @staticmethod
    def check_excel_empty_cells(file_path, sheet_name=None):
        """
        检查 Excel 的空值情况 (精准合并单元格判定版 V3)
        """
        try:
            import openpyxl

            wb = openpyxl.load_workbook(file_path, data_only=True)
            target_sheet_name = sheet_name
            if not target_sheet_name:
                for name in wb.sheetnames:
                    if "功能点拆分" in name:
                        target_sheet_name = name
                        break

            if not target_sheet_name or target_sheet_name not in wb.sheetnames:
                if not sheet_name and wb.sheetnames:
                    target_sheet_name = wb.sheetnames[0]
                else:
                    return {
                        "is_ok": False,
                        "errors": ["未找到有效的工作表进行空值校验"],
                    }

            ws = wb[target_sheet_name]

            # 1. 探测表头区域 (多行探测)
            header_start = -1
            header_end = -1
            # 探测前 30 行，寻找核心关键字
            for row_idx in range(1, 31):
                row_vals = [
                    str(ws.cell(row=row_idx, column=col).value) for col in range(1, 20)
                ]
                row_str = " ".join([v for v in row_vals if v != "None"])
                if "客户需求" in row_str or "一级模块" in row_str:
                    if header_start == -1:
                        header_start = row_idx
                    header_end = row_idx

            if header_start == -1:
                return {
                    "is_ok": False,
                    "errors": ["未能在工作表中定位到“客户需求”或“一级模块”表头行"],
                }

            # 2. 定位关键校验列
            target_keywords = [
                "客户需求",
                "一级模块",
                "二级模块",
                "三级模块",
                "功能用户",
                "触发事件",
                "功能过程",
                "子过程描述",
                "数据移动类型",
                "数据组",
                "数据属性",
                "复用度",
                "CFP",
            ]

            col_map = {}  # {keyword: col_index_1_based}
            for col_idx in range(1, ws.max_column + 1):
                # 检查 header_end 这一行，或其上方的表头行
                cell_val = ""
                for h_idx in range(header_start, header_end + 1):
                    val = ws.cell(row=h_idx, column=col_idx).value
                    if val:
                        cell_val += str(val)

                for kw in target_keywords:
                    if kw in cell_val:
                        col_map[kw] = col_idx
                        break

            if not col_map:
                return {
                    "is_ok": False,
                    "errors": ["未匹配到任何待校验的关键列，请检查表头名称"],
                }

            # 3. 确定有效数据范围
            last_valid_row = header_end
            for r in range(header_end + 1, ws.max_row + 1):
                # 检查是否触底 (图3中的注记文字)
                first_cell_val = str(ws.cell(row=r, column=1).value or "").strip()
                if first_cell_val.startswith("注：") or "请在正式提交时删除" in first_cell_val:
                    break
                
                # 检查整行是否有数据
                row_has_something = False
                for c in range(1, 21): # 检测前20列
                    if ws.cell(row=r, column=c).value is not None:
                        row_has_something = True
                        break
                
                if row_has_something:
                    last_valid_row = r

            if last_valid_row <= header_end:
                # 可能是个空表，除了表头没数据
                return {"is_ok": True, "errors": []}

            # 4. 建立合并单元格查询表 (row, col) -> (top_left_value)
            merged_lookup = {}
            for merged_range in ws.merged_cells.ranges:
                min_col, min_row, max_col, max_row = merged_range.bounds
                tl_val = ws.cell(row=min_row, column=min_col).value
                for r in range(min_row, max_row + 1):
                    for c in range(min_col, max_col + 1):
                        merged_lookup[(r, c)] = tl_val

            # 5. 遍历扫描 (核心：不再跳过整天空行)
            col_errors = {kw: [] for kw in col_map.keys()}

            for r_idx in range(header_end + 1, last_valid_row + 1):
                for kw, c_idx in col_map.items():
                    val = ws.cell(row=r_idx, column=c_idx).value

                    # 确定最终判定值
                    actual_val = val
                    if (
                        val is None
                        or str(val).strip() == ""
                        or str(val).lower() == "nan"
                    ):
                        if (r_idx, c_idx) in merged_lookup:
                            actual_val = merged_lookup[(r_idx, c_idx)]

                    if (
                        actual_val is None
                        or str(actual_val).strip() == ""
                        or str(actual_val).lower() == "nan"
                    ):
                        col_errors[kw].append(r_idx)

            # 6. 格式化错误信息 (合并连续行)
            final_errors = []
            for kw in target_keywords:  # 按预定义顺序排列
                if kw not in col_errors:
                    continue
                rows = sorted(list(set(col_errors[kw])))
                if not rows:
                    continue

                ranges = []
                start = rows[0]
                prev = start
                for curr in rows[1:]:
                    if curr == prev + 1:
                        prev = curr
                    else:
                        ranges.append(
                            f"{start}-{prev}" if start != prev else f"{start}"
                        )
                        start = curr
                        prev = curr
                ranges.append(f"{start}-{prev}" if start != prev else f"{start}")
                final_errors.append(f"列 [{kw}] -> 第 {', '.join(ranges)} 行为空")

            return {"is_ok": len(final_errors) == 0, "errors": final_errors}
        except Exception as e:
            return {"is_ok": False, "errors": [f"Excel 校验引擎异常 (V3): {str(e)}"]}

    @staticmethod
    def extract_excel_info(file_path):
        """
        提取 Excel 的工作表和列信息
        """
        try:
            excel_file = pd.ExcelFile(file_path)
            info = {}
            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=5)
                info[sheet_name] = df.columns.tolist()
            return info
        except Exception as e:
            print(f"提取 Excel 信息失败: {e}")
            return {}

    @staticmethod
    def get_excel_modules(file_path, sheet_name=None):
        """
        提取 Excel 中的三级模块名称集合
        """
        try:
            excel_file = pd.ExcelFile(file_path)
            target_sheet = sheet_name
            if not target_sheet:
                for sheet in excel_file.sheet_names:
                    if "功能点拆分" in sheet:
                        target_sheet = sheet
                        break

            if not target_sheet or target_sheet not in excel_file.sheet_names:
                if not sheet_name and excel_file.sheet_names:
                    target_sheet = excel_file.sheet_names[0]
                else:
                    return []

            # 1. 探测表头区域 (复用探测逻辑)
            df_preview = pd.read_excel(
                file_path, sheet_name=target_sheet, nrows=30, header=None
            )
            header_start = -1
            header_end = -1
            for i, row in df_preview.iterrows():
                row_str = " ".join([str(x) for x in row if pd.notna(x)])
                if "三级模块" in row_str or "一级模块" in row_str:
                    if header_start == -1:
                        header_start = i
                    header_end = i
            if header_start == -1:
                return []

            # 2. 读取数据
            df = pd.read_excel(
                file_path, sheet_name=target_sheet, skiprows=header_start
            )

            # 使用最后一行作为列名
            actual_header_row = df_preview.iloc[header_end]
            final_col_names = [
                str(x) if pd.notna(x) else f"Unnamed_{i}"
                for i, x in enumerate(actual_header_row)
            ]
            df.columns = final_col_names

            # 跳过表头
            df = df.iloc[header_end - header_start + 1 :].reset_index(drop=True)

            if "三级模块" in df.columns:
                return df["三级模块"].ffill().dropna().unique().tolist()
            return []
        except:
            return []

    @staticmethod
    def validate_hierarchy_matching(
        word_path,
        excel_path,
        header_row=0,
        level1_col=None,
        level2_col=None,
        level3_col=None,
        sheet_name=None,
        fuzzy_match=True,
        threshold=0.8,
    ):
        """
        节点5：层级匹配校验
        使用 HierarchicalMatcher 进行 Excel 一二三级模块与 Word 标题的层级对应校验
        """
        try:
            import pandas as pd
            import os

            matcher = HierarchicalMatcher(fuzzy_match=fuzzy_match, threshold=threshold)

            # 加载配置
            config = MatcherConfig.load()
            h_config = config.get(
                "hierarchy", MatcherConfig.get_defaults()["hierarchy"]
            )

            # 使用配置中的列索引 (如果参数未提供，则使用配置值)
            l1_c = (
                level1_col if level1_col is not None else h_config.get("level1_col", 1)
            )
            l2_c = (
                level2_col if level2_col is not None else h_config.get("level2_col", 2)
            )
            l3_c = (
                level3_col if level3_col is not None else h_config.get("level3_col", 3)
            )

            # 确定要使用的工作表
            target_sheet_name = sheet_name
            if target_sheet_name is None:
                sheet_idx = h_config.get(
                    "sheet_name", 2
                )  # Default to 3rd sheet (index 2)
                try:
                    xl = pd.ExcelFile(excel_path)
                    sheet_names = xl.sheet_names
                    if len(sheet_names) > sheet_idx:
                        target_sheet_name = sheet_names[sheet_idx]
                    elif len(sheet_names) > 0:
                        target_sheet_name = sheet_names[0]
                except Exception as e:
                    print(f"Error checking sheets: {e}")
                    target_sheet_name = 0

            print(
                f"[DEBUG] validate_hierarchy_matching: sheet={target_sheet_name}, l1={l1_c}, l2={l2_c}, l3={l3_c}, header={header_row}"
            )

            report = matcher.match_hierarchical_documents(
                word_path,
                excel_path,
                sheet_name=target_sheet_name,
                level1_col=l1_c,
                level2_col=l2_c,
                level3_col=l3_c,
                header=header_row,
            )

            # 保存报告
            base_name = os.path.splitext(os.path.basename(word_path))[0]
            report_filename = f"{base_name}-层级匹配报告.xlsx"
            saved_path = matcher.save_report(report, report_filename)
            report["report_path"] = saved_path

            return report
        except Exception as e:
            print(f"层级匹配校验失败: {e}")
            import traceback

            traceback.print_exc()
            return {"is_valid": False, "error": str(e)}

    @staticmethod
    def validate_functional_process(
        word_path,
        excel_path,
        header_row=0,
        func_col=None,
        sheet_name=None,
        fuzzy_match=True,
        threshold=0.8,
    ):
        """
        节点6：功能过程校验
        使用 HierarchicalMatcher (简单模式) 校验 Excel [功能过程] 在 Word 中的匹配情况
        """
        try:
            # 加载配置
            config = MatcherConfig.load()
            p_config = config.get("process", MatcherConfig.get_defaults()["process"])
            # 使用参数提供的列，否则使用配置
            func_proc_col = (
                func_col if func_col is not None else p_config.get("column", 6)
            )  # 0-indexed

            matcher = HierarchicalMatcher(fuzzy_match=fuzzy_match, threshold=threshold)

            # 确定要使用的工作表
            target_sheet = sheet_name
            if target_sheet is None:
                sheet_idx = p_config.get("sheet_name", 2)
                try:
                    xl = pd.ExcelFile(excel_path)
                    sheet_names = xl.sheet_names
                    # Sheet selection fallback
                    if len(sheet_names) > sheet_idx:
                        target_sheet = sheet_names[sheet_idx]
                    elif len(sheet_names) > 0:
                        target_sheet = sheet_names[0]
                    else:
                        target_sheet = 0
                except Exception as e:
                    print(f"Error checking excel structure: {e}")
                    target_sheet = 0

            print(
                f"[DEBUG] validate_functional_process: sheet={target_sheet}, func_col={func_proc_col}, header_row={header_row}"
            )

            try:
                # Column selection fallback
                # Ensure we can read the header to check columns
                # MUST pass the correct header row, otherwise index detection counts might be wrong due to merging in row 0
                df_header = pd.read_excel(
                    excel_path, sheet_name=target_sheet, nrows=0, header=header_row
                )
                if func_proc_col >= len(df_header.columns):
                    print(
                        f"[DEBUG] Column {func_proc_col} out of range (max {len(df_header.columns)-1}), fallback to 0"
                    )
                    func_proc_col = 0
            except Exception as e:
                print(f"Error checking column range: {e}")
                pass

            report = matcher.match_documents(
                word_path,
                excel_path,
                sheet_name=target_sheet,
                header=header_row,
                column=func_proc_col,
                full_text_search=True,  # 功能过程通常在正文中
            )

            # 保存报告
            base_name = os.path.splitext(os.path.basename(word_path))[0]
            report_filename = f"{base_name}-功能过程报告.xlsx"
            saved_path = matcher.save_report(report, report_filename)
            report["report_path"] = saved_path

            return report
        except Exception as e:
            print(f"功能过程校验失败: {e}")
            return {"is_valid": False, "error": str(e)}
