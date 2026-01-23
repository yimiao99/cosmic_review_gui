import os
import re
import pandas as pd
import openpyxl
from datetime import datetime
from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_COLOR_INDEX
from PySide6.QtCore import QThread, Signal
from extend.matcher_config import MatcherConfig


class ReceiptWorker(QThread):
    """异步处理回单生成的线程"""

    finished = Signal(dict)  # 改为发送字典，包含路径和统计数据
    error = Signal(str)

    def __init__(self, data, project_name=None):
        super().__init__()
        self.data = data
        self.project_name = project_name or "Receipt"

    def run(self):
        from utils.runtime_logger import RuntimeLogger

        # 设置项目名称用于日志
        RuntimeLogger.set_project(self.project_name)
        RuntimeLogger.log("开始生成回单...")

        try:
            output_path, stats = ReceiptProcessor.generate(self.data)
            self.finished.emit({"output_path": output_path, "stats": stats})
        except Exception as e:
            import traceback

            traceback.print_exc()
            self.error.emit(str(e))


class ReceiptProcessor:
    """回单生成处理器"""

    @staticmethod
    def generate(data):
        # 1. 提取评估报告数据 (如果存在)
        fp_data = {}
        if data.get("eval_report_path"):
            fp_data = ReceiptProcessor._parse_eval_report(data["eval_report_path"])
        else:
            fp_data = {
                "total_fp": "N/A",
                "new_fp": "N/A",
                "reuse_fp": "N/A",
                "legacy_fp": "N/A",
                "new_ratio": "N/A",
                "reuse_ratio": "N/A",
                "legacy_ratio": "N/A",
                "total_ratio": "N/A",
            }

        # 2. 提取评估认同表数据 (如果存在)
        consent_data = {}
        if data.get("eval_consent_path"):
            consent_data = ReceiptProcessor._parse_consent_form(
                data["eval_consent_path"]
            )
        else:
            consent_data = {
                "submission_days": "N/A",
                "eval_days": "N/A",
                "reduction_ratio": "N/A",
            }

        # 3. 选择模式 (严格对应按钮选择)
        mode = str(data.get("submission_mode", "线上"))  # 默认线上

        # 4. 合并所有需要填充的数据
        context = {
            "project_name": data["project_name"],
            "project_id": data["project_id"],
            "submission_unit": data["submission_unit"],
            "submitter": data["submitter"],
            "submission_time": data["submission_time"],
            "submission_mode": mode,
            "_consent_path": data["eval_consent_path"],  # 保留路径用于 Excel 公式
            **fp_data,
            **consent_data,
        }

        template_name = f"XXXXXXX项目评估确认单 - {mode}.docx"
        template_path = os.path.join("folder", template_name)
        if not os.path.exists(template_path):
            # 兼容性处理
            template_path = os.path.join(
                os.path.dirname(__file__), "..", "folder", template_name
            )

        if not os.path.exists(template_path):
            raise FileNotFoundError(f"找不到模板文件: {template_path}")

        # 5. 填充模板并保存
        doc = Document(template_path)
        ReceiptProcessor._fill_template(doc, context)

        output_name = f"{data['project_name']}项目评估确认单.docx"
        # 从配置中加载回单存放位置
        config = MatcherConfig.load()
        output_dir = config.get("storage", {}).get("receipt")

        if not output_dir:
            # 兼容：优先保存在报告所在目录
            output_dir = (
                os.path.dirname(data["eval_report_path"])
                if data.get("eval_report_path")
                else "."
            )

        final_path = os.path.join(output_dir, output_name)
        counter = 1
        while True:
            try:
                # 检查输出目录是否存在
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                doc.save(final_path)
                break
            except (IOError, PermissionError):
                # 如果文件正在被打开，则生成副本
                base, ext = os.path.splitext(output_name)
                final_path = os.path.join(output_dir, f"{base}({counter}){ext}")
                counter += 1

        # 6. 回写结果到评估报告 PQR 列 (使用副本回写以避免被占用)
        if data.get("eval_report_path"):
            try:
                ReceiptProcessor._write_back_to_excel(data["eval_report_path"], context)
            except RuntimeError as e:
                # 如果回写失败（通常是文件占用），生成一个带 _结果 的副本
                base, ext = os.path.splitext(data["eval_report_path"])
                copy_path = f"{base}_回写结果{ext}"
                import shutil

                try:
                    shutil.copy2(data["eval_report_path"], copy_path)
                    ReceiptProcessor._write_back_to_excel(copy_path, context)
                    # 如果回写成功，通知或者记录日志（由于目前是直接返回 output_path，这里仅作为防御逻辑）
                except Exception:
                    pass

        # 7. 计算统计数据用于卡片显示
        stats = {
            "new_fp": fp_data.get("new_fp", "N/A"),
            "reuse_fp": fp_data.get("reuse_fp", "N/A"),
            "legacy_fp": fp_data.get("legacy_fp", "N/A"),
            "total_fp": fp_data.get("total_fp", "N/A"),
            "new_ratio": fp_data.get("new_ratio", "0.0%"),
            "reuse_ratio": fp_data.get("reuse_ratio", "0.0%"),
            "legacy_ratio": fp_data.get("legacy_ratio", "0.0%"),
            "total_ratio": fp_data.get("total_ratio", "100.0%"),
            "submission_days": consent_data.get("submission_days", "0.00"),
            "eval_days": consent_data.get("eval_days", "0.00"),
            "reduction_ratio": consent_data.get("reduction_ratio", "0.00%"),
        }

        return final_path, stats

    @staticmethod
    def _write_back_to_excel(path, context):
        """将统计结果和结论表数据一并回写到评估报告 Excel 的 PQR 列"""
        try:
            wb = openpyxl.load_workbook(path)
            # 查找功能点拆分表
            ws = None
            for sheet in wb.sheetnames:
                if "功能点拆分表" in sheet:
                    ws = wb[sheet]
                    break
            if not ws:
                ws = wb.active

            # 找到起始行 (从 Q 列第 10 行开始寻找第一个空行)
            start_row = 10
            for r in range(10, 100):
                if (
                    ws.cell(row=r, column=17).value is None
                    and ws.cell(row=r + 1, column=17).value is None
                ):
                    start_row = r
                    break

            # P(16): 标签, Q(17): 数值(公式), R(18): 占比(公式)
            col_letter = openpyxl.utils.get_column_letter(12)  # L
            total_r_idx = start_row + 3

            # 1. 功能点公式回写
            fp_rows = [
                (
                    "新增",
                    f'=COUNTIF({col_letter}:{col_letter}, "新增")',
                    f"=Q{{row}}/Q{total_r_idx}",
                ),
                (
                    "复用",
                    f'=COUNTIF({col_letter}:{col_letter}, "复用")',
                    f"=Q{{row}}/Q{total_r_idx}",
                ),
                (
                    "利旧",
                    f'=COUNTIF({col_letter}:{col_letter}, "利旧")',
                    f"=Q{{row}}/Q{total_r_idx}",
                ),
                ("合计", f"=Q{start_row}+Q{start_row+1}+Q{start_row+2}", "100.0%"),
            ]

            curr_r = start_row
            for i, (label, val_fmt, ratio_fmt) in enumerate(fp_rows):
                ws.cell(row=curr_r, column=16).value = label
                ws.cell(row=curr_r, column=17).value = val_fmt
                if i < 3:
                    ws.cell(row=curr_r, column=18).value = ratio_fmt.replace(
                        "{row}", str(curr_r)
                    )
                else:
                    ws.cell(row=curr_r, column=18).value = "100.0%"
                curr_r += 1

            # 2. 结论认同表数值回写 (使用绝对路径公式，解决闭卷 #REF! 问题)
            curr_r += 2
            consent_path = context.get("_consent_path")
            if consent_path:
                abs_path = os.path.abspath(consent_path)
                dir_name = os.path.dirname(abs_path)
                file_name = os.path.basename(abs_path)
                sheet_name = context.get("found_sheet") or "结论认同表"

                # Excel 闭卷引用格式: ='C:\路径\[文件名.xlsx]工作表'!$A$1
                # 注意：跨文件引用必须使用绝对路径，否则关闭源文件后会显示 #REF!
                ref_prefix = f"='{dir_name}\\[{file_name}]{sheet_name}'"

                consent_data = [
                    (
                        "送审人天",
                        f"{ref_prefix}!{context.get('submission_addr', '$C$3')}",
                        "",
                    ),
                    (
                        "核定人天",
                        f"{ref_prefix}!{context.get('eval_addr', '$D$3')}",
                        "",
                    ),
                    (
                        "核减比例",
                        f"{ref_prefix}!{context.get('ratio_addr', '$E$3')}",
                        "",
                    ),
                ]
            else:
                consent_data = [
                    ("送审人天", context.get("submission_days", "0"), ""),
                    ("核定人天", context.get("eval_days", "0.00"), ""),
                    ("核减比例", context.get("reduction_ratio", "0.0%"), ""),
                ]

            for label, val, ratio in consent_data:
                ws.cell(row=curr_r, column=16).value = label
                ws.cell(row=curr_r, column=17).value = val
                ws.cell(row=curr_r, column=18).value = ratio
                curr_r += 1

            wb.save(path)
        except PermissionError:
            raise RuntimeError(
                f"无法保存 Excel：文件已被打开，请关闭 {os.path.basename(path)}"
            )
        except Exception as e:
            raise RuntimeError(f"Excel 回写失败: {str(e)}")

    @staticmethod
    def _parse_eval_report(path):
        """解析评估报告获取项目统计 (严格基于复用度列行数统计)"""
        try:
            df_dict = pd.read_excel(path, sheet_name=None)
        except Exception as e:
            raise RuntimeError(f"无法读取评估报告: {str(e)}")

        target_df = None
        for name, df in df_dict.items():
            if "功能点拆分表" in name:
                target_df = df
                break

        if target_df is None:
            target_df = list(df_dict.values())[0]

        # 1. 查找“复用度”列索引
        type_col = -1
        start_data_idx = 0
        for i in range(min(15, len(target_df))):
            row = target_df.iloc[i].values
            for idx, val in enumerate(row):
                if isinstance(val, str) and "复用度" in val:
                    type_col = idx
                    start_data_idx = i + 1
                    break
            if type_col != -1:
                break

        if type_col == -1:
            type_col = 11  # 默认 L 列

        # 2. 统计逻辑：严格统计行数
        new_count = 0
        reuse_count = 0
        legacy_count = 0

        for idx in range(start_data_idx, len(target_df)):
            row = target_df.iloc[idx]
            val = str(row.iloc[type_col]) if type_col < len(row) else ""
            clean_val = val.strip()
            if "新增" in clean_val:
                new_count += 1
            elif "复用" in clean_val or "优化" in clean_val:
                reuse_count += 1
            elif "利旧" in clean_val:
                legacy_count += 1

        total_fp = new_count + reuse_count + legacy_count

        def safe_ratio(part, total):
            if total == 0:
                return "0.0%"
            return f"{(part / total):.1%}"

        return {
            "total_fp": total_fp,
            "new_fp": new_count,
            "reuse_fp": reuse_count,
            "legacy_fp": legacy_count,
            "new_ratio": safe_ratio(new_count, total_fp),
            "reuse_ratio": safe_ratio(reuse_count, total_fp),
            "legacy_ratio": safe_ratio(legacy_count, total_fp),
            "total_ratio": "100.0%",  # 合计占比始终是100%
        }

    @staticmethod
    def _parse_consent_form(path):
        """解析结论认同表并记录数值所在的单元格地址，以便生成公式"""
        results = {
            "submission_days": "0",
            "submission_addr": "$C$3",
            "eval_days": "0.00",
            "eval_addr": "$D$3",
            "reduction_ratio": "0.0%",
            "ratio_addr": "$E$3",
            "found_sheet": None,
        }

        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            header_map = {
                "送审人天": ("submission_days", "submission_addr"),
                "核定人天": ("eval_days", "eval_addr"),
                "评估人天": ("eval_days", "eval_addr"),
                "核定工作量": ("eval_days", "eval_addr"),
                "评估工作量": ("eval_days", "eval_addr"),
                "核减比例": ("reduction_ratio", "ratio_addr"),
                "核减率": ("reduction_ratio", "ratio_addr"),
            }
            all_header_keys = list(header_map.keys())

            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                for r in range(1, 101):
                    for c in range(1, 31):
                        cell_val = ws.cell(row=r, column=c).value
                        if cell_val and isinstance(cell_val, str):
                            clean_text = (
                                cell_val.replace(" ", "")
                                .replace("\n", "")
                                .replace("\r", "")
                            )
                            for key, (attr, addr_attr) in header_map.items():
                                if key in clean_text:
                                    # 如果已经在一个 Sheet 找到了数据，且当前 Sheet 不是那个 Sheet，则跳过（除非还没找到 Sheet）
                                    if (
                                        results["found_sheet"]
                                        and results["found_sheet"] != sheet_name
                                    ):
                                        # 如果这个属性已经有值了，就不再从别的 Sheet 覆盖
                                        if (
                                            results[attr] != "0"
                                            and results[attr] != "0.00"
                                            and results[attr] != "0.0%"
                                        ):
                                            continue

                                    # 智能探测：全方位探测数值 (先右后下)
                                    # 1. 探测右侧单元格 (需排除右侧也是另一个表头的情况)
                                    val_right = ws.cell(row=r, column=c + 1).value
                                    is_right_valid = val_right is not None
                                    if is_right_valid and isinstance(val_right, str):
                                        rt_clean = val_right.replace(" ", "")
                                        # 如果右侧单元格内容包含任一已知表头关键字，则认为它是下一个表头
                                        if any(k in rt_clean for k in all_header_keys):
                                            is_right_valid = False

                                    # 2. 探测下方单元格 (跳过可能的表头)
                                    target_r = r + 1
                                    if target_r <= 2:
                                        target_r = 3
                                    val_below = ws.cell(row=target_r, column=c).value

                                    # 优先级：右侧(非表头) > 下方
                                    if is_right_valid and (
                                        isinstance(val_right, (int, float))
                                        or (
                                            isinstance(val_right, str)
                                            and len(val_right.strip()) > 0
                                        )
                                    ):
                                        val_cell = ws.cell(row=r, column=c + 1)
                                        target_col = c + 1
                                        target_row = r
                                    else:
                                        val_cell = ws.cell(row=target_r, column=c)
                                        target_col = c
                                        target_row = target_r

                                    results["found_sheet"] = sheet_name

                                    # 记录绝对地址，如 $C$3
                                    col_letter = openpyxl.utils.get_column_letter(
                                        target_col
                                    )
                                    results[addr_attr] = f"${col_letter}${target_row}"

                                    val = val_cell.value
                                    if val is not None:
                                        if isinstance(val, (int, float)):
                                            if "比例" in key:
                                                results[attr] = (
                                                    f"{val:.2%}"
                                                    if val < 1
                                                    else f"{val}%"
                                                )
                                            else:
                                                results[attr] = f"{float(val):.2f}"
                                        else:
                                            # 处理文本形式的百分比
                                            results[attr] = str(val).strip()
                                    break  # 找到一个关键词就跳出关键词循环
            wb.close()
        except Exception:
            pass
        return results

    @staticmethod
    def _fill_template(doc, context):
        """填充 Word 模板 (根据语义隔离精准匹配数据)"""

        # 1. 准备数据包 (强制总数为整数)
        try:
            total_fp = int(float(context.get("total_fp", 0)))
        except:
            total_fp = context.get("total_fp", "0")

        data_bundle = {
            "total_fp": str(total_fp),
            "new_fp": str(context.get("new_fp", "0")),
            "new_ratio": str(context.get("new_ratio", "0%")),
            "reuse_fp": str(context.get("reuse_fp", "0")),
            "reuse_ratio": str(context.get("reuse_ratio", "0%")),
            "legacy_fp": str(context.get("legacy_fp", "0")),
            "legacy_ratio": str(context.get("legacy_ratio", "0%")),
            "submission_days": str(context.get("submission_days", "0")),
            "eval_days": str(context.get("eval_days", "0")),
            "reduction_ratio": str(context.get("reduction_ratio", "0%")),
        }

        # 2. 普通文本替换 (全局兜底)
        reduction_val = context.get("reduction_ratio", "0%")

        # 处理送审时间 (从 yyyy-MM-dd 转为 中文年月日)
        date_str = context["submission_time"]
        try:
            dt = datetime.strptime(date_str, "%Y-%m-%d")
            cn_date = (
                dt.strftime("%Y年%m月%d日").replace("年0", "年").replace("月0", "月")
            )
        except:
            cn_date = date_str

        # 准备基础替换词
        text_replacements = {
            "2025 年 x 月 x 日": cn_date,
            "2025年x月x日": cn_date,
            "2025年X月X日": cn_date,
            "2026 年 x 月 x 日": cn_date,
            "2026年x月x日": cn_date,
            "2026年X月X日": cn_date,
            "XXX人天": context["eval_days"],
            "xx.xx%": reduction_val,
            "XXXXX%": reduction_val,
            "XXXX%": reduction_val,
            "XXX%": reduction_val,
            "xxxxx%": reduction_val,
            "xxxx%": reduction_val,
            "xxx%": reduction_val,
        }

        # 动态补充项目名称占位符 (XX, XXX, XXXX, ... XXXXXXXX)
        # 注意：不再添加纯X占位符，统一由Section C的正则处理
        # for i in range(10, 1, -1):
        #     # 大写 X
        #     ph_upper = "X" * i
        #     if ph_upper not in text_replacements:
        #         text_replacements[ph_upper] = context["project_name"]
        #     # 小写 x
        #     ph_lower = "x" * i
        #     if ph_lower not in text_replacements:
        #         text_replacements[ph_lower] = context["project_name"]
        #     # 混合/带数字
        #     if i <= 5:
        #         text_replacements[f"XX{i}"] = context["project_name"]

        # 增加一些特殊变体
        # text_replacements["Xx"] = context["project_name"]
        # text_replacements["Xxx"] = context["project_name"]

        # 定义标签追加 (如 项目名称：后追加)
        labels = {
            "项目名称：": context["project_name"],
            "项目名称:": context["project_name"],
            "项目名称": context["project_name"],
            "项目编号：": context["project_id"],
            "项目编号:": context["project_id"],
            "项目编号": context["project_id"],
            "送审单位：": context["submission_unit"],
            "送审单位:": context["submission_unit"],
            "送审单位": context["submission_unit"],
            "送审人：": context["submitter"],
            "送审人:": context["submitter"],
            "送审人": context["submitter"],
            "项目负责人：": context["submitter"],
            "项目负责人:": context["submitter"],
            "项目负责人": context["submitter"],
            "送审时间：": context["submission_time"],
            "送审时间:": context["submission_time"],
            "送审时间": context["submission_time"],
            "日期：": context["submission_time"],
            "日期": context["submission_time"],
        }

        # 定义一个统一的处理器
        def process_paragraph(p, extra_context=""):
            if hasattr(p, "_processed_by_agent"):
                return

            # A. 处理高亮色块
            ReceiptProcessor._smart_replace_highlights(p, data_bundle, extra_context)

            # B. 处理特定文本替换 (日期、百分比等) - 移动到 X 模式之前，并支持跨 Run 替换
            p_text = p.text
            sorted_replaces = sorted(
                text_replacements.items(), key=lambda x: len(x[0]), reverse=True
            )
            replaced_any = False
            for old, new in sorted_replaces:
                if old in p_text:
                    # 尝试直接在 Run 中替换
                    found_in_run = False
                    for run in p.runs:
                        if old in run.text:
                            run.text = run.text.replace(old, str(new))
                            found_in_run = True

                    # 如果 Run 中没找到（说明被碎裂了），则全段替换并尝试保留样式
                    if not found_in_run:
                        p_text = p_text.replace(old, str(new))
                        replaced_any = True

            if replaced_any:
                # 获取原始样式并重组段落
                sample_run = p.runs[0] if p.runs else None
                for r in p.runs:
                    if not sample_run and r.text.strip():
                        sample_run = r
                    r.text = ""
                if p.runs:
                    p.runs[0].text = p_text
                    target_run = p.runs[0]
                else:
                    target_run = p.add_run(p_text)

                # 恢复样板样式
                if sample_run:
                    if sample_run.font.name:
                        target_run.font.name = sample_run.font.name
                        try:
                            target_run._element.rPr.rFonts.set(
                                qn("w:eastAsia"), sample_run.font.name
                            )
                        except:
                            pass
                    if sample_run.font.size:
                        target_run.font.size = sample_run.font.size
                    if sample_run.font.bold is not None:
                        target_run.font.bold = sample_run.font.bold
                    if sample_run.font.color and sample_run.font.color.rgb:
                        try:
                            target_run.font.color.rgb = sample_run.font.color.rgb
                        except:
                            pass

                # 标记已处理并退出，避免后续 X 替换和标签追加再次处理
                p._processed_by_agent = True
                return

            # C. 统一占位符替换 (使用正则处理所有 X/x 连缀)
            # 优化正则：排除后面紧跟 % 的情况，避免误杀比例
            x_pattern = re.compile(r"[Xx]+(?:[.,、\s]*[Xx]+)*(?:\.\d+)?(?:\d+)?(?![%])")

            p_text = p.text
            # 检查是否包含 x 或 X
            if any(c in p_text for c in "xX") and x_pattern.search(p_text):
                # 寻找样板 Run (用于继承样式)
                sample_run = p.runs[0] if p.runs else None
                for r in p.runs:
                    if any(c in r.text for c in "xX"):
                        sample_run = r
                        break

                # 执行正则替换：将所有的 X 连缀块替换为项目名称
                def replace_x_pattern(match):
                    return context["project_name"]

                new_text = x_pattern.sub(replace_x_pattern, p_text)

                if new_text != p_text:
                    # 获取原始样式
                    fname = sample_run.font.name if sample_run else None
                    fsize = sample_run.font.size if sample_run else None
                    fbold = sample_run.font.bold if sample_run else None
                    fcolor = (
                        sample_run.font.color.rgb
                        if (sample_run and sample_run.font.color)
                        else None
                    )

                    # 重组段落 (防止碎裂同时彻底清除旧 placeholder)
                    for r in p.runs:
                        r.text = ""
                    if p.runs:
                        p.runs[0].text = new_text
                        target_run = p.runs[0]
                    else:
                        target_run = p.add_run(new_text)

                    # 恢复样式
                    if fname:
                        target_run.font.name = fname
                        try:
                            target_run._element.rPr.rFonts.set(qn("w:eastAsia"), fname)
                        except:
                            pass
                    if fsize:
                        target_run.font.size = fsize
                    if fbold is not None:
                        target_run.font.bold = fbold
                    if fcolor:
                        try:
                            target_run.font.color.rgb = fcolor
                        except:
                            pass

                    # 只要处理了 X 占位符，就标记已完成，防止后续标签逻辑再次追加
                    p._processed_by_agent = True
                    return

            # D. 处理标签追加 (项目名称：[空/弱值])
            p_text_clean = "".join(p.text.split()).replace(" ", "")
            for label, val in labels.items():
                label_clean = "".join(label.split()).replace(" ", "")
                if p_text_clean.startswith(label_clean):
                    # 检查标签后面是否是“弱值”
                    suffix = p_text_clean[len(label_clean) :].strip()
                    weak_base = [
                        "",
                        "N/A",
                        "[]",
                        "【】",
                        "：",
                        ":",
                        "1",
                        "x",
                        "X",
                        "unknown",
                    ]
                    is_weak = (
                        not suffix
                        or suffix in weak_base
                        or all(c in "_.-·— " for c in suffix)
                    )

                    if is_weak:
                        if val and str(val) != "N/A":
                            # 清理弱值
                            if suffix and suffix not in ["：", ":"]:
                                for r in p.runs:
                                    if suffix in r.text:
                                        r.text = r.text.replace(suffix, "", 1)

                            if not p.text.strip().endswith(("：", ":")):
                                p.add_run("：")

                            new_run = p.add_run(f"{val}")
                            # 继承该段落之前的样式
                            ref = p.runs[0] if p.runs else None
                            if ref:
                                new_run.font.name = ref.font.name
                                new_run.font.size = ref.font.size
                                new_run.font.bold = ref.font.bold
                                try:
                                    new_run._element.rPr.rFonts.set(
                                        qn("w:eastAsia"), ref.font.name
                                    )
                                except:
                                    pass
                        p._processed_by_agent = True
                        return

            p._processed_by_agent = True

        # 3. 遍历所有段落
        for p in doc.paragraphs:
            process_paragraph(p)

        # 4. 遍历所有表格 (使用列头信息处理高亮块和普通文本)
        for table in doc.tables:
            # 记录表头
            headers = []
            if len(table.rows) > 0:
                for cell in table.rows[0].cells:
                    headers.append(cell.text.strip())

            for row in table.rows:
                for idx, cell in enumerate(row.cells):
                    # 获取列上下文 (表头)
                    col_header = headers[idx] if idx < len(headers) else ""
                    for p in cell.paragraphs:
                        process_paragraph(p, col_header)

    @staticmethod
    def _smart_replace_highlights(p, data_bundle, col_header=""):
        """
        根据"标点隔离+权重路径+全句环视"原则精准填报。
        改进版：将连续的高亮 runs 识别为一个高亮块，每个块只填充一次。
        """
        runs = p.runs
        header = col_header.strip().replace(" ", "")
        major_seps = ["。", "；", "！", "？", "\n", "\r", "\t"]
        minor_seps = ["，", "：", ":", " ", ","]
        all_seps = major_seps + minor_seps

        # 第一步：识别所有高亮块（连续的高亮 runs 视为一个块）
        highlight_blocks = []
        current_block = []

        for i, run in enumerate(runs):
            # 兼容性修复：某些 docx 文件的高亮值为 'none' 会导致 python-docx 抛出 ValueError
            try:
                has_highlight = run.font.highlight_color is not None
            except (ValueError, Exception):
                has_highlight = False

            if has_highlight:
                current_block.append(i)
            else:
                if current_block:
                    highlight_blocks.append(current_block)
                    current_block = []

        # 最后一个块
        if current_block:
            highlight_blocks.append(current_block)

        # 第二步：对每个高亮块，只填充一次
        for block_indices in highlight_blocks:
            if not block_indices:
                continue

            # 使用第一个 run 的索引进行上下文分析
            i = block_indices[0]

            try:
                # 1. 动态提取本句内的短上下文 (隔离)
                lc = ""  # 左侧短语境
                for j in range(i - 1, -1, -1):
                    txt = runs[j].text
                    has_sep = False
                    for s in all_seps:
                        if s in txt:
                            lc = txt[txt.rfind(s) + 1 :] + lc
                            has_sep = True
                            break
                    if has_sep:
                        break
                    lc = txt + lc

                # 右侧短语境 - 从高亮块结束后开始
                rc = ""
                last_idx = block_indices[-1]
                for j in range(last_idx + 1, len(runs)):
                    txt = runs[j].text
                    has_sep = False
                    for s in all_seps:
                        if s in txt:
                            rc += txt[: txt.find(s)]
                            has_sep = True
                            break
                    if has_sep:
                        break
                    rc += txt

                lc_clean = lc.strip().replace(" ", "")
                rc_clean = rc.strip().replace(" ", "")

                # 合并整个高亮块的文本用于判定
                own_text = (
                    "".join([runs[idx].text for idx in block_indices])
                    .strip()
                    .replace(" ", "")
                )

                # 2. 维度判定 (Dimension) - 优先级：右侧 > 自身 > 左侧 > 表头
                dim = "count"
                if (
                    "%" in rc_clean
                    or "%" in own_text
                    or "占比" in lc_clean
                    or "比例" in lc_clean
                    or "比例" in header
                ):
                    dim = "ratio"
                elif (
                    "人天" in rc_clean
                    or "工作量" in rc_clean
                    or "人天" in lc_clean
                    or "工作量" in lc_clean
                    or "人天" in header
                ):
                    dim = "days"
                elif "个" in rc_clean or "功能点" in rc_clean or "功能点" in header:
                    dim = "count"

                # 3. 对象判定 (Object Scoring)
                obj_map = {
                    "核减": "reduction",
                    "新增": "new",
                    "增加": "new",
                    "复用": "reuse",
                    "优化": "reuse",
                    "利旧": "legacy",
                    "送审": "submission",
                    "核定": "eval",
                    "核准": "eval",
                    "评估": "eval",
                    "评定": "eval",
                }

                obj = None
                # A. 优先判定后置语义 (rc 中的"计新增"等)
                if "计新增" in rc_clean:
                    obj = "new"
                    dim = "ratio"
                elif "计复用" in rc_clean:
                    obj = "reuse"
                    dim = "ratio"
                elif "计利旧" in rc_clean:
                    obj = "legacy"
                    dim = "ratio"

                if obj is None:
                    # B. 寻找离色块最近的操作对象 (全句追溯，不停于逗号)
                    # 我们向上回溯直到大标点（句号/分号）
                    long_lc = ""
                    for j in range(i - 1, -1, -1):
                        txt = runs[j].text
                        if any(s in txt for s in major_seps):
                            parts = [txt.rfind(s) for s in major_seps if s in txt]
                            long_lc = txt[max(parts) + 1 :] + long_lc
                            break
                        long_lc = txt + long_lc

                    long_lc = long_lc.replace(" ", "")
                    last_pos = -1
                    for kw, target_obj in obj_map.items():
                        p_idx = long_lc.rfind(kw)
                        if p_idx > last_pos:
                            last_pos = p_idx
                            obj = target_obj

                    # C. 兜底看表头
                    if obj is None:
                        for kw, target_obj in obj_map.items():
                            if kw in header:
                                obj = target_obj
                                break

                # 4. 映射数据
                val = None
                if obj == "new":
                    val = (
                        data_bundle["new_ratio"]
                        if dim == "ratio"
                        else data_bundle["new_fp"]
                    )
                elif obj == "reuse":
                    val = (
                        data_bundle["reuse_ratio"]
                        if dim == "ratio"
                        else data_bundle["reuse_fp"]
                    )
                elif obj == "legacy":
                    val = (
                        data_bundle["legacy_ratio"]
                        if dim == "ratio"
                        else data_bundle["legacy_fp"]
                    )
                elif obj == "reduction":
                    val = data_bundle["reduction_ratio"]
                elif obj == "submission":
                    val = (
                        data_bundle["submission_days"]
                        if dim == "days"
                        else data_bundle["total_fp"]
                    )
                elif obj == "eval":
                    val = (
                        data_bundle["eval_days"]
                        if dim == "days"
                        else data_bundle["total_fp"]
                    )

                # 5. 执行替换 - 只在第一个 run 中填充，其他 run 清空
                if val is not None:
                    val_str = str(val)
                    # 清理重复的百分号
                    if "%" in val_str and ("%" in rc_clean or own_text.endswith("%")):
                        val_str = val_str.replace("%", "")

                    # 第一个 run 填充值
                    runs[block_indices[0]].text = val_str
                    runs[block_indices[0]].font.highlight_color = None

                    # 其他 run 清空（删除多余的 X 占位符）
                    for idx in block_indices[1:]:
                        runs[idx].text = ""
                        runs[idx].font.highlight_color = None
            except Exception:
                pass


def datetime_str():
    return datetime.now().strftime("%Y%m%d_%H%M%S")
