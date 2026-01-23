import pandas as pd
import os
import re
from datetime import datetime
from extend.matcher_config import MatcherConfig


class ReportGenerator:
    """自动化报表生成工具"""

    @staticmethod
    def _clean_project_name(name):
        """净化项目名，去除附件前缀、各种冗余后缀、版本号等"""
        # 1. 基础处理：手动移除扩展名
        name = os.path.basename(name).strip()
        for ext in [".docx", ".doc", ".xlsx", ".XLSX", ".DOCX"]:
            if name.lower().endswith(ext.lower()):
                name = name[: -len(ext)].strip()
                break

        # 2. 去掉开头的“附件X：”或“1.”等，匹配附件/数字+标点
        name = re.sub(r"^附件\s*\d+\s*[：:.\-\s]*\s*", "", name)
        # 匹配 1-3位数字 + 点 + 可选空格。1-3位是为了避开 2024. 这种年份开头
        name = re.sub(r"^\d{1,3}[\.．]\s*", "", name)
        name = name.strip()

        # 3. 循环清理末尾后缀
        while True:
            prev_name = name

            # (a) 明显的后缀词
            suffixes = [
                "产品需求说明书",
                "需求规格说明书",
                "需求规格书",
                "需求说明书",
                "规格说明书",
                "规格书",
                "功能点拆分表",
                "功能拆分表",
                "拆分表",
                "审计方案",
                "测试用例",
                "说明书",
                "文档",
                "需求",
            ]
            for s in suffixes:
                if name.endswith(s):
                    new_n = name[: -len(s)].strip()
                    if len(new_n) >= 2:
                        name = new_n

            # (b) 版本号日期以及 (1) (2) 这种重复标识
            name = re.sub(r"[vV]\d+(?:\.\d+)*\s*$", "", name)
            name = re.sub(r"(?:20)?\d{6,8}\s*$", "", name)
            name = re.sub(r"[\(（]\d+[\)）]\s*$", "", name)
            name = re.sub(r"\s*-\s*副本\s*$", "", name)

            # (c) 末尾连接词
            name = re.sub(
                r"(?:的项目|项目|的需求|的产品|产品|的|之)+$", "", name
            ).strip()

            # (d) 标点
            name = name.strip(" :：-－_|'\"().（）")

            if name == prev_name:
                break

        return name if name else "未知项目"

    @classmethod
    def generate_validation_report(cls, task_name, results, output_dir=None):
        """
        生成全量比对报表
        """
        try:
            if output_dir is None:
                config = MatcherConfig.load()
                output_dir = config.get("storage", {}).get("initial_review", ".")

            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)

            details = results.get("all_details", [])
            if not details:
                return None

            # 净化项目名
            clean_name = cls._clean_project_name(task_name)

            # 转换为 DataFrame
            df = pd.DataFrame(details)

            column_map = {
                "level": "层级",
                "chapter": "标准章节",
                "status": "判定结果",
                "tpl_sentence": "模板参考句 (对照)",
                "target_sentence": "上传文本句 (对照)",
                "score": "重合度",
            }
            # 过滤并重命名
            existing_cols = [c for c in df.columns if c in column_map]
            df = df[existing_cols].rename(columns=column_map)

            # 格式化百分比
            if "重合度" in df.columns:
                df["重合度"] = df["重合度"].apply(
                    lambda x: f"{x*100:.1f}%" if isinstance(x, (int, float)) else x
                )

            # 确定文件名: xxx项目模板校验评估报告.xlsx
            base_filename = f"{clean_name}模板校验评估报告.xlsx"
            target_path = os.path.join(output_dir, base_filename)

            # 防占用处理：如果文件已存在且无法写入，尝试生成副本
            final_path = target_path
            counter = 1
            while True:
                try:
                    # 尝试以写模式打开文件以检测占用
                    if os.path.exists(final_path):
                        with open(final_path, "a"):
                            pass
                    break  # 如果没报错，说明可以写入
                except (IOError, PermissionError):
                    # 文件被占用，尝试新名称
                    name, ext = os.path.splitext(base_filename)
                    final_path = os.path.join(output_dir, f"{name}({counter}){ext}")
                    counter += 1

            # 使用 ExcelWriter 保存，带一点样式
            with pd.ExcelWriter(final_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="校验详情")

                workbook = writer.book
                worksheet = writer.sheets["校验详情"]
                from openpyxl.styles import Alignment

                # 调整列宽并开启自动换行
                for i, col in enumerate(df.columns):
                    column_len = (
                        max(
                            (
                                df[col].astype(str).map(len).max()
                                if not df[col].empty
                                else 10
                            ),
                            len(col),
                        )
                        + 2
                    )
                    col_letter = chr(65 + i)
                    worksheet.column_dimensions[col_letter].width = min(column_len, 60)

                    # 为正文对照列开启自动换行
                    if "对照" in col:
                        for cell in worksheet[col_letter]:
                            cell.alignment = Alignment(
                                wrap_text=True, vertical="center"
                            )
                    else:
                        for cell in worksheet[col_letter]:
                            cell.alignment = Alignment(vertical="center")

            return final_path
        except Exception as e:
            print(f"生成报表失败: {e}")
            return None
