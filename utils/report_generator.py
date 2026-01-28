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
        生成全量比对报表，包含所有校验步骤的汇总。
        """
        try:
            if output_dir is None:
                config = MatcherConfig.load()
                output_dir = config.get("storage", {}).get("initial_review", ".")

            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)

            # 净化项目名
            clean_name = cls._clean_project_name(task_name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_filename = f"{clean_name}_{timestamp}_全量校验报告.xlsx"
            final_path = os.path.join(output_dir, base_filename)

            # 准备各步骤数据
            with pd.ExcelWriter(final_path, engine="openpyxl") as writer:
                # 1. 概览页
                summary_data = []
                for node_key, node_name in [
                    ("tpl_res", "模板合规性"),
                    ("null_res", "空值校验"),
                    ("ratio_res", "比例校验"),
                    ("factor_check", "因子提取"),
                    ("hierarchy_res", "层级匹配"),
                    ("process_res", "功能过程"),
                    ("move_res", "数据移动"),
                ]:
                    res = results.get(node_key, {})
                    status = "✅ 通过" if res.get("is_valid", True) else "❌ 异常"
                    if res.get("skipped"): status = "⏩ 跳过"
                    
                    summary_data.append({
                        "校验节点": node_name,
                        "状态": status,
                        "详情": str(res.get("statistics", res.get("status", "")))
                    })
                pd.DataFrame(summary_data).to_excel(writer, sheet_name="📊 核查概览", index=False)

                # 2. Node 1: 模板详情
                tpl_details = results.get("all_details", [])
                if tpl_details:
                    df_tpl = pd.DataFrame(tpl_details)
                    column_map = {
                        "level": "层级", "chapter": "标准章节", "status": "判定结果",
                        "tpl_sentence": "模板句", "target_sentence": "匹配句", "score": "相似度"
                    }
                    df_tpl = df_tpl[[c for c in df_tpl.columns if c in column_map]].rename(columns=column_map)
                    df_tpl.to_excel(writer, sheet_name="1.模板校验详情", index=False)

                # 3. Node 5: 层级详情
                h_res = results.get("hierarchy_res", {})
                h_all = (h_res.get("exact_matched", []) + h_res.get("fuzzy_matched", []) + 
                        h_res.get("hierarchy_mismatched", []) + h_res.get("not_found_in_word", []))
                if h_all:
                    pd.DataFrame(h_all).to_excel(writer, sheet_name="5.层级匹配详情", index=False)

                # 4. Node 6: 功能过程详情
                p_res = results.get("process_res", {})
                p_all = (p_res.get("exact_matched", []) + p_res.get("fuzzy_matched", []) + 
                        p_res.get("not_found_in_word", []))
                if p_all:
                    df_p = pd.DataFrame(p_all)
                    # 确保关键列存在且排在前面
                    p_cols = ["Excel一级模块", "Excel二级模块", "Excel三级模块", "Excel功能点", "Word匹配项", "相似度", "位置", "匹配状态", "简略描述"]
                    existing_p_cols = [c for c in p_cols if c in df_p.columns]
                    # 加入其他可能存在的列
                    other_cols = [c for c in df_p.columns if c not in p_cols]
                    df_p = df_p[existing_p_cols + other_cols]
                    df_p.to_excel(writer, sheet_name="6.功能过程详情", index=False)

                # 5. Node 7: 数据移动详情
                m_res = results.get("move_res", {})
                m_details = m_res.get("details", [])
                if m_details:
                    pd.DataFrame(m_details).to_excel(writer, sheet_name="7.数据移动详情", index=False)

            return final_path
        except Exception as e:
            print(f"生成综合报表失败: {e}")
            import traceback
            traceback.print_exc()
            return None

        except Exception as e:
            print(f"生成报表失败: {e}")
            return None
