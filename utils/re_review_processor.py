import os
import re
import openpyxl
from openpyxl.utils import get_column_letter
from PySide6.QtCore import QThread, Signal


class ReReviewWorker(QThread):
    """异步处理重评任务的线程"""

    progress = Signal(int)
    finished = Signal(tuple)
    error = Signal(str)

    def __init__(self, excel1_path, excel2_path, output_dir, project_name=None):
        super().__init__()
        self.excel1_path = excel1_path
        self.excel2_path = excel2_path
        self.output_dir = output_dir
        self.project_name = project_name or "ReReview"

    def run(self):
        from utils.runtime_logger import RuntimeLogger
        
        # 设置项目名称用于日志
        RuntimeLogger.set_project(self.project_name)
        RuntimeLogger.log("开始重评处理...")
        
        try:
            p1, p2 = ReReviewProcessor.process_re_review(
                self.excel1_path, self.excel2_path, self.output_dir, self.progress.emit
            )
            self.finished.emit((p1, p2))
        except Exception as e:
            import traceback

            traceback.print_exc()
            self.error.emit(str(e))


class ReReviewProcessor:
    """重评自动标注处理器"""

    @staticmethod
    def find_column(ws, names, default_col, search_rows=10):
        # ... existing code ...
        """在工作表中搜索包含特定名称的列"""
        if isinstance(names, str):
            names = [names]

        for r in range(1, search_rows + 1):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(row=r, column=c).value
                if val:
                    val_str = str(val)
                    for name in names:
                        if name in val_str:
                            return c
        return default_col

    @staticmethod
    def transform_remark(old_remark, new_id):
        """智能生成新备注：[新标号] [剩余文字]"""
        if old_remark is None:
            return ""

        txt = str(old_remark).strip()
        if not txt:
            return ""

        # N 或 O 保持不变
        if txt.upper() in ["N", "O"]:
            return txt.upper()

        # 匹配第一个数字序列
        match = re.search(r"\d+", txt)
        if match:
            start, end = match.span()
            # 替换该数字序列为新 ID
            return txt[:start] + str(new_id) + txt[end:]
        else:
            # 无数字说明是纯文字，按 MD 逻辑： T5 & " " & txt
            return f"{new_id} {txt}"

    @staticmethod
    def process_re_review(excel1_path, excel2_path, output_dir, progress_callback=None):
        """
        处理重评逻辑：
        excel1_path: 评估报告 (需要添加公式)
        excel2_path: 重评回单 (需要同步结果)
        output_dir: 保存目录
        """
        if progress_callback:
            progress_callback(5)

        # 1. 加载工作簿
        # wb1: 评估报告 (Excel1) - 需要加载两次：一次用于公式保存，一次用于数据读取
        wb1_formula = openpyxl.load_workbook(excel1_path)
        wb1_val = openpyxl.load_workbook(excel1_path, data_only=True)

        if progress_callback:
            progress_callback(15)

        # wb2: 重评回单 (Excel2) - 用于读取数据和保存结果
        wb2_val = openpyxl.load_workbook(excel2_path, data_only=True)
        wb2_final = openpyxl.load_workbook(excel2_path)

        if progress_callback:
            progress_callback(25)

        def get_split_sheet(wb):
            names = ["功能点拆分表", "2、功能点拆分表", "拆分表"]
            for name in names:
                if name in wb.sheetnames:
                    return wb[name]
            for name in wb.sheetnames:
                if "功能点拆分" in name:
                    return wb[name]
            raise ValueError(f"无法在文件中找到功能点拆分表工作表")

        ws1_formula = get_split_sheet(wb1_formula)
        ws1_val = get_split_sheet(wb1_val)
        ws2_val = get_split_sheet(wb2_val)
        ws2_final = get_split_sheet(wb2_final)

        # 获取列索引 (支持多种可能的列表头名称)
        desc_col1 = ReReviewProcessor.find_column(
            ws1_val, ["子过程描述", "功能点拆分描述", "子过程"], 8
        )
        func_col1 = ReReviewProcessor.find_column(
            ws1_val, ["功能过程", "功能简述", "功能点"], 6
        )

        desc_col2 = ReReviewProcessor.find_column(
            ws2_val, ["子过程描述", "功能点拆分描述", "子过程"], 8
        )
        func_col2 = ReReviewProcessor.find_column(
            ws2_val, ["功能过程", "功能简述", "功能点"], 6
        )
        remark_col2 = ReReviewProcessor.find_column(ws2_val, ["备注", "原备注"], 15)

        start_row = 5

        # ========== 建立数据映射用于结果计算 ==========
        # 使用 (功能过程, 子过程描述) 作为联合主键
        excel1_idx_to_data = {}  # {row_idx: (func, desc)}
        excel1_key_to_idx = {}  # {(func, desc): row_idx}

        for r in range(start_row, ws1_val.max_row + 1):
            d_val = ws1_val.cell(row=r, column=desc_col1).value
            f_val = ws1_val.cell(row=r, column=func_col1).value
            if d_val:
                desc = str(d_val).strip()
                func = str(f_val).strip() if f_val else ""
                excel1_idx_to_data[r] = (func, desc)
                excel1_key_to_idx[(func, desc)] = r

        excel2_key_to_idx = {}  # {(func, desc): row_idx}
        for r in range(start_row, ws2_val.max_row + 1):
            d_val = ws2_val.cell(row=r, column=desc_col2).value
            f_val = ws2_val.cell(row=r, column=func_col2).value
            if d_val or f_val:
                desc = str(d_val).strip() if d_val else ""
                func = str(f_val).strip() if f_val else ""
                excel2_key_to_idx[(func, desc)] = r

        if progress_callback:
            progress_callback(35)

        # ========== 步骤 1: 在 Excel1 中创建辅助工作表 Sheet1 ==========
        if "Sheet1" in wb1_formula.sheetnames:
            del wb1_formula["Sheet1"]
        ws_sheet1 = wb1_formula.create_sheet("Sheet1")

        # 匹配截图中的表头结构
        # A: Excel1行号 | B: Excel1映射主键 | C: Excel2映射主键 | D: Excel2行号
        ws_sheet1.cell(row=1, column=1).value = 1
        ws_sheet1.cell(row=1, column=4).value = 1
        ws_sheet1.cell(row=3, column=2).value = "Excel1组合键"
        ws_sheet1.cell(row=3, column=3).value = "Excel2组合键"

        # 填充 Excel1 数据到 Sheet1
        for r, data in excel1_idx_to_data.items():
            ws_sheet1.cell(row=r, column=1).value = r
            # 使用公式拼接主表的 功能过程 和 子过程描述
            f_ref = f"'{ws1_formula.title}'!{get_column_letter(func_col1)}{r}"
            d_ref = f"'{ws1_formula.title}'!{get_column_letter(desc_col1)}{r}"
            ws_sheet1.cell(row=r, column=2).value = f"=TRIM({f_ref}) & TRIM({d_ref})"

        # 填充 Excel2 数据到 Sheet1 (C, D列)
        for key, r in excel2_key_to_idx.items():
            # 组合键 (Python 直接拼接)
            ws_sheet1.cell(row=r, column=3).value = key[0] + key[1]
            ws_sheet1.cell(row=r, column=4).value = r

        # ========== 步骤 2 & 3: 在 Excel1 中添加公式 ==========
        col_q = 17  # Q (17) 匹配检查
        col_r = 18  # R (18) 备注对应描述
        col_s = 19  # S (19) 备注对应标号 (新增处理列)
        col_o = 15  # O (15) 备注
        col_t = 20  # T (20) 新标号

        ws1_formula.cell(row=start_row - 1, column=col_q).value = "匹配检查"
        ws1_formula.cell(row=start_row - 1, column=col_r).value = "备注对应描述"
        ws1_formula.cell(row=start_row - 1, column=col_s).value = "备注处理(ID)"
        ws1_formula.cell(row=start_row - 1, column=col_t).value = "新标号"

        for r in range(start_row, ws1_formula.max_row + 1):
            col_o_ref = f"{get_column_letter(col_o)}{r}"
            col_r_ref = f"{get_column_letter(col_r)}{r}"
            col_s_ref = f"{get_column_letter(col_s)}{r}"

            # S列: 备注处理 (提取其中的 ID)
            # 使用 Excel 技巧提取字符串中的第一个数字序列
            # 如果是 n,123 或 123,文字，该公式能较好地提取出 123
            ws1_formula.cell(row=r, column=col_s).value = (
                f'=IFERROR(LOOKUP(9.9E+307,--LEFT(MID({col_o_ref},MIN(FIND({{0,1,2,3,4,5,6,7,8,9}},{col_o_ref}&"0123456789")),99),ROW($1:$99))),"")'
            )

            # Q列: 匹配检查
            curr_key_formula = f"TRIM({get_column_letter(func_col1)}{r}) & TRIM({get_column_letter(desc_col1)}{r})"
            ws1_formula.cell(row=r, column=col_q).value = (
                f'=IFERROR(VLOOKUP({curr_key_formula}, Sheet1!$C:$C, 1, 0), "")'
            )

            # R列: 备注对应描述 (现在查找的是 S 列处理后的 ID)
            ws1_formula.cell(row=r, column=col_r).value = (
                f'=IFERROR(VLOOKUP({col_s_ref}, Sheet1!$A:$B, 2, 0), "")'
            )

            # T列: 新标号
            formula_t = (
                f'=IF({col_o_ref}="", "", '
                f'IF(OR(UPPER(TRIM({col_o_ref}))="N", UPPER(TRIM({col_o_ref}))="O"), {col_o_ref}, '
                f'IFERROR(VLOOKUP({col_r_ref}, Sheet1!$C:$D, 2, 0), "")))'
            )

            ws1_formula.cell(row=r, column=col_t).value = formula_t

        if progress_callback:
            progress_callback(60)

        # ========== 步骤 4: 处理 Excel2 (重评回单) 的结果赋值到 P 列 ==========
        col_p = 16  # P 列 (结果存放列)
        ws2_final.cell(row=start_row - 1, column=col_p).value = "新标号(P)"

        for r in range(start_row, ws2_val.max_row + 1):
            desc2 = ws2_val.cell(row=r, column=desc_col2).value
            func2 = ws2_val.cell(row=r, column=func_col2).value

            if not desc2 and not func2:
                ws2_final.cell(row=r, column=col_p).value = ""
                continue

            desc2_str = str(desc2).strip() if desc2 else ""
            func2_str = str(func2).strip() if func2 else ""
            key2 = (func2_str, desc2_str)

            # 保证excel2的子功能描述和功能过程都和excel1一样，才进行填充
            if key2 in excel1_key_to_idx:
                # 找到 excel1 中对应的行
                r1 = excel1_key_to_idx[key2]
                # 获取 excel1 中的备注 (作为同步源)
                remark1 = ws1_val.cell(row=r1, column=col_o).value

                if not remark1:
                    # 原来是空的，新的也应该是空
                    ws2_final.cell(row=r, column=col_p).value = ""
                elif str(remark1).strip().upper() in ["N", "O"]:
                    # n/N 都原样输出
                    ws2_final.cell(row=r, column=col_p).value = str(remark1).strip()
                else:
                    # 处理带标号的情况
                    txt = str(remark1).strip()
                    match = re.search(r"\d+", txt)
                    if match:
                        old_id_str = match.group()
                        try:
                            old_row_idx = int(old_id_str)
                            # 查询 excel1 中该标号指向的描述和功能
                            if old_row_idx in excel1_idx_to_data:
                                ref_key = excel1_idx_to_data[old_row_idx]
                                # 查询该描述+功能在当前 excel2 中的新行号
                                if ref_key in excel2_key_to_idx:
                                    new_row_idx = excel2_key_to_idx[ref_key]
                                    start, end = match.span()
                                    
                                    # 如果原始文本就是一个单纯的数字，则输出为数字类型
                                    if txt == old_id_str:
                                        ws2_final.cell(row=r, column=col_p).value = new_row_idx
                                    else:
                                        # 否则拼接字符串
                                        new_val = txt[:start] + str(new_row_idx) + txt[end:]
                                        ws2_final.cell(row=r, column=col_p).value = new_val
                                else:
                                    ws2_final.cell(row=r, column=col_p).value = txt
                            else:
                                ws2_final.cell(row=r, column=col_p).value = txt
                        except:
                            ws2_final.cell(row=r, column=col_p).value = txt
                    else:
                        ws2_final.cell(row=r, column=col_p).value = txt
            else:
                # 描述或功能不匹配，不填充
                ws2_final.cell(row=r, column=col_p).value = ""

        # ========== 保存文件 ==========
        if progress_callback:
            progress_callback(85)

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        filename1 = os.path.basename(excel1_path)
        filename2 = os.path.basename(excel2_path)

        path1 = os.path.join(output_dir, filename1)
        path2 = os.path.join(output_dir, filename2)

        wb1_formula.save(path1)
        wb2_final.save(path2)

        if progress_callback:
            progress_callback(100)

        wb1_formula.close()
        wb1_val.close()
        wb2_val.close()
        wb2_final.close()

        return path1, path2
