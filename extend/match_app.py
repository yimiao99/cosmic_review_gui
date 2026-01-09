import streamlit as st
from docx import Document
import pandas as pd
import re
from fuzzywuzzy import fuzz
from io import BytesIO


# --- 文本清洗函数 ---
def clean_text(text):
    """清洗文本：去除编号、括号内容、多余空格，并转为小写"""
    if not isinstance(text, str):
        return ""
    text = re.sub(r'^\d+(\.\d+)*\s*', '', text)  # 去除开头的编号
    text = re.sub(r'\(.*?\)', '', text)  # 去除括号内容
    text = text.strip()  # 去除首尾空格
    text = re.sub(r'\s+', ' ', text)  # 多个空格合并为一个
    text = text.lower()  # 转为小写
    return text


# --- 从Word文档中提取标题并清洗 ---
def extract_word_headings(word_file_buffer):
    """从Word文档中提取所有标题样式的段落"""
    doc = Document(word_file_buffer)
    original_headings = []
    cleaned_headings = []

    for para in doc.paragraphs:
        if para.style.name.startswith('Heading'):
            original_headings.append(para.text)
            cleaned_text = clean_text(para.text)
            if cleaned_text:
                cleaned_headings.append(cleaned_text)

    # 去重
    unique_cleaned_headings = list(set(cleaned_headings))

    # 创建清洗后到原始文本的映射
    cleaned_to_original_map = {}
    for original, cleaned in zip(original_headings, cleaned_headings):
        if cleaned not in cleaned_to_original_map:
            cleaned_to_original_map[cleaned] = original

    return unique_cleaned_headings, cleaned_to_original_map


# --- 从Excel文件中读取任务名称并清洗 ---
def read_excel_tasks(excel_file_buffer, sheet_name, task_col):
    """从Excel文件中读取指定工作表和列的任务名称"""
    st.info(f"DEBUG: read_excel_tasks 函数内部尝试读取Excel工作表: '{sheet_name}'")

    if not sheet_name:
        raise ValueError("Excel工作表名称不能为空。")

    try:
        df = pd.read_excel(excel_file_buffer, sheet_name=sheet_name)
    except ValueError as e:
        if f"Worksheet named '{sheet_name}' not found" in str(e) or f"No sheet named '{sheet_name}'" in str(e):
            raise ValueError(f"Excel文件中未找到名为 '{sheet_name}' 的工作表。请检查名称是否正确。")
        else:
            raise e

    if task_col not in df.columns:
        raise ValueError(f"Excel文件 '{sheet_name}' 工作表中未找到列 '{task_col}'。请检查列名是否正确。")

    original_tasks = []
    cleaned_tasks = []

    for task_name in df[task_col].dropna().astype(str):
        original_tasks.append(task_name)
        cleaned_text = clean_text(task_name)
        if cleaned_text:
            cleaned_tasks.append(cleaned_text)

    # 去重
    unique_cleaned_tasks = list(set(cleaned_tasks))

    # 创建清洗后到原始文本的映射
    cleaned_to_original_map = {}
    for original, cleaned in zip(original_tasks, cleaned_tasks):
        if cleaned not in cleaned_to_original_map:
            cleaned_to_original_map[cleaned] = original

    return unique_cleaned_tasks, cleaned_to_original_map


# --- 比较两个列表并生成报告 ---
def compare_lists(word_list_cleaned, excel_list_cleaned,
                  original_word_map, original_excel_map, threshold=90):
    """比较Word标题和Excel任务，生成匹配报告"""
    results = []

    # 检查Word中的每个条目在Excel中是否存在
    for word_item_cleaned in word_list_cleaned:
        match_found = False
        match_type = "未找到匹配项"
        matched_excel_original = ""

        # 首先尝试精确匹配
        if word_item_cleaned in excel_list_cleaned:
            match_found = True
            match_type = "精确匹配"
            matched_excel_original = original_excel_map.get(word_item_cleaned, "")
        else:
            # 尝试模糊匹配
            best_score = 0
            best_match_excel_cleaned = None
            for excel_item_cleaned in excel_list_cleaned:
                score = fuzz.ratio(word_item_cleaned, excel_item_cleaned)
                if score > best_score:
                    best_score = score
                    best_match_excel_cleaned = excel_item_cleaned

            if best_score >= threshold:
                match_found = True
                match_type = f"模糊匹配 (相似度: {best_score}%)"
                matched_excel_original = original_excel_map.get(best_match_excel_cleaned, "")

        results.append({
            '来源': 'Word目录',
            '原始条目': original_word_map.get(word_item_cleaned, word_item_cleaned),
            '清洗后条目': word_item_cleaned,
            '在对方文件中找到': '是' if match_found else '否',
            '匹配方式': match_type,
            '匹配到的对方原始条目': matched_excel_original
        })

    # 检查Excel中独有的条目（Word中不存在）
    for excel_item_cleaned in excel_list_cleaned:
        match_found = False

        # 检查精确匹配
        if excel_item_cleaned in word_list_cleaned:
            match_found = True
        else:
            # 检查模糊匹配
            best_score = 0
            for word_item_cleaned in word_list_cleaned:
                score = fuzz.ratio(excel_item_cleaned, word_item_cleaned)
                if score > best_score:
                    best_score = score

            if best_score >= threshold:
                match_found = True

        # 只添加Word中不存在的Excel条目
        if not match_found:
            results.append({
                '来源': 'Excel任务',
                '原始条目': original_excel_map.get(excel_item_cleaned, excel_item_cleaned),
                '清洗后条目': excel_item_cleaned,
                '在对方文件中找到': '否',
                '匹配方式': 'Word目录中不存在',
                '匹配到的对方原始条目': ''
            })

    return pd.DataFrame(results)


# --- Streamlit 应用界面 ---
def main():
    st.set_page_config(page_title="文档匹配工具", layout="wide")
    st.title("📄 Word目录与Excel任务匹配工具")
    st.markdown("上传Word文档和Excel文件，比对Word目录条目与Excel功能过程是否匹配。")

    st.sidebar.header("文件上传")
    word_file = st.sidebar.file_uploader("上传Word文档 (.docx)", type=["docx"], key="word_uploader")
    excel_file = st.sidebar.file_uploader("上传Excel文件 (.xlsx)", type=["xlsx"], key="excel_uploader")

    st.sidebar.header("配置选项")

    # 直接使用控件返回值，不依赖 session_state
    excel_sheet_name = st.sidebar.text_input(
        "Excel工作表名称",
        value="Sheet1"
    )
    excel_task_column = st.sidebar.text_input(
        "Excel任务列名",
        value="功能过程"
    )
    fuzzy_threshold = st.sidebar.slider(
        "模糊匹配相似度阈值 (%)", 0, 100, 85
    )

    if word_file and excel_file:
        st.success("文件已上传。点击 '开始比对' 进行处理。")
        if st.button("开始比对"):
            with st.spinner("正在处理，请稍候..."):
                try:
                    # 重置文件指针到开头（关键修复）
                    word_file.seek(0)
                    excel_file.seek(0)

                    st.info(f"DEBUG: 实际使用的Excel工作表名称: '{excel_sheet_name}'")
                    st.info(f"DEBUG: 实际使用的Excel任务列名: '{excel_task_column}'")

                    # 处理Word文档
                    word_file_buffer = BytesIO(word_file.read())
                    word_headings_cleaned, original_word_map = extract_word_headings(word_file_buffer)
                    st.info(f"Word文档中提取并清洗到 {len(word_headings_cleaned)} 个标题。")

                    # 处理Excel文件
                    excel_file_buffer = BytesIO(excel_file.read())
                    excel_tasks_cleaned, original_excel_map = read_excel_tasks(
                        excel_file_buffer,
                        excel_sheet_name,
                        excel_task_column
                    )
                    st.info(f"Excel文件中提取并清洗到 {len(excel_tasks_cleaned)} 个任务。")

                    # 生成比对报告
                    comparison_report = compare_lists(
                        word_headings_cleaned,
                        excel_tasks_cleaned,
                        original_word_map,
                        original_excel_map,
                        fuzzy_threshold
                    )

                    st.success("比对完成！")
                    st.subheader("比对报告")
                    st.dataframe(comparison_report, use_container_width=True)

                    # 生成Excel下载
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        comparison_report.to_excel(writer, index=False, sheet_name='匹配报告')
                    output.seek(0)

                    st.download_button(
                        label="📥 下载比对报告 (Excel)",
                        data=output,
                        file_name="匹配报告_纯匹配.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except ValueError as e:
                    st.error(f"❌ 配置错误：{e}")
                except Exception as e:
                    st.error(f"❌ 处理过程中发生错误：{e}")
                    st.exception(e)
    else:
        st.info("请在左侧边栏上传Word文档和Excel文件。")


if __name__ == "__main__":
    main()