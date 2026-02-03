# extend/word_outline_extractor_debug.py
"""
Word 大纲提取器（增强版 + 详细调试日志）
- 严格使用 OutlineLevel 1-9 过滤（纯净大纲）
- 同时支持自动编号（ListFormat）和手动编号（正则提取）
- 添加详细的日志追踪，显示Word文件打开的每个步骤
"""
import win32com.client
import pythoncom
import os
import re  # 新增：用于手动编号提取
from typing import List, Tuple
import time  # 新增：用于性能监测


def extract_word_outline_with_detailed_log(doc_path: str, max_level: int = 9) -> List[Tuple[str, str, int]]:
    """
    稳定提取 Word 大纲（OutlineLevel 1~9），支持自动/手动编号 + 详细日志

        返回: [(编号, 标题, 展示级别), ...]
        - 过滤口径：仍使用 Word 的 OutlineLevel 1~9 来筛选"纯净大纲"段落
        - 展示口径：为了让目录树呈现为 4.6 / 4.6.1 / 4.6.1.1 ... 的缩进结构，
            对于形如 4.6.1.1 的数字编号，展示级别按"点的个数"计算：
            4.6 -> 1，4.6.1 -> 2，4.6.1.1 -> 3 ...（上限 9）
    """
    pythoncom.CoInitialize()
    word = None
    doc = None
    start_time = time.time()
    
    try:
        doc_path = os.path.abspath(doc_path)
        
        # ====== 第一阶段：文件检查 ======
        print(f"\n{'='*80}")
        print(f"[Word 大纲提取] 阶段 0: 文件检查")
        print(f"{'='*80}")
        print(f"📁 目标文件路径: {doc_path}")
        
        if not os.path.exists(doc_path):
            raise FileNotFoundError(f"文件不存在: {doc_path}")
        
        file_size = os.path.getsize(doc_path)
        file_ext = os.path.splitext(doc_path)[1].lower()
        print(f"✓ 文件存在")
        print(f"  • 文件大小: {file_size / 1024:.1f} KB")
        print(f"  • 文件类型: {file_ext}")
        print(f"  • 文件名: {os.path.basename(doc_path)}")

        # ====== 第二阶段：COM 初始化 ======
        print(f"\n{'='*80}")
        print(f"[Word 大纲提取] 阶段 1: Word COM 接口初始化")
        print(f"{'='*80}")
        
        print(f"🔧 初始化 COM...")
        init_time = time.time()
        word = win32com.client.DispatchEx("Word.Application")
        init_duration = time.time() - init_time
        print(f"✓ COM 初始化成功 (耗时: {init_duration:.3f}秒)")
        
        print(f"⚙️  配置 Word 应用属性...")
        word.Visible = False  # 不显示Word窗口
        word.DisplayAlerts = 0  # 不显示对话框
        word.AutomationSecurity = 3  # 禁用所有宏
        print(f"✓ Word 应用配置完成")
        print(f"  • 可见性: False")
        print(f"  • 警告提示: 禁用")
        print(f"  • 安全级别: 高 (禁用宏)")

        # ====== 第三阶段：文档打开 ======
        print(f"\n{'='*80}")
        print(f"[Word 大纲提取] 阶段 2: 打开 Word 文档")
        print(f"{'='*80}")
        
        print(f"📖 调用 Word.Documents.Open()...")
        open_start = time.time()
        doc = word.Documents.Open(
            FileName=doc_path,
            ReadOnly=True,  # 以只读模式打开
            AddToRecentFiles=False,  # 不添加到最近使用列表
            ConfirmConversions=False  # 不确认格式转换
        )
        open_duration = time.time() - open_start
        print(f"✓ 文档打开成功 (耗时: {open_duration:.3f}秒)")

        if doc is None:
            raise RuntimeError("文档对象为空（Documents.Open 返回 None）")
        
        # ====== 第四阶段：文档信息收集 ======
        print(f"\n{'='*80}")
        print(f"[Word 大纲提取] 阶段 3: 分析文档结构")
        print(f"{'='*80}")
        
        print(f"📊 正在读取文档属性...")
        total_paras = doc.Paragraphs.Count
        total_styles = doc.Styles.Count
        total_sections = doc.Sections.Count
        
        print(f"✓ 文档属性读取完成")
        print(f"  • 段落总数: {total_paras}")
        print(f"  • 样式数量: {total_styles}")
        print(f"  • 分节数: {total_sections}")

        # ====== 第五阶段：大纲提取 ======
        print(f"\n{'='*80}")
        print(f"[Word 大纲提取] 阶段 4: 扫描大纲段落 (OutlineLevel 1-9)")
        print(f"{'='*80}")
        
        result = []
        outline_level_stats = {}  # 统计各级别大纲数量
        error_count = 0
        extraction_start = time.time()

        print(f"🔍 开始逐段扫描... (共 {total_paras} 个段落)")
        
        for para_idx, para in enumerate(doc.Paragraphs):
            try:
                # 获取段落文本
                text = para.Range.Text
                if not text or text.strip() in ['\r', '\x07', '']:
                    continue

                # 检查大纲级别（1-9为有效大纲）
                outline_level = para.OutlineLevel
                if outline_level < 1 or outline_level > 9:
                    continue
                if outline_level > max_level:
                    continue

                # 统计
                outline_level_stats[outline_level] = outline_level_stats.get(outline_level, 0) + 1

                # ===== 核心增强：双重编号提取 =====
                # 1. 优先尝试自动编号（Word 原生 ListFormat）
                try:
                    list_text = para.Range.ListFormat.ListString.strip()
                except:
                    list_text = ""
                    
                title = text.strip().replace('\r', '').replace('\x07', '')

                # 2. 如果自动编号为空，尝试从文本开头提取手动编号
                if not list_text:
                    # 匹配模式: 1.1.1.1 / 4.1.1.2 / 4.1.1.2、 / 4.1.1.2 查询...
                    # 限制: 最多4级编号（避免匹配"2023年"等）
                    match = re.match(
                        r'^(\d+(?:[\.\．]\d+){0,3})[\s\.．、：:\u3000]?',  # 允许编号后无空格直接接文字
                        title
                    )
                    if match:
                        list_text = match.group(1).replace('．', '.')  # 统一为半角点
                        # 移除编号部分（包括可能的分隔符）
                        title = title[len(match.group(0)):].strip()
                else:
                    # 自动编号存在时，清理标题中重复的编号前缀
                    if title.startswith(list_text):
                        title = title[len(list_text):].strip()
                    # 处理 ListString="4.1.1." 但标题="4.1.1 查询" 的情况
                    elif list_text.endswith('.'):
                        prefix = list_text.rstrip('.')
                        if title.startswith(prefix):
                            title = title[len(prefix):].strip()

                # 清理编号末尾标点
                num_clean = list_text.rstrip('.').rstrip(')').rstrip('：').rstrip(':').strip()

                # 展示级别：优先按数字编号的"点数"映射目录层级
                display_level = outline_level
                if num_clean and re.match(r"^\d+(?:\.\d+)*$", num_clean):
                    # 修正：段数 = 点数 + 1
                    # "4" (0个点) -> 1级, "4.1" (1个点) -> 2级, "4.1.1" (2个点) -> 3级
                    display_level = min(9, max(1, num_clean.count('.') + 1))

                if title:
                    result.append([num_clean, title, display_level])

            except pythoncom.com_error as e:
                # 这是预期的，某些段落可能无法读取COM属性
                error_count += 1
                continue
            except Exception as e:
                # 跳过无法处理的段落
                error_count += 1
                continue

        extraction_duration = time.time() - extraction_start
        
        # ====== 第六阶段：结果统计 ======
        print(f"\n{'='*80}")
        print(f"[Word 大纲提取] 阶段 5: 提取结果统计")
        print(f"{'='*80}")
        
        print(f"✓ 扫描完成 (耗时: {extraction_duration:.3f}秒)")
        print(f"  • 处理段落数: {total_paras}")
        print(f"  • 成功提取大纲数: {len(result)}")
        print(f"  • 错误段落数: {error_count}")
        
        print(f"\n📈 按级别统计:")
        if outline_level_stats:
            for level in sorted(outline_level_stats.keys()):
                count = outline_level_stats[level]
                bar_length = int(count / max(outline_level_stats.values()) * 30)
                bar = "█" * bar_length
                print(f"  • L{level}: {count:4d} 项  {bar}")
        else:
            print(f"  ⚠️  未找到任何大纲段落")

        if result:
            print(f"\n📋 大纲示例 (前10项):")
            for i, (num, title, lvl) in enumerate(result[:10]):
                indent = "  " * (lvl - 1)
                title_short = title[:60] + "..." if len(title) > 60 else title
                print(f"  • {indent}[{num:8s}] {title_short}")

        return result

    finally:
        # ====== 清理阶段 ======
        print(f"\n{'='*80}")
        print(f"[Word 大纲提取] 阶段 6: 资源清理")
        print(f"{'='*80}")
        
        if doc:
            try:
                print(f"🔐 正在关闭文档...")
                close_start = time.time()
                doc.Close(False)
                close_duration = time.time() - close_start
                print(f"✓ 文档已关闭 (耗时: {close_duration:.3f}秒)")
            except Exception as e:
                print(f"⚠️  关闭文档时出错: {str(e)[:100]}")
        
        if word:
            try:
                print(f"🔐 正在退出 Word 应用...")
                quit_start = time.time()
                word.Quit()
                quit_duration = time.time() - quit_start
                print(f"✓ Word 应用已退出 (耗时: {quit_duration:.3f}秒)")
            except Exception as e:
                print(f"⚠️  退出 Word 时出错: {str(e)[:100]}")
        
        pythoncom.CoUninitialize()
        
        total_duration = time.time() - start_time
        print(f"\n{'='*80}")
        print(f"[Word 大纲提取] 完成")
        print(f"{'='*80}")
        print(f"⏱️  总耗时: {total_duration:.3f}秒")
        print(f"{'='*80}\n")


# 为了向后兼容，保留原始函数名
def extract_word_outline(doc_path: str, max_level: int = 9) -> List[Tuple[str, str, int]]:
    """原始函数，现已重定向到带日志版本"""
    return extract_word_outline_with_detailed_log(doc_path, max_level)


def print_word_outline(outline: List[Tuple[str, str, int]]) -> None:
    """简洁输出大纲结构"""
    if not outline:
        print("\n⚠️ 未提取到任何大纲项")
        print("💡 请在 Word 中验证：【视图】→【大纲】是否显示层级结构")
        return

    print(f"\n{'=' * 70}")
    print("Word 大纲视图 (OutlineLevel 1-9)")
    print(f"{'=' * 70}\n")

    for num, title, level in outline:
        indent = '  ' * (level - 1)
        if num:
            print(f'{indent}[{num}] {title}')
        else:
            print(f'{indent}[] {title}')  # 显式显示空编号

    from collections import Counter
    level_dist = Counter(level for _, _, level in outline)
    print(f"\n{'=' * 70}")
    print(f"✓ 共 {len(outline)} 项 | 级别分布: ", end=" ")
    for lvl in sorted(level_dist.keys()):
        print(f"L{lvl}:{level_dist[lvl]}  ", end=" ")
    print()
    print(f"{'=' * 70}")


if __name__ == "__main__":
    import sys

    doc_path = sys.argv[1] if len(sys.argv) > 1 else input("请输入Word文档路径: ")
    print("正在提取 Word 大纲（OutlineLevel 1-9）...")
    try:
        outline = extract_word_outline(doc_path, max_level=9)
        print_word_outline(outline)
    except Exception as e:
        print(f"❌ 提取失败: {str(e)}")
        import traceback

        traceback.print_exc()
        sys.exit(1)
