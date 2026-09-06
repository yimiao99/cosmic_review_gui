import re
from difflib import SequenceMatcher


class SimilarityChecker:
    """模板相似度校验类"""

    REQUIRED_WORD_MODULES = [
        "1. 需求说明",
        "1.1. 总体描述",
        "1.2. 建设目标",
        "1.3. 建设必要性",
        "2. 功能架构图",
        "3. 功能需求",
        "3.1. 功能需求1(请注明本需求是：新增、优化）",
        "3.1.1. 关键时序图/业务逻辑图",
        "3.1.2. 功能描述",
        "3.4. 需求功能清单",
        "4. 附加值调整因子说明",
        "5.1. 需求变更规模因子",
        "5.2. 应用领域",
        "5.3. 质量及特性",
        "5.4. 开发语言",
        "5.5. 开发团队背景",
        "5.6. 完整性级别调整因子",
    ]

    @staticmethod
    def calculate_string_similarity(a, b):
        if not a or not b:
            return 0.0
        return SequenceMatcher(None, a, b).ratio()

    @classmethod
    def validate_template(cls, target_sections, excel_info, template_sections=None):
        def clean_text(text):
            if not text:
                return ""
            # 去除括号内容 (如指令文字)
            text = re.sub(r"\(.*?\)", "", text)
            text = re.sub(r"（.*?）", "", text)
            # 仅保留中文字符、字母，移除数字、标点和层级符号
            return re.sub(r"[^\u4e00-\u9fa5a-zA-Z]", "", text)

        # 1. 索引目标文档 (支持模糊匹配标题末尾部分)
        # 用标题核心词做 key，解决 1. > 1.1 分级路径导致无法匹配的问题
        target_lookup = {}
        target_level_map = {}  # 新增：记录标题层级

        for s in target_sections:
            if not s:
                continue
            # [FIX] 彻底解决 AttributeError: 'str' object has no attribute 'get'
            if isinstance(s, str):
                t_title = s
                t_level = 2  # 默认
            else:
                # 兼容性修复：优先使用 title，退而求其次使用 display 或 cleaned
                t_title = s.get("title", s.get("display", s.get("cleaned", "")))
                t_level = s.get("level", 2)

            if t_title:
                clean_t = clean_text(t_title)
                target_lookup[clean_t] = s
                target_level_map[clean_t] = t_level

        # 2. 索引模板内容 (作为参考库)
        template_content_map = {}
        if template_sections:
            for s in template_sections:
                if not s:
                    continue
                # [FIX] 同步修复模板提取结果
                if isinstance(s, str):
                    tm_title = s
                else:
                    tm_title = s.get("title", s.get("display", s.get("cleaned", "")))

                if tm_title:
                    t_content = s.get("content", "") if isinstance(s, dict) else ""
                    t_key = clean_text(tm_title)
                    # 同名章节可能因多来源大纲提取出现重复项，空内容不得覆盖已有正文
                    if t_key and (
                        t_key not in template_content_map
                        or (t_content and not template_content_map[t_key])
                    ):
                        template_content_map[t_key] = t_content

        results = {
            "is_valid": True,
            "matched_modules": [],
            "suspect_modules": [],
            "missing_modules": [],
            "all_details": [],  # 每一项是一个查重行
        }

        # 3. 按照 17 个要求的章节进行循环
        for m in cls.REQUIRED_WORD_MODULES:
            m_clean = clean_text(m)
            target = target_lookup.get(m_clean)

            # [NEW] 增强匹配逻辑：解决“功能需求”等项目特化标题导致缺失的问题
            if not target:
                # A. 针对“功能需求1”等包含指令文字的模块，通过包含关系匹配
                if "功能需求" in m_clean:
                    # 尝试在 target_lookup 中找包含“功能”的 level 2/3 标题
                    for t_key, t_val in target_lookup.items():
                        if "功能" in t_key and target_level_map.get(t_key) >= 2:
                            # 只要包含“功能”且是二级及以下标题，且还没被其他必填项占用，就尝试匹配
                            # 这能大幅减少由于“流量控制平台”等实际业务标题导致的缺失误报
                            target = t_val
                            break
                # B. 针对其他可能的同义词
                synonyms = {
                    "总体描述": ["项目概述", "项目简介", "总体说明"],
                    "建设目标": ["建设内容", "建设思路"],
                    "功能架构图": ["系统架构图", "技术架构图", "功能架构"],
                    "质量及特性": [
                        "非功能性需求",
                        "性能需求",
                        "非功能需求",
                        "质量需求",
                    ],
                    "功能需求": [
                        "系统功能需求",
                        "业务功能需求",
                        "功能说明",
                        "主要功能",
                    ],
                }
                if not target and m_clean in synonyms:
                    for syn in synonyms[m_clean]:
                        if syn in target_lookup:
                            target = target_lookup[syn]
                            break

            # 层级计算 (仅为 UI 展示)
            num_match = re.match(r"^(\d+(\.\d+)*)\.?\s*", m)
            level = len(num_match.group(1).split(".")) if num_match else 1

            # --- 章节级结论 ---
            chapter_status = "缺失"
            target_content = ""
            if target:
                chapter_status = "已匹配"  # 占位
                target_content = target.get("content", "").strip()
            else:
                results["missing_modules"].append({"template": m})
                if level == 1:
                    results["is_valid"] = False

            # 获取模板预设值
            base_content = template_content_map.get(m_clean, "").strip()

            # --- 逐句拆分对比 ---
            base_sents = [
                s.strip()
                for s in re.split(r"[。\n!！?？;；]", base_content)
                if len(s.strip()) > 3
            ]
            target_sents = [
                s.strip()
                for s in re.split(r"[。\n!！?？;；]", target_content)
                if len(s.strip()) > 3
            ]

            # 情形 A: 章节缺失
            if not target:
                results["all_details"].append(
                    {
                        "level": level,
                        "chapter": m,
                        "status": "缺失",
                        "tpl_sentence": (
                            base_content[:50] + "..."
                            if base_content
                            else "模板无预设正文"
                        ),
                        "target_sentence": "-",
                        "score": 0.0,
                    }
                )
                continue

            # 情形 B: 既有模板又有内容 -> 逐句对照
            marked_tpl_idx = set()
            section_suspect_count = 0

            # 记录所有比对行项
            comparison_rows = []

            for ts in target_sents:
                best_sim = 0
                best_tpl = ""
                best_idx = -1

                for idx, bs in enumerate(base_sents):
                    sim = SequenceMatcher(None, ts, bs).ratio()
                    if sim > best_sim:
                        best_sim = sim
                        best_tpl = bs
                        best_idx = idx

                status = "✅ 原创"
                tpl_text = "-"
                if best_sim > 0.4:
                    tpl_text = best_tpl
                    marked_tpl_idx.add(best_idx)
                    status = "🚨 雷同" if best_sim > 0.8 else "🔹 参考"
                    if best_sim > 0.85:
                        section_suspect_count += 1

                comparison_rows.append(
                    {
                        "level": level,
                        "chapter": m,
                        "status": status,
                        "tpl_sentence": tpl_text,
                        "target_sentence": ts,
                        "score": best_sim,
                    }
                )

            # 情形 C: 检查模板中有但是文本没用上的句子 (可能被删减)
            for idx, bs in enumerate(base_sents):
                if idx not in marked_tpl_idx:
                    comparison_rows.append(
                        {
                            "level": level,
                            "chapter": m,
                            "status": "🔹 模板参考",
                            "tpl_sentence": bs,
                            "target_sentence": "(未提及)",
                            "score": 0.0,
                        }
                    )

            # 如果章节本身没内容
            if not target_sents and not base_sents:
                results["all_details"].append(
                    {
                        "level": level,
                        "chapter": m,
                        "status": "已匹配",
                        "tpl_sentence": "(空)",
                        "target_sentence": "(空)",
                        "score": 1.0,
                    }
                )
            else:
                for row in comparison_rows:
                    results["all_details"].append(row)

            # 更新模块统计结果
            if section_suspect_count > 0:
                results["suspect_modules"].append({"template": m})
            elif target:
                results["matched_modules"].append({"template": m})

        # 4. 补充额外章节 (存在于目标文档但不在 17 个必填项中的)
        seen_mandatory_cleans = {clean_text(m) for m in cls.REQUIRED_WORD_MODULES}
        extra_chapters = []
        for s in target_sections:
            s_title = s.get("title", s.get("display", s.get("cleaned", "")))
            s_clean = clean_text(s_title)
            if s_clean and s_clean not in seen_mandatory_cleans:
                extra_chapters.append(s)

        for target in extra_chapters:
            m = target.get(
                "title", target.get("display", target.get("cleaned", "未知章节"))
            )
            m_clean = clean_text(m)
            target_content = target.get("content", "").strip()
            base_content = template_content_map.get(m_clean, "").strip()
            level = target.get("level", 2)

            # --- 逐句拆分对比 ---
            base_sents = [
                s.strip()
                for s in re.split(r"[。\n!！?？;；]", base_content)
                if len(s.strip()) > 3
            ]
            target_sents = [
                s.strip()
                for s in re.split(r"[。\n!！?？;；]", target_content)
                if len(s.strip()) > 3
            ]

            marked_tpl_idx = set()
            comparison_rows = []

            for ts in target_sents:
                best_sim = 0
                best_tpl = ""
                best_idx = -1
                for idx, bs in enumerate(base_sents):
                    sim = SequenceMatcher(None, ts, bs).ratio()
                    if sim > best_sim:
                        best_sim = sim
                        best_tpl = bs
                        best_idx = idx

                status = "✅ 原创"
                tpl_text = "-"
                if best_sim > 0.4:
                    tpl_text = best_tpl
                    marked_tpl_idx.add(best_idx)
                    status = "🚨 雷同" if best_sim > 0.8 else "🔹 参考"

                comparison_rows.append(
                    {
                        "level": level,
                        "chapter": m,
                        "status": status,
                        "tpl_sentence": tpl_text,
                        "target_sentence": ts,
                        "score": best_sim,
                    }
                )

            if not target_sents and not base_sents:
                results["all_details"].append(
                    {
                        "level": level,
                        "chapter": m,
                        "status": "已匹配",
                        "tpl_sentence": "(空)",
                        "target_sentence": "(空)",
                        "score": 1.0,
                    }
                )
            else:
                for row in comparison_rows:
                    results["all_details"].append(row)

        # 最后计算总体匹配率
        total_mandatory = len(cls.REQUIRED_WORD_MODULES)
        found_mandatory = total_mandatory - len(results["missing_modules"])
        results["template_match_rate"] = (
            found_mandatory / total_mandatory if total_mandatory > 0 else 0
        )

        return results

    @classmethod
    def validate_function_matching(cls, word_sections, excel_modules):
        """
        比对 Word 三级章节与 Excel 三级模块的匹配度
        """

        def clean_title(t):
            return re.sub(r"[^\u4e00-\u9fa5a-zA-Z]", "", t)

        # 1. 提取 Word 三级标题集 (排除掉“需求说明”等框架章节)
        word_modules = []
        for s in word_sections:
            if s.get("level") == 3:
                title = s.get("title", s.get("display", s.get("cleaned", "")))
                if not title:
                    continue
                # 过滤掉模板框架章节
                if any(x in title for x in ["需求说明", "总体描述", "建设内容"]):
                    continue
                word_modules.append(title)

        # 2. 准备结果容器
        results = {
            "is_ok": True,
            "word_extra": [],  # Word 有，Excel 无
            "excel_extra": [],  # Excel 有，Word 无
            "match_count": 0,
        }

        # --- 性能优化：预处理 ---
        excel_cleans = [clean_title(m) for m in excel_modules]
        excel_set = set(excel_cleans)
        word_cleans = [clean_title(wm) for wm in word_modules]
        word_set = set(word_cleans)

        # 3. 双向比对
        # Word -> Excel
        for wm, wm_clean in zip(word_modules, word_cleans):
            # 优先 O(1) 精确查找
            if wm_clean in excel_set:
                results["match_count"] += 1
                continue

            # 模糊匹配
            found = False
            for ec in excel_cleans:
                if cls.calculate_string_similarity(wm_clean, ec) > 0.6:
                    found = True
                    results["match_count"] += 1
                    break
            if not found:
                results["word_extra"].append(wm)

        # Excel -> Word
        for em, em_clean in zip(excel_modules, excel_cleans):
            # 优先 O(1) 精确查找
            if em_clean in word_set:
                continue

            # 模糊匹配
            found = False
            for wc in word_cleans:
                if cls.calculate_string_similarity(em_clean, wc) > 0.6:
                    found = True
                    break
            if not found:
                results["excel_extra"].append(em)

        if results["word_extra"] or results["excel_extra"]:
            results["is_ok"] = False

        return results
