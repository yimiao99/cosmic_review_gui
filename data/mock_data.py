import random

def get_mock_tasks():
    """获取模拟任务数据"""
    colors = ["#2563eb", "#10b981", "#f59e0b", "#a855f7", "#ef4444", "#06b6d4", "#ec4899"]
    return [
        {
            'filename': 'PRJ-Error-Case. docx',
            'type_label': '结算',
            'days': 120,
            'border_color': random.choice(colors),
            'steps': [
                ('模板校验', 'fail'),
                ('空值检查', 'fail'),
                ('送审比例', 'fail'),
                ('附加值因子', 'warn'),
                ('层级匹配', 'fail'),
                ('功能过程', 'fail')
            ],
            'default_log': '🔴 点击上方红色节点查看具体错误详情。',
            'logs': {
                1: "❌ [模板搬运] Word 文档内容与空白模板相似度 >95%，未检测到实际项目内容修改。",
                2: "❌ [空值检查] Excel Sheet2 存在 5 个空单元格，已标红。",
                3: "❌ [送审比例] 输入人天(120) 与 Excel 行数(450) 偏差过大 (3.75 > 2.0)。",
                4: "⚠️ [附加值因子] 未找到关键质量特性描述，使用默认值。",
                5: "❌ [层级匹配] Word 3.1章 '系统管理' 在 Excel 目录表中未找到对应层级。",
                6: "❌ [功能过程] Word 缺少 '退款流程' 的详细描述。"
            },
            'report_sections': [
                {
                    'title': '1. 模板内容校验',
                    'rows':  [
                        {'key': '模板版本', 'value': 'V3.0 (最新)'},
                        {'key': '未替换占位符', 'value': '12 处', 'style': 'color:  #ef4444; font-weight: bold;'},
                        {'key': '原创度检测', 'value': '极低 (疑似未改)', 'style': 'color:  #ef4444; font-weight: bold;'}
                    ]
                },
                {
                    'title':  '2. 人天 vs 范围计算',
                    'rows':  [
                        {'key': '输入人天', 'value':  '120'},
                        {'key': 'Excel功能行数', 'value': '450'},
                        {'key': '逻辑结论', 'value': '偏差异常 (>30%)', 'style': 'color: #ef4444; font-weight: bold;'}
                    ]
                }
            ]
        },
        {
            'filename': 'PRJ-Budget-Warning.docx',
            'type_label': '预算',
            'days': 85,
            'border_color':  random.choice(["#2563eb", "#10b981", "#f59e0b", "#a855f7", "#ef4444", "#06b6d4", "#ec4899"]),
            'steps': [
                ('模板校验', 'done'),
                ('空值检查', 'done'),
                ('送审比例', 'done'),
                ('附加值因子', 'warn'),
                ('层级匹配', 'done'),
                ('功能过程', 'done')
            ],
            'default_log': '⚠️ 注意：此为预算项目，请关注步骤4的因子提示。',
            'logs': {
                1: "✅ 模板校验通过，未发现遗留占位符。",
                2: "✅ Excel 无空值。",
                3: "✅ 送审比例 (5.3) 在合理范围内。",
                4: "⚠️ [附加值因子] 检测为预算项目。已找到4个需求因子，请人工确认是否合理。",
                5: "✅ 目录结构匹配成功。",
                6: "✅ Excel 中定义的 20 个功能过程，在 Word 中均已描述。"
            },
            'report_sections': [
                {
                    'title': '总体评估',
                    'rows':  [
                        {'key': '状态', 'value': '通过（有警告）', 'style': 'color:  #f59e0b; font-weight: bold;'}
                    ]
                }
            ]
        }
    ]