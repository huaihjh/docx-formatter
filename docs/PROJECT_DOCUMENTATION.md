# Docx 文档格式自动调整工具 - 项目文档

## 1. 项目概述

### 1.1 项目目标
将中文 Word 文档的格式调整流程自动化。用户选择 `.docx` 文件后，输入自然语言格式要求（或选择模板），系统自动识别文档结构并应用格式，导出新文档与识别报告。

### 1.2 目标用户
- 学生：论文格式统一
- 职场用户：报告、公文、简历排版
- 不熟悉 Word 高级格式功能的用户

### 1.3 非目标（当前版本）
- 不编辑文档语义内容
- 不做联网/云端/多人协作
- 不接入大模型 API

## 2. 当前进度（截至当前代码）

### 2.1 已完成
- 桌面 GUI（PySide6）
- 选择本地 `.docx`
- 自然语言规则解析（关键词/正则）
- 通用结构识别流水线：
  - 逻辑块切分
  - 特征提取
  - 规则+打分分类
  - 正文基线统计
- 按分类应用格式（识别与应用已解耦）
- 识别诊断报告 JSON 导出
- 表格单元格递归遍历与处理
- 软换行正文块的“真实拆段”后再应用缩进
- run 级别显式清粗体

### 2.2 正在优化
- 跨文档泛化准确率（标题/列表/行内小标题边界）
- 分类体系与字段命名一致性
- 中英文混排、复杂列表、图表说明判定

### 2.3 当前阶段判断
项目已经从“可运行 Demo”进入“可迭代的通用识别引擎阶段”，核心链路已建立，当前重点是识别准确率与策略参数调优。

## 3. 技术栈与运行环境

- Python 3
- PySide6（桌面界面）
- python-docx（文档读写）
- VS Code（开发环境）

安装与运行：

```bash
pip install -r requirements.txt
python main.py
```

## 4. 目录结构

```text
main.py
ui/
  main_window.py
models/
  format_rule.py
  structure.py
services/
  docx_reader.py
  rule_parser.py
  structure_analyzer.py
  formatter.py
  diagnostic_reporter.py
  exporter.py
```

## 5. 系统架构

统一流程：
1. 读取文档
2. 切分逻辑块
3. 提取特征
4. 分类打分
5. 应用格式
6. 导出文档
7. 导出识别报告

关键设计原则：
- 不把 `paragraph` 直接等同于“逻辑段落”
- 识别与应用分离
- 不确定时允许 `unknown`，避免误改
- 处理结果可诊断、可复核

## 6. 核心模块说明

### 6.1 `services/rule_parser.py`
职责：将自然语言规则解析为 `FormatRule`。

支持要点：
- 字体（宋体/黑体/微软雅黑/楷体/仿宋）
- 中文字号（小四、三号等）与 pt
- 加粗 / 不加粗
- 对齐（左/中/右）
- 行距（单倍/1.5/2倍）
- 首行缩进（含“段落缩进2”表达）
- 段前段后

输出模型：`models/format_rule.py` 中的 `FormatRule` / `SectionRule`。

### 6.2 `services/structure_analyzer.py`
职责：结构识别，不改格式。

#### 逻辑块来源
- 普通段落
- 表格单元格段落（含嵌套表）
- 软换行拆出的行
- 行内拆分（前缀标题 + 正文）

#### 特征
- 文本、长度、样式、对齐、加粗比例、字号
- 是否在表格
- 编号/列表/图表说明模式
- 冒号/句号结尾
- 中文占比、标点密度、长句特征
- 来源位置信息

#### 分类（当前实现）
- `main_heading`
- `sub_heading`
- `inline_subheading`
- `body`
- `list_item`
- `caption`
- `unknown`

说明：`table_cell` 当前作为 `location_type`，不是 `semantic_label`。

#### 分类机制
- 规则+打分
- 正文基线对比（字号/对齐/长度/加粗）
- 上下文（邻接 list_item）
- 低置信时 body fallback 或 unknown

### 6.3 `services/formatter.py`
职责：只按分类结果套规则。

策略：
- `main_heading/sub_heading -> rule.title`
- `inline_subheading -> rule.inline_subheading`
- `body -> rule.body`
- `list_item -> rule.list_item`
- `caption -> rule.caption`
- `location_type == table_cell -> rule.table`
- `unknown -> skip`

新增能力：
- 对“同一真实 paragraph 中由软换行拆出的多个 body/list_item 逻辑块”，先做真实拆段再格式化
- run 级设置字体/字号/加粗，`bold=False` 时显式清除
- 输出 `paragraph_split_applied` 与 `runs_bold_cleared`

### 6.4 `services/diagnostic_reporter.py`
职责：生成识别报告 JSON。

每块输出：
- 文本、分类、得分、原因
- 来源位置、软换行/行内拆分标记
- 特征详情
- 应用规则
- `paragraph_split_applied`
- `runs_bold_cleared`

### 6.5 `ui/main_window.py`
职责：交互流程与参数输入。

包含：
- 文件选择
- 模板选择
- 格式输入
- 结构识别策略面板
- 日志与结果路径显示

## 7. 数据模型

### 7.1 格式规则模型
`FormatRule`：
- `title`
- `body`
- `table`
- `list_item`
- `inline_subheading`
- `caption`

`SectionRule` 字段：
- `font_name`
- `font_size`
- `bold`
- `alignment`
- `line_spacing`
- `first_line_indent`
- `space_before`
- `space_after`

### 7.2 结构模型
- `AnalyzerConfig`
- `BlockFeatures`
- `LogicalBlock`
- `StructureAnalysisResult`

## 8. 诊断报告说明

输出文件：`<输出文档名>_识别报告.json`

用途：
- 复盘为什么某段被判成某类型
- 定位误判来源
- 校验格式应用是否按预期执行

关键字段：
- `semantic_label`
- `location_type`
- `scores`
- `reasons`
- `body_baseline_match_score`
- `final_confidence`
- `applied_rule`
- `paragraph_split_applied`
- `runs_bold_cleared`

## 9. 已知问题与风险

1. 分类集合与文档期望尚有差异
- 需求曾提出 `table_cell` 作为语义类，当前实现将其作为位置类。

2. 复杂文档中的泛化仍需增强
- 深层嵌套编号、手工排版异常、混合样式可能导致置信度不稳。

3. 字体名在中文 Word 的 EastAsia 兼容性
- 目前主要设置 `run.font.name`，某些环境可能需要补充 EastAsia 设置。

4. 真实拆段后的列表语义保持
- 当前策略优先保证缩进与格式一致，复杂编号列表的原编号逻辑需持续验证。

## 10. 下一阶段建议（按优先级）

1. 统一分类设计
- 明确 `semantic_label` 与 `location_type` 的最终规范，保持全链路一致。

2. 提升可配置性
- 在 UI 增加“是否启用真实拆段”“拆段仅限表格/全文”等开关。

3. 建立回归样本集
- 增加 10~30 份多类型中文文档，固化识别与格式回归测试。

4. 增加自动化测试
- `rule_parser` 单测
- `structure_analyzer` 打分与分类单测
- `formatter` 拆段与 run 级格式单测

5. 提升导出可控性
- 支持自定义输出目录与命名策略

## 11. 使用说明（当前推荐）

1. 启动 `python main.py`
2. 选择 `.docx`
3. 输入规则（如：`标题黑体三号居中，正文宋体小四首行缩进2字符，1.5倍行距，正文不要加粗`）
4. 必要时调整“结构识别策略”
5. 点击“开始处理”
6. 查看输出文档与识别报告

## 12. 当前结论

项目核心难点已从“格式参数解析”转为“结构识别准确率与可解释性”。当前代码已经具备通用识别框架与诊断闭环，下一步应围绕样本回归和策略稳定性做工程化完善。
