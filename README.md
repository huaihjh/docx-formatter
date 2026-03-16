# Docx 文档格式自动调整工具

桌面应用（PySide6 + python-docx），用于根据自然语言格式要求自动调整 `.docx` 文档格式，并导出新文档。

## 架构目标
项目已重构为通用“结构识别 + 格式应用”系统：
1. 解析 docx 原始结构
2. 切分逻辑块
3. 提取特征
4. 规则+打分分类
5. 按分类应用格式
6. 输出诊断报告

## 逻辑块与识别
- 不把 Word `paragraph` 直接等同逻辑段落
- 逻辑块来源：
  - 普通段落
  - 表格单元格段落（含嵌套表格）
  - 软换行拆分后的行
  - 行内“前缀+正文”拆分后的子块
- 分类标签：
  - `main_heading`
  - `sub_heading`
  - `inline_subheading`
  - `body`
  - `list_item`
  - `table_cell`
  - `caption`
  - `unknown`

## 特征（示例）
每个逻辑块会提取：
- 原始文本、清理文本、文本长度、是否为空
- 样式名、是否在表格、对齐方式
- 加粗/斜体比例、主字号
- 是否软换行拆分
- 编号模式、列表项模式
- 冒号/句号结尾
- 中文字符占比、标点密度、长句特征
- 来源位置（文档/表格/行/拆分块）

## 正文基线分析
识别前会先统计文档基线（最常见正文字体/字号/对齐/长度/加粗），分类时用“块特征 vs 文档基线”比较，而非固定模板。

## 识别与应用解耦
- `services/structure_analyzer.py`：仅负责结构识别，不改格式
- `services/formatter.py`：仅按分类结果套格式，不做临时识别判断

## 规则应用策略
- `main_heading` / `sub_heading`：标题规则
- `inline_subheading`：行内小标题规则
- `body`：正文规则
- `list_item`：列表规则
- `caption`：图表说明规则
- `table_cell`：表格规则
- `unknown`：默认不改

## 诊断报告
每次处理会输出 `*_识别报告.json`，包含每个逻辑块：
- 原文本
- 分类结果
- 分类得分
- 识别原因
- 来源位置
- 是否在表格中
- 是否软换行拆分
- 应用了哪种格式规则

## 可配置策略
界面“结构识别策略”面板可调：
- `split_soft_breaks`
- `split_inline_subheading`
- `keep_list_item_integrity`
- `unknown_score_threshold`
- `unknown_margin_threshold`
- `body_marker_tokens`
- 调试输出开关

## 目录结构
- `main.py`：程序入口
- `ui/main_window.py`：主界面与交互流程
- `services/docx_reader.py`：文档读取
- `services/rule_parser.py`：自然语言规则解析
- `services/structure_analyzer.py`：逻辑块切分、特征、分类打分
- `services/formatter.py`：按分类结果应用格式
- `services/diagnostic_reporter.py`：生成识别诊断报告
- `services/exporter.py`：导出新文档
- `models/format_rule.py`：格式规则模型
- `models/structure.py`：结构分析模型与配置

## 运行方式
```bash
pip install -r requirements.txt
python main.py
```

## 完整项目文档
- [项目文档](docs/PROJECT_DOCUMENTATION.md)
