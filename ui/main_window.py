from __future__ import annotations

import json
from pathlib import Path

from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from models.structure import AnalyzerConfig
from services.diagnostic_reporter import (
    DiagnosticReporter,
    ReportBuildError,
    ReportWriteError,
)
from services.docx_reader import DocxReader
from services.exporter import ExportError, Exporter
from services.formatter import Formatter
from services.rule_parser import RuleParseError, RuleParser
from services.structure_analyzer import StructureAnalyzer


class MainWindow(QMainWindow):
    PRESET_TEMPLATES = {
        "不使用模板": "",
        "论文模板": "标题黑体三号居中，正文宋体小四首行缩进2字符，1.5倍行距",
        "公文模板": "标题黑体二号居中加粗，正文仿宋四号首行缩进2字符，单倍行距",
        "简历模板": "标题微软雅黑三号居中加粗，正文宋体五号左对齐，表格宋体五号",
    }

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Docx 文档格式自动调整工具")
        self.resize(980, 760)
        self._setup_ui()

    def _setup_ui(self) -> None:
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)

        container = QWidget(self)
        layout = QVBoxLayout(container)

        row_file = QHBoxLayout()
        self.btn_choose = QPushButton("选择文档")
        self.btn_choose.clicked.connect(self.on_choose_file)
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("请选择 .docx 文件")
        row_file.addWidget(self.btn_choose)
        row_file.addWidget(self.file_path_edit)
        layout.addLayout(row_file)

        row_template = QHBoxLayout()
        row_template.addWidget(QLabel("预设模板"))
        self.template_combo = QComboBox()
        self.template_combo.addItems(self.PRESET_TEMPLATES.keys())
        self.template_combo.currentTextChanged.connect(self.on_template_changed)
        row_template.addWidget(self.template_combo)
        layout.addLayout(row_template)

        self.rule_input = QPlainTextEdit()
        self.rule_input.setPlaceholderText(
            "请输入格式要求，例如：标题黑体三号居中，正文宋体小四首行缩进2字符，1.5倍行距，表格宋体五号"
        )
        self.rule_input.setFixedHeight(130)
        layout.addWidget(self.rule_input)

        self._setup_rule_override_panel(layout)

        self.btn_run = QPushButton("开始处理")
        self.btn_run.clicked.connect(self.on_process)
        layout.addWidget(self.btn_run)

        layout.addWidget(QLabel("日志输出"))
        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output)

        layout.addWidget(QLabel("导出结果"))
        self.result_output = QLineEdit()
        self.result_output.setReadOnly(True)
        layout.addWidget(self.result_output)

        scroll.setWidget(container)
        self.setCentralWidget(scroll)

    def _setup_rule_override_panel(self, parent_layout: QVBoxLayout) -> None:
        group = QGroupBox("\u89c4\u5219\u89e3\u6790\u9884\u89c8\u4e0e\u8986\u76d6")
        form = QFormLayout(group)

        self.title_align_override = QComboBox()
        self.title_align_override.addItems(["\u81ea\u52a8(\u4e0d\u8986\u76d6)", "\u5de6\u5bf9\u9f50", "\u5c45\u4e2d", "\u53f3\u5bf9\u9f50"])
        form.addRow("\u6807\u9898\u5bf9\u9f50\u8986\u76d6", self.title_align_override)

        self.title_bold_override = QComboBox()
        self.title_bold_override.addItems(["\u81ea\u52a8(\u4e0d\u8986\u76d6)", "\u52a0\u7c97", "\u4e0d\u52a0\u7c97"])
        form.addRow("\u6807\u9898\u52a0\u7c97\u8986\u76d6", self.title_bold_override)

        self.body_bold_override = QComboBox()
        self.body_bold_override.addItems(["\u81ea\u52a8(\u4e0d\u8986\u76d6)", "\u52a0\u7c97", "\u4e0d\u52a0\u7c97"])
        form.addRow("\u6b63\u6587\u52a0\u7c97\u8986\u76d6", self.body_bold_override)

        self.table_bold_override = QComboBox()
        self.table_bold_override.addItems(["\u81ea\u52a8(\u4e0d\u8986\u76d6)", "\u52a0\u7c97", "\u4e0d\u52a0\u7c97"])
        form.addRow("\u8868\u683c\u52a0\u7c97\u8986\u76d6", self.table_bold_override)

        self.table_indent_override = QComboBox()
        self.table_indent_override.addItems(["\u81ea\u52a8(\u4e0d\u8986\u76d6)", "\u7ee7\u627f\u6b63\u6587\u9996\u884c\u7f29\u8fdb", "\u4e0d\u7f29\u8fdb"])
        form.addRow("\u8868\u683c\u9996\u884c\u7f29\u8fdb\u8986\u76d6", self.table_indent_override)

        self.btn_preview_rule = QPushButton("\u89e3\u6790\u5e76\u9884\u89c8\u89c4\u5219")
        self.btn_preview_rule.clicked.connect(self.on_preview_rule)
        form.addRow(self.btn_preview_rule)

        self.rule_preview = QPlainTextEdit()
        self.rule_preview.setReadOnly(True)
        self.rule_preview.setFixedHeight(140)
        form.addRow("\u6700\u7ec8\u89c4\u5219\u9884\u89c8(JSON)", self.rule_preview)

        parent_layout.addWidget(group)

    @staticmethod
    def _apply_bold_override(combo_text: str, section_rule) -> None:
        if combo_text == "\u52a0\u7c97":
            section_rule.bold = True
        elif combo_text == "\u4e0d\u52a0\u7c97":
            section_rule.bold = False

    def _apply_rule_overrides(self, rule) -> None:
        align_map = {"\u5de6\u5bf9\u9f50": "left", "\u5c45\u4e2d": "center", "\u53f3\u5bf9\u9f50": "right"}
        title_align = align_map.get(self.title_align_override.currentText())
        if title_align:
            rule.title.alignment = title_align

        self._apply_bold_override(self.title_bold_override.currentText(), rule.title)
        self._apply_bold_override(self.body_bold_override.currentText(), rule.body)
        self._apply_bold_override(self.table_bold_override.currentText(), rule.table)

        table_indent_text = self.table_indent_override.currentText()
        if table_indent_text == "\u7ee7\u627f\u6b63\u6587\u9996\u884c\u7f29\u8fdb":
            if rule.body.first_line_indent is not None:
                rule.table.first_line_indent = rule.body.first_line_indent
        elif table_indent_text == "\u4e0d\u7f29\u8fdb":
            rule.table.first_line_indent = 0.0

    def _build_rule_from_inputs(self, rule_text: str):
        rule = RuleParser.parse(rule_text)
        self._apply_rule_overrides(rule)
        rule.normalize()
        return rule

    def on_preview_rule(self) -> None:
        rule_text = self.rule_input.toPlainText().strip()
        if not rule_text:
            self._show_error("\u8bf7\u5148\u8f93\u5165\u683c\u5f0f\u8981\u6c42\u6216\u9009\u62e9\u6a21\u677f")
            return

        try:
            rule = self._build_rule_from_inputs(rule_text)
            self.rule_preview.setPlainText(json.dumps(rule.to_dict(), ensure_ascii=False, indent=2))
            self.log("\u89c4\u5219\u9884\u89c8\u5df2\u66f4\u65b0\uff08\u7ed3\u6784\u5316\u8986\u76d6\u4f18\u5148\u4e8e\u81ea\u7136\u8bed\u8a00\u89e3\u6790\uff09")
        except RuleParseError as exc:
            self._show_error(f"\u89c4\u5219\u89e3\u6790\u5931\u8d25: {exc}")

    def _build_analyzer_config(self) -> AnalyzerConfig:
        return AnalyzerConfig()

    def log(self, text: str) -> None:
        self.log_output.appendPlainText(text)

    def on_template_changed(self, template_name: str) -> None:
        template_text = self.PRESET_TEMPLATES.get(template_name, "")
        if template_text:
            self.rule_input.setPlainText(template_text)
            self.log(f"已应用模板: {template_name}")

    def on_choose_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择 DOCX 文件",
            "",
            "Word 文档 (*.docx)",
        )
        if file_path:
            self.file_path_edit.setText(file_path)
            self.log(f"已选择文件: {file_path}")

    def on_process(self) -> None:
        file_path = self.file_path_edit.text().strip()
        rule_text = self.rule_input.toPlainText().strip()

        if not file_path:
            self._show_error("请先选择 .docx 文件")
            return
        if not rule_text:
            self._show_error("请先输入格式要求或选择模板")
            return

        path_obj = Path(file_path)
        if path_obj.suffix.lower() != ".docx":
            self._show_error("仅支持 .docx 文件")
            return
        if not path_obj.exists():
            self._show_error("文件不存在，请重新选择")
            return

        try:
            self.log("开始解析规则...")
            rule = RuleParser.parse(rule_text)
            rule.normalize()
            self.log(f"规则解析成功: {rule.to_dict()}")

            self.log("开始读取文档...")
            doc = DocxReader.load(path_obj)

            cfg = self._build_analyzer_config()
            self.log(f"识别策略: {cfg}")
            self.log("开始结构识别...")
            analysis = StructureAnalyzer.analyze(doc, config=cfg, debug=False)
            self.log(f"结构识别完成: {analysis.summary()}")

            self.log("开始应用格式...")
            applied_map = Formatter.apply(doc, rule, analysis)

            self.log("开始导出文档...")
            output_path = Exporter.save(doc, path_obj)

            report_path = output_path.with_name(f"{output_path.stem}_识别报告.json")
            report_payload = DiagnosticReporter.build_report(
                analysis,
                applied_map,
                source_path=path_obj,
                output_path=output_path,
            )
            DiagnosticReporter.write_json(report_path, report_payload)

            self.result_output.setText(str(output_path))
            self.log(f"处理完成，输出文件: {output_path}")
            self.log(f"识别报告: {report_path}")
            QMessageBox.information(self, "完成", "处理完成")
        except RuleParseError as exc:
            self._show_error(f"规则解析失败: {exc}")
        except PermissionError:
            self._show_error("导出失败，可能没有写入权限")
        except Exception as exc:  # noqa: BLE001
            self._show_error(f"处理失败: {exc}")

    def _show_error(self, message: str) -> None:
        self.log(message)
        QMessageBox.warning(self, "错误", message)



