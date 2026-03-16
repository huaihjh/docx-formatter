from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
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
    QVBoxLayout,
    QWidget,
)

from models.structure import AnalyzerConfig
from services.diagnostic_reporter import DiagnosticReporter
from services.docx_reader import DocxReader
from services.exporter import Exporter
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

        self._setup_strategy_panel(layout)

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

        self.setCentralWidget(container)

    def _setup_strategy_panel(self, parent_layout: QVBoxLayout) -> None:
        group = QGroupBox("结构识别策略")
        form = QFormLayout(group)

        self.chk_split_soft = QCheckBox("拆分软换行")
        self.chk_split_soft.setChecked(True)
        form.addRow(self.chk_split_soft)

        self.chk_split_inline = QCheckBox("拆分行内小标题")
        self.chk_split_inline.setChecked(True)
        form.addRow(self.chk_split_inline)

        self.chk_keep_list = QCheckBox("保留列表项整体")
        self.chk_keep_list.setChecked(True)
        form.addRow(self.chk_keep_list)

        self.spin_unknown_score = QDoubleSpinBox()
        self.spin_unknown_score.setRange(0.0, 10.0)
        self.spin_unknown_score.setSingleStep(0.1)
        self.spin_unknown_score.setValue(2.0)
        form.addRow("unknown 最低分阈值", self.spin_unknown_score)

        self.spin_unknown_margin = QDoubleSpinBox()
        self.spin_unknown_margin.setRange(0.0, 5.0)
        self.spin_unknown_margin.setSingleStep(0.05)
        self.spin_unknown_margin.setValue(0.35)
        form.addRow("unknown 分差阈值", self.spin_unknown_margin)

        self.body_markers_edit = QLineEdit("本,该,通过,针对,为了,其")
        self.body_markers_edit.setPlaceholderText("逗号分隔，例如：本,该,通过")
        form.addRow("正文起始词", self.body_markers_edit)

        self.chk_debug = QCheckBox("启用结构调试输出")
        self.chk_debug.setChecked(os.getenv("DOCX_STRUCTURE_DEBUG", "0") == "1")
        form.addRow(self.chk_debug)

        parent_layout.addWidget(group)

    def _build_analyzer_config(self) -> AnalyzerConfig:
        marker_text = self.body_markers_edit.text().strip()
        markers = tuple(m.strip() for m in marker_text.split(",") if m.strip())

        return AnalyzerConfig(
            split_soft_breaks=self.chk_split_soft.isChecked(),
            split_inline_subheading=self.chk_split_inline.isChecked(),
            keep_list_item_integrity=self.chk_keep_list.isChecked(),
            unknown_score_threshold=float(self.spin_unknown_score.value()),
            unknown_margin_threshold=float(self.spin_unknown_margin.value()),
            body_marker_tokens=markers or AnalyzerConfig().body_marker_tokens,
        )

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
            analysis = StructureAnalyzer.analyze(doc, config=cfg, debug=self.chk_debug.isChecked())
            self.log(f"结构识别完成: {analysis.summary()}")

            self.log("开始应用格式...")
            applied_map = Formatter.apply(doc, rule, analysis)

            self.log("开始导出文档...")
            output_path = Exporter.save(doc, path_obj)

            report_path = output_path.with_name(f"{output_path.stem}_识别报告.json")
            report_payload = DiagnosticReporter.build_report(analysis, applied_map)
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
