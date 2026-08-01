"""主界面：拖入 Excel → 选模板 → 配字段映射 → 勾选数据行 → 打印。"""
from __future__ import annotations

import sys
import traceback
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFormLayout,
    QGroupBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QSpinBox,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from ..core.backend import BackendError, PrintBackend
from ..core.excel_reader import SheetData, load_excel
from ..core.profiles import (
    SOURCE_COLUMN, SOURCE_FIXED, SOURCE_NONE,
    FieldMapping, Profile, guess_mapping, load_profile, save_profile,
)

OPTION_NONE = "（留空）"
OPTION_FIXED = "（固定文本）"


def app_data_dir() -> Path:
    """exe 同目录（打包后）或项目根目录（开发时）。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[2]


class PrintWorker(QThread):
    """后台打印线程，避免卡住界面。"""

    progress = Signal(int, int, str)   # 当前, 总数, 消息
    log = Signal(str)
    done = Signal(bool, str)           # 是否成功, 总结消息

    def __init__(self, backend: PrintBackend, template: str, printer: str,
                 jobs: list[tuple[int, dict[str, str], int]], parent=None):
        super().__init__(parent)
        self._backend = backend
        self._template = template
        self._printer = printer
        self._jobs = jobs              # (行号, 字段值, 张数)
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        ok, failed = 0, 0
        try:
            self._backend.open(self._template, self._printer)
            total = len(self._jobs)
            for i, (row_no, values, copies) in enumerate(self._jobs, 1):
                if self._cancelled:
                    self.done.emit(False, f"已取消（完成 {ok} 行，失败 {failed} 行）")
                    return
                label = values.get("发货新编号") or f"第 {row_no} 行"
                try:
                    self._backend.print_label(values, copies)
                    ok += 1
                    self.progress.emit(i, total, f"已发送：{label} × {copies} 张")
                except Exception as exc:
                    failed += 1
                    self.log.emit(f"[失败] 第 {row_no} 行（{label}）：{exc}")
            msg = f"打印完成：成功 {ok} 行" + (f"，失败 {failed} 行" if failed else "")
            self.done.emit(failed == 0, msg)
        except Exception as exc:
            self.done.emit(False, f"打印中断：{exc}")
        finally:
            try:
                self._backend.close()
            except Exception:
                pass


class MainWindow(QMainWindow):
    def __init__(self, backend: PrintBackend):
        super().__init__()
        self.backend = backend
        self.sheet: SheetData | None = None
        self.template_fields: list[str] = []
        self.worker: PrintWorker | None = None

        self.setWindowTitle(f"标签自动打印工具 — {backend.name}")
        self.resize(1100, 800)
        self.setAcceptDrops(True)

        self._build_ui()
        self._scan_templates()

    # ---------------------------------------------------------------- UI

    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)

        # --- 1. Excel ---
        excel_box = QGroupBox("1. Excel 清单（可把文件拖进本窗口）")
        form = QFormLayout(excel_box)
        row = QHBoxLayout()
        self.excel_path = QLineEdit()
        self.excel_path.setPlaceholderText("选择或拖入 .xls / .xlsx 文件…")
        self.excel_path.setReadOnly(True)
        btn_excel = QPushButton("浏览…")
        btn_excel.clicked.connect(self._browse_excel)
        row.addWidget(self.excel_path, 1)
        row.addWidget(btn_excel)
        form.addRow("文件：", row)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("表头在第"))
        self.header_row = QSpinBox()
        self.header_row.setMinimum(1)
        self.header_row.setMaximum(50)
        self.header_row.setValue(1)
        self.header_row.valueChanged.connect(self._reload_excel)
        row2.addWidget(self.header_row)
        row2.addWidget(QLabel("行（一般自动识别，无需修改）"))
        row2.addStretch(1)
        form.addRow("表头：", row2)
        layout.addWidget(excel_box)

        # --- 2. 模板 ---
        tpl_box = QGroupBox("2. BarTender 标签模板")
        tform = QFormLayout(tpl_box)
        trow = QHBoxLayout()
        self.template_combo = QComboBox()
        self.template_combo.currentIndexChanged.connect(self._template_changed)
        btn_tpl = QPushButton("浏览…")
        btn_tpl.clicked.connect(self._browse_template)
        trow.addWidget(self.template_combo, 1)
        trow.addWidget(btn_tpl)
        tform.addRow("模板：", trow)
        layout.addWidget(tpl_box)

        # --- 3. 字段映射 ---
        map_box = QGroupBox("3. 字段映射（模板字段 ← 取哪一列 / 固定文本）")
        mlay = QVBoxLayout(map_box)
        self.map_table = QTableWidget(0, 3)
        self.map_table.setHorizontalHeaderLabels(["模板字段", "取值来源", "固定文本 / 列名"])
        self.map_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.map_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.map_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        mlay.addWidget(self.map_table)
        layout.addWidget(map_box)

        # --- 4. 数据预览 ---
        data_box = QGroupBox("4. 数据预览（勾选要打印的行，可改每行张数）")
        dlay = QVBoxLayout(data_box)
        drow = QHBoxLayout()
        btn_all = QPushButton("全选")
        btn_none = QPushButton("全不选")
        btn_all.clicked.connect(lambda: self._set_all_checked(True))
        btn_none.clicked.connect(lambda: self._set_all_checked(False))
        self.data_summary = QLabel("未加载 Excel")
        drow.addWidget(btn_all)
        drow.addWidget(btn_none)
        drow.addStretch(1)
        drow.addWidget(self.data_summary)
        dlay.addLayout(drow)
        self.data_table = QTableWidget(0, 0)
        dlay.addWidget(self.data_table)
        layout.addWidget(data_box, 1)

        # --- 5. 打印 ---
        print_box = QGroupBox("5. 打印")
        play = QVBoxLayout(print_box)
        prow = QHBoxLayout()
        prow.addWidget(QLabel("打印机："))
        self.printer_combo = QComboBox()
        self.printer_combo.setEditable(False)
        self._refresh_printers()
        btn_printer = QPushButton("刷新")
        btn_printer.clicked.connect(self._refresh_printers)
        prow.addWidget(self.printer_combo, 1)
        prow.addWidget(btn_printer)
        self.btn_test = QPushButton("试打第一张")
        self.btn_test.clicked.connect(lambda: self._start_print(test_only=True))
        self.btn_print = QPushButton("全部打印")
        self.btn_print.clicked.connect(lambda: self._start_print(test_only=False))
        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._cancel_print)
        prow.addWidget(self.btn_test)
        prow.addWidget(self.btn_print)
        prow.addWidget(self.btn_cancel)
        play.addLayout(prow)
        self.progress = QProgressBar()
        self.progress.setValue(0)
        play.addWidget(self.progress)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumBlockCount(2000)
        self.log_view.setPlaceholderText("日志…")
        play.addWidget(self.log_view)
        layout.addWidget(print_box)

    # ------------------------------------------------------------ 拖放

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith((".xls", ".xlsx")):
                self._load_excel(path)
                return
            if path.lower().endswith(".btw"):
                self._add_template(path)
                return
        self.log("拖入的文件不是 .xls/.xlsx/.btw，已忽略")

    # ---------------------------------------------------------- Excel

    def _browse_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 清单", str(app_data_dir()),
            "Excel 文件 (*.xls *.xlsx)")
        if path:
            self._load_excel(path)

    def _load_excel(self, path: str, header_row: int | None = None):
        try:
            sheet = load_excel(path, header_row)
        except Exception as exc:
            QMessageBox.critical(self, "Excel 读取失败", str(exc))
            return
        self.sheet = sheet
        self.excel_path.setText(path)
        # 用探测结果回填 spinbox（暂时屏蔽信号避免递归重载）
        self.header_row.blockSignals(True)
        self.header_row.setValue(sheet.header_row + 1)
        self.header_row.blockSignals(False)
        self._fill_data_table()
        self._auto_map()
        self.log(f"已加载 Excel：{Path(path).name}，共 {len(sheet)} 行数据"
                 f"（表头第 {sheet.header_row + 1} 行）")

    def _reload_excel(self):
        if self.excel_path.text():
            self._load_excel(self.excel_path.text(),
                             header_row=self.header_row.value() - 1)

    def _fill_data_table(self):
        sheet = self.sheet
        assert sheet is not None
        cols = ["打印", "张数"] + sheet.columns
        self.data_table.setRowCount(len(sheet))
        self.data_table.setColumnCount(len(cols))
        self.data_table.setHorizontalHeaderLabels(cols)
        for r, row in enumerate(sheet.rows):
            chk = QCheckBox()
            chk.setChecked(True)
            self.data_table.setCellWidget(r, 0, chk)
            spin = QSpinBox()
            spin.setMinimum(1)
            spin.setMaximum(999)
            spin.setValue(sheet.copies[r])
            self.data_table.setCellWidget(r, 1, spin)
            for c, val in enumerate(row):
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.data_table.setItem(r, 2 + c, item)
        self.data_table.resizeColumnsToContents()
        total = sum(sheet.copies)
        self.data_summary.setText(f"共 {len(sheet)} 行，预计 {total} 张")

    def _set_all_checked(self, checked: bool):
        for r in range(self.data_table.rowCount()):
            w = self.data_table.cellWidget(r, 0)
            if isinstance(w, QCheckBox):
                w.setChecked(checked)

    # ---------------------------------------------------------- 模板

    def _scan_templates(self):
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        tpl_dir = app_data_dir() / "模板"
        if tpl_dir.is_dir():
            for p in sorted(tpl_dir.glob("*.btw")):
                self.template_combo.addItem(p.name, str(p))
        self.template_combo.blockSignals(False)
        if self.template_combo.count():
            self._template_changed()
        else:
            self.log("未在「模板」目录找到 .btw 文件，请点“浏览…”手动选择")

    def _browse_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 BarTender 模板", str(app_data_dir()),
            "BarTender 模板 (*.btw)")
        if path:
            self._add_template(path)

    def _add_template(self, path: str):
        idx = self.template_combo.findData(path)
        if idx < 0:
            self.template_combo.addItem(Path(path).name, path)
            idx = self.template_combo.count() - 1
        self.template_combo.setCurrentIndex(idx)

    def _template_changed(self):
        path = self.template_combo.currentData()
        if not path:
            return
        try:
            self.template_fields = self.backend.list_template_fields(path)
        except BackendError as exc:
            QMessageBox.critical(self, "模板读取失败", str(exc))
            self.template_fields = []
        if not self.template_fields:
            self.map_table.setRowCount(0)
            self.log(f"模板 {Path(path).name} 中没有命名数据源。\n"
                     "请先在 BarTender 里给要动态填充的对象起“数据源名称”。")
            return
        self.log(f"模板 {Path(path).name}：发现 {len(self.template_fields)} 个字段："
                 + "、".join(self.template_fields))
        self._fill_mapping_table()

    # ---------------------------------------------------------- 映射

    def _fill_mapping_table(self):
        columns = self.sheet.columns if self.sheet else []
        saved = None
        path = self.template_combo.currentData()
        if path:
            saved = load_profile(path, app_data_dir() / "profiles")

        self.map_table.setRowCount(len(self.template_fields))
        for r, field_name in enumerate(self.template_fields):
            item = QTableWidgetItem(field_name)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            self.map_table.setItem(r, 0, item)

            mapping = (saved.mapping_for(field_name) if saved else None)
            if mapping is None:
                mapping = guess_mapping(field_name, columns)

            combo = QComboBox()
            combo.addItem(OPTION_NONE)
            combo.addItem(OPTION_FIXED)
            for col in columns:
                combo.addItem(col)
            edit = QLineEdit()
            edit.setPlaceholderText("选择 Excel 列，或选固定文本后在此输入")

            if mapping.source == SOURCE_COLUMN and mapping.value in columns:
                combo.setCurrentText(mapping.value)
                edit.setText(mapping.value)
                edit.setReadOnly(True)
            elif mapping.source == SOURCE_FIXED:
                combo.setCurrentText(OPTION_FIXED)
                edit.setText(mapping.value)
            else:
                combo.setCurrentText(OPTION_NONE)
                edit.setText("")
                edit.setReadOnly(True)

            combo.currentTextChanged.connect(
                lambda text, e=edit: self._mapping_source_changed(text, e))
            self.map_table.setCellWidget(r, 1, combo)
            self.map_table.setCellWidget(r, 2, edit)
        self.map_table.resizeColumnsToContents()

    @staticmethod
    def _mapping_source_changed(text: str, edit: QLineEdit):
        if text == OPTION_FIXED:
            edit.setReadOnly(False)
            edit.setPlaceholderText("输入固定文本")
            edit.setFocus()
        elif text == OPTION_NONE:
            edit.setReadOnly(True)
            edit.setText("")
        else:
            edit.setReadOnly(True)
            edit.setText(text)

    def _auto_map(self):
        if self.template_fields:
            self._fill_mapping_table()

    def _current_mappings(self) -> list[FieldMapping]:
        out = []
        for r in range(self.map_table.rowCount()):
            field_name = self.map_table.item(r, 0).text()
            combo: QComboBox = self.map_table.cellWidget(r, 1)
            edit: QLineEdit = self.map_table.cellWidget(r, 2)
            src_text = combo.currentText()
            if src_text == OPTION_NONE:
                out.append(FieldMapping(field_name, SOURCE_NONE, ""))
            elif src_text == OPTION_FIXED:
                out.append(FieldMapping(field_name, SOURCE_FIXED, edit.text()))
            else:
                out.append(FieldMapping(field_name, SOURCE_COLUMN, src_text))
        return out

    # ---------------------------------------------------------- 打印

    def _refresh_printers(self):
        current = self.printer_combo.currentText()
        printers = self.backend.list_printers()
        self.printer_combo.clear()
        self.printer_combo.addItem("（使用模板默认打印机）", "")
        for p in printers:
            self.printer_combo.addItem(p, p)
        idx = self.printer_combo.findText(current)
        if idx >= 0:
            self.printer_combo.setCurrentIndex(idx)

    def _build_jobs(self, test_only: bool) -> list[tuple[int, dict[str, str], int]] | None:
        """收集勾选行的打印任务；返回 None 表示输入有误（已提示用户）。"""
        if self.sheet is None:
            QMessageBox.warning(self, "缺少数据", "请先加载 Excel 清单")
            return None
        if not self.template_combo.currentData():
            QMessageBox.warning(self, "缺少模板", "请先选择 BarTender 模板")
            return None

        mappings = self._current_mappings()
        col_index = {c: i for i, c in enumerate(self.sheet.columns)}
        jobs = []
        for r in range(self.data_table.rowCount()):
            chk: QCheckBox = self.data_table.cellWidget(r, 0)
            if not chk.isChecked():
                continue
            spin: QSpinBox = self.data_table.cellWidget(r, 1)
            row = self.sheet.rows[r]
            values = {}
            for m in mappings:
                if m.source == SOURCE_COLUMN and m.value in col_index:
                    values[m.field] = row[col_index[m.value]]
                elif m.source == SOURCE_FIXED:
                    values[m.field] = m.value
            jobs.append((r + 1, values, spin.value()))
            if test_only:
                break
        if not jobs:
            QMessageBox.warning(self, "没有选中行", "请至少勾选一行数据")
            return None
        return jobs

    def _start_print(self, test_only: bool):
        if self.worker is not None and self.worker.isRunning():
            return
        jobs = self._build_jobs(test_only)
        if jobs is None:
            return

        template = self.template_combo.currentData()
        printer = self.printer_combo.currentData() or ""

        # 打印前保存映射配置
        save_profile(Profile(template, printer, self._current_mappings()),
                     app_data_dir() / "profiles")

        total_labels = sum(c for _, _, c in jobs)
        if not test_only:
            ret = QMessageBox.question(
                self, "确认打印",
                f"将向「{printer or '模板默认打印机'}」发送 {len(jobs)} 行、"
                f"共 {total_labels} 张标签，是否继续？")
            if ret != QMessageBox.Yes:
                return

        self.log(f"开始{'试打' if test_only else '打印'}：模板 {Path(template).name}，"
                 f"打印机 {printer or '（模板默认）'}，{len(jobs)} 行 / {total_labels} 张")
        self.worker = PrintWorker(self.backend, template, printer, jobs, self)
        self.worker.progress.connect(self._on_progress)
        self.worker.log.connect(self.log)
        self.worker.done.connect(self._on_done)
        self._set_printing_state(True)
        self.progress.setMaximum(len(jobs))
        self.progress.setValue(0)
        self.worker.start()

    def _cancel_print(self):
        if self.worker is not None:
            self.worker.cancel()
            self.log("正在取消…")

    def _on_progress(self, current: int, total: int, message: str):
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        self.log(message)

    def _on_done(self, success: bool, message: str):
        self._set_printing_state(False)
        self.log(("✔ " if success else "✘ ") + message)
        if not success:
            QMessageBox.warning(self, "打印结果", message)

    def _set_printing_state(self, printing: bool):
        self.btn_print.setEnabled(not printing)
        self.btn_test.setEnabled(not printing)
        self.btn_cancel.setEnabled(printing)

    # ---------------------------------------------------------- 杂项

    def log(self, message: str):
        self.log_view.appendPlainText(message)

    def closeEvent(self, event):
        if self.worker is not None and self.worker.isRunning():
            self.worker.cancel()
            self.worker.wait(3000)
        try:
            self.backend.close()
        except Exception:
            traceback.print_exc()
        super().closeEvent(event)
