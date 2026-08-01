"""主界面：拖入 Excel → 选模板 → 配字段映射 → 勾选数据行 → 预览/预检 → 打印。

线程纪律：所有后端（BarTender COM / mock）操作都在 QThread 内完成，
后端实例由 worker 在线程内通过工厂创建、使用、销毁，绝不跨线程共享。
"""
from __future__ import annotations

import json
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QThread, QTimer, Signal
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QFileDialog, QFormLayout,
    QGroupBox, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QSpinBox,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from ..core.backend import BackendError, PrintBackend
from ..core.excel_reader import SheetData, load_excel
from ..core.paths import logs_dir, previews_dir, profiles_dir, resource_dir
from ..core.profiles import (
    SOURCE_COLUMN, SOURCE_FIXED, SOURCE_NONE,
    FieldMapping, Profile, guess_mapping, load_profile, save_profile,
    validate_jobs,
)

OPTION_NONE = "（留空）"
OPTION_FIXED = "（固定文本）"


def _co_init():
    """在当前线程初始化 COM；无 pywin32 时静默跳过（mock 不需要）。"""
    try:
        import pythoncom
        pythoncom.CoInitialize()
        return pythoncom
    except Exception:
        return None


class _BackendWorker(QThread):
    """基类：线程内 CoInitialize → 工厂创建后端 → 子类做事 → close → CoUninitialize。"""

    def __init__(self, backend_factory, parent=None):
        super().__init__(parent)
        self._factory = backend_factory
        self.backend: PrintBackend | None = None
        self.thread_id: int | None = None

    def _setup(self):
        self.thread_id = threading.get_ident()
        self._pythoncom = _co_init()
        self.backend = self._factory()  # 关键：后端在本线程内创建

    def _teardown(self):
        try:
            if self.backend is not None:
                self.backend.close()
        except Exception:
            pass
        if getattr(self, "_pythoncom", None) is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass


class TemplateScanWorker(_BackendWorker):
    """异步枚举模板字段，避免阻塞 GUI 线程。"""

    fields_ready = Signal(int, list)      # epoch, 字段名列表
    failed = Signal(int, str)             # epoch, 错误消息

    def __init__(self, backend_factory, template_path: str, epoch: int, parent=None):
        super().__init__(backend_factory, parent)
        self._template_path = template_path
        self._epoch = epoch

    def run(self):
        try:
            self._setup()
            fields = self.backend.list_template_fields(self._template_path)
            self.fields_ready.emit(self._epoch, fields)
        except Exception as exc:
            self.failed.emit(self._epoch, str(exc))
        finally:
            self._teardown()


class PreviewWorker(_BackendWorker):
    """渲染一张标签预览图。"""

    preview_ready = Signal(str)   # 图片路径
    failed = Signal(str)

    def __init__(self, backend_factory, template: str, printer: str,
                 values: dict[str, str], out_path: str, parent=None):
        super().__init__(backend_factory, parent)
        self._template = template
        self._printer = printer
        self._values = values
        self._out_path = out_path

    def run(self):
        try:
            self._setup()
            self.backend.open(self._template, self._printer)
            path = self.backend.render_preview(self._values, self._out_path)
            self.preview_ready.emit(path)
        except Exception as exc:
            self.failed.emit(str(exc))
        finally:
            self._teardown()


class PrintWorker(_BackendWorker):
    """后台打印线程，避免卡住界面。"""

    progress = Signal(int, int, str)   # 当前, 总数, 消息
    log = Signal(str)
    done = Signal(bool, str, list)     # 是否成功, 总结, 每行状态 list[dict]

    def __init__(self, backend_factory, template: str, printer: str,
                 jobs: list[tuple[int, dict[str, str], int]], parent=None):
        super().__init__(backend_factory, parent)
        self._template = template
        self._printer = printer
        self._jobs = jobs              # (行号, 字段值, 张数)
        self._cancelled = False
        self.statuses: list[dict] = []

    def cancel(self):
        self._cancelled = True

    def run(self):
        ok, failed = 0, 0
        self.statuses = []
        try:
            self._setup()
            self.backend.open(self._template, self._printer)
            total = len(self._jobs)
            for i, (row_no, values, copies) in enumerate(self._jobs, 1):
                if self._cancelled:
                    self.done.emit(False, f"已取消（完成 {ok} 行，失败 {failed} 行）",
                                   self.statuses)
                    return
                label = next((v for v in values.values() if v), f"第 {row_no} 行")
                entry = {"row": row_no, "values": values, "copies": copies,
                         "status": "submitted", "error": ""}
                try:
                    self.backend.print_label(values, copies)
                    ok += 1
                    self.statuses.append(entry)
                    self.progress.emit(i, total, f"已发送：{label} × {copies} 张")
                except Exception as exc:
                    failed += 1
                    entry["status"] = "failed"
                    entry["error"] = str(exc)
                    self.statuses.append(entry)
                    self.log.emit(f"[失败] 第 {row_no} 行（{label}）：{exc}")
            msg = f"打印完成：成功 {ok} 行" + (f"，失败 {failed} 行" if failed else "")
            self.done.emit(failed == 0, msg, self.statuses)
        except Exception as exc:
            self.done.emit(False, f"打印中断：{exc}", self.statuses)
        finally:
            self._teardown()


class MainWindow(QMainWindow):
    def __init__(self, backend_factory, backend_name: str = ""):
        super().__init__()
        self._backend_factory = backend_factory
        self.sheet: SheetData | None = None
        self.template_fields: list[str] = []
        self.worker: PrintWorker | None = None
        self._scan_worker: TemplateScanWorker | None = None
        self._preview_worker: PreviewWorker | None = None
        self._scan_epoch = 0
        self._saved_printer = ""  # 当前模板 profile 里保存的打印机

        title = "标签自动打印工具"
        if backend_name:
            title += f" — {backend_name}"
        self.setWindowTitle(title)
        self.resize(1150, 850)
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
        row2.addWidget(QLabel("工作表："))
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentIndexChanged.connect(self._sheet_changed)
        row2.addWidget(self.sheet_combo, 1)
        row2.addWidget(QLabel("表头在第"))
        self.header_row = QSpinBox()
        self.header_row.setMinimum(1)
        self.header_row.setMaximum(50)
        self.header_row.setValue(1)
        self.header_row.setKeyboardTracking(False)
        row2.addWidget(self.header_row)
        row2.addWidget(QLabel("行"))
        btn_reload = QPushButton("重新读取")
        btn_reload.clicked.connect(self._reload_excel)
        row2.addWidget(btn_reload)
        row2.addStretch(1)
        form.addRow("选项：", row2)
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

        # --- 3. 字段映射 + 标签预览 ---
        mid_row = QHBoxLayout()
        map_box = QGroupBox("3. 字段映射（模板字段 ← 取哪一列 / 固定文本）")
        mlay = QVBoxLayout(map_box)
        self.map_table = QTableWidget(0, 3)
        self.map_table.setHorizontalHeaderLabels(["模板字段", "取值来源", "固定文本 / 列名"])
        self.map_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.map_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.map_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        mlay.addWidget(self.map_table)
        mid_row.addWidget(map_box, 3)

        preview_box = QGroupBox("标签预览（打印前核对）")
        pvlay = QVBoxLayout(preview_box)
        self.preview_label = QLabel("选择一行数据后点击「刷新预览」")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setMinimumSize(360, 200)
        self.preview_label.setStyleSheet("QLabel { background: #fafafa; border: 1px dashed #bbb; }")
        pvlay.addWidget(self.preview_label, 1)
        btn_preview = QPushButton("刷新预览（当前选中行）")
        btn_preview.clicked.connect(self._refresh_preview)
        pvlay.addWidget(btn_preview)
        mid_row.addWidget(preview_box, 2)
        layout.addLayout(mid_row, 1)

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
        self.data_table.itemSelectionChanged.connect(self._schedule_preview)
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
        self.btn_test = QPushButton("试打当前行")
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

        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(300)
        self._preview_timer.timeout.connect(self._refresh_preview)

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
            self, "选择 Excel 清单", str(resource_dir()),
            "Excel 文件 (*.xls *.xlsx)")
        if path:
            self._load_excel(path)

    def _snapshot_row_state(self) -> list[tuple[bool, int]]:
        """记住当前每行的勾选状态与张数，供重读后按行号恢复。"""
        state = []
        for r in range(self.data_table.rowCount()):
            chk = self.data_table.cellWidget(r, 0)
            spin = self.data_table.cellWidget(r, 1)
            if isinstance(chk, QCheckBox) and isinstance(spin, QSpinBox):
                state.append((chk.isChecked(), spin.value()))
        return state

    def _load_excel(self, path: str, header_row: int | None = None,
                    sheet_name: str | None = None):
        prev_state = self._snapshot_row_state()
        try:
            sheet = load_excel(path, sheet_name=sheet_name, header_row=header_row)
        except Exception as exc:
            QMessageBox.critical(self, "Excel 读取失败", str(exc))
            return
        self.sheet = sheet
        self.excel_path.setText(path)

        # 回填 sheet 下拉（屏蔽信号）
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        for name in sheet.sheet_names:
            self.sheet_combo.addItem(name, name)
        idx = self.sheet_combo.findText(sheet.sheet_name)
        if idx >= 0:
            self.sheet_combo.setCurrentIndex(idx)
        self.sheet_combo.blockSignals(False)
        self.sheet_combo.setEnabled(len(sheet.sheet_names) > 1)

        # 回填表头行（屏蔽信号，不再自动重载）
        self.header_row.blockSignals(True)
        self.header_row.setValue(sheet.header_row + 1)
        self.header_row.blockSignals(False)

        self._fill_data_table(prev_state)
        self._auto_map()
        msg = (f"已加载 Excel：{Path(path).name}（工作表「{sheet.sheet_name}」"
               f"{f'，共 {len(sheet.sheet_names)} 个工作表' if len(sheet.sheet_names) > 1 else ''}），"
               f"共 {len(sheet)} 行数据（表头第 {sheet.header_row + 1} 行）")
        self.log(msg)

    def _reload_excel(self):
        if self.excel_path.text():
            self._load_excel(self.excel_path.text(),
                             header_row=self.header_row.value() - 1,
                             sheet_name=self.sheet_combo.currentData())

    def _sheet_changed(self):
        if self.sheet is not None and self.excel_path.text():
            name = self.sheet_combo.currentData()
            if name and name != self.sheet.sheet_name:
                self._load_excel(self.excel_path.text(), sheet_name=name)

    def _fill_data_table(self, prev_state: list[tuple[bool, int]] | None = None):
        sheet = self.sheet
        assert sheet is not None
        cols = ["打印", "张数"] + sheet.columns
        self.data_table.setRowCount(len(sheet))
        self.data_table.setColumnCount(len(cols))
        self.data_table.setHorizontalHeaderLabels(cols)
        for r, row in enumerate(sheet.rows):
            chk = QCheckBox()
            checked, copies = (True, sheet.copies[r])
            if prev_state and r < len(prev_state):
                checked, copies = prev_state[r]
            chk.setChecked(checked)
            self.data_table.setCellWidget(r, 0, chk)
            spin = QSpinBox()
            spin.setMinimum(1)
            spin.setMaximum(999)
            spin.setValue(copies)
            self.data_table.setCellWidget(r, 1, spin)
            for c, val in enumerate(row):
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.data_table.setItem(r, 2 + c, item)
        self.data_table.resizeColumnsToContents()
        self._update_summary()

    def _update_summary(self):
        if self.sheet is None:
            self.data_summary.setText("未加载 Excel")
            return
        checked = sum(
            1 for r in range(self.data_table.rowCount())
            if isinstance(self.data_table.cellWidget(r, 0), QCheckBox)
            and self.data_table.cellWidget(r, 0).isChecked())
        total = sum(
            self.data_table.cellWidget(r, 1).value()
            for r in range(self.data_table.rowCount())
            if isinstance(self.data_table.cellWidget(r, 1), QSpinBox)
            and self.data_table.cellWidget(r, 0).isChecked())
        self.data_summary.setText(f"共 {len(self.sheet)} 行，勾选 {checked} 行 / 预计 {total} 张")

    def _set_all_checked(self, checked: bool):
        for r in range(self.data_table.rowCount()):
            w = self.data_table.cellWidget(r, 0)
            if isinstance(w, QCheckBox):
                w.setChecked(checked)
        self._update_summary()

    # ---------------------------------------------------------- 模板

    def _scan_templates(self):
        self.template_combo.blockSignals(True)
        self.template_combo.clear()
        tpl_dir = resource_dir() / "模板"
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
            self, "选择 BarTender 模板", str(resource_dir()),
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
        self._scan_epoch += 1
        self.template_fields = []
        self.map_table.setRowCount(0)
        self.log(f"正在读取模板字段：{Path(path).name} …")
        worker = TemplateScanWorker(self._backend_factory, path, self._scan_epoch, self)
        worker.fields_ready.connect(self._on_template_fields)
        worker.failed.connect(self._on_template_scan_failed)
        worker.finished.connect(worker.deleteLater)
        self._scan_worker = worker
        worker.start()

    def _on_template_fields(self, epoch: int, fields: list[str]):
        if epoch != self._scan_epoch:
            return  # 已切到别的模板，丢弃过期结果
        self.template_fields = fields
        path = self.template_combo.currentData()
        name = Path(path).name if path else ""
        if not fields:
            self.map_table.setRowCount(0)
            self.log(f"模板 {name} 中没有命名数据源。\n"
                     "请先在 BarTender 里给要动态填充的对象起“数据源名称”。")
            return
        self.log(f"模板 {name}：发现 {len(fields)} 个字段：" + "、".join(fields))

        # 模板切换时恢复 profile 保存的打印机
        saved = load_profile(path, profiles_dir()) if path else None
        if saved and saved.printer:
            self._saved_printer = saved.printer
            idx = self.printer_combo.findText(saved.printer)
            if idx >= 0:
                self.printer_combo.setCurrentIndex(idx)
            else:
                self.log(f"提示：上次使用的打印机「{saved.printer}」当前不存在，已用默认")
        else:
            self._saved_printer = ""

        self._fill_mapping_table(saved)

    def _on_template_scan_failed(self, epoch: int, message: str):
        if epoch != self._scan_epoch:
            return
        self.template_fields = []
        self.map_table.setRowCount(0)
        QMessageBox.critical(self, "模板读取失败", message)

    # ---------------------------------------------------------- 映射

    def _fill_mapping_table(self, saved: Profile | None = None):
        columns = self.sheet.columns if self.sheet else []
        if saved is None:
            path = self.template_combo.currentData()
            if path:
                saved = load_profile(path, profiles_dir())

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

    def _values_for_row(self, r: int) -> dict[str, str]:
        """按当前映射计算某一行的字段值。"""
        if self.sheet is None:
            return {}
        col_index = {c: i for i, c in enumerate(self.sheet.columns)}
        row = self.sheet.rows[r]
        values = {}
        for m in self._current_mappings():
            if m.source == SOURCE_COLUMN and m.value in col_index:
                values[m.field] = row[col_index[m.value]]
            elif m.source == SOURCE_FIXED:
                values[m.field] = m.value
        return values

    # ---------------------------------------------------------- 预览

    def _schedule_preview(self):
        self._preview_timer.start()

    def _current_row(self) -> int | None:
        r = self.data_table.currentRow()
        if 0 <= r < self.data_table.rowCount():
            return r
        return None

    def _refresh_preview(self):
        if self.sheet is None or not self.template_fields:
            return
        if not self.template_combo.currentData():
            return
        if self._preview_worker is not None and self._preview_worker.isRunning():
            return  # 上一张还在渲染，跳过本次
        r = self._current_row()
        if r is None:
            return
        values = self._values_for_row(r)
        out_dir = previews_dir()
        out_dir.mkdir(parents=True, exist_ok=True)
        for old in out_dir.glob("label_*.png"):
            old.unlink(missing_ok=True)
        out_path = str(out_dir / "preview.png")

        self.preview_label.setText("预览生成中…")
        worker = PreviewWorker(self._backend_factory,
                               self.template_combo.currentData(),
                               self.printer_combo.currentData() or "",
                               values, out_path, self)
        worker.preview_ready.connect(self._on_preview_ready)
        worker.failed.connect(self._on_preview_failed)
        worker.finished.connect(worker.deleteLater)
        self._preview_worker = worker
        worker.start()

    def _on_preview_ready(self, path: str):
        pix = QPixmap(path)
        if pix.isNull():
            self.preview_label.setText("预览文件无法显示")
            return
        self.preview_label.setPixmap(pix.scaled(
            self.preview_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))

    def _on_preview_failed(self, message: str):
        self.preview_label.setText(f"预览失败：{message}")
        self.log(f"[预览失败] {message}")

    # ---------------------------------------------------------- 打印

    def _refresh_printers(self):
        current = self.printer_combo.currentText()
        printers = self._list_printers_safe()
        self.printer_combo.clear()
        self.printer_combo.addItem("（使用模板默认打印机）", "")
        for p in printers:
            self.printer_combo.addItem(p, p)
        idx = self.printer_combo.findText(current)
        if idx >= 0:
            self.printer_combo.setCurrentIndex(idx)

    def _list_printers_safe(self) -> list[str]:
        """win32print 枚举（非 COM，主线程安全）；失败时用临时后端兜底。"""
        try:
            import win32print
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            return [p[2] for p in win32print.EnumPrinters(flags)]
        except Exception:
            pass
        try:
            backend = self._backend_factory()
            printers = backend.list_printers()
            backend.close()
            return printers
        except Exception:
            return []

    def _checked_rows(self) -> list[int]:
        rows = []
        for r in range(self.data_table.rowCount()):
            chk = self.data_table.cellWidget(r, 0)
            if isinstance(chk, QCheckBox) and chk.isChecked():
                rows.append(r)
        return rows

    def _build_jobs(self, test_only: bool) -> list[tuple[int, dict[str, str], int]] | None:
        """收集勾选行的打印任务；返回 None 表示输入有误（已提示用户）。"""
        if self.sheet is None:
            QMessageBox.warning(self, "缺少数据", "请先加载 Excel 清单")
            return None
        if not self.template_combo.currentData():
            QMessageBox.warning(self, "缺少模板", "请先选择 BarTender 模板")
            return None
        if not self.template_fields:
            QMessageBox.warning(self, "模板无字段",
                                "模板还没有读出命名数据源，请稍候或检查模板配置")
            return None

        checked = self._checked_rows()
        if test_only:
            current = self._current_row()
            if current is not None and current in checked:
                checked = [current]
            elif checked:
                checked = checked[:1]

        jobs = []
        for r in checked:
            spin: QSpinBox = self.data_table.cellWidget(r, 1)
            jobs.append((r + 1, self._values_for_row(r), spin.value()))
        if not jobs:
            QMessageBox.warning(self, "没有选中行", "请至少勾选一行数据")
            return None

        # 预检
        rows_data = [self.sheet.rows[r - 1] for r, _, _ in jobs]
        copies_list = [c for _, _, c in jobs]
        errors, warnings = validate_jobs(
            self._current_mappings(), self.sheet.columns, rows_data, copies_list)
        if errors:
            QMessageBox.critical(self, "预检未通过", "\n".join(errors))
            return None
        if warnings:
            text = "预检发现以下警告：\n\n" + "\n".join(f"· {w}" for w in warnings[:10])
            if len(warnings) > 10:
                text += f"\n· …共 {len(warnings)} 条"
            text += "\n\n是否继续打印？"
            if QMessageBox.question(self, "预检警告", text) != QMessageBox.Yes:
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

        # 打印前保存映射配置（失败只记日志，不阻断打印）
        try:
            save_profile(Profile(template, printer, self._current_mappings()),
                         profiles_dir())
        except OSError as exc:
            self.log(f"[警告] 映射配置保存失败（不影响本次打印）：{exc}")

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
        self.worker = PrintWorker(self._backend_factory, template, printer, jobs, self)
        self.worker.progress.connect(self._on_progress)
        self.worker.log.connect(self.log)
        self.worker.done.connect(self._on_done)
        self.worker.finished.connect(self.worker.deleteLater)
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

    def _on_done(self, success: bool, message: str, statuses: list):
        self._set_printing_state(False)
        self.log(("✔ " if success else "✘ ") + message)
        self._write_job_log(success, statuses)
        if not success:
            QMessageBox.warning(self, "打印结果", message)

    def _write_job_log(self, success: bool, statuses: list):
        """每次打印（含试打）写一份任务日志 JSON，便于查错与补打。"""
        try:
            out = logs_dir()
            out.mkdir(parents=True, exist_ok=True)
            path = out / f"{datetime.now():%Y%m%d_%H%M%S}.json"
            record = {
                "time": datetime.now().isoformat(timespec="seconds"),
                "success": success,
                "excel": self.excel_path.text(),
                "sheet": self.sheet.sheet_name if self.sheet else "",
                "template": self.template_combo.currentData() or "",
                "printer": self.printer_combo.currentData() or "",
                "rows": statuses,
            }
            path.write_text(json.dumps(record, ensure_ascii=False, indent=2),
                            encoding="utf-8")
            self.log(f"任务日志已保存：{path}")
        except OSError as exc:
            self.log(f"[警告] 任务日志写入失败：{exc}")

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
            if not self.worker.wait(10000):
                ret = QMessageBox.question(
                    self, "打印仍在进行",
                    "打印任务 10 秒内未能停止。\n"
                    "「继续等待」将保留窗口；「强制退出」可能残留 BarTender 后台进程。",
                    "继续等待", "强制退出")
                if ret == 0:
                    event.ignore()
                    return
        for w in (self._scan_worker, self._preview_worker):
            if w is not None and w.isRunning():
                w.wait(3000)
        super().closeEvent(event)
