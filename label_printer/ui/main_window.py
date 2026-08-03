"""主界面：拖入 Excel → 选模板 → 配字段映射 → 勾选数据行 → 预览/预检 → 打印。

线程纪律：所有后端（BarTender COM / mock）操作都在 QThread 内完成，
后端实例由 worker 在线程内通过工厂创建、使用、销毁，绝不跨线程共享。
"""
from __future__ import annotations

import json
import shutil
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QThread, QTimer, Signal, QSettings, QSize
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QFileDialog, QFrame,
    QGridLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow,
    QMessageBox, QPlainTextEdit, QProgressBar, QPushButton, QScrollArea,
    QSizePolicy, QSpinBox, QSplitter, QTableWidget, QTableWidgetItem,
    QTabWidget, QVBoxLayout, QWidget, QAbstractItemView,
)

from ..core.backend import BackendError, PrintBackend
from ..core.excel_reader import SheetData, load_excel
from ..core.paths import logs_dir, previews_dir, profiles_dir, resource_dir
from ..core.profiles import (
    SOURCE_COLUMN, SOURCE_FIXED, SOURCE_NONE,
    FieldMapping, Profile, guess_mapping, load_profile, save_profile,
    validate_jobs,
)
from .widgets import CollapsiblePanel, PreviewCanvas, PreviewToolbar, StatusPill

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
        self._preview_epoch = 0
        self._preview_active_epoch = 0
        self._preview_pending = False
        self._latest_preview_path = ""
        self._saved_printer = ""  # 当前模板 profile 里保存的打印机
        self._backend_name = backend_name
        self._settings = QSettings("BG0", "LabelPrinter")
        self._log_count = 0

        title = "标签自动打印工具"
        if backend_name:
            title += f" — {backend_name}"
        self.setWindowTitle(title)
        self._compact_screen = False
        screen = QApplication.primaryScreen().availableGeometry()
        self.setAcceptDrops(True)

        self._build_ui()

        # 小屏（如 1366×768 @150% 缩放，逻辑约 910×480）：整体放进滚动区，
        # 保证底部打印操作条始终可达；普通屏直接定最小尺寸。
        if screen.width() < 1040 or screen.height() < 660:
            self._compact_screen = True
            self._root_panel.setMinimumSize(980, 620)
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setFrameShape(QFrame.NoFrame)
            scroll.setWidget(self._root_panel)
            self.setCentralWidget(scroll)
            self.setMinimumSize(min(940, screen.width()), min(470, screen.height()))
        else:
            self.setCentralWidget(self._root_panel)
            self.setMinimumSize(min(1080, screen.width()), min(700, screen.height()))
        self.resize(min(1440, int(screen.width() * .90)),
                    min(900, int(screen.height() * .90)))

        geometry = self._settings.value("windowGeometry")
        if geometry:
            self.restoreGeometry(geometry)
            # 恢复的几何可能来自更大/更高 DPI 的屏幕，夹回当前可用区域
            self.resize(min(self.width(), screen.width()),
                        min(self.height(), screen.height()))
        splitter_state = self._settings.value("mainSplitter")
        if splitter_state:
            self.main_splitter.restoreState(splitter_state)
        self.side_tabs.setCurrentIndex(int(self._settings.value("sideTab", 0)))
        self.log_panel.set_expanded(
            str(self._settings.value("logExpanded", "false")).lower() == "true")
        self._scan_templates()

    # ---------------------------------------------------------------- UI

    def _panel(self, object_name: str = "workPanel") -> tuple[QFrame, QVBoxLayout]:
        frame = QFrame()
        frame.setObjectName(object_name)
        lay = QVBoxLayout(frame)
        lay.setContentsMargins(14, 12, 14, 12)
        lay.setSpacing(9)
        return frame, lay

    def _section_title(self, title: str, hint: str = "") -> QHBoxLayout:
        row = QHBoxLayout()
        label = QLabel(title)
        label.setObjectName("sectionTitle")
        row.addWidget(label)
        if hint:
            h = QLabel(hint)
            h.setObjectName("sectionHint")
            row.addWidget(h)
        row.addStretch(1)
        return row

    def _build_ui(self):
        root = QWidget()
        root.setObjectName("appRoot")
        self._root_panel = root  # setCentralWidget 在确认屏幕尺寸后进行
        layout = QVBoxLayout(root)
        layout.setContentsMargins(12, 10, 12, 8)
        layout.setSpacing(8)

        # 顶部品牌与运行状态
        top, top_lay = self._panel("topBar")
        top_lay.setContentsMargins(16, 10, 16, 10)
        top_row = QHBoxLayout()
        title_col = QVBoxLayout()
        title_col.setSpacing(1)
        title = QLabel("标签自动打印")
        title.setObjectName("appTitle")
        subtitle = QLabel("Excel 数据校验 · BarTender 校样 · 批量打印")
        subtitle.setObjectName("appSubtitle")
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        top_row.addLayout(title_col)
        top_row.addStretch(1)
        self.runtime_status = StatusPill(
            f"{self._backend_name or '后端'} 已就绪",
            "success" if "干跑" not in self._backend_name else "warning")
        top_row.addWidget(self.runtime_status)
        self.top_summary = QLabel("等待载入数据")
        self.top_summary.setObjectName("sectionHint")
        top_row.addWidget(self.top_summary)
        top_lay.addLayout(top_row)
        layout.addWidget(top)

        # 紧凑输入区：Excel 与模板各一行
        inputs, grid = QFrame(), QGridLayout()
        inputs.setObjectName("inputPanel")
        inputs.setLayout(grid)
        grid.setContentsMargins(14, 11, 14, 11)
        grid.setHorizontalSpacing(9)
        grid.setVerticalSpacing(8)
        grid.addWidget(QLabel("Excel 数据"), 0, 0)
        self.excel_path = QLineEdit()
        self.excel_path.setPlaceholderText("拖入或选择 Excel 清单…")
        self.excel_path.setReadOnly(True)
        grid.addWidget(self.excel_path, 0, 1)
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(145)
        self.sheet_combo.currentIndexChanged.connect(self._sheet_changed)
        grid.addWidget(self.sheet_combo, 0, 2)
        self.header_row = QSpinBox()
        self.header_row.setRange(1, 50)
        self.header_row.setValue(1)
        self.header_row.setPrefix("表头第 ")
        self.header_row.setSuffix(" 行")
        self.header_row.setKeyboardTracking(False)
        self.header_row.setMinimumWidth(120)
        grid.addWidget(self.header_row, 0, 3)
        btn_reload = QPushButton("重新读取")
        btn_reload.clicked.connect(self._reload_excel)
        grid.addWidget(btn_reload, 0, 4)
        btn_excel = QPushButton("选择 Excel")
        btn_excel.clicked.connect(self._browse_excel)
        grid.addWidget(btn_excel, 0, 5)

        grid.addWidget(QLabel("标签模板"), 1, 0)
        self.template_combo = QComboBox()
        self.template_combo.currentIndexChanged.connect(self._template_changed)
        grid.addWidget(self.template_combo, 1, 1, 1, 4)
        btn_tpl = QPushButton("选择模板")
        btn_tpl.clicked.connect(self._browse_template)
        grid.addWidget(btn_tpl, 1, 5)
        grid.setColumnStretch(1, 1)
        layout.addWidget(inputs)

        # 主工作区：数据表优先，右侧是按需使用的校样/映射
        self.main_splitter = QSplitter(Qt.Horizontal)
        self.main_splitter.setChildrenCollapsible(False)
        self.main_splitter.setHandleWidth(6)

        data_panel, data_lay = self._panel()
        data_panel.setMinimumWidth(540)
        data_lay.addLayout(self._section_title("数据工作台", "勾选要打印的行，并核对每行张数"))
        data_tools = QHBoxLayout()
        btn_all = QPushButton("全选")
        btn_none = QPushButton("全不选")
        btn_all.setProperty("role", "compact")
        btn_none.setProperty("role", "compact")
        btn_all.clicked.connect(lambda: self._set_all_checked(True))
        btn_none.clicked.connect(lambda: self._set_all_checked(False))
        data_tools.addWidget(btn_all)
        data_tools.addWidget(btn_none)
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("搜索编号、香型或任意内容…")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setMaximumWidth(330)
        self.search_edit.textChanged.connect(self._filter_rows)
        data_tools.addWidget(self.search_edit)
        data_tools.addStretch(1)
        self.data_summary = QLabel("未加载 Excel")
        self.data_summary.setObjectName("sectionHint")
        data_tools.addWidget(self.data_summary)
        data_lay.addLayout(data_tools)

        self.data_table = QTableWidget(0, 0)
        self.data_table.setAlternatingRowColors(True)
        self.data_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.data_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.data_table.setTextElideMode(Qt.ElideRight)
        self.data_table.verticalHeader().setDefaultSectionSize(36)
        self.data_table.horizontalHeader().setMinimumHeight(36)
        self.data_table.itemSelectionChanged.connect(self._schedule_preview)
        self.data_table.setMinimumHeight(150)
        data_lay.addWidget(self.data_table, 1)
        self.main_splitter.addWidget(data_panel)

        self.side_tabs = QTabWidget()
        self.side_tabs.setMinimumWidth(320)
        preview_tab = QWidget()
        pvlay = QVBoxLayout(preview_tab)
        pvlay.setContentsMargins(10, 10, 10, 10)
        pvlay.setSpacing(8)
        pvlay.addLayout(self._section_title("标签校样", "打印前按像素核对"))
        self.preview_canvas = PreviewCanvas()
        self.preview_toolbar = PreviewToolbar(self.preview_canvas)
        self.preview_toolbar.openRequested.connect(self._open_preview_dialog)
        pvlay.addWidget(self.preview_toolbar)
        pvlay.addWidget(self.preview_canvas, 1)
        self.btn_preview = QPushButton("刷新当前行校样")
        self.btn_preview.clicked.connect(self._refresh_preview)
        pvlay.addWidget(self.btn_preview)
        self.side_tabs.addTab(preview_tab, "标签校样")

        mapping_tab = QWidget()
        mlay = QVBoxLayout(mapping_tab)
        mlay.setContentsMargins(10, 10, 10, 10)
        mlay.setSpacing(8)
        mlay.addLayout(self._section_title("字段映射", "模板字段 ← Excel 列 / 固定文本，可加前后缀"))
        self.map_table = QTableWidget(0, 5)
        self.map_table.setHorizontalHeaderLabels(
            ["模板字段", "取值来源", "固定文本 / 列名", "前缀", "后缀"])
        self.map_table.setAlternatingRowColors(True)
        self.map_table.verticalHeader().setDefaultSectionSize(36)
        self.map_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.map_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.map_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.map_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.map_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        mlay.addWidget(self.map_table)
        self.side_tabs.addTab(mapping_tab, "字段映射")
        self.main_splitter.addWidget(self.side_tabs)
        self.main_splitter.setStretchFactor(0, 65)
        self.main_splitter.setStretchFactor(1, 35)
        self.main_splitter.setSizes([900, 500])
        layout.addWidget(self.main_splitter, 1)

        # 固定打印操作条
        print_bar, play = self._panel("printBar")
        play.setContentsMargins(14, 9, 14, 9)
        prow = QHBoxLayout()
        prow.addWidget(QLabel("打印机"))
        self.printer_combo = QComboBox()
        self.printer_combo.setMinimumWidth(300)
        self.printer_combo.currentIndexChanged.connect(self._schedule_preview)
        self._refresh_printers()
        btn_printer = QPushButton("刷新")
        btn_printer.setProperty("role", "compact")
        btn_printer.clicked.connect(self._refresh_printers)
        prow.addWidget(self.printer_combo, 1)
        prow.addWidget(btn_printer)
        self.btn_test = QPushButton("试打当前行")
        self.btn_test.clicked.connect(lambda: self._start_print(test_only=True))
        self.btn_print = QPushButton("打印全部")
        self.btn_print.setProperty("role", "primary")
        self.btn_print.clicked.connect(lambda: self._start_print(test_only=False))
        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setProperty("role", "danger")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._cancel_print)
        prow.addWidget(self.btn_test)
        prow.addWidget(self.btn_print)
        prow.addWidget(self.btn_cancel)
        play.addLayout(prow)
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setValue(0)
        play.addWidget(self.progress)
        layout.addWidget(print_bar)

        # 运行记录默认收起，错误时自动展开
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setMaximumBlockCount(2000)
        self.log_view.setMaximumHeight(180)
        self.log_view.setPlaceholderText("运行记录…")
        self.log_panel = CollapsiblePanel("运行记录 · 0 条", self.log_view, False)
        layout.addWidget(self.log_panel)

        self.statusBar().showMessage("就绪 · 可拖入 Excel 或 .btw 模板")
        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(250)
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
        self._latest_preview_path = ""  # 数据已变，旧校样不可再打开
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
            chk.setAccessibleName(f"打印第 {r + 1} 行")
            checked, copies = (True, sheet.copies[r])
            if prev_state and r < len(prev_state):
                checked, copies = prev_state[r]
            chk.setChecked(checked)
            # clicked 只响应用户手动切换；批量全选/全不选在循环后统一刷新一次，
            # 避免 N 行每行触发一次全表扫描形成 O(N²)。
            chk.clicked.connect(lambda _checked: self._update_summary())
            self.data_table.setCellWidget(r, 0, chk)
            spin = QSpinBox()
            spin.setAccessibleName(f"第 {r + 1} 行张数")
            spin.setMinimum(1)
            spin.setMaximum(999)
            spin.setValue(copies)
            spin.valueChanged.connect(lambda _copies: self._update_summary())
            self.data_table.setCellWidget(r, 1, spin)
            for c, val in enumerate(row):
                item = QTableWidgetItem(val)
                item.setToolTip(val)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.data_table.setItem(r, 2 + c, item)
        self._apply_data_column_widths()
        # setRowHidden 挂在行号上、重建 item 不会清除：换文件/切表后必须
        # 按当前搜索词重新过滤，否则隐藏状态与搜索框内容不一致。
        self._filter_rows(self.search_edit.text())
        for r in range(self.data_table.rowCount()):
            if not self.data_table.isRowHidden(r):
                self.data_table.selectRow(r)
                break
        self._update_summary()

    def _apply_data_column_widths(self):
        """使用业务语义列宽，避免长文本把整个表格挤扁。"""
        header = self.data_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        widths = {"打印": 56, "张数": 76, "序号": 58}
        for c in range(self.data_table.columnCount()):
            name = self.data_table.horizontalHeaderItem(c).text()
            if name in widths:
                width = widths[name]
            elif "香型" in name or name == "英文":
                width = 190
            elif "编号" in name:
                width = 135
            elif "粒径" in name or "日期" in name:
                width = 112
            elif name.startswith("列"):
                width = 72
            else:
                width = 120
            self.data_table.setColumnWidth(c, width)

    def _filter_rows(self, text: str):
        needle = text.strip().lower()
        for r in range(self.data_table.rowCount()):
            matched = not needle
            if needle:
                for c in range(2, self.data_table.columnCount()):
                    item = self.data_table.item(r, c)
                    if item and needle in item.text().lower():
                        matched = True
                        break
            self.data_table.setRowHidden(r, not matched)

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
        summary = f"共 {len(self.sheet)} 行 · 勾选 {checked} 行 · 预计 {total} 张"
        self.data_summary.setText(summary)
        self.top_summary.setText(f"数据 {len(self.sheet)} 行  |  待打印 {checked} 行 / {total} 张")
        self.btn_print.setText(f"打印 {checked} 行 · {total} 张")
        self.btn_print.setEnabled(checked > 0 and not self._worker_running(self.worker))
        self.btn_print.setToolTip("" if checked else "请至少勾选一行数据")

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
        # 模板切换后旧校样即失效，避免「独立查看」打开上个模板的图
        self._latest_preview_path = ""
        self.preview_canvas.clear_preview("模板已切换，读取字段后将自动刷新校样")
        self.log(f"正在读取模板字段：{Path(path).name} …")
        worker = TemplateScanWorker(self._backend_factory, path, self._scan_epoch, self)
        worker.fields_ready.connect(self._on_template_fields)
        worker.failed.connect(self._on_template_scan_failed)
        worker.finished.connect(worker.deleteLater)
        # worker 销毁后属性置空，避免 closeEvent 等处对悬空 C++ 对象调 isRunning
        worker.finished.connect(self._clear_scan_worker)
        self._scan_worker = worker
        worker.start()

    def _clear_scan_worker(self):
        # 旧 worker 的 finished 可能晚于新 worker 启动，只清自己
        if self._scan_worker is self.sender():
            self._scan_worker = None

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
            # 前缀/后缀：加到每一行该字段取值上的固定文字（如「编号：」）
            prefix_edit = QLineEdit(mapping.prefix)
            suffix_edit = QLineEdit(mapping.suffix)
            prefix_edit.setPlaceholderText("（无）")
            suffix_edit.setPlaceholderText("（无）")

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
                # 留空的字段不会被填充，前后缀无意义，禁用以免误解
                prefix_edit.setEnabled(False)
                suffix_edit.setEnabled(False)

            combo.currentTextChanged.connect(
                lambda text, e=edit, p=prefix_edit, s=suffix_edit:
                self._mapping_source_changed(text, e, p, s))
            combo.currentTextChanged.connect(lambda _text: self._schedule_preview())
            edit.textChanged.connect(lambda _text: self._schedule_preview())
            prefix_edit.textChanged.connect(lambda _text: self._schedule_preview())
            suffix_edit.textChanged.connect(lambda _text: self._schedule_preview())
            self.map_table.setCellWidget(r, 1, combo)
            self.map_table.setCellWidget(r, 2, edit)
            self.map_table.setCellWidget(r, 3, prefix_edit)
            self.map_table.setCellWidget(r, 4, suffix_edit)
        self.map_table.resizeColumnsToContents()

    @staticmethod
    def _mapping_source_changed(text: str, edit: QLineEdit,
                                prefix_edit: QLineEdit, suffix_edit: QLineEdit):
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
        has_source = text != OPTION_NONE
        prefix_edit.setEnabled(has_source)
        suffix_edit.setEnabled(has_source)

    def _auto_map(self):
        if self.template_fields:
            self._fill_mapping_table()

    def _current_mappings(self) -> list[FieldMapping]:
        out = []
        for r in range(self.map_table.rowCount()):
            field_name = self.map_table.item(r, 0).text()
            combo: QComboBox = self.map_table.cellWidget(r, 1)
            edit: QLineEdit = self.map_table.cellWidget(r, 2)
            prefix_edit: QLineEdit = self.map_table.cellWidget(r, 3)
            suffix_edit: QLineEdit = self.map_table.cellWidget(r, 4)
            prefix, suffix = prefix_edit.text(), suffix_edit.text()
            src_text = combo.currentText()
            if src_text == OPTION_NONE:
                out.append(FieldMapping(field_name, SOURCE_NONE, ""))
            elif src_text == OPTION_FIXED:
                out.append(FieldMapping(field_name, SOURCE_FIXED, edit.text(), prefix, suffix))
            else:
                out.append(FieldMapping(field_name, SOURCE_COLUMN, src_text, prefix, suffix))
        return out

    def _values_for_row(self, r: int) -> dict[str, str]:
        """按当前映射计算某一行的字段值（含前缀/后缀组合）。"""
        if self.sheet is None:
            return {}
        col_index = {c: i for i, c in enumerate(self.sheet.columns)}
        row = self.sheet.rows[r]
        values = {}
        for m in self._current_mappings():
            if m.source == SOURCE_COLUMN and m.value in col_index:
                values[m.field] = m.render(row[col_index[m.value]])
            elif m.source == SOURCE_FIXED:
                values[m.field] = m.render(m.value)
        return values

    # ---------------------------------------------------------- 预览

    def _schedule_preview(self):
        self._preview_epoch += 1
        if hasattr(self, "_preview_timer"):
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
        if self._worker_running(self._preview_worker):
            self._preview_pending = True
            return
        r = self._current_row()
        if r is None:
            return
        epoch = self._preview_epoch
        self._preview_active_epoch = epoch
        self._preview_pending = False
        values = self._values_for_row(r)
        # 固定目录 + 渲染前清空：校样只保留最新一份，previews 目录不膨胀
        out_dir = previews_dir() / "current"
        out_dir.mkdir(parents=True, exist_ok=True)
        for old in out_dir.glob("*.png"):
            try:
                old.unlink()
            except OSError:
                pass
        for legacy in previews_dir().glob("request_*"):  # v1.2 早期版本的按次目录
            try:
                shutil.rmtree(legacy, ignore_errors=True)
            except OSError:
                pass
        out_path = str(out_dir / "preview.png")

        self._latest_preview_path = ""  # 渲染期间旧校样不可再打开
        self.preview_canvas.clear_preview(f"正在渲染第 {r + 1} 行校样…")
        self.statusBar().showMessage(f"正在渲染第 {r + 1} 行标签校样…")
        worker = PreviewWorker(self._backend_factory,
                               self.template_combo.currentData(),
                               self.printer_combo.currentData() or "",
                               values, out_path, self)
        worker.preview_ready.connect(
            lambda path, e=epoch: self._on_preview_ready(e, path))
        worker.failed.connect(
            lambda message, e=epoch: self._on_preview_failed(e, message))
        worker.finished.connect(worker.deleteLater)
        worker.finished.connect(self._on_preview_worker_finished)
        self._preview_worker = worker
        worker.start()

    def _on_preview_ready(self, epoch: int, path: str):
        if epoch != self._preview_epoch:
            return  # 过期校样不得覆盖用户当前选中的行
        pix = QPixmap(path)
        if pix.isNull():
            self.preview_canvas.clear_preview("校样文件无法显示")
            return
        self._latest_preview_path = path
        self.preview_canvas.set_pixmap(pix)
        self.preview_toolbar.set_source_info(pix)
        self.statusBar().showMessage("校样已更新 · 可用 100% 或独立查看核对细节", 5000)

    def _on_preview_failed(self, epoch: int, message: str):
        if epoch != self._preview_epoch:
            return
        self._latest_preview_path = ""  # 失败/失效后不得再打开上一张旧校样
        self.preview_canvas.clear_preview(f"校样生成失败\n{message}")
        self.log(f"[预览失败] {message}")
        self.log_panel.set_expanded(True)
        self.statusBar().showMessage("校样生成失败，已展开运行记录", 6000)

    def _on_preview_worker_finished(self):
        if self._preview_worker is self.sender():
            self._preview_worker = None
        if self._preview_pending or self._preview_active_epoch != self._preview_epoch:
            QTimer.singleShot(0, self._refresh_preview)

    def _open_preview_dialog(self):
        if not self._latest_preview_path or not Path(self._latest_preview_path).exists():
            QMessageBox.information(self, "暂无校样", "请先选择一行并生成标签校样。")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("标签校样 · 独立查看")
        dialog.resize(1000, 700)
        lay = QVBoxLayout(dialog)
        canvas = PreviewCanvas()
        canvas.set_pixmap(QPixmap(self._latest_preview_path))
        toolbar = PreviewToolbar(canvas)
        toolbar.open_button.hide()
        toolbar.set_source_info(canvas.source_pixmap)
        lay.addWidget(toolbar)
        lay.addWidget(canvas, 1)
        dialog.exec()

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
        if self._worker_running(self.worker):
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
        self.worker.finished.connect(self._clear_print_worker)
        self._set_printing_state(True)
        self.progress.setMaximum(len(jobs))
        self.progress.setValue(0)
        self.worker.start()

    def _cancel_print(self):
        if self.worker is not None:
            self.worker.cancel()
            self.log("正在取消…")

    def _clear_print_worker(self):
        if self.worker is self.sender():
            self.worker = None

    def _on_progress(self, current: int, total: int, message: str):
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        self.log(message)

    def _on_done(self, success: bool, message: str, statuses: list):
        self._set_printing_state(False)
        self.log(("✔ " if success else "✘ ") + message)
        self._write_job_log(success, statuses)
        self.runtime_status.set_state(
            "打印完成" if success else "打印有失败", "success" if success else "error")
        self.statusBar().showMessage(message, 8000)
        if not success:
            self.log_panel.set_expanded(True)
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
        self.btn_print.setEnabled(not printing and bool(self._checked_rows()))
        self.btn_test.setEnabled(not printing)
        self.btn_cancel.setEnabled(printing)
        self.runtime_status.set_state(
            "正在提交打印任务" if printing else f"{self._backend_name or '后端'} 已就绪",
            "warning" if printing else ("success" if "干跑" not in self._backend_name else "warning"))

    # ---------------------------------------------------------- 杂项

    def log(self, message: str):
        self.log_view.appendPlainText(message)
        self._log_count += 1
        self.log_panel.set_title(f"运行记录 · {self._log_count} 条")

    @staticmethod
    def _worker_running(worker) -> bool:
        """worker 可能已被 deleteLater 销毁（属性尚未置空），按未运行处理。"""
        if worker is None:
            return False
        try:
            return worker.isRunning()
        except RuntimeError:
            return False

    def closeEvent(self, event):
        if self._worker_running(self.worker):
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
            if self._worker_running(w):
                w.wait(3000)
        self._settings.setValue("windowGeometry", self.saveGeometry())
        self._settings.setValue("mainSplitter", self.main_splitter.saveState())
        self._settings.setValue("sideTab", self.side_tabs.currentIndex())
        self._settings.setValue("logExpanded", self.log_panel.is_expanded())
        super().closeEvent(event)
