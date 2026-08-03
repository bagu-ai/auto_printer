"""截取说明书用图：真实 BarTender 后端 + 真实数据，按部件抓取（非离屏）。

用法：python tools/capture_manual_shots.py
输出：docs/images/*.png
"""
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

# 必须先于 MainWindow（其 _refresh_printers 会 import win32print）：
# win32print 先导入会导致 worker 线程 import win32com 失败（pywin32 加载顺序坑）
import pywintypes  # noqa: F401,E402
import win32com.client  # noqa: F401,E402
import win32print  # noqa: F401,E402

from PySide6.QtWidgets import QApplication, QMessageBox  # noqa: E402


def wait_ms(app, ms):
    """替代 QTest.qWait（本机 QtTest 的 DLL 会干扰 COM 线程）。"""
    t0 = time.time()
    while (time.time() - t0) * 1000 < ms:
        app.processEvents()
        time.sleep(0.02)


from label_printer.core.bartender_backend import BarTenderBackend  # noqa: E402
from label_printer.ui.main_window import MainWindow, PreviewWorker  # noqa: E402
from label_printer.ui.theme import apply_theme  # noqa: E402


def _ensure_app_fresh(self):
    """本机有一个杀不掉的半死 BarTender 服务进程；截图用 DispatchEx 强制新实例。"""
    if self._app is None:
        self._app = self._client.DispatchEx("BarTender.Application")
    return self._app


BarTenderBackend._ensure_app = _ensure_app_fresh

# 截图不需要后台预览线程（本机冷启动慢且多个 worker 排队）：主线程直接渲染
# 一张真实校样后注入画布，预览 worker 全程禁用。
PreviewWorker.start = lambda self: None

_orig_run = PreviewWorker.run


def _noisy_run(self):
    print("[pw] run start", flush=True)
    _orig_run(self)
    print("[pw] run end", flush=True)


PreviewWorker.run = _noisy_run

# 截图脚本无人值守：屏蔽一切模态弹窗，避免阻塞
QMessageBox.critical = staticmethod(lambda *a, **k: print("[box:critical]", a[2] if len(a) > 2 else ""))
QMessageBox.warning = staticmethod(lambda *a, **k: print("[box:warning]", a[2] if len(a) > 2 else ""))
QMessageBox.information = staticmethod(lambda *a, **k: None)
QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.Yes)

OUT = ROOT / "docs" / "images"
OUT.mkdir(parents=True, exist_ok=True)

app = QApplication(sys.argv)
apply_theme(app)
win = MainWindow(BarTenderBackend, BarTenderBackend.name)
win.resize(1500, 920)
win.show()
wait_ms(app, 800)
print("[step] window shown", flush=True)

excel = next(ROOT.glob("韩国-Samwha*.xls*"))
win._load_excel(str(excel))
print("[step] excel loaded", flush=True)

# 等模板字段扫描完成
t0 = time.time()
while not win.template_fields and time.time() - t0 < 20:
    app.processEvents()
    time.sleep(0.1)
print("[step] template fields:", win.template_fields, flush=True)

# 说明书展示「前缀/后缀」功能：给三个字段加固定前缀，校样与映射图同步呈现
# （字段名可能带全角冒号尾巴，如「香型：」，按去冒号后的名字匹配；
#   粒径改映射到短值的「发货粒径」列——「粒径」列带括号备注，加前缀后超出标签）
for r in range(win.map_table.rowCount()):
    field_name = win.map_table.item(r, 0).text().strip("：:")
    if field_name == "粒径":
        combo = win.map_table.cellWidget(r, 1)
        if combo.findText("发货粒径") >= 0:
            combo.setCurrentText("发货粒径")
    prefix = {"编号": "编号：", "香型": "香型：", "粒径": "粒径："}.get(field_name)
    if prefix:
        win.map_table.cellWidget(r, 3).setText(prefix)

# 第 9 行（索引 8）香型为短文本 "Mint"：加前缀后也不会超出标签边界，
# 说明书插图避免展示文字被裁剪的"残次"效果
win.data_table.selectRow(8)

# 截图期间彻底停用后台预览刷新（定时器 250ms 后会清掉注入的图）
win._preview_timer.stop()
win._schedule_preview = lambda: None
win._refresh_preview = lambda: None
win.printer_combo.setCurrentIndex(0)  # 截图统一显示「使用模板默认打印机」

# 主线程直接渲染第 3 行的真实校样并注入画布（不经 PreviewWorker）
import pythoncom  # noqa: E402
from PySide6.QtGui import QPixmap  # noqa: E402
pythoncom.CoInitialize()
try:
    backend = BarTenderBackend()
    backend.open(win.template_combo.currentData(),
                 win.printer_combo.currentData() or "")
    import tempfile  # noqa: E402
    shot_path = backend.render_preview(
        win._values_for_row(8),
        str(Path(tempfile.mkdtemp(prefix="manual_shot_")) / "_preview_src.png"))
    backend.close()
finally:
    pythoncom.CoUninitialize()
pix = QPixmap(shot_path)
win._latest_preview_path = shot_path
win.preview_canvas.set_pixmap(pix)
win.preview_toolbar.set_source_info(pix)
win.statusBar().showMessage("校样已更新 · 可用 100% 或独立查看核对细节")
win.log_panel.set_expanded(False)  # 主界面图用默认的收起状态
wait_ms(app, 400)
print("[step] preview injected:", not pix.isNull(), flush=True)

win.grab().save(str(OUT / "main_window.png"), "PNG")
win.side_tabs.grab().save(str(OUT / "preview_panel.png"), "PNG")

# 打印操作条
print_bar = win.btn_print.parentWidget()
while print_bar is not None and print_bar.objectName() != "printBar":
    print_bar = print_bar.parentWidget()
if print_bar is not None:
    print_bar.grab().save(str(OUT / "print_bar.png"), "PNG")

# 字段映射页（放宽右侧面板，让 5 列内容清晰可读）
win.side_tabs.setCurrentIndex(1)
win.main_splitter.setSizes([620, 880])
wait_ms(app, 400)
win.side_tabs.grab().save(str(OUT / "mapping_panel.png"), "PNG")

# 运行记录展开状态
win.side_tabs.setCurrentIndex(0)
win.main_splitter.setSizes([900, 500])
win.log_panel.set_expanded(True)
wait_ms(app, 400)
win.log_panel.grab().save(str(OUT / "log_panel.png"), "PNG")

win.close()
print("saved:", [p.name for p in sorted(OUT.glob('*.png'))])
