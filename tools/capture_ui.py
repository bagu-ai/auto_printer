"""生成 v1.2 UI 审查截图（离屏也可运行）。"""
import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from PySide6.QtWidgets import QApplication
from label_printer.core.mock_backend import MockBackend
from label_printer.ui.main_window import MainWindow
from label_printer.ui.theme import apply_theme

app = QApplication([])
apply_theme(app)
win = MainWindow(MockBackend, MockBackend.name)
win.resize(1440, 900)
win.show()
win._load_excel(str(ROOT / "韩国-Samwha（GWG507）35个胶囊样品清单20260727.xls"))
if win._scan_worker:
    win._scan_worker.wait(10000)
app.processEvents()
win.data_table.selectRow(2)
app.processEvents()
path = ROOT / "ui_review_v1.2.png"
win.grab().save(str(path), "PNG")
print(path)
