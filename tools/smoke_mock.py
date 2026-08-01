"""开发机自测：mock 后端 + 离屏渲染，跑通 Excel 加载 → 映射 → 模拟打印全流程。

用法：python tools/smoke_mock.py
"""
import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PySide6.QtWidgets import QApplication, QComboBox, QLineEdit

from label_printer.core.mock_backend import MockBackend
from label_printer.core.profiles import SOURCE_COLUMN
from label_printer.ui.main_window import MainWindow

SAMPLE_XLS = r"E:\AI_EX\BG0\auto_printer\韩国-Samwha（GWG507）35个胶囊样品清单20260727.xls"
SAMPLE_BTW = r"E:\AI_EX\BG0\auto_printer\模板\中文版无云南巴菰公司信息标签模板（60-30mm）.btw"

app = QApplication([])
backend = MockBackend()
win = MainWindow(backend)

# 1. 加载 Excel
win._load_excel(SAMPLE_XLS)
assert win.sheet is not None and len(win.sheet) == 35, "Excel 加载失败"
print(f"[ok] Excel：{len(win.sheet)} 行，列：{win.sheet.columns[:6]}...")

# 2. 加载模板（mock 返回示例字段）
win._add_template(SAMPLE_BTW)
assert win.template_fields, "模板字段为空"
print(f"[ok] 模板字段：{win.template_fields}")

# 3. 检查自动映射：发货新编号/客户香型/发货粒径 应自动对应同名 Excel 列
auto = {}
for r in range(win.map_table.rowCount()):
    name = win.map_table.item(r, 0).text()
    combo: QComboBox = win.map_table.cellWidget(r, 1)
    auto[name] = combo.currentText()
for f in ("发货新编号", "客户香型", "发货粒径"):
    assert auto.get(f) == f, f"字段 {f} 未自动映射：{auto.get(f)}"
print(f"[ok] 自动映射：{auto}")

# 4. 构建打印任务（第一行试打）
jobs = win._build_jobs(test_only=True)
assert jobs and len(jobs) == 1
row_no, values, copies = jobs[0]
assert values["发货新编号"] == "TB-70584", values
assert copies == 3, f"张数应为 3（10g*3瓶），实际 {copies}"
print(f"[ok] 试打任务：行 {row_no}，{copies} 张，值 {values}")

# 5. 全量任务 + 同步执行打印线程
jobs = win._build_jobs(test_only=False)
assert len(jobs) == 35
worker = win.worker = None
from label_printer.ui.main_window import PrintWorker
w = PrintWorker(backend, SAMPLE_BTW, "TOSHIBA B-EX4T1 (305 dpi)", jobs)
w.run()  # 同步执行，不开线程
assert len(backend.printed) == 35, f"应打印 35 行，实际 {len(backend.printed)}"
total = sum(c for _, c in backend.printed)
print(f"[ok] 模拟打印完成：35 行，共 {total} 张")
print("[ok] 最后一行内容：", backend.printed[-1])

print("\n全部通过 ✔")
