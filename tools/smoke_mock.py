"""开发机自测：mock 后端 + 离屏渲染，跑通 Excel 加载 → 映射 → 预览 → 模拟打印全流程。

用法：python tools/smoke_mock.py
"""
import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
# 状态目录指向临时目录，不污染真实 APPDATA
_tmp_state = tempfile.mkdtemp(prefix="label_printer_smoke_")
os.environ["LABEL_PRINTER_HOME"] = _tmp_state
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PySide6.QtWidgets import QApplication, QComboBox

from label_printer.core.mock_backend import MockBackend
from label_printer.core.paths import logs_dir, previews_dir
from label_printer.ui.main_window import MainWindow, PrintWorker, PreviewWorker

SAMPLE_XLS = Path(__file__).resolve().parents[1] / "韩国-Samwha（GWG507）35个胶囊样品清单20260727.xls"
SAMPLE_BTW = Path(__file__).resolve().parents[1] / "模板" / "中文版无云南巴菰公司信息标签模板（60-30mm）.btw"

app = QApplication([])
win = MainWindow(MockBackend, MockBackend.name)

# 1. 加载 Excel
win._load_excel(str(SAMPLE_XLS))
assert win.sheet is not None and len(win.sheet) == 35, "Excel 加载失败"
first = dict(zip(win.sheet.columns, win.sheet.rows[0]))
assert first["生产日期"] == "2026-01-03", f"日期格式化错误: {first['生产日期']}"
print(f"[ok] Excel：{len(win.sheet)} 行，生产日期已格式化：{first['生产日期']}")

# 2. 模板字段（mock 后端，直接同步走扫描 worker 逻辑）
win._add_template(str(SAMPLE_BTW))
# TemplateScanWorker 是异步的，等线程结束后处理事件队列以投递信号
if win._scan_worker is not None:
    assert win._scan_worker.wait(10000), "模板扫描超时"
app.processEvents()
assert win.template_fields, "模板字段为空"
print(f"[ok] 模板字段：{win.template_fields}")

# 3. 检查自动映射
auto = {}
for r in range(win.map_table.rowCount()):
    name = win.map_table.item(r, 0).text()
    combo: QComboBox = win.map_table.cellWidget(r, 1)
    auto[name] = combo.currentText()
for f in ("发货新编号", "客户香型", "发货粒径"):
    assert auto.get(f) == f, f"字段 {f} 未自动映射：{auto.get(f)}"
print(f"[ok] 自动映射：{auto}")

# 4. 预览（真实起线程）
win.data_table.selectRow(0)
out = previews_dir() / "preview.png"
pw = PreviewWorker(MockBackend, str(SAMPLE_BTW), "TOSHIBA B-EX4T1 (305 dpi)",
                   win._values_for_row(0), str(out))
pw.start()
assert pw.wait(10000), "预览渲染超时"
assert out.exists() and out.stat().st_size > 100
print(f"[ok] 预览 PNG 已生成：{out}（{out.stat().st_size} 字节）")

# 5. 试打任务（当前选中行）
jobs = win._build_jobs(test_only=True)
assert jobs and len(jobs) == 1
row_no, values, copies = jobs[0]
assert values["发货新编号"] == "TB-70584", values
assert copies == 3, f"张数应为 3（10g*3瓶），实际 {copies}"
print(f"[ok] 试打任务：行 {row_no}，{copies} 张")

# 6. 全量打印（真实起线程，验证线程模型）
factory_calls = []
def factory():
    factory_calls.append(1)
    return MockBackend()

worker = PrintWorker(factory, str(SAMPLE_BTW), "TOSHIBA B-EX4T1 (305 dpi)", jobs := win._build_jobs(test_only=False))
assert jobs and len(jobs) == 35
done_result = {}
worker.done.connect(lambda ok, msg, statuses: done_result.update(ok=ok, msg=msg, n=len(statuses)))
worker.start()
assert worker.wait(30000), "打印线程超时"
app.processEvents()
backend = worker.backend
assert len(backend.printed) == 35, f"应打印 35 行，实际 {len(backend.printed)}"
total = sum(c for _, c in backend.printed)
assert backend.closed_count >= 1, "后端必须在线程内被 close"
print(f"[ok] 模拟打印完成：35 行，共 {total} 张（线程内创建后端：{len(factory_calls)} 次）")

# 7. 任务日志
win._write_job_log(True, worker.statuses)
logs = list(logs_dir().glob("*.json"))
assert logs, "任务日志未写入"
print(f"[ok] 任务日志：{logs[0].name}")

print("\n全部通过 ✔")
