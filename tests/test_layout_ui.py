import pandas as pd
import pytest
from PySide6.QtCore import QRect
from PySide6.QtWidgets import QApplication, QScrollArea

from label_printer.core.mock_backend import MockBackend
from label_printer.ui.main_window import MainWindow
from label_printer.ui.theme import apply_theme


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance() or QApplication([])
    apply_theme(app)
    return app


class FakeScreen:
    """模拟指定可用区域的屏幕（离屏默认 800x600 只会走小屏分支）。"""

    def __init__(self, w, h):
        self._rect = QRect(0, 0, w, h)

    def availableGeometry(self):
        return self._rect


def test_normal_screen_sets_central_widget(qapp, monkeypatch):
    """普通屏：root 直接是 centralWidget（修复漏 setCentralWidget 导致白屏）。"""
    monkeypatch.setattr(QApplication, "primaryScreen",
                        staticmethod(lambda: FakeScreen(1920, 1040)))
    win = MainWindow(MockBackend, MockBackend.name)
    assert win.centralWidget() is win._root_panel
    win.close()


def test_compact_screen_wraps_in_scroll_area(qapp, monkeypatch):
    """小屏（1366x768 @150% 逻辑约 910x480）：滚动容器保证打印条可达。"""
    monkeypatch.setattr(QApplication, "primaryScreen",
                        staticmethod(lambda: FakeScreen(910, 480)))
    win = MainWindow(MockBackend, MockBackend.name)
    central = win.centralWidget()
    assert isinstance(central, QScrollArea)
    assert central.widget() is win._root_panel
    assert win.minimumHeight() <= 480  # 窗口本身不得超出可用高度
    win.close()


def make_book(path, rows=20):
    pd.DataFrame({
        "序号": list(range(1, rows + 1)),
        "发货新编号": [f"TB-{i:05d}" for i in range(rows)],
        "客户香型": ["Sweet Watermelon" if i % 2 else "Lemon Tea Saffron" for i in range(rows)],
        "发货粒径": ["2.8mm"] * rows,
        "送样数量": ["3瓶"] * rows,
    }).to_excel(path, index=False)


def test_layout_prioritizes_data_table(qapp, tmp_path):
    path = tmp_path / "清单.xlsx"
    make_book(path)
    win = MainWindow(MockBackend, MockBackend.name)
    win.resize(1360, 860)
    win.show()
    win._load_excel(str(path))
    qapp.processEvents()
    assert win.data_table.viewport().height() >= 300
    assert not win.log_panel.is_expanded()
    assert win.main_splitter.sizes()[0] > win.main_splitter.sizes()[1]
    win.close()


def test_search_filters_without_changing_checks(qapp, tmp_path):
    path = tmp_path / "清单.xlsx"
    make_book(path, 4)
    win = MainWindow(MockBackend, MockBackend.name)
    win._load_excel(str(path))
    before = [win.data_table.cellWidget(r, 0).isChecked() for r in range(4)]
    win.search_edit.setText("Watermelon")
    qapp.processEvents()
    visible = sum(not win.data_table.isRowHidden(r) for r in range(4))
    after = [win.data_table.cellWidget(r, 0).isChecked() for r in range(4)]
    assert visible == 2
    assert before == after
    win.close()


def test_business_column_widths_are_bounded(qapp, tmp_path):
    path = tmp_path / "清单.xlsx"
    make_book(path, 2)
    win = MainWindow(MockBackend, MockBackend.name)
    win._load_excel(str(path))
    names = [win.data_table.horizontalHeaderItem(c).text()
             for c in range(win.data_table.columnCount())]
    assert win.data_table.columnWidth(names.index("打印")) <= 60
    assert 150 <= win.data_table.columnWidth(names.index("客户香型")) <= 220
    win.close()


def test_reload_reapplies_search_filter(qapp, tmp_path):
    """换文件后行隐藏状态必须按当前搜索词重建，且选中行必须可见。"""
    path_a = tmp_path / "A.xlsx"
    make_book(path_a, 4)
    win = MainWindow(MockBackend, MockBackend.name)
    win._load_excel(str(path_a))
    win.search_edit.setText("Watermelon")
    qapp.processEvents()
    assert sum(not win.data_table.isRowHidden(r) for r in range(4)) == 2

    path_b = tmp_path / "B.xlsx"
    pd.DataFrame({
        "序号": [1, 2, 3],
        "发货新编号": ["TB-100", "TB-101", "TB-102"],
        "客户香型": ["Sweet Watermelon", "Lemon Tea Saffron", "Sweet Watermelon"],
        "发货粒径": ["2.8mm"] * 3,
        "送样数量": ["3瓶"] * 3,
    }).to_excel(path_b, index=False)
    win._load_excel(str(path_b))
    qapp.processEvents()
    visible = [r for r in range(3) if not win.data_table.isRowHidden(r)]
    assert visible == [0, 2]  # 新数据按旧搜索词重新过滤
    assert win.data_table.currentRow() in visible  # 预览/试打不会落在隐藏行
    win.close()


def test_close_after_workers_finished_is_clean(qapp, tmp_path):
    """扫描 worker 销毁（deleteLater）后关窗不得抛 RuntimeError，设置正常保存。"""
    path = tmp_path / "清单.xlsx"
    make_book(path, 3)
    win = MainWindow(MockBackend, MockBackend.name)
    win.show()
    win._load_excel(str(path))
    if win._scan_worker is not None:
        win._scan_worker.wait(5000)
    qapp.processEvents()  # 让 finished 槽与 deleteLater 全部跑完
    qapp.processEvents()
    win.close()  # 修复前：Internal C++ object already deleted
    qapp.processEvents()


def test_stale_preview_results_are_dropped(qapp, tmp_path):
    """过期 epoch 的校样结果不得覆盖当前行；失败后不得残留旧校样路径。"""
    path = tmp_path / "清单.xlsx"
    make_book(path, 3)
    win = MainWindow(MockBackend, MockBackend.name)
    win._load_excel(str(path))
    win._preview_epoch = 5
    win._on_preview_ready(3, "不存在的旧图.png")  # 过期 epoch
    assert win.preview_canvas.source_pixmap.isNull()
    assert win._latest_preview_path == ""
    win._latest_preview_path = "上一张.png"
    win._on_preview_failed(5, "模拟失败")
    assert win._latest_preview_path == ""
    win.close()


def test_settings_are_isolated_in_tests(qapp, tmp_path):
    """conftest 的隔离夹具生效：主窗口的 QSettings 不落到真实用户注册表。"""
    from pathlib import Path
    win = MainWindow(MockBackend, MockBackend.name)
    assert Path(win._settings.fileName()).parent == tmp_path
    win.close()
