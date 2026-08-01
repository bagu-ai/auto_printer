"""数据表汇总的 GUI 回归测试。"""
import pandas as pd
import pytest
from PySide6.QtWidgets import QApplication, QCheckBox, QSpinBox

from label_printer.core.mock_backend import MockBackend
from label_printer.ui.main_window import MainWindow


@pytest.fixture(scope="session")
def qapp():
    return QApplication.instance() or QApplication([])


def make_book(path):
    pd.DataFrame({
        "序号": [1, 2],
        "发货新编号": ["TB-1", "TB-2"],
        "客户香型": ["Mint", "Lemon"],
        "发货粒径": ["2.8mm", "3.0mm"],
        "送样数量": ["3瓶", "2瓶"],
    }).to_excel(path, index=False)


def test_manual_checkbox_updates_summary(qapp, tmp_path):
    path = tmp_path / "清单.xlsx"
    make_book(path)
    win = MainWindow(MockBackend, MockBackend.name)
    win._load_excel(str(path))
    assert "勾选 2 行 · 预计 5 张" in win.data_summary.text()

    check: QCheckBox = win.data_table.cellWidget(0, 0)
    check.click()  # 模拟用户鼠标/键盘操作（发出 clicked）
    qapp.processEvents()

    assert "勾选 1 行 · 预计 2 张" in win.data_summary.text()
    win.close()


def test_manual_copies_change_updates_summary(qapp, tmp_path):
    path = tmp_path / "清单.xlsx"
    make_book(path)
    win = MainWindow(MockBackend, MockBackend.name)
    win._load_excel(str(path))
    assert "勾选 2 行 · 预计 5 张" in win.data_summary.text()

    copies: QSpinBox = win.data_table.cellWidget(0, 1)
    copies.setValue(7)
    qapp.processEvents()

    assert "勾选 2 行 · 预计 9 张" in win.data_summary.text()
    win.close()
