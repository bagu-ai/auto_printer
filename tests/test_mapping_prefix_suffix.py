"""字段映射的前缀/后缀功能测试。

场景：客户希望每行标签在 Excel 取值前后加固定文字（如「编号：」「粒径：」），
这些固定内容不在 Excel 里。前缀/后缀随映射一起保存，并作用到校样与打印。
"""
import pandas as pd
import pytest
from PySide6.QtWidgets import QApplication, QComboBox, QLineEdit

from label_printer.core.mock_backend import MockBackend
from label_printer.core.profiles import Profile, load_profile, save_profile
from label_printer.ui.main_window import (
    OPTION_FIXED,
    OPTION_NONE,
    MainWindow,
)
from label_printer.ui.theme import apply_theme


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance() or QApplication([])
    apply_theme(app)
    return app


@pytest.fixture()
def win(qapp, tmp_path):
    path = tmp_path / "清单.xlsx"
    pd.DataFrame({
        "序号": [1, 2],
        "发货新编号": ["TB-71717", "TB-71718"],
        "客户香型": ["红酒", "Lemon Tea Saffron"],
        "发货粒径": ["3.0mm", "2.8mm"],
    }).to_excel(path, index=False)
    w = MainWindow(MockBackend, MockBackend.name)
    w._load_excel(str(path))
    # 不经过扫描 worker，直接注入字段并用空 profile 避免读真实配置
    w.template_fields = ["编号", "香型", "粒径"]
    w._fill_mapping_table(saved=Profile(template_path="test.btw"))
    yield w
    w.close()


def _row_widgets(win, r):
    return (win.map_table.cellWidget(r, 1),   # 来源 combo
            win.map_table.cellWidget(r, 2),   # 列名/固定文本
            win.map_table.cellWidget(r, 3),   # 前缀
            win.map_table.cellWidget(r, 4))   # 后缀


def test_prefix_suffix_composed_into_row_values(win):
    # 自动映射：编号→发货新编号、香型→客户香型、粒径→发货粒径
    _row_widgets(win, 0)[2].setText("编号：")
    _row_widgets(win, 1)[2].setText("香型：")
    _row_widgets(win, 2)[2].setText("粒径：")
    _row_widgets(win, 2)[3].setText("（样品）")
    assert win._values_for_row(0) == {
        "编号": "编号：TB-71717",
        "香型": "香型：红酒",
        "粒径": "粒径：3.0mm（样品）",
    }
    # 每一行都会加（第 2 行）
    assert win._values_for_row(1)["编号"] == "编号：TB-71718"


def test_prefix_suffix_apply_to_fixed_text(win):
    combo, edit, prefix_edit, suffix_edit = _row_widgets(win, 0)
    combo.setCurrentText(OPTION_FIXED)
    edit.setText("样品")
    prefix_edit.setText("【")
    suffix_edit.setText("】")
    assert win._values_for_row(0)["编号"] == "【样品】"


def test_prefix_suffix_disabled_when_unmapped(win):
    combo, _edit, prefix_edit, suffix_edit = _row_widgets(win, 0)
    assert prefix_edit.isEnabled() and suffix_edit.isEnabled()
    combo.setCurrentText(OPTION_NONE)
    assert not prefix_edit.isEnabled() and not suffix_edit.isEnabled()
    # 留空字段不产出值
    assert "编号" not in win._values_for_row(0)
    combo.setCurrentText("发货新编号")
    assert prefix_edit.isEnabled() and suffix_edit.isEnabled()


def test_prefix_suffix_saved_and_restored(win, tmp_path, monkeypatch):
    monkeypatch.setattr("label_printer.ui.main_window.profiles_dir", lambda: tmp_path)
    _row_widgets(win, 0)[2].setText("编号：")
    _row_widgets(win, 2)[2].setText("粒径：")
    _row_widgets(win, 2)[3].setText("（实测）")

    tpl = r"C:\模板\测试模板.btw"
    save_profile(Profile(tpl, "", win._current_mappings()), tmp_path)
    loaded = load_profile(tpl, tmp_path)

    # 重新填表（模拟下次启动）后前后缀应恢复
    win._fill_mapping_table(saved=loaded)
    assert _row_widgets(win, 0)[2].text() == "编号："
    assert _row_widgets(win, 2)[2].text() == "粒径："
    assert _row_widgets(win, 2)[3].text() == "（实测）"
    assert win._values_for_row(0)["编号"] == "编号：TB-71717"


def test_current_mappings_carry_prefix_suffix(win):
    _row_widgets(win, 1)[2].setText("香型：")
    mappings = {m.field: m for m in win._current_mappings()}
    assert mappings["香型"].prefix == "香型："
    assert mappings["香型"].suffix == ""
    assert mappings["编号"].prefix == ""
