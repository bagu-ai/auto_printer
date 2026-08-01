import datetime as dt
from pathlib import Path

import pandas as pd
import pytest

from label_printer.core.excel_reader import (
    format_cell, list_sheet_names, load_excel, parse_copies,
)

SAMPLE_XLS = Path(__file__).resolve().parents[1] / "韩国-Samwha（GWG507）35个胶囊样品清单20260727.xls"

HEADERS = ["序号", "发货新编号", "客户香型", "发货粒径", "送样数量", "生产日期"]


def make_xlsx(path: Path, sheets: dict[str, list[list]]) -> Path:
    """在临时目录造一份测试工作簿。sheets: {sheet名: 所有行(含表头)}。"""
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, rows in sheets.items():
            df = pd.DataFrame(rows[1:], columns=rows[0])
            df.to_excel(writer, sheet_name=name, index=False)
    return path


@pytest.fixture
def simple_xlsx(tmp_path):
    return make_xlsx(tmp_path / "清单.xlsx", {
        "Sheet1": [HEADERS,
                   [1, "TB-001", "Mint", "2.8mm", "10g*3瓶", dt.datetime(2026, 1, 3)],
                   [2, "TB-002", "Lemon", "3.0mm", "5g*2瓶", dt.datetime(2026, 2, 1)]],
    })


# ---------- format_cell ----------

@pytest.mark.parametrize("value,expected", [
    (dt.datetime(2026, 1, 3), "2026-01-03"),
    (dt.datetime(2026, 1, 3, 14, 30), "2026-01-03 14:30"),
    (dt.date(2026, 1, 3), "2026-01-03"),
    (pd.Timestamp("2026-06-05"), "2026-06-05"),
    (3.0, "3"),
    (2.8, "2.8"),
    ("  TB-70584 ", "TB-70584"),
    (None, ""),
    (float("nan"), ""),
    (pd.NaT, ""),
])
def test_format_cell(value, expected):
    assert format_cell(value) == expected


# ---------- load_excel ----------

def test_load_simple_xlsx(simple_xlsx):
    data = load_excel(str(simple_xlsx))
    assert data.sheet_name == "Sheet1"
    assert data.sheet_names == ["Sheet1"]
    assert len(data) == 2
    first = dict(zip(data.columns, data.rows[0]))
    assert first["发货新编号"] == "TB-001"
    assert first["生产日期"] == "2026-01-03"   # 不是 2026-01-03 00:00:00
    assert data.copies == [3, 2]


def test_multi_sheet(tmp_path):
    path = make_xlsx(tmp_path / "多表.xlsx", {
        "客户A": [HEADERS, [1, "TB-100", "A", "2.8mm", "3瓶", dt.date(2026, 1, 1)]],
        "客户B": [HEADERS, [9, "TB-200", "B", "3.0mm", "1瓶", dt.date(2026, 1, 2)],
                           [10, "TB-201", "C", "3.0mm", "2瓶", dt.date(2026, 1, 3)]],
    })
    assert list_sheet_names(str(path)) == ["客户A", "客户B"]
    a = load_excel(str(path))
    assert len(a) == 1 and a.rows[0][1] == "TB-100"
    b = load_excel(str(path), sheet_name="客户B")
    assert len(b) == 2 and b.rows[0][1] == "TB-200"
    with pytest.raises(ValueError, match="不存在"):
        load_excel(str(path), sheet_name="客户C")


def test_header_autodetect_with_title_rows(tmp_path):
    rows = [["某公司样品清单", "", "", "", "", ""],
            ["", "", "", "", "", ""],
            HEADERS,
            [1, "TB-1", "X", "2.8mm", "3瓶", dt.date(2026, 1, 1)]]
    path = make_xlsx(tmp_path / "带标题.xlsx", {"S": rows})
    data = load_excel(str(path))
    assert data.header_row == 2
    assert len(data) == 1


# ---------- 真实样例（文件存在才跑，不硬依赖本机路径） ----------

@pytest.mark.skipif(not SAMPLE_XLS.exists(), reason="真实样例文件不在本机")
def test_real_sample():
    data = load_excel(str(SAMPLE_XLS))
    assert data.header_row == 2
    assert len(data) == 35
    first = dict(zip(data.columns, data.rows[0]))
    assert first["发货新编号"] == "TB-70584"
    assert first["客户香型"] == "Sweet Watermelon"
    assert first["发货粒径"] == "2.8mm"
    assert first["生产日期"] == "2026-01-03"
    assert data.copies[0] == 3


# ---------- parse_copies ----------

@pytest.mark.parametrize("text,expected", [
    ("10g*3瓶", 3),
    ("3瓶", 3),
    ("5g*10瓶", 10),
    ("10g*2 瓶", 2),
    ("x4", 4),
    ("*5", 5),
    ("7", 7),
    ("", 1),
    ("10g", 1),
    ("若干", 1),
])
def test_parse_copies(text, expected):
    assert parse_copies(text) == expected
