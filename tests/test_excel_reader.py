import pytest

from label_printer.core.excel_reader import load_excel, parse_copies

SAMPLE_XLS = r"E:\AI_EX\BG0\auto_printer\韩国-Samwha（GWG507）35个胶囊样品清单20260727.xls"


def test_header_autodetect_and_columns():
    data = load_excel(SAMPLE_XLS)
    assert data.header_row == 2
    for col in ("发货新编号", "客户香型", "发货粒径", "送样数量"):
        assert col in data.columns


def test_row_count_and_first_row():
    data = load_excel(SAMPLE_XLS)
    assert len(data) == 35
    first = dict(zip(data.columns, data.rows[0]))
    assert first["发货新编号"] == "TB-70584"
    assert first["客户香型"] == "Sweet Watermelon"
    assert first["发货粒径"] == "2.8mm"


def test_copies_default_from_quantity_column():
    data = load_excel(SAMPLE_XLS)
    assert data.copies_column == "送样数量"
    # 示例清单里送样数量都是 "10g*3瓶" → 3 张
    assert data.copies[0] == 3
    assert all(c >= 1 for c in data.copies)


def test_explicit_header_row():
    data = load_excel(SAMPLE_XLS, header_row=2)
    assert len(data) == 35
    with pytest.raises(Exception):
        load_excel(SAMPLE_XLS, header_row=999)


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
