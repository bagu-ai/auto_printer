from label_printer.core.profiles import (
    SOURCE_COLUMN,
    SOURCE_FIXED,
    SOURCE_NONE,
    FieldMapping,
    Profile,
    guess_mapping,
    load_profile,
    save_profile,
)


def test_save_and_load_roundtrip(tmp_path):
    profile = Profile(
        template_path=r"C:\模板\中文版标签模板（60-30mm）.btw",
        printer="TOSHIBA B-EX4T1 (305 dpi)",
        mappings=[
            FieldMapping("发货新编号", SOURCE_COLUMN, "发货新编号"),
            FieldMapping("客户香型", SOURCE_COLUMN, "客户香型"),
            FieldMapping("备注", SOURCE_FIXED, "样品"),
            FieldMapping("没用的字段", SOURCE_NONE, ""),
        ],
    )
    save_profile(profile, tmp_path)
    loaded = load_profile(profile.template_path, tmp_path)
    assert loaded is not None
    assert loaded.printer == "TOSHIBA B-EX4T1 (305 dpi)"
    assert loaded.mapping_for("客户香型").value == "客户香型"
    assert loaded.mapping_for("备注").source == SOURCE_FIXED
    assert loaded.mapping_for("备注").value == "样品"
    assert loaded.mapping_for("不存在") is None


def test_load_missing_returns_none(tmp_path):
    assert load_profile(r"C:\模板\nope.btw", tmp_path) is None


def test_filename_sanitization(tmp_path):
    profile = Profile(template_path=r"C:\a\b\c?d*e.btw")
    path = save_profile(profile, tmp_path)
    assert "?" not in path.name and "*" not in path.name
    assert load_profile(profile.template_path, tmp_path) is not None


def test_guess_mapping():
    cols = ["序号", "发货新编号", "客户香型", "发货粒径"]
    assert guess_mapping("发货新编号", cols).source == SOURCE_COLUMN
    assert guess_mapping("客户香型", cols).value == "客户香型"
    # 包含匹配：模板字段叫"粒径"，Excel 列叫"发货粒径"
    g = guess_mapping("粒径", cols)
    assert g.source == SOURCE_COLUMN and g.value == "发货粒径"
    # 完全匹配不上 → 留空
    assert guess_mapping("二维码", cols).source == SOURCE_NONE
