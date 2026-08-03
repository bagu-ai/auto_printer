from label_printer.core.profiles import (
    SOURCE_COLUMN,
    SOURCE_FIXED,
    SOURCE_NONE,
    FieldMapping,
    Profile,
    guess_mapping,
    load_profile,
    profile_path_for,
    save_profile,
    validate_jobs,
)

TPL = r"C:\模板\中文版标签模板（60-30mm）.btw"


def make_profile(template=TPL):
    return Profile(
        template_path=template,
        printer="TOSHIBA B-EX4T1 (305 dpi)",
        mappings=[
            FieldMapping("发货新编号", SOURCE_COLUMN, "发货新编号"),
            FieldMapping("客户香型", SOURCE_COLUMN, "客户香型"),
            FieldMapping("备注", SOURCE_FIXED, "样品"),
            FieldMapping("没用的字段", SOURCE_NONE, ""),
        ],
    )


def test_save_and_load_roundtrip(tmp_path):
    profile = make_profile()
    save_profile(profile, tmp_path)
    loaded = load_profile(profile.template_path, tmp_path)
    assert loaded is not None
    assert loaded.printer == "TOSHIBA B-EX4T1 (305 dpi)"
    assert loaded.mapping_for("客户香型").value == "客户香型"
    assert loaded.mapping_for("备注").source == SOURCE_FIXED
    assert loaded.mapping_for("不存在") is None


def test_prefix_suffix_roundtrip(tmp_path):
    """前缀/后缀随 profile 保存并恢复。"""
    profile = Profile(
        template_path=TPL,
        mappings=[
            FieldMapping("编号", SOURCE_COLUMN, "发货新编号", prefix="编号："),
            FieldMapping("粒径", SOURCE_COLUMN, "发货粒径", prefix="粒径：", suffix="（实测）"),
            FieldMapping("备注", SOURCE_FIXED, "样品", suffix="·勿食用"),
        ])
    save_profile(profile, tmp_path)
    loaded = load_profile(TPL, tmp_path)
    assert loaded.mapping_for("编号").prefix == "编号："
    assert loaded.mapping_for("编号").suffix == ""
    assert loaded.mapping_for("粒径").prefix == "粒径："
    assert loaded.mapping_for("粒径").suffix == "（实测）"
    assert loaded.mapping_for("备注").suffix == "·勿食用"


def test_json_without_prefix_suffix_loads_defaults(tmp_path):
    """v1.2 及更早的 profile（无 prefix/suffix 键）按空前/后缀加载。"""
    path = profile_path_for(TPL, tmp_path)
    path.write_text(
        '{"template_path": "%s", "printer": "", '
        '"mappings": [{"field": "编号", "source": "column", "value": "发货新编号"}]}'
        % TPL.replace("\\", "\\\\"),
        encoding="utf-8")
    loaded = load_profile(TPL, tmp_path)
    m = loaded.mapping_for("编号")
    assert m.prefix == "" and m.suffix == ""


def test_render_composes_prefix_suffix():
    m = FieldMapping("编号", SOURCE_COLUMN, "发货新编号", prefix="编号：", suffix="（新）")
    assert m.render("TB-71717") == "编号：TB-71717（新）"
    assert FieldMapping("编号", SOURCE_COLUMN, "发货新编号").render("TB-1") == "TB-1"
    # 空值也照拼：前缀本身就是提示信息（预检会对空值另行告警）
    assert FieldMapping("编号", SOURCE_COLUMN, "x", prefix="编号：").render("") == "编号："


def test_load_missing_returns_none(tmp_path):
    assert load_profile(r"C:\模板\nope.btw", tmp_path) is None


def test_filename_sanitization(tmp_path):
    profile = Profile(template_path=r"C:\a\b\c?d*e.btw")
    path = save_profile(profile, tmp_path)
    assert "?" not in path.name and "*" not in path.name
    assert load_profile(profile.template_path, tmp_path) is not None


def test_same_name_different_dirs_no_collision(tmp_path):
    """不同目录的同名模板不再串配置。"""
    tpl_a = r"D:\客户A\中文版标签模板.btw"
    tpl_b = r"D:\客户B\中文版标签模板.btw"
    assert profile_path_for(tpl_a, tmp_path) != profile_path_for(tpl_b, tmp_path)

    pa = Profile(tpl_a, "打印机A", [FieldMapping("字段", SOURCE_COLUMN, "列A")])
    pb = Profile(tpl_b, "打印机B", [FieldMapping("字段", SOURCE_COLUMN, "列B")])
    save_profile(pa, tmp_path)
    save_profile(pb, tmp_path)
    assert load_profile(tpl_a, tmp_path).printer == "打印机A"
    assert load_profile(tpl_b, tmp_path).printer == "打印机B"


def test_legacy_profile_migrated(tmp_path):
    """v1.0 旧命名（仅模板名）加载后自动迁移为新命名。"""
    legacy = tmp_path / "中文版标签模板（60-30mm）.json"
    legacy.write_text(
        '{"template_path": "%s", "printer": "旧打印机", '
        '"mappings": [{"field": "发货新编号", "source": "column", "value": "发货新编号"}]}'
        % TPL.replace("\\", "\\\\"),
        encoding="utf-8")
    loaded = load_profile(TPL, tmp_path)
    assert loaded is not None
    assert loaded.printer == "旧打印机"
    # 已迁移：新路径存在，旧文件删除
    assert profile_path_for(TPL, tmp_path).exists()
    assert not legacy.exists()


def test_guess_mapping():
    cols = ["序号", "发货新编号", "客户香型", "发货粒径"]
    assert guess_mapping("发货新编号", cols).source == SOURCE_COLUMN
    assert guess_mapping("客户香型", cols).value == "客户香型"
    g = guess_mapping("粒径", cols)
    assert g.source == SOURCE_COLUMN and g.value == "发货粒径"
    assert guess_mapping("二维码", cols).source == SOURCE_NONE


def test_guess_mapping_strips_trailing_colons():
    """数据源名带全角/半角冒号尾巴（如「香型：」）也能自动匹配列。"""
    cols = ["序号", "客户香型", "发货粒径"]
    assert guess_mapping("香型：", cols).value == "客户香型"
    assert guess_mapping("粒径:", cols).value == "发货粒径"
    assert guess_mapping("：", cols).source == SOURCE_NONE


# ---------- validate_jobs ----------

COLS = ["序号", "发货新编号", "客户香型"]
MAPPINGS = [
    FieldMapping("发货新编号", SOURCE_COLUMN, "发货新编号"),
    FieldMapping("客户香型", SOURCE_COLUMN, "客户香型"),
]


def test_validate_ok():
    rows = [["1", "TB-1", "Mint"], ["2", "TB-2", "Lemon"]]
    errors, warnings = validate_jobs(MAPPINGS, COLS, rows, [3, 1])
    assert errors == []
    assert warnings == []


def test_validate_no_rows():
    errors, _ = validate_jobs(MAPPINGS, COLS, [], [])
    assert errors


def test_validate_empty_cells_warned():
    rows = [["1", "TB-1", "Mint"], ["2", "TB-2", ""]]
    errors, warnings = validate_jobs(MAPPINGS, COLS, rows, [1, 1])
    assert errors == []
    assert any("客户香型" in w and "1 行为空" in w for w in warnings)


def test_validate_duplicates_warned():
    rows = [["1", "TB-1", "Mint"], ["2", "TB-1", "Lemon"]]
    _, warnings = validate_jobs(MAPPINGS, COLS, rows, [1, 1])
    assert any("重复" in w for w in warnings)


def test_validate_unmapped_field_warned():
    mappings = MAPPINGS + [FieldMapping("二维码", SOURCE_NONE, "")]
    _, warnings = validate_jobs(mappings, COLS, [["1", "TB-1", "Mint"]], [1])
    assert any("二维码" in w for w in warnings)


def test_validate_bad_copies_is_error():
    errors, _ = validate_jobs(MAPPINGS, COLS, [["1", "TB-1", "Mint"]], [0])
    assert any("张数" in e for e in errors)
