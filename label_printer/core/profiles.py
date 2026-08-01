"""映射配置：每个模板的「模板字段 ← Excel 列 / 固定文本」对应关系存成 JSON。"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

# 字段映射的三种来源
SOURCE_NONE = "none"        # 留空（打印空白）
SOURCE_COLUMN = "column"    # 取 Excel 某列
SOURCE_FIXED = "fixed"      # 固定文本


@dataclass
class FieldMapping:
    field: str                 # 模板字段名（BarTender 命名数据源）
    source: str = SOURCE_NONE  # none / column / fixed
    value: str = ""            # source=column 时是列名；source=fixed 时是固定文本


@dataclass
class Profile:
    template_path: str
    printer: str = ""                      # 空 = 用模板里保存的打印机
    mappings: list[FieldMapping] = field(default_factory=list)

    def mapping_for(self, field_name: str) -> FieldMapping | None:
        for m in self.mappings:
            if m.field == field_name:
                return m
        return None


def profile_path_for(template_path: str, profiles_dir: Path) -> Path:
    """按模板文件名生成 profile 路径（同名模板共享一份配置）。"""
    stem = Path(template_path).stem
    safe = "".join(c if c not in '\\/:*?"<>|' else "_" for c in stem)
    return profiles_dir / f"{safe}.json"


def save_profile(profile: Profile, profiles_dir: Path) -> Path:
    profiles_dir.mkdir(parents=True, exist_ok=True)
    path = profile_path_for(profile.template_path, profiles_dir)
    data = {
        "template_path": profile.template_path,
        "printer": profile.printer,
        "mappings": [
            {"field": m.field, "source": m.source, "value": m.value}
            for m in profile.mappings
        ],
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def load_profile(template_path: str, profiles_dir: Path) -> Profile | None:
    path = profile_path_for(template_path, profiles_dir)
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    mappings = [
        FieldMapping(
            field=m.get("field", ""),
            source=m.get("source", SOURCE_NONE),
            value=m.get("value", ""),
        )
        for m in data.get("mappings", [])
    ]
    return Profile(
        template_path=data.get("template_path", template_path),
        printer=data.get("printer", ""),
        mappings=mappings,
    )


def guess_mapping(field_name: str, columns: list[str]) -> FieldMapping:
    """按名字自动猜映射：精确匹配 > 包含匹配 > 留空。"""
    name = field_name.strip()
    if name in columns:
        return FieldMapping(field=field_name, source=SOURCE_COLUMN, value=name)
    for col in columns:
        if name and (name in col or col in name):
            return FieldMapping(field=field_name, source=SOURCE_COLUMN, value=col)
    return FieldMapping(field=field_name, source=SOURCE_NONE)
