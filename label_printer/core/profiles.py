"""映射配置：每个模板的「模板字段 ← Excel 列 / 固定文本」对应关系存成 JSON。

profile 文件名 = 模板名 + 模板全路径短哈希，避免不同目录的同名模板互相串配置；
兼容 v1.0 的旧命名（仅模板名），加载到旧配置后自动迁移为新名。
"""
from __future__ import annotations

import hashlib
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
    prefix: str = ""           # 加在每行取值前的固定文本（如「编号：」）
    suffix: str = ""           # 加在每行取值后的固定文本（如「（样品）」）

    def render(self, cell: str) -> str:
        """最终填充值 = 前缀 + 取值 + 后缀。"""
        return f"{self.prefix}{cell}{self.suffix}"


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


def _safe_stem(template_path: str) -> str:
    stem = Path(template_path).stem
    return "".join(c if c not in '\\/:*?"<>|' else "_" for c in stem)


def _path_hash(template_path: str) -> str:
    normalized = str(Path(template_path).resolve()).lower()
    return hashlib.sha1(normalized.encode("utf-8")).hexdigest()[:8]


def profile_path_for(template_path: str, profiles_dir: Path) -> Path:
    """按模板名+路径哈希生成 profile 路径。"""
    return profiles_dir / f"{_safe_stem(template_path)}_{_path_hash(template_path)}.json"


def _legacy_path_for(template_path: str, profiles_dir: Path) -> Path:
    """v1.0 旧命名（仅模板名），用于迁移兼容。"""
    return profiles_dir / f"{_safe_stem(template_path)}.json"


def save_profile(profile: Profile, profiles_dir: Path) -> Path:
    profiles_dir.mkdir(parents=True, exist_ok=True)
    path = profile_path_for(profile.template_path, profiles_dir)
    data = {
        "template_path": profile.template_path,
        "printer": profile.printer,
        "mappings": [
            {"field": m.field, "source": m.source, "value": m.value,
             "prefix": m.prefix, "suffix": m.suffix}
            for m in profile.mappings
        ],
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def load_profile(template_path: str, profiles_dir: Path) -> Profile | None:
    path = profile_path_for(template_path, profiles_dir)
    legacy = _legacy_path_for(template_path, profiles_dir)
    migrating = False
    if not path.exists() and legacy.exists():
        path, migrating = legacy, True
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None
    profile = Profile(
        template_path=data.get("template_path", template_path),
        printer=data.get("printer", ""),
        mappings=[
            FieldMapping(
                field=m.get("field", ""),
                source=m.get("source", SOURCE_NONE),
                value=m.get("value", ""),
                prefix=m.get("prefix", ""),   # 旧版配置没有这两个键
                suffix=m.get("suffix", ""),
            )
            for m in data.get("mappings", [])
        ],
    )
    if migrating:
        # 旧配置迁移为新命名；失败不影响使用
        try:
            save_profile(profile, profiles_dir)
            path.unlink(missing_ok=True)
        except OSError:
            pass
    return profile


def guess_mapping(field_name: str, columns: list[str]) -> FieldMapping:
    """按名字自动猜映射：精确匹配 > 包含匹配 > 留空。

    匹配前剥掉字段名首尾的空格和冒号——不少模板的数据源名带全角冒号尾巴
    （如「香型：」「粒径：」），不剥掉就连「客户香型」都匹配不上。
    """
    name = field_name.strip().strip("：:").strip()
    if not name:
        return FieldMapping(field=field_name, source=SOURCE_NONE)
    if name in columns:
        return FieldMapping(field=field_name, source=SOURCE_COLUMN, value=name)
    for col in columns:
        if name in col or col in name:
            return FieldMapping(field=field_name, source=SOURCE_COLUMN, value=col)
    return FieldMapping(field=field_name, source=SOURCE_NONE)


def validate_jobs(mappings: list[FieldMapping], columns: list[str],
                  rows: list[list[str]], copies: list[int]) -> tuple[list[str], list[str]]:
    """打印前预检。返回 (errors, warnings)：errors 非空应阻止打印。

    rows/copies 只包含已勾选的行。
    """
    errors: list[str] = []
    warnings: list[str] = []

    if not rows:
        errors.append("没有勾选任何数据行")
        return errors, warnings

    col_index = {c: i for i, c in enumerate(columns)}
    unmapped = [m.field for m in mappings if m.source == SOURCE_NONE]
    if unmapped:
        warnings.append("以下模板字段留空不打印：" + "、".join(unmapped))

    for i, copies_n in enumerate(copies):
        if copies_n < 1:
            errors.append(f"第 {i + 1} 行的打印张数无效（{copies_n}）")

    for m in mappings:
        if m.source != SOURCE_COLUMN or m.value not in col_index:
            continue
        idx = col_index[m.value]
        empty_rows = [str(i + 1) for i, row in enumerate(rows) if not row[idx]]
        if empty_rows:
            shown = "、".join(empty_rows[:5]) + ("…" if len(empty_rows) > 5 else "")
            warnings.append(f"「{m.value}」有 {len(empty_rows)} 行为空（第 {shown} 行）")

    # 发货新编号类字段查重（取第一个映射到含“编号”列的字段）
    for m in mappings:
        if m.source == SOURCE_COLUMN and "编号" in m.value and m.value in col_index:
            idx = col_index[m.value]
            seen: dict[str, int] = {}
            dups: list[str] = []
            for row in rows:
                v = row[idx]
                if not v:
                    continue
                seen[v] = seen.get(v, 0) + 1
                if seen[v] == 2:
                    dups.append(v)
            if dups:
                warnings.append(f"「{m.value}」存在重复值：{'、'.join(dups[:5])}")
            break

    return errors, warnings
