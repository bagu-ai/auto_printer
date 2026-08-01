"""Excel 清单读取：多 sheet 支持、自动探测表头行、单元格显示值格式化、每行默认张数。

纯 Python / pandas 实现，不依赖 GUI 与 BarTender，可单独测试。
"""
from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field

import pandas as pd

# 用来识别“表头行”的关键词：某一行里命中 >= MIN_KEYWORD_HITS 个即认为是表头
HEADER_KEYWORDS = (
    "发货新编号", "客户香型", "发货粒径", "序号",
    "胶囊原编号", "香精编号", "送样数量",
)
MIN_KEYWORD_HITS = 2
MAX_HEADER_SCAN_ROWS = 20

# “送样数量”列的候选名字（用于解析默认张数）
COPIES_SOURCE_COLUMNS = ("送样数量", "数量", "份数")


@dataclass
class SheetData:
    """一份已解析的 Excel 清单。"""

    path: str
    columns: list[str]
    rows: list[list[str]]            # 全部按显示值字符串化，空值为 ""
    copies: list[int]                # 每行默认打印张数（从送样数量解析，默认 1）
    header_row: int                  # 表头在原表中的行号（0 起）
    copies_column: str | None = field(default=None)
    sheet_names: list[str] = field(default_factory=list)  # 工作簿全部 sheet 名
    sheet_name: str = ""             # 当前读取的 sheet 名

    def __len__(self) -> int:
        return len(self.rows)


def list_sheet_names(path: str) -> list[str]:
    """返回工作簿的全部 sheet 名。"""
    return list(pd.ExcelFile(path).sheet_names)


def detect_header_row(raw: pd.DataFrame) -> int:
    """在未设表头的 DataFrame 里找表头行号，找不到返回 0。"""
    best_row, best_hits = 0, 0
    for i in range(min(MAX_HEADER_SCAN_ROWS, len(raw))):
        cells = {str(v).strip() for v in raw.iloc[i] if pd.notna(v)}
        hits = sum(1 for kw in HEADER_KEYWORDS if kw in cells)
        if hits > best_hits:
            best_row, best_hits = i, hits
    return best_row if best_hits >= MIN_KEYWORD_HITS else 0


def format_cell(v) -> str:
    """把单元格值转成适合打印的显示文本。

    - 日期时间：零时间 → YYYY-MM-DD；否则 YYYY-MM-DD HH:MM
    - 整值浮点：3.0 → "3"（避免标签上出现多余的 .0）
    - 其余：去首尾空白的字符串
    """
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, (pd.Timestamp, _dt.datetime)):
        if v.hour == 0 and v.minute == 0 and v.second == 0:
            return v.strftime("%Y-%m-%d")
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, _dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def parse_copies(text: str) -> int:
    """从 '10g*3瓶' / '3瓶' / 'x3' 之类文本解析瓶数，解析不出返回 1。"""
    if not text:
        return 1
    m = re.search(r"(\d+)\s*瓶", text)
    if m:
        return max(1, int(m.group(1)))
    m = re.search(r"[*x×]\s*(\d+)", text, re.IGNORECASE)
    if m:
        return max(1, int(m.group(1)))
    m = re.fullmatch(r"\s*(\d+)\s*", text)
    if m:
        return max(1, int(m.group(1)))
    return 1


def _dedupe_columns(names: list[str]) -> list[str]:
    """空列名补 '列N'，重名加后缀，保证列名唯一可下拉选择。"""
    seen: dict[str, int] = {}
    out = []
    for idx, name in enumerate(names):
        name = name.strip() or f"列{idx + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        out.append(name)
    return out


def load_excel(path: str, sheet_name: str | None = None,
               header_row: int | None = None) -> SheetData:
    """加载 .xls/.xlsx 清单。

    sheet_name 为 None 时读第一个 sheet；header_row 为 None 时自动探测（0 起原表行号）。
    """
    sheet_names = list_sheet_names(path)
    if sheet_name is None:
        sheet_name = sheet_names[0]
    elif sheet_name not in sheet_names:
        raise ValueError(f"工作表「{sheet_name}」不存在（可选：{'、'.join(sheet_names)}）")

    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
    if raw.empty:
        raise ValueError(f"工作表「{sheet_name}」是空的")

    if header_row is None:
        header_row = detect_header_row(raw)

    columns = _dedupe_columns([str(v) for v in raw.iloc[header_row].fillna("")])
    body = raw.iloc[header_row + 1:].reset_index(drop=True)

    rows: list[list[str]] = []
    for _, r in body.iterrows():
        values = [format_cell(v) for v in r]
        if any(values):  # 跳过整行空白
            rows.append(values)

    copies_column = next((c for c in COPIES_SOURCE_COLUMNS if c in columns), None)
    copies = [1] * len(rows)
    if copies_column:
        idx = columns.index(copies_column)
        copies = [parse_copies(r[idx]) for r in rows]

    return SheetData(
        path=path,
        columns=columns,
        rows=rows,
        copies=copies,
        header_row=header_row,
        copies_column=copies_column,
        sheet_names=sheet_names,
        sheet_name=sheet_name,
    )
