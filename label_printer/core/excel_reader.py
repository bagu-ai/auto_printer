"""Excel 清单读取：自动探测表头行、解析每行默认打印张数。

纯 Python / pandas 实现，不依赖 GUI 与 BarTender，可单独测试。
"""
from __future__ import annotations

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
    rows: list[list[str]]            # 全部字符串化，空值为 ""
    copies: list[int]                # 每行默认打印张数（从送样数量解析，默认 1）
    header_row: int                  # 表头在原表中的行号（0 起）
    copies_column: str | None = field(default=None)

    def __len__(self) -> int:
        return len(self.rows)


def detect_header_row(raw: pd.DataFrame) -> int:
    """在未设表头的 DataFrame 里找表头行号，找不到返回 0。"""
    best_row, best_hits = 0, 0
    for i in range(min(MAX_HEADER_SCAN_ROWS, len(raw))):
        cells = {str(v).strip() for v in raw.iloc[i] if pd.notna(v)}
        hits = sum(1 for kw in HEADER_KEYWORDS if kw in cells)
        if hits > best_hits:
            best_row, best_hits = i, hits
    return best_row if best_hits >= MIN_KEYWORD_HITS else 0


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


def load_excel(path: str, header_row: int | None = None) -> SheetData:
    """加载 .xls/.xlsx 清单。

    header_row 为 None 时自动探测；否则使用指定行（0 起，指原表行号）。
    """
    raw = pd.read_excel(path, header=None, dtype=object)
    if raw.empty:
        raise ValueError("Excel 文件是空的")

    if header_row is None:
        header_row = detect_header_row(raw)

    columns = _dedupe_columns([str(v) for v in raw.iloc[header_row].fillna("")])
    body = raw.iloc[header_row + 1:].reset_index(drop=True)

    rows: list[list[str]] = []
    for _, r in body.iterrows():
        values = ["" if pd.isna(v) else str(v).strip() for v in r]
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
    )
