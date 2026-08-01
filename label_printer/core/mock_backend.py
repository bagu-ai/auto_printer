"""干跑后端：不依赖 BarTender，用于开发机调试与演示。

字段列表返回一组示例字段；打印动作只记录到内存日志。
"""
from __future__ import annotations

from .backend import PrintBackend

MOCK_FIELDS = ["发货新编号", "客户香型", "发货粒径", "英文", "胶囊原编号"]
MOCK_PRINTERS = ["TOSHIBA B-EX4T1 (305 dpi)", "ARGOX CP-2140M", "Microsoft Print to PDF"]


class MockBackend(PrintBackend):
    name = "干跑模式（未连接 BarTender）"

    def __init__(self) -> None:
        self.template_path: str | None = None
        self.printer: str = ""
        self.printed: list[tuple[dict[str, str], int]] = []

    def list_template_fields(self, template_path: str) -> list[str]:
        return list(MOCK_FIELDS)

    def list_printers(self) -> list[str]:
        return list(MOCK_PRINTERS)

    def open(self, template_path: str, printer: str = "") -> None:
        self.template_path = template_path
        self.printer = printer or MOCK_PRINTERS[0]

    def print_label(self, values: dict[str, str], copies: int = 1) -> None:
        self.printed.append((dict(values), copies))

    def close(self) -> None:
        self.template_path = None
