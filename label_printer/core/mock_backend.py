"""干跑后端：不依赖 BarTender，用于开发机调试与演示。

字段列表返回一组示例字段；打印动作只记录到内存日志；
预览用 Qt 画一张示意 PNG（字段名=值），方便演示预览流程。
"""
from __future__ import annotations

from .backend import BackendError, PrintBackend

MOCK_FIELDS = ["发货新编号", "客户香型", "发货粒径", "英文", "胶囊原编号"]
MOCK_PRINTERS = ["TOSHIBA B-EX4T1 (305 dpi)", "ARGOX CP-2140M", "Microsoft Print to PDF"]


class MockBackend(PrintBackend):
    name = "干跑模式（未连接 BarTender）"

    def __init__(self) -> None:
        self.template_path: str | None = None
        self.printer: str = ""
        self.printed: list[tuple[dict[str, str], int]] = []
        self.closed_count = 0

    def list_template_fields(self, template_path: str) -> list[str]:
        return list(MOCK_FIELDS)

    def list_printers(self) -> list[str]:
        return list(MOCK_PRINTERS)

    def open(self, template_path: str, printer: str = "") -> None:
        self.template_path = template_path
        self.printer = printer or MOCK_PRINTERS[0]

    def print_label(self, values: dict[str, str], copies: int = 1) -> None:
        if self.template_path is None:
            raise BackendError("内部错误：尚未打开模板就调用打印")
        self.printed.append((dict(values), copies))

    def render_preview(self, values: dict[str, str], out_path: str) -> str:
        if self.template_path is None:
            raise BackendError("内部错误：尚未打开模板就调用预览")
        import os  # noqa: PLC0415

        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)

        from PySide6.QtCore import Qt
        from PySide6.QtGui import QColor, QImage, QPainter

        img = QImage(600, 300, QImage.Format_RGB32)
        img.fill(QColor("white"))
        p = QPainter(img)
        p.setPen(QColor("#cccccc"))
        p.drawRect(1, 1, 598, 298)
        p.setPen(QColor("black"))
        y = 45
        p.drawText(20, y, "【干跑预览】60×30mm 标签示意")
        for name, value in values.items():
            y += 40
            p.drawText(20, y, f"{name}：{value}")
        p.end()
        if not img.save(out_path, "PNG"):
            raise BackendError(f"预览图片写入失败：{out_path}")
        return out_path

    def close(self) -> None:
        self.template_path = None
        self.closed_count += 1
