"""打印后端抽象：GUI 只依赖这个接口，不关心底下是 BarTender 还是干跑模拟。"""
from __future__ import annotations

from abc import ABC, abstractmethod


class PrintBackend(ABC):
    """一个打印后端对应一次 BarTender 会话（或模拟会话）。"""

    #: 后端显示名，用于界面和日志
    name: str = "abstract"

    @abstractmethod
    def list_template_fields(self, template_path: str) -> list[str]:
        """枚举模板里的命名数据源（字段名列表）。"""

    @abstractmethod
    def list_printers(self) -> list[str]:
        """返回系统可用打印机名列表。"""

    @abstractmethod
    def open(self, template_path: str, printer: str = "") -> None:
        """打开模板；printer 为空则用模板保存的打印机。"""

    @abstractmethod
    def print_label(self, values: dict[str, str], copies: int = 1) -> None:
        """用给定字段值打印一张标签（copies 张相同标签）。

        调用前必须先 open()。values 的 key 是模板字段名。
        """

    @abstractmethod
    def render_preview(self, values: dict[str, str], out_path: str) -> str:
        """用给定字段值渲染一张预览图（PNG）到 out_path，返回实际文件路径。

        调用前必须先 open()。用于打印前的视觉核对，不需要真实打印机。
        """

    @abstractmethod
    def close(self) -> None:
        """关闭模板并释放后端资源。可重复调用。"""

    def __enter__(self) -> "PrintBackend":
        return self

    def __exit__(self, *exc) -> None:
        self.close()


class BackendError(Exception):
    """后端相关错误（BarTender 未安装、模板打不开、打印失败等），message 面向用户。"""


class PrintOutcomeUnknown(BackendError):
    """当前行可能已经下发；必须核对队列和实物，禁止自动重试。"""


class BackendCancelled(BackendError):
    """操作在提交当前打印命令前取消。"""
