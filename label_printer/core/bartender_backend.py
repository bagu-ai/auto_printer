"""BarTender COM 后端：驱动 BarTender Enterprise Automation 填值并打印。

线程纪律（重要）：本类的实例必须在**同一个线程**内创建、使用、销毁，
且该线程须已调用 pythoncom.CoInitialize()。由 worker 线程负责工厂创建，
不要跨线程传递本实例或其中的 COM 对象。

所有 COM 错误都转成 BackendError（中文提示）。
"""
from __future__ import annotations

from .backend import BackendError, PrintBackend

# Formats.Open 第二参数：关闭 BarTender 启动时自动打开的空白 Document1.btw
# （官方文档签名：Formats.Open(PathAndFileName, CloseOutFirstDocument, UsePrinter)）
BT_CLOSE_OUT_FIRST_DOCUMENT = True
BT_SAVE_DO_NOT_SAVE = 0  # BtSaveOptions.btSaveDoNotSaveChanges

# BtPrintResult 枚举（0=成功）
BT_PRINT_SUCCESS = 0


def _import_win32com():
    try:
        import win32com.client  # noqa: PLC0415
        return win32com.client
    except ImportError as exc:
        raise BackendError(
            "缺少 pywin32 组件，无法连接 BarTender。\n"
            "请运行：pip install pywin32"
        ) from exc


def _messages_text(ret) -> str:
    """Format.Print 的返回可能是 (result, messages) 元组，提取消息文本。"""
    if not isinstance(ret, tuple) or len(ret) < 2:
        return ""
    msgs = ret[1]
    try:
        return "；".join(str(m.Message) for m in msgs) or ""
    except Exception:
        return str(msgs) if msgs else ""


class BarTenderBackend(PrintBackend):
    name = "BarTender"

    def __init__(self) -> None:
        self._client = _import_win32com()
        self._app = None
        self._format = None

    # ---- 会话管理 -------------------------------------------------------

    def _ensure_app(self):
        if self._app is None:
            try:
                self._app = self._client.Dispatch("BarTender.Application")
            except Exception as exc:
                raise BackendError(
                    "无法启动 BarTender。\n"
                    "请确认本机已安装 BarTender（Automation / Enterprise 版），"
                    "且当前账号有运行权限。\n\n"
                    f"原始错误：{exc}"
                ) from exc
        return self._app

    def list_template_fields(self, template_path: str) -> list[str]:
        """打开模板枚举命名数据源后立即关闭。"""
        app = self._ensure_app()
        fmt = None
        try:
            fmt = app.Formats.Open(template_path, BT_CLOSE_OUT_FIRST_DOCUMENT, "")
            return self._collect_fields(fmt)
        except BackendError:
            raise
        except Exception as exc:
            raise BackendError(f"打不开模板文件：\n{template_path}\n\n原始错误：{exc}") from exc
        finally:
            if fmt is not None:
                try:
                    fmt.Close(BT_SAVE_DO_NOT_SAVE)
                except Exception:
                    pass

    @staticmethod
    def _collect_fields(fmt) -> list[str]:
        """命名子字符串（旧 API）和新式命名数据源取并集。"""
        fields: list[str] = []

        def _add(collection) -> None:
            try:
                for item in collection:
                    name = str(item.Name)
                    if name and name not in fields:
                        fields.append(name)
            except Exception:
                pass

        _add(getattr(fmt, "NamedSubStrings", []))
        _add(getattr(fmt, "NamedDataSources", []))
        return fields

    def list_printers(self) -> list[str]:
        """只走 win32print（非 COM，可在任意线程调用）。"""
        try:
            import win32print  # noqa: PLC0415
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            return [p[2] for p in win32print.EnumPrinters(flags)]
        except Exception:
            return []

    # ---- 打印 -----------------------------------------------------------

    def open(self, template_path: str, printer: str = "") -> None:
        self.close_format()
        app = self._ensure_app()
        try:
            self._format = app.Formats.Open(template_path, BT_CLOSE_OUT_FIRST_DOCUMENT, printer or "")
            if printer:
                self._format.PrintSetup.Printer = printer
        except Exception as exc:
            self._format = None
            raise BackendError(f"打不开模板文件：\n{template_path}\n\n原始错误：{exc}") from exc

    def print_label(self, values: dict[str, str], copies: int = 1) -> None:
        if self._format is None:
            raise BackendError("内部错误：尚未打开模板就调用打印")
        try:
            for name, value in values.items():
                self._set_field(name, value)
            setup = self._format.PrintSetup
            setup.IdenticalCopiesOfLabel = max(1, int(copies))
            # Format.Print(作业名, 等提交spooler, 超时ms, 消息输出) → (BtPrintResult, Messages)
            ret = self._format.Print("", True, -1, None)
            result = ret[0] if isinstance(ret, tuple) else ret
            if result != BT_PRINT_SUCCESS:
                detail = _messages_text(ret) or f"BtPrintResult={result}"
                raise BackendError(f"BarTender 打印失败：{detail}")
        except BackendError:
            raise
        except Exception as exc:
            raise BackendError(f"打印失败：{exc}") from exc

    def render_preview(self, values: dict[str, str], out_path: str) -> str:
        """填值后把标签渲染成 PNG（ExportPrintPreviewToImage，2016 起支持）。"""
        if self._format is None:
            raise BackendError("内部错误：尚未打开模板就调用预览")
        import os  # noqa: PLC0415

        out_dir = os.path.dirname(os.path.abspath(out_path))
        os.makedirs(out_dir, exist_ok=True)
        try:
            for name, value in values.items():
                self._set_field(name, value)
            # (目录, 文件名模板, 类型, 颜色=24位, DPI=300, 背景=白, 保存选项=2, 含边距=False, 含边框=False, 消息)
            ret = self._format.ExportPrintPreviewToImage(
                out_dir, "label_%PageNumber%.png", "png",
                6, 300, 16777215, 2, False, False, None)
            result = ret[0] if isinstance(ret, tuple) else ret
            if result != BT_PRINT_SUCCESS:
                detail = _messages_text(ret) or f"BtPrintResult={result}"
                raise BackendError(f"生成预览失败：{detail}")
        except BackendError:
            raise
        except Exception as exc:
            raise BackendError(f"生成预览失败：{exc}") from exc

        expected = os.path.join(out_dir, "label_1.png")
        if os.path.exists(expected):
            return expected
        # 文件名页码格式可能随版本不同（label_1.png / label_01.png），兜底找最新的
        candidates = sorted(
            (f for f in os.listdir(out_dir) if f.startswith("label_") and f.endswith(".png")),
            key=lambda f: os.path.getmtime(os.path.join(out_dir, f)),
            reverse=True)
        if candidates:
            return os.path.join(out_dir, candidates[0])
        raise BackendError("生成预览失败：BarTender 报告成功但找不到输出文件")

    def _set_field(self, name: str, value: str) -> None:
        fmt = self._format
        try:
            fmt.SetNamedSubStringValue(name, value)
            return
        except Exception:
            pass
        try:
            fmt.NamedDataSources.Item(name).Value = value
        except Exception as exc:
            raise BackendError(f"模板里找不到字段「{name}」，请检查映射配置") from exc

    # ---- 清理 -----------------------------------------------------------

    def close_format(self) -> None:
        if self._format is not None:
            try:
                self._format.Close(BT_SAVE_DO_NOT_SAVE)
            except Exception:
                pass
            self._format = None

    def close(self) -> None:
        self.close_format()
        if self._app is not None:
            try:
                self._app.Quit(BT_SAVE_DO_NOT_SAVE)
            except Exception:
                pass
            self._app = None
