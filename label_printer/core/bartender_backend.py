"""BarTender COM 后端：驱动 BarTender Enterprise Automation 填值并打印。

只在装有 BarTender 的目标机上使用；所有 COM 错误都转成 BackendError（中文提示）。
"""
from __future__ import annotations

from .backend import BackendError, PrintBackend

# BarTender 常量（避免依赖 makepy 生成的常量模块）
BT_FORMAT_OPEN_READ_ONLY = True
BT_SAVE_DO_NOT_SAVE = 0  # BtSaveOptions.btSaveDoNotSaveChanges


def _import_win32com():
    try:
        import win32com.client  # noqa: PLC0415
        return win32com.client
    except ImportError as exc:
        raise BackendError(
            "缺少 pywin32 组件，无法连接 BarTender。\n"
            "请运行：pip install pywin32"
        ) from exc


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
            fmt = app.Formats.Open(template_path, BT_FORMAT_OPEN_READ_ONLY, "")
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
        try:
            import win32print  # noqa: PLC0415
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            return [p[2] for p in win32print.EnumPrinters(flags)]
        except Exception:
            # 退回 BarTender 的打印机列表
            try:
                app = self._ensure_app()
                return [str(p) for p in app.Printers]
            except Exception:
                return []

    # ---- 打印 -----------------------------------------------------------

    def open(self, template_path: str, printer: str = "") -> None:
        self.close_format()
        app = self._ensure_app()
        try:
            self._format = app.Formats.Open(template_path, False, printer or "")
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
            self._format.PrintOut(False, False)  # 不弹打印对话框、不等完成
        except BackendError:
            raise
        except Exception as exc:
            raise BackendError(f"打印失败：{exc}") from exc

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
