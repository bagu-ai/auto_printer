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

# IBtFormat.Print 的 ActiveX DISPID（=48）。
# 已在本机 BarTender 12.0 类型库与真实调用验证；2016 ActiveX 官方文档的接口
# 顺序一致，但仍需在客户的 2016 主机做最终兼容验证。
# Print 最后一个 Messages 参数为 [out] IBtMessages**；pywin32 动态 Dispatch
# 不能给这个参数传 None，必须用 InvokeTypes 让它自动创建 out 参数。
DISPID_PRINT = 48
VT_OUT_DISPATCH = 16393  # pythoncom.VT_BYREF | pythoncom.VT_DISPATCH


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

    def _invoke_print(self):
        """调用含 [out] Messages 的 Format.Print。

        动态 Dispatch 直接传 None 会报「参数 4 类型不匹配」；这里按 ActiveX
        类型库显式描述输入/输出类型，让 pywin32 返回 (BtPrintResult, Messages)。
        """
        import pythoncom  # noqa: PLC0415

        return self._format._oleobj_.InvokeTypes(
            DISPID_PRINT, 0, pythoncom.DISPATCH_METHOD,
            (pythoncom.VT_I4, 0),
            (
                (pythoncom.VT_BSTR, pythoncom.PARAMFLAG_FIN),
                (pythoncom.VT_BOOL, pythoncom.PARAMFLAG_FIN),
                (pythoncom.VT_I4, pythoncom.PARAMFLAG_FIN),
                (VT_OUT_DISPATCH, pythoncom.PARAMFLAG_FOUT),
            ),
            "", True, -1,
        )

    def print_label(self, values: dict[str, str], copies: int = 1) -> None:
        if self._format is None:
            raise BackendError("内部错误：尚未打开模板就调用打印")
        try:
            for name, value in values.items():
                self._set_field(name, value)
            setup = self._format.PrintSetup
            setup.IdenticalCopiesOfLabel = max(1, int(copies))
            ret = self._invoke_print()
            result = ret[0] if isinstance(ret, tuple) else ret
            if result != BT_PRINT_SUCCESS:
                detail = _messages_text(ret) or f"BtPrintResult={result}"
                raise BackendError(f"BarTender 打印失败：{detail}")
        except BackendError:
            raise
        except Exception as exc:
            raise BackendError(f"打印失败：{exc}") from exc

    def render_preview(self, values: dict[str, str], out_path: str) -> str:
        """填值后把**单张标签设计面**直接渲染成 300 DPI PNG。

        优先使用 ExportImage（真实验证：60×30mm → 708×354 px）；旧版本若无该
        方法则回退 ExportToFile。不要用 ExportPrintPreviewToImage——它导出整张
        A4 打印页（本机为 2126×3189），缩进校样台后文字必然显得很小、模糊。

        ExportImage/ExportToFile 是旧 ActiveX 方法，没有 [out] Messages 参数
        （不同于 Print/ExportPrintPreviewToImage），失败表现为抛 COM 异常或不
        生成文件，因此这里的诊断靠异常原文 + 文件存在性检查。回退只在「方法
        不存在」时发生，真实渲染错误不会被第二个调用掩盖。
        """
        if self._format is None:
            raise BackendError("内部错误：尚未打开模板就调用预览")
        import os  # noqa: PLC0415

        out_dir = os.path.dirname(os.path.abspath(out_path))
        os.makedirs(out_dir, exist_ok=True)
        # 尊重调用方给的文件名（backend 契约：写入 out_path）
        output = os.path.join(out_dir, os.path.basename(out_path))
        try:
            for name, value in values.items():
                self._set_field(name, value)
            # 打印份数与单张设计面导出无关，但旧模板里若保存了多份数，
            # 某些版本会按份数导出多页/多图；统一钳为 1 保证只出单张。
            self._format.PrintSetup.IdenticalCopiesOfLabel = 1
            export_image = getattr(self._format, "ExportImage", None)
            if export_image is not None:
                # (文件, 类型, 24位色, DPI, 覆盖, 背景白, 画标签边框)
                export_image(output, "png", 6, 300, 2, 16777215, True)
            else:
                # BarTender 2016 的旧 ActiveX 接口：导出单张设计面
                self._format.ExportToFile(output, "png", 6, 300, 2)
        except Exception as exc:
            raise BackendError(f"生成校样失败：{exc}") from exc

        if os.path.exists(output):
            return output
        raise BackendError("生成校样失败：BarTender 未生成图片文件")

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
