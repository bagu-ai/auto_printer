"""BarTender COM 后端：驱动 BarTender Enterprise Automation 填值并打印。

线程纪律（重要）：本类的实例必须在**同一个线程**内创建、使用、销毁，
且该线程须已调用 pythoncom.CoInitialize()。由 worker 线程负责工厂创建，
不要跨线程传递本实例或其中的 COM 对象。

所有 COM 错误都转成 BackendError（中文提示）。
"""
from __future__ import annotations

from time import monotonic
from uuid import uuid4

from .backend import BackendError, PrintBackend, PrintOutcomeUnknown

# Formats.Open 第二参数：关闭 BarTender 启动时自动打开的空白 Document1.btw
# （官方文档签名：Formats.Open(PathAndFileName, CloseOutFirstDocument, UsePrinter)）
BT_CLOSE_OUT_FIRST_DOCUMENT = True
BT_SAVE_DO_NOT_SAVE = 1  # BtSaveOptions.btDoNotSaveChanges；0 是 btPromptSave

# BtPrintResult 枚举（0=成功）
BT_PRINT_SUCCESS = 0
BT_PRINT_TIMEOUT = 1
BT_PRINT_FAILURE = 2
DEFAULT_PRINT_TIMEOUT_MS = 60_000

# BtColors 枚举（12.0 类型库实测）：Mono=0, 16色=1, 256色=2, 24Bit=4, 32Bit=5, 16Bit=6。
# 前四项（0/1/2/4）从旧版本沿用至今；5/6 是后来追加的。曾误用 6（以为是"24位色"），
# 在 BarTender 2016 (11.0) 上枚举越界，报「为第一个参数指定的颜色不正确」
# (DISP_E_EXCEPTION / E_UNEXPECTED)。4 = btColors24Bit 在 2016 与 12.x 上均有效。
BT_COLORS_24BIT = 4

# IBtFormat.Print 的 ActiveX DISPID（=48）。
# 已在本机 BarTender 12.0 类型库与真实调用验证；2016 ActiveX 官方文档的接口
# 顺序一致，但仍需在客户的 2016 主机做最终兼容验证。
# Print 最后一个 Messages 参数为 [out] IBtMessages**；pywin32 动态 Dispatch
# 不能给这个参数传 None，必须用 InvokeTypes 让它自动创建 out 参数。
DISPID_PRINT = 48
VT_OUT_DISPATCH = 16393  # pythoncom.VT_BYREF | pythoncom.VT_DISPATCH

# 参数个数/类型不匹配的 HRESULT：这类错误只可能意味着「该版本的接口签名不同」，
# 与模板内容无关，因此可以安全地回退到旧接口再试一次。
_SIGNATURE_MISMATCH_HRESULTS = frozenset({
    -2147352561,  # DISP_E_PARAMNOTOPTIONAL「非选择性的参数」
    -2147352562,  # DISP_E_BADPARAMCOUNT
    -2147352571,  # DISP_E_TYPEMISMATCH「类型不匹配」
})


def _is_signature_mismatch(exc: Exception) -> bool:
    hresult = getattr(exc, "hresult", None)
    if hresult is None and getattr(exc, "args", None):
        hresult = exc.args[0]
    return hresult in _SIGNATURE_MISMATCH_HRESULTS


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
    if msgs is None:
        return ""
    try:
        # InvokeTypes 返回的是原始 PyIDispatch；它没有高层 COM 集合的迭代接口。
        # 只包装原始接口，不能把普通字符串当 ProgID 传给 Dispatch。
        import pythoncom  # noqa: PLC0415
        if isinstance(msgs, pythoncom.TypeIIDs[pythoncom.IID_IDispatch]):
            msgs = _import_win32com().Dispatch(msgs)
        if isinstance(msgs, str):
            return msgs
        return "；".join(str(m.Message) for m in msgs)
    except Exception as exc:
        return f"无法读取 BarTender 消息详情：{exc}"


class BarTenderBackend(PrintBackend):
    name = "BarTender"

    def __init__(self, print_timeout_ms: int = DEFAULT_PRINT_TIMEOUT_MS,
                 optimize_repeated_values: bool = False,
                 wait_for_completion: bool = False) -> None:
        if (isinstance(print_timeout_ms, bool) or not isinstance(print_timeout_ms, int)
                or not 1 <= print_timeout_ms <= 2_147_483_647):
            raise ValueError("print_timeout_ms 必须是 1 到 2147483647 之间的毫秒整数")
        if not isinstance(wait_for_completion, bool):
            raise TypeError("wait_for_completion 必须是 bool")
        self._client = _import_win32com()
        self._app = None
        self._format = None
        self._print_timeout_ms = print_timeout_ms
        self._wait_for_completion = wait_for_completion
        # 模板可能包含 OnProcessData 等脚本；跳过重复赋值必须由调用方明确启用。
        # 提交兼容模式保持逐行完整赋值，不能与快速缓存一起启用。
        self._optimize_repeated_values = bool(optimize_repeated_values) and wait_for_completion
        self._observer = None
        self._job_name = ""
        self._active_job_name = ""
        self._pending_cleanup_errors: list[str] = []
        self._reset_format_cache()

    def set_observer(self, callback) -> None:
        """设置阶段观察器；回调异常不吞掉，防止日志失败后继续发送新任务。"""
        if callback is not None and not callable(callback):
            raise TypeError("observer 必须可调用或为 None")
        self._observer = callback

    def set_job_name(self, name: str) -> None:
        """调用方可在每行开始前设置可追溯的任务名；空值使用自动唯一名。"""
        self._job_name = str(name)
        self._active_job_name = ""

    def _observe(self, phase: str, operation: str, **details) -> None:
        if self._observer is not None:
            self._observer({"phase": phase, "operation": operation,
                            "job_name": self._active_job_name or self._job_name,
                            **details})

    def _reset_format_cache(self) -> None:
        self._print_setup = None
        self._field_values: dict[str, str] = {}
        self._copies = None

    def _setup_for_format(self):
        if self._print_setup is None:
            self._print_setup = self._format.PrintSetup
        return self._print_setup

    # ---- 会话管理 -------------------------------------------------------

    def _ensure_app(self):
        if self._app is None:
            self._observe("before", "application_start")
            started = monotonic()
            try:
                # 不通过 ROT 连接用户已经打开的 BarTender；清理只针对本次新建实例。
                self._app = self._client.DispatchEx("BarTender.Application")
            except Exception as exc:
                raise BackendError(
                    "无法启动 BarTender。\n"
                    "请确认本机已安装 BarTender（Automation / Enterprise 版），"
                    "且当前账号有运行权限。\n\n"
                    f"原始错误：{exc}"
                ) from exc
            self._observe("after", "application_start",
                          elapsed_ms=round((monotonic() - started) * 1000, 3))
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
                except Exception as exc:
                    # 保留枚举结果或原业务异常；由最终 close() 单独报告清理问题。
                    self._pending_cleanup_errors.append(f"关闭临时模板失败：{exc}")

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
        self._observe("before", "open", template=template_path, printer=printer)
        started = monotonic()
        try:
            self._format = app.Formats.Open(template_path, BT_CLOSE_OUT_FIRST_DOCUMENT, printer or "")
            if printer:
                self._setup_for_format().Printer = printer
        except Exception as exc:
            self._format = None
            self._reset_format_cache()
            raise BackendError(f"打不开模板文件：\n{template_path}\n\n原始错误：{exc}") from exc
        self._observe("after", "open",
                      elapsed_ms=round((monotonic() - started) * 1000, 3))

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
            self._active_job_name or self._job_name, self._wait_for_completion, self._print_timeout_ms,
        )

    def print_label(self, values: dict[str, str], copies: int = 1) -> None:
        if self._format is None:
            raise BackendError("内部错误：尚未打开模板就调用打印")
        self._active_job_name = self._job_name or f"label-{uuid4().hex[:24]}"
        try:
            self._observe("before", "prepare_label")
            for name, value in values.items():
                cached = self._optimize_repeated_values and self._field_values.get(name) == value
                self._observe("before", "set_field", field=name, cached=cached)
                started = monotonic()
                if not cached:
                    self._set_field(name, value)
                    self._field_values[name] = value  # 仅成功后缓存
                self._observe("after", "set_field", field=name, cached=cached,
                              elapsed_ms=round((monotonic() - started) * 1000, 3))
            try:
                quantity = max(1, int(copies))
            except (TypeError, ValueError, OverflowError) as exc:
                raise BackendError(f"标签份数无效：{copies!r}") from exc
            cached = self._optimize_repeated_values and self._copies == quantity
            self._observe("before", "set_copies", copies=quantity, cached=cached)
            started = monotonic()
            if not cached:
                try:
                    self._setup_for_format().IdenticalCopiesOfLabel = quantity
                except Exception as exc:
                    raise BackendError(f"设置标签份数失败，尚未调用打印：{exc}") from exc
                self._copies = quantity
            self._observe("after", "set_copies", copies=quantity, cached=cached,
                          elapsed_ms=round((monotonic() - started) * 1000, 3))
            self._observe("after", "prepare_label")
            # 必须先让观察器成功持久化即将发送的记录，再进入可能阻塞的 COM 调用。
            self._observe("before", "print", timeout_ms=self._print_timeout_ms,
                          copies=quantity, wait_for_completion=self._wait_for_completion)
        except Exception:
            self._reset_format_cache()
            raise

        started = monotonic()
        try:
            ret = self._invoke_print()
        except Exception as exc:
            self._reset_format_cache()
            raise PrintOutcomeUnknown(
                f"BarTender 调用中断，任务是否已提交无法确定，请核对队列后处理，勿自动重打。原始错误：{exc}"
            ) from exc
        result = ret[0] if isinstance(ret, tuple) and ret else ret
        detail = _messages_text(ret)
        try:
            self._observe("after", "print", result=result, messages=detail,
                          wait_for_completion=self._wait_for_completion,
                          elapsed_ms=round((monotonic() - started) * 1000, 3))
        except Exception as exc:
            self._reset_format_cache()
            raise PrintOutcomeUnknown(
                f"打印调用已返回，但结果记录失败；请核对队列及实物，勿自动重打：{exc}"
            ) from exc
        if result != BT_PRINT_SUCCESS:
            self._reset_format_cache()
            detail = detail or f"BtPrintResult={result}"
            if result == BT_PRINT_FAILURE:
                # btFailure 不能证明整行的所有份数均未出纸；保留部分提交的可能性。
                raise PrintOutcomeUnknown(
                    f"BarTender 返回打印失败，但可能已有部分标签提交或出纸，"
                    f"请核对队列及实物，勿自动重打：{detail}"
                )
            if result == BT_PRINT_TIMEOUT:
                raise PrintOutcomeUnknown(
                    f"BarTender 等待超过 {self._print_timeout_ms / 1000:g} 秒，"
                    f"任务可能已进入队列，请核对后处理，勿自动重打：{detail}"
                )
            raise PrintOutcomeUnknown(f"BarTender 返回未知结果，请核对队列后处理：{detail}")

    def render_preview(self, values: dict[str, str], out_path: str) -> str:
        """填值后把**单张标签设计面**直接渲染成 300 DPI PNG。

        优先使用 ExportImage（真实验证：60×30mm → 708×354 px）；旧版本若无该
        方法则回退 ExportToFile。不要用 ExportPrintPreviewToImage——它导出整张
        A4 打印页（本机为 2126×3189），缩进校样台后文字必然显得很小、模糊。

        ExportImage/ExportToFile 是旧 ActiveX 方法，没有 [out] Messages 参数
        （不同于 Print/ExportPrintPreviewToImage），失败表现为抛 COM 异常或不
        生成文件，因此这里的诊断靠异常原文 + 文件存在性检查。回退只在「方法
        不存在」或「签名不匹配（参数个数/类型不符）」时发生，真实渲染错误
        （如色深越界）不会被第二个调用掩盖。
        """
        if self._format is None:
            raise BackendError("内部错误：尚未打开模板就调用预览")
        # 预览可能触发模板脚本，不沿用打印字段与份数缓存。
        self._reset_format_cache()
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
            self._export_label_image(output)
        except Exception as exc:
            raise BackendError(f"生成校样失败：{exc}") from exc

        if os.path.exists(output):
            return output
        raise BackendError("生成校样失败：BarTender 未生成图片文件")

    def _export_label_image(self, output: str) -> None:
        """导出单张标签设计面 PNG，兼容 2016 (11.0) 与 12.x。

        色深统一传 BT_COLORS_24BIT（=4）：该枚举值在各版本中含义一致；
        不要改成 6——那是 12.x 才追加的 btColors16Bit，2016 上越界报错。

        12.x 的 ExportImage 为 7 参（含背景色/边框）；老版本若方法缺失
        （getattr 为 None）或参数个数不同（报「非选择性的参数」等签名错误），
        回退自早期版本即有的 5 参旧接口 ExportToFile。
        """
        export_image = getattr(self._format, "ExportImage", None)
        if export_image is not None:
            try:
                # (文件, 类型, 24位色, DPI, 覆盖, 背景白, 画标签边框)
                export_image(output, "png", BT_COLORS_24BIT, 300, 2, 16777215, True)
                return
            except Exception as exc:
                if not _is_signature_mismatch(exc):
                    raise  # 真实渲染错误原样上抛，不用旧接口掩盖
                # 签名不符：落入旧接口再试一次
        self._format.ExportToFile(output, "png", BT_COLORS_24BIT, 300, 2)

    def _set_field(self, name: str, value: str) -> None:
        fmt = self._format
        try:
            fmt.SetNamedSubStringValue(name, value)
            return
        except Exception as legacy_exc:
            first_error = legacy_exc
        try:
            fmt.NamedDataSources.Item(name).Value = value
        except Exception as exc:
            raise BackendError(
                f"设置模板字段「{name}」失败，请检查字段与 BarTender 状态。"
                f"旧接口：{first_error}；新接口：{exc}"
            ) from exc

    # ---- 清理 -----------------------------------------------------------

    def close_format(self) -> None:
        self._reset_format_cache()
        fmt, self._format = self._format, None
        if fmt is not None:
            try:
                fmt.Close(BT_SAVE_DO_NOT_SAVE)
            except Exception as exc:
                raise BackendError(f"关闭 BarTender 模板失败：{exc}") from exc

    def close(self) -> None:
        errors, self._pending_cleanup_errors = self._pending_cleanup_errors, []
        try:
            self.close_format()
        except Exception as exc:
            errors.append(str(exc))
        # 模板关闭失败也要尝试 Quit；只释放本实例引用，不处理其他 BarTender 进程。
        app, self._app = self._app, None
        if app is not None:
            try:
                app.Quit(BT_SAVE_DO_NOT_SAVE)
            except Exception as exc:
                errors.append(f"退出 BarTender 失败：{exc}")
        if errors:
            raise BackendError("BarTender 清理未完全成功：" + "；".join(errors))
