"""BarTender ActiveX 调用契约测试（无需安装/启动 BarTender）。

- 锁定含 [out] Messages 的 Format.Print 必须通过 InvokeTypes 调用，防止以后
  又退回 `method(..., None)` 导致 COM「类型不匹配」。
- 锁定 render_preview 的实现：色深必须用 btColors24Bit(4)——6 是 12.x 才追加
  的 btColors16Bit，BarTender 2016 报「为第一个参数指定的颜色不正确」；
  ExportImage 方法缺失或签名不符（参数个数/类型）时回退 ExportToFile；
  真实渲染错误（含色深越界）不被回退掩盖；写入调用方指定的 out_path 文件名。
"""
import pythoncom
import pytest
from types import SimpleNamespace

from label_printer.core.backend import BackendError, PrintOutcomeUnknown
from label_printer.core import bartender_backend as backend_module
from label_printer.core.bartender_backend import (
    BT_COLORS_24BIT,
    DEFAULT_PRINT_TIMEOUT_MS,
    DISPID_PRINT,
    BarTenderBackend,
    _messages_text,
)


class FakeOleObj:
    def __init__(self, result=(0, "messages")):
        self.result = result
        self.calls = []

    def InvokeTypes(self, *args):
        self.calls.append(args)
        if isinstance(self.result, Exception):
            raise self.result
        return self.result


class FakePrintSetup:
    def __init__(self):
        self._copies = 7  # 模板里保存的非 1 份数
        self.copy_writes = []

    @property
    def IdenticalCopiesOfLabel(self):
        return self._copies

    @IdenticalCopiesOfLabel.setter
    def IdenticalCopiesOfLabel(self, value):
        self.copy_writes.append(value)
        self._copies = value


class FakeFormat:
    """模拟带新式 ExportImage 的 Format。"""

    def __init__(self, create_file: bool = True, export_error: Exception | None = None):
        self._oleobj_ = FakeOleObj()
        self.PrintSetup = FakePrintSetup()
        self.create_file = create_file
        self.export_error = export_error
        self.export_image_calls = []
        self.export_to_file_calls = []
        self.named_values = {}
        self.field_writes = []

    def SetNamedSubStringValue(self, name, value):
        self.field_writes.append((name, value))
        self.named_values[name] = value

    def Close(self, save_option):
        pass

    def ExportImage(self, *args):
        self.export_image_calls.append(args)
        if self.export_error is not None:
            raise self.export_error
        if self.create_file:
            with open(args[0], "wb") as f:
                f.write(b"png")

    def ExportToFile(self, *args):
        self.export_to_file_calls.append(args)
        if self.create_file:
            with open(args[0], "wb") as f:
                f.write(b"png")


class LegacyFormat:
    """模拟 BarTender 2016 的旧 Format：没有 ExportImage 方法。"""

    def __init__(self):
        self._oleobj_ = FakeOleObj()
        self.PrintSetup = FakePrintSetup()
        self.export_to_file_calls = []

    def SetNamedSubStringValue(self, name, value):
        pass

    def ExportToFile(self, *args):
        self.export_to_file_calls.append(args)
        with open(args[0], "wb") as f:
            f.write(b"png")


class ShortSignatureFormat:
    """模拟假设中「ExportImage 存在但参数更少」的老版本 Format。

    用 7 参调用时报「非选择性的参数」(DISP_E_PARAMNOTOPTIONAL)，
    与 BarTender 12.0 实探测得的参数个数错误一致。
    """

    PARAM_NOT_OPTIONAL = -2147352561

    def __init__(self):
        self._oleobj_ = FakeOleObj()
        self.PrintSetup = FakePrintSetup()
        self.export_image_calls = []
        self.export_to_file_calls = []

    def SetNamedSubStringValue(self, name, value):
        pass

    def ExportImage(self, *args):
        self.export_image_calls.append(args)
        raise pythoncom.com_error(
            self.PARAM_NOT_OPTIONAL, "非选择性的参数。", None, None)

    def ExportToFile(self, *args):
        self.export_to_file_calls.append(args)
        with open(args[0], "wb") as f:
            f.write(b"png")


def bare_backend(fmt, **kwargs):
    # 构造函数只导入库，只有 open 才创建真实 COM 应用。
    backend = BarTenderBackend(**kwargs)
    backend._format = fmt
    backend._app = None
    return backend


def test_close_uses_do_not_save_for_format_and_application_without_prompting():
    calls = []
    fmt = SimpleNamespace(Close=lambda option: calls.append(("Format.Close", option)))
    backend = bare_backend(fmt)
    backend._app = SimpleNamespace(Quit=lambda option: calls.append(("Application.Quit", option)))

    backend.close()
    backend.close()  # 清理保持幂等，不重复关闭已释放的会话。

    # 官方 BtSaveOptions 定义：1=不保存，0=提示保存。不能依赖代码常量自证。
    assert calls == [("Format.Close", 1), ("Application.Quit", 1)]
    assert backend._format is None
    assert backend._app is None


def test_template_scan_closes_temporary_format_without_save_prompt():
    close_options = []
    fmt = SimpleNamespace(
        NamedSubStrings=[SimpleNamespace(Name="编号")],
        NamedDataSources=[],
        Close=close_options.append,
    )
    backend = bare_backend(None)
    backend._app = SimpleNamespace(Formats=SimpleNamespace(Open=lambda *args: fmt))

    assert backend.list_template_fields("synthetic-template.btw") == ["编号"]
    assert close_options == [1]


@pytest.mark.parametrize("format_fails,quit_fails", [(True, False), (False, True), (True, True)])
def test_cleanup_reports_errors_and_still_attempts_quit(format_fails, quit_fails):
    calls = []

    def close_format(option):
        calls.append(("Format.Close", option))
        if format_fails:
            raise RuntimeError("synthetic Close failure")

    def quit_app(option):
        calls.append(("Application.Quit", option))
        if quit_fails:
            raise RuntimeError("synthetic Quit failure")

    backend = bare_backend(SimpleNamespace(Close=close_format))
    backend._app = SimpleNamespace(Quit=quit_app)
    backend._print_setup = object()
    backend._field_values = {"编号": "previous"}
    backend._copies = 3

    with pytest.raises(BackendError) as error:
        backend.close()

    assert calls == [("Format.Close", 1), ("Application.Quit", 1)]
    assert ("synthetic Close failure" in str(error.value)) is format_fails
    assert ("synthetic Quit failure" in str(error.value)) is quit_fails
    assert backend._format is None and backend._app is None
    assert backend._print_setup is None and backend._field_values == {} and backend._copies is None
    backend.close()  # 释放后不重复退出，不重放已上报的清理异常。
    assert len(calls) == 2


def test_close_format_failure_is_reported_before_opening_another_template():
    calls = []

    def close_format(option):
        calls.append("close")
        raise RuntimeError("synthetic Close failure")

    backend = bare_backend(SimpleNamespace(Close=close_format))
    backend._app = SimpleNamespace(Formats=SimpleNamespace(Open=lambda *args: calls.append("open")))

    with pytest.raises(BackendError, match="synthetic Close failure"):
        backend.open("another-template.btw")

    assert calls == ["close"]
    assert backend._format is None


@pytest.mark.parametrize("scan_fails", [False, True])
def test_template_scan_cleanup_error_is_deferred_without_hiding_scan_result(monkeypatch, scan_fails):
    calls = []

    def close_format(option):
        calls.append(("Format.Close", option))
        raise RuntimeError("synthetic temporary Close failure")

    fmt = SimpleNamespace(NamedSubStrings=[SimpleNamespace(Name="编号")], Close=close_format)
    backend = bare_backend(None)
    backend._app = SimpleNamespace(
        Formats=SimpleNamespace(Open=lambda *args: fmt),
        Quit=lambda option: calls.append(("Application.Quit", option)),
    )
    if scan_fails:
        def fail_scan(_):
            raise RuntimeError("synthetic field enumeration failure")
        monkeypatch.setattr(backend, "_collect_fields", fail_scan)
        with pytest.raises(BackendError, match="synthetic field enumeration failure") as error:
            backend.list_template_fields("synthetic-template.btw")
        assert "temporary Close failure" not in str(error.value)
    else:
        assert backend.list_template_fields("synthetic-template.btw") == ["编号"]

    with pytest.raises(BackendError, match="synthetic temporary Close failure"):
        backend.close()
    assert calls == [("Format.Close", 1), ("Application.Quit", 1)]
    assert backend._app is None
    backend.close()
    assert len(calls) == 2


def test_print_out_messages_uses_invoke_types():
    backend = bare_backend(FakeFormat())
    ret = backend._invoke_print()
    assert ret == (0, "messages")
    args = backend._format._oleobj_.calls[0]
    assert args[0] == DISPID_PRINT
    assert args[2] == pythoncom.DISPATCH_METHOD
    expected_types = (
        (pythoncom.VT_BSTR, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_BOOL, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_I4, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_BYREF | pythoncom.VT_DISPATCH, pythoncom.PARAMFLAG_FOUT),
    )
    assert args[4] == expected_types
    assert args[5:] == ("", False, DEFAULT_PRINT_TIMEOUT_MS)


@pytest.mark.parametrize("timeout", [-1, 0, True, 1.5, 2_147_483_648])
def test_print_timeout_rejects_infinite_or_invalid_values(timeout):
    with pytest.raises(ValueError):
        BarTenderBackend(print_timeout_ms=timeout)


@pytest.mark.parametrize("wait_for_completion", [True, False])
def test_configured_timeout_and_job_name_are_sent_to_com(wait_for_completion):
    fmt = FakeFormat()
    backend = bare_backend(fmt, print_timeout_ms=1500, wait_for_completion=wait_for_completion)
    backend.set_job_name("job-20261110-row-9")
    backend.print_label({"编号": "TB-9"}, copies=3)
    assert fmt._oleobj_.calls[0][5:] == ("job-20261110-row-9", wait_for_completion, 1500)
    assert fmt.PrintSetup.IdenticalCopiesOfLabel == 3


@pytest.mark.parametrize("value", [0, 1, None, "0", "1", "false"])
def test_completion_wait_requires_a_real_bool(value):
    with pytest.raises(TypeError, match="wait_for_completion"):
        BarTenderBackend(wait_for_completion=value)


def test_default_job_names_are_unique_across_calls():
    fmt = FakeFormat()
    backend = bare_backend(fmt)
    backend.print_label({"编号": "A"})
    backend.print_label({"编号": "B"})
    names = [call[5] for call in fmt._oleobj_.calls]
    assert all(names)
    assert len(set(names)) == 2


def test_application_uses_dispatchex_not_existing_rot_instance(monkeypatch):
    calls = []
    app = object()

    def new_instance(progid):
        calls.append(progid)
        return app

    def existing_instance(*args):
        pytest.fail("不得调用 Dispatch 复用用户已经打开的 BarTender")

    monkeypatch.setattr(backend_module, "_import_win32com", lambda: SimpleNamespace(
        DispatchEx=new_instance, Dispatch=existing_instance))
    backend = BarTenderBackend()
    assert backend._ensure_app() is app
    assert backend._ensure_app() is app
    assert calls == ["BarTender.Application"]


@pytest.mark.parametrize("wait_for_completion", [True, False])
def test_timeout_and_com_interruption_are_unknown_without_retry(wait_for_completion):
    for result in ((1, "device busy"), RuntimeError("RPC disconnected"), (99, None)):
        fmt = FakeFormat()
        fmt._oleobj_.result = result
        backend = bare_backend(fmt, optimize_repeated_values=True,
                               wait_for_completion=wait_for_completion)
        with pytest.raises(PrintOutcomeUnknown):
            backend.print_label({"编号": "A"})
        assert len(fmt._oleobj_.calls) == 1
        assert fmt._oleobj_.calls[0][6] is wait_for_completion
        assert backend._field_values == {}
        assert backend._print_setup is None


@pytest.mark.parametrize("wait_for_completion", [True, False])
def test_explicit_bartender_failure_with_multiple_copies_is_unknown(wait_for_completion):
    fmt = FakeFormat()
    fmt._oleobj_.result = (2, "synthetic printer rejection")
    backend = bare_backend(fmt, wait_for_completion=wait_for_completion)
    with pytest.raises(PrintOutcomeUnknown, match="synthetic printer rejection") as exc:
        backend.print_label({"编号": "A"}, copies=5)
    assert "部分标签" in str(exc.value)
    assert fmt.PrintSetup.IdenticalCopiesOfLabel == 5
    assert len(fmt._oleobj_.calls) == 1


@pytest.mark.parametrize("wait_for_completion", [True, False])
def test_observer_failure_before_print_prevents_submission(wait_for_completion):
    fmt = FakeFormat()
    backend = bare_backend(fmt, wait_for_completion=wait_for_completion)

    def disk_full(event):
        if event["phase"] == "before" and event["operation"] == "print":
            raise OSError("synthetic disk full")

    backend.set_observer(disk_full)
    with pytest.raises(OSError, match="synthetic disk full"):
        backend.print_label({"编号": "A"})
    assert fmt._oleobj_.calls == []


@pytest.mark.parametrize("wait_for_completion", [True, False])
def test_observer_failure_after_print_is_unknown_and_does_not_retry(wait_for_completion):
    fmt = FakeFormat()
    backend = bare_backend(fmt, wait_for_completion=wait_for_completion)

    def disk_full(event):
        if event["phase"] == "after" and event["operation"] == "print":
            raise OSError("synthetic disk full")

    backend.set_observer(disk_full)
    with pytest.raises(PrintOutcomeUnknown, match="结果记录失败"):
        backend.print_label({"编号": "A"})
    assert len(fmt._oleobj_.calls) == 1


@pytest.mark.parametrize("wait_for_completion", [True, False])
def test_observer_has_ordered_before_after_stages_and_print_result(wait_for_completion):
    fmt = FakeFormat()
    backend = bare_backend(fmt, wait_for_completion=wait_for_completion)
    backend.set_job_name("batch-row-9")
    events = []
    backend.set_observer(events.append)
    backend.print_label({"编号": "A"})
    assert [(e["phase"], e["operation"]) for e in events] == [
        ("before", "prepare_label"), ("before", "set_field"),
        ("after", "set_field"), ("before", "set_copies"),
        ("after", "set_copies"), ("after", "prepare_label"),
        ("before", "print"), ("after", "print"),
    ]
    assert all(e["job_name"] == "batch-row-9" for e in events)
    assert events[-1]["result"] == 0
    assert events[-1]["elapsed_ms"] >= 0
    assert [e["wait_for_completion"] for e in events if e["operation"] == "print"] == [
        wait_for_completion, wait_for_completion]
    assert "value" not in events[1]  # 阶段日志默认不额外复制标签内容


def test_repeated_field_value_optimization_is_opt_in():
    fmt = FakeFormat()
    backend = bare_backend(fmt)
    backend.print_label({"编号": "A"}, 2)
    backend.print_label({"编号": "A"}, 2)
    assert fmt.field_writes == [("编号", "A"), ("编号", "A")]
    assert fmt.PrintSetup.copy_writes == [2, 2]


@pytest.mark.parametrize("mode_kwargs", [{}, {"wait_for_completion": False}])
def test_default_and_explicit_submission_modes_write_all_fields_and_copies(mode_kwargs):
    fmt = FakeFormat()
    backend = bare_backend(fmt, optimize_repeated_values=True, **mode_kwargs)
    for value in ["A", "B", "C"]:
        backend.print_label({"编号": value, "公司": "unchanged"}, copies=2)
    assert fmt.field_writes == [
        ("编号", "A"), ("公司", "unchanged"),
        ("编号", "B"), ("公司", "unchanged"),
        ("编号", "C"), ("公司", "unchanged"),
    ]
    assert fmt.PrintSetup.copy_writes == [2, 2, 2]
    assert len(fmt._oleobj_.calls) == 3
    assert all(call[6] is False for call in fmt._oleobj_.calls)


def test_opt_in_cache_skips_unchanged_values_but_keeps_every_print():
    fmt = FakeFormat()
    backend = bare_backend(fmt, optimize_repeated_values=True, wait_for_completion=True)
    backend.print_label({"编号": "A", "公司": "C"}, 2)
    backend.print_label({"编号": "B", "公司": "C"}, 2)
    assert fmt.field_writes == [("编号", "A"), ("公司", "C"), ("编号", "B")]
    assert fmt.PrintSetup.copy_writes == [2]
    assert len(fmt._oleobj_.calls) == 2


def test_cache_is_reset_when_template_is_reopened():
    fmt1, fmt2 = FakeFormat(), FakeFormat()
    backend = bare_backend(fmt1, optimize_repeated_values=True, wait_for_completion=True)
    backend.print_label({"编号": "A"}, 2)
    backend._app = SimpleNamespace(Formats=SimpleNamespace(Open=lambda *args: fmt2))
    backend.open("second.btw")
    backend.print_label({"编号": "A"}, 2)
    assert fmt2.field_writes == [("编号", "A")]
    assert fmt2.PrintSetup.copy_writes == [2]


def test_printsetup_reference_is_cached_without_skipping_assignments():
    class CountSetupFormat(FakeFormat):
        setup_reads = 0

        @property
        def PrintSetup(self):
            self.setup_reads += 1
            return self._setup_value

        @PrintSetup.setter
        def PrintSetup(self, value):
            self._setup_value = value

    fmt = CountSetupFormat()
    backend = bare_backend(fmt)
    backend.print_label({"编号": "A"})
    backend.print_label({"编号": "B"})
    assert fmt.setup_reads == 1
    assert fmt._setup_value.copy_writes == [1, 1]


def test_field_failure_clears_partial_cache_and_does_not_print(monkeypatch):
    fmt = FakeFormat()
    backend = bare_backend(fmt, optimize_repeated_values=True, wait_for_completion=True)
    real_set_field = backend._set_field

    def fail_second(name, value):
        if name == "B":
            raise BackendError("synthetic invalid field")
        real_set_field(name, value)

    monkeypatch.setattr(backend, "_set_field", fail_second)
    with pytest.raises(BackendError, match="synthetic invalid field"):
        backend.print_label({"A": "first", "B": "second"})
    assert fmt._oleobj_.calls == []
    assert backend._field_values == {}
    monkeypatch.setattr(backend, "_set_field", real_set_field)
    backend.print_label({"A": "first", "B": "second"})
    assert fmt.field_writes == [("A", "first"), ("A", "first"), ("B", "second")]


def test_messages_wraps_real_raw_dispatch_without_starting_bartender():
    from win32com.server.util import Collection, wrap

    class Message:
        _public_attrs_ = ["Message"]
        _public_methods_ = []
        Message = "SYNTHETIC_ERROR: printer unavailable"

    pythoncom.CoInitialize()
    try:
        item = wrap(Message(), iid=pythoncom.IID_IDispatch)
        raw_messages = wrap(Collection([item]), iid=pythoncom.IID_IDispatch)
        assert _messages_text((2, raw_messages)) == Message.Message
    finally:
        pythoncom.CoUninitialize()


def test_color_depth_constant_is_btcolors_24bit():
    # 4 = btColors24Bit（12.0 类型库实测）；6 是 12.x 追加的 btColors16Bit，
    # 在 BarTender 2016 (11.0) 枚举越界，报「为第一个参数指定的颜色不正确」。
    assert BT_COLORS_24BIT == 4


def test_render_preview_uses_export_image_and_out_path_basename(tmp_path):
    fmt = FakeFormat()
    backend = bare_backend(fmt)
    out = backend.render_preview({"编号": "A1"}, str(tmp_path / "preview.png"))
    # 写入调用方指定的文件名（backend 契约），而非硬编码 label.png
    assert out.endswith("preview.png")
    assert fmt.export_image_calls == [
        (str(tmp_path / "preview.png"), "png", 4, 300, 2, 16777215, True)]
    assert fmt.export_to_file_calls == []
    # 预览强制单份，避免模板里保存的多份数导致多图导出
    assert fmt.PrintSetup.IdenticalCopiesOfLabel == 1
    assert fmt.named_values == {"编号": "A1"}


def test_render_preview_falls_back_only_when_method_missing(tmp_path):
    fmt = LegacyFormat()  # 无 ExportImage，回退旧接口
    backend = bare_backend(fmt)
    out = backend.render_preview({}, str(tmp_path / "preview.png"))
    assert out.endswith("preview.png")
    assert fmt.export_to_file_calls == [(str(tmp_path / "preview.png"), "png", 4, 300, 2)]


def test_render_preview_falls_back_on_signature_mismatch(tmp_path):
    """老版本 ExportImage 参数个数不同（报「非选择性的参数」）→ 回退旧接口。"""
    fmt = ShortSignatureFormat()
    backend = bare_backend(fmt)
    out = backend.render_preview({}, str(tmp_path / "preview.png"))
    assert out.endswith("preview.png")
    # ExportImage 尝试过、被拒，随后 ExportToFile 成功
    assert fmt.export_image_calls == [
        (str(tmp_path / "preview.png"), "png", 4, 300, 2, 16777215, True)]
    assert fmt.export_to_file_calls == [(str(tmp_path / "preview.png"), "png", 4, 300, 2)]


def test_render_preview_real_error_not_masked_by_fallback(tmp_path):
    fmt = FakeFormat(export_error=RuntimeError("模板渲染失败：字段越界"))
    backend = bare_backend(fmt)
    with pytest.raises(BackendError) as err:
        backend.render_preview({}, str(tmp_path / "preview.png"))
    # 真实错误原文必须保留，且不得再调旧接口掩盖
    assert "模板渲染失败：字段越界" in str(err.value)
    assert fmt.export_to_file_calls == []


def test_render_preview_color_depth_error_not_masked(tmp_path):
    """色深被某版本拒绝（如 2016 收到 6）属真实渲染错误：原样上抛、不回退。"""
    color_error = pythoncom.com_error(
        -2147352567, "发生意外。",
        (0, "BarTend", "为第一个参数指定的颜色不正确。", None, 0, -2147418113), None)
    fmt = FakeFormat(export_error=color_error)
    backend = bare_backend(fmt)
    with pytest.raises(BackendError) as err:
        backend.render_preview({}, str(tmp_path / "preview.png"))
    assert "颜色不正确" in str(err.value)
    assert fmt.export_to_file_calls == []


def test_render_preview_missing_file_raises(tmp_path):
    fmt = FakeFormat(create_file=False)
    backend = bare_backend(fmt)
    with pytest.raises(BackendError, match="未生成图片文件"):
        backend.render_preview({}, str(tmp_path / "preview.png"))
