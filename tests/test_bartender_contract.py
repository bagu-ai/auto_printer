"""BarTender ActiveX 调用契约测试（无需安装/启动 BarTender）。

锁定含 [out] Messages 的两个方法必须通过 InvokeTypes 调用，防止以后又退回
`method(..., None)` 导致 COM「类型不匹配」。
"""
import pythoncom

from label_printer.core.bartender_backend import (
    BarTenderBackend, DISPID_EXPORT_PRINT_PREVIEW, DISPID_PRINT,
)


class FakeOleObj:
    def __init__(self, result=(0, "messages")):
        self.result = result
        self.calls = []

    def InvokeTypes(self, *args):
        self.calls.append(args)
        return self.result


class FakeFormat:
    def __init__(self):
        self._oleobj_ = FakeOleObj()


def bare_backend():
    backend = object.__new__(BarTenderBackend)
    backend._format = FakeFormat()
    backend._app = None
    return backend


def test_print_out_messages_uses_invoke_types():
    backend = bare_backend()
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
    assert args[5:] == ("", True, -1)


def test_preview_out_messages_uses_invoke_types(tmp_path):
    backend = bare_backend()
    ret = backend._invoke_preview(str(tmp_path))
    assert ret == (0, "messages")
    args = backend._format._oleobj_.calls[0]
    assert args[0] == DISPID_EXPORT_PRINT_PREVIEW
    assert args[2] == pythoncom.DISPATCH_METHOD
    expected_types = (
        (pythoncom.VT_BSTR, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_BSTR, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_BSTR, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_I4, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_I4, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_I4, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_I4, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_BOOL, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_BOOL, pythoncom.PARAMFLAG_FIN),
        (pythoncom.VT_BYREF | pythoncom.VT_DISPATCH, pythoncom.PARAMFLAG_FOUT),
    )
    assert args[4] == expected_types
    assert args[5:] == (
        str(tmp_path), "label_%PageNumber%.png", "png",
        6, 300, 16777215, 2, False, False,
    )
