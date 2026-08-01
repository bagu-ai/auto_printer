"""BarTender ActiveX 调用契约测试（无需安装/启动 BarTender）。

- 锁定含 [out] Messages 的 Format.Print 必须通过 InvokeTypes 调用，防止以后
  又退回 `method(..., None)` 导致 COM「类型不匹配」。
- 锁定 render_preview 的新实现：ExportImage 参数、方法缺失时回退
  ExportToFile、真实错误不被回退掩盖、写入调用方指定的 out_path 文件名。
"""
import pythoncom
import pytest

from label_printer.core.backend import BackendError
from label_printer.core.bartender_backend import BarTenderBackend, DISPID_PRINT


class FakeOleObj:
    def __init__(self, result=(0, "messages")):
        self.result = result
        self.calls = []

    def InvokeTypes(self, *args):
        self.calls.append(args)
        return self.result


class FakePrintSetup:
    def __init__(self):
        self.IdenticalCopiesOfLabel = 7  # 模板里保存的非 1 份数


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

    def SetNamedSubStringValue(self, name, value):
        self.named_values[name] = value

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


def bare_backend(fmt):
    backend = object.__new__(BarTenderBackend)
    backend._format = fmt
    backend._app = None
    return backend


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
    assert args[5:] == ("", True, -1)


def test_render_preview_uses_export_image_and_out_path_basename(tmp_path):
    fmt = FakeFormat()
    backend = bare_backend(fmt)
    out = backend.render_preview({"编号": "A1"}, str(tmp_path / "preview.png"))
    # 写入调用方指定的文件名（backend 契约），而非硬编码 label.png
    assert out.endswith("preview.png")
    assert fmt.export_image_calls == [
        (str(tmp_path / "preview.png"), "png", 6, 300, 2, 16777215, True)]
    assert fmt.export_to_file_calls == []
    # 预览强制单份，避免模板里保存的多份数导致多图导出
    assert fmt.PrintSetup.IdenticalCopiesOfLabel == 1
    assert fmt.named_values == {"编号": "A1"}


def test_render_preview_falls_back_only_when_method_missing(tmp_path):
    fmt = LegacyFormat()  # 无 ExportImage，回退旧接口
    backend = bare_backend(fmt)
    out = backend.render_preview({}, str(tmp_path / "preview.png"))
    assert out.endswith("preview.png")
    assert fmt.export_to_file_calls == [(str(tmp_path / "preview.png"), "png", 6, 300, 2)]


def test_render_preview_real_error_not_masked_by_fallback(tmp_path):
    fmt = FakeFormat(export_error=RuntimeError("模板渲染失败：字段越界"))
    backend = bare_backend(fmt)
    with pytest.raises(BackendError) as err:
        backend.render_preview({}, str(tmp_path / "preview.png"))
    # 真实错误原文必须保留，且不得再调旧接口掩盖
    assert "模板渲染失败：字段越界" in str(err.value)
    assert fmt.export_to_file_calls == []


def test_render_preview_missing_file_raises(tmp_path):
    fmt = FakeFormat(create_file=False)
    backend = bare_backend(fmt)
    with pytest.raises(BackendError, match="未生成图片文件"):
        backend.render_preview({}, str(tmp_path / "preview.png"))
