"""直接入口冻结打印模式；替换 Qt 窗口和事件循环，不启动 BarTender。"""
import importlib
from functools import partial

import pytest
from PySide6 import QtWidgets

from label_printer.core import runtime
from label_printer.core.isolated_backend import IsolatedBarTenderBackend
from label_printer.ui import main_window, theme


@pytest.mark.parametrize("env_value,expected_mode,title_fragment", [
    (None, False, "提交模式"),
    ("1", True, "排查用"),
])
def test_direct_entry_freezes_default_or_explicit_wait_mode(
        monkeypatch, env_value, expected_mode, title_fragment):
    monkeypatch.delenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", raising=False)
    if env_value is not None:
        monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", env_value)
    monkeypatch.setattr(runtime, "configure_runtime", lambda: None)
    entry = importlib.import_module("label_printer.__main__")
    monkeypatch.setattr(entry.sys, "argv", ["label_printer.exe"])
    captured = {}

    class FakeApplication:
        @staticmethod
        def instance():
            return None

        def __init__(self, argv):
            captured["argv"] = list(argv)

        def setApplicationName(self, name):
            captured["application_name"] = name

        def exec(self):
            return 17

    class FakeWindow:
        def __init__(self, backend_factory, backend_name, *, wait_for_completion):
            captured.update(factory=backend_factory, backend_name=backend_name,
                            wait_for_completion=wait_for_completion)

        def show(self):
            captured["shown"] = True

    monkeypatch.setattr(QtWidgets, "QApplication", FakeApplication)
    monkeypatch.setattr(QtWidgets.QMessageBox, "warning",
                        lambda *args: pytest.fail("valid mode must not show a fallback dialog"))
    monkeypatch.setattr(theme, "apply_theme", lambda app: captured.update(theme_applied=True))
    monkeypatch.setattr(main_window, "MainWindow", FakeWindow)
    monkeypatch.setattr(IsolatedBarTenderBackend, "_start",
                        lambda self: pytest.fail("entry test must not create a backend process"))

    assert entry.main() == 17
    assert captured["shown"] and captured["theme_applied"]
    assert title_fragment in captured["backend_name"]
    assert captured["wait_for_completion"] is expected_mode
    factory = captured["factory"]
    assert isinstance(factory, partial)
    assert factory.func is IsolatedBarTenderBackend
    assert factory.keywords["wait_for_completion"] is expected_mode

    # 入口之后环境变化也不能改变该窗口创建的实际后端模式。
    monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", "1" if not expected_mode else "0")
    backend = factory()
    assert backend.wait_for_completion is expected_mode
    assert backend.child_pid is None
