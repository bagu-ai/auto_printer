"""pytest 根配置：保证不设 PYTHONPATH 也能导入包；Qt 测试默认离屏。"""
import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


import pytest  # noqa: E402


@pytest.fixture(autouse=True)
def _isolate_qsettings(tmp_path, monkeypatch):
    """QSettings 重定向到临时文件：测试不得读写真实用户注册表。

    注：setDefaultFormat 在 Windows/PySide6 下不影响 QSettings(org, app)
    双参数构造（仍走注册表），故直接替换 main_window 里的 QSettings 名称。
    """
    from PySide6.QtCore import QSettings
    ini = tmp_path / "settings.ini"
    monkeypatch.setattr(
        "label_printer.ui.main_window.QSettings",
        lambda *args, **kwargs: QSettings(str(ini), QSettings.IniFormat))
    yield tmp_path
