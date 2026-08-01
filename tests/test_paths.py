"""路径约定测试：state_dir 可被环境变量覆盖（测试不污染真实 APPDATA）。"""
import os

from label_printer.core.paths import logs_dir, previews_dir, profiles_dir, resource_dir, state_dir


def test_state_dir_env_override(tmp_path, monkeypatch):
    monkeypatch.setenv("LABEL_PRINTER_HOME", str(tmp_path))
    assert state_dir() == tmp_path
    assert profiles_dir() == tmp_path / "profiles"
    assert logs_dir() == tmp_path / "logs"
    assert previews_dir() == tmp_path / "previews"


def test_state_dir_defaults_to_appdata(monkeypatch):
    monkeypatch.delenv("LABEL_PRINTER_HOME", raising=False)
    monkeypatch.setenv("APPDATA", r"C:\Users\Test\AppData\Roaming")
    assert state_dir() == r"C:\Users\Test\AppData\Roaming\标签自动打印工具" or \
           str(state_dir()).endswith("标签自动打印工具")


def test_resource_dir_exists():
    # 开发模式下是项目根，应包含 label_printer 包
    assert (resource_dir() / "label_printer").is_dir()
