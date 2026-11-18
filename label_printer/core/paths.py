"""路径约定：只读资源目录 vs 可写状态目录。

- resource_dir()：模板等随程序分发的只读资源（exe 同目录 / 项目根）
- state_dir()：程序目录下运行数据，可用 LABEL_PRINTER_HOME 覆盖。
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

APP_NAME = "标签自动打印工具"


def resource_dir() -> Path:
    """只读资源目录：打包后是 exe 所在目录，开发时是项目根。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[2]


def state_dir() -> Path:
    """可写状态目录（profiles/logs/previews 的父目录）。"""
    override = os.environ.get("LABEL_PRINTER_HOME")
    if override:
        return Path(override)
    # Portable distribution: configurable state stays next to the D-drive app.
    # Do not migrate or mutate an earlier version's APPDATA profiles implicitly.
    return resource_dir() / "运行数据"


def profiles_dir() -> Path:
    return state_dir() / "profiles"


def logs_dir() -> Path:
    return state_dir() / "logs"


def previews_dir() -> Path:
    return state_dir() / "previews"
