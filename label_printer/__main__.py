"""入口：python -m label_printer [--mock]"""
from __future__ import annotations

import argparse
import sys
from functools import partial

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()

from .core.runtime import configure_runtime
configure_runtime()

# pywin32 加载顺序坑：若 win32print 在主线程先导入，之后其他线程再 import
# win32com 会因 pywintypes 初始化时序报「DLL load failed …找不到指定的程序」
# （仅源码运行环境如此，PyInstaller 打包不受影响）。这里统一先加载 win32com。
try:
    import pywintypes  # noqa: F401
    import win32com.client  # noqa: F401
    import win32print  # noqa: F401
except Exception:
    pass  # 无 pywin32 时由后端给出中文提示；mock 模式也不需要


def main() -> int:
    parser = argparse.ArgumentParser(description="标签自动打印工具")
    parser.add_argument(
        "--mock", action="store_true",
        help="干跑模式：不连接 BarTender，打印动作只写日志（调试用）")
    parser.add_argument("--self-test", metavar="OUTPUT_DIR",
                        help="运行本地故障自检并写入结果（不会连接真实打印机）")
    parser.add_argument("--ui-self-test", metavar="OUTPUT_DIR",
                        help="离屏校验界面、Excel 与模拟提交（不会连接真实打印机）")
    args = parser.parse_args()

    if args.self_test:
        from .core.diagnostics import run_self_test
        return run_self_test(args.self_test)

    if args.ui_self_test:
        from .core.gui_diagnostics import run_gui_self_test
        return run_gui_self_test(args.ui_self_test)

    wait_for_completion = False
    if args.mock:
        from label_printer.core.mock_backend import MockBackend
        backend_factory = MockBackend
        backend_name = MockBackend.name
    else:
        try:
            from label_printer.core.isolated_backend import (
                IsolatedBarTenderBackend, wait_for_completion_from_env)
            wait_for_completion = wait_for_completion_from_env()
            backend_factory = partial(IsolatedBarTenderBackend,
                                      wait_for_completion=wait_for_completion)
            backend_name = ("BarTender（完成反馈等待，排查用）" if wait_for_completion
                            else "BarTender（提交模式）")
        except Exception as exc:
            # BarTender 不可用时给个选择：直接退出或退回干跑模式
            from PySide6.QtWidgets import QApplication, QMessageBox
            from label_printer.ui.theme import apply_theme
            app = QApplication(sys.argv)
            apply_theme(app)
            ret = QMessageBox.warning(
                None, "BarTender 不可用",
                f"无法连接 BarTender：\n{exc}\n\n"
                "是否以「干跑模式」启动（只演示流程，不真正打印）？",
                QMessageBox.Yes | QMessageBox.No)
            if ret != QMessageBox.Yes:
                return 1
            from label_printer.core.mock_backend import MockBackend
            backend_factory = MockBackend
            backend_name = MockBackend.name

    from PySide6.QtWidgets import QApplication
    from label_printer.ui.main_window import MainWindow
    from label_printer.ui.theme import apply_theme

    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName("标签自动打印工具")
    apply_theme(app)
    window = MainWindow(backend_factory, backend_name,
                        wait_for_completion=wait_for_completion)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
