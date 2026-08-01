"""入口：python -m label_printer [--mock]"""
from __future__ import annotations

import argparse
import sys


def main() -> int:
    parser = argparse.ArgumentParser(description="标签自动打印工具")
    parser.add_argument(
        "--mock", action="store_true",
        help="干跑模式：不连接 BarTender，打印动作只写日志（调试用）")
    args = parser.parse_args()

    if args.mock:
        from label_printer.core.mock_backend import MockBackend
        backend = MockBackend()
    else:
        try:
            from label_printer.core.bartender_backend import BarTenderBackend
            backend = BarTenderBackend()
        except Exception as exc:
            # BarTender 不可用时给个选择：直接退出或退回干跑模式
            from PySide6.QtWidgets import QApplication, QMessageBox
            app = QApplication(sys.argv)
            ret = QMessageBox.warning(
                None, "BarTender 不可用",
                f"无法连接 BarTender：\n{exc}\n\n"
                "是否以「干跑模式」启动（只演示流程，不真正打印）？",
                QMessageBox.Yes | QMessageBox.No)
            if ret != QMessageBox.Yes:
                return 1
            from label_printer.core.mock_backend import MockBackend
            backend = MockBackend()

    from PySide6.QtWidgets import QApplication
    from label_printer.ui.main_window import MainWindow

    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName("标签自动打印工具")
    window = MainWindow(backend)
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
