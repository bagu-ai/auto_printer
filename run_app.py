"""PyInstaller 打包入口（同 `python -m label_printer`，支持 --mock 参数）。"""
from label_printer.__main__ import main

if __name__ == "__main__":
    raise SystemExit(main())
