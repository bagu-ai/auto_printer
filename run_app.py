"""PyInstaller 打包入口（同 `python -m label_printer`，支持 --mock 参数）。"""
if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    from label_printer.__main__ import main
    raise SystemExit(main())
