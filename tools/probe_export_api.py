"""探针：在本机 BarTender 上逐一试验导出图片接口的不同签名/色深语义。

目的：客户机 BarTender 2016 (11.0.1.3045) 调用
    ExportImage(path, "png", 6, 300, 2, 16777215, True)
报「为第一个参数指定的颜色不正确」。需要弄清：
  1) 12.0 的 ExportImage 是否拒绝位深写法 24（验证 6 是枚举而非位深）
  2) 12.0 是否有 ExportToFile、它接受 24 还是 6
  3) app.Version / app.Build 返回什么（用于版本分流与日志诊断）

用法：python tools/probe_export_api.py
只读探测，不打印、不改动模板。
"""
import os
import sys
import tempfile

import pythoncom
import win32com.client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(ROOT, "模板", "中文版无云南巴菰公司信息标签模板（60-30mm）.btw")

OUT_DIR = tempfile.mkdtemp(prefix="bt_probe_")


def probe(label, fn):
    try:
        result = fn()
        print(f"[ OK ] {label} -> {result!r}")
    except Exception as exc:  # noqa: BLE001 - 探针就是要看原始错误
        print(f"[FAIL] {label}\n       {exc!r}")


def main():
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("BarTender.Application")
        probe("app.Version", lambda: app.Version)
        probe("app.Build", lambda: getattr(app, "Build", "<无此属性>"))
        fmt = app.Formats.Open(TEMPLATE, True, "")
        try:
            fmt.PrintSetup.IdenticalCopiesOfLabel = 1
            has_export_image = getattr(fmt, "ExportImage", None) is not None
            has_export_to_file = getattr(fmt, "ExportToFile", None) is not None
            print(f"[info] ExportImage 存在: {has_export_image}  "
                  f"ExportToFile 存在: {has_export_to_file}")

            out = os.path.join(OUT_DIR, "probe.png")

            def _export_image(*args):
                def _call():
                    fmt.ExportImage(*args)
                    return os.path.exists(out)
                return _call

            def _export_to_file(*args):
                def _call():
                    fmt.ExportToFile(*args)
                    return os.path.exists(out)
                return _call

            probe('ExportImage(out,"png",6,300,2,16777215,True) [12.0 已知可用]',
                  _export_image(out, "png", 6, 300, 2, 16777215, True))
            probe('ExportImage(out,"png",24,300,2,16777215,True) [位深写法]',
                  _export_image(out, "png", 24, 300, 2, 16777215, True))
            probe('ExportImage(out,"png",6,300,2) [仅5参]',
                  _export_image(out, "png", 6, 300, 2))
            if has_export_to_file:
                probe('ExportToFile(out,"png",24,300,True) [位深24+布尔覆盖]',
                      _export_to_file(out, "png", 24, 300, True))
                probe('ExportToFile(out,"png",6,300,True) [枚举6+布尔覆盖]',
                      _export_to_file(out, "png", 6, 300, True))
                probe('ExportToFile(out,"png",6,300,2) [当前代码的回退写法]',
                      _export_to_file(out, "png", 6, 300, 2))
        finally:
            fmt.Close(0)
            app.Quit(0)
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    sys.exit(main())
