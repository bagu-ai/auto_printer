"""Offscreen end-to-end GUI diagnostic using only the in-memory mock backend."""
from __future__ import annotations

import json
import os
import time
import traceback
from pathlib import Path


def run_gui_self_test(output_dir) -> int:
    """Exercise packaged Excel/Qt/preview/worker/journal paths without a printer.

    Run in a dedicated process: this function creates and destroys QApplication.
    All diagnostic data, settings, logs, and pictures stay under output_dir.
    """
    output = Path(output_dir).expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    temporary = output / ".tmp"
    temporary.mkdir(exist_ok=True)
    env_names = ("QT_QPA_PLATFORM", "LABEL_PRINTER_HOME", "TEMP", "TMP")
    saved_env = {name: os.environ.get(name) for name in env_names}
    os.environ.update(QT_QPA_PLATFORM="offscreen",
                      LABEL_PRINTER_HOME=str(output / "state"),
                      TEMP=str(temporary), TMP=str(temporary))
    result = {
        "success": False,
        "scope": "Offscreen GUI, generated XLSX, MockBackend only; no BarTender or printer connection",
        "output_dir": str(output),
        "expected_rows": 20,
        "dialogs": [],
        "qt_destroyed": False,
    }
    app = None
    window = None
    originals = {}
    ui = None
    message_box = None
    try:
        import pandas as pd
        import shiboken6
        from PySide6.QtCore import QCoreApplication, QEvent
        from PySide6.QtGui import QFont, QFontDatabase, QFontMetrics
        from PySide6.QtWidgets import QApplication, QMessageBox

        from .mock_backend import MOCK_FIELDS, MockBackend
        from .paths import resource_dir
        from ..ui import main_window as ui
        from ..ui.theme import apply_theme

        if QApplication.instance() is not None:
            raise RuntimeError("GUI self-test must run in a separate process with no existing QApplication")
        app = QApplication([])
        # Windows' offscreen plugin may not populate its font database. Load
        # existing system fonts explicitly for this diagnostic process only.
        # Product theme and the normal Windows application stay unchanged.
        font_directory = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
        font_records = []
        loaded_families = []
        for font_file in ("msyh.ttc", "msyhbd.ttc", "segoeui.ttf", "arial.ttf", "consola.ttf"):
            path = font_directory / font_file
            if not path.is_file():
                font_records.append({"path": str(path), "exists": False})
                continue
            font_id = QFontDatabase.addApplicationFont(str(path))
            families = list(QFontDatabase.applicationFontFamilies(font_id)) if font_id >= 0 else []
            font_records.append({"path": str(path), "exists": True,
                                 "font_id": font_id, "families": families})
            loaded_families.extend(families)
        chosen = next((family for family in ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI")
                       if family in loaded_families), None)
        result["font_loading"] = font_records
        if chosen is None:
            raise RuntimeError("Offscreen diagnostic could not load a usable system font")
        app.setFont(QFont(chosen, 10))
        metrics = QFontMetrics(app.font())
        glyph_checks = {character: metrics.inFontUcs4(ord(character)) for character in "标签打印ABC123"}
        result["diagnostic_font"] = {"family": chosen, "glyph_checks": glyph_checks}
        if not all(glyph_checks.values()):
            raise RuntimeError("Offscreen diagnostic font is missing required Chinese or Latin glyphs")
        apply_theme(app)
        message_box = QMessageBox

        def capture_dialog(kind, args):
            result["dialogs"].append({"kind": kind,
                                      "title": str(args[1]) if len(args) > 1 else "",
                                      "text": str(args[2]) if len(args) > 2 else ""})

        def answer_yes(*args, **kwargs):
            capture_dialog("question", args)
            return QMessageBox.Yes

        def record_message(kind):
            def record(*args, **kwargs):
                capture_dialog(kind, args)
                return QMessageBox.Ok
            return record

        for name in ("question", "warning", "critical", "information"):
            originals[name] = getattr(QMessageBox, name)
            setattr(QMessageBox, name, answer_yes if name == "question" else record_message(name))
        originals["printers"] = ui.MainWindow._list_printers_safe
        originals["co_init"] = ui._co_init
        # Do not enumerate real Windows printers or initialize a COM apartment.
        ui.MainWindow._list_printers_safe = lambda self: MockBackend().list_printers()
        ui._co_init = lambda: None

        sample = output / "gui-self-test-20rows.xlsx"
        pd.DataFrame({
            "发货新编号": [f"DIAG-{i:04d}" for i in range(1, 21)],
            "客户香型": [f"Mock flavor {i}" for i in range(1, 21)],
            "发货粒径": ["2.8mm"] * 20,
            "英文": [f"Diagnostic sample {i}" for i in range(1, 21)],
            "胶囊原编号": [f"SOURCE-{i:04d}" for i in range(1, 21)],
            "送样数量": ["1瓶"] * 20,
        }).to_excel(sample, index=False)
        result["excel"] = str(sample)

        templates = list((resource_dir() / "模板").glob("*.btw"))
        if not templates:
            raise AssertionError("Bundled template directory has no .btw files")
        result["bundled_templates"] = [str(path) for path in templates]
        from .isolated_backend import wait_for_completion_from_env
        wait_for_completion = wait_for_completion_from_env()
        result["wait_for_completion"] = wait_for_completion
        mode_name = MockBackend.name if wait_for_completion else "模拟后端（提交模式）"
        window = ui.MainWindow(MockBackend, mode_name,
                               wait_for_completion=wait_for_completion)
        window.resize(1280, 860)
        window.show()
        window._load_excel(str(sample))

        def pump_until(predicate, label, timeout=15):
            deadline = time.monotonic() + timeout
            while time.monotonic() < deadline:
                app.processEvents()
                if predicate():
                    return
                time.sleep(0.01)
            raise TimeoutError(f"GUI self-test timeout: {label}")

        def background_idle():
            return (not window._worker_running(window._scan_worker) and
                    not window._scan_pending and
                    not window._worker_running(window._preview_worker) and
                    not window._preview_pending and
                    not window._preview_timer.isActive())

        pump_until(lambda: bool(window.template_fields) and background_idle(),
                   "template scan, mapping, and initial preview")
        if window.sheet is None or len(window.sheet) != 20:
            raise AssertionError("Generated XLSX did not load as 20 rows")
        mappings = {mapping.field: mapping.value for mapping in window._current_mappings()}
        if any(mappings.get(field) != field for field in MOCK_FIELDS):
            raise AssertionError(f"Named-field auto mapping failed: {mappings}")
        result["mappings"] = mappings

        window.data_table.selectRow(0)
        window._preview_timer.stop()
        window._refresh_preview()
        pump_until(lambda: bool(window._latest_preview_path) and background_idle(),
                   "mock label preview")
        preview = Path(window._latest_preview_path)
        if not preview.is_file() or preview.stat().st_size < 100:
            raise AssertionError("Mock preview PNG missing or empty")
        result["preview_png"] = str(preview)

        window._start_print(test_only=False)
        worker = window.worker
        if worker is None:
            raise AssertionError("GUI did not start PrintWorker")
        summary = worker.summary_path
        journal = worker.journal_path
        batch_id = worker.batch_id
        result["batch_id"] = batch_id
        pump_until(lambda: window.worker is None and not window._printing,
                   "actual QThread.finished and control unlock", timeout=30)
        record = json.loads(summary.read_text(encoding="utf-8"))
        if record.get("wait_for_completion") is not wait_for_completion:
            raise AssertionError("Batch summary does not preserve the selected print mode")
        expected_mode = "standard" if wait_for_completion else "submit_compatibility"
        if record.get("print_mode") != expected_mode:
            raise AssertionError("Batch mode name disagrees with the selected mode")
        if not wait_for_completion and "提交模式" not in window.windowTitle():
            raise AssertionError("Compatibility mode must remain visible in the window title")
        rows = record["rows"]
        if not record["success"] or len(rows) != 20:
            raise AssertionError(f"Unexpected batch result: {record['message']}")
        if any(row["status"] != "submitted" or row["copies"] != 1 for row in rows):
            raise AssertionError("Expected exactly 20 submitted rows with one copy each")
        if [row["values"]["发货新编号"] for row in rows] != [f"DIAG-{i:04d}" for i in range(1, 21)]:
            raise AssertionError("Submitted identifiers have a duplicate, missing row, or wrong order")
        if window._needs_reconcile or not window.template_combo.isEnabled():
            raise AssertionError("Successful finished batch did not restore the GUI state")
        events = [json.loads(line) for line in journal.read_text(encoding="utf-8").splitlines()]
        if events[0]["snapshot"].get("wait_for_completion") is not wait_for_completion:
            raise AssertionError("Durable batch start does not preserve the selected print mode")
        if sum(event["event"] == "row_submitted" for event in events) != 20:
            raise AssertionError("Durable journal does not contain all submitted rows")
        if events[-1]["event"] != "batch_end" or not events[-1]["success"]:
            raise AssertionError("Durable journal does not finish successfully")
        result.update(summary_json=str(summary), journal_jsonl=str(journal),
                      submitted_rows=len(rows), submitted_copies=sum(row["copies"] for row in rows),
                      gui_controls_restored=True, finished_signal_observed=True)

        screenshot = output / "gui-self-test-window.png"
        app.processEvents()
        pixmap = window.grab()
        if not pixmap.save(str(screenshot), "PNG"):
            raise AssertionError("Cannot save GUI screenshot")
        result.update(window_png=str(screenshot),
                      window_png_size=[pixmap.width(), pixmap.height()])
        if any(dialog["kind"] != "question" for dialog in result["dialogs"]):
            raise AssertionError("Unexpected GUI warning/error during diagnostic")
        result["success"] = True
    except Exception as exc:
        result.update(error=str(exc), traceback=traceback.format_exc())
    finally:
        try:
            if window is not None:
                window._preview_timer.stop()
                if not window.close():
                    raise RuntimeError("GUI refused to close because a worker was still active")
                window.deleteLater()
                QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
                app.processEvents()
            if ui is not None and "printers" in originals:
                ui.MainWindow._list_printers_safe = originals["printers"]
                ui._co_init = originals["co_init"]
            if message_box is not None:
                for name in ("question", "warning", "critical", "information"):
                    if name in originals:
                        setattr(message_box, name, originals[name])
            if app is not None:
                app.quit()
                shiboken6.delete(app)
                result["qt_destroyed"] = True
        except Exception as exc:
            result["success"] = False
            result["cleanup_error"] = str(exc)
        for name, previous in saved_env.items():
            if previous is None:
                os.environ.pop(name, None)
            else:
                os.environ[name] = previous
        (output / "gui-self-test-result.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if result["success"] else 1
