import time

import pytest
from PySide6.QtCore import QCoreApplication, QEvent
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import QApplication, QMessageBox

from label_printer.core.mock_backend import MockBackend
from label_printer.ui.main_window import MainWindow


@pytest.fixture(scope="session")
def qapp():
    return QApplication.instance() or QApplication([])


@pytest.fixture
def window(qapp, monkeypatch):
    monkeypatch.setattr(MainWindow, "_scan_templates", lambda self: None)
    win = MainWindow(MockBackend, MockBackend.name)
    yield win
    win.worker = None
    win._scan_worker = None
    win._preview_worker = None
    win.close()
    win.deleteLater()
    QCoreApplication.sendPostedEvents(None, QEvent.DeferredDelete)
    qapp.processEvents()


class RunningWorker:
    def __init__(self):
        self.cancelled = False
        self.backend = None
    def isRunning(self):
        return True
    def cancel(self):
        self.cancelled = True
    def wait(self, timeout):
        return False


def test_printing_locks_inputs_and_guards_programmatic_paths(window):
    window._preview_timer.start()
    window._set_printing_state(True)
    assert not window._preview_timer.isActive()
    assert not window.acceptDrops()
    assert not window.template_combo.isEnabled()
    assert not window.data_table.isEnabled()
    assert not window.map_table.isEnabled()
    assert not window.printer_combo.isEnabled()
    window._load_excel("must-not-be-opened.xlsx")
    window._add_template("must-not-be-added.btw")
    window._refresh_preview()
    window._schedule_preview()
    assert window.sheet is None and window.template_combo.count() == 0
    assert not window._preview_timer.isActive()
    window._set_printing_state(False)
    assert window.acceptDrops() and window.template_combo.isEnabled()


def test_start_rejects_active_preview_before_building_jobs(window, monkeypatch):
    messages = []
    monkeypatch.setattr(QMessageBox, "information", lambda *args: messages.append(args[2]))
    monkeypatch.setattr(window, "_build_jobs", lambda _: pytest.fail("must wait for preview"))
    window._preview_worker = RunningWorker()
    window._start_print(False)
    assert messages and window.worker is None


def test_close_rejects_unfinished_thread_instead_of_destroying(window, monkeypatch):
    warnings = []
    monkeypatch.setattr(QMessageBox, "warning", lambda *args: warnings.append(args[2]))
    worker = RunningWorker()
    window.worker = worker
    event = QCloseEvent()
    window.closeEvent(event)
    assert worker.cancelled
    assert not event.isAccepted()
    assert warnings and "窗口暂时保留" in warnings[0]


def test_done_signal_does_not_release_edit_controls_before_finished(window):
    window._set_printing_state(True)
    window._on_done(True, "提交完成，出纸请核对", [])
    assert window._printing
    assert not window.template_combo.isEnabled()
    assert window._last_print_result[0] is True


def test_confirmation_dialog_cannot_start_print_over_new_preview(window, monkeypatch):
    messages = []
    monkeypatch.setattr(window, "_build_jobs", lambda _: [(1, {"id": "A"}, 1)])
    monkeypatch.setattr(window, "_current_mappings", lambda: [])
    window.template_combo.blockSignals(True)
    window.template_combo.addItem("template", "fixture.btw")
    window.template_combo.blockSignals(False)
    def confirm(*args):
        window._preview_worker = RunningWorker()
        return QMessageBox.Yes
    monkeypatch.setattr(QMessageBox, "question", confirm)
    monkeypatch.setattr(QMessageBox, "information", lambda *args: messages.append(args[2]))
    window._start_print(False)
    assert window.worker is None
    assert messages and "刚刚启动" in messages[0]


def test_unreconciled_batch_requires_explicit_acknowledgement(window, monkeypatch):
    questions = []
    window._set_reconcile_required(True, "previous-batch.jsonl")
    monkeypatch.setattr(window, "_build_jobs", lambda _: [(1, {"id": "A"}, 1)])
    monkeypatch.setattr(window, "_current_mappings", lambda: [])
    window.template_combo.blockSignals(True)
    window.template_combo.addItem("template", "fixture.btw")
    window.template_combo.blockSignals(False)
    monkeypatch.setattr(QMessageBox, "question",
                        lambda *args: questions.append(args[2]) or QMessageBox.No)
    window._start_print(True)
    assert window.worker is None and window._needs_reconcile
    assert "previous-batch.jsonl" in questions[0]
    assert "不要直接重打" in questions[0]


def test_missing_worker_result_keeps_reconciliation_marker(window, monkeypatch):
    monkeypatch.setattr(QMessageBox, "warning", lambda *args: None)
    window._set_reconcile_required(True, "incomplete.jsonl")
    window._set_printing_state(True)
    worker = object()
    window.worker = worker
    monkeypatch.setattr(window, "sender", lambda: worker)
    window._last_print_result = None
    window._clear_print_worker()
    assert window.worker is None
    assert window._needs_reconcile
    assert str(window._settings.value("needsReconcile")).lower() == "true"


def test_real_qt_completion_unlocks_after_finished_and_preserves_log(window, qapp,
                                                                   monkeypatch, tmp_path):
    monkeypatch.setenv("LABEL_PRINTER_HOME", str(tmp_path / "state"))
    monkeypatch.setattr(window, "_build_jobs", lambda _: [(1, {"id": "A"}, 1)])
    monkeypatch.setattr(window, "_current_mappings", lambda: [])
    window.template_combo.blockSignals(True)
    window.template_combo.addItem("template", "fixture.btw")
    window.template_combo.blockSignals(False)
    window._start_print(True)
    worker = window.worker
    assert worker is not None and window._printing
    summary = worker.summary_path
    deadline = time.monotonic() + 5
    while window.worker is not None and time.monotonic() < deadline:
        qapp.processEvents()
        time.sleep(0.01)
    assert window.worker is None
    assert not window._printing and window.template_combo.isEnabled()
    assert not window._needs_reconcile
    assert summary.exists()
    assert "提交完成" in window.log_view.toPlainText()
