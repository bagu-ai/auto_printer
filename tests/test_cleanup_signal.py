"""清理失败单独通知；读取字段和校样的业务结果不被清理结果覆盖。"""
from types import SimpleNamespace

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from label_printer.core.backend import BackendError
from label_printer.ui.main_window import _BackendWorker, PreviewWorker, TemplateScanWorker


@pytest.fixture(scope="session")
def qapp():
    return QApplication.instance() or QApplication([])


class CleanupBackend:
    def __init__(self, mode, operation_fails=False):
        self.mode = mode
        self.operation_fails = operation_fails
        self.cleanup_warning = ""
        self.closed_count = 0

    def list_template_fields(self, template_path):
        if self.operation_fails:
            raise BackendError("synthetic scan failure")
        return ["编号", "香型："]

    def open(self, template_path, printer):
        self.opened = (template_path, printer)

    def render_preview(self, values, out_path):
        if self.operation_fails:
            raise BackendError("synthetic preview failure")
        return out_path

    def close(self):
        self.closed_count += 1
        if self.mode in {"warning", "both"}:
            self.cleanup_warning = "synthetic supervisor cleanup warning"
        if self.mode in {"exception", "both"}:
            raise BackendError("synthetic Close/Quit exception")


def assert_cleanup_notice(mode, notices):
    if mode == "success":
        assert notices == []
        return
    assert len(notices) == 1
    if mode in {"exception", "both"}:
        assert "synthetic Close/Quit exception" in notices[0]
    if mode in {"warning", "both"}:
        assert "synthetic supervisor cleanup warning" in notices[0]


@pytest.mark.parametrize("mode", ["success", "exception", "warning", "both"])
def test_base_teardown_reports_cleanup_and_always_uninitializes_com(qapp, mode):
    backend = CleanupBackend(mode)
    worker = _BackendWorker(lambda: backend)
    worker.backend = backend
    uninitialized = []
    worker._pythoncom = SimpleNamespace(CoUninitialize=lambda: uninitialized.append(True))
    notices = []
    worker.cleanup_failed.connect(notices.append, Qt.DirectConnection)

    result = worker._teardown()

    assert backend.closed_count == 1
    assert uninitialized == [True]
    assert_cleanup_notice(mode, notices)
    if mode == "success":
        assert result == ""
    else:
        for detail in ("synthetic Close/Quit exception", "synthetic supervisor cleanup warning"):
            if detail in notices[0]:
                assert detail in result


def make_worker(kind, mode, out_path, operation_fails=False):
    created = []

    def factory():
        backend = CleanupBackend(mode, operation_fails)
        created.append(backend)
        return backend

    if kind == "scan":
        worker = TemplateScanWorker(factory, "synthetic-template.btw", 17)
        ready_signal = worker.fields_ready
    else:
        worker = PreviewWorker(factory, "synthetic-template.btw", "synthetic-printer",
                               {"编号": "TB-1"}, str(out_path))
        ready_signal = worker.preview_ready
    results, failures, notices = [], [], []
    ready_signal.connect(lambda *args: results.append(args), Qt.DirectConnection)
    worker.failed.connect(lambda *args: failures.append(args), Qt.DirectConnection)
    worker.cleanup_failed.connect(notices.append, Qt.DirectConnection)
    return worker, created, results, failures, notices


@pytest.mark.parametrize("kind", ["scan", "preview"])
@pytest.mark.parametrize("mode", ["success", "exception", "warning", "both"])
def test_worker_preserves_successful_result_and_reports_cleanup(qapp, tmp_path, kind, mode):
    out_path = tmp_path / "synthetic-preview.png"
    worker, created, results, failures, notices = make_worker(kind, mode, out_path)

    worker.start()
    assert worker.wait(5000), "worker did not finish cleanup"

    assert len(created) == 1 and created[0].closed_count == 1
    assert failures == []
    assert results == ([(17, ["编号", "香型："])] if kind == "scan" else [(str(out_path),)])
    assert_cleanup_notice(mode, notices)


@pytest.mark.parametrize("kind", ["scan", "preview"])
@pytest.mark.parametrize("mode", ["exception", "warning"])
def test_worker_keeps_business_failure_separate_from_cleanup(qapp, tmp_path, kind, mode):
    worker, created, results, failures, notices = make_worker(
        kind, mode, tmp_path / "synthetic-preview.png", operation_fails=True)

    worker.start()
    assert worker.wait(5000), "worker did not finish cleanup"

    assert created[0].closed_count == 1
    assert results == []
    assert failures == ([(17, "synthetic scan failure")] if kind == "scan"
                        else [("synthetic preview failure",)])
    assert_cleanup_notice(mode, notices)
