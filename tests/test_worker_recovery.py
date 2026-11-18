import json
import threading

import pytest
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from label_printer.core.backend import BackendError, PrintOutcomeUnknown
from label_printer.core.job_journal import JobJournal, JournalError
from label_printer.core.mock_backend import MockBackend
from label_printer.ui.main_window import PrintWorker


@pytest.fixture(scope="session")
def qapp():
    return QApplication.instance() or QApplication([])


def events(worker):
    return [json.loads(line) for line in worker.journal_path.read_text(encoding="utf-8").splitlines()]


def jobs(count=12):
    return [(i, {"id": f"A{i}"}, 1) for i in range(1, count + 1)]


def run(worker):
    worker.start()
    assert worker.wait(5000)
    return worker.result


def test_unknown_stops_remaining_rows_with_durable_intent(qapp, tmp_path):
    class Backend(MockBackend):
        def set_job_name(self, name):
            self.job_name = name
        def print_label(self, values, copies=1):
            audit = events(worker)
            assert audit[-1]["event"] == "row_call_begin"
            assert audit[-2]["event"] == "row_start"
            assert len(self.job_name) <= 31
            if len(self.printed) == 8:
                raise PrintOutcomeUnknown("simulated timeout")
            super().print_label(values, copies)
    worker = PrintWorker(Backend, "a.btw", "p", jobs(), journal_dir=tmp_path)
    success, message, statuses = run(worker)
    assert not success and "不要直接重打未知行" in message
    assert len(worker.backend.printed) == 8
    assert [entry["status"] for entry in statuses] == ["submitted"] * 8 + ["unknown"] + ["not_attempted"] * 3
    assert "row_unknown" in [event["event"] for event in events(worker)]
    assert events(worker)[-1]["event"] == "batch_end"
    assert worker.backend.closed_count >= 1


def test_failure_stops_instead_of_continuing(qapp, tmp_path):
    class Backend(MockBackend):
        def print_label(self, values, copies=1):
            raise BackendError("broken field")
    worker = PrintWorker(Backend, "a.btw", "p", jobs(3), journal_dir=tmp_path)
    success, _, statuses = run(worker)
    assert not success
    assert [entry["status"] for entry in statuses] == ["failed", "not_attempted", "not_attempted"]


def test_cancel_reaches_active_backend_and_done_follows_cleanup(qapp, tmp_path):
    entered, released = threading.Event(), threading.Event()
    callback_closed_states = []
    class Backend(MockBackend):
        def print_label(self, values, copies=1):
            entered.set()
            assert released.wait(3)
            raise PrintOutcomeUnknown("cancelled while submitting")
        def cancel(self):
            released.set()
    worker = PrintWorker(Backend, "a.btw", "p", jobs(2), journal_dir=tmp_path)
    worker.done.connect(lambda *_: callback_closed_states.append(worker.backend.closed_count),
                        Qt.DirectConnection)
    worker.start()
    assert entered.wait(2)
    worker.cancel()
    assert worker.wait(3000)
    assert callback_closed_states and callback_closed_states[0] >= 1
    assert worker.statuses[0]["status"] == "unknown"
    assert worker.statuses[1]["status"] == "not_attempted"


def test_log_failure_before_call_prevents_submission(qapp, tmp_path, monkeypatch):
    original = JobJournal.append
    def fail_row_start(self, event, **kwargs):
        if event == "row_start":
            raise JournalError("simulated full disk")
        return original(self, event, **kwargs)
    monkeypatch.setattr(JobJournal, "append", fail_row_start)
    worker = PrintWorker(MockBackend, "a.btw", "p", jobs(2), journal_dir=tmp_path)
    success, message, _ = run(worker)
    assert not success and "full disk" in message
    assert worker.backend.printed == []


def test_start_snapshot_is_independent_of_callers_and_summary_matches(qapp, tmp_path):
    original_jobs = jobs(1)
    context = {"excel": "initial.xlsx", "sheet": "first"}
    worker = PrintWorker(MockBackend, "initial.btw", "P1", original_jobs,
                         context=context, journal_dir=tmp_path)
    original_jobs[0][1]["id"] = "changed"
    context["excel"] = "changed.xlsx"
    worker.snapshot["template"] = "changed.btw"
    success, message, _ = run(worker)
    assert success and "提交完成" in message and "出纸请核对" in message
    summary = json.loads(worker.summary_path.read_text(encoding="utf-8"))
    assert summary["excel"] == "initial.xlsx"
    assert summary["template"] == "initial.btw"
    assert summary["printer"] == "P1"
    assert summary["rows"][0]["values"]["id"] == "A1"


def test_cancel_before_start_does_not_create_backend(qapp, tmp_path):
    def factory():
        raise AssertionError("Backend should not start")
    worker = PrintWorker(factory, "a.btw", "p", jobs(2), journal_dir=tmp_path)
    worker.cancel()
    success, _, statuses = run(worker)
    assert not success
    assert all(entry["status"] == "not_attempted" for entry in statuses)
    assert "cancelled" in [event["event"] for event in events(worker)]


def test_cleanup_warning_preserves_submitted_rows(qapp, tmp_path):
    class Backend(MockBackend):
        cleanup_warning = ""
        def close(self):
            super().close()
            self.cleanup_warning = "simulated close deadline"
    worker = PrintWorker(Backend, "a.btw", "p", jobs(1), journal_dir=tmp_path)
    success, message, statuses = run(worker)
    assert success and statuses[0]["status"] == "submitted"
    assert "后台清理未确认" in message
    assert "cleanup_warning" in [event["event"] for event in events(worker)]


def test_phase_journal_failure_does_not_allow_backend_to_send(qapp, tmp_path, monkeypatch):
    original = JobJournal.append
    def fail_print_phase(self, event, **kwargs):
        if event == "phase" and kwargs["detail"]["operation"] == "print":
            raise JournalError("simulated phase write failure")
        return original(self, event, **kwargs)
    monkeypatch.setattr(JobJournal, "append", fail_print_phase)
    class Backend(MockBackend):
        def set_observer(self, observer):
            self.observer = observer
        def print_label(self, values, copies=1):
            self.observer({"operation": "print", "phase": "before"})
            super().print_label(values, copies)
    worker = PrintWorker(Backend, "a.btw", "p", jobs(2), journal_dir=tmp_path)
    success, _, statuses = run(worker)
    assert not success and worker.backend.printed == []
    assert statuses[0]["status"] == "unknown"
    assert statuses[1]["status"] == "not_attempted"
