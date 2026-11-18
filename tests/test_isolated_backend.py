"""Real spawn/IPC tests, using only the importable fake diagnostic backend."""
import json
import threading
import time
from functools import partial

import pytest

from label_printer.core.backend import BackendCancelled, BackendError, PrintOutcomeUnknown
from label_printer.core.diagnostics import DiagnosticBackend
from label_printer.core.isolated_backend import IsolatedBarTenderBackend, wait_for_completion_from_env


class ModeDiagnosticBackend(DiagnosticBackend):
    """可被 spawn 导入的后端，记录实际收到的配置而不启动 COM。"""
    def __init__(self, *, wait_for_completion=False, optimize_repeated_values=False,
                 print_timeout_ms=None, **kwargs):
        super().__init__(**kwargs)
        self.wait_for_completion = wait_for_completion
        self.optimize_repeated_values = optimize_repeated_values

    def open(self, template, printer=""):
        super().open(template, printer)
        self._record("configuration", wait_for_completion=self.wait_for_completion,
                     optimize_repeated_values=self.optimize_repeated_values)


def proxy(tmp_path, **faults):
    return IsolatedBarTenderBackend(
        backend_factory=DiagnosticBackend,
        backend_kwargs={"evidence_path": str(tmp_path / "calls.jsonl"), **faults},
        startup_timeout_s=10, operation_timeout_s=3, print_timeout_s=0.25, close_timeout_s=0.2)


def calls(tmp_path):
    return [json.loads(line) for line in (tmp_path / "calls.jsonl").read_text().splitlines()]


@pytest.mark.parametrize("value,expected", [(None, False), ("1", True), ("0", False)])
def test_completion_mode_env_is_strict_and_defaults_to_submission(monkeypatch, value, expected):
    monkeypatch.delenv("LABEL_PRINTER_PRINT_TIMEOUT_MS", raising=False)
    monkeypatch.delenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", raising=False)
    if value is not None:
        monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", value)
    assert wait_for_completion_from_env() is expected
    backend = IsolatedBarTenderBackend()
    assert backend.wait_for_completion is expected
    assert backend._kwargs["wait_for_completion"] is expected
    assert backend._print_timeout == 75.0


@pytest.mark.parametrize("value", ["", "false", "true", " 0", "0 ", "2", "-1"])
def test_invalid_completion_mode_env_fails_before_start(monkeypatch, value):
    monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", value)
    with pytest.raises(BackendError, match="LABEL_PRINTER_WAIT_FOR_COMPLETION"):
        IsolatedBarTenderBackend()


@pytest.mark.parametrize("value", [None, "0"])
def test_default_and_explicit_submission_mode_is_frozen_and_disables_fast_mode(monkeypatch, value):
    monkeypatch.delenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", raising=False)
    if value is not None:
        monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", value)
    monkeypatch.setenv("LABEL_PRINTER_FAST_MODE", "1")
    backend = IsolatedBarTenderBackend()
    monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", "1")
    assert backend.wait_for_completion is False
    assert backend._kwargs["wait_for_completion"] is False
    assert backend._kwargs["optimize_repeated_values"] is False
    assert backend.child_pid is None


@pytest.mark.parametrize("mode", [True, False])
def test_explicit_mode_overrides_env_and_remains_frozen(monkeypatch, mode):
    monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", "invalid-but-overridden")
    backend = IsolatedBarTenderBackend(wait_for_completion=mode)
    assert backend.wait_for_completion is mode
    assert backend._kwargs["wait_for_completion"] is mode


@pytest.mark.parametrize("mode", [0, 1, "0", "1"])
def test_explicit_mode_rejects_non_bool_before_start(mode):
    with pytest.raises(BackendError, match="wait_for_completion"):
        IsolatedBarTenderBackend(wait_for_completion=mode)
    with pytest.raises(BackendError, match="wait_for_completion"):
        IsolatedBarTenderBackend(backend_kwargs={"wait_for_completion": mode})


def test_explicit_backend_kwargs_remain_authoritative_and_are_copied(monkeypatch):
    monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", "0")
    kwargs = {"wait_for_completion": True, "optimize_repeated_values": True}
    backend = IsolatedBarTenderBackend(backend_kwargs=kwargs, wait_for_completion=False)
    kwargs["wait_for_completion"] = False
    assert backend.wait_for_completion is True
    assert backend._kwargs == {"wait_for_completion": True, "optimize_repeated_values": True}


def test_explicit_kwargs_without_mode_follow_backend_default_without_adding_option(monkeypatch):
    monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", "1")
    backend = IsolatedBarTenderBackend(
        backend_kwargs={"optimize_repeated_values": True}, wait_for_completion=True)
    assert backend.wait_for_completion is False
    assert backend._kwargs == {"optimize_repeated_values": False}


def test_explicit_diagnostic_kwargs_do_not_receive_unknown_options(monkeypatch, tmp_path):
    monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", "0")
    backend = proxy(tmp_path)
    assert backend.wait_for_completion is False
    assert "wait_for_completion" not in backend._kwargs
    assert "optimize_repeated_values" not in backend._kwargs
    try:
        backend.open("fake.btw")
        backend.print_label({"id": "no-extra-kwargs"})
    finally:
        backend.close()
    assert sum(record["operation"] == "print" for record in calls(tmp_path)) == 1


@pytest.mark.parametrize("mode", [True, False])
def test_completion_mode_reaches_spawned_child_and_disables_incompatible_cache(tmp_path, mode):
    backend = IsolatedBarTenderBackend(
        backend_factory=ModeDiagnosticBackend,
        backend_kwargs={"evidence_path": str(tmp_path / "calls.jsonl"),
                        "wait_for_completion": mode, "optimize_repeated_values": True},
        startup_timeout_s=10, operation_timeout_s=3, close_timeout_s=1)
    try:
        backend.open("fake.btw")
        backend.print_label({"id": "mode-check"}, copies=2)
    finally:
        backend.close()
    records = calls(tmp_path)
    configuration = next(record for record in records if record["operation"] == "configuration")
    assert configuration["wait_for_completion"] is mode
    assert configuration["optimize_repeated_values"] is mode
    assert sum(record["operation"] == "print" for record in records) == 1


@pytest.mark.parametrize("configured_mode,mode", [(None, False), (True, True), (False, False)])
def test_default_and_explicit_env_mode_is_frozen_before_child_spawn(monkeypatch, tmp_path, configured_mode, mode):
    monkeypatch.delenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", raising=False)
    if configured_mode is not None:
        monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", "1" if configured_mode else "0")
    monkeypatch.setenv("LABEL_PRINTER_FAST_MODE", "1")
    factory = partial(ModeDiagnosticBackend, evidence_path=str(tmp_path / "calls.jsonl"))
    backend = IsolatedBarTenderBackend(backend_factory=factory, startup_timeout_s=10)
    monkeypatch.setenv("LABEL_PRINTER_WAIT_FOR_COMPLETION", "0" if mode else "1")
    try:
        backend.open("fake.btw")
    finally:
        backend.close()
    configuration = next(record for record in calls(tmp_path) if record["operation"] == "configuration")
    assert configuration["wait_for_completion"] is mode
    assert configuration["optimize_repeated_values"] is mode


def test_compatibility_mode_still_has_bounded_external_wait_without_retry(tmp_path):
    backend = IsolatedBarTenderBackend(
        backend_factory=ModeDiagnosticBackend,
        backend_kwargs={"evidence_path": str(tmp_path / "calls.jsonl"),
                        "wait_for_completion": False, "hang_at": 1},
        startup_timeout_s=10, print_timeout_s=0.25, close_timeout_s=0.2)
    try:
        backend.open("fake.btw")
        started = time.monotonic()
        with pytest.raises(PrintOutcomeUnknown):
            backend.print_label({"id": "unknown"})
        assert time.monotonic() - started < 4
        assert backend.child_pid is None
        with pytest.raises(BackendError, match="会话已结束"):
            backend.print_label({"id": "must-not-retry"})
    finally:
        backend.close()
    assert sum(record["operation"] == "print" for record in calls(tmp_path)) == 1


def test_process_reused_without_reopening_template(tmp_path):
    backend = proxy(tmp_path)
    try:
        backend.open("fake.btw")
        pid = backend.child_pid
        for n in range(20):
            backend.set_job_name(f"RUN-r{n:04}")
            backend.print_label({"id": str(n)})
            assert backend.child_pid == pid
    finally:
        backend.close()
    records = calls(tmp_path)
    assert sum(r["operation"] == "open" for r in records) == 1
    assert len([r for r in records if r["operation"] == "print"]) == 20
    assert records[-2]["job_name"] == "RUN-r0019"


@pytest.mark.parametrize("fault", [{"hang_at": 9}, {"crash_at": 9}])
def test_stall_or_crash_is_unknown_no_retry(tmp_path, fault):
    backend = proxy(tmp_path, **fault)
    try:
        backend.open("fake.btw")
        for n in range(8):
            backend.print_label({"id": str(n)})
        start = time.monotonic()
        with pytest.raises(PrintOutcomeUnknown):
            backend.print_label({"id": "ninth"})
        assert time.monotonic() - start < 4
        assert backend.child_pid is None
        with pytest.raises(BackendError, match="会话已结束"):
            backend.print_label({"id": "must-not-retry"})
    finally:
        backend.close()
    assert sum(r["operation"] == "print" for r in calls(tmp_path)) == 9


def test_cancel_interrupts_wait_but_does_not_claim_unprinted(tmp_path):
    backend = proxy(tmp_path, hang_at=1)
    backend._print_timeout = 10
    backend.open("fake.btw")
    timer = threading.Timer(0.1, backend.cancel)
    timer.start()
    try:
        with pytest.raises(PrintOutcomeUnknown, match="可能已进入打印队列"):
            backend.print_label({"id": "cancel"})
        assert backend.child_pid is None
    finally:
        timer.join()
        backend.close()


def test_cancel_before_start_does_not_create_child(tmp_path):
    backend = proxy(tmp_path)
    backend.cancel()
    with pytest.raises(BackendCancelled):
        backend.open("fake.btw")
    assert backend.child_pid is None
    assert not (tmp_path / "calls.jsonl").exists()


def test_journal_failure_before_ack_prevents_print(tmp_path):
    backend = proxy(tmp_path)
    backend.open("fake.btw")
    def observer(event):
        raise OSError("disk full")
    backend.set_observer(observer)
    try:
        with pytest.raises(PrintOutcomeUnknown, match="记录调用阶段失败"):
            backend.print_label({"id": "must-not-send"})
    finally:
        backend.close()
    assert not any(r["operation"] == "print" for r in calls(tmp_path))


def test_close_is_bounded_and_reports_cleanup_warning(tmp_path):
    backend = proxy(tmp_path, hang_close=True)
    backend.open("fake.btw")
    started = time.monotonic()
    backend.close()
    assert time.monotonic() - started < 4
    assert backend.child_pid is None
    assert "清理未正常完成" in backend.cleanup_warning
