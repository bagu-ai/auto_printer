"""Safe fault-injection diagnostics usable from the frozen executable too.

No real COM application, printer or spooler request is created by this module.
"""
import json
import os
import threading
import time
from pathlib import Path

from .backend import BackendError, PrintOutcomeUnknown
from .isolated_backend import IsolatedBarTenderBackend


class DiagnosticBackend:
    def __init__(self, *, evidence_path, hang_at=0, crash_at=0, hang_close=False):
        self.path = Path(evidence_path)
        self.hang_at, self.crash_at, self.hang_close = hang_at, crash_at, hang_close
        self.count = 0
        self.observer = lambda event: None
        self.job_name = ""

    def set_observer(self, callback):
        self.observer = callback

    def set_job_name(self, name):
        self.job_name = name

    def _record(self, operation, **data):
        with self.path.open("a", encoding="utf-8") as f:
            f.write(json.dumps({"operation": operation, **data}) + "\n")
            f.flush()

    def open(self, template, printer=""):
        self._record("open")

    def print_label(self, values, copies=1):
        self.count += 1
        self.observer({"phase": "before", "operation": "print", "job_name": self.job_name})
        self._record("print", number=self.count, values=values, copies=copies, job_name=self.job_name)
        if self.count == self.crash_at:
            os._exit(41)
        if self.count == self.hang_at:
            while True:
                time.sleep(60)
        self.observer({"phase": "after", "operation": "print", "result": 0})

    def close(self):
        if self.hang_close:
            while True:
                time.sleep(60)
        self._record("close")

    def list_template_fields(self, template):
        return ["id"]

    def render_preview(self, values, out_path):
        raise BackendError("自检后端不生成校样")


def _new(path, **kwargs):
    return IsolatedBarTenderBackend(
        backend_factory=DiagnosticBackend, backend_kwargs={"evidence_path": str(path), **kwargs},
        startup_timeout_s=20, operation_timeout_s=5, print_timeout_s=0.6, close_timeout_s=0.4)


def run_self_test(output_dir):
    output = Path(output_dir).resolve()
    output.mkdir(parents=True, exist_ok=True)
    run = output / f"probe-{time.time_ns()}"
    run.mkdir()
    results = {"uses_real_printer": False, "checks": [], "run": str(run)}
    try:
        normal = _new(run / "normal.jsonl")
        normal.open("fake.btw", "FAKE_ONLY")
        pid = normal.child_pid
        for n in range(100):
            normal.set_job_name(f"SIM-{n:03}")
            normal.print_label({"id": str(n)}, 1)
            assert normal.child_pid == pid
        normal.close()
        normal_rows = [json.loads(s) for s in (run / "normal.jsonl").read_text().splitlines()]
        assert sum(x["operation"] == "print" for x in normal_rows) == 100
        assert sum(x["operation"] == "open" for x in normal_rows) == 1
        results["checks"].append({"name": "100_rows_reuse_one_process_and_template", "pass": True})

        stalled = _new(run / "stalled.jsonl", hang_at=9)
        stalled.open("fake.btw")
        started = time.monotonic()
        try:
            for n in range(12):
                stalled.print_label({"id": str(n)})
            raise AssertionError("A blocked print unexpectedly succeeded")
        except PrintOutcomeUnknown:
            elapsed = time.monotonic() - started
            assert elapsed < 5
            assert stalled.child_pid is None
        finally:
            stalled.close()
        stalled_rows = [json.loads(s) for s in (run / "stalled.jsonl").read_text().splitlines()]
        assert sum(x["operation"] == "print" for x in stalled_rows) == 9
        results["checks"].append({"name": "ninth_call_stall_is_bounded_unknown_no_retries", "pass": True,
                                   "elapsed_s": round(elapsed, 3)})

        cancel = _new(run / "cancel.jsonl", hang_at=1)
        cancel._print_timeout = 10
        cancel.open("fake.btw")
        timer = threading.Timer(0.15, cancel.cancel)
        timer.start()
        started = time.monotonic()
        try:
            cancel.print_label({"id": "cancel"})
            raise AssertionError("Cancelled in-flight print unexpectedly succeeded")
        except PrintOutcomeUnknown:
            assert time.monotonic() - started < 5
        finally:
            timer.join()
            cancel.close()
        results["checks"].append({"name": "cancel_breaks_inflight_wait", "pass": True})

        crashing = _new(run / "crash.jsonl", crash_at=1)
        crashing.open("fake.btw")
        try:
            crashing.print_label({"id": "crash"})
            raise AssertionError("Crashing child unexpectedly succeeded")
        except PrintOutcomeUnknown:
            assert crashing.child_pid is None
        finally:
            crashing.close()
        results["checks"].append({"name": "child_crash_is_unknown", "pass": True})

        blocked_log = _new(run / "log_failure.jsonl")
        blocked_log.open("fake.btw")
        def fail_log(event):
            if event.get("operation") == "print":
                raise OSError("injected disk error")
        blocked_log.set_observer(fail_log)
        try:
            blocked_log.print_label({"id": "must-not-send"})
            raise AssertionError("Print should be blocked by logging failure")
        except PrintOutcomeUnknown:
            pass
        finally:
            blocked_log.close()
        assert '"operation": "print"' not in (run / "log_failure.jsonl").read_text()
        results["checks"].append({"name": "failed_intent_log_prevents_send", "pass": True})

        bad_close = _new(run / "close.jsonl", hang_close=True)
        bad_close.open("fake.btw")
        started = time.monotonic()
        bad_close.close()
        assert time.monotonic() - started < 5 and bad_close.child_pid is None
        assert bad_close.cleanup_warning
        results["checks"].append({"name": "cleanup_hang_is_bounded", "pass": True})
        results["success"] = True
    except Exception as exc:
        results["success"] = False
        results["error"] = repr(exc)
    (output / "self-test-result.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if results["success"] else 1
