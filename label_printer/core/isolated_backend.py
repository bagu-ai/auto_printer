"""Bounded, cancellable IPC around a single-threaded BarTender COM session.

Only this app's Python child is terminated on a stall. No global BarTender or
spooler process is killed, and no queue is cleared. A lost response to a print
request always means UNKNOWN, not safe-to-retry. Each failed channel is retired.
"""
from __future__ import annotations

import multiprocessing as mp
import os
import threading
import time
from typing import Callable

from .backend import BackendCancelled, BackendError, PrintBackend, PrintOutcomeUnknown


def _bartender_factory(**kwargs):
    from .bartender_backend import BarTenderBackend
    return BarTenderBackend(**kwargs)


def _serve(connection, factory, kwargs):
    """Importable spawn target: creates, uses and disposes COM in this thread."""
    backend = None
    pythoncom = None
    try:
        if os.name == "nt":
            import pythoncom
            pythoncom.CoInitialize()
        backend = factory(**kwargs)

        def event_observer(event):
            connection.send({"kind": "event", "event": dict(event)})
            # The controller persists the print intent before acknowledging.
            # If logging fails or it exits, do not continue into the COM call.
            ack = connection.recv()
            if ack != {"kind": "ack"}:
                raise BackendError("阶段记录未获确认，已停止提交")

        observe = getattr(backend, "set_observer", None)
        if observe:
            observe(event_observer)
        connection.send({"kind": "ready"})
        while True:
            command = connection.recv()
            operation = command["operation"]
            if operation not in {"open", "list_template_fields", "render_preview", "print_label", "close"}:
                raise BackendError("未知后台操作")
            try:
                if operation == "print_label":
                    naming = getattr(backend, "set_job_name", None)
                    if naming:
                        naming(command.get("job_name", ""))
                result = getattr(backend, operation)(*command.get("args", ()))
                connection.send({"kind": "result", "value": result})
            except Exception as exc:
                category = ("unknown" if isinstance(exc, PrintOutcomeUnknown)
                            else "cancelled" if isinstance(exc, BackendCancelled) else "error")
                connection.send({"kind": "error", "category": category, "message": str(exc)})
            if operation == "close":
                backend = None  # close was attempted; avoid an unbounded second close.
                break
    except (EOFError, BrokenPipeError, OSError):
        pass
    except Exception as exc:
        try:
            connection.send({"kind": "error", "category": "error", "message": str(exc)})
        except (EOFError, BrokenPipeError, OSError):
            pass
    finally:
        # Parent also bounds shutdown, including COM Close/Quit hanging here.
        if backend is not None:
            try:
                backend.close()
            except Exception:
                pass
        if pythoncom is not None:
            pythoncom.CoUninitialize()
        connection.close()


def _env_timeout():
    value = os.environ.get("LABEL_PRINTER_PRINT_TIMEOUT_MS", "60000")
    try:
        result = int(value)
    except ValueError as exc:
        raise BackendError("打印超时设置必须为毫秒整数") from exc
    if not 1000 <= result <= 600000:
        raise BackendError("打印超时设置须在 1000～600000 毫秒之间")
    return result


def wait_for_completion_from_env() -> bool:
    """默认采用已验证的提交模式；显式 1 保留用于研发回归，批次开始时冻结。"""
    value = os.environ.get("LABEL_PRINTER_WAIT_FOR_COMPLETION", "0")
    if value not in {"0", "1"}:
        raise BackendError("LABEL_PRINTER_WAIT_FOR_COMPLETION 必须是 0 或 1；默认 0 为提交模式")
    return value == "1"


class IsolatedBarTenderBackend(PrintBackend):
    name = "BarTender（防卡住版）"

    def __init__(self, *, backend_factory=None, backend_kwargs=None,
                 startup_timeout_s=60.0, operation_timeout_s=60.0,
                 print_timeout_s=None, close_timeout_s=4.0,
                 wait_for_completion=None):
        api_timeout = _env_timeout()
        if wait_for_completion is not None and not isinstance(wait_for_completion, bool):
            raise BackendError("wait_for_completion 必须是 bool 或 None")
        self._factory = backend_factory or _bartender_factory
        if backend_kwargs is None:
            self._kwargs = {
                "print_timeout_ms": api_timeout,
                "optimize_repeated_values": os.environ.get("LABEL_PRINTER_FAST_MODE") == "1",
                "wait_for_completion": (wait_for_completion_from_env()
                                        if wait_for_completion is None else wait_for_completion),
            }
        else:
            # 显式参数完整优先；不向诊断等自定义工厂强塞不支持的参数。
            self._kwargs = dict(backend_kwargs)
        self._wait_for_completion = self._kwargs.get("wait_for_completion", False)
        if not isinstance(self._wait_for_completion, bool):
            raise BackendError("wait_for_completion 必须是 bool")
        if not self._wait_for_completion and "optimize_repeated_values" in self._kwargs:
            self._kwargs["optimize_repeated_values"] = False
        self._startup_timeout = float(startup_timeout_s)
        self._operation_timeout = float(operation_timeout_s)
        self._print_timeout = float(print_timeout_s if print_timeout_s is not None else api_timeout / 1000 + 15)
        self._close_timeout = float(close_timeout_s)
        self._process = None
        self._connection = None
        self._cancel = threading.Event()
        self._observer: Callable | None = None
        self._job_name = ""
        self._retired = False
        self.cleanup_warning = ""

    def set_observer(self, callback):
        self._observer = callback

    @property
    def wait_for_completion(self) -> bool:
        return self._wait_for_completion

    def set_job_name(self, name):
        self._job_name = str(name)

    def cancel(self):
        """Thread-safe request; the owner thread tears down IPC and its child."""
        self._cancel.set()

    @property
    def child_pid(self):
        return self._process.pid if self._process else None

    def _abort(self):
        self._retired = True
        connection, process = self._connection, self._process
        self._connection = None
        self._process = None
        if connection is not None:
            connection.close()
        if process is not None:
            if process.is_alive():
                process.terminate()
                process.join(1.5)
            if process.is_alive():
                process.kill()
                process.join(1.0)
            if not process.is_alive():
                process.close()

    def _lost(self, operation, reason):
        self._abort()
        if operation == "print_label":
            raise PrintOutcomeUnknown(
                f"{reason}。当前行可能已进入打印队列，结果未知；已停止后续行。"
                "请核对队列和实际出纸编号后再决定补打；旧任务可能仍在 BarTender 或打印机中。")
        raise BackendError(f"{reason}，已结束本次后台会话")

    def _wait(self, operation, timeout, *, ignore_cancel=False):
        deadline = time.monotonic() + timeout
        last_heartbeat = time.monotonic()
        while True:
            if not ignore_cancel and self._cancel.is_set():
                if operation == "print_label":
                    self._lost(operation, "已取消当前等待")
                self._abort()
                raise BackendCancelled("已取消，未继续提交后续任务")
            if time.monotonic() >= deadline:
                self._lost(operation, f"{operation} 超过 {timeout:g} 秒未返回")
            try:
                available = self._connection.poll(min(0.05, max(0.0, deadline - time.monotonic())))
                if available:
                    message = self._connection.recv()
                    kind = message.get("kind")
                    if kind == "event":
                        try:
                            if self._observer:
                                self._observer(message["event"])
                        except Exception as exc:
                            self._lost(operation, f"记录调用阶段失败：{exc}")
                        self._connection.send({"kind": "ack"})
                        continue
                    if kind in ("ready", "result"):
                        return message.get("value")
                    if kind == "error":
                        category = message.get("category")
                        error_type = {"unknown": PrintOutcomeUnknown, "cancelled": BackendCancelled}.get(category, BackendError)
                        raise error_type(message.get("message", "后台操作失败"))
                    self._lost(operation, "后台返回了无法识别的消息")
                if not self._process.is_alive():
                    self._lost(operation, "后台进程意外退出")
            except (EOFError, BrokenPipeError, OSError) as exc:
                self._lost(operation, f"后台连接中断：{exc}")
            if self._observer and time.monotonic() - last_heartbeat >= 2:
                last_heartbeat = time.monotonic()
                try:
                    self._observer({"phase": "waiting", "operation": operation,
                                    "remaining_s": round(max(0, deadline - last_heartbeat), 1)})
                except Exception as exc:
                    self._lost(operation, f"记录等待状态失败：{exc}")

    def _start(self):
        if self._retired:
            raise BackendError("后台会话已结束，须核对后创建新批次")
        if self._cancel.is_set():
            raise BackendCancelled("已取消，尚未启动打印会话")
        if self._process:
            return
        context = mp.get_context("spawn")
        parent, child = context.Pipe(duplex=True)
        self._connection = parent
        self._process = context.Process(target=_serve, args=(child, self._factory, self._kwargs), daemon=True)
        try:
            self._process.start()
            child.close()
            self._wait("startup", self._startup_timeout)
        except Exception:
            child.close()
            self._abort()
            raise

    def _call(self, operation, *args, timeout=None):
        self._start()
        if self._cancel.is_set():
            self._abort()
            raise BackendCancelled("已取消，未发送当前行")
        try:
            self._connection.send({"operation": operation, "args": args, "job_name": self._job_name})
        except (EOFError, BrokenPipeError, OSError) as exc:
            self._lost(operation, f"发送后台请求失败：{exc}")
        return self._wait(operation, timeout or self._operation_timeout)

    def open(self, template_path, printer=""):
        return self._call("open", template_path, printer)

    def list_template_fields(self, template_path):
        return self._call("list_template_fields", template_path)

    def list_printers(self):
        try:
            import win32print
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            return [p[2] for p in win32print.EnumPrinters(flags)]
        except Exception:
            return []

    def print_label(self, values, copies=1):
        return self._call("print_label", dict(values), copies, timeout=self._print_timeout)

    def render_preview(self, values, out_path):
        return self._call("render_preview", dict(values), out_path)

    def close(self):
        if self._process is None:
            return
        try:
            self._connection.send({"operation": "close", "args": ()})
            self._wait("close", self._close_timeout, ignore_cancel=True)
        except Exception as exc:
            # Previously submitted rows do not become unprinted on cleanup failure.
            self.cleanup_warning = f"后台清理未正常完成：{exc}；可能残留本次 BarTender 会话"
        finally:
            self._abort()
