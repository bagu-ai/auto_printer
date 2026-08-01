"""Worker 线程模型测试：后端必须在 worker 线程内创建、使用、销毁。"""
import threading

import pytest
from PySide6.QtWidgets import QApplication

from label_printer.core.mock_backend import MockBackend
from label_printer.ui.main_window import PrintWorker, PreviewWorker


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app


class RecordingFactory:
    """记录 backend 在哪个线程被创建；每个实例记录自己被 close 的次数。"""

    def __init__(self, backend_class=MockBackend):
        self.created_in_threads: list[int] = []
        self.instances: list = []
        self._backend_class = backend_class

    def __call__(self):
        self.created_in_threads.append(threading.get_ident())
        backend = self._backend_class()
        self.instances.append(backend)
        return backend


def run_worker(qapp, worker, timeout=10000):
    worker.start()
    assert worker.wait(timeout), "worker 超时未结束"


def test_backend_created_in_worker_thread(qapp):
    factory = RecordingFactory()
    jobs = [(1, {"发货新编号": "TB-1"}, 3), (2, {"发货新编号": "TB-2"}, 1)]
    worker = PrintWorker(factory, r"C:\模板\t.btw", "P", jobs)
    main_tid = threading.get_ident()
    run_worker(qapp, worker)

    assert factory.created_in_threads, "工厂没被调用"
    assert all(t != main_tid for t in factory.created_in_threads), \
        "backend 必须在 worker 线程内创建（COM 线程亲和）"
    assert len(factory.instances) == 1
    backend = factory.instances[0]
    assert len(backend.printed) == 2
    assert backend.closed_count >= 1, "worker 结束后必须 close 后端"
    # 打印行数与状态记录一致
    assert len(worker.statuses) == 2
    assert all(s["status"] == "submitted" for s in worker.statuses)


def test_cancel_stops_loop(qapp):
    class SlowMock(MockBackend):
        def print_label(self, values, copies=1):
            super().print_label(values, copies)
            if len(self.printed) == 2:
                worker_ref.cancel()

    factory = RecordingFactory(SlowMock)
    jobs = [(i, {"f": f"v{i}"}, 1) for i in range(1, 11)]
    worker_ref = PrintWorker(factory, r"C:\模板\t.btw", "P", jobs)
    run_worker(qapp, worker_ref)

    backend = factory.instances[0]
    assert len(backend.printed) < 10, "取消后不应打完全部行"
    assert backend.closed_count >= 1


def test_backend_closed_on_exception(qapp):
    class FailMock(MockBackend):
        def print_label(self, values, copies=1):
            raise RuntimeError("模拟打印失败")

    factory = RecordingFactory(FailMock)
    jobs = [(1, {"f": "v"}, 1)]
    worker = PrintWorker(factory, r"C:\模板\t.btw", "P", jobs)
    run_worker(qapp, worker)

    backend = factory.instances[0]
    assert backend.closed_count >= 1, "异常路径也必须 close 后端（防 COM 泄漏）"
    assert worker.statuses[0]["status"] == "failed"
    assert "模拟打印失败" in worker.statuses[0]["error"]


def test_preview_worker_renders_png(qapp, tmp_path):
    factory = RecordingFactory()
    out = tmp_path / "preview.png"
    worker = PreviewWorker(factory, r"C:\模板\t.btw", "P",
                           {"发货新编号": "TB-1", "客户香型": "Mint"}, str(out))
    run_worker(qapp, worker)

    assert out.exists() and out.stat().st_size > 100
    # PNG 魔数
    assert out.read_bytes()[:4] == b"\x89PNG"
    assert factory.instances[0].closed_count >= 1
