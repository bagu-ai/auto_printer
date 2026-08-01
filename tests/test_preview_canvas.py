import pytest
from PySide6.QtCore import QPointF, Qt
from PySide6.QtGui import QColor, QPixmap
from PySide6.QtWidgets import QApplication

from label_printer.ui.widgets import PreviewCanvas


@pytest.fixture(scope="session")
def qapp():
    return QApplication.instance() or QApplication([])


def pixmap(w=600, h=300):
    p = QPixmap(w, h)
    p.fill(QColor("white"))
    return p


class FakeMouseEvent:
    """鸭子类型事件：PreviewCanvas 的鼠标处理只用 position()/button()/accept()。"""

    def __init__(self, pos, button=Qt.LeftButton):
        self._pos = QPointF(*pos)
        self._button = button

    def position(self):
        return self._pos

    def button(self):
        return self._button

    def accept(self):
        pass


def test_fit_never_upscales_source(qapp):
    canvas = PreviewCanvas()
    canvas.resize(1200, 800)
    canvas.set_pixmap(pixmap(600, 300))
    canvas.fit_to_window()
    assert canvas.zoom == 1.0


def test_fit_downscales_large_source(qapp):
    canvas = PreviewCanvas()
    canvas.resize(400, 250)
    canvas.set_pixmap(pixmap(1200, 600))
    canvas.fit_to_window()
    assert 0.25 <= canvas.zoom < 1.0


def test_actual_and_zoom_limits(qapp):
    canvas = PreviewCanvas()
    canvas.resize(500, 300)
    canvas.set_pixmap(pixmap())
    canvas.actual_size()
    assert canvas.zoom == 1.0
    for _ in range(20):
        canvas.zoom_in()
    assert canvas.zoom == canvas.MAX_ZOOM
    for _ in range(40):
        canvas.zoom_out()
    assert canvas.zoom == canvas.MIN_ZOOM


def test_zoom_out_never_enlarges_when_fit_below_min(qapp):
    """超大标签 fit < MIN_ZOOM 时，点「−」不得被钳到 MIN_ZOOM 反而放大。"""
    canvas = PreviewCanvas()
    canvas.resize(400, 250)
    canvas.set_pixmap(pixmap(2400, 1200))  # fit ≈ 0.147 < 0.25
    canvas.fit_to_window()
    fit_zoom = canvas.zoom
    assert fit_zoom < canvas.MIN_ZOOM
    canvas.zoom_out()
    assert canvas.zoom <= fit_zoom


def test_resize_recomputes_from_source(qapp):
    canvas = PreviewCanvas()
    canvas.resize(400, 250)
    canvas.show()
    qapp.processEvents()
    canvas.set_pixmap(pixmap(1200, 600))
    first = canvas.zoom
    canvas.resize(800, 500)
    qapp.processEvents()
    assert canvas.source_pixmap.width() == 1200
    assert canvas.zoom > first


def test_display_pixmap_scales_in_device_pixels(qapp):
    """高 DPI：显示副本按 zoom×dpr 生成并带 DPR 标记，100% 才是真 1:1。"""
    canvas = PreviewCanvas()
    canvas.resize(800, 500)
    canvas.set_pixmap(pixmap(600, 300))
    canvas.actual_size()
    canvas.devicePixelRatioF = lambda: 2.0  # 离屏默认 1.0，实例属性遮蔽模拟 200%
    scaled = canvas._display_pixmap()
    assert scaled.size().width() == 1200
    assert scaled.size().height() == 600
    assert scaled.devicePixelRatio() == 2.0


def test_drag_pans_zoomed_image(qapp):
    canvas = PreviewCanvas()
    canvas.resize(400, 250)
    canvas.set_pixmap(pixmap(600, 300))
    canvas._mode = "manual"
    canvas._set_zoom(2.0)  # 图像 1200×600，超出画布
    canvas.mousePressEvent(FakeMouseEvent((10, 10)))
    canvas.mouseMoveEvent(FakeMouseEvent((110, 60)))
    assert canvas._pan.x() == 100
    assert canvas._pan.y() == 50
    canvas.mouseReleaseEvent(FakeMouseEvent((110, 60)))
    assert not canvas._dragging


def test_pan_clamped_and_reset_on_fit(qapp):
    canvas = PreviewCanvas()
    canvas.resize(400, 250)
    canvas.set_pixmap(pixmap(600, 300))
    canvas._mode = "manual"
    canvas._set_zoom(2.0)
    canvas._pan = QPointF(99999, -99999)
    canvas._clamp_pan()
    assert abs(canvas._pan.x()) <= 1200
    canvas.fit_to_window()
    assert canvas._pan.x() == 0 and canvas._pan.y() == 0

