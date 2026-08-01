"""标签打印工具的可复用 UI 组件。"""
from __future__ import annotations

from PySide6.QtCore import QPointF, QRect, Qt, Signal
from PySide6.QtGui import QColor, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QFrame, QHBoxLayout, QLabel, QPushButton, QSizePolicy,
    QToolButton, QVBoxLayout, QWidget,
)


class StatusPill(QFrame):
    """带文字的状态徽标；颜色只是辅助，状态仍由文字表达。"""

    def __init__(self, text: str = "就绪", state: str = "neutral", parent=None):
        super().__init__(parent)
        self.setObjectName("statusPill")
        lay = QHBoxLayout(self)
        lay.setContentsMargins(9, 4, 10, 4)
        lay.setSpacing(6)
        self.dot = QLabel("●")
        self.dot.setObjectName("statusDot")
        self.text_label = QLabel(text)
        lay.addWidget(self.dot)
        lay.addWidget(self.text_label)
        self.set_state(text, state)

    def set_state(self, text: str, state: str = "neutral"):
        self.text_label.setText(text)
        self.setProperty("state", state)
        # 圆点颜色由父级属性选择器决定，父子都要刷新样式缓存，
        # 否则运行中状态迁移后圆点保持旧颜色。
        for w in (self, self.dot, self.text_label):
            self.style().unpolish(w)
            self.style().polish(w)


class CollapsiblePanel(QFrame):
    """紧凑可折叠面板，适合运行日志等低频内容。"""

    toggled = Signal(bool)

    def __init__(self, title: str, content: QWidget, expanded: bool = False, parent=None):
        super().__init__(parent)
        self.setObjectName("collapsiblePanel")
        self._content = content
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self.header = QToolButton()
        self.header.setObjectName("collapseHeader")
        self.header.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.header.setArrowType(Qt.DownArrow if expanded else Qt.RightArrow)
        self.header.setText(title)
        self.header.setCheckable(True)
        self.header.setChecked(expanded)
        self.header.clicked.connect(self.set_expanded)
        root.addWidget(self.header)
        root.addWidget(content)
        content.setVisible(expanded)

    def is_expanded(self) -> bool:
        return self.header.isChecked()

    def set_title(self, title: str):
        self.header.setText(title)

    def set_expanded(self, expanded: bool):
        self.header.setChecked(expanded)
        self.header.setArrowType(Qt.DownArrow if expanded else Qt.RightArrow)
        self._content.setVisible(expanded)
        self.toggled.emit(expanded)


class PreviewCanvas(QWidget):
    """校样画布：始终从原始 QPixmap 计算显示，避免重复缩放导致模糊。

    高 DPI 下按设备像素生成缩放副本（100% = 1 图像素 : 1 物理像素），
    放大超过画布时可按住左键拖动平移。
    """

    zoomChanged = Signal(float)
    MIN_ZOOM = 0.25
    MAX_ZOOM = 4.0

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("previewCanvas")
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMinimumSize(280, 190)
        self._source = QPixmap()
        self._mode = "fit"
        self._zoom = 1.0
        self._fit_zoom = 1.0
        self._pan = QPointF(0, 0)
        self._dragging = False
        self._message = "选择数据行后刷新校样"
        # 无键盘交互，不做 Tab 停留（避免焦点停在无响应控件上）
        self.setFocusPolicy(Qt.NoFocus)

    @property
    def source_pixmap(self) -> QPixmap:
        return self._source

    @property
    def zoom(self) -> float:
        return self._zoom

    @property
    def mode(self) -> str:
        return self._mode

    def set_message(self, text: str):
        self._message = text
        if self._source.isNull():
            self.update()

    def set_pixmap(self, pixmap: QPixmap):
        self._source = QPixmap(pixmap)
        self._source.setDevicePixelRatio(1.0)
        self._message = ""
        self._pan = QPointF(0, 0)
        if self._mode == "fit":
            self._recompute_fit()
        self.update()

    def clear_preview(self, message: str = "选择数据行后刷新校样"):
        self._source = QPixmap()
        self._pan = QPointF(0, 0)
        self._message = message
        self.update()

    def fit_to_window(self):
        self._mode = "fit"
        self._pan = QPointF(0, 0)
        self._recompute_fit()
        self.update()

    def actual_size(self):
        self._mode = "manual"
        self._set_zoom(1.0)

    def zoom_in(self):
        self._mode = "manual"
        self._set_zoom(self._zoom * 1.25)

    def zoom_out(self):
        self._mode = "manual"
        self._set_zoom(self._zoom / 1.25)

    def _min_zoom(self) -> float:
        # fit 可能小于 MIN_ZOOM（超大标签），手动缩放下限随之放宽，
        # 否则 fit=0.20 时点「−」反而被钳到 0.25（放大）。
        return min(self.MIN_ZOOM, self._fit_zoom)

    def _set_zoom(self, value: float):
        self._zoom = max(self._min_zoom(), min(self.MAX_ZOOM, value))
        self._clamp_pan()
        self.zoomChanged.emit(self._zoom)
        self.update()

    def _recompute_fit(self):
        if self._source.isNull():
            self._zoom = 1.0
            self._fit_zoom = 1.0
            return
        avail_w = max(1, self.width() - 48)
        avail_h = max(1, self.height() - 48)
        # 默认不把源图放大超过 100%；需要放大由用户明确操作。
        self._fit_zoom = min(1.0, avail_w / self._source.width(),
                             avail_h / self._source.height())
        self._zoom = self._fit_zoom
        self.zoomChanged.emit(self._zoom)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._mode == "fit":
            self._recompute_fit()
        self._clamp_pan()

    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier and not self._source.isNull():
            self._mode = "manual"
            self._set_zoom(self._zoom * (1.15 if event.angleDelta().y() > 0 else 1 / 1.15))
            event.accept()
            return
        super().wheelEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and not self._source.isNull():
            self._dragging = True
            self._drag_start = event.position()
            self._pan_start = QPointF(self._pan)
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._dragging:
            self._pan = self._pan_start + (event.position() - self._drag_start)
            self._clamp_pan()
            self.update()
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if self._dragging and event.button() == Qt.LeftButton:
            self._dragging = False
            self.unsetCursor()
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def mouseDoubleClickEvent(self, event):
        if not self._source.isNull():
            self.actual_size() if self._mode == "fit" else self.fit_to_window()
        super().mouseDoubleClickEvent(event)

    def _image_rect(self) -> QRect:
        """图像在画布中的目标矩形（逻辑像素，含平移）。"""
        w = max(1, round(self._source.width() * self._zoom))
        h = max(1, round(self._source.height() * self._zoom))
        x = (self.width() - w) / 2 + self._pan.x()
        y = (self.height() - h) / 2 + self._pan.y()
        return QRect(round(x), round(y), w, h)

    def _clamp_pan(self):
        """图像小于画布时居中；大于画布时至少保留 48px 可见。"""
        if self._source.isNull():
            self._pan = QPointF(0, 0)
            return
        w = self._source.width() * self._zoom
        h = self._source.height() * self._zoom
        keep = 48
        px, py = self._pan.x(), self._pan.y()
        if w <= self.width():
            px = 0
        else:
            half = (w - self.width()) / 2
            px = max(keep - w + half, min(self.width() - keep + half, px))
        if h <= self.height():
            py = 0
        else:
            half = (h - self.height()) / 2
            py = max(keep - h + half, min(self.height() - keep + half, py))
        self._pan = QPointF(px, py)

    def _display_pixmap(self) -> QPixmap:
        """按当前缩放与 devicePixelRatio 生成显示用副本（设备像素空间）。"""
        dpr = self.devicePixelRatioF()
        w = max(1, round(self._source.width() * self._zoom * dpr))
        h = max(1, round(self._source.height() * self._zoom * dpr))
        scaled = self._source.scaled(w, h, Qt.IgnoreAspectRatio,
                                     Qt.SmoothTransformation)
        scaled.setDevicePixelRatio(dpr)
        return scaled

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.SmoothPixmapTransform, True)
        p.fillRect(self.rect(), QColor("#E9EEF2"))

        # 纸张校样台的细刻度边框
        p.setPen(QPen(QColor("#C2CCD5"), 1, Qt.DashLine))
        stage = self.rect().adjusted(14, 14, -14, -14)
        p.drawRect(stage)

        if self._source.isNull():
            p.setPen(QColor("#6A7682"))
            p.drawText(stage, Qt.AlignCenter | Qt.TextWordWrap, self._message)
            return

        target = self._image_rect()
        p.fillRect(target.adjusted(-1, -1, 1, 1), QColor("#BBC5CE"))
        # 缩放副本带 DPR 标记，drawPixmap 按逻辑坐标绘制后仍是 1:1 物理像素
        p.drawPixmap(target.topLeft(), self._display_pixmap())


class PreviewToolbar(QFrame):
    """校样画布工具条。"""

    openRequested = Signal()

    def __init__(self, canvas: PreviewCanvas, parent=None):
        super().__init__(parent)
        self.canvas = canvas
        self.setObjectName("previewToolbar")
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
        for text, tip, callback in (
            ("适合窗口", "按当前校样台大小显示，不放大超过原图", canvas.fit_to_window),
            ("100%", "按源图像素 1:1 显示（可按住拖动平移）", canvas.actual_size),
            ("−", "缩小校样", canvas.zoom_out),
            ("+", "放大校样", canvas.zoom_in),
        ):
            button = QPushButton(text)
            button.setProperty("role", "compact")
            button.setToolTip(tip)
            button.setAccessibleName(tip)
            button.clicked.connect(callback)
            lay.addWidget(button)
        self.info = QLabel("尚无校样")
        self.info.setObjectName("previewInfo")
        lay.addStretch(1)
        lay.addWidget(self.info)
        self.open_button = QPushButton("独立查看")
        self.open_button.setProperty("role", "compact")
        self.open_button.clicked.connect(self.openRequested)
        lay.addWidget(self.open_button)
        canvas.zoomChanged.connect(self._update_zoom)

    def set_source_info(self, pixmap: QPixmap):
        if pixmap.isNull():
            self.info.setText("尚无校样")
        else:
            self.info.setText(f"{pixmap.width()} × {pixmap.height()} px · {round(self.canvas.zoom * 100)}%")

    def _update_zoom(self, zoom: float):
        prefix = self.info.text().split(" · ")[0]
        self.info.setText(f"{prefix} · {round(zoom * 100)}%")
