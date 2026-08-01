"""集中式视觉主题：精密仪器工作台。"""
from __future__ import annotations

APP_QSS = r"""
* {
    font-family: "Microsoft YaHei UI", "Segoe UI";
    font-size: 13px;
    color: #17212B;
}
QMainWindow, QWidget#appRoot { background: #F4F6F8; }
QFrame#topBar, QFrame#inputPanel, QFrame#workPanel, QFrame#printBar,
QFrame#collapsiblePanel, QFrame#previewPanel {
    background: #FFFFFF;
    border: 1px solid #D7DEE5;
    border-radius: 6px;
}
QLabel#appTitle { font-size: 20px; font-weight: 700; color: #17212B; }
QLabel#appSubtitle { font-size: 12px; color: #6A7682; }
QLabel#sectionTitle { font-size: 14px; font-weight: 600; color: #273440; }
QLabel#sectionHint, QLabel#previewInfo { color: #6A7682; font-size: 12px; }
QFrame#statusPill { background: #EEF2F5; border: 1px solid #D7DEE5; border-radius: 12px; }
QFrame#statusPill[state="success"] { background: #E9F6EF; border-color: #B8DEC8; }
QFrame#statusPill[state="warning"] { background: #FFF5E5; border-color: #E8C98E; }
QFrame#statusPill[state="error"] { background: #FCEDEC; border-color: #E9BBB6; }
QFrame#statusPill[state="success"] QLabel#statusDot { color: #18864B; }
QFrame#statusPill[state="warning"] QLabel#statusDot { color: #B76E00; }
QFrame#statusPill[state="error"] QLabel#statusDot { color: #C0392B; }
QFrame#statusPill[state="neutral"] QLabel#statusDot { color: #758391; }
QLineEdit, QComboBox, QSpinBox {
    background: #FFFFFF; border: 1px solid #C8D1DA; border-radius: 4px;
    min-height: 30px; padding: 0 9px;
    selection-background-color: #087EA4; selection-color: white;
}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus { border: 2px solid #087EA4; padding: 0 8px; }
QLineEdit:disabled, QComboBox:disabled { background: #EEF2F5; color: #7C8894; }
QPushButton {
    min-height: 30px; padding: 0 14px; background: #FFFFFF;
    border: 1px solid #C8D1DA; border-radius: 4px; font-weight: 500;
}
QPushButton:hover { border-color: #087EA4; color: #066C8D; background: #F2F9FC; }
QPushButton:pressed { background: #E4F1F6; }
QPushButton:focus { border: 2px solid #087EA4; }
QPushButton[role="primary"] { background: #087EA4; border-color: #087EA4; color: white; font-weight: 600; }
QPushButton[role="primary"]:hover { background: #066C8D; border-color: #066C8D; color: white; }
QPushButton[role="danger"] { color: #C0392B; }
QPushButton[role="compact"] { min-height: 26px; padding: 0 9px; font-size: 12px; }
QPushButton:disabled { background: #EEF2F5; border-color: #DDE3E8; color: #9AA5AF; }
QTableWidget {
    background: #FFFFFF; alternate-background-color: #F7F9FA;
    border: 1px solid #D7DEE5; border-radius: 4px;
    gridline-color: #E5EAEE; selection-background-color: #DDF1F7;
    selection-color: #17212B;
}
QHeaderView::section {
    background: #EEF2F5; color: #384654; font-weight: 600;
    border: 0; border-right: 1px solid #D7DEE5; border-bottom: 1px solid #D7DEE5;
    padding: 7px 8px;
}
QTableWidget::item { padding: 5px 7px; }
QTabWidget::pane { border: 1px solid #D7DEE5; background: #FFFFFF; border-radius: 4px; }
QTabBar::tab { background: #EEF2F5; border: 1px solid #D7DEE5; border-bottom: 0; padding: 8px 16px; }
QTabBar::tab:selected { background: #FFFFFF; color: #087EA4; font-weight: 600; }
QTabBar::tab:focus { border: 2px solid #087EA4; padding: 7px 15px; }
QProgressBar { border: 1px solid #D7DEE5; border-radius: 3px; background: #EEF2F5; min-height: 7px; max-height: 7px; text-align: center; }
QProgressBar::chunk { background: #087EA4; border-radius: 3px; }
QPlainTextEdit { background: #111820; color: #D7E2EA; border: 0; padding: 8px; font-family: "Cascadia Mono", "Consolas"; font-size: 12px; }
QToolButton#collapseHeader { background: #FFFFFF; border: 1px solid transparent; border-radius: 6px; padding: 6px 9px; text-align: left; font-weight: 600; }
QToolButton#collapseHeader:hover { background: #F3F7F9; }
QToolButton#collapseHeader:focus { border-color: #087EA4; }
QSplitter::handle { background: #D7DEE5; width: 5px; }
QSplitter::handle:hover { background: #87B9CC; }
QToolTip { background: #17212B; color: white; border: 0; padding: 5px 7px; }
QStatusBar { background: #EEF2F5; border-top: 1px solid #D7DEE5; color: #5E6B78; }
QScrollBar:vertical { background: transparent; width: 11px; margin: 0; }
QScrollBar::handle:vertical { background: #C3CDD6; min-height: 30px; border-radius: 5px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
"""


def apply_theme(app) -> None:
    app.setStyleSheet(APP_QSS)
