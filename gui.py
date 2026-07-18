"""
Stock Merge - PySide6 GUI 入口。

界面分三区：路径选择 / 主按钮 / 进度与日志。
自动持久化路径到 exe 同目录的 config.json，下次启动预填。
"""
from __future__ import annotations

import os
import platform
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import QObject, QThread, Qt, Signal, QUrl, QSize
from PySide6.QtGui import QAction, QColor, QDesktopServices, QFont, QIcon, QPainter, QPen, QPixmap, QBrush
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QStatusBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

import merge_stock_files as core

APP_NAME = "Stock Merge"
APP_VERSION = "1.0"

# ------------------- 主题（现代暗色 / 渐变高亮 / 玻璃质感） -------------------
# 颜色按层级组织: 背景三阶 / 前景两阶 / 品牌色 / 语义色。
# 所有阴影、圆角、间距都用 token 集中管理,改主题只改这里。
BG_BASE = "#0e0f13"          # 最底层
BG_PANEL = "#181a21"         # 卡片底
BG_RAISED = "#1f222a"        # 升起层(输入框 / 按钮)
BG_HOVER = "#262a33"
BG_FOCUS = "#2d323d"
BORDER = "#2a2e38"
BORDER_LIGHT = "#3a3f4a"
TEXT_PRIMARY = "#e8eaed"
TEXT_SECONDARY = "#b3b8c2"
TEXT_MUTED = "#7a8092"
TEXT_DIM = "#5a6072"

# 品牌色 (cyan -> teal -> purple 渐变)
ACCENT_CYAN = "#22d3ee"
ACCENT_TEAL = "#14b8a6"
ACCENT_PURPLE = "#a855f7"
ACCENT_PINK = "#ec4899"

# 品牌渐变
GRAD_BRAND = f"qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {ACCENT_CYAN}, stop:0.5 {ACCENT_TEAL}, stop:1 {ACCENT_PURPLE})"
GRAD_BG_HERO = f"qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #181a21, stop:1 #0e0f13)"
GRAD_PROGRESS = f"qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {ACCENT_CYAN}, stop:0.5 {ACCENT_TEAL}, stop:1 {ACCENT_PURPLE})"

# 语义色
SUCCESS = "#22c55e"
WARNING = "#f59e0b"
DANGER = "#ef4444"
INFO = "#3b82f6"

# 间距/圆角
RADIUS_SM = 6
RADIUS_MD = 10
RADIUS_LG = 14
RADIUS_XL = 18

# 阴影(QSS 不直接支持 box-shadow,改用 QGraphicsDropShadowEffect 在 widget 上挂)


# ------------------- 图标 (内联 SVG 工厂) -------------------
# 用 SVG 字符串而不是外部文件,避免打包资源路径问题。所有图标 stroke="currentColor",
# 渲染时用 QPainter 把画笔颜色换色,以适配浅/深背景。

_SVG_TEMPLATE = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">{paths}</svg>'

_ICONS: dict[str, str] = {
    "folder": '<path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>',
    "file": '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/>',
    "play": '<polygon points="6 4 20 12 6 20 6 4" fill="{color}" stroke="none"/>',
    "check": '<polyline points="4 12 10 18 20 6"/>',
    "alert": '<circle cx="12" cy="12" r="9"/><line x1="12" y1="8" x2="12" y2="13"/><line x1="12" y1="16" x2="12" y2="16"/>',
    "sparkle": '<path d="M12 3l2 5 5 2-5 2-2 5-2-5-5-2 5-2z"/>',
    "shield": '<path d="M12 3l8 3v6c0 5-3.5 8-8 9-4.5-1-8-4-8-9V6z"/><polyline points="9 12 11 14 15 10"/>',
    "lock": '<rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>',
    "wand": '<path d="M3 21l12-12"/><path d="M14 7l3-3 3 3-3 3z"/><path d="M19 3l1 2 2 1-2 1-1 2-1-2-2-1 2-1z"/>',
}


def make_icon(name: str, color: str = "#e8eaed", size: int = 18) -> QIcon:
    """渲染一个 SVG 图标。color 通常用前景文字色,size 是像素边长。"""
    body = _ICONS.get(name, "")
    svg = _SVG_TEMPLATE.format(color=color, paths=body)
    renderer = QSvgRenderer(svg.encode("utf-8"))
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.Antialiasing)
    renderer.render(painter)
    painter.end()
    return QIcon(pm)


def add_shadow(widget: QWidget, blur: int = 24, dx: int = 0, dy: int = 4,
               color: QColor | None = None) -> QGraphicsDropShadowEffect:
    """给 widget 加一个柔和的投影效果,QSS 替代品。"""
    eff = QGraphicsDropShadowEffect(widget)
    eff.setBlurRadius(blur)
    eff.setOffset(dx, dy)
    eff.setColor(color or QColor(0, 0, 0, 110))
    widget.setGraphicsEffect(eff)
    return eff


# ------------------- 工作线程 -------------------
class Worker(QObject):
    progress = Signal(int, int, str)   # current, total, message
    log = Signal(str)                   # 单行日志
    finished = Signal(object)           # core.ProcessResult
    failed = Signal(str)                # 异常文本

    def __init__(self, cfg: dict):
        super().__init__()
        self._cfg = cfg

    def run(self):
        try:
            def cb(cur: int, total: int, msg: str):
                self.progress.emit(cur, total, msg)
                self.log.emit(msg)
            result = core.process(self._cfg, progress_cb=cb)
            self.finished.emit(result)
        except Exception as e:
            self.failed.emit(f"{type(e).__name__}: {e}")


# ------------------- 主窗口 -------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.setMinimumSize(720, 560)
        self._worker: Worker | None = None
        self._thread: QThread | None = None
        self._build_ui()
        self._load_initial_config()

    # ------------------- UI 构建 -------------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(28, 24, 28, 20)
        root.setSpacing(18)

        root.addWidget(self._build_hero())

        # 路径卡片区
        cards_row = QVBoxLayout()
        cards_row.setSpacing(10)
        self.src_card, self.src_edit, self._src_btn = self._build_path_card(
            "folder", "源文件夹",
            "包含子表 .xlsx 的源目录", self._pick_src
        )
        self.idx_card, self.idx_edit, self._idx_btn = self._build_path_card(
            "file", "索引文件",
            "分类表 分类.xlsx 的绝对路径", self._pick_idx
        )
        self.out_card, self.out_edit, self._out_btn = self._build_path_card(
            "folder", "输出文件夹",
            "合并结果将写入此目录(脚本会先清空,勿选桌面/文档)", self._pick_out
        )
        cards_row.addWidget(self.src_card)
        cards_row.addWidget(self.idx_card)
        cards_row.addWidget(self.out_card)
        root.addLayout(cards_row)

        # 选项条
        opts_row = QHBoxLayout()
        opts_row.setSpacing(20)
        self._num_label = QLabel("统一列数")
        self._num_label.setStyleSheet(f"color: {TEXT_SECONDARY}; font-weight: 600;")
        self.num_spin = QSpinBox()
        self.num_spin.setRange(1, 200)
        self.num_spin.setValue(41)
        self.num_spin.setObjectName("numSpin")
        self.num_spin.valueChanged.connect(self._persist_config)
        self.remember_chk = QCheckBox("完成后自动打开输出目录")
        self.remember_chk.setChecked(True)
        self.remember_chk.setObjectName("rememberChk")
        opts_row.addWidget(self._num_label)
        opts_row.addWidget(self.num_spin)
        opts_row.addStretch(1)
        opts_row.addWidget(self.remember_chk)
        root.addLayout(opts_row)

        # 主按钮
        self.run_btn = QPushButton("  开始合并")
        self.run_btn.setIcon(make_icon("play", "#ffffff", size=18))
        self.run_btn.setIconSize(QSize(18, 18))
        self.run_btn.setObjectName("runBtn")
        self.run_btn.setMinimumHeight(48)
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.clicked.connect(self._on_run_clicked)
        root.addWidget(self.run_btn)

        # 进度 + 状态徽章
        prog_row = QHBoxLayout()
        prog_row.setSpacing(12)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setObjectName("progress")
        self.progress.setTextVisible(True)
        self.progress.setFormat("  %p%  ·  %v / %m")
        prog_row.addWidget(self.progress, 1)
        self._badge = QLabel("就绪")
        self._badge.setObjectName("badge")
        self._badge.setAlignment(Qt.AlignCenter)
        self._badge.setMinimumWidth(86)
        prog_row.addWidget(self._badge)
        root.addLayout(prog_row)

        # 日志
        log_header = QHBoxLayout()
        log_title = QLabel("日志输出")
        log_title.setObjectName("logTitle")
        log_clear = QPushButton("清空")
        log_clear.setObjectName("ghostBtn")
        log_clear.setCursor(Qt.PointingHandCursor)
        log_clear.clicked.connect(lambda: self.log_view.clear())
        log_header.addWidget(log_title)
        log_header.addStretch(1)
        log_header.addWidget(log_clear)
        root.addLayout(log_header)

        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setObjectName("logView")
        self.log_view.setFont(QFont("Menlo, Consolas, monospace", 11))
        self.log_view.setMinimumHeight(200)
        root.addWidget(self.log_view, 1)

        # 状态栏
        sb = QStatusBar()
        sb.setObjectName("statusBar")
        sb.showMessage(f"  {APP_NAME}  ·  v{APP_VERSION}  ·  路径已自动同步到 config.json")
        self.setStatusBar(sb)

        # 全局样式 + 微交互定时器
        self._apply_qss()
        self._install_path_watchers()

        # 阴影:QSS 不支持 box-shadow,改用 QGraphicsDropShadowEffect
        add_shadow(self.run_btn, blur=28, dy=6, color=QColor(34, 211, 238, 70))
        add_shadow(self._badge, blur=14, dy=2, color=QColor(0, 0, 0, 80))

    # ----- Hero 顶部 -----
    def _build_hero(self) -> QFrame:
        hero = QFrame()
        hero.setObjectName("hero")
        layout = QHBoxLayout(hero)
        layout.setContentsMargins(24, 18, 24, 18)
        layout.setSpacing(20)

        # 左侧:大标题 + 副标题
        left = QVBoxLayout()
        left.setSpacing(4)
        title_row = QHBoxLayout()
        title_row.setSpacing(10)
        sparkle = QLabel()
        sparkle.setPixmap(make_icon("sparkle", ACCENT_CYAN, size=26).pixmap(26, 26))
        title = QLabel(APP_NAME)
        title.setObjectName("heroTitle")
        title_row.addWidget(sparkle)
        title_row.addWidget(title)
        title_row.addStretch(1)
        left.addLayout(title_row)
        subtitle = QLabel("按 分类.xlsx 分组合并 Excel 库存表   ·   一键导出按 A 列归类的合并文件")
        subtitle.setObjectName("heroSubtitle")
        subtitle.setWordWrap(True)
        left.addWidget(subtitle)
        left.addStretch(1)

        # 右侧:版本徽章 + 安全徽章
        right = QVBoxLayout()
        right.setSpacing(8)
        right.setAlignment(Qt.AlignTop | Qt.AlignRight)
        ver = QLabel(f"v{APP_VERSION}")
        ver.setObjectName("versionPill")
        ver.setAlignment(Qt.AlignCenter)
        ver.setMinimumWidth(72)
        right.addWidget(ver, 0, Qt.AlignRight)
        safety = QLabel("防误删已启用")
        safety.setObjectName("safetyPill")
        safety.setAlignment(Qt.AlignCenter)
        safety.setMinimumWidth(110)
        right.addWidget(safety, 0, Qt.AlignRight)

        layout.addLayout(left, 1)
        layout.addLayout(right)
        return hero

    # ----- 路径行卡片 -----
    def _build_path_card(self, icon_name: str, label: str, hint: str,
                         on_pick) -> tuple[QFrame, QLineEdit, QPushButton]:
        card = QFrame()
        card.setObjectName("pathCard")
        layout = QHBoxLayout(card)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(14)

        icon_lbl = QLabel()
        icon_lbl.setPixmap(make_icon(icon_name, ACCENT_CYAN, size=20).pixmap(20, 20))
        icon_lbl.setFixedWidth(24)
        icon_lbl.setAlignment(Qt.AlignCenter)
        layout.addWidget(icon_lbl)

        text_col = QVBoxLayout()
        text_col.setSpacing(2)
        lbl = QLabel(label)
        lbl.setObjectName("cardLabel")
        text_col.addWidget(lbl)
        sub = QLabel(hint)
        sub.setObjectName("cardHint")
        sub.setWordWrap(True)
        text_col.addWidget(sub)
        layout.addLayout(text_col, 1)

        edit = QLineEdit()
        edit.setObjectName("pathInput")
        edit.setPlaceholderText(hint)
        edit.editingFinished.connect(self._persist_config)
        layout.addWidget(edit, 2)

        btn = QPushButton("浏览")
        btn.setObjectName("ghostBtn")
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(on_pick)
        layout.addWidget(btn)

        return card, edit, btn

    def _install_path_watchers(self):
        """路径框内容变化时,自动持久化并把 border 染红(空)/正常(有值)。"""
        for edit in (self.src_edit, self.idx_edit, self.out_edit):
            edit.textChanged.connect(lambda _t, e=edit: self._refresh_path_state(e))

    def _refresh_path_state(self, edit: QLineEdit):
        edit.setProperty("hasValue", bool(edit.text().strip()))
        # 触发属性刷新
        edit.style().unpolish(edit)
        edit.style().polish(edit)

    def _set_badge(self, text: str, kind: str = "ready"):
        """更新状态徽章文字 + 颜色。kind: ready / running / ok / err / warn"""
        self._badge.setProperty("kind", kind)
        self._badge.setText(text)
        self._badge.style().unpolish(self._badge)
        self._badge.style().polish(self._badge)

    def _flash_status(self, text: str, ms: int = 1600):
        """状态栏闪一下成功消息,然后自动还原。"""
        sb = self.statusBar()
        sb.showMessage(f"  ✓  {text}")
        # 用单次定时器还原
        from PySide6.QtCore import QTimer
        QTimer.singleShot(ms, lambda: sb.showMessage(
            f"  {APP_NAME}  ·  v{APP_VERSION}  ·  路径已自动同步到 config.json"
        ))

    def _apply_qss(self):
        """全局 QSS。QSS 不支持 box-shadow / transform / animation,所以:
        - 投影靠 QGraphicsDropShadowEffect (在 build_ui 时挂)
        - 按钮 hover 效果靠 :hover 状态 + 颜色变化代替 transform
        """
        self.setStyleSheet(f"""
            /* ---------- 全局基色 ---------- */
            QMainWindow, QWidget {{
                background: {BG_BASE};
                color: {TEXT_PRIMARY};
                font-family: -apple-system, "SF Pro Text", "Helvetica Neue", "PingFang SC", "Microsoft YaHei", sans-serif;
                font-size: 13px;
            }}
            QToolTip {{
                background: {BG_PANEL};
                color: {TEXT_PRIMARY};
                border: 1px solid {BORDER_LIGHT};
                border-radius: 6px;
                padding: 6px 8px;
            }}

            /* ---------- Hero 顶部 ---------- */
            QFrame#hero {{
                background: {GRAD_BG_HERO};
                border: 1px solid {BORDER};
                border-radius: {RADIUS_XL}px;
            }}
            QFrame#hero QLabel#heroTitle {{
                font-size: 24px;
                font-weight: 700;
                color: {TEXT_PRIMARY};
                letter-spacing: -0.3px;
                background: transparent;
            }}
            QFrame#hero QLabel#heroSubtitle {{
                font-size: 13px;
                color: {TEXT_SECONDARY};
                background: transparent;
            }}
            QLabel#versionPill {{
                background: {BG_RAISED};
                color: {TEXT_SECONDARY};
                border: 1px solid {BORDER_LIGHT};
                border-radius: 12px;
                padding: 5px 12px;
                font-weight: 600;
                font-size: 11px;
                letter-spacing: 0.5px;
            }}
            QLabel#safetyPill {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {ACCENT_TEAL}, stop:1 {ACCENT_CYAN});
                color: #061b1c;
                border: none;
                border-radius: 12px;
                padding: 5px 14px;
                font-weight: 700;
                font-size: 11px;
                letter-spacing: 0.3px;
            }}

            /* ---------- 路径卡 ---------- */
            QFrame#pathCard {{
                background: {BG_PANEL};
                border: 1px solid {BORDER};
                border-radius: {RADIUS_MD}px;
            }}
            QFrame#pathCard:hover {{
                border-color: {BORDER_LIGHT};
                background: {BG_RAISED};
            }}
            QFrame#pathCard QLabel#cardLabel {{
                color: {TEXT_PRIMARY};
                font-weight: 600;
                font-size: 13px;
                background: transparent;
            }}
            QFrame#pathCard QLabel#cardHint {{
                color: {TEXT_MUTED};
                font-size: 11px;
                background: transparent;
            }}
            QLineEdit#pathInput {{
                background: {BG_RAISED};
                border: 1px solid {BORDER};
                border-radius: 8px;
                padding: 9px 12px;
                color: {TEXT_PRIMARY};
                selection-background-color: {ACCENT_CYAN};
                selection-color: #061b1c;
            }}
            QLineEdit#pathInput:focus {{
                border: 1px solid {ACCENT_CYAN};
                background: {BG_FOCUS};
            }}
            QLineEdit#pathInput[hasValue="true"] {{
                border-color: {BORDER_LIGHT};
            }}

            /* ---------- 次级按钮 (浏览 / 清空) ---------- */
            QPushButton#ghostBtn {{
                background: {BG_RAISED};
                border: 1px solid {BORDER_LIGHT};
                border-radius: 8px;
                padding: 8px 16px;
                color: {TEXT_SECONDARY};
                font-weight: 500;
            }}
            QPushButton#ghostBtn:hover {{
                background: {BG_HOVER};
                color: {TEXT_PRIMARY};
                border-color: {TEXT_MUTED};
            }}
            QPushButton#ghostBtn:pressed {{
                background: {BG_FOCUS};
            }}

            /* ---------- 主按钮 (渐变 + 大圆角) ---------- */
            QPushButton#runBtn {{
                background: {GRAD_BRAND};
                border: none;
                border-radius: {RADIUS_MD}px;
                color: white;
                font-size: 15px;
                font-weight: 700;
                letter-spacing: 1px;
                padding: 4px 18px;
                text-align: center;
            }}
            QPushButton#runBtn:hover {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {ACCENT_PINK}, stop:0.5 {ACCENT_PURPLE}, stop:1 {ACCENT_CYAN});
            }}
            QPushButton#runBtn:pressed {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 {ACCENT_TEAL}, stop:1 {ACCENT_CYAN});
            }}
            QPushButton#runBtn:disabled {{
                background: {BG_RAISED};
                color: {TEXT_DIM};
            }}

            /* ---------- 进度条 (流光渐变) ---------- */
            QProgressBar#progress {{
                background: {BG_RAISED};
                border: 1px solid {BORDER};
                border-radius: 8px;
                text-align: center;
                color: {TEXT_PRIMARY};
                height: 22px;
                font-weight: 600;
            }}
            QProgressBar#progress::chunk {{
                background: {GRAD_PROGRESS};
                border-radius: 7px;
                margin: 1px;
            }}

            /* ---------- 徽章 ---------- */
            QLabel#badge {{
                background: {BG_RAISED};
                color: {TEXT_SECONDARY};
                border: 1px solid {BORDER_LIGHT};
                border-radius: 12px;
                padding: 6px 14px;
                font-weight: 700;
                font-size: 11px;
                letter-spacing: 0.5px;
            }}
            QLabel#badge[kind="ready"]   {{ background: {BG_RAISED};        color: {TEXT_SECONDARY}; border-color: {BORDER_LIGHT}; }}
            QLabel#badge[kind="running"] {{ background: {ACCENT_CYAN};      color: #061b1c;          border: none; }}
            QLabel#badge[kind="ok"]      {{ background: {SUCCESS};          color: #061b1c;          border: none; }}
            QLabel#badge[kind="err"]     {{ background: {DANGER};           color: white;            border: none; }}
            QLabel#badge[kind="warn"]    {{ background: {WARNING};          color: #1a1206;          border: none; }}

            /* ---------- 选项 (列数 / 复选) ---------- */
            QLabel#logTitle {{
                color: {TEXT_SECONDARY};
                font-weight: 600;
                font-size: 12px;
                letter-spacing: 0.5px;
            }}
            QSpinBox#numSpin {{
                background: {BG_RAISED};
                border: 1px solid {BORDER_LIGHT};
                border-radius: 8px;
                padding: 6px 12px;
                color: {TEXT_PRIMARY};
                selection-background-color: {ACCENT_CYAN};
                min-width: 70px;
            }}
            QSpinBox#numSpin:focus {{
                border: 1px solid {ACCENT_CYAN};
            }}
            QSpinBox#numSpin::up-button, QSpinBox#numSpin::down-button {{
                background: transparent;
                border: none;
                width: 16px;
            }}
            QCheckBox#rememberChk {{
                color: {TEXT_SECONDARY};
                spacing: 8px;
                background: transparent;
            }}
            QCheckBox#rememberChk::indicator {{
                width: 18px;
                height: 18px;
                border: 1.5px solid {BORDER_LIGHT};
                border-radius: 5px;
                background: {BG_RAISED};
            }}
            QCheckBox#rememberChk::indicator:hover {{
                border-color: {ACCENT_CYAN};
            }}
            QCheckBox#rememberChk::indicator:checked {{
                background: {ACCENT_CYAN};
                border: 1.5px solid {ACCENT_CYAN};
                image: none;
            }}
            QCheckBox#rememberChk:checked {{
                color: {TEXT_PRIMARY};
            }}

            /* ---------- 日志 ---------- */
            QTextEdit#logView {{
                background: {BG_PANEL};
                border: 1px solid {BORDER};
                border-radius: {RADIUS_MD}px;
                padding: 12px;
                color: {TEXT_SECONDARY};
                selection-background-color: {ACCENT_PURPLE};
                selection-color: white;
            }}

            /* ---------- 状态栏 ---------- */
            QStatusBar#statusBar {{
                background: {BG_PANEL};
                color: {TEXT_MUTED};
                border-top: 1px solid {BORDER};
                font-size: 11px;
            }}
            QStatusBar#statusBar::item {{
                border: none;
            }}
        """)

    # ------------------- 控件辅助 -------------------
    def _mk_btn(self, text: str, slot) -> QPushButton:
        b = QPushButton(text)
        b.setCursor(Qt.PointingHandCursor)
        b.clicked.connect(slot)
        return b

    def _h(self, *widgets) -> QWidget:
        w = QWidget()
        layout = QHBoxLayout(w)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        for x in widgets:
            layout.addWidget(x)
        layout.itemAt(0).widget().setSizeIncrement(0, 0)
        return w

    def _wrap(self, layout) -> QWidget:
        w = QWidget()
        w.setLayout(layout)
        return w

    def _labeled(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet(f"color: {TEXT_MUTED}; min-width: 90px;")
        return lbl

    # ------------------- 路径选择 -------------------
    def _pick_src(self):
        d = QFileDialog.getExistingDirectory(self, "选择源文件夹", self.src_edit.text() or str(Path.home()))
        if d:
            self.src_edit.setText(d)
            self._persist_config()

    def _pick_idx(self):
        f, _ = QFileDialog.getOpenFileName(
            self, "选择索引文件", self.idx_edit.text() or str(Path.home()),
            "Excel 文件 (*.xlsx)"
        )
        if f:
            self.idx_edit.setText(f)
            self._persist_config()

    def _pick_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出文件夹", self.out_edit.text() or str(Path.home()))
        if d:
            self.out_edit.setText(d)
            self._persist_config()

    # ------------------- 配置持久化 -------------------
    def _load_initial_config(self):
        cfg = core.load_config()
        # 安全防御:启动时若持久化的路径指向危险位置,自动清空(避免下次 Run 时 rmtree)
        for field in ("src_dir", "index_file", "output_dir"):
            v = cfg.get(field, "")
            if v and core._is_dangerous_output_dir(Path(v)) is not None:
                self._log(
                    f"[安全] 检测到 {field} 指向 {v} (危险路径),已自动清空"
                    f"——请重新选择更具体的子目录。"
                )
                cfg[field] = ""

        self.src_edit.setText(cfg.get("src_dir", ""))
        self.idx_edit.setText(cfg.get("index_file", ""))
        self.out_edit.setText(cfg.get("output_dir", ""))
        try:
            self.num_spin.setValue(int(cfg.get("num_cols", 41)))
        except Exception:
            pass
        # 把清理过的 cfg 立刻回写,避免下次启动再触发同样的"被修改成危险路径"问题
        try:
            core.save_config(cfg)
        except Exception as e:
            self._log(f"[配置回写失败] {e}")

        # 触发 path 状态刷新 + 徽章初始化
        for e in (self.src_edit, self.idx_edit, self.out_edit):
            self._refresh_path_state(e)
        self._set_badge("就绪", "ready")

    def _current_cfg(self) -> dict:
        return {
            "src_dir": self.src_edit.text().strip(),
            "index_file": self.idx_edit.text().strip(),
            "output_dir": self.out_edit.text().strip(),
            "num_cols": self.num_spin.value(),
        }

    def _persist_config(self):
        try:
            core.save_config(self._current_cfg())
            self._flash_status("已保存到 config.json")
        except Exception as e:
            self._log(f"[配置保存失败] {e}")

    # ------------------- 日志输出 -------------------
    def _log(self, line: str):
        """输出带时间戳的日志,按关键字自动染色 + 联动状态徽章。

        配色规则:
          - [错误]/Error/Exception/失败   → 红色
          - [警告]/警告/跳过              → 黄色
          - ✓/完成/成功                   → 绿色
          - [安全]/防误删                 → 青色
          - 其它                          → 默认前景色
        """
        from datetime import datetime
        from PySide6.QtGui import QTextCharFormat, QTextCursor
        ts = datetime.now().strftime("%H:%M:%S")
        text = f"[{ts}] {line}\n"

        cursor = self.log_view.textCursor()
        cursor.movePosition(QTextCursor.End)
        fmt = QTextCharFormat()
        lower = line.lower()
        if any(k in line for k in ("[错误]", "Exception", "失败", "ERROR", "Traceback")):
            fmt.setForeground(QColor(DANGER))
            self._set_badge("异常", "err")
        elif any(k in line for k in ("[警告]", "警告", "跳过", "缺文件", "WARN")):
            fmt.setForeground(QColor(WARNING))
            self._set_badge("警告", "warn")
        elif any(k in line for k in ("✓", "完成", "成功", "[结果]")):
            fmt.setForeground(QColor(SUCCESS))
            self._set_badge("已完成", "ok")
        elif any(k in line for k in ("[安全]", "防误删", "BLOCKED")):
            fmt.setForeground(QColor(ACCENT_CYAN))
        elif line.startswith("[过滤]") or line.startswith("=== "):
            fmt.setForeground(QColor(ACCENT_PURPLE))
        elif line.startswith("[") and "]" in line and len(line.split("]")[0]) < 12:
            fmt.setForeground(QColor(INFO))
        cursor.insertText(text, fmt)

        sb = self.log_view.verticalScrollBar()
        sb.setValue(sb.maximum())

    # ------------------- 运行 -------------------
    def _validate(self) -> str | None:
        cfg = self._current_cfg()
        for k in ("src_dir", "index_file", "output_dir"):
            if not cfg[k]:
                return f"请填写 {k}"
        return None

    @staticmethod
    def _summarize_existing(out: Path) -> tuple[int, str] | None:
        """统计 out 目录里已有文件的数量和总体积;若无内容返回 None。"""
        if not out.exists():
            return None
        try:
            n = 0
            total = 0
            for root, dirs, files in os.walk(out):
                for f in files:
                    p = Path(root) / f
                    try:
                        total += p.stat().st_size
                        n += 1
                    except OSError:
                        continue
            if n == 0:
                return None
            # 人性化的体积
            for unit in ("B", "KB", "MB", "GB"):
                if total < 1024 or unit == "GB":
                    return n, f"{total:.1f} {unit}"
                total /= 1024
        except OSError:
            pass
        return None

    def _on_run_clicked(self):
        err = self._validate()
        if err:
            QMessageBox.warning(self, APP_NAME, err)
            return
        cfg = self._current_cfg()
        out = Path(cfg["output_dir"])

        # 二次确认:output_dir 里已有文件,数量/大小暴露
        existing_summary = self._summarize_existing(out)
        if existing_summary is not None:
            n, size_human = existing_summary
            self._log(f"[安全检查] output_dir 已存在内容: {n} 个文件 / 约 {size_human}")
            btn = QMessageBox.question(
                self, APP_NAME,
                f"输出目录已存在内容:\n  {out}\n"
                f"  共 {n} 个文件,约 {size_human}\n\n"
                f"脚本会先【清空整个目录】再写入合并结果。\n"
                f"是否继续?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if btn != QMessageBox.Yes:
                self._log("[取消] 用户放弃合并")
                return

        self._persist_config()
        self.run_btn.setEnabled(False)
        self.run_btn.setText("  处理中…")
        self.log_view.clear()
        self.progress.setValue(0)
        self.progress.setFormat("  %p%  ·  %v / %m")
        self._set_badge("运行中", "running")
        self._log(f"启动合并: src={cfg['src_dir']}  →  out={cfg['output_dir']}")

        self._thread = QThread(self)
        self._worker = Worker(cfg)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._on_progress)
        self._worker.log.connect(self._log)
        self._worker.finished.connect(self._on_finished)
        self._worker.failed.connect(self._on_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.start()

    def _on_progress(self, cur: int, total: int, msg: str):
        self.progress.setRange(0, total)
        self.progress.setValue(cur)
        self.progress.setFormat(f"  %p%  ·  {cur} / {total}  ·  {msg}")

    def _on_finished(self, result):
        self.run_btn.setEnabled(True)
        self.run_btn.setText("  开始合并")
        miss = len(result.missing_files)
        summary = (
            f"合并完成: {result.groups_merged}/{result.groups_total} 组，"
            f"有效 {result.rows_total} 行，过滤 {result.rows_filtered_total} 行「不显示」，"
            f"复制 {result.files_copied} 个源文件"
        )
        if miss:
            summary += f"，跳过 {miss} 个缺文件"
        self._log(summary)
        self.progress.setValue(self.progress.maximum())
        self.progress.setFormat("  ✓ 完成  ·  100%")
        self._set_badge("完成", "ok")
        QMessageBox.information(
            self, APP_NAME,
            summary + "\n\n输出目录已自动打开。" if self.remember_chk.isChecked() else summary
        )
        if self.remember_chk.isChecked():
            self._open_in_file_manager(Path(self._current_cfg()["output_dir"]))

    def _on_failed(self, msg: str):
        self.run_btn.setEnabled(True)
        self.run_btn.setText("  开始合并")
        self._log(f"[错误] {msg}")
        self._set_badge("失败", "err")
        QMessageBox.critical(self, APP_NAME, f"合并失败:\n\n{msg}")

    @staticmethod
    def _open_in_file_manager(p: Path):
        p = str(p)
        sysname = platform.system()
        try:
            if sysname == "Darwin":
                subprocess.Popen(["open", p])
            elif sysname == "Windows":
                os.startfile(p)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", p])
        except Exception as e:
            print(f"[打开文件管理器失败] {e}")


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    win = MainWindow()
    win.show()
    # 仅用于打包后自动化测试：环境变量驱动自动点开始
    if os.environ.get("STOCK_MERGE_AUTO_RUN") == "1":
        from PySide6.QtCore import QTimer
        def _maybe_quit():
            if win.run_btn.isEnabled() and win.run_btn.text() == "开始合并" and win.progress.maximum() > 0:
                QTimer.singleShot(500, app.quit)
            else:
                QTimer.singleShot(200, _maybe_quit)
        QTimer.singleShot(300, win._on_run_clicked)
        QTimer.singleShot(500, _maybe_quit)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
