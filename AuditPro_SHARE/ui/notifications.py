"""
Système de notifications toast pour AuditPro.
Gère les notifications non bloquantes avec animation et pile verticale.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Callable, Optional

from PyQt6.QtCore import QEasingCurve, QPropertyAnimation, QRect, QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QHBoxLayout, QLabel, QProgressBar, QPushButton, QVBoxLayout, QWidget


class NotificationType(Enum):
    """Types de notifications disponibles."""

    INFO = "info"
    SUCCESS = "success"
    WARNING = "warning"
    ERROR = "error"
    PROGRESS = "progress"


@dataclass
class NotificationConfig:
    """Configuration d'une notification."""

    type: NotificationType
    title: str
    message: str
    duration: int = 3000
    action_text: Optional[str] = None
    action_callback: Optional[Callable[[], None]] = None


class Toast(QWidget):
    """Widget toast animé."""

    closed = pyqtSignal()
    action_triggered = pyqtSignal()

    COLORS = {
        NotificationType.INFO: {"bg": "#E3F2FD", "text": "#1565C0", "border": "#90CAF9", "accent": "#1565C0"},
        NotificationType.SUCCESS: {"bg": "#E8F5E9", "text": "#2E7D32", "border": "#81C784", "accent": "#2E7D32"},
        NotificationType.WARNING: {"bg": "#FFF3E0", "text": "#E65100", "border": "#FFB74D", "accent": "#E65100"},
        NotificationType.ERROR: {"bg": "#FFEBEE", "text": "#C62828", "border": "#EF5350", "accent": "#C62828"},
        NotificationType.PROGRESS: {"bg": "#F3E5F5", "text": "#6A1B9A", "border": "#BA68C8", "accent": "#6A1B9A"},
    }

    ICONS = {
        NotificationType.INFO: "i",
        NotificationType.SUCCESS: "OK",
        NotificationType.WARNING: "!",
        NotificationType.ERROR: "x",
        NotificationType.PROGRESS: "...",
    }

    def __init__(self, config: NotificationConfig, parent: QWidget | None = None):
        super().__init__(parent)
        self.config = config
        self._closing = False

        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.ToolTip)
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating, True)

        self._build_ui()
        self._setup_styling()
        self._setup_animations()

        if self.config.type != NotificationType.PROGRESS:
            self._setup_auto_close()

    def _build_ui(self):
        self.setMinimumWidth(360)
        self.setMaximumWidth(420)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(8)

        header_layout = QHBoxLayout()
        header_layout.setSpacing(10)

        icon_label = QLabel(self.ICONS.get(self.config.type, "•"))
        icon_label.setObjectName("ToastIcon")
        icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_label.setFixedSize(28, 28)
        icon_label.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        header_layout.addWidget(icon_label)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)

        title_label = QLabel(self.config.title)
        title_label.setObjectName("ToastTitle")
        title_label.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        text_layout.addWidget(title_label)

        self.message_label = QLabel(self.config.message)
        self.message_label.setObjectName("ToastMessage")
        self.message_label.setWordWrap(True)
        self.message_label.setFont(QFont("Segoe UI", 10))
        text_layout.addWidget(self.message_label)

        header_layout.addLayout(text_layout, 1)

        close_btn = QPushButton("×")
        close_btn.setObjectName("ToastCloseButton")
        close_btn.setFixedSize(24, 24)
        close_btn.clicked.connect(self.dismiss)
        header_layout.addWidget(close_btn, 0, Qt.AlignmentFlag.AlignTop)

        layout.addLayout(header_layout)

        self.progress_bar: Optional[QProgressBar] = None
        if self.config.type == NotificationType.PROGRESS:
            self.progress_bar = QProgressBar()
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.progress_bar.setTextVisible(False)
            self.progress_bar.setFixedHeight(6)
            layout.addWidget(self.progress_bar)

        if self.config.action_text:
            btn_layout = QHBoxLayout()
            btn_layout.addStretch()
            action_btn = QPushButton(self.config.action_text)
            action_btn.setObjectName("ToastActionButton")
            action_btn.clicked.connect(self._on_action)
            btn_layout.addWidget(action_btn)
            layout.addLayout(btn_layout)

        self.adjustSize()

    def _setup_styling(self):
        colors = self.COLORS[self.config.type]
        self.setStyleSheet(
            f"""
            Toast {{
                background-color: {colors['bg']};
                border: 1px solid {colors['border']};
                border-radius: 10px;
                color: {colors['text']};
            }}
            QLabel#ToastIcon {{
                background-color: rgba(255, 255, 255, 0.65);
                border: 1px solid {colors['border']};
                border-radius: 14px;
                color: {colors['accent']};
            }}
            QPushButton#ToastCloseButton {{
                border: none;
                background: transparent;
                color: {colors['text']};
                font-size: 16px;
            }}
            QPushButton#ToastCloseButton:hover {{
                background-color: rgba(0, 0, 0, 0.06);
                border-radius: 12px;
            }}
            QPushButton#ToastActionButton {{
                background-color: transparent;
                color: {colors['accent']};
                border: 1px solid {colors['accent']};
                border-radius: 6px;
                padding: 6px 12px;
                font-weight: bold;
            }}
            QPushButton#ToastActionButton:hover {{
                background-color: rgba(255, 255, 255, 0.45);
            }}
            QProgressBar {{
                border: none;
                background-color: rgba(0, 0, 0, 0.1);
                border-radius: 3px;
            }}
            QProgressBar::chunk {{
                background-color: {colors['accent']};
                border-radius: 3px;
            }}
            """
        )

    def _setup_animations(self):
        self.animation_in = QPropertyAnimation(self, b"geometry")
        self.animation_in.setDuration(220)
        self.animation_in.setEasingCurve(QEasingCurve.Type.OutCubic)

        self.animation_out = QPropertyAnimation(self, b"geometry")
        self.animation_out.setDuration(180)
        self.animation_out.setEasingCurve(QEasingCurve.Type.InCubic)
        self.animation_out.finished.connect(self._on_animation_finished)

    def _setup_auto_close(self):
        timer = QTimer(self)
        timer.setSingleShot(True)
        timer.timeout.connect(self.dismiss)
        timer.start(self.config.duration)

    def show_animated(self, target_geometry: QRect):
        start_geometry = QRect(
            target_geometry.x() + 24,
            target_geometry.y(),
            target_geometry.width(),
            target_geometry.height(),
        )
        self.setGeometry(start_geometry)
        self.show()
        self.raise_()
        self.animation_in.stop()
        self.animation_in.setStartValue(start_geometry)
        self.animation_in.setEndValue(target_geometry)
        self.animation_in.start()

    def move_animated(self, target_geometry: QRect):
        self.animation_in.stop()
        self.animation_in.setStartValue(self.geometry())
        self.animation_in.setEndValue(target_geometry)
        self.animation_in.start()

    def update_message(self, message: str):
        self.message_label.setText(message)
        self.adjustSize()

    def set_progress(self, value: int):
        if self.progress_bar is not None:
            self.progress_bar.setValue(max(0, min(100, value)))

    def dismiss(self):
        if self._closing:
            return
        self._closing = True
        # Stop any in-progress repositioning animation so it cannot interfere
        # with the slide-out animation that follows.
        self.animation_in.stop()
        current = self.geometry()
        end_rect = QRect(current.x() + 30, current.y(), current.width(), current.height())
        self.animation_out.stop()
        self.animation_out.setStartValue(current)
        self.animation_out.setEndValue(end_rect)
        self.animation_out.start()

    def _on_action(self):
        self.action_triggered.emit()
        if self.config.action_callback:
            self.config.action_callback()
        self.dismiss()

    def _on_animation_finished(self):
        self.closed.emit()
        self.hide()
        self.deleteLater()


class NotificationManager(QWidget):
    """Gestionnaire de toasts rattaché à une fenêtre parente."""

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.notifications: list[Toast] = []
        self.spacing = 12
        self.margin_top = 20
        self.margin_right = 20

    def show(self, config: NotificationConfig) -> Toast:
        toast = Toast(config, self.parentWidget())
        toast.closed.connect(lambda: self._on_notification_closed(toast))
        self.notifications.append(toast)
        self._position_notifications()
        return toast

    def show_info(
        self,
        title: str,
        message: str,
        duration: int = 3000,
        action_text: Optional[str] = None,
        action_callback: Optional[Callable[[], None]] = None,
    ) -> Toast:
        return self.show(NotificationConfig(NotificationType.INFO, title, message, duration, action_text, action_callback))

    def show_success(
        self,
        title: str,
        message: str,
        duration: int = 3000,
        action_text: Optional[str] = None,
        action_callback: Optional[Callable[[], None]] = None,
    ) -> Toast:
        return self.show(NotificationConfig(NotificationType.SUCCESS, title, message, duration, action_text, action_callback))

    def show_warning(
        self,
        title: str,
        message: str,
        duration: int = 4500,
        action_text: Optional[str] = None,
        action_callback: Optional[Callable[[], None]] = None,
    ) -> Toast:
        return self.show(NotificationConfig(NotificationType.WARNING, title, message, duration, action_text, action_callback))

    def show_error(
        self,
        title: str,
        message: str,
        duration: int = 5500,
        action_text: Optional[str] = None,
        action_callback: Optional[Callable[[], None]] = None,
    ) -> Toast:
        return self.show(NotificationConfig(NotificationType.ERROR, title, message, duration, action_text, action_callback))

    def show_progress(self, title: str, message: str) -> Toast:
        return self.show(NotificationConfig(NotificationType.PROGRESS, title, message, duration=60_000))

    def reposition(self):
        self._position_notifications()

    def _position_notifications(self):
        parent = self.parentWidget()
        if parent is None:
            return

        y = self.margin_top
        available_width = max(320, parent.width() - self.margin_right)

        for notification in list(self.notifications):
            # Skip toasts that are already animating out — repositioning them
            # would start animation_in and conflict with the dismiss animation.
            if notification._closing:
                continue
            notification.adjustSize()
            width = min(notification.width(), available_width)
            height = notification.height()
            x = parent.width() - width - self.margin_right
            geometry = QRect(x, y, width, height)
            if notification.isVisible():
                notification.move_animated(geometry)
            else:
                notification.show_animated(geometry)
            y += height + self.spacing

    def _on_notification_closed(self, toast: Toast):
        if toast in self.notifications:
            self.notifications.remove(toast)
        self._position_notifications()