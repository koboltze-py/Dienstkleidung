"""
DRK Dienstkleidung - Dashboard-Ansicht
Zeigt Uhrzeit und Kalender.
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QFrame, QCalendarWidget,
)
from PySide6.QtCore import Qt, QTimer, QTime, QDate
from PySide6.QtGui import QFont


_WOCHENTAGE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
_MONATE = ["Januar", "Februar", "März", "April", "Mai", "Juni",
           "Juli", "August", "September", "Oktober", "November", "Dezember"]


class DashboardView(QWidget):
    """Dashboard mit Uhrzeit und Kalender."""

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._setup_ui()
        self._update_clock()

        self._clock_timer = QTimer(self)
        self._clock_timer.timeout.connect(self._update_clock)
        self._clock_timer.start(1000)

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 28, 40, 28)
        layout.setSpacing(20)

        lbl_title = QLabel("Dashboard")
        lbl_title.setObjectName("page_title")
        lbl_sub = QLabel("DRK Erste-Hilfe-Station Flughafen Köln")
        lbl_sub.setObjectName("page_subtitle")
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_sub)

        content = QHBoxLayout()
        content.setSpacing(28)
        content.setAlignment(Qt.AlignmentFlag.AlignTop)

        # --- Uhr ---
        clock_frame = QFrame()
        clock_frame.setObjectName("stat_card")
        clock_layout = QVBoxLayout(clock_frame)
        clock_layout.setContentsMargins(40, 36, 40, 36)
        clock_layout.setSpacing(10)
        clock_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self._lbl_time = QLabel("00:00:00")
        time_font = QFont("Segoe UI", 52)
        time_font.setBold(True)
        self._lbl_time.setFont(time_font)
        self._lbl_time.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._lbl_time.setStyleSheet("color: #344F6B; letter-spacing: 2px;")
        clock_layout.addWidget(self._lbl_time)

        self._lbl_date = QLabel("")
        date_font = QFont("Segoe UI", 13)
        self._lbl_date.setFont(date_font)
        self._lbl_date.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._lbl_date.setStyleSheet("color: #555; margin-top: 4px;")
        clock_layout.addWidget(self._lbl_date)

        content.addWidget(clock_frame, stretch=1)

        # --- Kalender ---
        cal_frame = QFrame()
        cal_frame.setObjectName("stat_card")
        cal_layout = QVBoxLayout(cal_frame)
        cal_layout.setContentsMargins(16, 14, 16, 14)

        self._calendar = QCalendarWidget()
        self._calendar.setGridVisible(True)
        self._calendar.setNavigationBarVisible(True)
        self._calendar.setMinimumSize(380, 280)
        cal_layout.addWidget(self._calendar)

        content.addWidget(cal_frame, stretch=2)

        layout.addLayout(content)
        layout.addStretch()

    def _update_clock(self):
        now = QTime.currentTime()
        self._lbl_time.setText(now.toString("HH:mm:ss"))
        today = QDate.currentDate()
        day_name = _WOCHENTAGE[today.dayOfWeek() - 1]
        month_name = _MONATE[today.month() - 1]
        self._lbl_date.setText(f"{day_name}, {today.day()}. {month_name} {today.year()}")

    def showEvent(self, event):
        super().showEvent(event)
        self._update_clock()
