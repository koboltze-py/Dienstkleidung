"""
DRK Dienstkleidung - Dashboard-Ansicht
Zeigt Uhrzeit, Kalender und niedrige Bestände.
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QFrame, QCalendarWidget, QTableWidget, QTableWidgetItem,
    QHeaderView, QScrollArea, QPushButton, QMessageBox,
)
from PySide6.QtCore import Qt, QTimer, QTime, QDate
from PySide6.QtGui import QFont, QColor


_WOCHENTAGE = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]
_MONATE = ["Januar", "Februar", "März", "April", "Mai", "Juni",
           "Juli", "August", "September", "Oktober", "November", "Dezember"]


class DashboardView(QWidget):
    """Dashboard mit Uhrzeit, Kalender und Niedrig-Bestand-Übersicht."""

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._setup_ui()
        self._update_clock()
        self._refresh_low_stock()

        self._clock_timer = QTimer(self)
        self._clock_timer.timeout.connect(self._update_clock)
        self._clock_timer.start(1000)

        self._stock_timer = QTimer(self)
        self._stock_timer.timeout.connect(self._refresh_low_stock)
        self._stock_timer.start(60_000)

    def _setup_ui(self):
        outer = QScrollArea(self)
        outer.setWidgetResizable(True)
        outer.setFrameShape(QFrame.Shape.NoFrame)
        container = QWidget()
        outer.setWidget(container)

        layout = QVBoxLayout(container)
        layout.setContentsMargins(32, 24, 32, 24)
        layout.setSpacing(20)

        lbl_title = QLabel("Dashboard")
        lbl_title.setObjectName("page_title")
        lbl_sub = QLabel("DRK Erste-Hilfe-Station Flughafen Köln")
        lbl_sub.setObjectName("page_subtitle")
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_sub)

        # Obere Zeile: Uhr + Kalender
        top_row = QHBoxLayout()
        top_row.setSpacing(24)
        top_row.setAlignment(Qt.AlignmentFlag.AlignTop)

        # --- Uhr ---
        clock_frame = QFrame()
        clock_frame.setObjectName("stat_card")
        clock_layout = QVBoxLayout(clock_frame)
        clock_layout.setContentsMargins(40, 40, 40, 40)
        clock_layout.setSpacing(12)
        clock_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self._lbl_time = QLabel("00:00:00")
        time_font = QFont("Segoe UI", 80)
        time_font.setBold(True)
        self._lbl_time.setFont(time_font)
        self._lbl_time.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._lbl_time.setStyleSheet("color: #344F6B; letter-spacing: 4px;")
        clock_layout.addWidget(self._lbl_time)

        self._lbl_date = QLabel("")
        date_font = QFont("Segoe UI", 15)
        self._lbl_date.setFont(date_font)
        self._lbl_date.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._lbl_date.setStyleSheet("color: #555; margin-top: 6px;")
        clock_layout.addWidget(self._lbl_date)

        top_row.addWidget(clock_frame, stretch=3)

        # --- Kalender ---
        cal_frame = QFrame()
        cal_frame.setObjectName("stat_card")
        cal_layout = QVBoxLayout(cal_frame)
        cal_layout.setContentsMargins(14, 12, 14, 12)

        self._calendar = QCalendarWidget()
        self._calendar.setGridVisible(True)
        self._calendar.setNavigationBarVisible(True)
        self._calendar.setMinimumSize(360, 260)
        cal_layout.addWidget(self._calendar)

        top_row.addWidget(cal_frame, stretch=2)
        layout.addLayout(top_row)

        # --- Niedrig-Bestand ---
        low_frame = QFrame()
        low_frame.setObjectName("stat_card")
        low_layout = QVBoxLayout(low_frame)
        low_layout.setContentsMargins(16, 14, 16, 14)
        low_layout.setSpacing(10)

        hdr_row = QHBoxLayout()
        lbl_low = QLabel("⚠  Niedriger Bestand")
        f = QFont(); f.setBold(True); f.setPointSize(12)
        lbl_low.setFont(f)
        lbl_low.setStyleSheet("color: #E65100;")
        hdr_row.addWidget(lbl_low)
        hdr_row.addStretch()
        self._lbl_low_count = QLabel("")
        self._lbl_low_count.setStyleSheet("color:#888; font-size:11px;")
        hdr_row.addWidget(self._lbl_low_count)

        btn_set_min = QPushButton("Mindestbestand 3 setzen (alle ohne Wert)")
        btn_set_min.setObjectName("btn_secondary")
        btn_set_min.setToolTip("Setzt min_menge=3 für alle Artikel, die noch keinen Mindestbestand haben")
        btn_set_min.clicked.connect(self._set_default_min)
        hdr_row.addWidget(btn_set_min)
        low_layout.addLayout(hdr_row)

        self._tbl_low = QTableWidget(0, 4)
        self._tbl_low.setHorizontalHeaderLabels(["Kleidungsart", "Größe", "Auf Lager", "Mindestbestand"])
        self._tbl_low.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._tbl_low.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_low.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_low.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_low.verticalHeader().setVisible(False)
        self._tbl_low.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._tbl_low.setAlternatingRowColors(True)
        self._tbl_low.setMaximumHeight(260)
        low_layout.addWidget(self._tbl_low)

        layout.addWidget(low_frame)
        layout.addStretch()

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.addWidget(outer)

    def _update_clock(self):
        now = QTime.currentTime()
        self._lbl_time.setText(now.toString("HH:mm:ss"))
        today = QDate.currentDate()
        day_name = _WOCHENTAGE[today.dayOfWeek() - 1]
        month_name = _MONATE[today.month() - 1]
        self._lbl_date.setText(f"{day_name}, {today.day()}. {month_name} {today.year()}")

    def _refresh_low_stock(self):
        if not self.db.kleidung_db_exists():
            return
        low = self.db.get_niedrig_bestand()
        warn_color = QColor("#FFF3CD")
        crit_color = QColor("#FFCCBC")
        self._tbl_low.setRowCount(len(low))
        for r, item in enumerate(low):
            menge = int(item["menge"])
            min_m = int(item["min_menge"])
            cells = [
                item["art_name"],
                str(item["groesse"]),
                str(menge),
                str(min_m),
            ]
            color = crit_color if menge == 0 else warn_color
            for c, txt in enumerate(cells):
                it = QTableWidgetItem(txt)
                it.setBackground(color)
                self._tbl_low.setItem(r, c, it)
        count = len(low)
        self._lbl_low_count.setText(f"{count} Artikel" if count else "Alles in Ordnung ✓")
        self._lbl_low_count.setStyleSheet(
            "color: #E65100; font-weight:bold;" if count else "color: #388E3C; font-weight:bold;"
        )

    def _set_default_min(self):
        if not self.db.kleidung_db_exists():
            return
        count = self.db.set_default_min_menge(3)
        self._refresh_low_stock()
        QMessageBox.information(
            self,
            "Mindestbestand gesetzt",
            f"{count} Artikel auf Mindestbestand 3 gesetzt." if count
            else "Alle Artikel haben bereits einen Mindestbestand.",
        )

    def showEvent(self, event):
        super().showEvent(event)
        self._update_clock()
        self._refresh_low_stock()
