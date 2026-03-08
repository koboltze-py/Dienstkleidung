"""
DRK Dienstkleidung - Dashboard-Ansicht
Zeigt Statistiken, Warnungen und aktuelle Aktivitäten.
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QFrame, QScrollArea, QTableWidget, QTableWidgetItem,
    QPushButton, QSizePolicy,
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont

from utils import format_datum


def _stat_card(value: str, label: str, warn: bool = False) -> QFrame:
    """Erstellt eine Statistik-Karte."""
    frame = QFrame()
    frame.setObjectName("stat_card_warn" if warn else "stat_card")
    frame.setFixedHeight(100)
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(16, 14, 16, 14)
    layout.setSpacing(4)

    lbl_val = QLabel(str(value))
    lbl_val.setObjectName("stat_value_warn" if warn else "stat_value")
    font = QFont()
    font.setPointSize(24)
    font.setBold(True)
    lbl_val.setFont(font)
    lbl_val.setAlignment(Qt.AlignmentFlag.AlignLeft)

    lbl_txt = QLabel(label.upper())
    lbl_txt.setObjectName("stat_label")
    lbl_txt.setAlignment(Qt.AlignmentFlag.AlignLeft)

    layout.addWidget(lbl_val)
    layout.addWidget(lbl_txt)
    return frame, lbl_val


class DashboardView(QWidget):
    """Übersichts-Dashboard der App."""

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._setup_ui()
        self._refresh()

        # Automatisch alle 60 Sekunden aktualisieren
        self._timer = QTimer(self)
        self._timer.timeout.connect(self._refresh)
        self._timer.start(60_000)

    def _setup_ui(self):
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        scroll.setWidget(container)
        self._main_layout = QVBoxLayout(container)
        self._main_layout.setContentsMargins(24, 20, 24, 20)
        self._main_layout.setSpacing(20)

        # Titel
        lbl_title = QLabel("Dashboard")
        lbl_title.setObjectName("page_title")
        lbl_sub = QLabel("Übersicht Dienstkleidung · DRK Erste-Hilfe-Station Flughafen Köln")
        lbl_sub.setObjectName("page_subtitle")

        self._main_layout.addWidget(lbl_title)
        self._main_layout.addWidget(lbl_sub)

        # Stat-Karten
        cards_row = QHBoxLayout()
        cards_row.setSpacing(16)

        (c2, self._val_types)     = _stat_card("–", "Kleidungsarten")
        (c3, self._val_emp)       = _stat_card("–", "MA mit Kleidung")
        (c4, self._val_low)       = _stat_card("–", "Niedriger Bestand", warn=True)
        (c5, self._val_today)     = _stat_card("–", "Ausgaben heute")

        for card in (c2, c3, c4, c5):
            card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            cards_row.addWidget(card)

        self._main_layout.addLayout(cards_row)

        # Niedrig-Bestand + Letzte Aktivitäten nebeneinander
        two_col = QHBoxLayout()
        two_col.setSpacing(16)

        # Linke Spalte: Niedriger Bestand
        left_widget = self._build_low_stock_panel()
        two_col.addWidget(left_widget, stretch=1)

        # Rechte Spalte: Letzte Buchungen
        right_widget = self._build_recent_panel()
        two_col.addWidget(right_widget, stretch=2)

        self._main_layout.addLayout(two_col)
        self._main_layout.addStretch()

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

    def _build_low_stock_panel(self) -> QWidget:
        frame = QFrame()
        frame.setObjectName("stat_card")
        frame.setMinimumHeight(240)
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(8)

        header = QHBoxLayout()
        lbl = QLabel("⚠ Niedriger Bestand")
        lbl_font = QFont()
        lbl_font.setBold(True)
        lbl.setFont(lbl_font)
        lbl.setStyleSheet("color: #F57C00;")
        header.addWidget(lbl)
        header.addStretch()
        layout.addLayout(header)

        self._tbl_low = QTableWidget(0, 3)
        self._tbl_low.setHorizontalHeaderLabels(["Kleidungsart", "Größe", "Lager / Min."])
        self._tbl_low.horizontalHeader().setStretchLastSection(True)
        self._tbl_low.verticalHeader().setVisible(False)
        self._tbl_low.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._tbl_low.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self._tbl_low.setAlternatingRowColors(True)
        layout.addWidget(self._tbl_low)
        return frame

    def _build_recent_panel(self) -> QWidget:
        frame = QFrame()
        frame.setObjectName("stat_card")
        frame.setMinimumHeight(240)
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(8)

        lbl = QLabel("Letzte Buchungen")
        lbl_font = QFont()
        lbl_font.setBold(True)
        lbl.setFont(lbl_font)
        layout.addWidget(lbl)

        self._tbl_recent = QTableWidget(0, 5)
        self._tbl_recent.setHorizontalHeaderLabels(
            ["Datum", "Typ", "Kleidungsart", "Größe / Menge", "Mitarbeiter"]
        )
        self._tbl_recent.horizontalHeader().setStretchLastSection(True)
        self._tbl_recent.verticalHeader().setVisible(False)
        self._tbl_recent.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self._tbl_recent.setAlternatingRowColors(True)
        layout.addWidget(self._tbl_recent)
        return frame

    def _refresh(self):
        """Daten neu laden."""
        if not self.db.kleidung_db_exists():
            return

        stats = self.db.get_dashboard_stats()
        self._val_types.setText(str(stats["total_types"]))
        self._val_emp.setText(str(stats["employees_with_clothing"]))
        self._val_low.setText(str(stats["low_stock"]))
        self._val_today.setText(str(stats["today_ausgabe"]))

        # Niedrig-Bestand-Tabelle
        low = self.db.get_niedrig_bestand()
        self._tbl_low.setRowCount(len(low))
        for r, item in enumerate(low):
            self._tbl_low.setItem(r, 0, QTableWidgetItem(item["art_name"]))
            self._tbl_low.setItem(r, 1, QTableWidgetItem(str(item["groesse"])))
            self._tbl_low.setItem(r, 2, QTableWidgetItem(
                f"{item['menge']} / {item['min_menge']}"
            ))
            for c in range(3):
                it = self._tbl_low.item(r, c)
                if it:
                    it.setBackground(Qt.GlobalColor.yellow)

        # Letzte Buchungen
        recent = self.db.get_recent_buchungen(15)
        TYP_LABEL = {
            "ausgabe": "📤 Ausgabe",
            "rueckgabe": "📥 Rückgabe",
            "eingang": "📦 Eingang",
            "korrektur": "✏ Korrektur",
        }
        self._tbl_recent.setRowCount(len(recent))
        for r, b in enumerate(recent):
            datum_str = format_datum(b.get("datum", ""))
            typ_str = TYP_LABEL.get(b.get("typ", ""), b.get("typ", ""))
            menge_str = str(abs(b.get("menge", 0))) + " Stk."
            self._tbl_recent.setItem(r, 0, QTableWidgetItem(datum_str))
            self._tbl_recent.setItem(r, 1, QTableWidgetItem(typ_str))
            self._tbl_recent.setItem(r, 2, QTableWidgetItem(b.get("art_name", "")))
            self._tbl_recent.setItem(r, 3, QTableWidgetItem(
                f"{b.get('groesse','')} × {menge_str}"
            ))
            self._tbl_recent.setItem(r, 4, QTableWidgetItem(b.get("mitarbeiter_name", "")))

    def showEvent(self, event):
        super().showEvent(event)
        self._refresh()
