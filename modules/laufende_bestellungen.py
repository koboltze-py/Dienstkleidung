"""
DRK Dienstkleidung – Laufende Bestellungen
Übersicht offener Bestellungen, Abschluss bei Wareneingang,
sowie Historiensicht mit Filter nach Jahr/Monat.
"""

from datetime import datetime
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QFrame, QMessageBox, QDialog, QDialogButtonBox, QTextEdit,
    QFormLayout, QComboBox, QSplitter, QTabWidget, QLineEdit,
    QSizePolicy,
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QFont, QColor


# ---------------------------------------------------------------------------
# Detail-Dialog: Positionen einer Bestellung anzeigen
# ---------------------------------------------------------------------------

class BestellungDetailDialog(QDialog):
    def __init__(self, bestellung: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Bestellung {bestellung.get('bestellnummer', '')} – Details")
        self.setMinimumWidth(560)
        self.setMinimumHeight(400)
        self.setModal(True)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(10)

        # Meta
        meta_frame = QFrame()
        meta_frame.setObjectName("stat_card")
        meta_layout = QFormLayout(meta_frame)
        meta_layout.setSpacing(6)
        meta_layout.setContentsMargins(12, 10, 12, 10)
        meta_layout.addRow("Bestellnummer:", QLabel(str(bestellung.get("bestellnummer", ""))))
        meta_layout.addRow("Datum:", QLabel(str(bestellung.get("datum", ""))))
        status_lbl = QLabel(
            "✅ Abgeschlossen" if bestellung.get("status") == "abgeschlossen" else "🔄 Offen"
        )
        meta_layout.addRow("Status:", status_lbl)
        if bestellung.get("abgeschlossen_am"):
            meta_layout.addRow("Abgeschlossen am:", QLabel(str(bestellung["abgeschlossen_am"])))
        if bestellung.get("bemerkung"):
            meta_layout.addRow("Bemerkung:", QLabel(str(bestellung["bemerkung"])))
        layout.addWidget(meta_frame)

        # Positionen-Tabelle
        lbl = QLabel("Bestellpositionen")
        f = QFont(); f.setPointSize(10); f.setBold(True)
        lbl.setFont(f)
        layout.addWidget(lbl)

        tbl = QTableWidget(0, 4)
        tbl.setHorizontalHeaderLabels(["Kleidungsart", "Größe", "Menge", "Bemerkung"])
        tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        tbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        tbl.setAlternatingRowColors(True)

        positionen = bestellung.get("positionen", [])
        tbl.setRowCount(len(positionen))
        for r, pos in enumerate(positionen):
            tbl.setItem(r, 0, QTableWidgetItem(pos.get("art_name", "")))
            tbl.setItem(r, 1, QTableWidgetItem(str(pos.get("groesse", ""))))
            tbl.setItem(r, 2, QTableWidgetItem(str(pos.get("menge", ""))))
            tbl.setItem(r, 3, QTableWidgetItem(pos.get("bemerkung", "")))
        layout.addWidget(tbl, stretch=1)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


# ---------------------------------------------------------------------------
# Abschluss-Dialog
# ---------------------------------------------------------------------------

class AbschlussDialog(QDialog):
    def __init__(self, bestellung: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Bestellung abschließen")
        self.setMinimumWidth(420)
        self.setModal(True)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(10)

        lbl = QLabel(
            f"<b>Bestellung {bestellung.get('bestellnummer', '')} abschließen</b><br>"
            f"Datum: {bestellung.get('datum', '')}<br>"
            f"Positionen: {len(bestellung.get('positionen', []))}<br><br>"
            "Ware wurde vollständig geliefert und eingelagert?"
        )
        lbl.setWordWrap(True)
        lbl.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(lbl)

        form = QFormLayout()
        self.le_bemerkung = QLineEdit()
        self.le_bemerkung.setPlaceholderText("Optional: Bemerkung zum Abschluss ...")
        form.addRow("Bemerkung:", self.le_bemerkung)
        layout.addLayout(form)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("✅ Jetzt abschließen")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Abbrechen")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_bemerkung(self) -> str:
        return self.le_bemerkung.text().strip()


# ---------------------------------------------------------------------------
# Haupt-View
# ---------------------------------------------------------------------------

class LaufendeBestellungenView(QWidget):
    """Tab-Widget mit 'Laufende Bestellungen' und 'Historie'."""

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._setup_ui()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        # Titel
        lbl_title = QLabel("Laufende Bestellungen")
        lbl_title.setObjectName("page_title")
        lbl_sub = QLabel("Offene Bestellungen verwalten und Wareneingang bestätigen")
        lbl_sub.setObjectName("page_subtitle")
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_sub)

        # Tabs: Offen | Historie
        self._tabs = QTabWidget()
        self._tabs.setObjectName("inner_tabs")

        # --- Tab: Offen ---
        self._tab_offen = QWidget()
        self._build_offen_tab()
        self._tabs.addTab(self._tab_offen, "🔄  Offene Bestellungen")

        # --- Tab: Historie ---
        self._tab_historie = QWidget()
        self._build_historie_tab()
        self._tabs.addTab(self._tab_historie, "📋  Bestellhistorie")

        layout.addWidget(self._tabs, stretch=1)

        self._load_offen()
        self._load_historie()

    def _build_offen_tab(self):
        layout = QVBoxLayout(self._tab_offen)
        layout.setContentsMargins(0, 12, 0, 0)
        layout.setSpacing(10)

        # Toolbar
        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)

        btn_refresh = QPushButton("🔄  Aktualisieren")
        btn_refresh.setObjectName("btn_secondary")
        btn_refresh.clicked.connect(self._load_offen)
        toolbar.addWidget(btn_refresh)

        toolbar.addStretch()

        btn_abschliessen = QPushButton("✅  Bestellung abschließen")
        btn_abschliessen.setObjectName("btn_primary")
        btn_abschliessen.clicked.connect(self._abschliessen)
        toolbar.addWidget(btn_abschliessen)

        btn_detail = QPushButton("🔍  Details anzeigen")
        btn_detail.setObjectName("btn_secondary")
        btn_detail.clicked.connect(self._show_detail_offen)
        toolbar.addWidget(btn_detail)

        btn_loeschen = QPushButton("🗑  Löschen")
        btn_loeschen.setObjectName("btn_secondary")
        btn_loeschen.clicked.connect(self._loeschen_offen)
        toolbar.addWidget(btn_loeschen)

        layout.addLayout(toolbar)

        # Tabelle
        card = QFrame()
        card.setObjectName("stat_card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(6)

        self._tbl_offen = QTableWidget(0, 5)
        self._tbl_offen.setHorizontalHeaderLabels(
            ["Bestellnummer", "Datum", "Positionen", "Gesamt Stück", "Bemerkung"]
        )
        self._tbl_offen.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_offen.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_offen.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_offen.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_offen.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self._tbl_offen.verticalHeader().setVisible(False)
        self._tbl_offen.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._tbl_offen.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._tbl_offen.setAlternatingRowColors(True)
        self._tbl_offen.doubleClicked.connect(lambda _: self._show_detail_offen())
        card_layout.addWidget(self._tbl_offen)

        layout.addWidget(card, stretch=1)

        self._lbl_offen_info = QLabel("")
        self._lbl_offen_info.setObjectName("page_subtitle")
        layout.addWidget(self._lbl_offen_info)

    def _build_historie_tab(self):
        layout = QVBoxLayout(self._tab_historie)
        layout.setContentsMargins(0, 12, 0, 0)
        layout.setSpacing(10)

        # Filter-Zeile
        filter_bar = QHBoxLayout()
        filter_bar.setSpacing(10)

        filter_bar.addWidget(QLabel("Jahr:"))
        self._cb_jahr = QComboBox()
        self._cb_jahr.setMinimumWidth(90)
        now = datetime.now()
        for y in range(now.year, now.year - 6, -1):
            self._cb_jahr.addItem(str(y), y)
        self._cb_jahr.currentIndexChanged.connect(self._load_historie)
        filter_bar.addWidget(self._cb_jahr)

        filter_bar.addWidget(QLabel("Monat:"))
        self._cb_monat = QComboBox()
        self._cb_monat.setMinimumWidth(130)
        self._cb_monat.addItem("Alle Monate", 0)
        monate = [
            "Januar", "Februar", "März", "April", "Mai", "Juni",
            "Juli", "August", "September", "Oktober", "November", "Dezember"
        ]
        for i, name in enumerate(monate, 1):
            self._cb_monat.addItem(name, i)
        # Default: aktueller Monat
        self._cb_monat.setCurrentIndex(now.month)
        self._cb_monat.currentIndexChanged.connect(self._load_historie)
        filter_bar.addWidget(self._cb_monat)

        filter_bar.addWidget(QLabel("Status:"))
        self._cb_status_hist = QComboBox()
        self._cb_status_hist.addItem("Alle", "")
        self._cb_status_hist.addItem("🔄 Offen", "offen")
        self._cb_status_hist.addItem("✅ Abgeschlossen", "abgeschlossen")
        self._cb_status_hist.currentIndexChanged.connect(self._load_historie)
        filter_bar.addWidget(self._cb_status_hist)

        btn_refresh_hist = QPushButton("🔄  Aktualisieren")
        btn_refresh_hist.setObjectName("btn_secondary")
        btn_refresh_hist.clicked.connect(self._load_historie)
        filter_bar.addWidget(btn_refresh_hist)

        filter_bar.addStretch()

        btn_detail_hist = QPushButton("🔍  Details anzeigen")
        btn_detail_hist.setObjectName("btn_secondary")
        btn_detail_hist.clicked.connect(self._show_detail_hist)
        filter_bar.addWidget(btn_detail_hist)

        layout.addLayout(filter_bar)

        # Tabelle
        card = QFrame()
        card.setObjectName("stat_card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(12, 10, 12, 10)
        card_layout.setSpacing(6)

        self._tbl_hist = QTableWidget(0, 6)
        self._tbl_hist.setHorizontalHeaderLabels(
            ["Bestellnummer", "Datum", "Positionen", "Gesamt Stück", "Status", "Abgeschlossen am"]
        )
        self._tbl_hist.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_hist.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_hist.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_hist.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_hist.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_hist.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        self._tbl_hist.verticalHeader().setVisible(False)
        self._tbl_hist.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._tbl_hist.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._tbl_hist.setAlternatingRowColors(True)
        self._tbl_hist.doubleClicked.connect(lambda _: self._show_detail_hist())
        card_layout.addWidget(self._tbl_hist)

        layout.addWidget(card, stretch=1)

        self._lbl_hist_info = QLabel("")
        self._lbl_hist_info.setObjectName("page_subtitle")
        layout.addWidget(self._lbl_hist_info)

    # ------------------------------------------------------------------
    # Daten laden
    # ------------------------------------------------------------------

    def _load_offen(self):
        rows = self.db.get_laufende_bestellungen(status="offen")
        self._offen_data = rows
        self._tbl_offen.setRowCount(len(rows))
        for r, b in enumerate(rows):
            positionen = b.get("positionen", [])
            gesamt = sum(int(p.get("menge", 1)) for p in positionen)
            self._tbl_offen.setItem(r, 0, QTableWidgetItem(b.get("bestellnummer", "")))
            self._tbl_offen.setItem(r, 1, QTableWidgetItem(b.get("datum", "")))
            self._tbl_offen.setItem(r, 2, QTableWidgetItem(str(len(positionen))))
            self._tbl_offen.setItem(r, 3, QTableWidgetItem(str(gesamt)))
            self._tbl_offen.setItem(r, 4, QTableWidgetItem(b.get("bemerkung", "")))
            # Zeile einfärben
            for col in range(5):
                item = self._tbl_offen.item(r, col)
                if item:
                    item.setForeground(QColor("#1565c0"))
        self._lbl_offen_info.setText(
            f"{len(rows)} offene Bestellung(en)" if rows else "Keine offenen Bestellungen vorhanden."
        )

    def _load_historie(self):
        jahr = self._cb_jahr.currentData()
        monat = self._cb_monat.currentData()
        status = self._cb_status_hist.currentData() or None
        rows = self.db.get_laufende_bestellungen(
            status=status,
            jahr=jahr,
            monat=monat if monat else None,
        )
        self._hist_data = rows
        self._tbl_hist.setRowCount(len(rows))
        for r, b in enumerate(rows):
            positionen = b.get("positionen", [])
            gesamt = sum(int(p.get("menge", 1)) for p in positionen)
            status_txt = "✅ Abgeschlossen" if b.get("status") == "abgeschlossen" else "🔄 Offen"
            self._tbl_hist.setItem(r, 0, QTableWidgetItem(b.get("bestellnummer", "")))
            self._tbl_hist.setItem(r, 1, QTableWidgetItem(b.get("datum", "")))
            self._tbl_hist.setItem(r, 2, QTableWidgetItem(str(len(positionen))))
            self._tbl_hist.setItem(r, 3, QTableWidgetItem(str(gesamt)))
            self._tbl_hist.setItem(r, 4, QTableWidgetItem(status_txt))
            self._tbl_hist.setItem(r, 5, QTableWidgetItem(b.get("abgeschlossen_am", "") or ""))
            # Farbe je Status
            color = QColor("#2e7d32") if b.get("status") == "abgeschlossen" else QColor("#1565c0")
            for col in range(6):
                item = self._tbl_hist.item(r, col)
                if item:
                    item.setForeground(color)
        self._lbl_hist_info.setText(f"{len(rows)} Einträge")

    def reload(self):
        """Wird extern aufgerufen, um beide Tabs zu aktualisieren."""
        self._load_offen()
        self._load_historie()

    # ------------------------------------------------------------------
    # Aktionen
    # ------------------------------------------------------------------

    def _get_selected_offen(self) -> dict | None:
        rows = self._tbl_offen.selectedItems()
        if not rows:
            return None
        row = self._tbl_offen.currentRow()
        if 0 <= row < len(self._offen_data):
            return self._offen_data[row]
        return None

    def _get_selected_hist(self) -> dict | None:
        rows = self._tbl_hist.selectedItems()
        if not rows:
            return None
        row = self._tbl_hist.currentRow()
        if 0 <= row < len(self._hist_data):
            return self._hist_data[row]
        return None

    def _show_detail_offen(self):
        b = self._get_selected_offen()
        if not b:
            QMessageBox.information(self, "Hinweis", "Bitte eine Bestellung auswählen.")
            return
        dlg = BestellungDetailDialog(b, parent=self)
        dlg.exec()

    def _show_detail_hist(self):
        b = self._get_selected_hist()
        if not b:
            QMessageBox.information(self, "Hinweis", "Bitte eine Bestellung auswählen.")
            return
        dlg = BestellungDetailDialog(b, parent=self)
        dlg.exec()

    def _abschliessen(self):
        b = self._get_selected_offen()
        if not b:
            QMessageBox.information(self, "Hinweis", "Bitte eine offene Bestellung auswählen.")
            return
        dlg = AbschlussDialog(b, parent=self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        bemerkung = dlg.get_bemerkung()
        ok, msg = self.db.abschliessen_laufende_bestellung(b["id"], bemerkung)
        if ok:
            QMessageBox.information(self, "Abgeschlossen", msg)
            self._load_offen()
            self._load_historie()
        else:
            QMessageBox.critical(self, "Fehler", msg)

    def _loeschen_offen(self):
        b = self._get_selected_offen()
        if not b:
            QMessageBox.information(self, "Hinweis", "Bitte eine Bestellung auswählen.")
            return
        ans = QMessageBox.question(
            self, "Bestellung löschen",
            f"Bestellung {b.get('bestellnummer', '')} wirklich löschen?\n"
            "Diese Aktion kann nicht rückgängig gemacht werden.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if ans != QMessageBox.StandardButton.Yes:
            return
        ok, msg = self.db.delete_laufende_bestellung(b["id"])
        if ok:
            self._load_offen()
            self._load_historie()
        else:
            QMessageBox.critical(self, "Fehler", msg)
