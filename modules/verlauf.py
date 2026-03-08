"""
DRK Dienstkleidung - Buchungsverlauf
Vollständige Transaktionshistorie mit Filter und Export.
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QLineEdit, QDateEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QMessageBox, QFrame,
)
from PySide6.QtCore import Qt, QDate
from PySide6.QtGui import QColor

from utils import format_datum, export_table_to_csv

PAGE_SIZE = 100


class VerlaufView(QWidget):
    """Buchungsverlauf / Transaktionshistorie."""

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._offset = 0
        self._total = 0
        self._setup_ui()
        self._load_filter_combos()
        self._load_data()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        # Titel
        lbl_title = QLabel("Buchungsverlauf")
        lbl_title.setObjectName("page_title")
        lbl_sub = QLabel("Alle Ein-/Ausgaben und Korrekturen in chronologischer Reihenfolge")
        lbl_sub.setObjectName("page_subtitle")
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_sub)

        # Filter-Zeile 1
        f1 = QHBoxLayout()
        f1.setSpacing(10)

        f1.addWidget(QLabel("Kleidungsart:"))
        self.cb_art = QComboBox()
        self.cb_art.setMinimumWidth(170)
        f1.addWidget(self.cb_art)

        f1.addWidget(QLabel("Typ:"))
        self.cb_typ = QComboBox()
        self.cb_typ.addItems(["Alle", "ausgabe", "rueckgabe", "eingang", "korrektur"])
        self.cb_typ.setMinimumWidth(120)
        f1.addWidget(self.cb_typ)

        f1.addWidget(QLabel("Von:"))
        self.de_von = QDateEdit()
        self.de_von.setDisplayFormat("dd.MM.yyyy")
        self.de_von.setCalendarPopup(True)
        self.de_von.setDate(QDate.currentDate().addMonths(-3))
        self.de_von.setMaximumWidth(130)
        f1.addWidget(self.de_von)

        f1.addWidget(QLabel("Bis:"))
        self.de_bis = QDateEdit(QDate.currentDate())
        self.de_bis.setDisplayFormat("dd.MM.yyyy")
        self.de_bis.setCalendarPopup(True)
        self.de_bis.setMaximumWidth(130)
        f1.addWidget(self.de_bis)

        f1.addStretch()
        layout.addLayout(f1)

        # Filter-Zeile 2
        f2 = QHBoxLayout()
        f2.setSpacing(10)

        f2.addWidget(QLabel("Suche:"))
        self.le_suche = QLineEdit()
        self.le_suche.setPlaceholderText("Name, Kleidungsart oder Größe ...")
        self.le_suche.setMinimumWidth(250)
        f2.addWidget(self.le_suche)

        btn_suchen = QPushButton("🔍  Suchen")
        btn_suchen.setObjectName("btn_primary")
        btn_suchen.clicked.connect(self._search)
        f2.addWidget(btn_suchen)

        btn_reset = QPushButton("Zurücksetzen")
        btn_reset.setObjectName("btn_secondary")
        btn_reset.clicked.connect(self._reset_filter)
        f2.addWidget(btn_reset)

        f2.addStretch()

        btn_export = QPushButton("📄 CSV-Export")
        btn_export.setObjectName("btn_secondary")
        btn_export.clicked.connect(self._export)
        f2.addWidget(btn_export)

        layout.addLayout(f2)

        # Tabelle
        self._tbl = QTableWidget(0, 8)
        self._tbl.setHorizontalHeaderLabels([
            "Datum", "Typ", "Kleidungsart", "Größe", "Menge",
            "Mitarbeiter", "Ausgeg. von", "Bemerkung"
        ])
        self._tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeMode.Stretch)
        self._tbl.verticalHeader().setVisible(False)
        self._tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._tbl.setAlternatingRowColors(True)
        self._tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout.addWidget(self._tbl)

        # Paginierung
        pag_row = QHBoxLayout()
        self.lbl_count = QLabel("–")
        self.lbl_count.setObjectName("page_subtitle")
        pag_row.addWidget(self.lbl_count)
        pag_row.addStretch()

        self.btn_prev = QPushButton("◀ Zurück")
        self.btn_prev.setObjectName("btn_secondary")
        self.btn_prev.clicked.connect(self._prev_page)
        self.btn_prev.setEnabled(False)
        pag_row.addWidget(self.btn_prev)

        self.lbl_page = QLabel("Seite 1")
        self.lbl_page.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_page.setMinimumWidth(80)
        pag_row.addWidget(self.lbl_page)

        self.btn_next = QPushButton("Weiter ▶")
        self.btn_next.setObjectName("btn_secondary")
        self.btn_next.clicked.connect(self._next_page)
        pag_row.addWidget(self.btn_next)

        layout.addLayout(pag_row)

    def _load_filter_combos(self):
        self.cb_art.clear()
        self.cb_art.addItem("Alle Kleidungsarten", None)
        for art in self.db.get_kleidungsarten():
            self.cb_art.addItem(art["name"], art["id"])

    def _get_filter_args(self) -> dict:
        art_id = self.cb_art.currentData()
        typ_raw = self.cb_typ.currentText()
        typ = None if typ_raw == "Alle" else typ_raw
        datum_von = self.de_von.date().toString("yyyy-MM-dd")
        datum_bis = self.de_bis.date().toString("yyyy-MM-dd")
        suche = self.le_suche.text().strip() or None
        return dict(art_id=art_id, typ=typ, datum_von=datum_von, datum_bis=datum_bis, suche=suche)

    def _search(self):
        self._offset = 0
        self._load_data()

    def _reset_filter(self):
        self.cb_art.setCurrentIndex(0)
        self.cb_typ.setCurrentIndex(0)
        self.de_von.setDate(QDate.currentDate().addMonths(-3))
        self.de_bis.setDate(QDate.currentDate())
        self.le_suche.clear()
        self._offset = 0
        self._load_data()

    def _load_data(self):
        args = self._get_filter_args()
        self._total = self.db.get_buchungen_count(**args)
        buchungen = self.db.get_buchungen(limit=PAGE_SIZE, offset=self._offset, **args)
        self._fill_table(buchungen)
        self._update_pagination()

    def _fill_table(self, data: list[dict]):
        TYP_COLOR = {
            "ausgabe":   QColor("#FFF9C4"),
            "rueckgabe": QColor("#E8F5E9"),
            "eingang":   QColor("#E3F2FD"),
            "korrektur": QColor("#F3E5F5"),
        }
        TYP_LABEL = {
            "ausgabe":   "📤 Ausgabe",
            "rueckgabe": "📥 Rückgabe",
            "eingang":   "📦 Eingang",
            "korrektur": "✏  Korrektur",
        }
        self._tbl.setRowCount(len(data))
        for r, b in enumerate(data):
            typ = b.get("typ", "")
            menge = b.get("menge", 0)
            menge_str = f"+{menge}" if menge > 0 else str(menge)
            color = TYP_COLOR.get(typ, QColor("white"))

            cells = [
                format_datum(b.get("datum", "")),
                TYP_LABEL.get(typ, typ),
                b.get("art_name", ""),
                str(b.get("groesse", "")),
                menge_str,
                b.get("mitarbeiter_name", ""),
                b.get("ausgegeben_von", ""),
                b.get("bemerkung", ""),
            ]
            for c, text in enumerate(cells):
                it = QTableWidgetItem(text)
                it.setBackground(color)
                self._tbl.setItem(r, c, it)

    def _update_pagination(self):
        page = self._offset // PAGE_SIZE + 1
        total_pages = max(1, (self._total + PAGE_SIZE - 1) // PAGE_SIZE)
        self.lbl_page.setText(f"Seite {page} / {total_pages}")
        self.lbl_count.setText(f"{self._total} Buchung(en) gefunden")
        self.btn_prev.setEnabled(self._offset > 0)
        self.btn_next.setEnabled(self._offset + PAGE_SIZE < self._total)

    def _prev_page(self):
        self._offset = max(0, self._offset - PAGE_SIZE)
        self._load_data()

    def _next_page(self):
        self._offset += PAGE_SIZE
        self._load_data()

    def _export(self):
        args = self._get_filter_args()
        # Alle Datensätze ohne Limit
        all_data = self.db.get_buchungen(limit=100_000, offset=0, **args)
        headers = ["Datum", "Typ", "Kleidungsart", "Größe", "Menge",
                   "Mitarbeiter", "Ausgegeben von", "Bemerkung"]
        rows = [
            [
                format_datum(b.get("datum", "")),
                b.get("typ", ""),
                b.get("art_name", ""),
                str(b.get("groesse", "")),
                str(b.get("menge", "")),
                b.get("mitarbeiter_name", ""),
                b.get("ausgegeben_von", ""),
                b.get("bemerkung", ""),
            ]
            for b in all_data
        ]
        ok, msg = export_table_to_csv(headers, rows, "Buchungsverlauf")
        if ok:
            QMessageBox.information(self, "Export erfolgreich", f"Datei gespeichert:\n{msg}")
        else:
            QMessageBox.warning(self, "Export fehlgeschlagen", msg)

    def showEvent(self, event):
        super().showEvent(event)
        self._load_filter_combos()
        self._load_data()
