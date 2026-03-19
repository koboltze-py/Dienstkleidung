"""
DRK Dienstkleidung - Ausgabe & Rückgabe
Kleidung an Mitarbeiter ausgeben oder zurücknehmen.
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QLineEdit, QSpinBox, QDateEdit, QTextEdit,
    QTabWidget, QTableWidget, QTableWidgetItem, QGroupBox,
    QFormLayout, QMessageBox, QHeaderView, QAbstractItemView,
    QFrame, QScrollArea, QSizePolicy, QDialog, QDialogButtonBox,
    QSplitter, QRadioButton, QButtonGroup,
)
from PySide6.QtCore import Qt, QDate, QTimer

from utils import today_iso, format_datum
from modules.word_protokoll import (
    ProtokollAbfrageDialog, create_ausgabe_protokoll,
    create_rueckgabe_protokoll, open_document,
    get_ausgabe_dir, get_ruecknahme_dir,
)


class MaSidebar(QWidget):
    """Mitarbeiter-Seitenleiste mit Suche und Lazy-Loading (25 pro Seite)."""

    PAGE_SIZE = 25

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._all_items: list[dict] = []
        self._shown_count = self.PAGE_SIZE
        self._select_cb = None
        self._setup_ui()

    def set_select_callback(self, cb):
        self._select_cb = cb

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 4, 8)
        layout.setSpacing(6)

        top = QHBoxLayout()
        top.setSpacing(4)
        self._le_suche = QLineEdit()
        self._le_suche.setPlaceholderText("\U0001f50d  Name suchen ...")
        self._le_suche.textChanged.connect(self._on_search)
        top.addWidget(self._le_suche)

        btn_reload = QPushButton("\u21bb")
        btn_reload.setObjectName("btn_secondary")
        btn_reload.setFixedWidth(30)
        btn_reload.setToolTip("Aktualisieren")
        btn_reload.clicked.connect(self.load)
        top.addWidget(btn_reload)
        layout.addLayout(top)

        lbl = QLabel("Mitarbeiter")
        lbl.setStyleSheet("font-weight: bold; color: #555; font-size: 12px;")
        layout.addWidget(lbl)

        self._tbl = QTableWidget(0, 2)
        self._tbl.setHorizontalHeaderLabels(["Name", "\u2709"])
        self._tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.verticalHeader().setVisible(False)
        self._tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._tbl.setAlternatingRowColors(True)
        self._tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._tbl.setMinimumWidth(180)
        self._tbl.currentItemChanged.connect(self._on_row_clicked)
        layout.addWidget(self._tbl)

        self._btn_mehr = QPushButton("\u25bc  Weitere 25 laden")
        self._btn_mehr.setObjectName("btn_secondary")
        self._btn_mehr.clicked.connect(self._load_more)
        self._btn_mehr.setVisible(False)
        layout.addWidget(self._btn_mehr)

        self._lbl_count = QLabel("")
        self._lbl_count.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._lbl_count.setStyleSheet("color: #888; font-size: 11px;")
        layout.addWidget(self._lbl_count)

    def load(self):
        self._shown_count = self.PAGE_SIZE
        self._le_suche.blockSignals(True)
        self._le_suche.clear()
        self._le_suche.blockSignals(False)
        self._fetch_items()
        self._render()

    def _fetch_items(self):
        ma_mit_kleidung = {
            mk["mitarbeiter_id"]: mk
            for mk in self.db.get_mitarbeiter_mit_kleidung()
        }
        alle_ma = self.db.get_alle_mitarbeiter()
        added_ids: set = set()
        self._all_items = []
        for ma in alle_ma:
            mk_info = ma_mit_kleidung.get(ma["id"], {})
            self._all_items.append({
                "id": ma["id"],
                "name": f"{ma['nachname']}, {ma['vorname']}",
                "kleidung_anzahl": mk_info.get("anzahl_positionen", 0),
            })
            added_ids.add(ma["id"])
        for mk_id, mk_info in ma_mit_kleidung.items():
            if mk_id not in added_ids:
                self._all_items.append({
                    "id": mk_id,
                    "name": mk_info.get("mitarbeiter_name", f"ID {mk_id}"),
                    "kleidung_anzahl": mk_info.get("anzahl_positionen", 0),
                })
        self._all_items.sort(key=lambda x: x["name"])

    def _get_filtered(self) -> list[dict]:
        text = self._le_suche.text().lower()
        if text:
            return [e for e in self._all_items if text in e["name"].lower()]
        return self._all_items

    def _render(self):
        filtered = self._get_filtered()
        shown = filtered[:self._shown_count]
        self._tbl.blockSignals(True)
        self._tbl.setSortingEnabled(False)
        self._tbl.setRowCount(len(shown))
        for r, entry in enumerate(shown):
            name_cell = QTableWidgetItem(entry["name"])
            name_cell.setData(Qt.ItemDataRole.UserRole, entry)
            if entry["kleidung_anzahl"] > 0:
                name_cell.setForeground(Qt.GlobalColor.darkGreen)
            self._tbl.setItem(r, 0, name_cell)
            count_cell = QTableWidgetItem()
            count_cell.setText(str(entry["kleidung_anzahl"]) if entry["kleidung_anzahl"] > 0 else "\u2013")
            count_cell.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if entry["kleidung_anzahl"] > 0:
                count_cell.setForeground(Qt.GlobalColor.darkGreen)
            self._tbl.setItem(r, 1, count_cell)
        self._tbl.setSortingEnabled(True)
        self._tbl.blockSignals(False)
        total = len(filtered)
        show_more = total > self._shown_count
        self._btn_mehr.setVisible(show_more)
        if show_more:
            remaining = total - self._shown_count
            self._btn_mehr.setText(f"\u25bc  Weitere 25 laden ({remaining} verbleibend)")
        self._lbl_count.setText(f"{min(self._shown_count, total)} von {total}")

    def _on_search(self):
        self._shown_count = self.PAGE_SIZE
        self._render()

    def _load_more(self):
        self._shown_count += self.PAGE_SIZE
        self._render()

    def _on_row_clicked(self, current, _prev):
        if current is None or self._select_cb is None:
            return
        cell = self._tbl.item(current.row(), 0)
        if not cell:
            return
        entry = cell.data(Qt.ItemDataRole.UserRole)
        if entry:
            self._select_cb(entry)


class AusgabeTab(QWidget):
    """Reiter: Kleidung ausgeben – mehrere Positionen gleichzeitig."""

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._artikel_rows: list[dict] = []
        self._setup_ui()
        self._load_combos()
        self._sidebar.load()

    def _setup_ui(self):
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        scroll.setWidget(container)
        main_layout = QVBoxLayout(container)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(16)

        # --- Mitarbeiter & Meta ---
        grp_ma = QGroupBox("Mitarbeiter")
        form_ma = QFormLayout(grp_ma)
        form_ma.setSpacing(12)
        form_ma.setContentsMargins(16, 20, 16, 16)

        ma_row = QHBoxLayout()
        self.cb_ma = QComboBox()
        self.cb_ma.setMinimumWidth(300)
        self.cb_ma.setEditable(True)
        self.cb_ma.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.cb_ma.completer().setFilterMode(Qt.MatchFlag.MatchContains)
        self.cb_ma.completer().setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        ma_row.addWidget(self.cb_ma)
        self.le_ma_freitext = QLineEdit()
        self.le_ma_freitext.setPlaceholderText("Oder Freitext (falls nicht in Liste)")
        self.le_ma_freitext.setMinimumWidth(220)
        ma_row.addWidget(QLabel("oder:"))
        ma_row.addWidget(self.le_ma_freitext)
        form_ma.addRow("Mitarbeiter:", ma_row)

        # Bestehende Kleidung des gewählten MA
        self._grp_vorhanden = QGroupBox("Bereits zugeteilte Kleidung")
        self._grp_vorhanden.setVisible(False)
        vorhanden_layout = QVBoxLayout(self._grp_vorhanden)
        vorhanden_layout.setContentsMargins(12, 12, 12, 8)
        vorhanden_layout.setSpacing(4)
        self._tbl_vorhanden = QTableWidget(0, 4)
        self._tbl_vorhanden.setHorizontalHeaderLabels(["Kleidungsart", "Größe", "Anzahl", "Ausgabe-Datum"])
        self._tbl_vorhanden.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._tbl_vorhanden.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_vorhanden.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_vorhanden.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_vorhanden.verticalHeader().setVisible(False)
        self._tbl_vorhanden.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._tbl_vorhanden.setMaximumHeight(110)
        vorhanden_layout.addWidget(self._tbl_vorhanden)
        form_ma.addRow(self._grp_vorhanden)

        # Debounce: Datenbank nicht bei jedem Tastendruck abfragen
        self._vorhanden_timer = QTimer(self)
        self._vorhanden_timer.setSingleShot(True)
        self._vorhanden_timer.timeout.connect(self._update_vorhanden)
        self.cb_ma.currentIndexChanged.connect(lambda: self._vorhanden_timer.start(250))

        self.de_datum = QDateEdit(QDate.currentDate())
        self.de_datum.setDisplayFormat("dd.MM.yyyy")
        self.de_datum.setCalendarPopup(True)
        self.de_datum.setMaximumWidth(160)
        form_ma.addRow("Datum:", self.de_datum)

        self.le_von = QLineEdit()
        self.le_von.setPlaceholderText("Kürzel oder Name ...")
        self.le_von.setMaximumWidth(180)
        form_ma.addRow("Ausgegeben von:", self.le_von)

        self.le_bem = QLineEdit()
        self.le_bem.setPlaceholderText("Optional ...")
        form_ma.addRow("Bemerkung:", self.le_bem)

        main_layout.addWidget(grp_ma)

        # --- Kleidungspositionen ---
        grp_artikel = QGroupBox("Kleidungspositionen")
        grp_layout = QVBoxLayout(grp_artikel)
        grp_layout.setContentsMargins(16, 20, 16, 16)
        grp_layout.setSpacing(6)

        # Spaltenköpfe
        hdr = QHBoxLayout()
        hdr.setSpacing(8)
        for txt, w in [("Kleidungsart", 170), ("Größe", 190), ("Verfügbar", 75), ("Anzahl", 75)]:
            lbl = QLabel(txt)
            lbl.setFixedWidth(w)
            lbl.setStyleSheet("font-weight: bold; color: #555;")
            hdr.addWidget(lbl)
        hdr.addStretch()
        grp_layout.addLayout(hdr)

        # Container für dynamische Zeilen
        self._artikel_container = QVBoxLayout()
        self._artikel_container.setSpacing(4)
        grp_layout.addLayout(self._artikel_container)

        btn_add = QPushButton("+ Artikel hinzufügen")
        btn_add.setObjectName("btn_secondary")
        btn_add.clicked.connect(self._add_zeile)
        grp_layout.addWidget(btn_add)

        main_layout.addWidget(grp_artikel)

        # Aktions-Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        btn_ordner = QPushButton("Protokoll-Ordner oeffnen")
        btn_ordner.setObjectName("btn_secondary")
        btn_ordner.setToolTip("Ordner öffnen: Ausgabe Protokolle")
        def _open_ausgabe_ordner():
            import os as _os; d = get_ausgabe_dir(); _os.makedirs(d, exist_ok=True); _os.startfile(d)
        btn_ordner.clicked.connect(_open_ausgabe_ordner)
        btn_row.addWidget(btn_ordner)

        btn_reset = QPushButton("Zurücksetzen")
        btn_reset.setObjectName("btn_secondary")
        btn_reset.clicked.connect(self._reset_form)
        btn_row.addWidget(btn_reset)
        btn_submit = QPushButton("✔  Ausgabe speichern")
        btn_submit.setObjectName("btn_primary")
        btn_submit.clicked.connect(self._submit)
        btn_row.addWidget(btn_submit)
        main_layout.addLayout(btn_row)

        self.lbl_result = QLabel("")
        self.lbl_result.setVisible(False)
        self.lbl_result.setWordWrap(True)
        main_layout.addWidget(self.lbl_result)
        main_layout.addStretch()

        self._sidebar = MaSidebar(self.db)
        self._sidebar.set_select_callback(self._on_ma_selected_sidebar)
        self._sidebar.setMinimumWidth(180)
        self._sidebar.setMaximumWidth(280)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self._sidebar)
        splitter.addWidget(scroll)
        splitter.setSizes([230, 800])
        outer.addWidget(splitter)

    def _on_ma_selected_sidebar(self, entry: dict):
        idx = self.cb_ma.findData(entry["id"])
        if idx >= 0:
            self.cb_ma.blockSignals(True)
            self.cb_ma.setCurrentIndex(idx)
            self.cb_ma.lineEdit().setText(self.cb_ma.itemText(idx))
            self.cb_ma.blockSignals(False)
            # Completer-Popup schließen, damit er den Text nicht überschreibt
            popup = self.cb_ma.completer().popup()
            if popup and popup.isVisible():
                popup.hide()
            self._update_vorhanden()
        else:
            self.cb_ma.setCurrentIndex(0)
            self.le_ma_freitext.setText(entry["name"])

    def _load_combos(self):
        self.cb_ma.blockSignals(True)
        self.cb_ma.clear()
        self.cb_ma.addItem("– Aus Liste wählen –", None)
        for ma in self.db.get_alle_mitarbeiter():
            display = f"{ma['nachname']}, {ma['vorname']}"
            if ma.get("personalnummer"):
                display += f" ({ma['personalnummer']})"
            self.cb_ma.addItem(display, ma["id"])
        self.cb_ma.blockSignals(False)
        # Bestände in vorhandenen Zeilen aktualisieren
        for ri in self._artikel_rows:
            self._update_groessen_row(ri)
        if not self._artikel_rows:
            self._add_zeile()
        self._update_vorhanden()

    def _update_vorhanden(self):
        ma_id = self.cb_ma.currentData()
        if not ma_id:
            self._grp_vorhanden.setVisible(False)
            return
        items = self.db.get_mitarbeiter_kleidung(mitarbeiter_id=ma_id, status="ausgegeben")
        if not items:
            self._grp_vorhanden.setVisible(False)
            return
        self._grp_vorhanden.setVisible(True)
        self._tbl_vorhanden.setRowCount(len(items))
        for r, it in enumerate(items):
            self._tbl_vorhanden.setItem(r, 0, QTableWidgetItem(it["art_name"]))
            self._tbl_vorhanden.setItem(r, 1, QTableWidgetItem(str(it["groesse"])))
            self._tbl_vorhanden.setItem(r, 2, QTableWidgetItem(str(it["menge"])))
            self._tbl_vorhanden.setItem(r, 3, QTableWidgetItem(format_datum(it["ausgabe_datum"])))
            for c in range(4):
                cell = self._tbl_vorhanden.item(r, c)
                if cell:
                    cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)

    def _add_zeile(self):
        row_w = QWidget()
        hl = QHBoxLayout(row_w)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(8)

        cb_art = QComboBox()
        cb_art.setFixedWidth(170)
        for art in self.db.get_kleidungsarten():
            cb_art.addItem(art["name"], art["id"])

        cb_gr = QComboBox()
        cb_gr.setFixedWidth(190)

        lbl_v = QLabel("–")
        lbl_v.setFixedWidth(75)
        lbl_v.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_v.setStyleSheet("color: #555; font-weight: bold;")

        sb = QSpinBox()
        sb.setRange(1, 999)
        sb.setValue(1)
        sb.setFixedWidth(75)

        btn_x = QPushButton("×")
        btn_x.setFixedSize(28, 28)
        btn_x.setStyleSheet("QPushButton{color:#B20000;font-weight:bold;font-size:16px;border:none;background:transparent;}"
                            "QPushButton:hover{background:#fce4e4;border-radius:4px;}")

        hl.addWidget(cb_art)
        hl.addWidget(cb_gr)
        hl.addWidget(lbl_v)
        hl.addWidget(sb)
        hl.addWidget(btn_x)
        hl.addStretch()

        ri = {"widget": row_w, "cb_art": cb_art, "cb_gr": cb_gr, "lbl_v": lbl_v, "sb": sb}

        cb_art.currentIndexChanged.connect(lambda _, r=ri: self._update_groessen_row(r))
        cb_gr.currentIndexChanged.connect(lambda _, r=ri: self._update_verf_row(r))
        btn_x.clicked.connect(lambda _, rw=row_w, r=ri: self._remove_zeile(rw, r))

        self._artikel_rows.append(ri)
        self._artikel_container.addWidget(row_w)
        self._update_groessen_row(ri)

    def _remove_zeile(self, row_w, ri):
        if len(self._artikel_rows) <= 1:
            return
        self._artikel_rows.remove(ri)
        self._artikel_container.removeWidget(row_w)
        row_w.deleteLater()

    def _update_groessen_row(self, ri):
        art_id = ri["cb_art"].currentData()
        ri["cb_gr"].blockSignals(True)
        ri["cb_gr"].clear()
        if art_id:
            for g in self.db.get_groessen_fuer_art(art_id):
                ri["cb_gr"].addItem(f"{g['groesse']}  ({g['menge']} Stk.)", g["groesse"])
        ri["cb_gr"].blockSignals(False)
        self._update_verf_row(ri)

    def _update_verf_row(self, ri):
        art_id = ri["cb_art"].currentData()
        groesse = ri["cb_gr"].currentData()
        if not art_id or not groesse:
            ri["lbl_v"].setText("–")
            ri["sb"].setMaximum(999)
            return
        item = self.db.get_bestand_item(art_id, groesse)
        if item:
            m = int(item["menge"])
            ri["lbl_v"].setText(str(m))
            ri["sb"].setMaximum(max(1, m))
        else:
            ri["lbl_v"].setText("0")
            ri["sb"].setMaximum(0)

    def _reset_form(self):
        self.cb_ma.setCurrentIndex(0)
        self.le_ma_freitext.clear()
        self.de_datum.setDate(QDate.currentDate())
        self.le_von.clear()
        self.le_bem.clear()
        self.lbl_result.setVisible(False)
        for ri in list(self._artikel_rows):
            ri["widget"].deleteLater()
        self._artikel_rows.clear()
        self._add_zeile()

    def _submit(self):
        ma_id = self.cb_ma.currentData()
        ma_name = ""
        if ma_id:
            ma_name = self.cb_ma.currentText().split(" (")[0].strip()
        elif self.le_ma_freitext.text().strip():
            ma_name = self.le_ma_freitext.text().strip()
        else:
            QMessageBox.warning(self, "Pflichtfeld", "Bitte einen Mitarbeiter auswählen oder eingeben.")
            return

        datum = self.de_datum.date().toString("yyyy-MM-dd")
        ausgegeben_von = self.le_von.text().strip()
        bemerkung = self.le_bem.text().strip()

        # Artikel sammeln (ohne DB-Schreibzugriff)
        artikel = []
        for ri in self._artikel_rows:
            art_id = ri["cb_art"].currentData()
            art_name = ri["cb_art"].currentText()
            groesse = ri["cb_gr"].currentData()
            if not groesse:
                continue
            artikel.append({
                "art_id": art_id,
                "art_name": art_name,
                "groesse": groesse,
                "menge": ri["sb"].value(),
            })

        if not artikel:
            QMessageBox.warning(self, "Nichts gespeichert", "Bitte mindestens eine Kleidungsposition vollständig ausfüllen.")
            return

        # Protokoll-Popup
        zeilen = [
            f"<b>Mitarbeiter:</b> {ma_name}",
            f"<b>Datum:</b> {self.de_datum.date().toString('dd.MM.yyyy')}",
        ] + [f"&nbsp;• {a['art_name']}  {a['groesse']}  × {a['menge']}" for a in artikel]
        dlg_prot = ProtokollAbfrageDialog("Ausgabe", zeilen, self)
        if dlg_prot.exec() == QDialog.DialogCode.Rejected:
            return

        if dlg_prot.result_action == ProtokollAbfrageDialog.PROTOKOLL_AND_SAVE:
            ok_p, path = create_ausgabe_protokoll(
                ma_name, datum, ausgegeben_von, bemerkung, artikel
            )
            if ok_p:
                open_document(path)
            else:
                QMessageBox.warning(self, "Protokoll-Fehler", path)

        # In Datenbank speichern
        saved, errors = 0, []
        for art in artikel:
            ok, msg = self.db.ausgabe_kleidung(
                ma_id, ma_name, art["art_id"], art["art_name"], art["groesse"],
                art["menge"], datum, ausgegeben_von, bemerkung
            )
            if ok:
                saved += 1
            else:
                errors.append(f"{art['art_name']} {art['groesse']}: {msg}")

        if saved:
            self._reset_form()
            self._load_combos()
        self.lbl_result.setVisible(True)
        if saved and not errors:
            self.lbl_result.setText(f"✔  {saved} Position(en) erfolgreich ausgegeben.")
            self.lbl_result.setStyleSheet("background:#E8F5E9; border-left:4px solid #388E3C; padding:8px; border-radius:4px;")
        if errors:
            prefix = f"✔  {saved} gespeichert  \n" if saved else ""
            self.lbl_result.setText(prefix + "⚠  " + "\n".join(errors))
            self.lbl_result.setStyleSheet("background:#FFF3E0; border-left:4px solid #F57C00; padding:8px; border-radius:4px;")

    def showEvent(self, event):
        super().showEvent(event)
        self._load_combos()
        self._sidebar.load()


class RueckgabeDialog(QDialog):
    """Popup: Mengenaufteilung Lager / Entsorgung pro zurückgenommenem Artikel."""

    def __init__(self, items: list[dict], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Rückgabe buchen")
        self.setMinimumWidth(660)
        self.setModal(True)
        self._rows: list[dict] = []
        self._setup_ui(items)

    def _setup_ui(self, items):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(12)

        lbl = QLabel(
            "Für jeden Artikel festlegen, wie viele Stücke zurück ins Lager kommen "
            "und wie viele entsorgt werden."
        )
        lbl.setWordWrap(True)
        layout.addWidget(lbl)

        # Spaltenköpfe
        hdr_w = QWidget()
        hdr = QHBoxLayout(hdr_w)
        hdr.setContentsMargins(0, 0, 0, 0)
        hdr.setSpacing(8)
        for txt, w in [("Kleidungsart", 155), (" Größe", 85), ("Gesamt", 70), ("→ Zurück ins Lager", 115), ("→ Entsorgung", 115)]:
            lb = QLabel(txt)
            lb.setFixedWidth(w)
            lb.setStyleSheet("font-weight:bold; color:#444;")
            hdr.addWidget(lb)
        hdr.addStretch()
        layout.addWidget(hdr_w)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#ddd;")
        layout.addWidget(sep)

        # Artikel-Zeilen in ScrollArea
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setMaximumHeight(280)
        container = QWidget()
        rows_layout = QVBoxLayout(container)
        rows_layout.setContentsMargins(0, 4, 0, 4)
        rows_layout.setSpacing(6)
        scroll.setWidget(container)
        layout.addWidget(scroll)

        for item in items:
            menge = max(1, int(item.get("menge", 1)))
            row_w = QWidget()
            hl = QHBoxLayout(row_w)
            hl.setContentsMargins(0, 0, 0, 0)
            hl.setSpacing(8)

            for txt, w in [
                (item.get("art_name", ""), 155),
                (str(item.get("groesse", "")), 85),
                (f"{menge} Stk.", 70),
            ]:
                lb = QLabel(txt)
                lb.setFixedWidth(w)
                hl.addWidget(lb)

            sb_lager = QSpinBox()
            sb_lager.setRange(0, menge)
            sb_lager.setValue(menge)
            sb_lager.setFixedWidth(115)
            sb_lager.setSuffix(" Stk.")

            sb_entsorgt = QSpinBox()
            sb_entsorgt.setRange(0, 0)
            sb_entsorgt.setValue(0)
            sb_entsorgt.setFixedWidth(115)
            sb_entsorgt.setSuffix(" Stk.")

            def make_lager_cb(sl, se, m):
                def cb(val):
                    se.blockSignals(True)
                    rest = m - val
                    se.setMaximum(rest)
                    se.setValue(rest)
                    se.blockSignals(False)
                return cb

            def make_entsorgt_cb(sl, se, m):
                def cb(val):
                    sl.blockSignals(True)
                    rest = m - val
                    sl.setMaximum(rest)
                    sl.setValue(rest)
                    sl.blockSignals(False)
                return cb

            sb_lager.valueChanged.connect(make_lager_cb(sb_lager, sb_entsorgt, menge))
            sb_entsorgt.valueChanged.connect(make_entsorgt_cb(sb_lager, sb_entsorgt, menge))

            hl.addWidget(sb_lager)
            hl.addWidget(sb_entsorgt)
            hl.addStretch()

            rows_layout.addWidget(row_w)
            self._rows.append({
                "mk_id": item["id"],
                "sb_lager": sb_lager,
                "sb_entsorgt": sb_entsorgt,
            })

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.HLine)
        sep2.setStyleSheet("color:#ddd;")
        layout.addWidget(sep2)

        meta = QFormLayout()
        meta.setSpacing(8)
        self.de_datum = QDateEdit(QDate.currentDate())
        self.de_datum.setDisplayFormat("dd.MM.yyyy")
        self.de_datum.setCalendarPopup(True)
        self.de_datum.setMaximumWidth(140)
        meta.addRow("Rückgabedatum:", self.de_datum)

        self.le_bem = QLineEdit()
        self.le_bem.setPlaceholderText("Optional ...")
        meta.addRow("Bemerkung:", self.le_bem)
        layout.addLayout(meta)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.button(QDialogButtonBox.StandardButton.Ok).setText("✔  Rückgabe buchen")
        btns.button(QDialogButtonBox.StandardButton.Cancel).setText("Abbrechen")
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_result(self) -> tuple[list[dict], str, str]:
        return (
            [
                {"mk_id": r["mk_id"], "lager": r["sb_lager"].value(), "entsorgt": r["sb_entsorgt"].value()}
                for r in self._rows
            ],
            self.de_datum.date().toString("yyyy-MM-dd"),
            self.le_bem.text().strip(),
        )


class RueckgabeTab(QWidget):
    """Reiter: Kleidung zurücknehmen."""

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._setup_ui()
        self._load_mitarbeiter()
        self._sidebar.load()

    def _setup_ui(self):
        # Sidebar
        self._sidebar = MaSidebar(self.db)
        self._sidebar.set_select_callback(self._on_ma_selected_sidebar)
        self._sidebar.setMinimumWidth(180)
        self._sidebar.setMaximumWidth(280)

        # Rechtes Panel
        right = QWidget()
        layout = QVBoxLayout(right)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        # Mitarbeiter-Auswahl
        select_row = QHBoxLayout()
        select_row.addWidget(QLabel("Mitarbeiter:"))

        self.cb_ma_rueck = QComboBox()
        self.cb_ma_rueck.setMinimumWidth(300)
        self.cb_ma_rueck.setEditable(True)
        self.cb_ma_rueck.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.cb_ma_rueck.completer().setFilterMode(Qt.MatchFlag.MatchContains)
        self.cb_ma_rueck.completer().setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        # Beim Anklicken Text sofort überschreiben: selectAll via Timer nach Focus
        _le = self.cb_ma_rueck.lineEdit()
        _orig_focus = _le.focusInEvent
        def _on_focus(event, le=_le, orig=_orig_focus):
            orig(event)
            QTimer.singleShot(0, le.selectAll)
        _le.focusInEvent = _on_focus
        select_row.addWidget(self.cb_ma_rueck)

        btn_laden = QPushButton("Kleidung laden")
        btn_laden.setObjectName("btn_secondary")
        btn_laden.clicked.connect(self._load_kleidung)
        select_row.addWidget(btn_laden)
        select_row.addStretch()
        layout.addLayout(select_row)

        # Tabelle mit aktuell ausgegebener Kleidung
        lbl_tbl = QLabel("Aktuell ausgegebene Kleidung:")
        lbl_tbl.setStyleSheet("font-weight: bold; margin-top: 8px;")
        layout.addWidget(lbl_tbl)

        self._tbl = QTableWidget(0, 6)
        self._tbl.setHorizontalHeaderLabels(
            ["Kleidungsart", "Größe", "Anzahl", "Ausgabe-Datum", "Ausgegeben von", "Bemerkung"]
        )
        self._tbl.horizontalHeader().setStretchLastSection(True)
        self._tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl.verticalHeader().setVisible(False)
        self._tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._tbl.setAlternatingRowColors(True)
        layout.addWidget(self._tbl)

        # Aktions-Button
        bottom_row = QHBoxLayout()
        bottom_row.setSpacing(16)

        btn_ordner_r = QPushButton("Protokoll-Ordner oeffnen")
        btn_ordner_r.setObjectName("btn_secondary")
        btn_ordner_r.setToolTip("Ordner öffnen: Rücknahme Protokolle")
        def _open_rueck_ordner():
            import os as _os; d = get_ruecknahme_dir(); _os.makedirs(d, exist_ok=True); _os.startfile(d)
        btn_ordner_r.clicked.connect(_open_rueck_ordner)
        bottom_row.addWidget(btn_ordner_r)

        bottom_row.addStretch()

        btn_rueck = QPushButton("📥  Ausgewählte zurücknehmen")
        btn_rueck.setObjectName("btn_primary")
        btn_rueck.clicked.connect(self._submit_rueckgabe)
        bottom_row.addWidget(btn_rueck)

        layout.addLayout(bottom_row)

        self.lbl_result = QLabel("")
        self.lbl_result.setVisible(False)
        self.lbl_result.setWordWrap(True)
        layout.addWidget(self.lbl_result)

        # Gesamt-Layout als Splitter
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self._sidebar)
        splitter.addWidget(right)
        splitter.setSizes([230, 800])
        outer.addWidget(splitter)

    def _on_ma_selected_sidebar(self, entry: dict):
        idx = self.cb_ma_rueck.findData(entry["id"])
        if idx >= 0:
            self.cb_ma_rueck.blockSignals(True)
            self.cb_ma_rueck.setCurrentIndex(idx)
            self.cb_ma_rueck.lineEdit().setText(self.cb_ma_rueck.itemText(idx))
            self.cb_ma_rueck.blockSignals(False)
            popup = self.cb_ma_rueck.completer().popup()
            if popup and popup.isVisible():
                popup.hide()
        self._load_kleidung()

    def showEvent(self, event):
        super().showEvent(event)
        self._load_mitarbeiter()
        self._sidebar.load()

    def _load_mitarbeiter(self):
        self.cb_ma_rueck.clear()
        self.cb_ma_rueck.addItem("– Mitarbeiter suchen –", None)
        for ma in self.db.get_alle_mitarbeiter():
            display = f"{ma['nachname']}, {ma['vorname']}"
            self.cb_ma_rueck.addItem(display, ma["id"])
        # Auch Mitarbeiter aus mk-Tabelle (ohne ID)
        mit_kleidung = self.db.get_mitarbeiter_mit_kleidung()
        in_combo = {self.cb_ma_rueck.itemData(i) for i in range(self.cb_ma_rueck.count())}
        for mk in mit_kleidung:
            mk_id = mk.get("mitarbeiter_id")
            if mk_id not in in_combo:
                self.cb_ma_rueck.addItem(f"{mk['mitarbeiter_name']} (historisch)", mk_id)

    def _load_kleidung(self):
        self._tbl.setRowCount(0)
        self.lbl_result.setVisible(False)

        ma_id = self.cb_ma_rueck.currentData()
        ma_text = self.cb_ma_rueck.currentText()

        if ma_id:
            items = self.db.get_mitarbeiter_kleidung(mitarbeiter_id=ma_id, status="ausgegeben")
        elif ma_text and ma_text not in ("– Mitarbeiter suchen –",):
            # Suche nach Name
            name_part = ma_text.split(" (")[0].strip()
            items = self.db.get_mitarbeiter_kleidung(mitarbeiter_name=name_part, status="ausgegeben")
        else:
            QMessageBox.information(self, "Hinweis", "Bitte zuerst einen Mitarbeiter auswählen.")
            return

        if not items:
            self._tbl.setRowCount(0)
            return

        self._tbl.setRowCount(len(items))
        for r, it in enumerate(items):
            self._tbl.setItem(r, 0, QTableWidgetItem(it["art_name"]))
            self._tbl.setItem(r, 1, QTableWidgetItem(str(it["groesse"])))
            self._tbl.setItem(r, 2, QTableWidgetItem(str(it["menge"])))
            self._tbl.setItem(r, 3, QTableWidgetItem(format_datum(it["ausgabe_datum"])))
            self._tbl.setItem(r, 4, QTableWidgetItem(it.get("ausgegeben_von", "")))
            self._tbl.setItem(r, 5, QTableWidgetItem(it.get("bemerkung", "")))
            for c in range(6):
                cell = self._tbl.item(r, c)
                if cell:
                    cell.setData(Qt.ItemDataRole.UserRole, it["id"])   # mk_id speichern

    def _submit_rueckgabe(self):
        selected = self._tbl.selectedItems()
        if not selected:
            QMessageBox.warning(self, "Auswahl fehlt", "Bitte mindestens eine Zeile auswählen.")
            return

        rows_selected = sorted({self._tbl.row(it) for it in selected})
        items = []
        for row in rows_selected:
            cell = self._tbl.item(row, 0)
            if not cell:
                continue
            mk_id = cell.data(Qt.ItemDataRole.UserRole)
            if mk_id is None:
                continue
            try:
                menge = int(self._tbl.item(row, 2).text()) if self._tbl.item(row, 2) else 1
            except (ValueError, AttributeError):
                menge = 1
            items.append({
                "id": mk_id,
                "art_name": self._tbl.item(row, 0).text() if self._tbl.item(row, 0) else "",
                "groesse": self._tbl.item(row, 1).text() if self._tbl.item(row, 1) else "",
                "menge": menge,
            })

        if not items:
            return

        dlg = RueckgabeDialog(items, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        result, datum, bemerkung = dlg.get_result()

        # MA Name für Protokoll
        ma_name = self.cb_ma_rueck.currentText().split(" (")[0].strip()
        if ma_name == "– Mitarbeiter suchen –":
            ma_name = ""

        # Artikeldaten anreichern und Protokoll-Popup zeigen
        id_to_item = {it["id"]: it for it in items}
        proto_rows = []
        for r in result:
            item = id_to_item.get(r["mk_id"], {})
            proto_rows.append({
                "art_name": item.get("art_name", ""),
                "groesse":  item.get("groesse", ""),
                "menge":    item.get("menge", r["lager"] + r["entsorgt"]),
                "lager":    r["lager"],
                "entsorgt": r["entsorgt"],
            })

        datum_de = QDate.fromString(datum, "yyyy-MM-dd").toString("dd.MM.yyyy")
        zeilen = [
            f"<b>Mitarbeiter:</b> {ma_name}",
            f"<b>Datum:</b> {datum_de}",
        ] + [
            f"&nbsp;• {p['art_name']}  {p['groesse']}  ×{p['menge']}  "
            f"(Lager: {p['lager']}, Entsorgung: {p['entsorgt']})"
            for p in proto_rows
        ]
        dlg_prot = ProtokollAbfrageDialog("Rückgabe", zeilen, self)
        if dlg_prot.exec() == QDialog.DialogCode.Rejected:
            return

        if dlg_prot.result_action == ProtokollAbfrageDialog.PROTOKOLL_AND_SAVE:
            ok_p, path = create_rueckgabe_protokoll(
                ma_name, datum, bemerkung, proto_rows
            )
            if ok_p:
                open_document(path)
            else:
                QMessageBox.warning(self, "Protokoll-Fehler", path)

        saved, errors = 0, []
        for r in result:
            ok, msg = self.db.rueckgabe_mit_entsorgung(
                r["mk_id"], r["lager"], r["entsorgt"], datum, bemerkung
            )
            if ok:
                saved += 1
            else:
                errors.append(msg)

        self.lbl_result.setVisible(True)
        if saved:
            self._load_kleidung()
            self.lbl_result.setText(f"✔  {saved} Position(en) erfolgreich zurückgenommen.")
            self.lbl_result.setStyleSheet("background:#E8F5E9; border-left:4px solid #388E3C; padding:8px; border-radius:4px;")
        if errors:
            self.lbl_result.setText(self.lbl_result.text() + "\n⚠  " + "\n".join(errors))
            self.lbl_result.setStyleSheet("background:#FFF3E0; border-left:4px solid #F57C00; padding:8px; border-radius:4px;")

    def showEvent(self, event):
        super().showEvent(event)
        self._load_mitarbeiter()


class NachtraeglichesProtokollTab(QWidget):
    """Reiter: Nachträgliches Ausgabe- oder Rückgabeprotokoll erstellen (ohne DB-Buchung)."""

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._artikel_rows: list[dict] = []
        self._setup_ui()
        self._load_combos()
        self._sidebar.load()

    def _setup_ui(self):
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        scroll.setWidget(container)
        main_layout = QVBoxLayout(container)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(16)

        # Protokolltyp
        grp_typ = QGroupBox("Protokolltyp")
        typ_layout = QHBoxLayout(grp_typ)
        typ_layout.setContentsMargins(16, 12, 16, 12)
        self._rb_ausgabe = QRadioButton("📤  Ausgabeprotokoll")
        self._rb_rueckgabe = QRadioButton("📥  Rückgabeprotokoll")
        self._rb_ausgabe.setChecked(True)
        self._btn_grp = QButtonGroup(self)
        self._btn_grp.addButton(self._rb_ausgabe)
        self._btn_grp.addButton(self._rb_rueckgabe)
        typ_layout.addWidget(self._rb_ausgabe)
        typ_layout.addWidget(self._rb_rueckgabe)
        typ_layout.addStretch()
        main_layout.addWidget(grp_typ)

        # Mitarbeiter & Meta
        grp_ma = QGroupBox("Mitarbeiter & Angaben")
        form_ma = QFormLayout(grp_ma)
        form_ma.setSpacing(12)
        form_ma.setContentsMargins(16, 20, 16, 16)

        ma_row = QHBoxLayout()
        self.cb_ma = QComboBox()
        self.cb_ma.setMinimumWidth(300)
        self.cb_ma.setEditable(True)
        self.cb_ma.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.cb_ma.completer().setFilterMode(Qt.MatchFlag.MatchContains)
        self.cb_ma.completer().setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        ma_row.addWidget(self.cb_ma)
        self.le_ma_freitext = QLineEdit()
        self.le_ma_freitext.setPlaceholderText("Oder Freitext (falls nicht in Liste)")
        self.le_ma_freitext.setMinimumWidth(220)
        ma_row.addWidget(QLabel("oder:"))
        ma_row.addWidget(self.le_ma_freitext)
        form_ma.addRow("Mitarbeiter:", ma_row)

        self._gs_timer = QTimer(self)
        self._gs_timer.setSingleShot(True)
        self._gs_timer.timeout.connect(self._load_gespeicherte_kleidung)
        self.cb_ma.currentIndexChanged.connect(lambda: self._gs_timer.start(250))

        self.de_datum = QDateEdit(QDate.currentDate())
        self.de_datum.setDisplayFormat("dd.MM.yyyy")
        self.de_datum.setCalendarPopup(True)
        self.de_datum.setMaximumWidth(160)
        form_ma.addRow("Datum:", self.de_datum)

        self.le_von = QLineEdit()
        self.le_von.setPlaceholderText("Kürzel oder Name ...")
        self.le_von.setMaximumWidth(180)
        form_ma.addRow("Ausgegeben von:", self.le_von)

        self.le_bem = QLineEdit()
        self.le_bem.setPlaceholderText("Optional ...")
        form_ma.addRow("Bemerkung:", self.le_bem)
        main_layout.addWidget(grp_ma)

        # Gespeicherte Kleidung des Mitarbeiters
        self._grp_gespeichert = QGroupBox(
            "Gespeicherte Kleidung des Mitarbeiters  (✔ Häkchen setzen → ins Protokoll übernehmen)"
        )
        gs_layout = QVBoxLayout(self._grp_gespeichert)
        gs_layout.setContentsMargins(16, 12, 16, 12)
        gs_layout.setSpacing(6)

        self._tbl_gespeichert = QTableWidget(0, 5)
        self._tbl_gespeichert.setHorizontalHeaderLabels(["✔", "Kleidungsart", "Größe", "Anzahl", "Ausgabe-Datum"])
        self._tbl_gespeichert.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_gespeichert.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self._tbl_gespeichert.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_gespeichert.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_gespeichert.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        self._tbl_gespeichert.verticalHeader().setVisible(False)
        self._tbl_gespeichert.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._tbl_gespeichert.setAlternatingRowColors(True)
        self._tbl_gespeichert.setMaximumHeight(190)
        gs_layout.addWidget(self._tbl_gespeichert)

        gs_btn_row = QHBoxLayout()
        btn_alle_an = QPushButton("☑  Alle auswählen")
        btn_alle_an.setObjectName("btn_secondary")
        btn_alle_an.clicked.connect(lambda: self._set_all_checks(True))
        gs_btn_row.addWidget(btn_alle_an)
        btn_alle_ab = QPushButton("☐  Alle abwählen")
        btn_alle_ab.setObjectName("btn_secondary")
        btn_alle_ab.clicked.connect(lambda: self._set_all_checks(False))
        gs_btn_row.addWidget(btn_alle_ab)
        gs_btn_row.addStretch()
        gs_layout.addLayout(gs_btn_row)

        self._lbl_gs_hint = QLabel("← Mitarbeiter auswählen, um gespeicherte Kleidung zu laden")
        self._lbl_gs_hint.setStyleSheet("color:#888; font-size:11px;")
        gs_layout.addWidget(self._lbl_gs_hint)
        main_layout.addWidget(self._grp_gespeichert)

        # Kleidungspositionen (manuell ergänzen)
        grp_artikel = QGroupBox("Kleidungspositionen (manuell ergänzen)")
        grp_layout = QVBoxLayout(grp_artikel)
        grp_layout.setContentsMargins(16, 20, 16, 16)
        grp_layout.setSpacing(6)

        hdr = QHBoxLayout()
        hdr.setSpacing(8)
        for txt, w in [("Kleidungsart", 200), ("Größe", 130), ("Anzahl", 80)]:
            lbl = QLabel(txt)
            lbl.setFixedWidth(w)
            lbl.setStyleSheet("font-weight: bold; color: #555;")
            hdr.addWidget(lbl)
        hdr.addStretch()
        grp_layout.addLayout(hdr)

        self._artikel_container = QVBoxLayout()
        self._artikel_container.setSpacing(4)
        grp_layout.addLayout(self._artikel_container)

        btn_add = QPushButton("+ Artikel hinzufügen")
        btn_add.setObjectName("btn_secondary")
        btn_add.clicked.connect(self._add_zeile)
        grp_layout.addWidget(btn_add)
        main_layout.addWidget(grp_artikel)

        # Aktions-Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        btn_ordner = QPushButton("Protokoll-Ordner öffnen")
        btn_ordner.setObjectName("btn_secondary")
        def _open_ordner():
            import os as _os
            d = get_ausgabe_dir() if self._rb_ausgabe.isChecked() else get_ruecknahme_dir()
            _os.makedirs(d, exist_ok=True)
            _os.startfile(d)
        btn_ordner.clicked.connect(_open_ordner)
        btn_row.addWidget(btn_ordner)

        btn_reset = QPushButton("Zurücksetzen")
        btn_reset.setObjectName("btn_secondary")
        btn_reset.clicked.connect(self._reset_form)
        btn_row.addWidget(btn_reset)

        btn_protokoll = QPushButton("📄  Protokoll erstellen")
        btn_protokoll.setObjectName("btn_primary")
        btn_protokoll.clicked.connect(self._create_protokoll)
        btn_row.addWidget(btn_protokoll)
        main_layout.addLayout(btn_row)

        self.lbl_result = QLabel("")
        self.lbl_result.setVisible(False)
        self.lbl_result.setWordWrap(True)
        main_layout.addWidget(self.lbl_result)
        main_layout.addStretch()

        self._sidebar = MaSidebar(self.db)
        self._sidebar.set_select_callback(self._on_ma_selected_sidebar)
        self._sidebar.setMinimumWidth(180)
        self._sidebar.setMaximumWidth(280)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self._sidebar)
        splitter.addWidget(scroll)
        splitter.setSizes([230, 800])
        outer.addWidget(splitter)

    def _on_ma_selected_sidebar(self, entry: dict):
        idx = self.cb_ma.findData(entry["id"])
        if idx >= 0:
            self.cb_ma.blockSignals(True)
            self.cb_ma.setCurrentIndex(idx)
            self.cb_ma.lineEdit().setText(self.cb_ma.itemText(idx))
            self.cb_ma.blockSignals(False)
            popup = self.cb_ma.completer().popup()
            if popup and popup.isVisible():
                popup.hide()
        else:
            self.cb_ma.setCurrentIndex(0)
            self.le_ma_freitext.setText(entry["name"])
        self._load_gespeicherte_kleidung()

    def _load_combos(self):
        self.cb_ma.blockSignals(True)
        self.cb_ma.clear()
        self.cb_ma.addItem("– Aus Liste wählen –", None)
        for ma in self.db.get_alle_mitarbeiter():
            display = f"{ma['nachname']}, {ma['vorname']}"
            if ma.get("personalnummer"):
                display += f" ({ma['personalnummer']})"
            self.cb_ma.addItem(display, ma["id"])
        self.cb_ma.blockSignals(False)
        if not self._artikel_rows:
            self._add_zeile()

    def _add_zeile(self):
        row_w = QWidget()
        hl = QHBoxLayout(row_w)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(8)

        cb_art = QComboBox()
        cb_art.setFixedWidth(200)
        cb_art.setEditable(True)
        cb_art.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        for art in self.db.get_kleidungsarten():
            cb_art.addItem(art["name"], art["id"])

        le_gr = QLineEdit()
        le_gr.setPlaceholderText("Größe ...")
        le_gr.setFixedWidth(130)

        sb = QSpinBox()
        sb.setRange(1, 999)
        sb.setValue(1)
        sb.setFixedWidth(80)

        btn_x = QPushButton("×")
        btn_x.setFixedSize(28, 28)
        btn_x.setStyleSheet(
            "QPushButton{color:#B20000;font-weight:bold;font-size:16px;border:none;background:transparent;}"
            "QPushButton:hover{background:#fce4e4;border-radius:4px;}"
        )

        hl.addWidget(cb_art)
        hl.addWidget(le_gr)
        hl.addWidget(sb)
        hl.addWidget(btn_x)
        hl.addStretch()

        ri = {"widget": row_w, "cb_art": cb_art, "le_gr": le_gr, "sb": sb}
        btn_x.clicked.connect(lambda _, rw=row_w, r=ri: self._remove_zeile(rw, r))
        self._artikel_rows.append(ri)
        self._artikel_container.addWidget(row_w)

    def _remove_zeile(self, row_w, ri):
        if len(self._artikel_rows) <= 1:
            return
        self._artikel_rows.remove(ri)
        self._artikel_container.removeWidget(row_w)
        row_w.deleteLater()

    def _reset_form(self):
        self.cb_ma.setCurrentIndex(0)
        self.le_ma_freitext.clear()
        self.de_datum.setDate(QDate.currentDate())
        self.le_von.clear()
        self.le_bem.clear()
        self._rb_ausgabe.setChecked(True)
        self.lbl_result.setVisible(False)
        for ri in list(self._artikel_rows):
            ri["widget"].deleteLater()
        self._artikel_rows.clear()
        self._add_zeile()

    def _load_gespeicherte_kleidung(self):
        ma_id = self.cb_ma.currentData()
        self._tbl_gespeichert.setRowCount(0)
        if not ma_id:
            self._lbl_gs_hint.setText("← Mitarbeiter auswählen, um gespeicherte Kleidung zu laden")
            self._lbl_gs_hint.setVisible(True)
            return
        items = self.db.get_mitarbeiter_kleidung(mitarbeiter_id=ma_id, status="ausgegeben")
        if not items:
            self._lbl_gs_hint.setText("Keine gespeicherte Kleidung für diesen Mitarbeiter vorhanden.")
            self._lbl_gs_hint.setVisible(True)
            return
        self._lbl_gs_hint.setVisible(False)
        self._tbl_gespeichert.setRowCount(len(items))
        for r, it in enumerate(items):
            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            chk.setCheckState(Qt.CheckState.Checked)
            chk.setData(Qt.ItemDataRole.UserRole, it)
            self._tbl_gespeichert.setItem(r, 0, chk)
            self._tbl_gespeichert.setItem(r, 1, QTableWidgetItem(it["art_name"]))
            self._tbl_gespeichert.setItem(r, 2, QTableWidgetItem(str(it["groesse"])))
            self._tbl_gespeichert.setItem(r, 3, QTableWidgetItem(str(it["menge"])))
            self._tbl_gespeichert.setItem(r, 4, QTableWidgetItem(format_datum(it.get("ausgabe_datum", ""))))

    def _set_all_checks(self, checked: bool):
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for r in range(self._tbl_gespeichert.rowCount()):
            chk = self._tbl_gespeichert.item(r, 0)
            if chk:
                chk.setCheckState(state)

    def _create_protokoll(self):
        ma_id = self.cb_ma.currentData()
        ma_name = ""
        if ma_id:
            ma_name = self.cb_ma.currentText().split(" (")[0].strip()
        elif self.le_ma_freitext.text().strip():
            ma_name = self.le_ma_freitext.text().strip()
        else:
            QMessageBox.warning(self, "Pflichtfeld", "Bitte einen Mitarbeiter auswählen oder eingeben.")
            return

        datum = self.de_datum.date().toString("yyyy-MM-dd")
        ausgegeben_von = self.le_von.text().strip()
        bemerkung = self.le_bem.text().strip()

        # Gespeicherte Kleidung (angehakte Zeilen) bevorzugen
        artikel = []
        for r in range(self._tbl_gespeichert.rowCount()):
            chk = self._tbl_gespeichert.item(r, 0)
            if chk and chk.checkState() == Qt.CheckState.Checked:
                it = chk.data(Qt.ItemDataRole.UserRole)
                if it:
                    artikel.append({
                        "art_name": it["art_name"],
                        "groesse":  str(it["groesse"]),
                        "menge":    int(it["menge"]),
                    })

        # Fallback: manuell eingetragene Positionen
        if not artikel:
            for ri in self._artikel_rows:
                art_name = ri["cb_art"].currentText().strip()
                groesse = ri["le_gr"].text().strip()
                if not art_name:
                    continue
                artikel.append({
                    "art_name": art_name,
                    "groesse":  groesse,
                    "menge":    ri["sb"].value(),
                })

        if not artikel:
            QMessageBox.warning(
                self, "Nichts ausgewählt",
                "Bitte gespeicherte Kleidungsstücke per Häkchen auswählen "
                "oder manuell Positionen eintragen."
            )
            return

        if self._rb_ausgabe.isChecked():
            ok, path = create_ausgabe_protokoll(ma_name, datum, ausgegeben_von, bemerkung, artikel)
        else:
            proto_rows = [
                {"art_name": a["art_name"], "groesse": a["groesse"],
                 "menge": a["menge"], "lager": a["menge"], "entsorgt": 0}
                for a in artikel
            ]
            ok, path = create_rueckgabe_protokoll(ma_name, datum, bemerkung, proto_rows)

        if ok:
            open_document(path)
            self.lbl_result.setVisible(True)
            self.lbl_result.setText(f"✔  Protokoll erstellt:\n{path}")
            self.lbl_result.setStyleSheet(
                "background:#E8F5E9; border-left:4px solid #388E3C; padding:8px; border-radius:4px;"
            )
        else:
            QMessageBox.warning(self, "Protokoll-Fehler", path)

    def showEvent(self, event):
        super().showEvent(event)
        self._load_combos()
        self._sidebar.load()


class AusgabeView(QWidget):
    """Hauptansicht Ausgabe / Rückgabe."""

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        lbl_title = QLabel("Ausgabe & Rückgabe")
        lbl_title.setObjectName("page_title")
        lbl_sub = QLabel("Kleidung an Mitarbeiter ausgeben oder zurücknehmen")
        lbl_sub.setObjectName("page_subtitle")
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_sub)

        tabs = QTabWidget()
        tabs.addTab(AusgabeTab(self.db), "📤  Ausgabe")
        tabs.addTab(RueckgabeTab(self.db), "📥  Rückgabe")
        tabs.addTab(NachtraeglichesProtokollTab(self.db), "📄  Nachträgliches Protokoll")
        layout.addWidget(tabs)
