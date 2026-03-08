"""
DRK Dienstkleidung - Hauptfenster
Sidebar-Navigation + gestapelter Inhaltsbereich.
"""

import os
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QLabel,
    QPushButton, QStackedWidget, QStatusBar, QMessageBox,
    QButtonGroup, QFrame, QSizePolicy, QDialog, QDialogButtonBox,
    QLineEdit, QInputDialog,
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QFont, QIcon, QPixmap

from modules.dashboard import DashboardView
from modules.bestand import BestandView
from modules.ausgabe import AusgabeView
from modules.mitarbeiter import MitarbeiterView
from modules.verlauf import VerlaufView
from modules.einstellungen import EinstellungenView
from modules.bestellung import BestellungView
from utils import create_full_backup

APP_VERSION = "1.0.0"

# Navigations-Einträge: (Label, Icon, Klassen-Referenz)
NAV_ITEMS = [
    ("Dashboard",          "Dash",     DashboardView),
    ("Bestand",            "Bestand",  BestandView),
    ("Ausgabe / Rueckgabe","Ausgabe",  AusgabeView),
    ("Mitarbeiter",        "MA",       MitarbeiterView),
    ("Buchungsverlauf",    "Verlauf",  VerlaufView),
    ("Bestellung",         "Bestell",  BestellungView),
]


class SetupPromptDialog(QDialog):
    """Dialog: Datenbank noch nicht vorhanden, Ersteinrichtung anbieten."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ersteinrichtung")
        self.setModal(True)
        self.setMinimumWidth(500)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        lbl = QLabel(
            "<b>Kleidungsdatenbank noch nicht vorhanden.</b><br><br>"
            "Möchten Sie jetzt die Datenbank aus der Excel-Datei erstellen?<br>"
            "Alternativ können Sie <code>App/setup_db.py</code> manuell ausführen."
        )
        lbl.setWordWrap(True)
        lbl.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(lbl)

        buttons = QDialogButtonBox()
        btn_yes = buttons.addButton("Jetzt erstellen", QDialogButtonBox.ButtonRole.AcceptRole)
        btn_no  = buttons.addButton("Später", QDialogButtonBox.ButtonRole.RejectRole)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)


class MainWindow(QMainWindow):
    """Hauptfenster der DRK Dienstkleidung-App."""

    _PW_EINSTELLUNGEN = "mettwurst"

    def __init__(self, db, role: str = "etz", parent=None):
        super().__init__(parent)
        self.db = db
        self.role = role  # "etz" or "gast"
        self.setWindowTitle("DRK Dienstkleidung · Erste-Hilfe-Station Flughafen Köln")
        self.setMinimumSize(1050, 680)
        self.resize(1250, 780)
        self._setup_ui()
        self._apply_role()
        self._check_db()

    # ------------------------------------------------------------------
    # UI Aufbau
    # ------------------------------------------------------------------

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # === Sidebar ===
        sidebar = self._build_sidebar()
        root_layout.addWidget(sidebar)

        # === Content-Bereich ===
        self._stack = QStackedWidget()
        self._stack.setObjectName("content_area")
        root_layout.addWidget(self._stack, stretch=1)

        # Statusleiste
        self._statusbar = QStatusBar()
        self.setStatusBar(self._statusbar)
        self._statusbar.showMessage("DRK Dienstkleidung bereit  ·  v" + APP_VERSION)

        # Views instanziieren und zum Stack hinzufügen
        self._views = []
        self._bestellung_view: "BestellungView | None" = None
        for i, (_label, _icon, ViewClass) in enumerate(NAV_ITEMS):
            view = ViewClass(self.db)
            self._stack.addWidget(view)
            self._views.append(view)
            if ViewClass.__name__ == "BestellungView":
                self._bestellung_view = view

        # Einstellungen-View (separat, immer vorhanden)
        self._einstellungen_view = EinstellungenView(self.db)
        self._stack.addWidget(self._einstellungen_view)
        self._einstellungen_idx = len(NAV_ITEMS)  # index in stack

        # Warenkorb-Callback: Bestand → Bestellung
        bestand_view = next((v for v in self._views if isinstance(v, BestandView)), None)
        if bestand_view and self._bestellung_view:
            bestand_idx = NAV_ITEMS.index(
                next(n for n in NAV_ITEMS if n[2].__name__ == "BestellungView")
            )
            def _cart_callback(item, _bidx=bestand_idx):
                self._bestellung_view.add_item_from_bestand(item)
                # Zur Bestellungsmaske wechseln
                self._stack.setCurrentIndex(_bidx)
                for btn in self._nav_buttons:
                    btn.setChecked(False)
                self._nav_buttons[_bidx].setChecked(True)
            bestand_view.set_bestellung_callback(_cart_callback)

        # Ersten Tab aktivieren
        self._nav_buttons[0].setChecked(True)
        self._stack.setCurrentIndex(0)

    def _build_sidebar(self) -> QWidget:
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(240)
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Logo-Bereich
        _LOGO_PATH = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "Data", "Logo", "Depot.jpg"
        )
        lbl_logo = QLabel()
        lbl_logo.setObjectName("sidebar_logo")
        lbl_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pixmap = QPixmap(_LOGO_PATH)
        if not pixmap.isNull():
            lbl_logo.setPixmap(
                pixmap.scaled(236, 210, Qt.AspectRatioMode.KeepAspectRatio,
                              Qt.TransformationMode.SmoothTransformation)
            )
        else:
            lbl_logo.setText("DRK Dienstkleidung")
        lbl_logo.setContentsMargins(4, 10, 4, 4)
        layout.addWidget(lbl_logo)

        lbl_sub = QLabel("Erste-Hilfe-Station\nFlughafen Köln")
        lbl_sub.setObjectName("sidebar_subtitle")
        lbl_sub.setWordWrap(True)
        layout.addWidget(lbl_sub)

        # Trennlinie
        div = QFrame()
        div.setObjectName("sidebar_divider")
        div.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(div)

        layout.addSpacing(6)

        # Nav-Buttons
        self._btn_group = QButtonGroup(self)
        self._btn_group.setExclusive(True)
        self._nav_buttons: list[QPushButton] = []

        for i, (label, _icon, _cls) in enumerate(NAV_ITEMS):
            btn = QPushButton(f"  {label}")
            btn.setObjectName("nav_btn")
            btn.setCheckable(True)
            btn.setMinimumHeight(44)
            btn.clicked.connect(lambda checked, idx=i: self._navigate(idx))
            self._btn_group.addButton(btn, i)
            self._nav_buttons.append(btn)
            layout.addWidget(btn)

        layout.addStretch()

        # Backup-Button
        div2 = QFrame()
        div2.setObjectName("sidebar_divider")
        div2.setFrameShape(QFrame.Shape.HLine)
        layout.addWidget(div2)

        btn_backup = QPushButton("  Backup erstellen")
        btn_backup.setObjectName("nav_btn")
        btn_backup.setMinimumHeight(44)
        btn_backup.clicked.connect(self._do_backup)
        layout.addWidget(btn_backup)

        self._btn_einstellungen = QPushButton("  Einstellungen")
        self._btn_einstellungen.setObjectName("nav_btn")
        self._btn_einstellungen.setMinimumHeight(44)
        self._btn_einstellungen.clicked.connect(self._open_einstellungen)
        layout.addWidget(self._btn_einstellungen)

        lbl_ver = QLabel(f"Version {APP_VERSION}")
        lbl_ver.setObjectName("sidebar_version")
        lbl_ver.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(lbl_ver)
        layout.addSpacing(8)

        return sidebar

    # ------------------------------------------------------------------
    # Navigation
    # ------------------------------------------------------------------

    def _navigate(self, index: int):
        self._stack.setCurrentIndex(index)
        label = NAV_ITEMS[index][0]
        self._statusbar.showMessage(f"{label}  ·  DRK Dienstkleidung v{APP_VERSION}")

    def _open_einstellungen(self):
        """Fragt nach Passwort und öffnet die Einstellungen."""
        pw, ok = QInputDialog.getText(
            self, "Einstellungen",
            "Bitte Passwort eingeben:",
            QLineEdit.EchoMode.Password,
        )
        if not ok:
            return
        if pw != self._PW_EINSTELLUNGEN:
            QMessageBox.warning(self, "Falsches Passwort", "Das eingegebene Passwort ist falsch.")
            return
        # Deselect nav buttons
        checked = self._btn_group.checkedButton()
        if checked:
            checked.setChecked(False)
        self._stack.setCurrentIndex(self._einstellungen_idx)
        self._statusbar.showMessage(f"Einstellungen  ·  DRK Dienstkleidung v{APP_VERSION}")

    def _apply_role(self):
        """Wendet Rollen-Beschränkungen an (Gast = nur lesen + Ausgabe)."""
        if self.role == "gast":
            self._btn_einstellungen.setVisible(False)
            # Bestand + Mitarbeiter + Verlauf auf readonly setzen
            for view in self._views:
                if hasattr(view, "set_readonly"):
                    view.set_readonly()
            user_lbl = "Gast"
        else:
            user_lbl = "Etz"
        self._statusbar.showMessage(
            f"Angemeldet als: {user_lbl}  ·  DRK Dienstkleidung v{APP_VERSION}"
        )

    # ------------------------------------------------------------------
    # Backup
    # ------------------------------------------------------------------

    def _do_backup(self):
        ok, msg = create_full_backup(self.db)
        if ok:
            QMessageBox.information(self, "Backup erfolgreich", msg)
            self._statusbar.showMessage("Backup erfolgreich erstellt.")
        else:
            QMessageBox.warning(self, "Backup fehlgeschlagen", msg)

    # ------------------------------------------------------------------
    # Ersteinrichtung prüfen
    # ------------------------------------------------------------------

    def _check_db(self):
        if self.db.kleidung_db_exists():
            return

        dlg = SetupPromptDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self._run_setup()

    def _run_setup(self):
        setup_py = os.path.join(os.path.dirname(os.path.abspath(__file__)), "setup_db.py")
        try:
            import subprocess
            import sys
            proc = subprocess.run(
                [sys.executable, setup_py],
                input="j\n",
                text=True,
                capture_output=True,
                timeout=60,
            )
            if proc.returncode == 0:
                self.db.initialize()
                QMessageBox.information(
                    self, "Einrichtung abgeschlossen",
                    "Datenbank wurde erfolgreich erstellt.\n\n" + proc.stdout[-500:]
                )
                # Views neu laden
                for view in self._views:
                    if hasattr(view, "_load_data"):
                        view._load_data()
            else:
                QMessageBox.warning(
                    self, "Fehler",
                    "Einrichtung fehlgeschlagen:\n" + proc.stderr[-500:]
                )
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Konnte setup_db.py nicht ausführen:\n{e}")
