"""
DRK Dienstkleidung - Einstiegspunkt
Starten mit: python main.py
"""

import sys
import os

# Sicherstellen, dass App-Verzeichnis im Python-Pfad ist (nur im Dev-Modus)
if not getattr(sys, 'frozen', False):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import (
    QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QLineEdit, QComboBox, QFrame, QMessageBox,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from database import DatabaseManager
from main_window import MainWindow
from styles import MAIN_STYLE
from utils import setup_logging


# Anmeldedaten: Benutzername → Passwort (None = kein Passwort nötig)
_USERS: dict[str, str | None] = {
    "Etz":     "schinken",
    "Kurthen": "cologne",
    "Gast":    None,
}
# Role-Mapping
_ROLES: dict[str, str] = {
    "Etz":     "etz",
    "Kurthen": "etz",
    "Gast":    "gast",
}


class LoginDialog(QDialog):
    """Anmelde-Dialog: Benutzer auswählen und ggf. Passwort eingeben."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("DRK Dienstkleidung – Anmeldung")
        self.setMinimumWidth(360)
        self.setModal(True)
        self._role: str = ""
        self._setup_ui()

    @property
    def role(self) -> str:
        return self._role

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(28, 24, 28, 20)
        layout.setSpacing(14)

        lbl_title = QLabel("DRK Dienstkleidung")
        f = QFont()
        f.setBold(True)
        f.setPointSize(14)
        lbl_title.setFont(f)
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_title.setStyleSheet("color:#344F6B;")
        layout.addWidget(lbl_title)

        lbl_sub = QLabel("Erste-Hilfe-Station Flughafen Köln")
        lbl_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_sub.setStyleSheet("color:#666; font-size:11px;")
        layout.addWidget(lbl_sub)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#ddd;")
        layout.addWidget(sep)

        lbl_user = QLabel("Benutzer:")
        layout.addWidget(lbl_user)
        self.cb_user = QComboBox()
        for name in _USERS:
            self.cb_user.addItem(name)
        self.cb_user.currentTextChanged.connect(self._on_user_changed)
        layout.addWidget(self.cb_user)

        self.lbl_pw = QLabel("Passwort:")
        layout.addWidget(self.lbl_pw)
        self.le_pw = QLineEdit()
        self.le_pw.setEchoMode(QLineEdit.EchoMode.Password)
        self.le_pw.setPlaceholderText("Passwort eingeben ...")
        self.le_pw.returnPressed.connect(self._login)
        layout.addWidget(self.le_pw)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)
        btn_login = QPushButton("Anmelden")
        btn_login.setObjectName("btn_primary")
        btn_login.clicked.connect(self._login)
        btn_row.addStretch()
        btn_row.addWidget(btn_login)
        layout.addLayout(btn_row)

        # Initial: set correct state
        self._on_user_changed(self.cb_user.currentText())

    def _on_user_changed(self, name: str):
        needs_pw = _USERS.get(name) is not None
        self.lbl_pw.setVisible(needs_pw)
        self.le_pw.setVisible(needs_pw)
        if not needs_pw:
            self.le_pw.clear()

    def _login(self):
        name = self.cb_user.currentText()
        expected_pw = _USERS.get(name)
        if expected_pw is not None:
            entered = self.le_pw.text()
            if entered != expected_pw:
                QMessageBox.warning(self, "Falsches Passwort",
                                    "Das eingegebene Passwort ist falsch.")
                self.le_pw.clear()
                self.le_pw.setFocus()
                return
        self._role = _ROLES.get(name, "gast")
        self.accept()


def main():
    # Logging aufsetzen
    logger = setup_logging()
    logger.info("DRK Dienstkleidung wird gestartet...")

    # Qt High-DPI aktivieren
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")

    app = QApplication(sys.argv)
    app.setApplicationName("DRK Dienstkleidung")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("DRK Kreisverband Köln e.V.")

    # Globaler Stil
    app.setStyleSheet(MAIN_STYLE)
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    # Login
    login = LoginDialog()
    if login.exec() != QDialog.DialogCode.Accepted:
        sys.exit(0)
    role = login.role

    # Datenbankmanager
    db = DatabaseManager()
    db.initialize()

    # Tägliches Datenbank-Backup
    from db_backup import run_daily_backup
    run_daily_backup()

    # Hauptfenster
    window = MainWindow(db, role=role)
    window.show()

    logger.info(f"App bereit (Rolle: {role}).")
    exit_code = app.exec()
    logger.info(f"App beendet (Code {exit_code}).")
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
