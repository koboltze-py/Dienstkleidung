"""
DRK Dienstkleidung – Einstellungen
Pfad-Konfiguration, Backup und sonstige Systemeinstellungen.
Nur zugänglich mit Passwort.
"""

import os
from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QFileDialog, QFormLayout, QMessageBox, QFrame,
    QGroupBox,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

import config


class EinstellungenView(QWidget):
    """Einstellungen: Pfade konfigurieren + vollständiges Backup erstellen."""

    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._path_edits: dict[str, QLineEdit] = {}
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(16)

        lbl_title = QLabel("Einstellungen")
        lbl_title.setObjectName("page_title")
        lbl_sub = QLabel("Pfade konfigurieren und Systemfunktionen verwalten")
        lbl_sub.setObjectName("page_subtitle")
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_sub)

        # --- Pfad-Konfiguration ---
        grp_paths = QGroupBox("Verzeichnisse")
        grp_paths.setStyleSheet("QGroupBox{font-weight:bold; font-size:12px; border:1px solid #ddd;"
                                "border-radius:6px; margin-top:8px; padding-top:12px;}"
                                "QGroupBox::title{subcontrol-origin:margin; left:8px;}")
        form = QFormLayout(grp_paths)
        form.setSpacing(10)
        form.setContentsMargins(16, 16, 16, 16)

        cfg = config.load()
        for key, label in config.LABELS.items():
            row_w = QWidget()
            row_l = QHBoxLayout(row_w)
            row_l.setContentsMargins(0, 0, 0, 0)
            row_l.setSpacing(6)

            le = QLineEdit(cfg.get(key, config.DEFAULTS.get(key, "")))
            le.setMinimumWidth(320)
            self._path_edits[key] = le
            row_l.addWidget(le)

            btn_browse = QPushButton("Durchsuchen")
            btn_browse.setObjectName("btn_secondary")
            btn_browse.setFixedWidth(100)
            btn_browse.clicked.connect(lambda _, k=key: self._browse(k))
            row_l.addWidget(btn_browse)

            form.addRow(f"{label}:", row_w)

        layout.addWidget(grp_paths)

        # Save paths button
        btn_save = QPushButton("Pfade speichern")
        btn_save.setObjectName("btn_primary")
        btn_save.setMaximumWidth(200)
        btn_save.clicked.connect(self._save_paths)
        layout.addWidget(btn_save)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#ddd;")
        layout.addWidget(sep)

        # --- Backup ---
        grp_backup = QGroupBox("Vollstaendiges Backup")
        grp_backup.setStyleSheet("QGroupBox{font-weight:bold; font-size:12px; border:1px solid #ddd;"
                                 "border-radius:6px; margin-top:8px; padding-top:12px;}"
                                 "QGroupBox::title{subcontrol-origin:margin; left:8px;}")
        bk_layout = QVBoxLayout(grp_backup)
        bk_layout.setContentsMargins(16, 16, 16, 16)
        bk_layout.setSpacing(8)

        lbl_bk = QLabel(
            "Erstellt ein komplettes Backup im konfigurierten Backup-Verzeichnis:\n"
            "• Alle Datenbank-Dateien (.db)\n"
            "• Buchungsverlauf als CSV\n"
            "• Bestand als CSV\n"
            "• Mitarbeiter-Kleidung als Excel"
        )
        lbl_bk.setWordWrap(True)
        bk_layout.addWidget(lbl_bk)

        btn_backup = QPushButton("Vollstaendiges Backup erstellen")
        btn_backup.setObjectName("btn_primary")
        btn_backup.setMaximumWidth(260)
        btn_backup.clicked.connect(self._do_full_backup)
        bk_layout.addWidget(btn_backup)

        layout.addWidget(grp_backup)
        layout.addStretch()

    def _browse(self, key: str):
        current = self._path_edits[key].text()
        folder = QFileDialog.getExistingDirectory(
            self, f"Verzeichnis waehlen: {config.LABELS.get(key, key)}", current
        )
        if folder:
            self._path_edits[key].setText(folder)

    def _save_paths(self):
        cfg = config.load()
        for key, le in self._path_edits.items():
            val = le.text().strip()
            if val:
                cfg[key] = val
        config.save(cfg)
        QMessageBox.information(self, "Gespeichert",
                                "Pfade wurden gespeichert.\n"
                                "Aenderungen gelten ab dem naechsten Protokoll/Export.")

    def _do_full_backup(self):
        from utils import create_full_backup
        ok, msg = create_full_backup(self.db)
        if ok:
            QMessageBox.information(self, "Backup erfolgreich", msg)
        else:
            QMessageBox.warning(self, "Backup fehlgeschlagen", msg)

    def showEvent(self, event):
        super().showEvent(event)
        # Pfade beim Öffnen neu laden (könnten sich geändert haben)
        cfg = config.load()
        for key, le in self._path_edits.items():
            le.setText(cfg.get(key, config.DEFAULTS.get(key, "")))
