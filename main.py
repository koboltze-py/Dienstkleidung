"""
DRK Dienstkleidung - Einstiegspunkt
Starten mit: python main.py
"""

import sys
import os

# Sicherstellen, dass App-Verzeichnis im Python-Pfad ist
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from database import DatabaseManager
from main_window import MainWindow
from styles import MAIN_STYLE
from utils import setup_logging


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

    # Datenbankmanager
    db = DatabaseManager()
    db.initialize()

    # Hauptfenster
    window = MainWindow(db)
    window.show()

    logger.info("App bereit.")
    exit_code = app.exec()
    logger.info(f"App beendet (Code {exit_code}).")
    sys.exit(exit_code)


if __name__ == "__main__":
    main()
