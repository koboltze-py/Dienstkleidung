"""
DRK Dienstkleidung – Pfad-Hilfsfunktionen
Stellt sicher, dass App- und Basisverzeichnis sowohl im Entwicklungsmodus
als auch als gefrorene EXE (PyInstaller) korrekt aufgelöst werden.
"""
import sys
import os


def get_app_dir() -> str:
    """Gibt das App/-Verzeichnis zurück (wo die EXE bzw. main.py liegt)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_base_dir() -> str:
    """Gibt das Dienstkleidung/-Wurzelverzeichnis zurück."""
    return os.path.dirname(get_app_dir())
