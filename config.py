"""
DRK Dienstkleidung – Pfad-Konfiguration
Alle Pfade zentral verwaltet und über config.json anpassbar.
"""

import os
import json

_APP_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_APP_DIR)
_CONFIG_FILE = os.path.join(_APP_DIR, "config.json")

DEFAULTS: dict[str, str] = {
    "ausgabe_dir":    os.path.join(_BASE_DIR, "Data", "Ausgabe Protokolle"),
    "ruecknahme_dir": os.path.join(_BASE_DIR, "Data", "Rücknahme Protokolle"),
    "export_dir":     os.path.join(_BASE_DIR, "Export"),
    "backup_dir":     os.path.join(_BASE_DIR, "Backup"),
    "database_dir":   os.path.join(_BASE_DIR, "Database"),
}

LABELS: dict[str, str] = {
    "ausgabe_dir":    "Ausgabe-Protokolle",
    "ruecknahme_dir": "Rücknahme-Protokolle",
    "export_dir":     "Export-Verzeichnis",
    "backup_dir":     "Backup-Verzeichnis",
    "database_dir":   "Datenbank-Verzeichnis",
}


def load() -> dict:
    """Lädt Konfiguration aus config.json, ergänzt fehlende Keys mit Defaults."""
    if os.path.exists(_CONFIG_FILE):
        try:
            with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {**DEFAULTS, **data}
        except Exception:
            pass
    return dict(DEFAULTS)


def save(data: dict) -> None:
    """Speichert Konfiguration in config.json."""
    with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def get(key: str) -> str:
    """Gibt einen konfigurierten Pfad zurück."""
    return load().get(key, DEFAULTS.get(key, ""))
