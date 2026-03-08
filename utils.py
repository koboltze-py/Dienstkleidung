"""
DRK Dienstkleidung - Hilfsfunktionen (Backup, Logging, Export)
"""

import os
import shutil
import logging
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BACKUP_DIR = os.path.join(BASE_DIR, "Backup")
DATABASE_DIR = os.path.join(BASE_DIR, "Database")
EXPORT_DIR = os.path.join(BASE_DIR, "Export")
LOGS_DIR = os.path.join(BASE_DIR, "Logs")


def setup_logging():
    """Logging ins Log-Verzeichnis einrichten."""
    os.makedirs(LOGS_DIR, exist_ok=True)
    log_file = os.path.join(LOGS_DIR, f"dienstkleidung_{datetime.now().strftime('%Y%m')}.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )
    return logging.getLogger("DRKDienstkleidung")


def create_backup() -> tuple[bool, str]:
    """
    Erstellt eine Sicherungskopie aller Datenbanken im Backup-Ordner.
    Gibt (Erfolg, Meldung) zurück.
    """
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_folder = os.path.join(BACKUP_DIR, f"backup_{timestamp}")
        os.makedirs(backup_folder, exist_ok=True)

        copied = []
        if os.path.exists(DATABASE_DIR):
            for filename in os.listdir(DATABASE_DIR):
                if filename.endswith(".db"):
                    src = os.path.join(DATABASE_DIR, filename)
                    dst = os.path.join(backup_folder, filename)
                    shutil.copy2(src, dst)
                    copied.append(filename)

        if not copied:
            return False, "Keine Datenbankdateien gefunden."

        msg = f"Backup erstellt: {backup_folder}\nDateien: {', '.join(copied)}"
        return True, msg
    except Exception as e:
        return False, f"Backup fehlgeschlagen: {e}"


def get_backups() -> list[dict]:
    """Gibt eine Liste vorhandener Backups zurück."""
    if not os.path.exists(BACKUP_DIR):
        return []
    backups = []
    for folder in sorted(os.listdir(BACKUP_DIR), reverse=True):
        path = os.path.join(BACKUP_DIR, folder)
        if os.path.isdir(path) and folder.startswith("backup_"):
            # Parse timestamp from folder name: backup_YYYYMMDD_HHMMSS
            try:
                ts_str = folder.replace("backup_", "")
                ts = datetime.strptime(ts_str, "%Y%m%d_%H%M%S")
                datum_str = ts.strftime("%d.%m.%Y %H:%M:%S")
            except Exception:
                datum_str = folder
            files = os.listdir(path)
            backups.append({"name": folder, "path": path, "datum": datum_str, "dateien": files})
    return backups


def export_table_to_csv(headers: list[str], rows: list[list], filename: str) -> tuple[bool, str]:
    """Exportiert eine Tabelle als CSV-Datei in den Export-Ordner."""
    try:
        os.makedirs(EXPORT_DIR, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_filename = f"{filename}_{timestamp}.csv"
        filepath = os.path.join(EXPORT_DIR, safe_filename)

        import csv
        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(headers)
            writer.writerows(rows)

        return True, filepath
    except Exception as e:
        return False, f"Export fehlgeschlagen: {e}"


def format_datum(datum_str: str) -> str:
    """Formatiert ein ISO-Datum (YYYY-MM-DD) als deutsches Datum (DD.MM.YYYY)."""
    if not datum_str:
        return ""
    try:
        dt = datetime.strptime(datum_str[:10], "%Y-%m-%d")
        return dt.strftime("%d.%m.%Y")
    except Exception:
        return datum_str


def parse_datum(datum_de: str) -> str:
    """Konvertiert deutsches Datum (DD.MM.YYYY) in ISO-Format (YYYY-MM-DD)."""
    try:
        dt = datetime.strptime(datum_de.strip(), "%d.%m.%Y")
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return datum_de


def today_iso() -> str:
    """Gibt das heutige Datum als ISO-String zurück."""
    return datetime.now().strftime("%Y-%m-%d")
