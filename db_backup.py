"""
DRK Dienstkleidung – Automatisches Datenbank-Backup
Wird beim App-Start aufgerufen.
- Erstellt maximal ein Backup pro Tag
- Löscht Backups die älter als 5 Tage sind
- Zielordner: <BASE_DIR>/Db Backup/
"""

import os
import shutil
import logging
from datetime import datetime, timedelta

from pathutils import get_base_dir
from database import MITARBEITER_DB, KLEIDUNG_DB

_BACKUP_DIR = os.path.join(get_base_dir(), "Db Backup")
_KEEP_DAYS   = 5
_DATE_FORMAT = "%Y-%m-%d"

logger = logging.getLogger(__name__)


def run_daily_backup() -> None:
    """
    Erstellt ein tägliches Backup beider Datenbanken.
    Wird übersprungen wenn heute bereits ein Backup existiert.
    Backups älter als _KEEP_DAYS Tage werden gelöscht.
    """
    today = datetime.now().strftime(_DATE_FORMAT)

    try:
        os.makedirs(_BACKUP_DIR, exist_ok=True)
    except Exception as exc:
        logger.warning(f"Backup-Ordner konnte nicht erstellt werden: {exc}")
        return

    # Prüfen ob heute bereits ein Backup existiert
    existing = [
        f for f in os.listdir(_BACKUP_DIR)
        if f.startswith(today) and f.endswith(".db")
    ]
    if existing:
        logger.debug(f"Backup heute bereits vorhanden ({existing[0]}), übersprungen.")
        return

    # Alte Backups löschen (> _KEEP_DAYS Tage)
    cutoff = datetime.now() - timedelta(days=_KEEP_DAYS)
    for fname in os.listdir(_BACKUP_DIR):
        if not fname.endswith(".db"):
            continue
        # Dateiname-Format: YYYY-MM-DD_<name>.db
        date_part = fname[:10]
        try:
            fdate = datetime.strptime(date_part, _DATE_FORMAT)
            if fdate < cutoff:
                os.remove(os.path.join(_BACKUP_DIR, fname))
                logger.info(f"Altes Backup gelöscht: {fname}")
        except ValueError:
            pass  # Unbekanntes Format, ignorieren

    # Neue Backups erstellen
    for src_path in (MITARBEITER_DB, KLEIDUNG_DB):
        if not os.path.exists(src_path):
            logger.warning(f"Datenbank nicht gefunden, kein Backup: {src_path}")
            continue
        db_name = os.path.splitext(os.path.basename(src_path))[0]
        dest = os.path.join(_BACKUP_DIR, f"{today}_{db_name}.db")
        try:
            shutil.copy2(src_path, dest)
            logger.info(f"Backup erstellt: {dest}")
        except Exception as exc:
            logger.error(f"Backup fehlgeschlagen für {src_path}: {exc}")
