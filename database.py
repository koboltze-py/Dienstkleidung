"""
DRK Dienstkleidung - Datenbankmanager
Verbindet sich mit mitarbeiter.db und kleidung.db
Alle Datenbankoperationen sind hier zentralisiert.
"""

import sqlite3
import os
from typing import Optional

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MITARBEITER_DB = os.path.join(BASE_DIR, "Database", "mitarbeiter.db")
KLEIDUNG_DB = os.path.join(BASE_DIR, "Database", "kleidung.db")

_CREATE_TABLES_SQL = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS kleidungsarten (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    name         TEXT    NOT NULL UNIQUE,
    beschreibung TEXT    DEFAULT '',
    aktiv        INTEGER DEFAULT 1,
    erstellt_am  TEXT    DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS kleidungsbestand (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    art_id      INTEGER NOT NULL REFERENCES kleidungsarten(id) ON DELETE CASCADE,
    groesse     TEXT    NOT NULL,
    menge       INTEGER NOT NULL DEFAULT 0,
    min_menge   INTEGER DEFAULT 0,
    bemerkung   TEXT    DEFAULT '',
    erstellt_am TEXT    DEFAULT (datetime('now','localtime')),
    geaendert_am TEXT   DEFAULT (datetime('now','localtime')),
    UNIQUE(art_id, groesse)
);

CREATE TABLE IF NOT EXISTS mitarbeiter_kleidung (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    mitarbeiter_id   INTEGER,
    mitarbeiter_name TEXT    NOT NULL,
    art_id           INTEGER NOT NULL REFERENCES kleidungsarten(id),
    art_name         TEXT    NOT NULL,
    groesse          TEXT    NOT NULL,
    menge            INTEGER NOT NULL DEFAULT 1,
    ausgabe_datum    TEXT    NOT NULL,
    rueckgabe_datum  TEXT,
    status           TEXT    DEFAULT 'ausgegeben',
    ausgegeben_von   TEXT    DEFAULT '',
    bemerkung        TEXT    DEFAULT '',
    erstellt_am      TEXT    DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS buchungen (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    art_id           INTEGER,
    art_name         TEXT    NOT NULL,
    groesse          TEXT    NOT NULL,
    menge            INTEGER NOT NULL,
    typ              TEXT    NOT NULL,
    mitarbeiter_id   INTEGER,
    mitarbeiter_name TEXT    DEFAULT '',
    datum            TEXT    NOT NULL,
    ausgegeben_von   TEXT    DEFAULT '',
    bemerkung        TEXT    DEFAULT '',
    erstellt_am      TEXT    DEFAULT (datetime('now','localtime'))
);

CREATE INDEX IF NOT EXISTS idx_bestand_art        ON kleidungsbestand(art_id);
CREATE INDEX IF NOT EXISTS idx_buchungen_datum    ON buchungen(datum);
CREATE INDEX IF NOT EXISTS idx_buchungen_ma       ON buchungen(mitarbeiter_id);
CREATE INDEX IF NOT EXISTS idx_mk_mitarbeiter     ON mitarbeiter_kleidung(mitarbeiter_id);
CREATE INDEX IF NOT EXISTS idx_mk_name            ON mitarbeiter_kleidung(mitarbeiter_name);
CREATE INDEX IF NOT EXISTS idx_mk_status          ON mitarbeiter_kleidung(status);
"""


class DatabaseManager:
    """Zentrale Datenbankverwaltung für die DRK Dienstkleidung-App."""

    def __init__(self):
        self.mitarbeiter_db = MITARBEITER_DB
        self.kleidung_db = KLEIDUNG_DB

    # ------------------------------------------------------------------
    # Verbindungen
    # ------------------------------------------------------------------

    def _conn_ma(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.mitarbeiter_db)
        conn.row_factory = sqlite3.Row
        return conn

    def _conn_kl(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.kleidung_db)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    # ------------------------------------------------------------------
    # Initialisierung
    # ------------------------------------------------------------------

    def initialize(self):
        """Erstellt kleidung.db und alle Tabellen, falls noch nicht vorhanden."""
        if not os.path.exists(self.kleidung_db):
            self._create_tables()

    def _create_tables(self):
        conn = sqlite3.connect(self.kleidung_db)
        conn.executescript(_CREATE_TABLES_SQL)
        conn.commit()
        conn.close()

    def kleidung_db_exists(self) -> bool:
        return os.path.exists(self.kleidung_db)

    # ------------------------------------------------------------------
    # Mitarbeiter (aus mitarbeiter.db)
    # ------------------------------------------------------------------

    def get_alle_mitarbeiter(self) -> list[dict]:
        """Alle aktiven Mitarbeiter aus mitarbeiter.db"""
        if not os.path.exists(self.mitarbeiter_db):
            return []
        conn = self._conn_ma()
        cur = conn.execute(
            "SELECT id, vorname, nachname, personalnummer, position, abteilung "
            "FROM mitarbeiter WHERE status = 'aktiv' ORDER BY nachname, vorname"
        )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def get_mitarbeiter_by_id(self, mid: int) -> Optional[dict]:
        if not os.path.exists(self.mitarbeiter_db):
            return None
        conn = self._conn_ma()
        cur = conn.execute("SELECT * FROM mitarbeiter WHERE id = ?", (mid,))
        row = cur.fetchone()
        conn.close()
        return dict(row) if row else None

    def add_mitarbeiter(self, vorname: str, nachname: str, personalnummer: str = "",
                        position: str = "", abteilung: str = "") -> tuple[bool, str]:
        if not os.path.exists(self.mitarbeiter_db):
            return False, "Mitarbeiter-Datenbank nicht gefunden."
        conn = self._conn_ma()
        try:
            conn.execute(
                "INSERT INTO mitarbeiter (vorname, nachname, personalnummer, position, abteilung, status) "
                "VALUES (?, ?, ?, ?, ?, 'aktiv')",
                (vorname, nachname, personalnummer, position, abteilung)
            )
            conn.commit()
            return True, f"{vorname} {nachname} wurde angelegt."
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    def delete_mitarbeiter(self, mid: int) -> tuple[bool, str]:
        if not os.path.exists(self.mitarbeiter_db):
            return False, "Mitarbeiter-Datenbank nicht gefunden."
        conn = self._conn_ma()
        try:
            conn.execute("DELETE FROM mitarbeiter WHERE id = ?", (mid,))
            conn.commit()
            return True, "Mitarbeiter wurde gelöscht."
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    # ------------------------------------------------------------------
    # Kleidungsarten
    # ------------------------------------------------------------------

    def get_kleidungsarten(self) -> list[dict]:
        conn = self._conn_kl()
        cur = conn.execute(
            "SELECT id, name, beschreibung FROM kleidungsarten WHERE aktiv=1 ORDER BY name"
        )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def add_kleidungsart(self, name: str, beschreibung: str = "") -> int:
        conn = self._conn_kl()
        cur = conn.execute(
            "INSERT INTO kleidungsarten (name, beschreibung) VALUES (?, ?)", (name, beschreibung)
        )
        new_id = cur.lastrowid
        conn.commit()
        conn.close()
        return new_id

    def update_kleidungsart(self, art_id: int, name: str, beschreibung: str = "") -> None:
        conn = self._conn_kl()
        conn.execute(
            "UPDATE kleidungsarten SET name=?, beschreibung=? WHERE id=?",
            (name, beschreibung, art_id),
        )
        conn.commit()
        conn.close()

    def delete_kleidungsart(self, art_id: int) -> None:
        conn = self._conn_kl()
        conn.execute("UPDATE kleidungsarten SET aktiv=0 WHERE id=?", (art_id,))
        conn.commit()
        conn.close()

    # ------------------------------------------------------------------
    # Kleidungsbestand
    # ------------------------------------------------------------------

    def get_bestand(self, art_id: Optional[int] = None) -> list[dict]:
        conn = self._conn_kl()
        if art_id:
            cur = conn.execute(
                """SELECT b.id, a.name AS art_name, b.art_id, b.groesse,
                          b.menge, b.min_menge, b.bemerkung
                   FROM kleidungsbestand b
                   JOIN kleidungsarten a ON b.art_id = a.id
                   WHERE b.art_id = ?
                   ORDER BY b.groesse""",
                (art_id,),
            )
        else:
            cur = conn.execute(
                """SELECT b.id, a.name AS art_name, b.art_id, b.groesse,
                          b.menge, b.min_menge, b.bemerkung
                   FROM kleidungsbestand b
                   JOIN kleidungsarten a ON b.art_id = a.id
                   ORDER BY a.name, b.groesse"""
            )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def get_bestand_item(self, art_id: int, groesse: str) -> Optional[dict]:
        conn = self._conn_kl()
        cur = conn.execute(
            """SELECT b.id, a.name AS art_name, b.art_id, b.groesse, b.menge, b.min_menge, b.bemerkung
               FROM kleidungsbestand b
               JOIN kleidungsarten a ON b.art_id = a.id
               WHERE b.art_id=? AND b.groesse=?""",
            (art_id, groesse),
        )
        row = cur.fetchone()
        conn.close()
        return dict(row) if row else None

    def upsert_bestand(
        self,
        art_id: int,
        groesse: str,
        menge: int,
        min_menge: int = 0,
        bemerkung: str = "",
    ) -> None:
        conn = self._conn_kl()
        conn.execute(
            """INSERT INTO kleidungsbestand (art_id, groesse, menge, min_menge, bemerkung)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(art_id, groesse) DO UPDATE SET
                 menge       = excluded.menge,
                 min_menge   = excluded.min_menge,
                 bemerkung   = excluded.bemerkung,
                 geaendert_am = datetime('now','localtime')""",
            (art_id, groesse, menge, min_menge, bemerkung),
        )
        conn.commit()
        conn.close()

    def delete_bestand_item(self, item_id: int) -> None:
        conn = self._conn_kl()
        conn.execute("DELETE FROM kleidungsbestand WHERE id=?", (item_id,))
        conn.commit()
        conn.close()

    def set_default_min_menge(self, default: int = 3) -> int:
        """Setzt min_menge=default für alle Einträge wo min_menge noch 0 ist.
        Gibt die Anzahl aktualisierter Zeilen zurück."""
        conn = self._conn_kl()
        cur = conn.execute(
            "UPDATE kleidungsbestand SET min_menge=? WHERE min_menge=0 OR min_menge IS NULL",
            (default,),
        )
        count = cur.rowcount
        conn.commit()
        conn.close()
        return count

    def get_groessen_fuer_art(self, art_id: int) -> list[dict]:
        """Liefert alle Größen mit Menge > 0 für eine Kleidungsart."""
        conn = self._conn_kl()
        cur = conn.execute(
            "SELECT groesse, menge FROM kleidungsbestand WHERE art_id=? AND menge > 0 ORDER BY groesse",
            (art_id,),
        )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def get_bestand_summary(self) -> list[dict]:
        conn = self._conn_kl()
        cur = conn.execute(
            """SELECT a.id, a.name,
                      COUNT(b.id) AS anzahl_groessen,
                      COALESCE(SUM(b.menge), 0) AS gesamt_menge,
                      SUM(CASE WHEN b.menge <= b.min_menge AND b.min_menge > 0 THEN 1 ELSE 0 END) AS niedrig_anzahl
               FROM kleidungsarten a
               LEFT JOIN kleidungsbestand b ON a.id = b.art_id
               WHERE a.aktiv = 1
               GROUP BY a.id, a.name
               ORDER BY a.name"""
        )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def get_niedrig_bestand(self) -> list[dict]:
        conn = self._conn_kl()
        cur = conn.execute(
            """SELECT a.name AS art_name, b.groesse, b.menge, b.min_menge
               FROM kleidungsbestand b
               JOIN kleidungsarten a ON b.art_id = a.id
               WHERE b.menge <= b.min_menge AND b.min_menge > 0
               ORDER BY b.menge ASC"""
        )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    # ------------------------------------------------------------------
    # Ausgabe / Rückgabe / Eingang
    # ------------------------------------------------------------------

    def ausgabe_kleidung(
        self,
        mitarbeiter_id: Optional[int],
        mitarbeiter_name: str,
        art_id: int,
        art_name: str,
        groesse: str,
        menge: int,
        datum: str,
        ausgegeben_von: str,
        bemerkung: str = "",
    ) -> tuple[bool, str]:
        """Kleidung ausgeben. Gibt (Erfolg, Meldung) zurück."""
        conn = self._conn_kl()
        try:
            cur = conn.execute(
                "SELECT menge FROM kleidungsbestand WHERE art_id=? AND groesse=?",
                (art_id, groesse),
            )
            row = cur.fetchone()
            if not row:
                return False, f"Kein Bestand für {art_name} Größe {groesse} gefunden."
            if row[0] < menge:
                return (
                    False,
                    f"Nicht genug auf Lager. Verfügbar: {row[0]}, Angefordert: {menge}",
                )

            conn.execute(
                """UPDATE kleidungsbestand
                   SET menge = menge - ?, geaendert_am = datetime('now','localtime')
                   WHERE art_id=? AND groesse=?""",
                (menge, art_id, groesse),
            )
            conn.execute(
                """INSERT INTO mitarbeiter_kleidung
                   (mitarbeiter_id, mitarbeiter_name, art_id, art_name, groesse, menge,
                    ausgabe_datum, ausgegeben_von, bemerkung, status)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'ausgegeben')""",
                (mitarbeiter_id, mitarbeiter_name, art_id, art_name, groesse, menge,
                 datum, ausgegeben_von, bemerkung),
            )
            conn.execute(
                """INSERT INTO buchungen
                   (art_id, art_name, groesse, menge, typ, mitarbeiter_id, mitarbeiter_name,
                    datum, ausgegeben_von, bemerkung)
                   VALUES (?, ?, ?, ?, 'ausgabe', ?, ?, ?, ?, ?)""",
                (art_id, art_name, groesse, -menge, mitarbeiter_id, mitarbeiter_name,
                 datum, ausgegeben_von, bemerkung),
            )
            conn.commit()
            return True, "Ausgabe erfolgreich gespeichert."
        except Exception as e:
            conn.rollback()
            return False, f"Fehler bei der Ausgabe: {e}"
        finally:
            conn.close()

    def rueckgabe_kleidung(
        self, mk_id: int, rueckgabe_datum: str, bemerkung: str = ""
    ) -> tuple[bool, str]:
        """Kleidung zurücknehmen. Gibt (Erfolg, Meldung) zurück."""
        conn = self._conn_kl()
        try:
            cur = conn.execute("SELECT * FROM mitarbeiter_kleidung WHERE id=?", (mk_id,))
            row = cur.fetchone()
            if not row:
                return False, "Datensatz nicht gefunden."
            if row["status"] == "zurueckgegeben":
                return False, "Kleidung wurde bereits zurückgegeben."

            art_id = row["art_id"]
            art_name = row["art_name"]
            groesse = row["groesse"]
            menge = row["menge"]
            mitarbeiter_id = row["mitarbeiter_id"]
            mitarbeiter_name = row["mitarbeiter_name"]

            conn.execute(
                """UPDATE kleidungsbestand
                   SET menge = menge + ?, geaendert_am = datetime('now','localtime')
                   WHERE art_id=? AND groesse=?""",
                (menge, art_id, groesse),
            )
            conn.execute(
                """UPDATE mitarbeiter_kleidung
                   SET status='zurueckgegeben', rueckgabe_datum=?
                   WHERE id=?""",
                (rueckgabe_datum, mk_id),
            )
            conn.execute(
                """INSERT INTO buchungen
                   (art_id, art_name, groesse, menge, typ, mitarbeiter_id, mitarbeiter_name,
                    datum, bemerkung)
                   VALUES (?, ?, ?, ?, 'rueckgabe', ?, ?, ?, ?)""",
                (art_id, art_name, groesse, menge, mitarbeiter_id, mitarbeiter_name,
                 rueckgabe_datum, bemerkung),
            )
            conn.commit()
            return True, "Rückgabe erfolgreich gespeichert."
        except Exception as e:
            conn.rollback()
            return False, f"Fehler bei der Rückgabe: {e}"
        finally:
            conn.close()

    def rueckgabe_mit_entsorgung(
        self, mk_id: int, lager_menge: int, entsorgt_menge: int,
        rueckgabe_datum: str, bemerkung: str = ""
    ) -> tuple[bool, str]:
        """Kleidung aufgeteilt zurücknehmen: Teil ins Lager, Teil entsorgen."""
        conn = self._conn_kl()
        try:
            row = conn.execute("SELECT * FROM mitarbeiter_kleidung WHERE id=?", (mk_id,)).fetchone()
            if not row:
                return False, "Datensatz nicht gefunden."
            if row["status"] == "zurueckgegeben":
                return False, "Kleidung wurde bereits zurückgegeben."

            art_id = row["art_id"]
            art_name = row["art_name"]
            groesse = row["groesse"]
            ma_id = row["mitarbeiter_id"]
            ma_name = row["mitarbeiter_name"]

            if lager_menge > 0:
                conn.execute(
                    "UPDATE kleidungsbestand SET menge=menge+?, geaendert_am=datetime('now','localtime') "
                    "WHERE art_id=? AND groesse=?",
                    (lager_menge, art_id, groesse),
                )
                conn.execute(
                    "INSERT INTO buchungen (art_id, art_name, groesse, menge, typ, "
                    "mitarbeiter_id, mitarbeiter_name, datum, bemerkung) "
                    "VALUES (?,?,?,?,'rueckgabe',?,?,?,?)",
                    (art_id, art_name, groesse, lager_menge, ma_id, ma_name, rueckgabe_datum, bemerkung),
                )

            if entsorgt_menge > 0:
                bem_entsorgt = f"Entsorgung – {bemerkung}".strip(" –") if bemerkung else "Entsorgung"
                conn.execute(
                    "INSERT INTO buchungen (art_id, art_name, groesse, menge, typ, "
                    "mitarbeiter_id, mitarbeiter_name, datum, bemerkung) "
                    "VALUES (?,?,?,?,'entsorgung',?,?,?,?)",
                    (art_id, art_name, groesse, entsorgt_menge, ma_id, ma_name, rueckgabe_datum, bem_entsorgt),
                )

            conn.execute(
                "UPDATE mitarbeiter_kleidung SET status='zurueckgegeben', rueckgabe_datum=? WHERE id=?",
                (rueckgabe_datum, mk_id),
            )
            conn.commit()
            return True, "Rückgabe erfolgreich gespeichert."
        except Exception as e:
            conn.rollback()
            return False, f"Fehler bei der Rückgabe: {e}"
        finally:
            conn.close()

    def update_mitarbeiter_kleidung(
        self, mk_id: int, menge: int, ausgabe_datum: str,
        ausgegeben_von: str, bemerkung: str
    ) -> tuple[bool, str]:
        """Editiert eine bestehende Kleidungsposition (Ausgabe-Metadaten)."""
        conn = self._conn_kl()
        try:
            row = conn.execute(
                "SELECT * FROM mitarbeiter_kleidung WHERE id=?", (mk_id,)
            ).fetchone()
            if not row:
                return False, "Datensatz nicht gefunden."

            alte_menge = row["menge"]
            menge_diff = menge - alte_menge

            if menge_diff != 0:
                if menge_diff > 0:
                    bestand = conn.execute(
                        "SELECT menge FROM kleidungsbestand WHERE art_id=? AND groesse=?",
                        (row["art_id"], row["groesse"]),
                    ).fetchone()
                    verfuegbar = bestand[0] if bestand else 0
                    if verfuegbar < menge_diff:
                        return False, (
                            f"Nicht genug im Bestand. Verfügbar: {verfuegbar}, "
                            f"benötigt: {menge_diff} mehr."
                        )
                conn.execute(
                    "UPDATE kleidungsbestand SET menge=menge-?, "
                    "geaendert_am=datetime('now','localtime') "
                    "WHERE art_id=? AND groesse=?",
                    (menge_diff, row["art_id"], row["groesse"]),
                )

            conn.execute(
                """UPDATE mitarbeiter_kleidung
                   SET menge=?, ausgabe_datum=?, ausgegeben_von=?, bemerkung=?
                   WHERE id=?""",
                (menge, ausgabe_datum, ausgegeben_von, bemerkung, mk_id),
            )
            conn.commit()
            return True, "Eintrag erfolgreich gespeichert."
        except Exception as e:
            conn.rollback()
            return False, f"Fehler beim Speichern: {e}"
        finally:
            conn.close()

    def eingang_kleidung(
        self, art_id: int, art_name: str, groesse: str, menge: int,
        datum: str, bemerkung: str = ""
    ) -> tuple[bool, str]:
        """Neue Kleidung einbuchen (Wareneingang)."""
        conn = self._conn_kl()
        try:
            conn.execute(
                """INSERT INTO kleidungsbestand (art_id, groesse, menge)
                   VALUES (?, ?, ?)
                   ON CONFLICT(art_id, groesse) DO UPDATE SET
                     menge = menge + excluded.menge,
                     geaendert_am = datetime('now','localtime')""",
                (art_id, groesse, menge),
            )
            conn.execute(
                """INSERT INTO buchungen
                   (art_id, art_name, groesse, menge, typ, datum, bemerkung)
                   VALUES (?, ?, ?, ?, 'eingang', ?, ?)""",
                (art_id, art_name, groesse, menge, datum, bemerkung),
            )
            conn.commit()
            return True, f"Eingang von {menge}x {art_name} Gr. {groesse} gebucht."
        except Exception as e:
            conn.rollback()
            return False, f"Fehler beim Eingang: {e}"
        finally:
            conn.close()

    def ausbuchen_bestand(
        self, art_id: int, art_name: str, groesse: str, menge: int,
        datum: str, grund: str, bemerkung: str = ""
    ) -> tuple[bool, str]:
        """Ware aus dem Bestand ausbuchen (Defekt, Verlust, Fehlbuchung etc.)."""
        conn = self._conn_kl()
        try:
            row = conn.execute(
                "SELECT menge FROM kleidungsbestand WHERE art_id=? AND groesse=?",
                (art_id, groesse),
            ).fetchone()
            if not row:
                return False, f"Kein Bestand für {art_name} Größe {groesse} gefunden."
            if row[0] < menge:
                return False, f"Nicht genug auf Lager. Verfügbar: {row[0]}, Angefordert: {menge}"
            conn.execute(
                "UPDATE kleidungsbestand SET menge=menge-?, geaendert_am=datetime('now','localtime') "
                "WHERE art_id=? AND groesse=?",
                (menge, art_id, groesse),
            )
            bem_full = f"{grund}" + (f" – {bemerkung}" if bemerkung else "")
            conn.execute(
                "INSERT INTO buchungen (art_id, art_name, groesse, menge, typ, datum, bemerkung) "
                "VALUES (?,?,?,?,'ausbuchen',?,?)",
                (art_id, art_name, groesse, -menge, datum, bem_full),
            )
            conn.commit()
            return True, f"{menge}x {art_name} Gr. {groesse} ausgebucht."
        except Exception as e:
            conn.rollback()
            return False, f"Fehler beim Ausbuchen: {e}"
        finally:
            conn.close()

    def get_buchungen_fuer_bestand(self, art_id: int, groesse: str) -> list[dict]:
        """Alle Buchungen für eine bestimmte art_id + Größe."""
        conn = self._conn_kl()
        cur = conn.execute(
            "SELECT datum, typ, menge, mitarbeiter_name, ausgegeben_von, bemerkung, erstellt_am "
            "FROM buchungen WHERE art_id=? AND groesse=? "
            "ORDER BY datum DESC, erstellt_am DESC LIMIT 200",
            (art_id, groesse),
        )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def korrektur_bestand(
        self, art_id: int, art_name: str, groesse: str,
        neue_menge: int, datum: str, bemerkung: str = ""
    ) -> tuple[bool, str]:
        """Bestandskorrektur (direkte Mengensetzung)."""
        conn = self._conn_kl()
        try:
            cur = conn.execute(
                "SELECT menge FROM kleidungsbestand WHERE art_id=? AND groesse=?",
                (art_id, groesse),
            )
            row = cur.fetchone()
            alte_menge = row[0] if row else 0
            delta = neue_menge - alte_menge

            conn.execute(
                """UPDATE kleidungsbestand
                   SET menge=?, geaendert_am=datetime('now','localtime')
                   WHERE art_id=? AND groesse=?""",
                (neue_menge, art_id, groesse),
            )
            conn.execute(
                """INSERT INTO buchungen
                   (art_id, art_name, groesse, menge, typ, datum, bemerkung)
                   VALUES (?, ?, ?, ?, 'korrektur', ?, ?)""",
                (art_id, art_name, groesse, delta, datum, f"Korrektur: {alte_menge}→{neue_menge}. {bemerkung}"),
            )
            conn.commit()
            return True, f"Bestand korrigiert: {alte_menge} → {neue_menge}"
        except Exception as e:
            conn.rollback()
            return False, f"Fehler bei Korrektur: {e}"
        finally:
            conn.close()

    # ------------------------------------------------------------------
    # Mitarbeiter-Kleidung
    # ------------------------------------------------------------------

    def get_mitarbeiter_kleidung(
        self,
        mitarbeiter_id: Optional[int] = None,
        mitarbeiter_name: Optional[str] = None,
        status: str = "ausgegeben",
    ) -> list[dict]:
        conn = self._conn_kl()
        if mitarbeiter_id is not None:
            cur = conn.execute(
                """SELECT id, mitarbeiter_id, mitarbeiter_name, art_id, art_name,
                          groesse, menge, ausgabe_datum, rueckgabe_datum, status,
                          ausgegeben_von, bemerkung
                   FROM mitarbeiter_kleidung
                   WHERE mitarbeiter_id=? AND status=?
                   ORDER BY art_name, groesse""",
                (mitarbeiter_id, status),
            )
        elif mitarbeiter_name:
            cur = conn.execute(
                """SELECT id, mitarbeiter_id, mitarbeiter_name, art_id, art_name,
                          groesse, menge, ausgabe_datum, rueckgabe_datum, status,
                          ausgegeben_von, bemerkung
                   FROM mitarbeiter_kleidung
                   WHERE mitarbeiter_name LIKE ? AND status=?
                   ORDER BY art_name, groesse""",
                (f"%{mitarbeiter_name}%", status),
            )
        else:
            cur = conn.execute(
                """SELECT id, mitarbeiter_id, mitarbeiter_name, art_id, art_name,
                          groesse, menge, ausgabe_datum, rueckgabe_datum, status,
                          ausgegeben_von, bemerkung
                   FROM mitarbeiter_kleidung
                   WHERE status=?
                   ORDER BY mitarbeiter_name, art_name, groesse""",
                (status,),
            )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def get_mitarbeiter_mit_kleidung(self) -> list[dict]:
        """Alle Mitarbeiter, denen aktuell Kleidung zugeordnet ist."""
        conn = self._conn_kl()
        cur = conn.execute(
            """SELECT mitarbeiter_id, mitarbeiter_name,
                      COUNT(*) AS anzahl_positionen,
                      SUM(menge) AS gesamt_stueck
               FROM mitarbeiter_kleidung
               WHERE status='ausgegeben'
               GROUP BY mitarbeiter_id, mitarbeiter_name
               ORDER BY mitarbeiter_name"""
        )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    # ------------------------------------------------------------------
    # Buchungen / Verlauf
    # ------------------------------------------------------------------

    def get_buchungen(
        self,
        limit: int = 200,
        offset: int = 0,
        art_id: Optional[int] = None,
        typ: Optional[str] = None,
        datum_von: Optional[str] = None,
        datum_bis: Optional[str] = None,
        suche: Optional[str] = None,
    ) -> list[dict]:
        query = (
            "SELECT id, art_name, groesse, menge, typ, mitarbeiter_name, "
            "datum, ausgegeben_von, bemerkung, erstellt_am "
            "FROM buchungen WHERE 1=1"
        )
        params: list = []
        if art_id:
            query += " AND art_id=?"
            params.append(art_id)
        if typ:
            query += " AND typ=?"
            params.append(typ)
        if datum_von:
            query += " AND datum >= ?"
            params.append(datum_von)
        if datum_bis:
            query += " AND datum <= ?"
            params.append(datum_bis)
        if suche:
            query += " AND (mitarbeiter_name LIKE ? OR art_name LIKE ? OR groesse LIKE ?)"
            s = f"%{suche}%"
            params.extend([s, s, s])
        query += " ORDER BY datum DESC, erstellt_am DESC LIMIT ? OFFSET ?"
        params.extend([limit, offset])

        conn = self._conn_kl()
        cur = conn.execute(query, params)
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows

    def get_buchungen_count(
        self,
        art_id: Optional[int] = None,
        typ: Optional[str] = None,
        datum_von: Optional[str] = None,
        datum_bis: Optional[str] = None,
        suche: Optional[str] = None,
    ) -> int:
        query = "SELECT COUNT(*) FROM buchungen WHERE 1=1"
        params: list = []
        if art_id:
            query += " AND art_id=?"
            params.append(art_id)
        if typ:
            query += " AND typ=?"
            params.append(typ)
        if datum_von:
            query += " AND datum >= ?"
            params.append(datum_von)
        if datum_bis:
            query += " AND datum <= ?"
            params.append(datum_bis)
        if suche:
            query += " AND (mitarbeiter_name LIKE ? OR art_name LIKE ? OR groesse LIKE ?)"
            s = f"%{suche}%"
            params.extend([s, s, s])
        conn = self._conn_kl()
        cur = conn.execute(query, params)
        count = cur.fetchone()[0]
        conn.close()
        return count

    # ------------------------------------------------------------------
    # Buchungs-Editierung
    # ------------------------------------------------------------------

    def update_buchung(self, buchung_id: int, datum: str, bemerkung: str,
                       ausgegeben_von: str) -> tuple[bool, str]:
        """Aktualisiert editierbare Felder einer Buchung (kein Bestandsausgleich)."""
        try:
            conn = self._conn_kl()
            conn.execute(
                "UPDATE buchungen SET datum=?, bemerkung=?, ausgegeben_von=? WHERE id=?",
                (datum, bemerkung, ausgegeben_von, buchung_id),
            )
            conn.commit()
            conn.close()
            return True, "Buchung aktualisiert."
        except Exception as e:
            return False, str(e)

    def delete_buchung(self, buchung_id: int) -> tuple[bool, str]:
        """Löscht einen Buchungseintrag. Bestand wird NICHT automatisch korrigiert."""
        try:
            conn = self._conn_kl()
            conn.execute("DELETE FROM buchungen WHERE id=?", (buchung_id,))
            conn.commit()
            conn.close()
            return True, "Buchung gelöscht."
        except Exception as e:
            return False, str(e)

    # ------------------------------------------------------------------
    # Dashboard-Statistiken
    # ------------------------------------------------------------------

    def get_dashboard_stats(self) -> dict:
        conn = self._conn_kl()
        total_items = conn.execute(
            "SELECT COALESCE(SUM(menge),0) FROM kleidungsbestand"
        ).fetchone()[0]
        total_types = conn.execute(
            "SELECT COUNT(*) FROM kleidungsarten WHERE aktiv=1"
        ).fetchone()[0]
        employees_with_clothing = conn.execute(
            "SELECT COUNT(DISTINCT COALESCE(CAST(mitarbeiter_id AS TEXT), mitarbeiter_name)) "
            "FROM mitarbeiter_kleidung WHERE status='ausgegeben'"
        ).fetchone()[0]
        low_stock = conn.execute(
            "SELECT COUNT(*) FROM kleidungsbestand WHERE menge <= min_menge AND min_menge > 0"
        ).fetchone()[0]
        recent_activity = conn.execute(
            "SELECT COUNT(*) FROM buchungen WHERE datum >= date('now','-7 days','localtime')"
        ).fetchone()[0]
        today_ausgabe = conn.execute(
            "SELECT COUNT(*) FROM buchungen WHERE typ='ausgabe' AND datum=date('now','localtime')"
        ).fetchone()[0]
        conn.close()
        return {
            "total_items": total_items,
            "total_types": total_types,
            "employees_with_clothing": employees_with_clothing,
            "low_stock": low_stock,
            "recent_activity": recent_activity,
            "today_ausgabe": today_ausgabe,
        }

    def get_recent_buchungen(self, limit: int = 10) -> list[dict]:
        conn = self._conn_kl()
        cur = conn.execute(
            """SELECT art_name, groesse, menge, typ, mitarbeiter_name, datum, ausgegeben_von
               FROM buchungen ORDER BY erstellt_am DESC LIMIT ?""",
            (limit,),
        )
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        return rows
