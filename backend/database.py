"""
database.py — Base de datos SQLite local para tracking de bitácoras.

Almacena el estado de cada bitácora (pendiente, revisada, rechazada)
y permite que el tío marque su progreso desde el dashboard.
"""

import sqlite3
from pathlib import Path
from datetime import datetime
from typing import Optional

DB_PATH = Path(__file__).parent.parent / "data" / "autouncle.db"


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Crea las tablas si no existen."""
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with get_conn() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS bitacoras (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,

            -- Identificación
            ficha           TEXT NOT NULL,
            cc              TEXT,
            apprentice_name TEXT NOT NULL,
            program         TEXT,
            area            TEXT,
            company         TEXT,
            start_date      TEXT,
            end_date        TEXT,

            -- Archivo en OneDrive
            filename        TEXT NOT NULL,
            file_path       TEXT NOT NULL UNIQUE,
            file_ext        TEXT,
            bitacora_num    INTEGER,
            modified        TEXT,

            -- Seguimientos programados
            seg1_date       TEXT,
            seg2_date       TEXT,
            seg3_date       TEXT,

            -- Estado de revisión
            status          TEXT NOT NULL DEFAULT 'pending',
            -- 'pending'  → no revisada aún
            -- 'approved' → revisada y firmada
            -- 'rejected' → devuelta al aprendiz con observaciones
            -- 'skipped'  → ya estaba revisada antes de usar la app

            notes           TEXT,
            reviewed_at     TEXT,
            signed_pdf_path TEXT,   -- ruta en SharePoint del PDF firmado

            -- Control
            first_seen      TEXT NOT NULL DEFAULT (datetime('now')),
            last_synced     TEXT NOT NULL DEFAULT (datetime('now'))
        );

        CREATE TABLE IF NOT EXISTS sync_log (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            synced_at   TEXT NOT NULL DEFAULT (datetime('now')),
            new_found   INTEGER,
            total_found INTEGER
        );
        """)
    print(f"✓ Base de datos inicializada en {DB_PATH}")


def upsert_bitacoras(bitacoras: list) -> tuple[int, int]:
    """
    Inserta bitácoras nuevas. Si ya existe (por file_path), actualiza
    solo los metadatos pero NO toca el status (para no perder el progreso).
    Retorna (nuevas, actualizadas).
    """
    new_count = 0
    updated_count = 0

    with get_conn() as conn:
        for b in bitacoras:
            existing = conn.execute(
                "SELECT id, status FROM bitacoras WHERE file_path = ?", (b.file_path,)
            ).fetchone()

            if existing:
                conn.execute("""
                    UPDATE bitacoras SET
                        apprentice_name = ?, ficha = ?, cc = ?, program = ?,
                        area = ?, company = ?, start_date = ?, end_date = ?,
                        filename = ?, file_ext = ?, bitacora_num = ?, modified = ?,
                        seg1_date = ?, seg2_date = ?, seg3_date = ?,
                        last_synced = datetime('now')
                    WHERE file_path = ?
                """, (
                    b.apprentice_name, b.ficha, b.cc, b.program,
                    b.area, b.company, b.start_date, b.end_date,
                    b.filename, b.file_ext, b.bitacora_num, b.modified,
                    b.seg1_date, b.seg2_date, b.seg3_date,
                    b.file_path
                ))
                updated_count += 1
            else:
                conn.execute("""
                    INSERT INTO bitacoras (
                        ficha, cc, apprentice_name, program, area, company,
                        start_date, end_date, filename, file_path, file_ext,
                        bitacora_num, modified, seg1_date, seg2_date, seg3_date,
                        status
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,'pending')
                """, (
                    b.ficha, b.cc, b.apprentice_name, b.program, b.area, b.company,
                    b.start_date, b.end_date, b.filename, b.file_path, b.file_ext,
                    b.bitacora_num, b.modified, b.seg1_date, b.seg2_date, b.seg3_date
                ))
                new_count += 1

        conn.execute(
            "INSERT INTO sync_log (new_found, total_found) VALUES (?, ?)",
            (new_count, new_count + updated_count)
        )

    return new_count, updated_count


def get_pending() -> list[sqlite3.Row]:
    """Retorna todas las bitácoras pendientes, ordenadas por ficha y aprendiz."""
    with get_conn() as conn:
        return conn.execute("""
            SELECT * FROM bitacoras
            WHERE status = 'pending'
            ORDER BY ficha, apprentice_name, bitacora_num
        """).fetchall()


def get_all() -> list[sqlite3.Row]:
    """Retorna todas las bitácoras."""
    with get_conn() as conn:
        return conn.execute("""
            SELECT * FROM bitacoras
            ORDER BY status, ficha, apprentice_name, bitacora_num
        """).fetchall()


def update_status(bitacora_id: int, status: str, notes: str = "") -> bool:
    """Actualiza el estado de una bitácora."""
    valid = ("pending", "approved", "rejected", "skipped")
    if status not in valid:
        return False
    with get_conn() as conn:
        conn.execute("""
            UPDATE bitacoras
            SET status = ?, notes = ?, reviewed_at = datetime('now')
            WHERE id = ?
        """, (status, notes, bitacora_id))
    return True


def get_stats() -> dict:
    """Retorna estadísticas del estado actual."""
    with get_conn() as conn:
        rows = conn.execute("""
            SELECT status, COUNT(*) as count FROM bitacoras GROUP BY status
        """).fetchall()
    return {r["status"]: r["count"] for r in rows}


if __name__ == "__main__":
    init_db()
    stats = get_stats()
    print("Stats:", stats)
