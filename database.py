from __future__ import annotations

import os
import sqlite3
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

from schema import INDICATOR_FIELDS, REPORT_COLUMNS
from validators import parse_summary_workbook, validate_reports, to_int, to_text, normalize_text

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "banana_smartlink.db"
DRIVE_SUMMARY_PATH = BASE_DIR / "sample_data" / "drive_imported" / "TONG_HOP_va_THEO_DOI_TIEN_DO.xlsx"


def get_connection(db_path: Path | str = DB_PATH) -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    return conn


def init_db(db_path: Path | str = DB_PATH) -> None:
    conn = get_connection(db_path)
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS villages (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            village_name TEXT NOT NULL UNIQUE,
            commune_name TEXT NOT NULL DEFAULT 'Xã Bà Nà',
            status TEXT DEFAULT 'active',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    indicator_cols = ",\n".join(f"{field} INTEGER" for field in INDICATOR_FIELDS)
    cur.execute(
        f"""
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            village_id INTEGER,
            commune_name TEXT NOT NULL,
            village_name TEXT NOT NULL,
            period TEXT NOT NULL,
            report_date TEXT,
            reporter_name TEXT,
            reporter_title TEXT,
            phone TEXT,
            due_at TEXT,
            submitted_at TEXT,
            submission_status TEXT,
            days_late INTEGER,
            {indicator_cols},
            note TEXT,
            source_file TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(village_name, period),
            FOREIGN KEY(village_id) REFERENCES villages(id)
        )
        """
    )

    # Migrate older SQLite databases that were created before CT01-CT14
    # were fully added to the reports table. CREATE TABLE IF NOT EXISTS
    # does not add new columns to an existing table, so we add any missing
    # indicator columns explicitly and preserve the current data.
    existing_report_columns = {
        row["name"] for row in cur.execute("PRAGMA table_info(reports)").fetchall()
    }
    for field in INDICATOR_FIELDS:
        if field not in existing_report_columns:
            cur.execute(f"ALTER TABLE reports ADD COLUMN {field} INTEGER")


    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            village_id INTEGER,
            village_name TEXT NOT NULL,
            task_name TEXT NOT NULL,
            period TEXT NOT NULL,
            due_at TEXT,
            submitted_at TEXT,
            status TEXT,
            days_late INTEGER,
            reminder_status TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(village_name, task_name, period),
            FOREIGN KEY(village_id) REFERENCES villages(id)
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS validation_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            village_name TEXT,
            period TEXT,
            severity TEXT NOT NULL,
            error_type TEXT,
            field_name TEXT,
            message TEXT NOT NULL,
            source_file TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS app_meta (
            key TEXT PRIMARY KEY,
            value TEXT
        )
        """
    )

    conn.commit()
    conn.close()


def get_or_create_village(conn: sqlite3.Connection, village_name: str, commune_name: str = "Xã Bà Nà") -> int:
    village_name = to_text(village_name)
    commune_name = to_text(commune_name) or "Xã Bà Nà"
    cur = conn.cursor()
    cur.execute("SELECT id FROM villages WHERE village_name = ?", (village_name,))
    row = cur.fetchone()
    if row:
        cur.execute("UPDATE villages SET commune_name = ? WHERE id = ?", (commune_name, row["id"]))
        return int(row["id"])
    cur.execute(
        "INSERT INTO villages(village_name, commune_name, status) VALUES (?, ?, 'active')",
        (village_name, commune_name),
    )
    return int(cur.lastrowid)


def _db_value(value: Any) -> Any:
    if value is None:
        return None
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    if str(value) == "nan":
        return None
    return value


def upsert_reports(rows: Iterable[Dict[str, Any]], source_file: str = "") -> List[int]:
    conn = get_connection()
    cur = conn.cursor()
    saved_ids: List[int] = []

    for row in rows:
        row = dict(row)
        status_norm = normalize_text(row.get("submission_status"))
        if "chua_nop" in status_norm:
            # Not-submitted rows still become villages/tasks but not reports.
            village_id = get_or_create_village(conn, row.get("village_name"), row.get("commune_name") or "Xã Bà Nà")
            upsert_task_from_row(conn, row, village_id)
            continue

        village_name = to_text(row.get("village_name"))
        if not village_name:
            continue
        commune_name = to_text(row.get("commune_name")) or "Xã Bà Nà"
        period = to_text(row.get("period")) or "2026-Q2"
        village_id = get_or_create_village(conn, village_name, commune_name)
        row["source_file"] = source_file or row.get("source_file") or ""

        columns = [c for c in REPORT_COLUMNS if c in row]
        db_columns = ["village_id", *columns]
        values = [village_id, *[_db_value(row.get(c)) for c in columns]]
        placeholders = ", ".join("?" for _ in db_columns)
        update_cols = [c for c in db_columns if c not in {"village_id", "village_name", "period"}]
        update_expr = ", ".join(f"{c}=excluded.{c}" for c in update_cols)

        sql = f"""
            INSERT INTO reports({', '.join(db_columns)})
            VALUES ({placeholders})
            ON CONFLICT(village_name, period)
            DO UPDATE SET {update_expr}, created_at=CURRENT_TIMESTAMP
        """
        cur.execute(sql, values)
        cur.execute("SELECT id FROM reports WHERE village_name = ? AND period = ?", (village_name, period))
        saved = cur.fetchone()
        if saved:
            saved_ids.append(int(saved["id"]))
        upsert_task_from_row(conn, row, village_id)

    conn.commit()
    conn.close()
    return saved_ids


def upsert_task_from_row(conn: sqlite3.Connection, row: Dict[str, Any], village_id: int) -> None:
    village_name = to_text(row.get("village_name"))
    period = to_text(row.get("period")) or "2026-Q2"
    status = to_text(row.get("submission_status")) or "Đã nộp"
    status_norm = normalize_text(status)
    if "chua_nop" in status_norm:
        task_status = "Chưa nộp"
        reminder = "Quá hạn"
    elif "tre_han" in status_norm:
        task_status = "Trễ hạn"
        reminder = "Đã nộp trễ"
    else:
        task_status = "Hoàn thành"
        reminder = "Không cần"
    conn.execute(
        """
        INSERT INTO tasks(village_id, village_name, task_name, period, due_at, submitted_at, status, days_late, reminder_status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(village_name, task_name, period)
        DO UPDATE SET
            village_id=excluded.village_id,
            due_at=excluded.due_at,
            submitted_at=excluded.submitted_at,
            status=excluded.status,
            days_late=excluded.days_late,
            reminder_status=excluded.reminder_status,
            created_at=CURRENT_TIMESTAMP
        """,
        (
            village_id,
            village_name,
            "Báo cáo văn hóa - xã hội Quý II/2026",
            period,
            to_text(row.get("due_at")) or "2026-06-15 17:00",
            to_text(row.get("submitted_at")),
            task_status,
            to_int(row.get("days_late")) or 0,
            reminder,
        ),
    )


def insert_validation_logs(logs: Iterable[Dict[str, Any]], source_file: str = "") -> None:
    conn = get_connection()
    cur = conn.cursor()
    for log in logs:
        cur.execute(
            """
            INSERT INTO validation_logs(village_name, period, severity, error_type, field_name, message, source_file)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                log.get("village_name"),
                log.get("period"),
                log.get("severity", "WARNING"),
                log.get("error_type"),
                log.get("field_name"),
                log.get("message"),
                source_file or log.get("source_file") or "",
            ),
        )
    conn.commit()
    conn.close()


def seed_demo_data(force: bool = False) -> None:
    init_db()
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT value FROM app_meta WHERE key='seeded_drive_data'")
    already = cur.fetchone()
    conn.close()
    if already and not force:
        return

    if force and DB_PATH.exists():
        DB_PATH.unlink()
        init_db()

    if DRIVE_SUMMARY_PATH.exists():
        reports, _ = parse_summary_workbook(DRIVE_SUMMARY_PATH, DRIVE_SUMMARY_PATH.name)
        clean, issues = validate_reports(reports, allow_not_submitted_rows=True)
        upsert_reports(clean.to_dict(orient="records"), source_file=DRIVE_SUMMARY_PATH.name)
        insert_validation_logs(issues, source_file=DRIVE_SUMMARY_PATH.name)
    else:
        seed_fallback_data()

    conn = get_connection()
    conn.execute(
        "INSERT OR REPLACE INTO app_meta(key, value) VALUES ('seeded_drive_data', 'true')"
    )
    conn.commit()
    conn.close()


def seed_fallback_data() -> None:
    rows = []
    for i in range(1, 11):
        rows.append(
            {
                "commune_name": "Xã Bà Nà",
                "village_name": f"Thôn {i}",
                "period": "2026-Q2",
                "report_date": "2026-06-19",
                "submission_status": "Đúng hạn" if i < 9 else "Chưa nộp",
                "ct01_households": 200 + i * 10,
                "ct02_population": 800 + i * 30,
                "ct03_poor_households": 10 + i,
                "ct04_near_poor_households": 8 + i,
                "source_file": "fallback_seed",
            }
        )
    clean, issues = validate_reports(__import__("pandas").DataFrame(rows), allow_not_submitted_rows=True)
    upsert_reports(clean.to_dict(orient="records"), "fallback_seed")
    insert_validation_logs(issues, "fallback_seed")


def fetch_reports(period: Optional[str] = None) -> List[Dict[str, Any]]:
    conn = get_connection()
    query = "SELECT * FROM reports"
    params: List[Any] = []
    if period:
        query += " WHERE period = ?"
        params.append(period)
    query += " ORDER BY period DESC, village_name ASC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def fetch_villages() -> List[Dict[str, Any]]:
    conn = get_connection()
    rows = conn.execute("SELECT * FROM villages ORDER BY village_name ASC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def fetch_tasks(period: Optional[str] = None) -> List[Dict[str, Any]]:
    conn = get_connection()
    query = "SELECT * FROM tasks"
    params: List[Any] = []
    if period:
        query += " WHERE period = ?"
        params.append(period)
    query += " ORDER BY status DESC, village_name ASC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def fetch_validation_logs(period: Optional[str] = None) -> List[Dict[str, Any]]:
    conn = get_connection()
    query = "SELECT * FROM validation_logs"
    params: List[Any] = []
    if period:
        query += " WHERE period = ?"
        params.append(period)
    query += " ORDER BY severity ASC, village_name ASC, created_at DESC"
    rows = conn.execute(query, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def list_periods() -> List[str]:
    conn = get_connection()
    rows = conn.execute(
        """
        SELECT period FROM reports
        UNION
        SELECT period FROM tasks
        ORDER BY period DESC
        """
    ).fetchall()
    conn.close()
    periods = [r["period"] for r in rows if r["period"]]
    return periods or ["2026-Q2"]


def clear_validation_logs(period: Optional[str] = None) -> None:
    conn = get_connection()
    if period:
        conn.execute("DELETE FROM validation_logs WHERE period = ?", (period,))
    else:
        conn.execute("DELETE FROM validation_logs")
    conn.commit()
    conn.close()
