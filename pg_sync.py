"""
Hybrid SQLite/PostgreSQL sync engine.

Design:
- The app ALWAYS writes to the local SQLite database first (attendance.db,
  via database.py). Nothing about local read/write behavior changes here.
- Every staff/attendance row carries a `synced` flag (0 = pending, 1 = pushed
  to the server). insert_staff/update_staff and the check-in/out writers in
  main.py set synced=0 whenever they touch a row.
- This module is only responsible for pushing synced=0 rows up to a
  PostgreSQL server on demand ("Synchronize Now") or automatically right
  after a successful Connect, and flipping them to synced=1 once pushed.
- Connection settings live in server_config.json next to the app - NOT inside
  attendance.db/rohi_attendance.db - so a "Reset Database" (which only drops
  the local staff/attendance tables) never wipes out the server config.
- Uses pg8000 (pure-Python Postgres driver) instead of psycopg2. psycopg2's C
  extension needs a matching prebuilt wheel or a C toolchain to cross-compile
  for Android via buildozer/python-for-android, which is a common source of
  broken builds; pg8000 has no C dependency and packages cleanly.
- If pg8000 isn't installed, every function below degrades to a clear
  (ok=False, message=...) result instead of raising, so the rest of the app
  never crashes because of this optional feature.
"""

import os
import json
import logging
from datetime import datetime

logger = logging.getLogger("ROHIApp.pg_sync")

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "server_config.json")
LOCAL_DB_PATH = os.path.join(APP_DIR, "attendance.db")

try:
    import pg8000.native as pg8000_native
    PG8000_AVAILABLE = True
except ImportError:
    pg8000_native = None
    PG8000_AVAILABLE = False

# Module-level connection state. Simple by design: this app has a single
# active server connection at a time, controlled from one Settings screen.
_connection = None
_connected_config = None  # the config dict that produced _connection, if any


# -------------------------------------------------------------
# Config persistence (server_config.json)
# -------------------------------------------------------------
DEFAULT_CONFIG = {
    "server_name": "",
    "host": "",
    "port": 5432,
    "dbname": "",
    "username": "",
    "password": "",
    "sslmode": "prefer",   # disable | allow | prefer | require
    "last_sync_time": None,
}


def load_config():
    """Returns the saved server config dict, or a copy of DEFAULT_CONFIG if none saved yet."""
    if not os.path.exists(CONFIG_PATH):
        return dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        merged = dict(DEFAULT_CONFIG)
        merged.update(data)
        return merged
    except Exception:
        logger.exception("Failed to read server_config.json; using defaults.")
        return dict(DEFAULT_CONFIG)


def save_config(config: dict):
    """Persists connection settings. Does not touch last_sync_time unless present in config."""
    try:
        existing = load_config()
        existing.update(config)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(existing, f, indent=2)
        return True, "Settings saved."
    except Exception as e:
        logger.exception("Failed to save server_config.json:")
        return False, f"Could not save settings: {e}"


def _set_last_sync_time(dt=None):
    dt = dt or datetime.now()
    save_config({"last_sync_time": dt.strftime("%Y-%m-%d %H:%M:%S")})


# -------------------------------------------------------------
# Connection lifecycle
# -------------------------------------------------------------
def _open_connection(config):
    """Low-level pg8000 connect. Raises on failure; callers catch."""
    return pg8000_native.Connection(
        user=config.get("username") or "",
        password=config.get("password") or "",
        host=config.get("host") or "",
        port=int(config.get("port") or 5432),
        database=config.get("dbname") or "",
        ssl_context=True if (config.get("sslmode") in ("require", "prefer")) else None,
    )


def test_connection(config):
    """Opens a throwaway connection to verify the given settings work, then closes it."""
    if not PG8000_AVAILABLE:
        return False, "PostgreSQL driver (pg8000) is not installed in this build."
    try:
        conn = _open_connection(config)
        conn.run("SELECT 1;")
        conn.close()
        return True, "Connection successful."
    except Exception as e:
        logger.exception("test_connection failed:")
        return False, f"Connection failed: {e}"


def connect(config, auto_sync=True):
    """Opens and keeps the live connection used by Synchronize Now / status display.

    When auto_sync=True (the default - used for both the Connect button and any
    future auto-reconnect), a successful connect immediately triggers a sync of
    whatever local rows are pending, so the server never sits stale after a
    manual reconnect.
    """
    global _connection, _connected_config
    if not PG8000_AVAILABLE:
        return False, "PostgreSQL driver (pg8000) is not installed in this build."
    try:
        _connection = _open_connection(config)
        _connected_config = dict(config)
        _ensure_remote_tables(_connection)
        logger.info("Connected to PostgreSQL server '%s'.", config.get("server_name") or config.get("host"))
        message = "Connected."
        if auto_sync:
            ok, sync_message, _counts = synchronize_now(config)
            message = "Connected. " + sync_message
        return True, message
    except Exception as e:
        _connection = None
        _connected_config = None
        logger.exception("connect() failed:")
        return False, f"Could not connect: {e}"


def disconnect():
    global _connection, _connected_config
    if _connection is not None:
        try:
            _connection.close()
        except Exception:
            pass
    _connection = None
    _connected_config = None
    return True, "Disconnected."


def is_connected():
    return _connection is not None


def get_status():
    """Returns a small dict for the Settings/Server Connection status card."""
    config = load_config()
    return {
        "connected": is_connected(),
        "server_name": config.get("server_name") or config.get("host") or "Not configured",
        "last_sync_time": config.get("last_sync_time") or "Never",
        "pending_count": get_pending_count(),
    }


# -------------------------------------------------------------
# Remote schema
# -------------------------------------------------------------
def _ensure_remote_tables(conn):
    """Creates minimal mirror tables on the server if they don't exist yet.
    Kept intentionally narrow (unique_id / email as natural keys) - this is a
    sync target, not the system of record."""
    conn.run("""
        CREATE TABLE IF NOT EXISTS staff (
            id SERIAL PRIMARY KEY,
            unique_id TEXT UNIQUE,
            fullname TEXT, sex TEXT, dob TEXT, blood_group TEXT, marital_status TEXT,
            nationality TEXT, state_origin TEXT, lga TEXT, address TEXT, next_of_kin TEXT,
            next_of_kin_phone TEXT, employment_type TEXT, state_office TEXT, cluster TEXT,
            department TEXT, section TEXT, position TEXT, staff_number TEXT, phone TEXT,
            email TEXT UNIQUE, facebook TEXT, twitter TEXT, instagram TEXT, telegram TEXT,
            linkedin TEXT, gps_coordinate TEXT, photo TEXT,
            local_id INTEGER,
            synced_at TIMESTAMP
        );
    """)
    conn.run("""
        CREATE TABLE IF NOT EXISTS attendance (
            id SERIAL PRIMARY KEY,
            local_id INTEGER UNIQUE,
            email TEXT,
            check_in_time TEXT,
            check_out_time TEXT,
            late_duration TEXT,
            attendance_status TEXT,
            gps_location TEXT,
            check_out_gps_location TEXT,
            synced_at TIMESTAMP
        );
    """)


# -------------------------------------------------------------
# Pending-row helpers (local SQLite side)
# -------------------------------------------------------------
def _local_conn():
    import sqlite3
    return sqlite3.connect(LOCAL_DB_PATH)


def get_pending_count():
    """Total unsynced rows across staff + attendance. Used by the Settings badge."""
    try:
        conn = _local_conn()
        cursor = conn.cursor()
        total = 0
        for table in ("staff", "attendance"):
            try:
                cursor.execute(f"SELECT COUNT(*) FROM {table} WHERE synced = 0 OR synced IS NULL")
                total += cursor.fetchone()[0]
            except Exception:
                pass  # table/column may not exist yet on a brand-new install
        conn.close()
        return total
    except Exception:
        logger.exception("get_pending_count failed:")
        return 0


def _fetch_pending(table, columns):
    conn = _local_conn()
    conn.row_factory = None
    cursor = conn.cursor()
    cursor.execute(f"SELECT {', '.join(columns)} FROM {table} WHERE synced = 0 OR synced IS NULL")
    rows = cursor.fetchall()
    conn.close()
    return rows


def _mark_synced(table, ids):
    if not ids:
        return
    conn = _local_conn()
    cursor = conn.cursor()
    cursor.executemany(f"UPDATE {table} SET synced = 1 WHERE id = ?", [(i,) for i in ids])
    conn.commit()
    conn.close()


STAFF_COLUMNS = [
    "id", "unique_id", "fullname", "sex", "dob", "blood_group", "marital_status",
    "nationality", "state_origin", "lga", "address", "next_of_kin", "next_of_kin_phone",
    "employment_type", "state_office", "cluster", "department", "section", "position",
    "staff_number", "phone", "email", "facebook", "twitter", "instagram", "telegram",
    "linkedin", "gps_coordinate", "photo",
]

ATTENDANCE_COLUMNS = [
    "id", "email", "check_in_time", "check_out_time", "late_duration",
    "attendance_status", "gps_location", "check_out_gps_location",
]


def _push_staff(conn, rows):
    pushed_ids = []
    for row in rows:
        data = dict(zip(STAFF_COLUMNS, row))
        local_id = data.pop("id")
        data["local_id"] = local_id
        try:
            conn.run(
                """
                INSERT INTO staff (
                    unique_id, fullname, sex, dob, blood_group, marital_status,
                    nationality, state_origin, lga, address, next_of_kin, next_of_kin_phone,
                    employment_type, state_office, cluster, department, section, position,
                    staff_number, phone, email, facebook, twitter, instagram, telegram,
                    linkedin, gps_coordinate, photo, local_id, synced_at
                ) VALUES (
                    :unique_id, :fullname, :sex, :dob, :blood_group, :marital_status,
                    :nationality, :state_origin, :lga, :address, :next_of_kin, :next_of_kin_phone,
                    :employment_type, :state_office, :cluster, :department, :section, :position,
                    :staff_number, :phone, :email, :facebook, :twitter, :instagram, :telegram,
                    :linkedin, :gps_coordinate, :photo, :local_id, now()
                )
                ON CONFLICT (email) DO UPDATE SET
                    unique_id = EXCLUDED.unique_id, fullname = EXCLUDED.fullname, sex = EXCLUDED.sex,
                    dob = EXCLUDED.dob, blood_group = EXCLUDED.blood_group,
                    marital_status = EXCLUDED.marital_status, nationality = EXCLUDED.nationality,
                    state_origin = EXCLUDED.state_origin, lga = EXCLUDED.lga, address = EXCLUDED.address,
                    next_of_kin = EXCLUDED.next_of_kin, next_of_kin_phone = EXCLUDED.next_of_kin_phone,
                    employment_type = EXCLUDED.employment_type, state_office = EXCLUDED.state_office,
                    cluster = EXCLUDED.cluster, department = EXCLUDED.department, section = EXCLUDED.section,
                    position = EXCLUDED.position, staff_number = EXCLUDED.staff_number, phone = EXCLUDED.phone,
                    facebook = EXCLUDED.facebook, twitter = EXCLUDED.twitter, instagram = EXCLUDED.instagram,
                    telegram = EXCLUDED.telegram, linkedin = EXCLUDED.linkedin,
                    gps_coordinate = EXCLUDED.gps_coordinate, photo = EXCLUDED.photo,
                    local_id = EXCLUDED.local_id, synced_at = now();
                """,
                **data,
            )
            pushed_ids.append(local_id)
        except Exception:
            logger.exception("Failed to push staff row local_id=%s:", local_id)
    return pushed_ids


def _push_attendance(conn, rows):
    pushed_ids = []
    for row in rows:
        data = dict(zip(ATTENDANCE_COLUMNS, row))
        local_id = data.pop("id")
        data["local_id"] = local_id
        try:
            conn.run(
                """
                INSERT INTO attendance (
                    local_id, email, check_in_time, check_out_time, late_duration,
                    attendance_status, gps_location, check_out_gps_location, synced_at
                ) VALUES (
                    :local_id, :email, :check_in_time, :check_out_time, :late_duration,
                    :attendance_status, :gps_location, :check_out_gps_location, now()
                )
                ON CONFLICT (local_id) DO UPDATE SET
                    email = EXCLUDED.email, check_in_time = EXCLUDED.check_in_time,
                    check_out_time = EXCLUDED.check_out_time, late_duration = EXCLUDED.late_duration,
                    attendance_status = EXCLUDED.attendance_status, gps_location = EXCLUDED.gps_location,
                    check_out_gps_location = EXCLUDED.check_out_gps_location, synced_at = now();
                """,
                **data,
            )
            pushed_ids.append(local_id)
        except Exception:
            logger.exception("Failed to push attendance row local_id=%s:", local_id)
    return pushed_ids


def synchronize_now(config=None):
    """Pushes every synced=0 local staff/attendance row to PostgreSQL.

    Uses the live connection if one is open; otherwise opens a short-lived one
    with the given (or saved) config. Always closes short-lived connections.
    Returns (ok: bool, message: str, counts: dict).
    """
    global _connection
    if not PG8000_AVAILABLE:
        return False, "PostgreSQL driver (pg8000) is not installed in this build.", {}

    config = config or load_config()
    own_connection = False
    conn = _connection
    try:
        if conn is None:
            conn = _open_connection(config)
            _ensure_remote_tables(conn)
            own_connection = True

        staff_rows = _fetch_pending("staff", STAFF_COLUMNS)
        attendance_rows = _fetch_pending("attendance", ATTENDANCE_COLUMNS)

        staff_pushed = _push_staff(conn, staff_rows)
        attendance_pushed = _push_attendance(conn, attendance_rows)

        _mark_synced("staff", staff_pushed)
        _mark_synced("attendance", attendance_pushed)

        _set_last_sync_time()

        counts = {"staff": len(staff_pushed), "attendance": len(attendance_pushed)}
        message = f"Synced {counts['staff']} staff record(s) and {counts['attendance']} attendance record(s)."
        logger.info(message)
        return True, message, counts
    except Exception as e:
        logger.exception("synchronize_now failed:")
        return False, f"Sync failed: {e}", {}
    finally:
        if own_connection and conn is not None:
            try:
                conn.close()
            except Exception:
                pass
