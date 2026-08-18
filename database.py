import os
import sqlite3
import time
import logging

DB_NAME = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attendance.db")
DB_TIMEOUT = 15.0
logger = logging.getLogger("ROHIApp")


def _connect():
    """Open SQLite with Android-safe timeout and WAL/busy settings."""
    conn = sqlite3.connect(DB_NAME, timeout=DB_TIMEOUT)
    conn.execute("PRAGMA busy_timeout=15000")
    try:
        conn.execute("PRAGMA journal_mode=WAL")
    except sqlite3.DatabaseError:
        pass
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def _is_locked_error(exc):
    return isinstance(exc, sqlite3.OperationalError) and "locked" in str(exc).lower()


def _run_write(operation, retries=4):
    """Run a write transaction and retry transient Android SQLite locks."""
    last_error = None
    for attempt in range(retries):
        conn = None
        try:
            conn = _connect()
            cursor = conn.cursor()
            result = operation(cursor)
            conn.commit()
            return result
        except sqlite3.OperationalError as exc:
            last_error = exc
            if _is_locked_error(exc) and attempt < retries - 1:
                if conn:
                    try:
                        conn.rollback()
                    except Exception:
                        pass
                time.sleep(0.25 * (attempt + 1))
                continue
            raise
        finally:
            if conn:
                conn.close()
    raise last_error


def create_table():
    """Create/migrate all local tables safely."""
    conn = _connect()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS staff (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fullname TEXT,
                sex TEXT,
                dob TEXT,
                blood_group TEXT,
                marital_status TEXT,
                nationality TEXT,
                state_origin TEXT,
                lga TEXT,
                address TEXT,
                next_of_kin TEXT,
                next_of_kin_phone TEXT,
                employment_type TEXT,
                state_office TEXT,
                cluster TEXT,
                department TEXT,
                section TEXT,
                position TEXT,
                staff_number TEXT UNIQUE,
                phone TEXT,
                email TEXT UNIQUE,
                facebook TEXT,
                twitter TEXT,
                instagram TEXT,
                telegram TEXT,
                linkedin TEXT,
                gps_coordinate TEXT,
                photo TEXT,
                password TEXT,
                unique_id TEXT UNIQUE
            )
        """)
        migrations = [
            ("cluster", "TEXT"),
            ("unique_id", "TEXT"),
            ("synced", "INTEGER DEFAULT 0"),
            ("genotype", "TEXT"),
            ("reintegration_status", "TEXT"),
        ]
        for column, col_type in migrations:
            try:
                cursor.execute(f"ALTER TABLE staff ADD COLUMN {column} {col_type}")
            except sqlite3.OperationalError as exc:
                if "duplicate column" not in str(exc).lower():
                    raise
        conn.commit()
    finally:
        conn.close()

    create_attendance_table()
    create_leave_table()


def create_attendance_table():
    """Create/migrate the attendance table."""
    conn = _connect()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email TEXT,
                check_in_time TEXT,
                check_out_time TEXT,
                late_duration TEXT,
                attendance_status TEXT,
                gps_location TEXT,
                check_out_gps_location TEXT,
                synced INTEGER DEFAULT 0
            )
        """)
        for column, col_type in (
            ("check_in_time", "TEXT"),
            ("check_out_time", "TEXT"),
            ("late_duration", "TEXT"),
            ("attendance_status", "TEXT"),
            ("gps_location", "TEXT"),
            ("check_out_gps_location", "TEXT"),
            ("synced", "INTEGER DEFAULT 0"),
            ("gform_synced", "INTEGER DEFAULT 0"),
        ):
            try:
                cursor.execute(f"ALTER TABLE attendance ADD COLUMN {column} {col_type}")
            except sqlite3.OperationalError as exc:
                if "duplicate column" not in str(exc).lower():
                    raise
        # Performance: attendance is queried by email and check_in_time on
        # nearly every screen (dashboard, reports, timesheet, gform sync),
        # and grows by rows every working day. Without an index these are
        # full table scans; the composite index covers both the plain
        # "WHERE email = ?" lookups and the "email = ? AND check_in_time
        # LIKE 'YYYY-MM%'" monthly-report queries.
        cursor.execute(
            "CREATE INDEX IF NOT EXISTS idx_attendance_email_checkin "
            "ON attendance(email, check_in_time)"
        )
        cursor.execute(
            "CREATE INDEX IF NOT EXISTS idx_attendance_gform_synced "
            "ON attendance(gform_synced)"
        )
        conn.commit()
    finally:
        conn.close()


def get_pending_gform_attendance(limit=50):
    """Completed attendance rows (checked out) not yet pushed to the Google
    Form, joined with the staff record for name/department/etc. Only rows
    with a check_out_time are returned - a check-in-only row is submitted
    once, as a single complete record, after the person checks out."""
    conn = _connect()
    try:
        rows = conn.execute(
            """
            SELECT a.id, a.email, a.check_in_time, a.check_out_time, a.gps_location,
                   a.check_out_gps_location, s.fullname, s.staff_number, s.department,
                   s.section, s.position
            FROM attendance a
            LEFT JOIN staff s ON s.email = a.email
            WHERE a.check_out_time IS NOT NULL AND a.check_out_time != ''
              AND (a.gform_synced IS NULL OR a.gform_synced = 0)
            ORDER BY a.id ASC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
        return rows
    finally:
        conn.close()


def mark_gform_synced(attendance_id):
    def operation(cursor):
        cursor.execute("UPDATE attendance SET gform_synced = 1 WHERE id = ?", (attendance_id,))
    return _run_write(operation)


def email_exists(email, exclude_id=None):
    """Case-insensitive email check before INSERT/UPDATE."""
    value = (email or "").strip().lower()
    if not value:
        return False
    conn = _connect()
    try:
        if exclude_id is None:
            row = conn.execute(
                "SELECT id FROM staff WHERE LOWER(TRIM(email)) = ? LIMIT 1",
                (value,),
            ).fetchone()
        else:
            row = conn.execute(
                "SELECT id FROM staff WHERE LOWER(TRIM(email)) = ? AND id <> ? LIMIT 1",
                (value, exclude_id),
            ).fetchone()
        return row is not None
    finally:
        conn.close()


def staff_number_exists(staff_number, exclude_id=None):
    value = (staff_number or "").strip().lower()
    if not value:
        return False
    conn = _connect()
    try:
        if exclude_id is None:
            row = conn.execute(
                "SELECT id FROM staff WHERE LOWER(TRIM(staff_number)) = ? LIMIT 1",
                (value,),
            ).fetchone()
        else:
            row = conn.execute(
                "SELECT id FROM staff WHERE LOWER(TRIM(staff_number)) = ? AND id <> ? LIMIT 1",
                (value, exclude_id),
            ).fetchone()
        return row is not None
    finally:
        conn.close()


def insert_staff(staff_data):
    """Insert a staff record with transient-lock retry."""
    def operation(cursor):
        cursor.execute("""
            INSERT INTO staff (
                fullname, sex, dob, blood_group, marital_status, nationality,
                state_origin, lga, address, next_of_kin, next_of_kin_phone,
                employment_type, state_office, cluster, department, section,
                position, staff_number, phone, email, facebook, twitter,
                instagram, telegram, linkedin, gps_coordinate, photo, password,
                unique_id, synced, genotype, reintegration_status
            ) VALUES (
                :fullname, :sex, :dob, :blood_group, :marital_status, :nationality,
                :state_origin, :lga, :address, :next_of_kin, :next_of_kin_phone,
                :employment_type, :state_office, :cluster, :department, :section,
                :position, :staff_number, :phone, :email, :facebook, :twitter,
                :instagram, :telegram, :linkedin, :gps_coordinate, :photo, :password,
                :unique_id, 0, :genotype, :reintegration_status
            )
        """, staff_data)
    return _run_write(operation)


def update_staff(staff_id, staff_data):
    """Update an existing staff record with transient-lock retry."""
    data = dict(staff_data)
    data["id"] = staff_id

    def operation(cursor):
        cursor.execute("""
            UPDATE staff SET
                synced = 0,
                fullname = :fullname,
                sex = :sex,
                dob = :dob,
                blood_group = :blood_group,
                marital_status = :marital_status,
                nationality = :nationality,
                state_origin = :state_origin,
                lga = :lga,
                address = :address,
                next_of_kin = :next_of_kin,
                next_of_kin_phone = :next_of_kin_phone,
                employment_type = :employment_type,
                state_office = :state_office,
                cluster = :cluster,
                department = :department,
                section = :section,
                position = :position,
                staff_number = :staff_number,
                phone = :phone,
                email = :email,
                facebook = :facebook,
                twitter = :twitter,
                instagram = :instagram,
                telegram = :telegram,
                linkedin = :linkedin,
                gps_coordinate = :gps_coordinate,
                photo = :photo,
                password = :password,
                genotype = :genotype,
                reintegration_status = :reintegration_status
            WHERE id = :id
        """, data)
    return _run_write(operation)


def get_staff_by_id(staff_id):
    conn = _connect()
    try:
        return conn.execute("SELECT * FROM staff WHERE id = ?", (staff_id,)).fetchone()
    finally:
        conn.close()


def get_staff_count():
    """Return how many staff records exist in the local database. Used to
    enforce a single on-device registration (one phone = one staff account)."""
    conn = _connect()
    try:
        row = conn.execute("SELECT COUNT(*) FROM staff").fetchone()
        return int(row[0] or 0)
    finally:
        conn.close()


def clear_all_staff():
    """Remove all existing staff records from the local database. Called
    before saving a brand new registration so a phone never ends up
    holding more than one staff registration at a time."""
    def operation(cursor):
        cursor.execute("DELETE FROM staff")
    return _run_write(operation)


def verify_login(email_or_staff_num, password):
    conn = _connect()
    try:
        identifier = (email_or_staff_num or "").strip().lower()
        return conn.execute("""
            SELECT * FROM staff
            WHERE (LOWER(TRIM(email)) = ? OR LOWER(TRIM(staff_number)) = ?)
              AND password = ?
        """, (identifier, identifier, password)).fetchone()
    finally:
        conn.close()


def create_leave_table():
    """Create the leave request table used by the Leave Management module."""
    conn = _connect()
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS leave_requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                staff_email TEXT NOT NULL,
                staff_name TEXT,
                leave_type TEXT NOT NULL,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                days INTEGER NOT NULL,
                reason TEXT,
                status TEXT DEFAULT 'Pending',
                manager_email TEXT,
                manager_comment TEXT,
                manager_signature TEXT,
                submitted_at TEXT,
                reviewed_at TEXT,
                synced INTEGER DEFAULT 0
            )
        """)
        conn.commit()
    finally:
        conn.close()


def create_leave_request(data):
    """Save a new leave request locally and mark it pending."""
    def operation(cursor):
        cursor.execute("""
            INSERT INTO leave_requests (
                staff_email, staff_name, leave_type, start_date, end_date,
                days, reason, status, submitted_at, synced
            ) VALUES (
                :staff_email, :staff_name, :leave_type, :start_date, :end_date,
                :days, :reason, 'Pending', :submitted_at, 0
            )
        """, data)
        return cursor.lastrowid
    return _run_write(operation)


def get_leave_requests(staff_email):
    conn = _connect()
    try:
        return conn.execute("""
            SELECT id, leave_type, start_date, end_date, days, reason,
                   status, manager_comment, submitted_at
            FROM leave_requests
            WHERE LOWER(TRIM(staff_email)) = LOWER(TRIM(?))
            ORDER BY id DESC
        """, (staff_email or "",)).fetchall()
    finally:
        conn.close()



def get_leave_status_counts(staff_email):
    """Return pending, approved and rejected leave request counts for a staff member."""
    conn = _connect()
    try:
        rows = conn.execute("""
            SELECT status, COUNT(*)
            FROM leave_requests
            WHERE LOWER(TRIM(staff_email)) = LOWER(TRIM(?))
            GROUP BY status
        """, (staff_email or "",)).fetchall()
        counts = {"Pending": 0, "Approved": 0, "Rejected": 0}
        for status, count in rows:
            key = str(status or "Pending").strip().title()
            if key in counts:
                counts[key] = int(count or 0)
        return counts
    finally:
        conn.close()

def get_leave_usage(staff_email, leave_type, year):
    """Return approved leave days used for a leave type in a year."""
    conn = _connect()
    try:
        row = conn.execute("""
            SELECT COALESCE(SUM(days), 0)
            FROM leave_requests
            WHERE LOWER(TRIM(staff_email)) = LOWER(TRIM(?))
              AND leave_type = ?
              AND status = 'Approved'
              AND substr(start_date, 1, 4) = ?
        """, (staff_email or "", leave_type, str(year))).fetchone()
        return int(row[0] or 0)
    finally:
        conn.close()
