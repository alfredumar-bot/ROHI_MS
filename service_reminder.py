"""
Background reminder service for ROHI Attendance.

Declared in buildozer.spec as:
    services = Reminder:service_reminder.py,foreground

python-for-android runs this file as an independent Android Service/process
(started from main.py via App._start_background_reminder_service). That is
what lets the 07:50 / 07:55 / 07:59 check-in nudges (and the 15:50-series
check-out nudges) keep firing even while the main app is closed, backgrounded,
or the screen is off - the Kivy Clock-based reminder in main.py only runs
while the app is actually in the foreground.

Deliberately does NOT import main.py, Kivy, or KivyMD: this process only
needs sqlite3/json/plyer, so it starts fast and stays lightweight. It reads
the same session.json / attendance.db files the main app already writes, so
no extra state or IPC is needed between the two processes.

Every failure here is caught and swallowed - worst case the reminder simply
doesn't fire in the background and the app falls back to the same in-app
Clock-based reminder as before (i.e. this can never crash or block the app).
"""
import os
import json
import time
import sqlite3
from datetime import datetime, timedelta

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SESSION_FILE = os.path.join(APP_DIR, "session.json")
DB_PATH = os.path.join(APP_DIR, "attendance.db")
APP_SETTINGS_FILE = os.path.join(APP_DIR, "app_settings.json")

REMINDER_CHECKIN_HOUR, REMINDER_CHECKIN_MINUTE = 7, 50
REMINDER_CHECKOUT_HOUR, REMINDER_CHECKOUT_MINUTE = 15, 50
# Exact clock times: 07:50 (on-time), 07:55, 07:59 - mirrors main.py's
# CHECKIN_REMINDER_OFFSETS_MINUTES. Keep these two files in sync if the
# schedule ever changes.
CHECKIN_OFFSETS_MINUTES = [0, 5, 9]
CHECKOUT_OFFSETS_MINUTES = [0, 5, 10, 15]

POLL_SECONDS = 20


def _is_working_day(now):
    """Monday-Thursday only; Friday is WFH, Sat/Sun are off. Mirrors
    ROHIApp._is_attendance_working_day in main.py."""
    return now.weekday() in (0, 1, 2, 3)


def _reminders_enabled():
    try:
        if not os.path.exists(APP_SETTINGS_FILE):
            return True
        with open(APP_SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return bool(data.get("reminders_enabled", True))
    except Exception:
        return True


def _logged_in_email():
    try:
        if not os.path.exists(SESSION_FILE):
            return None
        with open(SESSION_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        email = str(data.get("email") or "").strip().lower()
        return email or None
    except Exception:
        return None


def _today_row(email):
    try:
        if not os.path.exists(DB_PATH):
            return None
        today_prefix = datetime.now().strftime("%Y-%m-%d")
        conn = sqlite3.connect(DB_PATH, timeout=10)
        try:
            cur = conn.cursor()
            cur.execute(
                "SELECT check_in_time, check_out_time FROM attendance "
                "WHERE email = ? AND check_in_time LIKE ? ORDER BY id DESC LIMIT 1",
                (email, f"{today_prefix}%"),
            )
            return cur.fetchone()
        finally:
            conn.close()
    except Exception:
        return None


def _notify(title, message):
    try:
        from plyer import notification
        notification.notify(title=title, message=message, app_name="ROHI Attendance", timeout=15)
    except Exception:
        pass


def _fire(kind, now, hour, minute, already_done, offsets, fired_slots):
    if already_done:
        return
    base_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    last_slot = len(offsets) - 1
    label = "Check-In" if kind == "checkin" else "Check-Out"
    verb = "checked in" if kind == "checkin" else "checked out"
    for slot, offset in enumerate(offsets):
        if slot in fired_slots:
            continue
        slot_time = base_time + timedelta(minutes=offset)
        if now >= slot_time:
            fired_slots.add(slot)
            if slot == 0:
                _notify(
                    "ROHI {} Reminder".format(label),
                    "It's time to {}. Open ROHI Attendance and tap {}.".format(label.lower(), label),
                )
            else:
                _notify(
                    "ROHI {} Reminder ({}/{})".format(label, slot, last_slot),
                    "You still haven't {} today. Please do so now.".format(verb),
                )


def run():
    fired = {"date": None, "checkin": set(), "checkout": set()}
    while True:
        try:
            now = datetime.now()
            if fired["date"] != now.date():
                fired = {"date": now.date(), "checkin": set(), "checkout": set()}
            if _is_working_day(now) and _reminders_enabled():
                email = _logged_in_email()
                if email:
                    row = _today_row(email)
                    checked_in = bool(row and row[0])
                    checked_out = bool(row and row[1])
                    _fire("checkin", now, REMINDER_CHECKIN_HOUR, REMINDER_CHECKIN_MINUTE,
                          checked_in, CHECKIN_OFFSETS_MINUTES, fired["checkin"])
                    _fire("checkout", now, REMINDER_CHECKOUT_HOUR, REMINDER_CHECKOUT_MINUTE,
                          checked_out, CHECKOUT_OFFSETS_MINUTES, fired["checkout"])
        except Exception:
            pass
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    run()
