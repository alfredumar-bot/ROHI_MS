[app]
title = Rohi IMS
package.name = rohiattendance
package.domain = org.rohi

source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,db,ttf,json,xlsx

version = 1.4

# kivymd is pinned to 1.2.0 to match what the app was built/tested against
# (KivyMD 2.0 changes several widget APIs used here, e.g. MDDropdownMenu).
# openpyxl -> Timesheet/Attendance Excel export. reportlab -> Attendance PDF export.
# pg8000 -> hybrid SQLite/PostgreSQL sync (Settings > Server Connection). Chosen
# over psycopg2 because it's pure-Python (no C extension), so it doesn't need a
# matching prebuilt wheel or C toolchain to cross-compile for Android via
# buildozer/python-for-android.
# python3 is pinned to 3.10.11 (rather than left to float to the newest
# available, e.g. 3.14) because reportlab's C accelerator (_rl_accel) uses
# CPython APIs removed/changed in 3.11+ (PyUnicode_AS_UNICODE, opaque frame
# objects) and fails to cross-compile against anything newer than 3.10.
requirements = python3==3.10.11,hostpython3==3.10.11,kivy==2.3.1,kivymd==1.2.0,plyer,pyjnius,sqlite3,openpyxl,et_xmlfile,reportlab,pypdf,cryptography,msoffcrypto-tool,pillow,pg8000

icon.filename = %(source.dir)s/rohi_logo.png
presplash.filename = %(source.dir)s/rohi_logo.png

orientation = portrait
fullscreen = 0

# Camera (staff photo capture), GPS (base location capture + check-in/out),
# and storage (sqlite db + saved staff photos) are all used by main.py.
# FOREGROUND_SERVICE / FOREGROUND_SERVICE_DATA_SYNC / WAKE_LOCK / RECEIVE_BOOT_COMPLETED
# are for the "Reminder" background service (service_reminder.py) below, which
# keeps firing the 07:50/07:55/07:59 check-in and 15:50-series check-out
# reminders even while the app itself is closed or the screen is off.
android.permissions = CAMERA,ACCESS_FINE_LOCATION,ACCESS_COARSE_LOCATION,READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,READ_MEDIA_IMAGES,INTERNET,POST_NOTIFICATIONS,FOREGROUND_SERVICE,FOREGROUND_SERVICE_DATA_SYNC,WAKE_LOCK,RECEIVE_BOOT_COMPLETED

# Background reminder service - runs service_reminder.py as an independent
# Android Service/process (started from main.py's on_start), so the
# check-in/out reminders keep firing while the app is closed. ",foreground"
# makes it a foreground service with a persistent low-priority notification,
# which is required on modern Android so the OS doesn't kill it under Doze /
# battery optimization.
services = Reminder:service_reminder.py:foreground

android.api = 34
android.minapi = 21
android.ndk = 25b
android.archs = arm64-v8a
android.allow_backup = True

# Required for CI: auto-accept Android SDK licenses (non-interactive build)
android.accept_sdk_license = True

# Build for arm64-v8a only, not armeabi-v7a. The cryptography package's Rust
# build (via cryptography-cffi/maturin) breaks on the 32-bit armeabi-v7a
# target with a LONG_BIT/glibc header mismatch that's a known incompatibility
# in python-for-android's Rust cross-compile support for 32-bit Android. All
# phones from ~2018 onward are 64-bit, so arm64-v8a-only covers virtually
# every real device, and it also roughly halves build time.

[buildozer]
log_level = 2
warn_on_root = 1
