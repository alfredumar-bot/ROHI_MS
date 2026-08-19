"""
Automatic Google Form submission for completed attendance rows.

Design:
- ROHI's Google Form is configured ONCE by pasting a "pre-filled link" (Forms
  menu -> Get pre-filled link) where the admin typed a fixed placeholder
  token into each question (see FIELD_TOKENS below). parse_prefilled_url()
  reads the entry.NNNNNNN=<token> pairs out of that URL and maps each token
  back to the field it belongs to - so nobody has to manually read off raw
  entry IDs.
- Once configured, submit_row() POSTs directly to the form's `formResponse`
  endpoint (the same endpoint the real Google Form page posts to), so a
  response appears in the linked Google Sheet exactly as if a person had
  filled the form in a browser - no browser, no manual tap needed.
- Uses only the Python standard library (urllib) - no extra dependency, and
  nothing new to add to buildozer.spec.
- Every function degrades to (ok=False, message=...) instead of raising, so
  a misconfigured/unreachable Google Form never crashes the app.
"""

import os
import json
import logging
import ssl
from urllib.parse import urlparse, parse_qs, urlencode
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError

logger = logging.getLogger("ROHIApp.gform_sync")

# See main.py for why this is needed: python-for-android builds don't expose
# the OS CA store to Python, so plain urlopen() on https:// fails SSL
# verification without an explicit CA bundle from certifi.
try:
    import certifi
    _SSL_CONTEXT = ssl.create_default_context(cafile=certifi.where())
except Exception:
    logger.warning("certifi not available - HTTPS requests may fail certificate verification on Android.")
    _SSL_CONTEXT = ssl.create_default_context()

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(APP_DIR, "gform_config.json")

# Placeholder token the admin types into each Google Form question when
# generating the "pre-filled link" in Google Forms. These exact strings (case
# sensitive) are what parse_prefilled_url() looks for - see the field-by-field
# instructions given to the admin.
FIELD_TOKENS = {
    "name": "ROHI_NAME",
    "staff_id": "ROHI_STAFFID",
    "department": "ROHI_DEPT",
    "section": "ROHI_SECTION",
    "position": "ROHI_POSITION",
    "date": "ROHI_DATE",
    "checkin": "ROHI_CHECKIN",
    "checkout": "ROHI_CHECKOUT",
    "gps": "ROHI_GPS",
}

DEFAULT_CONFIG = {
    "response_url": "",   # https://docs.google.com/forms/d/e/<id>/formResponse
    "entry_map": {},      # {"name": "entry.123", "staff_id": "entry.456", ...}
    "configured_fields": [],  # which of FIELD_TOKENS were actually found
}


# -------------------------------------------------------------
# Config persistence (gform_config.json)
# -------------------------------------------------------------
def load_config():
    if not os.path.exists(CONFIG_PATH):
        return dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        merged = dict(DEFAULT_CONFIG)
        merged.update(data)
        return merged
    except Exception:
        logger.exception("Failed to read gform_config.json; using defaults.")
        return dict(DEFAULT_CONFIG)


def save_config(config: dict):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f)
        return True, "Saved."
    except Exception:
        logger.exception("Failed to write gform_config.json:")
        return False, "Could not save Google Form configuration."


def is_configured():
    config = load_config()
    return bool(config.get("response_url")) and bool(config.get("entry_map"))


# -------------------------------------------------------------
# Parsing the admin's pasted "pre-filled link"
# -------------------------------------------------------------
def parse_prefilled_url(url: str):
    """Reads a Google Forms pre-filled link and returns
    (ok, message, entry_map, response_url).

    entry_map only contains the fields whose placeholder token
    (FIELD_TOKENS) was actually found in the URL - any question the admin
    left blank or typed something else into is simply skipped, and reported
    back in the message so they can fix it.
    """
    url = (url or "").strip()
    if not url:
        return False, "Paste the Google Form pre-filled link first.", {}, ""

    try:
        parsed = urlparse(url)
        if "docs.google.com" not in parsed.netloc or "/forms/" not in parsed.path:
            return False, "That doesn't look like a Google Forms link.", {}, ""

        query = parse_qs(parsed.query)
        # value-token -> raw entry id (entry.NNNNNNN)
        token_to_entry = {}
        for key, values in query.items():
            if not key.startswith("entry."):
                continue
            if not values:
                continue
            token_to_entry[values[0].strip()] = key

        entry_map = {}
        missing = []
        for field, token in FIELD_TOKENS.items():
            entry_id = token_to_entry.get(token)
            if entry_id:
                entry_map[field] = entry_id
            else:
                missing.append(field)

        if not entry_map:
            return (
                False,
                "No ROHI_* placeholder tokens were found in that link. "
                "Make sure you typed the exact tokens into the form before "
                "generating the pre-filled link.",
                {},
                "",
            )

        # .../forms/d/e/<form_id>/viewform?... -> .../forms/d/e/<form_id>/formResponse
        response_url = url.split("?")[0]
        if response_url.endswith("viewform"):
            response_url = response_url[: -len("viewform")] + "formResponse"
        elif not response_url.endswith("formResponse"):
            response_url = response_url.rstrip("/") + "/formResponse"

        if missing:
            message = (
                f"Linked {len(entry_map)} field(s). Not found (left as-is on the "
                f"form, won't be submitted): {', '.join(missing)}."
            )
        else:
            message = f"All {len(entry_map)} fields linked successfully."

        return True, message, entry_map, response_url

    except Exception:
        logger.exception("Failed to parse Google Form pre-filled link:")
        return False, "Could not read that link. Please check and try again.", {}, ""


# -------------------------------------------------------------
# Submitting a row
# -------------------------------------------------------------
def submit_row(row: dict, config=None):
    """Submits one attendance row to the configured Google Form.
    `row` keys should be a subset of FIELD_TOKENS keys (name, staff_id,
    department, section, position, date, checkin, checkout, gps).
    Returns (ok: bool, message: str).
    """
    config = config or load_config()
    response_url = config.get("response_url")
    entry_map = config.get("entry_map") or {}

    if not response_url or not entry_map:
        return False, "Google Form is not configured yet."

    form_payload = {}
    for field, value in row.items():
        entry_id = entry_map.get(field)
        if entry_id and value is not None:
            form_payload[entry_id] = str(value)

    if not form_payload:
        return False, "Nothing to submit - no configured fields matched this row."

    try:
        data = urlencode(form_payload).encode("utf-8")
        request = Request(response_url, data=data, method="POST")
        request.add_header("Content-Type", "application/x-www-form-urlencoded")
        request.add_header("User-Agent", "Mozilla/5.0 (ROHI-Attendance-App)")
        with urlopen(request, timeout=15, context=_SSL_CONTEXT) as response:
            status = response.status
        # Google Forms replies 200 on success even without validating
        # semantics; anything else is treated as a failure.
        if status == 200:
            return True, "Submitted to Google Form."
        return False, f"Google Form responded with status {status}."
    except HTTPError as exc:
        logger.exception("HTTP error submitting to Google Form:")
        return False, f"Google Form submission failed (HTTP {exc.code})."
    except URLError:
        logger.exception("Network error submitting to Google Form:")
        return False, "No internet connection - will retry on the next sync."
    except Exception:
        logger.exception("Unexpected error submitting to Google Form:")
        return False, "Unexpected error submitting to Google Form."
