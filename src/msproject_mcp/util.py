"""Date and duration helpers.

MS Project hands back ``pywintypes.datetime`` values for date properties. They
carry the local wall clock but a UTC ``tzinfo``, so converting them to local time
shifts dates by hours (and sometimes a whole day). We strip the tzinfo instead of
converting, which keeps the value the user actually sees in Project.
"""

from __future__ import annotations

import datetime as dt
import re
from typing import Any

_DURATION_RE = re.compile(r"^\s*(-?\d+(?:[.,]\d+)?)\s*([a-zA-Z]*)\s*$")
_TIME_RE = re.compile(r"^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*$")

_UNIT_FACTORS = {
    "": None,
    "m": 1.0,
    "min": 1.0,
    "mins": 1.0,
    "minute": 1.0,
    "minutes": 1.0,
    "h": 60.0,
    "hr": 60.0,
    "hrs": 60.0,
    "hour": 60.0,
    "hours": 60.0,
    "d": "day",
    "day": "day",
    "days": "day",
    "w": "week",
    "wk": "week",
    "wks": "week",
    "week": "week",
    "weeks": "week",
}

_DATE_FORMATS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y",
)


def to_naive(value: Any) -> dt.datetime | None:
    """Return the local wall-clock value as a naive datetime.

    pywin32 marshals a COM DATE (which has no timezone) into a UTC-tagged
    ``pywintypes.datetime``. Project's 2026-01-05 therefore arrives as
    ``2026-01-04 23:00+00:00`` on a UTC+1 machine, so we must convert to local
    time rather than merely dropping tzinfo.
    """
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        if value.tzinfo is not None:
            return value.astimezone().replace(tzinfo=None)
        return value
    if isinstance(value, dt.date):
        return dt.datetime(value.year, value.month, value.day)
    text = str(value).strip()
    if not text:
        return None
    return parse_datetime(text)


def parse_datetime(value: Any) -> dt.datetime:
    """Parse a date, datetime, ISO string, or HH:MM time into a naive datetime."""
    if isinstance(value, dt.datetime):
        return to_naive(value)  # type: ignore[return-value]
    if isinstance(value, dt.date):
        return dt.datetime(value.year, value.month, value.day)
    if not isinstance(value, str):
        raise ValueError(f"Expected a date or ISO string, got {type(value).__name__}.")

    text = value.strip()
    if not text:
        raise ValueError("Empty date value.")

    match = _TIME_RE.match(text)
    if match:
        hour, minute = int(match.group(1)), int(match.group(2))
        second = int(match.group(3) or 0)
        if hour > 23 or minute > 59:
            raise ValueError(f"Invalid time {value!r}.")
        today = dt.date.today()
        return dt.datetime(today.year, today.month, today.day, hour, minute, second)

    cleaned = text[:-1] if text.endswith(("Z", "z")) else text
    try:
        return to_naive(dt.datetime.fromisoformat(cleaned))  # type: ignore[return-value]
    except ValueError:
        pass

    for fmt in _DATE_FORMATS:
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            continue

    raise ValueError(
        f"Unrecognised date {value!r}. Use YYYY-MM-DD, YYYY-MM-DD HH:MM, or HH:MM."
    )


def format_datetime(value: Any) -> str | None:
    """Format a Project date as YYYY-MM-DD, or YYYY-MM-DD HH:MM when it has a time.

    Project returns the string "NA" for a few unset date properties (StatusDate
    among them), so output formatting must never raise.
    """
    try:
        parsed = to_naive(value)
    except (TypeError, ValueError):
        text = str(value).strip()
        return text or None
    if parsed is None:
        return None
    if parsed.hour == 0 and parsed.minute == 0:
        return parsed.strftime("%Y-%m-%d")
    return parsed.strftime("%Y-%m-%d %H:%M")


def parse_duration(
    value: Any, minutes_per_day: float = 480.0, minutes_per_week: float | None = None
) -> float:
    """Parse "5d", "12h", "480m", a number of minutes, or a timedelta into minutes."""
    if value is None:
        raise ValueError("Duration is required.")
    if isinstance(value, dt.timedelta):
        return value.total_seconds() / 60.0
    if isinstance(value, bool):
        raise ValueError(f"Unrecognised duration {value!r}.")
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    match = _DURATION_RE.match(text)
    if not match:
        raise ValueError(
            f"Unrecognised duration {value!r}. Use e.g. '5d', '12h', '480m', or a number of minutes."
        )

    number = float(match.group(1).replace(",", "."))
    unit = match.group(2).lower()
    if unit not in _UNIT_FACTORS:
        raise ValueError(
            f"Unrecognised duration unit in {value!r}. Use m, h, d, or w."
        )

    factor = _UNIT_FACTORS[unit]
    if factor == "day":
        factor = float(minutes_per_day or 480.0)
    elif factor == "week":
        factor = float(minutes_per_week or (float(minutes_per_day or 480.0) * 5))
    elif factor is None:
        factor = 1.0
    return number * float(factor)


def _trim(number: float) -> str:
    if abs(number - round(number)) < 1e-9:
        return str(int(round(number)))
    return f"{number:.2f}".rstrip("0").rstrip(".")


def format_duration(
    minutes: Any, minutes_per_day: float = 480.0, minutes_per_week: float | None = None
) -> str | None:
    """Format minutes as the largest clean unit: 2400 -> "5d", 90 -> "1.5h"."""
    if minutes is None:
        return None
    try:
        total = float(minutes)
    except (TypeError, ValueError):
        return None

    per_day = float(minutes_per_day or 480.0)
    per_week = float(minutes_per_week or (per_day * 5))

    for unit, size in (("w", per_week), ("d", per_day), ("h", 60.0)):
        if size and abs(total) >= size and abs(total % size) < 1e-6:
            return f"{_trim(total / size)}{unit}"
    return f"{_trim(total)}m"


def to_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    text = str(value).strip().lower()
    if text in {"1", "true", "yes", "y", "on"}:
        return True
    if text in {"0", "false", "no", "n", "off", ""}:
        return False
    raise ValueError(f"Expected a boolean, got {value!r}.")
