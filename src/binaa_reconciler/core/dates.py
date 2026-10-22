"""Date and time parsing without guessing.

Day/month order is decided once per column (`infer_date_order`); a single
value is never interpreted on its own when it is ambiguous. Mauritania uses
UTC+0 without daylight saving, so naive dates/times are sufficient.
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from datetime import date, datetime, time, timedelta
from enum import StrEnum

from binaa_reconciler.normalize.text import to_ascii_digits


class DateOrder(StrEnum):
    DMY = "DMY"
    MDY = "MDY"
    YMD = "YMD"


class DateParseError(ValueError):
    """Raised when a date/time cannot be parsed with certainty (code only)."""

    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


_DATE_RE = re.compile(r"^(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{1,4})$")
_TIME_RE = re.compile(r"^(\d{1,2})[:hH](\d{2})(?::(\d{2}))?$")

# Excel 1900 date system (with the Lotus 1900-02-29 bug): serial 1 = 1900-01-01.
_EXCEL_EPOCH = date(1899, 12, 30)
_EXCEL_FAKE_LEAP_DAY = 60


def _split(text: str) -> tuple[str, str, str]:
    m = _DATE_RE.match(to_ascii_digits(text).strip())
    if not m:
        raise DateParseError("bad_format")
    return m.group(1), m.group(2), m.group(3)


def _year(raw: str) -> int:
    if len(raw) == 4:
        return int(raw)
    if len(raw) == 2:
        return 2000 + int(raw)
    raise DateParseError("bad_year")


def parse_date(text: str, order: DateOrder) -> date:
    """Parse "25/03/2026", "2026-03-25", "25.03.26"... with an explicit order."""
    a, b, c = _split(text)
    if len(a) == 4 and order is not DateOrder.YMD:
        raise DateParseError("order_mismatch")
    try:
        if order is DateOrder.YMD:
            if len(a) != 4:
                raise DateParseError("order_mismatch")
            return date(int(a), int(b), int(c))
        if order is DateOrder.DMY:
            return date(_year(c), int(b), int(a))
        return date(_year(c), int(a), int(b))
    except ValueError as exc:
        if isinstance(exc, DateParseError):
            raise
        raise DateParseError("out_of_range") from exc


def parse_time(text: str) -> time:
    """Parse "14:05", "14h05", "14:05:33"."""
    m = _TIME_RE.match(to_ascii_digits(text).strip())
    if not m:
        raise DateParseError("bad_time")
    try:
        return time(int(m.group(1)), int(m.group(2)), int(m.group(3) or 0))
    except ValueError as exc:
        raise DateParseError("bad_time") from exc


def parse_datetime(text: str, order: DateOrder) -> tuple[date, time | None]:
    """Parse "25/03/2026 14:05" (time optional)."""
    parts = to_ascii_digits(text).strip().split()
    if len(parts) == 1:
        return parse_date(parts[0], order), None
    if len(parts) == 2:
        return parse_date(parts[0], order), parse_time(parts[1])
    raise DateParseError("bad_format")


def infer_date_order(values: Iterable[str]) -> DateOrder | None:
    """Infer the day/month order of a whole column.

    Returns None when the column is ambiguous (every value has day <= 12) or
    contradictory; the caller must then ask the user instead of guessing.
    Unparseable values are ignored here (they are reported elsewhere).
    """
    saw_ymd = saw_other = dmy_evidence = mdy_evidence = False
    for value in values:
        try:
            a, b, _ = _split(value)
        except DateParseError:
            continue
        if len(a) == 4:
            saw_ymd = True
            continue
        saw_other = True
        if int(a) > 12:
            dmy_evidence = True
        if int(b) > 12:
            mdy_evidence = True
    if saw_ymd and not saw_other:
        return DateOrder.YMD
    if saw_ymd and saw_other:
        return None
    if dmy_evidence and not mdy_evidence:
        return DateOrder.DMY
    if mdy_evidence and not dmy_evidence:
        return DateOrder.MDY
    return None


def excel_serial_to_datetime(serial: float) -> tuple[date, time | None]:
    """Convert an Excel (1900 system) serial number to date and optional time."""
    if serial < 1 or int(serial) == _EXCEL_FAKE_LEAP_DAY:
        raise DateParseError("bad_excel_serial")
    days = int(serial)
    if days < _EXCEL_FAKE_LEAP_DAY:
        days += 1  # serials before the fake leap day are shifted by one
    d = _EXCEL_EPOCH + timedelta(days=days)
    fraction = serial - int(serial)
    if fraction <= 0:
        return d, None
    seconds = round(fraction * 86400)
    if seconds >= 86400:
        return d + timedelta(days=1), None
    return d, (datetime.min + timedelta(seconds=seconds)).time()


def days_between(earlier: date, later: date) -> int:
    """Signed number of days from `earlier` to `later`."""
    return (later - earlier).days
