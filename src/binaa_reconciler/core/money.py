"""Money handling.

Amounts are stored as integers in minor units (1/100 of the major unit) to
avoid floating point errors. Parsing never guesses: an ambiguous string such
as "1,250" raises instead of choosing between 1.25 and 1250.

Error messages carry a code only, never the offending value, so they are safe
to log.
"""

from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation

from binaa_reconciler.normalize.text import to_ascii_digits

MINOR_UNITS_PER_MAJOR = 100
_FLOAT_NOISE = Decimal("0.000001")  # in minor units

_CURRENCY_RE = re.compile(
    r"(?i)\b(?:mru|mro|um|ouguiyas?|oug)\b|أوقية(?:\s+جديدة)?|اوقية(?:\s+جديدة)?"
)
_GROUP_SPACES = (" ", " ", " ", " ", "'", "’")


class AmountParseError(ValueError):
    """Raised when an amount cannot be parsed with certainty."""

    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


def _valid_groups(groups: list[str]) -> bool:
    """Thousands groups: first 1-3 digits, following exactly 3 digits."""
    if not groups or not groups[0].isdigit() or not 1 <= len(groups[0]) <= 3:
        return False
    return all(len(g) == 3 and g.isdigit() for g in groups[1:])


def parse_amount(text: str) -> int:
    """Parse a human formatted amount into minor units.

    Supports "1 250,00", "1.250,00", "1,250.00", "15000", "-2 000", "(2 000)",
    "2 000-", currency suffixes (MRU, UM, أوقية) and Arabic-Indic digits.
    """
    s = to_ascii_digits(text).strip()
    s = _CURRENCY_RE.sub("", s).strip()
    if not s:
        raise AmountParseError("empty")

    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative, s = True, s[1:-1].strip()
    if s.startswith(("-", "+")):
        negative, s = negative or s[0] == "-", s[1:].strip()
    elif s.endswith("-"):
        negative, s = True, s[:-1].strip()

    # Space-like grouping: validate groups before removing them.
    for sp in _GROUP_SPACES:
        s = s.replace(sp, " ")
    if " " in s:
        head, _, tail = s.rpartition(" ")
        int_tail = re.split(r"[.,]", tail, maxsplit=1)[0]
        if not _valid_groups([*head.split(" "), int_tail]):
            raise AmountParseError("bad_grouping")
        s = s.replace(" ", "")

    if not s or not re.fullmatch(r"[0-9.,]+", s):
        raise AmountParseError("invalid_characters")

    has_dot, has_comma = "." in s, "," in s
    if has_dot and has_comma:
        dec_sep = "." if s.rfind(".") > s.rfind(",") else ","
        grp_sep = "," if dec_sep == "." else "."
        int_part, _, frac = s.rpartition(dec_sep)
        if dec_sep in int_part or not _valid_groups(int_part.split(grp_sep)):
            raise AmountParseError("bad_grouping")
        int_digits = int_part.replace(grp_sep, "")
    elif has_dot or has_comma:
        sep = "." if has_dot else ","
        parts = s.split(sep)
        if len(parts) > 2:
            if not _valid_groups(parts):
                raise AmountParseError("bad_grouping")
            int_digits, frac = "".join(parts), ""
        else:
            int_digits, frac = parts
            if len(frac) == 3:
                raise AmountParseError("ambiguous_separator")
            if not int_digits or len(frac) not in (1, 2):
                raise AmountParseError("bad_decimals")
    else:
        int_digits, frac = s, ""

    if not int_digits.isdigit() or (frac and not frac.isdigit()) or len(frac) > 2:
        raise AmountParseError("bad_decimals")

    minor = int(int_digits) * MINOR_UNITS_PER_MAJOR + int((frac + "00")[:2])
    return -minor if negative else minor


def amount_from_number(value: int | float | Decimal) -> int:
    """Convert a numeric cell value (e.g. from Excel) to minor units.

    Floats may carry binary noise (0.1 + 0.2 = 0.30000000000000004); noise below
    `_FLOAT_NOISE` minor units is removed. More than two *meaningful* decimals
    is an error, never a silent rounding.
    """
    if isinstance(value, bool):
        raise AmountParseError("not_a_number")
    try:
        dec = Decimal(str(value)) if isinstance(value, float) else Decimal(value)
    except (InvalidOperation, ValueError) as exc:
        raise AmountParseError("not_a_number") from exc
    if not dec.is_finite():
        raise AmountParseError("not_a_number")
    scaled = dec * MINOR_UNITS_PER_MAJOR
    nearest = scaled.to_integral_value()
    tolerance = _FLOAT_NOISE if isinstance(value, float) else Decimal(0)
    if abs(scaled - nearest) > tolerance:
        raise AmountParseError("too_many_decimals")
    return int(nearest)


def format_amount(minor: int) -> str:
    """Format minor units as French style "1 250,00" (for display/reports)."""
    sign = "-" if minor < 0 else ""
    major, cents = divmod(abs(minor), MINOR_UNITS_PER_MAJOR)
    grouped = f"{major:,}".replace(",", " ")
    return f"{sign}{grouped},{cents:02d}"


NBSP = "\u00a0"  # no-break space: groups digits without splitting the number
LRM = "\u200e"  # left-to-right mark: keeps "3 000 MRU" in reading order in RTL text


def format_major(minor: int) -> str:
    """Display form of an amount: "3 000", "3 000,50" (no ",00" when whole).

    Groups are separated by a no-break space and decimals by a comma: both are
    common number separators for the bidi algorithm, so the number stays one
    left-to-right run inside Arabic text. Formatting only: nothing is rounded.
    """
    sign = "-" if minor < 0 else ""
    major, cents = divmod(abs(minor), MINOR_UNITS_PER_MAJOR)
    grouped = f"{major:,}".replace(",", NBSP)
    return f"{sign}{grouped},{cents:02d}" if cents else f"{sign}{grouped}"


def format_money(minor: int, currency: str = "MRU") -> str:
    """Amount with its unit after it: "3 000 MRU", "3 000,50 MRU" (no-break spaces)."""
    return f"{format_major(minor)}{NBSP}{currency}"


def is_round_amount(minor: int, step_major: int = 100) -> bool:
    """True for "common" amounts (multiples of `step_major`), weak evidence."""
    return minor != 0 and minor % (step_major * MINOR_UNITS_PER_MAJOR) == 0


def is_ouguiya_x10_suspect(a_minor: int, b_minor: int) -> bool:
    """True when one amount is exactly 10x the other (old vs new ouguiya)."""
    a, b = abs(a_minor), abs(b_minor)
    return a != 0 and b != 0 and (a == 10 * b or b == 10 * a)
