"""Mauritanian phone number normalization.

National numbers have 8 digits. Mobile numbers start with 2 (Chinguitel),
3 (Mattel) or 4 (Mauritel); fixed lines also start with 4. International
prefixes +222 / 00222 / 222 are removed.
"""

from __future__ import annotations

import re

from binaa_reconciler.normalize.text import to_ascii_digits

PHONE_LENGTH = 8
COUNTRY_CODE = "222"
VALID_FIRST_DIGITS = frozenset("234")

# A phone inside free text: optional country prefix, then 8 digits optionally
# grouped by 2 with space / dot / dash. Not glued to other digits.
_PHONE_IN_TEXT_RE = re.compile(
    r"(?<![\d+])"
    r"(?:(?:\+|00)?222[\s.\-]?)?"
    r"([234]\d(?:[\s.\-]?\d{2}){3})"
    r"(?![\d])"
)


def normalize_mr_phone(text: str) -> str | None:
    """Normalize a phone *field* to 8 digits, or None if it is not a valid MR number.

    Strict: any letter in the value, or a foreign country code, gives None.
    """
    s = to_ascii_digits(text).strip()
    if not s or re.search(r"[^\d\s+().\-/]", s):
        return None
    has_plus = s.lstrip("(").startswith("+")
    digits = re.sub(r"\D", "", s)
    if has_plus:
        if not digits.startswith(COUNTRY_CODE):
            return None
        digits = digits[len(COUNTRY_CODE) :]
    elif digits.startswith("00"):
        if not digits.startswith("00" + COUNTRY_CODE):
            return None
        digits = digits[2 + len(COUNTRY_CODE) :]
    elif digits.startswith(COUNTRY_CODE) and len(digits) == PHONE_LENGTH + len(COUNTRY_CODE):
        digits = digits[len(COUNTRY_CODE) :]
    if len(digits) == PHONE_LENGTH and digits[0] in VALID_FIRST_DIGITS:
        return digits
    return None


def extract_mr_phones(text: str) -> tuple[str, ...]:
    """Find Mauritanian phone numbers inside free text (e.g. a ledger label).

    Heuristic: an 8-digit operation reference starting with 2/3/4 would also
    match, so a phone found in text is evidence, never proof on its own.
    """
    found: list[str] = []
    for m in _PHONE_IN_TEXT_RE.finditer(to_ascii_digits(text)):
        digits = re.sub(r"\D", "", m.group(1))
        if digits not in found:
            found.append(digits)
    return tuple(found)
