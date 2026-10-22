"""Stable fingerprints used to remember decisions across re-imports.

Neither N° pièce nor N° écriture is required: they only strengthen the
fingerprint when present. Without them, RIM-Soft lines are identified by
(date, amount, side, normalized label, occurrence index), where the occurrence
index separates genuinely identical lines in file order.
"""

from __future__ import annotations

import hashlib
from collections.abc import Hashable, Sequence
from datetime import date, time
from pathlib import Path

from binaa_reconciler.core.models import Direction, LedgerSide

FINGERPRINT_VERSION = "1"
_SEP = "\x1f"


def _digest(prefix: str, parts: Sequence[str]) -> str:
    payload = _SEP.join([prefix, FINGERPRINT_VERSION, *parts]).encode("utf-8")
    return f"{prefix}{FINGERPRINT_VERSION}:{hashlib.sha256(payload).hexdigest()[:32]}"


def _norm_ref(value: str) -> str:
    return "".join(value.split()).upper()


def bankily_fingerprint(
    *,
    reference: str | None,
    tx_date: date,
    tx_time: time | None,
    amount_minor: int,
    direction: Direction,
    balance_minor: int | None,
) -> str:
    """Bankily: the operation reference when present, else a composite key.

    The running balance makes the composite key unique in practice.
    """
    if reference and reference.strip():
        return _digest("bk", ["ref", _norm_ref(reference)])
    return _digest(
        "bk",
        [
            "cmp",
            tx_date.isoformat(),
            tx_time.isoformat() if tx_time else "",
            str(amount_minor),
            direction.value,
            "" if balance_minor is None else str(balance_minor),
        ],
    )


def rimsoft_fingerprint(
    *,
    piece_number: str | None,
    tx_date: date,
    amount_minor: int,
    side: LedgerSide,
    label_norm: str,
    occurrence_index: int,
    ecriture_number: str | None = None,
) -> str:
    """RIM-Soft: optional piece / entry numbers + composite key + occurrence index.

    The entry number (N° écriture) is appended only when present, so
    fingerprints of files without that column are unchanged.
    """
    parts = [
        _norm_ref(piece_number) if piece_number and piece_number.strip() else "",
        tx_date.isoformat(),
        str(amount_minor),
        side.value,
        label_norm,
        str(occurrence_index),
    ]
    if ecriture_number and ecriture_number.strip():
        parts.append("ecr=" + _norm_ref(ecriture_number))
    return _digest("rs", parts)


def occurrence_indexes(keys: Sequence[Hashable]) -> list[int]:
    """For each key, how many identical keys appeared before it (0, 1, 2...)."""
    seen: dict[Hashable, int] = {}
    result: list[int] = []
    for key in keys:
        n = seen.get(key, 0)
        result.append(n)
        seen[key] = n + 1
    return result


def file_sha256(path: Path, chunk_size: int = 1 << 20) -> str:
    """SHA-256 of a local file (to identify an import without storing its content)."""
    h = hashlib.sha256()
    with path.open("rb") as fh:
        while chunk := fh.read(chunk_size):
            h.update(chunk)
    return h.hexdigest()
