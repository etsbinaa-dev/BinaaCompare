"""Vocabulary of RIM-Soft ledger column headers.

Header matching is exact on the normalized text (case, accents, punctuation
and "N°"/"Nº" variations removed), after stripping an optional "montant"
prefix and a currency suffix. No fuzzy matching: an unknown header is simply
not recognized, and missing required columns are reported instead of guessed.
"""

from __future__ import annotations

from enum import StrEnum

from binaa_reconciler.normalize.text import normalize_text


class ColumnRole(StrEnum):
    DATE = "DATE"
    LABEL = "LABEL"
    DEBIT = "DEBIT"
    CREDIT = "CREDIT"
    PIECE = "PIECE"
    ECRITURE = "ECRITURE"
    JOURNAL = "JOURNAL"
    REFERENCE = "REFERENCE"
    # Bankily statement columns (P2A)
    OPERATION = "OPERATION"
    BALANCE = "BALANCE"
    TIME = "TIME"


REQUIRED_ROLES: tuple[ColumnRole, ...] = (
    ColumnRole.DATE,
    ColumnRole.LABEL,
    ColumnRole.DEBIT,
    ColumnRole.CREDIT,
)
OPTIONAL_ROLES: tuple[ColumnRole, ...] = (
    ColumnRole.PIECE,
    ColumnRole.ECRITURE,
    ColumnRole.JOURNAL,
    ColumnRole.REFERENCE,
)

_SYNONYMS: dict[ColumnRole, frozenset[str]] = {
    ColumnRole.DATE: frozenset(
        {"date", "dates", "date operation", "date ecriture", "date piece", "date comptable"}
    ),
    ColumnRole.LABEL: frozenset(
        {
            "libelle",
            "libelles",
            "libelle ecriture",
            "libelle operation",
            "designation",
            "description",
        }
    ),
    ColumnRole.DEBIT: frozenset({"debit", "debits"}),
    ColumnRole.CREDIT: frozenset({"credit", "credits"}),
    ColumnRole.PIECE: frozenset(
        {"n piece", "no piece", "num piece", "numero piece", "n de piece", "piece", "pieces"}
    ),
    ColumnRole.ECRITURE: frozenset(
        {"n ecriture", "no ecriture", "num ecriture", "numero ecriture", "n d ecriture", "ecriture"}
    ),
    ColumnRole.JOURNAL: frozenset({"journal", "code journal", "jnl", "jrn"}),
    ColumnRole.REFERENCE: frozenset({"reference", "ref", "n reference", "no reference"}),
    ColumnRole.OPERATION: frozenset(
        {
            "n operation",
            "no operation",
            "num operation",
            "numero operation",
            "n d operation",
            "operation",
            "id operation",
            "reference operation",
            "n transaction",
            "id transaction",
        }
    ),
    ColumnRole.BALANCE: frozenset({"solde", "soldes", "solde courant", "nouveau solde"}),
    ColumnRole.TIME: frozenset({"heure", "heures"}),
}

_LOOKUP: dict[str, ColumnRole] = {s: role for role, syns in _SYNONYMS.items() for s in syns}
_CURRENCY_SUFFIXES = (" en mru", " mru", " en um", " um", " ouguiya", " ouguiyas")


def role_for_header(value: object) -> ColumnRole | None:
    """Column role of a header cell, or None when not recognized."""
    if not isinstance(value, str):
        return None
    text = normalize_text(value)
    if text.startswith("montant "):
        text = text[len("montant ") :]
    for suffix in _CURRENCY_SUFFIXES:
        if text.endswith(suffix):
            text = text[: -len(suffix)]
            break
    return _LOOKUP.get(text)
