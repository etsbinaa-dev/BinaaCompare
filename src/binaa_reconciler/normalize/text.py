"""Generic text normalization for Arabic and Latin (French) content.

The goal is comparison, never display: normalized strings are lossy on purpose.
The raw text is always kept separately by the caller.
"""

from __future__ import annotations

import re
import unicodedata

_ARABIC_INDIC_DIGITS = str.maketrans(
    "٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹",
    "01234567890123456789",
)

# Harakat, Quranic marks, superscript alef.
_TASHKEEL_RE = re.compile("[ؐ-ًؚ-ٰٟۖ-ۭ]")
_TATWEEL = "ـ"

# Letter unification that does not change meaning for matching purposes.
_ARABIC_LETTER_MAP = str.maketrans(
    {
        "أ": "ا",
        "إ": "ا",
        "آ": "ا",
        "ٱ": "ا",
        "ة": "ه",
        "ى": "ي",
        "ؤ": "و",
        "ئ": "ي",
        "ء": "",
        "گ": "ك",
        "ڤ": "ف",
    }
)

_NON_WORD_RE = re.compile(r"[^\w]+", re.UNICODE)
_SPACES_RE = re.compile(r"\s+")

# Tokens that carry no identifying information in a ledger label.
GENERIC_LABEL_TOKENS: frozenset[str] = frozenset(
    {
        "bankil",
        "bankily",
        "bankili",
        "bnk",
        "bpm",
        "versement",
        "vers",
        "vrst",
        "virement",
        "vir",
        "depot",
        "dpt",
        "paiement",
        "pmt",
        "reglement",
        "regl",
        "rglt",
        "encaissement",
        "enc",
        "recu",
        "recette",
        "transfert",
        "trf",
        "client",
        "clients",
        "divers",
        "operation",
        "op",
        "mru",
        "um",
        "du",
        "de",
        "le",
        "la",
        "par",
        "via",
        "بنكيلي",
        "دفع",
        "تحويل",
        "ايداع",
    }
)


def to_ascii_digits(text: str) -> str:
    """Convert Arabic-Indic and Persian digits to ASCII digits."""
    return text.translate(_ARABIC_INDIC_DIGITS)


def strip_latin_accents(text: str) -> str:
    """Remove combining marks (é -> e). Arabic letters are unaffected in meaning."""
    decomposed = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in decomposed if not unicodedata.combining(ch))


def normalize_arabic(text: str) -> str:
    """Remove diacritics and tatweel, unify alef/teh-marbuta/yeh variants."""
    text = _TASHKEEL_RE.sub("", text).replace(_TATWEEL, "")
    return text.translate(_ARABIC_LETTER_MAP)


def normalize_text(text: str) -> str:
    """Normalize free text for comparison.

    NFKC (folds Arabic presentation forms found in PDFs) -> ASCII digits ->
    Arabic unification -> casefold -> accent removal -> punctuation to space ->
    collapse whitespace.
    """
    text = unicodedata.normalize("NFKC", text)
    text = to_ascii_digits(text)
    text = normalize_arabic(text)
    text = text.casefold()
    text = strip_latin_accents(text)
    text = _NON_WORD_RE.sub(" ", text).replace("_", " ")
    return _SPACES_RE.sub(" ", text).strip()


def tokens(text: str) -> list[str]:
    """Normalized tokens of `text`."""
    normalized = normalize_text(text)
    return normalized.split(" ") if normalized else []


def is_generic_label(text: str) -> bool:
    """True when a ledger label carries no identifying information.

    A label is generic if, after removing generic words and short numbers
    (day/month fragments), nothing remains. Long digit runs (possible phone
    numbers or references) count as information.
    """
    for token in tokens(text):
        if token in GENERIC_LABEL_TOKENS:
            continue
        if token.isdigit() and len(token) < 5:
            continue
        return False
    return True
