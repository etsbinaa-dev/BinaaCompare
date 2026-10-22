"""Detection of identical rows. Duplicates are flagged, never removed:
two identical payments on the same day can be perfectly legitimate."""

from __future__ import annotations

from collections.abc import Hashable, Sequence


def find_duplicate_groups(keys: Sequence[Hashable]) -> list[list[int]]:
    """Positions of rows sharing the same key, for every key seen more than once."""
    groups: dict[Hashable, list[int]] = {}
    for pos, key in enumerate(keys):
        groups.setdefault(key, []).append(pos)
    return [g for g in groups.values() if len(g) > 1]
