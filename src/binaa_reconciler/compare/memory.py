"""The user's ✓ / ✗ decisions, kept on this computer only.

Written after every decision (a crash or a power cut loses nothing) to one
JSON file in the user's local profile. A decision is stored by what its lines
say (date, side, amount, description), not by Excel row, so it still applies
when a file is exported again with rows moved. Never sent anywhere.
"""

from __future__ import annotations

import json
import os
from collections.abc import Sequence
from pathlib import Path
from typing import Any

from binaa_reconciler.compare.engine import Kind, Line, Pair

ENV_OVERRIDE = "BINAA_COMPARE_DIR"  # tests: a temporary folder

Decision = tuple[bool, Pair]  # (confirmed?, pair)


def folder() -> Path:
    if custom := os.environ.get(ENV_OVERRIDE):
        return Path(custom)
    base = os.environ.get("LOCALAPPDATA")
    return Path(base) / "BinaaCompare" if base else Path.home() / ".binaa_compare"


def _file() -> Path:
    return folder() / "decisions.json"


def _sig(x: Line) -> list[Any]:
    return [x.day.isoformat(), x.side.value, x.amount, x.label]


def _read() -> dict[str, Any]:
    try:
        data = json.loads(_file().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def _name(names: tuple[str, str]) -> str:
    return f"{names[0]}|{names[1]}"


def save(names: tuple[str, str], decisions: Sequence[Decision]) -> None:
    """All decisions of these two parties, in order (atomic write)."""
    data = _read()
    data[_name(names)] = [
        {
            "ok": ok,
            "kind": pair.kind.value,
            "a": [_sig(x) for x in pair.a],
            "b": [_sig(x) for x in pair.b],
        }
        for ok, pair in decisions
    ]
    path = _file()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, path)


def load(names: tuple[str, str], a: Sequence[Line], b: Sequence[Line]) -> list[Decision]:
    """The saved decisions whose lines are all still in the files."""
    out: list[Decision] = []
    used: set[str] = set()

    def find(sigs: Any, lines: Sequence[Line]) -> tuple[Line, ...] | None:
        found: list[Line] = []
        for sig in sigs:
            x = next(
                (x for x in lines if x.key not in used and _sig(x) == sig and x not in found),
                None,
            )
            if x is None:
                return None
            found.append(x)
        return tuple(found)

    for entry in _read().get(_name(names), []):
        try:
            kind = Kind(entry["kind"])
            pa, pb = find(entry["a"], a), find(entry["b"], b)
        except (KeyError, TypeError, ValueError):
            continue
        if pa is None or pb is None:
            continue
        ok = bool(entry["ok"])
        if ok:  # a confirmed line belongs to one pair only
            used.update(x.key for x in (*pa, *pb))
        out.append((ok, Pair(kind, pa, pb)))
    return out


def forget(names: tuple[str, str]) -> None:
    save(names, [])


def _history_file() -> Path:
    return folder() / "sales_history.json"


def sales_history() -> dict[str, Any]:
    """The accounts' usual entries, learned from the sales reports reviewed."""
    try:
        data = json.loads(_history_file().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def save_sales_history(data: dict[str, Any]) -> None:
    path = _history_file()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, path)


def price_history() -> dict[str, Any]:
    """The products' usual prices, learned from the price files checked."""
    try:
        data = json.loads((folder() / "prices.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def save_price_history(data: dict[str, Any]) -> None:
    path = folder() / "prices.json"
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, path)


def load_json(name: str) -> dict[str, Any]:
    """A small file of what was learned (on this computer only)."""
    try:
        data = json.loads((folder() / name).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data if isinstance(data, dict) else {}


def save_json(name: str, data: dict[str, Any]) -> None:
    path = folder() / name
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, path)


def _settings() -> Path:
    return folder() / "settings.json"


def last_file(key: str) -> Path | None:
    """The file last chosen for side `key` ("A" / "B"), if it is still there."""
    try:
        data = json.loads(_settings().read_text(encoding="utf-8"))
        path = Path(data["files"][key])
    except (OSError, ValueError, KeyError, TypeError):
        return None
    return path if path.is_file() else None


def remember_file(key: str, path: Path) -> None:
    try:
        data = json.loads(_settings().read_text(encoding="utf-8"))
        if not isinstance(data, dict) or not isinstance(data.get("files"), dict):
            raise ValueError
    except (OSError, ValueError):
        data = {"files": {}}
    data["files"][key] = str(Path(path).resolve())
    target = _settings()
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, target)


__all__ = [
    "ENV_OVERRIDE",
    "Decision",
    "folder",
    "forget",
    "last_file",
    "load",
    "load_json",
    "price_history",
    "remember_file",
    "sales_history",
    "save",
    "save_json",
    "save_price_history",
    "save_sales_history",
]
