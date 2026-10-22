"""Entry point of BinaaCompare.exe (windowed): the simple two-file comparison."""

from __future__ import annotations

import sys


class _NullStream:
    def write(self, text: str) -> int:
        return len(text)

    def flush(self) -> None:
        return None


def main() -> int:
    if sys.stdout is None:
        sys.stdout = _NullStream()
    if sys.stderr is None:
        sys.stderr = _NullStream()
    from binaa_reconciler.compare.ui import main as run

    run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
