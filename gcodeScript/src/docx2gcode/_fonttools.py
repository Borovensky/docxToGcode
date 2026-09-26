"""Єдине місце, де вимагаємо fontTools — решта модулів імпортують звідси."""

from __future__ import annotations

import sys

try:
    from fontTools.ttLib import TTFont
    from fontTools.pens.basePen import BasePen
except ImportError:  # pragma: no cover
    sys.exit(
        "Не знайдено бібліотеку fontTools.\n"
        "Встановіть її командою:\n\n"
        "    python3 -m pip install fonttools\n"
    )

__all__ = ["TTFont", "BasePen"]
