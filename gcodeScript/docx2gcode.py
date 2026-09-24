#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Точка входу для droplet і ручного запуску.

Пакет лежить у src/docx2gcode — окремо від цього файлу, щоб
``import docx2gcode`` не зчепився з самим лаунчером.
"""

from __future__ import annotations

import sys
from pathlib import Path

_SRC = Path(__file__).resolve().parent / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

from docx2gcode.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
