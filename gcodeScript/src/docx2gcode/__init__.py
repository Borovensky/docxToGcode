"""Конвертер заповненої DOCX-таблиці у G-code для письмового плотера."""

from .cli import main
from .config import Config

__all__ = ["Config", "main"]
