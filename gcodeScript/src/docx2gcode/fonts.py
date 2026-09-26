"""Пошук системних шрифтів і перетворення тексту на полілінії.

Інтерфейс модуля: find_font, FontEngine, resolve_font.
Індекс і скелетизація сховані всередині — CLI і розкладка їх не чіпають.
"""

from __future__ import annotations

import json
import math
import os
import re
import sys
import unicodedata
from dataclasses import replace
from pathlib import Path

from .config import CACHE_VERSION, PROJECT_DIR, PT_PER_MM, Config
from .geom import Polyline, chaikin, dist, rdp
from .centerline import PolygonPen, centerline
from ._fonttools import TTFont

def _font_dirs() -> list[Path]:
    home = Path.home()
    # Тека самого скрипта — щоб покласти .ttf/.otf поруч і не встановлювати його
    # в систему: шрифт подорожує разом із проєктом. Обхід рекурсивний, тому
    # підтека fonts/ теж працює. Відлік від __file__, а не від робочої теки, бо
    # «Створити G-code.app» запускається з «/».
    project = [PROJECT_DIR, Path(__file__).resolve().parent]
    if sys.platform == "darwin":
        candidates = project + [
            home / "Library/Fonts",
            Path("/Library/Fonts"),
            Path("/System/Library/Fonts"),
            Path("/System/Library/Fonts/Supplemental"),
            Path("/Network/Library/Fonts"),
        ]
    elif os.name == "nt":
        win = Path(os.environ.get("WINDIR", r"C:\Windows"))
        local = os.environ.get("LOCALAPPDATA")
        candidates = project + [win / "Fonts"]
        if local:
            candidates.append(Path(local) / "Microsoft/Windows/Fonts")
    else:
        candidates = project + [
            home / ".fonts",
            home / ".local/share/fonts",
            Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts"),
        ]
    return [c for c in candidates if c.is_dir()]


def _cache_path() -> Path:
    if sys.platform == "darwin":
        base = Path.home() / "Library/Caches"
    elif os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA", Path.home()))
    else:
        base = Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache"))
    d = base / "docx2gcode"
    d.mkdir(parents=True, exist_ok=True)
    return d / "fontindex.json"


def _font_names(path: Path, index: int) -> tuple[set[str], str]:
    """Назви шрифта (родина, повна, PostScript) та накреслення (nameID 2)."""
    names: set[str] = set()
    style = ""
    try:
        tt = TTFont(str(path), fontNumber=index, lazy=True)
        table = tt["name"]
        for rec in table.names:
            try:
                value = str(rec.toUnicode()).strip()
            except Exception:
                continue
            if not value:
                continue
            if rec.nameID in (1, 4, 6, 16):
                names.add(value)
            elif rec.nameID == 2 and not style:
                style = value
    except Exception:
        pass
    return {n for n in names if n}, style


def build_font_index(rebuild: bool = False) -> list[dict]:
    """Індексує шрифти системи (з кешем), повертає [{path, index, names}]."""
    cache = _cache_path()
    dirs = _font_dirs()
    signature = sorted(
        (str(d), int(d.stat().st_mtime)) for d in dirs if d.exists()
    )
    if cache.exists() and not rebuild:
        try:
            data = json.loads(cache.read_text("utf-8"))
            if data.get("version") == CACHE_VERSION and data.get("signature") == signature:
                return data["fonts"]
        except Exception:
            pass

    fonts: list[dict] = []
    seen: set[str] = set()
    for d in dirs:
        for ext in ("*.ttf", "*.otf", "*.ttc", "*.otc", "*.TTF", "*.OTF", "*.TTC"):
            for f in d.rglob(ext):
                key = str(f)
                if key in seen:
                    continue
                seen.add(key)
                count = 1
                if f.suffix.lower() in (".ttc", ".otc"):
                    try:
                        from fontTools.ttLib import TTCollection

                        with TTCollection(str(f), lazy=True) as coll:
                            count = len(coll.fonts)
                    except Exception:
                        count = 1
                for i in range(count):
                    names, style = _font_names(f, i)
                    if names:
                        fonts.append(
                            {"path": key, "index": i, "names": sorted(names), "style": style}
                        )

    try:
        cache.write_text(
            json.dumps(
                {"version": CACHE_VERSION, "signature": signature, "fonts": fonts},
                ensure_ascii=False,
            ),
            "utf-8",
        )
    except Exception:
        pass
    return fonts


def _normalise(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", unicodedata.normalize("NFKD", s).lower())


def find_font(query: str, rebuild: bool = False) -> tuple[Path, int] | None:
    """Знаходить шрифт за назвою родини або повною назвою."""
    p = Path(query)
    if p.exists() and p.is_file():
        return p, 0

    target = _normalise(query)
    fonts = build_font_index(rebuild)

    # Запит на кшталт «Times New Roman» точно збігається і з родиною
    # звичайного накреслення, і з родиною Bold Italic. За інших рівних
    # умов беремо пряме (Regular) накреслення, якщо стиль не названо явно.
    styled_query = any(
        w in target for w in ("bold", "italic", "oblique", "light", "book", "regular")
    )
    regular = {"regular", "roman", "book", "normal", "plain", ""}

    scored: list[tuple[tuple[int, int, int], Path, int]] = []
    for entry in fonts:
        names = {_normalise(n) for n in entry["names"]}
        if target in names:
            rank = 0
        elif any(target in n for n in names):
            rank = 1
        else:
            continue
        style = _normalise(entry.get("style", ""))
        style_rank = 0 if (styled_query or style in regular) else 1
        scored.append(((rank, style_rank, len(str(entry["path"]))),
                       Path(entry["path"]), entry["index"]))

    if not scored:
        return None
    scored.sort(key=lambda t: t[0])
    return scored[0][1], scored[0][2]


class FontEngine:
    """Перетворює текст на полілінії у міліметрах."""

    def __init__(self, path: Path, index: int, size_pt: float, cfg: Config):
        self.path = path
        self.cfg = cfg
        self.size_pt = size_pt
        self.font = TTFont(str(path), fontNumber=index)
        self.glyph_set = self.font.getGlyphSet()
        self.cmap = self.font.getBestCmap()
        self.upem = self.font["head"].unitsPerEm
        self.hmtx = self.font["hmtx"]
        # мм на одиницю шрифта
        self.scale = (size_pt / PT_PER_MM) / self.upem
        self.kern = self._load_kern()
        self._glyph_cache: dict[str, list[Polyline]] = {}
        self.fallback: FontEngine | None = None
        self.missing: set[str] = set()

    def derive(self, size_pt: float) -> FontEngine:
        """Той самий шрифт з іншим кеглем — без повторного читання TTF."""
        if size_pt == self.size_pt:
            return self
        other = self._clone(size_pt=size_pt, word_spacing=self.cfg.word_spacing)
        other.fallback = self.fallback.derive(size_pt) if self.fallback is not None else None
        return other

    def with_word_spacing(self, word_spacing: float) -> FontEngine:
        """Той самий кегль, інша додаткова ширина пробілу."""
        if word_spacing == self.cfg.word_spacing:
            return self
        other = self._clone(size_pt=self.size_pt, word_spacing=word_spacing)
        other.fallback = (
            self.fallback.with_word_spacing(word_spacing) if self.fallback is not None else None
        )
        return other

    def _clone(self, size_pt: float, word_spacing: float) -> FontEngine:
        other = FontEngine.__new__(FontEngine)
        other.path = self.path
        other.cfg = replace(self.cfg, font_size_pt=size_pt, word_spacing=word_spacing)
        other.size_pt = size_pt
        other.font = self.font
        other.glyph_set = self.glyph_set
        other.cmap = self.cmap
        other.upem = self.upem
        other.hmtx = self.hmtx
        other.scale = (size_pt / PT_PER_MM) / other.upem
        other.kern = self.kern
        other._glyph_cache = {}
        other.missing = self.missing
        other.fallback = None
        return other

    # -- метрики -------------------------------------------------------------------

    def _load_kern(self) -> dict[tuple[str, str], int]:
        if not self.cfg.kerning or "kern" not in self.font:
            return {}
        table: dict[tuple[str, str], int] = {}
        try:
            for sub in self.font["kern"].kernTables:
                table.update(sub.kernTable)
        except Exception:
            return {}
        return table

    def glyph_name(self, ch: str) -> str | None:
        return self.cmap.get(ord(ch))

    def has(self, ch: str) -> bool:
        return self.glyph_name(ch) is not None

    def advance(self, ch: str) -> float:
        """Ширина символу в мм."""
        name = self.glyph_name(ch)
        if name is None:
            if self.fallback is not None and self.fallback.has(ch):
                return self.fallback.advance(ch)
            return 0.0
        adv = self.hmtx[name][0] * self.scale
        adv += self.cfg.letter_spacing
        if ch == " ":
            adv += self.cfg.word_spacing
        return adv

    def kern_between(self, a: str, b: str) -> float:
        if not self.kern:
            return 0.0
        ga, gb = self.glyph_name(a), self.glyph_name(b)
        if ga is None or gb is None:
            return 0.0
        return self.kern.get((ga, gb), 0) * self.scale

    def text_width(self, text: str) -> float:
        total = 0.0
        for i, ch in enumerate(text):
            total += self.advance(ch)
            if i + 1 < len(text):
                total += self.kern_between(ch, text[i + 1])
        return total

    @property
    def ascent_mm(self) -> float:
        return self.font["hhea"].ascender * self.scale

    # -- контури -------------------------------------------------------------------

    def glyph_contours(self, name: str) -> list[Polyline]:
        cached = self._glyph_cache.get(name)
        if cached is not None:
            return cached
        # Для осьової лінії потрібен щільніший обхід контуру.
        tol = max(self.cfg.simplify, 1e-4)
        step = self.cfg.step
        if self.cfg.mode == "centerline":
            tol = min(tol, 1.0 / self.cfg.raster_ppmm / 2.0)
            step = min(step, 1.0 / self.cfg.raster_ppmm)
        pen = PolygonPen(self.glyph_set, self.scale, step, tol)
        try:
            self.glyph_set[name].draw(pen)
        except Exception:
            self._glyph_cache[name] = []
            return []
        contours = pen.finish()

        if self.cfg.mode == "centerline" and contours:
            contours = centerline(contours, self.cfg)

        t = 0.25 * max(0.0, min(1.0, self.cfg.smooth))
        # Скелет після потоншення має сходинки в один піксель; якщо берегти
        # їх як «кути», згладжування нічого не дасть. Тому в режимі centerline
        # кутом вважається лише різкий злам.
        angle = self.cfg.corner_angle
        if angle is None:
            angle = 140.0 if self.cfg.mode == "centerline" else 50.0
        corner_cos = math.cos(math.radians(angle))
        processed: list[Polyline] = []
        for c in contours:
            closed = dist(c[0], c[-1]) < 1e-9
            pts = c[:-1] if closed else c
            pts = chaikin(pts, closed, t, corner_cos, iters=2)
            if closed:
                pts = pts + [pts[0]]
            pts = rdp(pts, self.cfg.simplify)
            if len(pts) >= 2:
                processed.append(pts)
        self._glyph_cache[name] = processed
        return processed

    def render(self, text: str, x: float, baseline_y: float, extra_space: float = 0.0) -> list[Polyline]:
        """Полілінії тексту; x, baseline_y — у мм у системі сторінки (Y вниз)."""
        out: list[Polyline] = []
        pen_x = x
        slant = math.tan(math.radians(self.cfg.slant))
        for i, ch in enumerate(text):
            engine = self
            name = self.glyph_name(ch)
            if name is None and self.fallback is not None:
                name = self.fallback.glyph_name(ch)
                if name is not None:
                    engine = self.fallback
            if name is None:
                if not ch.isspace():
                    self.missing.add(ch)
                pen_x += self.advance(ch)
                if ch == " ":
                    pen_x += extra_space
                continue
            if not ch.isspace():
                for contour in engine.glyph_contours(name):
                    poly = [
                        (pen_x + px + slant * py, baseline_y - py)
                        for px, py in contour
                    ]
                    out.append(poly)
            pen_x += engine.advance(ch)
            if ch == " ":
                pen_x += extra_space
            if i + 1 < len(text):
                pen_x += self.kern_between(ch, text[i + 1])
        return out



def resolve_font(cfg: Config, size_pt: float, rebuild: bool) -> FontEngine:
    found = find_font(cfg.font_query, rebuild)
    if found is None:
        raise SystemExit(
            f"Шрифт «{cfg.font_query}» не знайдено в системі.\n"
            "Варіанти:\n"
            "  • покладіть файл шрифта в теку проєкту — він знайдеться за назвою\n"
            "  • вкажіть шлях до файлу:  --font /шлях/до/шрифта.ttf\n"
            "  • подивіться, що є в системі:  --list-fonts"
        )
    path, index = found
    if cfg.font_index:
        index = cfg.font_index
    engine = FontEngine(path, index, size_pt, cfg)

    if cfg.fallback_font_query:
        fb = find_font(cfg.fallback_font_query, rebuild)
        if fb is None:
            print(
                f"Попередження: запасний шрифт «{cfg.fallback_font_query}» не знайдено.",
                file=sys.stderr,
            )
        else:
            engine.fallback = FontEngine(fb[0], fb[1], size_pt, cfg)
    return engine
