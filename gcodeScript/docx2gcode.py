#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx2gcode.py — конвертер заповненої DOCX-таблиці у G-code для письмового плотера.

Читає текст із заповненої таблиці (будь-який .docx), розкладає його
по клітинках порожнього бланка (типово "1 лист.docx") і генерує G-code, яким плотер
пише текст рукописним шрифтом (типово Segoe Script із теки проєкту) на бланку.

Залежності:  pip install fonttools
Платформи:   macOS, Windows, Linux
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Iterable, Iterator, Sequence

try:
    from fontTools.ttLib import TTFont
    from fontTools.pens.basePen import BasePen
except ImportError:  # pragma: no cover
    sys.exit(
        "Не знайдено бібліотеку fontTools.\n"
        "Встановіть її командою:\n\n"
        "    python3 -m pip install fonttools\n"
    )

# --------------------------------------------------------------------------------------
# Константи та одиниці
# --------------------------------------------------------------------------------------

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

CACHE_VERSION = 2  # схема кешу шрифтів; змінюйте при зміні полів

#: Бланк, який використовується, якщо -t не задано.
DEFAULT_TEMPLATE = "1 лист.docx"

#: Тека для G-code, якщо -o не задано (поруч із текою проєкту).
DEFAULT_OUTPUT_DIR = "gcodeOutput"

TWIPS_PER_MM = 1440.0 / 25.4  # 56.6929
PT_PER_MM = 72.0 / 25.4  # 2.83465

#: Ліміт файлів .gcode. Кегль підбирається так, щоб зайняти якомога більше
#: аркушів, але не вийти за цей ліміт.
DEFAULT_MAX_PAGES = 496

#: Верхня межа кегля письма (як <w:sz> = 24 у Word — півпункти).
MAX_FONT_SIZE_PT = 12.0
MIN_FONT_SIZE_PT = 6.0
FONT_SIZE_STEP_PT = 0.05


def _snap_font_pt(size_pt: float) -> float:
    return round(round(size_pt / FONT_SIZE_STEP_PT) * FONT_SIZE_STEP_PT, 2)

#: Метрики Times New Roman на випадок, якщо шрифт не знайдено в системі.
TNR_FALLBACK_METRICS = dict(upem=2048, ascent=1825, descent=443, line_gap=87)

#: Типографські символи, яких зазвичай немає в декоративних шрифтах.
TYPOGRAPHIC_SUBSTITUTES = {
    "\u2019": "'", "\u2018": "'", "\u201a": ",", "\u201b": "'",
    "\u201c": '"', "\u201d": '"', "\u201e": '"', "\u00ab": '"', "\u00bb": '"',
    "\u2013": "-", "\u2014": "-", "\u2212": "-", "\u2010": "-", "\u2011": "-",
    "\u2026": "...", "\u00a0": " ", "\u2007": " ", "\u202f": " ", "\u2009": " ",
    "\u2116": "N", "\u00b0": "o", "\u2033": '"', "\u2032": "'",
}


# --------------------------------------------------------------------------------------
# Конфігурація
# --------------------------------------------------------------------------------------


@dataclass
class Config:
    """Усі параметри перетворення. Значення за замовчуванням — зі специфікації плотера."""

    # --- плотер ---
    feed: float = 2000.0  # мм/хв, робоча подача (знижена через глибокий натиск)
    feed_travel: float = 3000.0  # мм/хв, холості переміщення (G0 ігнорує, але лишаємо)
    feed_z: float = 2000.0  # мм/хв, подача по Z
    z_safe: float = 1.0  # мм, висота підйому пера
    z_draw: float = -0.4  # мм, сила натиску (заглиблення; має бути нижче паперу)
    
    # --- шрифт ---
    #: Назва родини, а не імʼя файлу: «Segoe_Script.ttf» лежить у теці проєкту
    #: й індексується разом із системними шрифтами (див. _font_dirs).
    font_query: str = "Segoe Script"
    font_path: Path | None = None
    font_index: int = 0
    fallback_font_query: str | None = None
    font_size_pt: float | None = None  # None → взяти з файлу-джерела, не більше MAX_FONT_SIZE_PT
    smooth: float = 0.8  # згладжування сплайном, 0..1
    step: float = 0.4  # мм, крок дискретизації кривих
    corner_angle: float | None = None  # градуси; None → 50 для outline, 140 для centerline
    simplify: float = 0.02  # мм, допуск спрощення (прибирає зайві точки)
    mode: str = "centerline"  # centerline | outline
    raster_ppmm: float = 48.0  # роздільність растра для осьової лінії, px/мм
    prune: float = 0.35  # мм, довжина відгалужень, які прибираються
    dot_size: float = 0.9  # мм, елементи менші за це лишаються контуром (крапки)
    slant: float = 0.0  # градуси додаткового нахилу (штучний курсив)
    letter_spacing: float = 0.0  # мм, додатковий трекінг
    word_spacing: float = 0.0  # мм, додаткова ширина пробілу
    kerning: bool = True
    substitute: bool = True

    # --- сторінка ---
    page_w: float = 294.0  # мм
    page_h: float = 210.0  # мм 
    offset_x: float = -6.0  # мм, загальний зсув тексту (мінус = ліворуч)
    offset_y: float = -3.0  # мм, загальний зсув тексту і сітки (мінус = вгору)
    date_offset_x: float = -2.0  # мм, додатковий зсув колонки дати
    number_offset_x: float = -5.0  # мм, додатковий зсув колонки номерів
    baseline_offset: float = 0.0  # мм, ручне підстроювання базової лінії
    origin: str = "bottom-left"  # bottom-left | top-left
    mirror_x: bool = False
    #: Додаток до offset_y для парних аркушів (2-го, 4-го…), мм; мінус = вгору.
    #: Потрібен не через розкладку, а через друк: надрукований бланк парних
    #: аркушів лежить на папері вище, ніж непарних, тому запис на них теж треба
    #: підняти. Непарні аркуші не змінюються. 0 — усі аркуші однакові.
    even_offset_y: float = -7.0
    #: Компенсація вертикального недоходу плотера: він проходить по Y трохи
    #: менше, ніж наказано, тому рядки поступово «утікають» угору від бланка —
    #: на першому збігаються, до кінця аркуша набігає майже цілий рядок.
    #: Множник розтягує вертикальні відстані, відлічені від базової лінії
    #: першого рядка (там розбіжності немає). 1.0 — без компенсації.
    scale_y: float = 1.031

    # --- розкладка ---
    #: Скільки перших рядків таблиці пропустити перед письмом. None → визначити
    #: за наявністю тексту. Типово 2: перший рядок — заголовок, другий уже
    #: заповнений у бланку, тому запис починається з третього рядка.
    header_rows: int | None = 2
    start_number: int = 1
    number_column: bool = True
    justify: bool = True
    split_records: bool = True  # дозволяти розривати запис між аркушами
    max_rows: int | None = None  # None → з бланка
    #: Ліміт файлів .gcode. Кегль зменшується (від 12 pt), доки журнал
    #: не ввійде в цей ліміт. None або 0 — без обмеження.
    max_pages: int | None = DEFAULT_MAX_PAGES
    grid: str = "none"  # none | calib | full
    optimize: bool = True  # перевпорядкування контурів для коротших холостих ходів

    # --- вивід ---
    comments: bool = True  # писати заголовкові коментарі у G-code
    single_file: bool = False  # усі аркуші в один файл із паузою M0
    preview: Path | None = None  # явний шлях; None → поруч із G-code
    write_preview: bool = True  # чи створювати SVG-прев'ю
    verbose: bool = False


# --------------------------------------------------------------------------------------
# Геометрія: прості типи
# --------------------------------------------------------------------------------------

Point = tuple[float, float]
Polyline = list[Point]


def mm_from_twips(v: float) -> float:
    return v / TWIPS_PER_MM


def dist(a: Point, b: Point) -> float:
    return math.hypot(b[0] - a[0], b[1] - a[1])


# --------------------------------------------------------------------------------------
# Читання DOCX
# --------------------------------------------------------------------------------------


def _attr(el: ET.Element | None, name: str) -> str | None:
    return None if el is None else el.get(W + name)


def _val(parent: ET.Element | None, tag: str) -> str | None:
    if parent is None:
        return None
    el = parent.find(W + tag)
    return None if el is None else el.get(W + "val")


def _load_part(path: Path, part: str) -> ET.Element | None:
    with zipfile.ZipFile(path) as zf:
        try:
            return ET.fromstring(zf.read(part))
        except KeyError:
            return None


def _paragraph_text(p: ET.Element) -> str:
    """Текст абзацу з урахуванням <w:tab/> і <w:br/>."""
    out: list[str] = []
    for node in p.iter():
        tag = node.tag
        if tag == W + "t":
            out.append(node.text or "")
        elif tag == W + "tab":
            out.append("\t")
        elif tag == W + "br":
            out.append("\n")
    return "".join(out)


def _cell_paragraphs(tc: ET.Element) -> list[str]:
    """Список абзаців клітинки (порожні абзаці зберігаються)."""
    result: list[str] = []
    for p in tc.findall(W + "p"):
        text = _paragraph_text(p)
        for chunk in text.split("\n"):
            result.append(chunk)
    return result


#: Дата й час в одному абзаці колонки 2. Час відокремлюємо лише коли він іде
#: одразу за датою; «(?!\.\d)» не дає прийняти за час початок другої дати
#: («16.07.2026»), але лишає діапазон («04:36-04:38») цілим.
_DATE_TIME_RE = re.compile(
    r"^(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})[,;]?\s+(\d{1,2}[:.]\d{2}(?!\.\d).*)$"
)


def split_date_time(lines: Iterable[str]) -> list[str]:
    """Дата й час — завжди окремими рядками, як би їх не набрали в джерелі."""
    result: list[str] = []
    for line in lines:
        match = _DATE_TIME_RE.match(line)
        if match:
            result.extend(match.groups())
        else:
            result.append(line)
    return result


@dataclass
class SourceRecord:
    """Один запис (рядок) із заповненої таблиці."""

    date_lines: list[str]
    text: str
    note: str = ""
    number: str = ""


def read_source(path: Path) -> list[SourceRecord]:
    """Витягує записи із заповненої таблиці."""
    root = _load_part(path, "word/document.xml")
    if root is None:
        raise SystemExit(f"Не вдалося прочитати {path}")
    body = root.find(W + "body")
    if body is None:
        raise SystemExit(f"У {path} немає тіла документа")

    records: list[SourceRecord] = []
    for tbl in body.findall(W + "tbl"):
        for tr in tbl.findall(W + "tr"):
            cells = tr.findall(W + "tc")
            if len(cells) < 3:
                continue
            texts = [_cell_paragraphs(tc) for tc in cells]
            number = " ".join(x.strip() for x in texts[0] if x.strip())
            date_lines = split_date_time(x.strip() for x in texts[1] if x.strip())
            main = " ".join(x.strip() for x in texts[2] if x.strip())
            note = " ".join(x.strip() for x in texts[3] if x.strip()) if len(texts) > 3 else ""

            if not date_lines and not main:
                continue  # порожній або службовий рядок
            # Пропускаємо рядок-заголовок на кшталт "1 | 2 | 3 | 4"
            if main.strip() in {"3"} and number.strip() in {"1"}:
                continue
            main = re.sub(r"\s+", " ", main).strip()
            records.append(
                SourceRecord(date_lines=date_lines, text=main, note=note, number=number)
            )
    return records


@dataclass
class TableGeometry:
    """Геометрія бланка, обчислена з порожнього шаблону."""

    page_w: float
    page_h: float
    margin_left: float
    margin_top: float
    col_x: list[float]  # межі колонок, мм від лівого краю сторінки
    header_rows: int
    row_tops: list[float]  # верх кожного рядка даних, мм від верху сторінки
    row_height: float
    cell_margin_left: float
    cell_margin_right: float
    baseline_in_row: float  # від верху рядка до базової лінії, мм
    font_size_pt: float

    @property
    def data_rows(self) -> int:
        return len(self.row_tops)

    def col_bounds(self, index: int) -> tuple[float, float]:
        return self.col_x[index], self.col_x[index + 1]

    def text_bounds(self, index: int) -> tuple[float, float]:
        left, right = self.col_bounds(index)
        return left + self.cell_margin_left, right - self.cell_margin_right


def _font_line_metrics(query: str) -> dict:
    """Метрики шрифта бланка для обчислення висоти рядка."""
    found = find_font(query)
    if found is None:
        return dict(TNR_FALLBACK_METRICS)
    path, index = found
    try:
        tt = TTFont(str(path), fontNumber=index, lazy=True)
        upem = tt["head"].unitsPerEm
        hhea = tt["hhea"]
        return dict(
            upem=upem,
            ascent=hhea.ascender,
            descent=abs(hhea.descender),
            line_gap=hhea.lineGap,
        )
    except Exception:
        return dict(TNR_FALLBACK_METRICS)


def read_template(path: Path, cfg: Config) -> TableGeometry:
    """Обчислює геометрію бланка (розміри сторінки, колонок, рядків)."""
    root = _load_part(path, "word/document.xml")
    if root is None:
        raise SystemExit(f"Не вдалося прочитати {path}")
    body = root.find(W + "body")
    tbl = body.find(W + "tbl") if body is not None else None
    if tbl is None:
        raise SystemExit(f"У шаблоні {path} не знайдено таблиці")

    sect = body.find(W + "sectPr")
    pg_sz = sect.find(W + "pgSz") if sect is not None else None
    pg_mar = sect.find(W + "pgMar") if sect is not None else None

    doc_page_w = mm_from_twips(float(_attr(pg_sz, "w") or 16838))
    doc_page_h = mm_from_twips(float(_attr(pg_sz, "h") or 11906))
    margin_left = mm_from_twips(float(_attr(pg_mar, "left") or 850))
    margin_top = mm_from_twips(float(_attr(pg_mar, "top") or 993))

    tbl_pr = tbl.find(W + "tblPr")
    tbl_ind = mm_from_twips(float(_attr(tbl_pr.find(W + "tblInd"), "w") or 0)) if tbl_pr is not None else 0.0

    # Поля клітинок: з таблиці, інакше стандарт Word (108 twips ліворуч і праворуч).
    cell_left, cell_right = mm_from_twips(108), mm_from_twips(108)
    cell_mar = tbl_pr.find(W + "tblCellMar") if tbl_pr is not None else None
    if cell_mar is not None:
        left_el, right_el = cell_mar.find(W + "left"), cell_mar.find(W + "right")
        if left_el is not None:
            cell_left = mm_from_twips(float(_attr(left_el, "w") or 0))
        if right_el is not None:
            cell_right = mm_from_twips(float(_attr(right_el, "w") or 0))

    grid = tbl.find(W + "tblGrid")
    widths = [float(_attr(gc, "w") or 0) for gc in grid.findall(W + "gridCol")]
    table_left = margin_left + tbl_ind
    col_x = [table_left]
    for w in widths:
        col_x.append(col_x[-1] + mm_from_twips(w))

    # Розмір шрифта бланка: беремо перший явний <w:sz> у клітинках.
    size_half_pt = 24.0
    for sz in tbl.iter(W + "sz"):
        try:
            size_half_pt = float(sz.get(W + "val"))
            break
        except (TypeError, ValueError):
            pass
    template_font_pt = size_half_pt / 2.0

    # Шрифт бланка задає висоту рядка.
    template_font_name = "Times New Roman"
    for rf in tbl.iter(W + "rFonts"):
        name = rf.get(W + "ascii")
        if name:
            template_font_name = name
            break
    metrics = _font_line_metrics(template_font_name)
    upem = metrics["upem"]
    line_factor = (metrics["ascent"] + metrics["descent"] + metrics["line_gap"]) / upem
    ascent_factor = (metrics["ascent"] + metrics["line_gap"]) / upem

    natural_line_mm = template_font_pt / PT_PER_MM * line_factor
    baseline_mm = template_font_pt / PT_PER_MM * ascent_factor

    # Висоти рядків: Word бере max(вказана висота, висота вмісту).
    rows = tbl.findall(W + "tr")
    heights: list[float] = []
    lines_per_row: list[int] = []
    for tr in rows:
        tr_pr = tr.find(W + "trPr")
        spec = 0.0
        rule = "atLeast"
        if tr_pr is not None:
            h_el = tr_pr.find(W + "trHeight")
            if h_el is not None:
                spec = mm_from_twips(float(_attr(h_el, "val") or 0))
                rule = _attr(h_el, "hRule") or "atLeast"
        max_paras = 1
        for tc in tr.findall(W + "tc"):
            max_paras = max(max_paras, max(1, len(tc.findall(W + "p"))))
        content = natural_line_mm * max_paras
        heights.append(spec if rule == "exact" else max(spec, content))
        lines_per_row.append(max_paras)

    # Рядки-заголовки — ті, що містять текст; далі йдуть порожні рядки для запису.
    if cfg.header_rows is not None:
        header_rows = max(0, min(cfg.header_rows, len(rows)))
    else:
        header_rows = 0
        for tr in rows:
            has_text = any(t.text and t.text.strip() for t in tr.iter(W + "t"))
            if has_text:
                header_rows += 1
            else:
                break

    tops: list[float] = []
    y = margin_top
    for i, h in enumerate(heights):
        if i >= header_rows:
            tops.append(y)
        y += h

    data_height = heights[header_rows] if header_rows < len(heights) else natural_line_mm

    if cfg.max_rows is not None:
        tops = tops[: cfg.max_rows]

    return TableGeometry(
        page_w=cfg.page_w or doc_page_w,
        page_h=cfg.page_h or doc_page_h,
        margin_left=margin_left,
        margin_top=margin_top,
        col_x=col_x,
        header_rows=header_rows,
        row_tops=tops,
        row_height=data_height,
        cell_margin_left=cell_left,
        cell_margin_right=cell_right,
        baseline_in_row=baseline_mm + cfg.baseline_offset,
        font_size_pt=template_font_pt,
    )


# --------------------------------------------------------------------------------------
# Пошук шрифтів у системі
# --------------------------------------------------------------------------------------


def _font_dirs() -> list[Path]:
    home = Path.home()
    # Тека самого скрипта — щоб покласти .ttf/.otf поруч і не встановлювати його
    # в систему: шрифт подорожує разом із проєктом. Обхід рекурсивний, тому
    # підтека fonts/ теж працює. Відлік від __file__, а не від робочої теки, бо
    # «Створити G-code.app» запускається з «/».
    project = [Path(__file__).resolve().parent]
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


# --------------------------------------------------------------------------------------
# Обробка контурів: згладжування, спрощення
# --------------------------------------------------------------------------------------


def _lerp(a: Point, b: Point, t: float) -> Point:
    return (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)


def chaikin(pts: Polyline, closed: bool, t: float, corner_cos: float, iters: int) -> Polyline:
    """Згладжування зрізанням кутів (Chaikin) зі збереженням гострих кутів."""
    if t <= 0 or len(pts) < 3:
        return pts
    for _ in range(iters):
        n = len(pts)
        out: Polyline = []
        for i in range(n):
            cur = pts[i]
            has_prev = closed or i > 0
            has_next = closed or i < n - 1
            if not (has_prev and has_next):
                out.append(cur)
                continue
            prev = pts[(i - 1) % n]
            nxt = pts[(i + 1) % n]
            v1 = (cur[0] - prev[0], cur[1] - prev[1])
            v2 = (nxt[0] - cur[0], nxt[1] - cur[1])
            l1 = math.hypot(*v1)
            l2 = math.hypot(*v2)
            if l1 < 1e-12 or l2 < 1e-12:
                out.append(cur)
                continue
            cos_a = (v1[0] * v2[0] + v1[1] * v2[1]) / (l1 * l2)
            if cos_a < corner_cos:  # гострий кут — зберігаємо
                out.append(cur)
                continue
            out.append(_lerp(cur, prev, t))
            out.append(_lerp(cur, nxt, t))
        pts = out
        if len(pts) < 3:
            break
    return pts


def rdp(pts: Polyline, tol: float) -> Polyline:
    """Спрощення Дугласа–Пекера: прибирає точки, що лежать на прямій."""
    if tol <= 0 or len(pts) < 3:
        return pts

    keep = [False] * len(pts)
    keep[0] = keep[-1] = True
    stack = [(0, len(pts) - 1)]
    while stack:
        i0, i1 = stack.pop()
        if i1 <= i0 + 1:
            continue
        ax, ay = pts[i0]
        bx, by = pts[i1]
        dx, dy = bx - ax, by - ay
        norm = math.hypot(dx, dy)
        best_i, best_d = -1, -1.0
        for i in range(i0 + 1, i1):
            px, py = pts[i]
            if norm < 1e-12:
                d = math.hypot(px - ax, py - ay)
            else:
                d = abs(dy * px - dx * py + bx * ay - by * ax) / norm
            if d > best_d:
                best_d, best_i = d, i
        if best_d > tol:
            keep[best_i] = True
            stack.append((i0, best_i))
            stack.append((best_i, i1))
    return [p for p, k in zip(pts, keep) if k]


# --------------------------------------------------------------------------------------
# Осьова лінія (одноштриховий режим)
# --------------------------------------------------------------------------------------


def rasterize(contours: list[Polyline], ppmm: float, pad: int = 3):
    """Растеризує контури за правилом even-odd. Повертає (сітка, ширина, висота, x0, y0)."""
    xs = [p[0] for c in contours for p in c]
    ys = [p[1] for c in contours for p in c]
    if not xs:
        return None
    x0, x1 = min(xs), max(xs)
    y0, y1 = min(ys), max(ys)
    w = int(math.ceil((x1 - x0) * ppmm)) + 2 * pad + 1
    h = int(math.ceil((y1 - y0) * ppmm)) + 2 * pad + 1
    grid = bytearray(w * h)

    # Ребра у піксельних координатах.
    edges: list[tuple[float, float, float, float]] = []
    for c in contours:
        n = len(c)
        closed = dist(c[0], c[-1]) < 1e-9
        pts = c[:-1] if closed else c
        n = len(pts)
        for i in range(n):
            ax, ay = pts[i]
            bx, by = pts[(i + 1) % n]
            edges.append(
                ((ax - x0) * ppmm + pad, (ay - y0) * ppmm + pad,
                 (bx - x0) * ppmm + pad, (by - y0) * ppmm + pad)
            )

    for py in range(h):
        yc = py + 0.5
        crossings: list[float] = []
        for ax, ay, bx, by in edges:
            if (ay <= yc) != (by <= yc):
                t = (yc - ay) / (by - ay)
                crossings.append(ax + t * (bx - ax))
        if not crossings:
            continue
        crossings.sort()
        row = py * w
        for i in range(0, len(crossings) - 1, 2):
            sx = int(math.ceil(crossings[i] - 0.5))
            ex = int(math.floor(crossings[i + 1] - 0.5))
            if ex < sx:
                continue
            sx = max(0, sx)
            ex = min(w - 1, ex)
            for px in range(sx, ex + 1):
                grid[row + px] = 1
    return grid, w, h, x0, y0


_N8 = ((0, -1), (1, -1), (1, 0), (1, 1), (0, 1), (-1, 1), (-1, 0), (-1, -1))


def zhang_suen(grid: bytearray, w: int, h: int) -> bytearray:
    """Потоншення Чжана–Суеня до скелета завтовшки 1 піксель."""
    g = bytearray(grid)
    active = [i for i, v in enumerate(g) if v]
    changed = True
    while changed and active:
        changed = False
        for step in (0, 1):
            to_clear: list[int] = []
            still: list[int] = []
            for idx in active:
                if not g[idx]:
                    continue
                x, y = idx % w, idx // w
                if x <= 0 or y <= 0 or x >= w - 1 or y >= h - 1:
                    still.append(idx)
                    continue
                p = [g[(y + dy) * w + (x + dx)] for dx, dy in _N8]
                b = sum(p)
                if b < 2 or b > 6:
                    still.append(idx)
                    continue
                a = sum(1 for i in range(8) if p[i] == 0 and p[(i + 1) % 8] == 1)
                if a != 1:
                    still.append(idx)
                    continue
                # p = [N, NE, E, SE, S, SW, W, NW]
                n, e, s, wst = p[0], p[2], p[4], p[6]
                if step == 0:
                    ok = (n * e * s == 0) and (e * s * wst == 0)
                else:
                    ok = (n * e * wst == 0) and (n * s * wst == 0)
                if ok:
                    to_clear.append(idx)
                else:
                    still.append(idx)
            for idx in to_clear:
                g[idx] = 0
            if to_clear:
                changed = True
            active = still
    return g


def trace_skeleton(skel: bytearray, w: int, h: int) -> list[list[tuple[int, int]]]:
    """Розбирає скелет на послідовності пікселів (гілки та замкнені петлі)."""
    pixels = {i for i, v in enumerate(skel) if v}
    if not pixels:
        return []

    _cache: dict[int, list[int]] = {}

    def neigh(idx: int) -> list[int]:
        """Сусіди з відкиданням надлишкових діагональних зв'язків.

        У 8-зв'язному скелеті сходинка на кшталт «#» по діагоналі вже з'єднана
        через ортогонального сусіда; без цього фільтра ступінь пікселя
        завищується і шлях розпадається на дволанкові уламки.
        """
        cached = _cache.get(idx)
        if cached is not None:
            return cached
        x, y = idx % w, idx // w
        orth: set[tuple[int, int]] = set()
        out: list[int] = []
        for dx, dy in ((0, -1), (1, 0), (0, 1), (-1, 0)):
            nx, ny = x + dx, y + dy
            if 0 <= nx < w and 0 <= ny < h and (ny * w + nx) in pixels:
                orth.add((dx, dy))
                out.append(ny * w + nx)
        for dx, dy in ((1, -1), (1, 1), (-1, 1), (-1, -1)):
            nx, ny = x + dx, y + dy
            if not (0 <= nx < w and 0 <= ny < h):
                continue
            j = ny * w + nx
            if j not in pixels:
                continue
            if (dx, 0) in orth or (0, dy) in orth:
                continue
            out.append(j)
        _cache[idx] = out
        return out

    degree = {i: len(neigh(i)) for i in pixels}
    nodes = {i for i, d in degree.items() if d != 2}
    visited: set[tuple[int, int]] = set()
    paths: list[list[int]] = []

    def walk(start: int, first: int) -> list[int]:
        path = [start, first]
        visited.add((start, first))
        visited.add((first, start))
        prev, cur = start, first
        while degree.get(cur, 0) == 2:
            nxt = [n for n in neigh(cur) if n != prev]
            if not nxt:
                break
            n = nxt[0]
            if (cur, n) in visited:
                break
            visited.add((cur, n))
            visited.add((n, cur))
            path.append(n)
            prev, cur = cur, n
        return path

    for node in sorted(nodes):
        ns = neigh(node)
        if not ns:
            # Ізольований піксель (крапка, тире) — залишаємо як точковий слід.
            paths.append([node, node])
            continue
        for n in ns:
            if (node, n) not in visited:
                paths.append(walk(node, n))

    # Замкнені петлі без вузлів (наприклад, «о»).
    remaining = pixels - {p for path in paths for p in path}
    while remaining:
        start = min(remaining)
        ns = [n for n in neigh(start) if (start, n) not in visited]
        if not ns:
            remaining.discard(start)
            continue
        path = walk(start, ns[0])
        if path[-1] != start and start in neigh(path[-1]):
            path.append(start)
        paths.append(path)
        remaining -= set(path)

    return [[(p % w, p // w) for p in path] for path in paths if len(path) >= 2]


def join_paths(polys: list[Polyline], angle_limit: float = 135.0) -> list[Polyline]:
    """Зшиває штрихи зі спільними кінцями в довші, щоб перо рідше відривалося.

    З'єднуються лише шляхи, які вже дотикаються, тому малюнок не змінюється —
    змінюється тільки порядок і напрямок обходу.
    """
    paths = [list(p) for p in polys if len(p) >= 2]
    if len(paths) < 2:
        return paths
    limit_cos = math.cos(math.radians(angle_limit))

    def kf(pt: Point) -> tuple[float, float]:
        return (round(pt[0], 3), round(pt[1], 3))

    def outward(path: Polyline, at_end: bool) -> Point:
        a, b = (path[-2], path[-1]) if at_end else (path[1], path[0])
        dx, dy = b[0] - a[0], b[1] - a[1]
        n = math.hypot(dx, dy) or 1.0
        return dx / n, dy / n

    alive = dict(enumerate(paths))
    result: list[Polyline] = []
    while alive:
        cur = alive.pop(next(iter(alive)))
        extended = True
        while extended:
            extended = False
            for at_end in (True, False):
                tip = kf(cur[-1] if at_end else cur[0])
                dvec = outward(cur, at_end)
                best, best_cos, best_rev = None, limit_cos, False
                for j, q in alive.items():
                    for q_end in (False, True):
                        if kf(q[-1] if q_end else q[0]) != tip:
                            continue
                        # Орієнтуємо q так, щоб він починався зі спільної точки,
                        # і беремо напрямок, у якому він іде далі від стику.
                        oriented = q[::-1] if q_end else q
                        qd = outward(oriented, at_end=False)
                        qd = (-qd[0], -qd[1])
                        c = dvec[0] * qd[0] + dvec[1] * qd[1]
                        if c > best_cos:
                            best, best_cos, best_rev = j, c, q_end
                if best is None:
                    continue
                q = alive.pop(best)
                if best_rev:
                    q = q[::-1]
                cur = cur + q[1:] if at_end else q[::-1] + cur[1:]
                extended = True

        result.append(cur)
    return result


def _bbox(c: Polyline) -> tuple[float, float, float, float]:
    xs = [p[0] for p in c]
    ys = [p[1] for p in c]
    return min(xs), min(ys), max(xs), max(ys)


def _split_dots(contours: list[Polyline], limit: float) -> tuple[list[Polyline], list[Polyline]]:
    """Відокремлює дрібні самостійні елементи (крапки, двокрапки, крапка над «і»).

    Потоншення перетворює таку цятку на один піксель і вона зникає, тому її
    контур лишаємо як є. Внутрішні контури («вічко» літери «о») не рахуються
    крапками — вони лежать усередині зовнішнього контуру.
    """
    boxes = [_bbox(c) for c in contours]
    dots: list[Polyline] = []
    rest: list[Polyline] = []
    for i, c in enumerate(contours):
        x0, y0, x1, y1 = boxes[i]
        small = max(x1 - x0, y1 - y0) <= limit
        inside = any(
            j != i and bx0 <= x0 and by0 <= y0 and bx1 >= x1 and by1 >= y1
            for j, (bx0, by0, bx1, by1) in enumerate(boxes)
        )
        (dots if (small and not inside) else rest).append(c)
    return dots, rest


def centerline(contours: list[Polyline], cfg: Config) -> list[Polyline]:
    """Перетворює замкнені контури гліфа на одноштрихову осьову лінію."""
    dots, contours = _split_dots(contours, cfg.dot_size)
    if not contours:
        return dots
    raster = rasterize(contours, cfg.raster_ppmm)
    if raster is None:
        return dots
    grid, w, h, x0, y0 = raster
    skel = zhang_suen(grid, w, h)
    paths = trace_skeleton(skel, w, h)

    scale = 1.0 / cfg.raster_ppmm
    pad = 3
    polys: list[Polyline] = []
    for path in paths:
        poly = [((px - pad) * scale + x0, (py - pad) * scale + y0) for px, py in path]
        polys.append(poly)

    # Прибирання коротких відгалужень (артефакти засічок).
    if cfg.prune > 0:
        pruned: list[Polyline] = []
        endpoints: dict[Point, int] = {}
        for p in polys:
            for e in (p[0], p[-1]):
                key = (round(e[0], 4), round(e[1], 4))
                endpoints[key] = endpoints.get(key, 0) + 1
        for p in polys:
            length = sum(dist(p[i], p[i + 1]) for i in range(len(p) - 1))
            if length >= cfg.prune:
                pruned.append(p)
                continue
            a = (round(p[0][0], 4), round(p[0][1], 4))
            b = (round(p[-1][0], 4), round(p[-1][1], 4))
            free_a = endpoints.get(a, 0) == 1
            free_b = endpoints.get(b, 0) == 1
            # Рівно один вільний кінець → це відгалуження, прибираємо.
            # Обидва вільні → самостійний штрих (крапка, тире), лишаємо.
            if free_a != free_b:
                continue
            pruned.append(p)
        polys = pruned

    return join_paths(polys) + dots


class PolygonPen(BasePen):
    """Перо fontTools, що перетворює контури гліфа на полілінії у мм."""

    def __init__(self, glyph_set, scale: float, step: float, tol: float):
        super().__init__(glyph_set)
        self.scale = scale
        self.step = step
        self.tol = tol
        self.contours: list[Polyline] = []
        self._current: Polyline = []

    def _s(self, pt) -> Point:
        return (pt[0] * self.scale, pt[1] * self.scale)

    def _moveTo(self, pt):
        self._flush()
        self._current = [self._s(pt)]

    def _lineTo(self, pt):
        self._current.append(self._s(pt))

    def _curveToOne(self, p1, p2, p3):
        if not self._current:
            return
        p0 = self._current[-1]
        a, b, c = self._s(p1), self._s(p2), self._s(p3)
        n = self._cubic_segments(p0, a, b, c)
        for i in range(1, n + 1):
            t = i / n
            mt = 1 - t
            x = (mt**3 * p0[0] + 3 * mt**2 * t * a[0] + 3 * mt * t**2 * b[0] + t**3 * c[0])
            y = (mt**3 * p0[1] + 3 * mt**2 * t * a[1] + 3 * mt * t**2 * b[1] + t**3 * c[1])
            self._current.append((x, y))

    def _qCurveToOne(self, p1, p2):
        if not self._current:
            return
        p0 = self._current[-1]
        a, b = self._s(p1), self._s(p2)
        n = self._quad_segments(p0, a, b)
        for i in range(1, n + 1):
            t = i / n
            mt = 1 - t
            x = mt**2 * p0[0] + 2 * mt * t * a[0] + t**2 * b[0]
            y = mt**2 * p0[1] + 2 * mt * t * a[1] + t**2 * b[1]
            self._current.append((x, y))

    def _quad_segments(self, p0: Point, p1: Point, p2: Point) -> int:
        chord = dist(p0, p1) + dist(p1, p2)
        n_len = math.ceil(chord / self.step) if self.step > 0 else 1
        d = math.hypot(p0[0] - 2 * p1[0] + p2[0], p0[1] - 2 * p1[1] + p2[1])
        n_tol = math.ceil(math.sqrt(d / (8 * self.tol))) if self.tol > 0 and d > 0 else 1
        return max(1, min(200, max(n_len, n_tol)))

    def _cubic_segments(self, p0: Point, p1: Point, p2: Point, p3: Point) -> int:
        chord = dist(p0, p1) + dist(p1, p2) + dist(p2, p3)
        n_len = math.ceil(chord / self.step) if self.step > 0 else 1
        d1 = math.hypot(p0[0] - 2 * p1[0] + p2[0], p0[1] - 2 * p1[1] + p2[1])
        d2 = math.hypot(p1[0] - 2 * p2[0] + p3[0], p1[1] - 2 * p2[1] + p3[1])
        d = max(d1, d2)
        n_tol = math.ceil(math.sqrt(3 * d / (4 * self.tol))) if self.tol > 0 and d > 0 else 1
        return max(1, min(200, max(n_len, n_tol)))

    def _closePath(self):
        self._flush(close=True)

    def _endPath(self):
        self._flush()

    def _flush(self, close: bool = False):
        pts = self._current
        self._current = []
        if len(pts) < 2:
            return
        if close and dist(pts[0], pts[-1]) > 1e-9:
            pts.append(pts[0])
        self.contours.append(pts)

    def finish(self) -> list[Polyline]:
        self._flush()
        return self.contours


# --------------------------------------------------------------------------------------
# Робота зі шрифтом
# --------------------------------------------------------------------------------------


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
        other = FontEngine.__new__(FontEngine)
        other.path = self.path
        other.cfg = replace(self.cfg, font_size_pt=size_pt)
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
        other.fallback = self.fallback.derive(size_pt) if self.fallback is not None else None
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


# --------------------------------------------------------------------------------------
# Розкладка тексту
# --------------------------------------------------------------------------------------


#: Посилання на документ-джерело в тексті 3-ї колонки. Кожне починає новий
#: рядок, щоб не губилося всередині абзацу. Регістр будь-який.
LINE_BREAK_MARKERS = (
    "розділ",
    "бойове розпорядження",
    "позатермінове",
    "клопотання",
    "розпорядження",
    "БД",
)

#: Розрив ставимо перед дужкою, тому шукаємо позицію нульової ширини — сама
#: дужка лишається на початку нового рядка. «\b» не дає зачепити слово, що лише
#: починається з позначки («(розділів», «(розпорядженням»).
_MARKER_RE = re.compile(
    r"(?=\(\s*(?:" + "|".join(LINE_BREAK_MARKERS) + r")\b)",
    re.IGNORECASE,
)

#: Підпис командира зазвичай уже з нового рядка. «Командир N омбр» —
#: окремий рядок ліворуч; звання «підполковник» ліворуч і прізвище
#: праворуч — наступний рядок. Після блоку — порожній рядок.
COMMANDER_GAP_LINES = 1
_COMMANDER_NAME = r"(?:[A-Za-zА-Яа-яІіЇїЄєҐґ]\.)+|[^\s.,;:()]{2,}"
_COMMANDER_RE = re.compile(
    r"(?P<title>Командир\s+\d+\s+омбр)\s+"
    r"(?P<rank>підполковник)\s+"
    r"(?P<name>(?:(?:" + _COMMANDER_NAME + r")\s+){0,2}?(?:" + _COMMANDER_NAME + r"))"
    r"(?:\.|(?=\s*\()|(?=\s*$))",
    re.IGNORECASE,
)


def split_markers(text: str) -> list[str]:
    """Ріже текст на абзаци перед посиланнями на документ-джерело."""
    return [part.strip() for part in _MARKER_RE.split(text) if part.strip()]


def iter_signature_parts(text: str):
    """Чергує звичайний текст і підпис командира, щоб підпис почав новий рядок."""
    pos = 0
    for match in _COMMANDER_RE.finditer(text):
        before = text[pos:match.start()].strip()
        if before:
            yield "text", before
        yield "commander", match
        pos = match.end()
        if pos < len(text) and text[pos] == ".":
            pos += 1
    tail = text[pos:].strip()
    if tail:
        yield "text", tail


#: Номер документа «N 231» / «№ 50445» — «N» і цифри не розриваємо між рядками.
_DOC_NUM_RE = re.compile(
    r"(?:N|№|No\.?)\s*\d+[.,;:]?",
    re.IGNORECASE,
)
_DOC_MARK_RE = re.compile(r"^(N|№|No\.?)$", re.IGNORECASE)


def wrap_tokens(text: str) -> list[str]:
    """Слова для переносу: «N 231» лишається одним шматком."""
    tokens: list[str] = []
    pos = 0
    for match in _DOC_NUM_RE.finditer(text):
        tokens.extend(text[pos:match.start()].split())
        tokens.append(re.sub(r"\s+", " ", match.group(0)).strip())
        pos = match.end()
    tokens.extend(text[pos:].split())
    return tokens


def _join_split_doc_nums(lines: list[str]) -> list[str]:
    """Якщо рядок обірвався на «N», а наступний починається з номера — звести разом.

    «N» переїжджає на вже існуючий наступний рядок. Нових рядків не зʼявляється.
    """
    if len(lines) < 2:
        return lines
    out = [lines[0]]
    for nxt in lines[1:]:
        prev = out[-1].rstrip()
        head, sep, last = prev.rpartition(" ")
        if not sep:
            head, last = "", prev
        rest = nxt.lstrip()
        if _DOC_MARK_RE.fullmatch(last) and rest[:1].isdigit():
            if head:
                out[-1] = head
                out.append(f"{last} {rest}")
            else:
                out[-1] = f"{last} {rest}"
        else:
            out.append(nxt)
    return out


def wrap_text(text: str, engine: FontEngine, width: float) -> list[str]:
    """Розбиває текст на рядки, що вміщуються в задану ширину."""
    words = wrap_tokens(text)
    if not words:
        return []
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = word if not current else current + " " + word
        if engine.text_width(candidate) <= width or not current:
            current = candidate
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)

    # Слово, довше за колонку, розбиваємо посимвольно — але не «N 231».
    result: list[str] = []
    for line in lines:
        if _DOC_NUM_RE.fullmatch(line.strip()):
            result.append(line)
            continue
        while engine.text_width(line) > width and len(line) > 1:
            cut = len(line)
            while cut > 1 and engine.text_width(line[:cut]) > width:
                cut -= 1
            piece, rest = line[:cut], line[cut:].lstrip()
            if _DOC_MARK_RE.fullmatch(piece.rstrip().rsplit(" ", 1)[-1]) and rest[:1].isdigit():
                # Не лишати голе «N» на рядку — віддати номер разом униз.
                head, sep, _mark = piece.rstrip().rpartition(" ")
                if sep:
                    result.append(head)
                    line = f"{_mark} {rest}"
                    continue
            result.append(piece)
            line = rest
        if line:
            result.append(line)
    return _join_split_doc_nums(result)


def wrap_text_at_least(
    text: str, engine: FontEngine, width: float, min_lines: int
) -> list[str]:
    """Як wrap_text, але займає порожні рядки запису — хоча б одним словом.

    Типовий випадок: дата й час уже взяли два рядки бланка, а текст ліг
    в один — другий рядок колонки лишається порожнім. Трохи звужуємо
    перенесення, щоб слово зʼїхало вниз; якщо не виходить рівно,
    переносимо останні слова примусово.
    """
    lines = wrap_text(text, engine, width)
    if min_lines <= 1 or len(lines) >= min_lines:
        return lines
    words = wrap_tokens(text)
    if len(words) < 2:
        return lines

    for factor in (0.94, 0.88, 0.82, 0.76, 0.70, 0.62, 0.54, 0.45):
        trial = wrap_text(text, engine, width * factor)
        if len(trial) == min_lines:
            return trial
        if len(trial) > min_lines:
            break
    return _force_min_wrap_lines(lines, min_lines)


def _force_min_wrap_lines(lines: list[str], min_lines: int) -> list[str]:
    """Знімає слова з кінця на окремі рядки бланка, зберігаючи порядок."""
    words = [w for line in lines for w in wrap_tokens(line)]
    if len(words) < 2:
        return lines
    extra = min(max(min_lines - 1, 0), len(words) - 1)
    if extra <= 0:
        return lines
    return _join_split_doc_nums([" ".join(words[:-extra])] + words[-extra:])


def _fill_short_record(
    body: list[BodyLine], engine: FontEngine, width: float, min_lines: int
) -> list[BodyLine]:
    """Дописує слова в порожні рядки, які вже зайняті датою/часом."""
    content = [b for b in body if not b.is_blank]
    if len(content) >= min_lines:
        return body
    if any(len(b.fragments) != 1 for b in content):
        return body

    last_para_start = 0
    for i, line in enumerate(body):
        if line.ends_paragraph and i < len(body) - 1:
            last_para_start = i + 1
    para = [b for b in body[last_para_start:] if not b.is_blank]
    if not para or any(len(b.fragments) != 1 for b in para):
        return body

    other = len(content) - len(para)
    need = min_lines - other
    if need <= 1:
        return body
    text = " ".join(frag for b in para for frag, _align in b.fragments)
    new_lines = wrap_text_at_least(text, engine, width, need)
    return body[:last_para_start] + [
        BodyLine([(line, "")], ends_paragraph=i == len(new_lines) - 1)
        for i, line in enumerate(new_lines)
    ]


@dataclass
class BodyLine:
    """Один рядок 3-ї колонки. Кілька фрагментів пишуться в тому ж рядку бланка."""

    fragments: list[tuple[str, str]] = field(default_factory=list)  # (текст, align)
    ends_paragraph: bool = False

    @property
    def is_blank(self) -> bool:
        return not any(text for text, _align in self.fragments)


def wrap_record_text(text: str, engine: FontEngine, width: float) -> list[BodyLine]:
    """Рядки запису: звичайне перенесення, підпис командира, позначки-джерела.

    Абзац закінчується перед позначкою-джерелом, після підпису командира
    або в кінці запису. Такий рядок не розтягується по ширині — інакше
    короткий «хвіст» перед примусовим розривом розповзся б на всю колонку.
    """
    out: list[BodyLine] = []
    for kind, payload in iter_signature_parts(text):
        if kind == "commander":
            title = payload.group("title")
            rank = payload.group("rank")
            name = payload.group("name")
            out.append(BodyLine([(title, "left")], ends_paragraph=True))
            out.append(
                BodyLine([(rank, "left"), (name, "right")], ends_paragraph=True)
            )
            for _ in range(COMMANDER_GAP_LINES):
                out.append(BodyLine())
            continue
        for paragraph in split_markers(payload):
            lines = wrap_text(paragraph, engine, width)
            for i, line in enumerate(lines):
                out.append(
                    BodyLine([(line, "")], ends_paragraph=i == len(lines) - 1)
                )
    return out


@dataclass
class PlacedLine:
    """Один рядок тексту, прив'язаний до рядка бланка."""

    row: int
    column: int
    text: str
    align: str = "left"  # left | center | justify | right
    is_last: bool = False


@dataclass
class Page:
    lines: list[PlacedLine] = field(default_factory=list)


def _wrap_cache_key(rec_i: int, engine: FontEngine) -> tuple[int, float]:
    """Ключ без id(): CPython повторно видає адреси GC-нутих двигунів."""
    return rec_i, engine.size_pt


def _wrapped_body(
    rec_i: int,
    record: SourceRecord,
    engine: FontEngine,
    text_width: float,
    wrap_cache: dict[tuple[int, float], list[BodyLine]] | None,
) -> list[BodyLine]:
    cache_key = _wrap_cache_key(rec_i, engine)
    if wrap_cache is not None and cache_key in wrap_cache:
        return wrap_cache[cache_key]
    body_lines = wrap_record_text(record.text, engine, text_width)
    body_lines = _fill_short_record(
        body_lines, engine, text_width, max(len(record.date_lines), 1)
    )
    if wrap_cache is not None:
        wrap_cache[cache_key] = body_lines
    return body_lines


def layout(
    records: Sequence[SourceRecord],
    geo: TableGeometry,
    engine: FontEngine,
    cfg: Config,
    *,
    _wrap_cache: dict[tuple[int, float], list[BodyLine]] | None = None,
) -> list[Page]:
    """Розкладає записи по рядках бланка, повертає список аркушів."""
    text_left, text_right = geo.text_bounds(2)
    text_width = text_right - text_left

    pages: list[Page] = [Page()]
    row = 0
    rows_per_page = geo.data_rows
    number = cfg.start_number

    def new_page():
        nonlocal row
        pages.append(Page())
        row = 0

    for rec_i, record in enumerate(records):
        body_lines = _wrapped_body(rec_i, record, engine, text_width, _wrap_cache)
        needed = max(len(body_lines), len(record.date_lines), 1)

        if row >= rows_per_page:
            new_page()
        if not cfg.split_records and row + needed > rows_per_page:
            if needed <= rows_per_page:
                new_page()

        page = pages[-1]
        label = record.number.strip() or str(number)
        number += 1

        for i in range(needed):
            body = body_lines[i] if i < len(body_lines) else None
            is_blank = body is not None and body.is_blank

            if row >= rows_per_page:
                if is_blank:
                    continue
                new_page()
                page = pages[-1]
            if i == 0 and cfg.number_column:
                page.lines.append(PlacedLine(row, 0, label, "center"))
            if i < len(record.date_lines):
                page.lines.append(PlacedLine(row, 1, record.date_lines[i], "center"))
            if body is not None:
                last = i == len(body_lines) - 1
                for text, align in body.fragments:
                    if not text:
                        continue
                    if not align:
                        align = (
                            "left"
                            if (body.ends_paragraph or not cfg.justify)
                            else "justify"
                        )
                    page.lines.append(
                        PlacedLine(row, 2, text, align, is_last=last)
                    )
            row += 1

    return [p for p in pages if p.lines]


def _count_pages(
    records: Sequence[SourceRecord],
    geo: TableGeometry,
    engine: FontEngine,
    cfg: Config,
    wrap_cache: dict[tuple[int, float], list[BodyLine]] | None = None,
) -> int:
    """Як layout, але лише кількість аркушів — для підбору кегля."""
    text_left, text_right = geo.text_bounds(2)
    text_width = text_right - text_left
    row = 0
    pages = 1
    rows_per_page = geo.data_rows
    filled = False

    for rec_i, record in enumerate(records):
        body_lines = _wrapped_body(rec_i, record, engine, text_width, wrap_cache)
        needed = max(len(body_lines), len(record.date_lines), 1)

        if row >= rows_per_page:
            pages += 1
            row = 0
        if not cfg.split_records and row + needed > rows_per_page:
            if needed <= rows_per_page:
                pages += 1
                row = 0

        for i in range(needed):
            body = body_lines[i] if i < len(body_lines) else None
            is_blank = body is not None and body.is_blank
            if row >= rows_per_page:
                if is_blank:
                    continue
                pages += 1
                row = 0
            filled = True
            row += 1

    return pages if filled else 0


@dataclass
class FontFitInfo:
    """Що змінили, щоб увійти в --max-pages."""

    requested_pt: float
    font_size_pt: float
    pages: int
    uncompressed_pages: int


def fit_font_to_max_pages(
    records: Sequence[SourceRecord],
    geo: TableGeometry,
    engine: FontEngine,
    cfg: Config,
) -> tuple[FontEngine, FontFitInfo, dict[tuple[int, float], list[BodyLine]]]:
    """Найбільший кегль ≤ 12 pt, при якому аркушів не більше cfg.max_pages."""
    requested_pt = min(engine.size_pt, MAX_FONT_SIZE_PT)
    engine = engine.derive(requested_pt)
    wrap_cache: dict[tuple[int, float], list[BodyLine]] = {}
    start_pages = _count_pages(records, geo, engine, cfg, wrap_cache)
    max_pages = cfg.max_pages

    info = FontFitInfo(
        requested_pt=requested_pt,
        font_size_pt=requested_pt,
        pages=start_pages,
        uncompressed_pages=start_pages,
    )
    if max_pages is None or max_pages <= 0 or start_pages <= max_pages:
        return engine, info, wrap_cache

    lo = MIN_FONT_SIZE_PT
    hi = requested_pt
    lo_pages = _count_pages(records, geo, engine.derive(lo), cfg, wrap_cache)
    if lo_pages > max_pages:
        info.font_size_pt = lo
        info.pages = lo_pages
        return engine.derive(lo), info, wrap_cache

    best = lo
    best_pages = lo_pages
    while hi - lo > FONT_SIZE_STEP_PT:
        mid = _snap_font_pt((lo + hi) / 2.0)
        mid = min(max(mid, lo), hi)
        if mid == lo or mid == hi:
            break
        mid_pages = _count_pages(records, geo, engine.derive(mid), cfg, wrap_cache)
        if mid_pages <= max_pages:
            best, best_pages = mid, mid_pages
            lo = mid
        else:
            hi = mid

    best = _snap_font_pt(best)
    best = max(MIN_FONT_SIZE_PT, min(best, requested_pt))
    fitted = engine.derive(best)
    best_pages = _count_pages(records, geo, fitted, cfg, wrap_cache)
    while best_pages > max_pages and best > MIN_FONT_SIZE_PT + 1e-9:
        best = _snap_font_pt(best - FONT_SIZE_STEP_PT)
        best = max(best, MIN_FONT_SIZE_PT)
        fitted = engine.derive(best)
        best_pages = _count_pages(records, geo, fitted, cfg, wrap_cache)
    while True:
        nxt = _snap_font_pt(best + FONT_SIZE_STEP_PT)
        if nxt > requested_pt + 1e-9:
            break
        nxt_pages = _count_pages(records, geo, engine.derive(nxt), cfg, wrap_cache)
        if nxt_pages > max_pages:
            break
        best, best_pages = nxt, nxt_pages
        fitted = engine.derive(best)

    info.font_size_pt = best
    info.pages = best_pages
    return fitted, info, wrap_cache


# --------------------------------------------------------------------------------------
# Генерація полілiній сторінки
# --------------------------------------------------------------------------------------


def page_line_groups(
    page: Page, geo: TableGeometry, engine: FontEngine, cfg: Config
) -> list[list[Polyline]]:
    """Полілінії аркуша, згруповані за рядками бланка — у порядку письма."""
    out: list[list[Polyline]] = []
    for line in page.lines:
        left, right = geo.text_bounds(line.column)
        width = right - left
        extra_x = 0.0
        if line.column == 0:
            extra_x = cfg.number_offset_x
        elif line.column == 1:
            extra_x = cfg.date_offset_x
        offset = cfg.offset_x + extra_x
        baseline = geo.row_tops[line.row] + geo.baseline_in_row + cfg.offset_y

        text = line.text
        extra = 0.0
        x = left + offset

        if line.align == "center":
            x += (width - engine.text_width(text)) / 2.0
        elif line.align == "right":
            x += width - engine.text_width(text)
        elif line.align == "justify":
            spaces = text.count(" ")
            if spaces:
                slack = width - engine.text_width(text)
                if slack > 0:
                    extra = slack / spaces

        out.append(engine.render(text, x, baseline, extra))
    return out


def page_polylines(page: Page, geo: TableGeometry, engine: FontEngine, cfg: Config) -> list[Polyline]:
    return [p for group in page_line_groups(page, geo, engine, cfg) for p in group]


def grid_polylines(geo: TableGeometry, cfg: Config, header: bool = True) -> list[Polyline]:
    """Лінії бланка (для перевірки суміщення або друку на чистому аркуші)."""
    out: list[Polyline] = []
    # Сітка має рухатися разом із текстом, інакше зсув по Y роз'їжджає
    # написане з лініями бланка на прев'ю та в режимі --grid full.
    dy = cfg.offset_y
    top = (geo.row_tops[0] if geo.row_tops else geo.margin_top) + dy
    bottom = ((geo.row_tops[-1] + geo.row_height) if geo.row_tops else geo.margin_top) + dy
    if header:
        top = geo.margin_top + dy
    x0, x1 = geo.col_x[0] + cfg.offset_x, geo.col_x[-1] + cfg.offset_x
    for x in geo.col_x:
        out.append([(x + cfg.offset_x, top), (x + cfg.offset_x, bottom)])
    ys = [t + dy for t in geo.row_tops] + [bottom]
    if header:
        ys = [geo.margin_top + dy] + ys
    for y in ys:
        out.append([(x0, y), (x1, y)])
    return out


def optimize_order(polys: list[Polyline]) -> list[Polyline]:
    """Найближчий сусід: зменшує холості переміщення."""
    if len(polys) < 3:
        return polys
    remaining = list(polys)
    result = [remaining.pop(0)]
    cur = result[-1][-1]
    while remaining:
        best_i, best_d, flip = 0, float("inf"), False
        for i, p in enumerate(remaining):
            d0 = dist(cur, p[0])
            if d0 < best_d:
                best_i, best_d, flip = i, d0, False
            closed = dist(p[0], p[-1]) < 1e-9
            if not closed:
                d1 = dist(cur, p[-1])
                if d1 < best_d:
                    best_i, best_d, flip = i, d1, True
        p = remaining.pop(best_i)
        if flip:
            p = p[::-1]
        result.append(p)
        cur = p[-1]
    return result


# --------------------------------------------------------------------------------------
# Вивід: G-code та SVG
# --------------------------------------------------------------------------------------


def y_anchor(geo: TableGeometry, cfg: Config) -> float:
    """Базова лінія першого рядка запису — нерухома точка вертикальної калібровки.

    Прив'язка саме до неї, бо на плотері перший рядок збігається з бланком, а
    розбіжність накопичується з віддаленням від нього.
    """
    top = geo.row_tops[0] if geo.row_tops else geo.margin_top
    return top + geo.baseline_in_row + cfg.offset_y


def page_configs(pages_count: int, cfg: Config) -> list[Config]:
    """Налаштування для кожного аркуша: парні зсунуті по Y, непарні — як є.

    Зсув вносимо саме через offset_y, бо його враховують і текст, і сітка
    бланка, і прив'язка вертикальної калібровки — аркуш зсувається цілком.
    """
    return [
        cfg if i % 2 == 0 else replace(cfg, offset_y=cfg.offset_y + cfg.even_offset_y)
        for i in range(pages_count)
    ]


def to_machine(pt: Point, geo: TableGeometry, cfg: Config) -> Point:
    x, y = pt
    if cfg.scale_y != 1.0:
        anchor = y_anchor(geo, cfg)
        y = anchor + (y - anchor) * cfg.scale_y
    if cfg.mirror_x:
        x = cfg.page_w - x
    if cfg.origin == "bottom-left":
        y = cfg.page_h - y
    return x, y


_TRANSLIT = {
    "а": "a", "б": "b", "в": "v", "г": "h", "ґ": "g", "д": "d", "е": "e", "є": "ie",
    "ж": "zh", "з": "z", "и": "y", "і": "i", "ї": "i", "й": "i", "к": "k", "л": "l",
    "м": "m", "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "kh", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "shch", "ь": "",
    "ю": "iu", "я": "ia", "ы": "y", "э": "e", "ъ": "", "ё": "e",
}


def _ascii(text: str) -> str:
    """ASCII-безпечний варіант рядка для коментарів у G-code."""
    out: list[str] = []
    for ch in text:
        if ord(ch) < 128:
            out.append(ch)
            continue
        low = ch.lower()
        rep = _TRANSLIT.get(low)
        if rep is None:
            rep = unicodedata.normalize("NFKD", ch).encode("ascii", "ignore").decode()
        if ch.isupper() and rep:
            rep = rep.capitalize()
        out.append(rep or "?")
    return "".join(out)


def _fmt(v: float) -> str:
    return f"{v:.3f}".rstrip("0").rstrip(".") or "0"


def emit_gcode(
    polys: list[Polyline],
    geo: TableGeometry,
    cfg: Config,
    title: str,
) -> tuple[str, dict]:
    lines: list[str] = []
    a = lines.append

    # Коментарі — лише ASCII і короткими рядками: буфер рядка GRBL 80 байтів,
    # а частина відправників не приймає не-ASCII байти навіть після ";".
    if cfg.comments:
        a(f"; {title}")
        a(f"; font: {_ascii(cfg.font_query)} {_fmt(cfg.font_size_pt or geo.font_size_pt)}pt"
          f" mode={cfg.mode}")
        a(f"; page: {_fmt(cfg.page_w)}x{_fmt(cfg.page_h)}mm origin={cfg.origin}")
        a(f"; offset X{_fmt(cfg.offset_x)} Y{_fmt(cfg.offset_y)}"
          f" number X{_fmt(cfg.number_offset_x)} date X{_fmt(cfg.date_offset_x)}")
        if cfg.scale_y != 1.0:
            a(f"; Y calibration: scale={cfg.scale_y:.4f}"
              f" anchor={_fmt(y_anchor(geo, cfg))}mm from page top")
        a(f"; smooth={_fmt(cfg.smooth)} step={_fmt(cfg.step)}mm feed={_fmt(cfg.feed)}")
        a(f"; Z safe={_fmt(cfg.z_safe)} draw={_fmt(cfg.z_draw)}")
    a("G21")
    a("G90")
    a("G17")
    a(f"G0 Z{_fmt(cfg.z_safe)}")

    draw_len = 0.0
    travel_len = 0.0
    z_down_len = 0.0
    z_up_len = 0.0
    z_stroke = abs(cfg.z_safe - cfg.z_draw)
    last: Point | None = None
    min_x = min_y = float("inf")
    max_x = max_y = float("-inf")

    for poly in polys:
        machine = [to_machine(p, geo, cfg) for p in poly]
        if len(machine) < 2:
            continue
        for x, y in machine:
            min_x, max_x = min(min_x, x), max(max_x, x)
            min_y, max_y = min(min_y, y), max(max_y, y)
        start = machine[0]
        if last is not None:
            travel_len += dist(last, start)
        a(f"G0 X{_fmt(start[0])} Y{_fmt(start[1])}")
        a(f"G1 Z{_fmt(cfg.z_draw)} F{_fmt(cfg.feed_z)}")
        z_down_len += z_stroke
        a(f"G1 F{_fmt(cfg.feed)}")
        prev = start
        for x, y in machine[1:]:
            a(f"G1 X{_fmt(x)} Y{_fmt(y)}")
            draw_len += dist(prev, (x, y))
            prev = (x, y)
        a(f"G0 Z{_fmt(cfg.z_safe)}")
        z_up_len += z_stroke
        last = prev

    a(f"G0 Z{_fmt(cfg.z_safe)}")
    a("G0 X0 Y0")
    a("M2")

    # Перо піднімається й опускається на кожен контур, а контурів тисячі:
    # при глибокому натиску хід по Z займає стільки ж часу, як саме письмо,
    # тому враховуємо його в оцінці, а не лише переміщення в площині.
    minutes = 0.0
    if cfg.feed:
        minutes = draw_len / cfg.feed + travel_len / cfg.feed_travel
        if cfg.feed_z:
            minutes += z_down_len / cfg.feed_z
        if cfg.feed_travel:
            minutes += z_up_len / cfg.feed_travel

    stats = {
        "contours": len(polys),
        "points": sum(len(p) for p in polys),
        "draw_mm": draw_len,
        "travel_mm": travel_len,
        "z_mm": z_down_len + z_up_len,
        "minutes": minutes,
        "bbox": (min_x, min_y, max_x, max_y) if polys else (0, 0, 0, 0),
    }
    return "\n".join(lines) + "\n", stats


def emit_svg(pages: list[list[Polyline]], geo: TableGeometry, cfgs: Sequence[Config]) -> str:
    """Прев'ю 1:1 у мм — можна роздрукувати й накласти на бланк.

    Налаштування свої для кожного аркуша, бо парні зсунуті по Y: сітка бланка
    має поїхати разом із текстом, інакше прев'ю показало б розбіжність, якої
    на папері немає — там бланк парного аркуша й надрукований вище.
    """
    cfg = cfgs[0]
    parts: list[str] = []
    total_h = cfg.page_h * len(pages)
    parts.append(
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{cfg.page_w}mm" '
        f'height="{total_h}mm" viewBox="0 0 {cfg.page_w} {total_h}">'
    )
    parts.append('<rect width="100%" height="100%" fill="white"/>')
    for i, polys in enumerate(pages):
        dy = i * cfg.page_h
        parts.append(f'<g transform="translate(0,{dy})">')
        parts.append(
            f'<rect x="0" y="0" width="{cfg.page_w}" height="{cfg.page_h}" '
            'fill="none" stroke="#cccccc" stroke-width="0.2"/>'
        )
        for poly in grid_polylines(geo, cfgs[i]):
            d = "M " + " L ".join(f"{x:.3f},{y:.3f}" for x, y in poly)
            parts.append(f'<path d="{d}" fill="none" stroke="#b0c4de" stroke-width="0.12"/>')
        for poly in polys:
            d = "M " + " L ".join(f"{x:.3f},{y:.3f}" for x, y in poly)
            parts.append(
                f'<path d="{d}" fill="none" stroke="#111111" stroke-width="0.22" '
                'stroke-linecap="round" stroke-linejoin="round"/>'
            )
        parts.append("</g>")
    parts.append("</svg>")
    return "\n".join(parts)


# --------------------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------------------


def substitute_text(text: str) -> str:
    return "".join(TYPOGRAPHIC_SUBSTITUTES.get(ch, ch) for ch in text)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="docx2gcode",
        description="Конвертує текст із заповненої DOCX-таблиці у G-code для письмового плотера.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    # Дефолти беруться з Config, щоб не задавати ті самі значення двічі:
    # правка поля в Config одразу стає новим типовим значенням для CLI.
    d = Config()

    p.add_argument("source", nargs="?", type=Path, help="заповнена таблиця (.docx)")
    p.add_argument("-t", "--template", type=Path,
                   help=f"порожній бланк (.docx); типово «{DEFAULT_TEMPLATE}» поруч із джерелом")
    p.add_argument("-o", "--output", type=Path,
                   help=f"файл або тека для G-code; типово «{DEFAULT_OUTPUT_DIR}» поруч із текою проєкту")

    g = p.add_argument_group("плотер")
    g.add_argument("--feed", type=float, default=d.feed, help="робоча подача, мм/хв")
    g.add_argument("--feed-z", type=float, default=d.feed_z, help="подача по Z, мм/хв")
    g.add_argument("--feed-travel", type=float, default=d.feed_travel, help="швидкість холостих ходів (для оцінки часу)")
    g.add_argument("--z-safe", type=float, default=d.z_safe, help="висота підйому пера, мм")
    g.add_argument("--z-draw", type=float, default=d.z_draw, help="сила натиску (Z під час письма), мм")

    g = p.add_argument_group("шрифт")
    g.add_argument("--font", default=d.font_query, help="назва або шлях до шрифта")
    g.add_argument("--font-index", type=int, default=d.font_index, help="індекс шрифта у .ttc")
    g.add_argument("--fallback-font", default=d.fallback_font_query,
                   help="запасний шрифт для відсутніх символів")
    g.add_argument("--font-size", type=float, default=d.font_size_pt,
                   help="кегль у пунктах (типово — як у джерелі)")
    g.add_argument("--mode", choices=["centerline", "outline"], default=d.mode,
                   help="centerline — одноштрихове письмо; outline — обвід контуру літери")
    g.add_argument("--raster-ppmm", type=float, default=d.raster_ppmm,
                   help="роздільність растра для режиму centerline, px/мм")
    g.add_argument("--prune", type=float, default=d.prune,
                   help="прибирати відгалуження, коротші за це (centerline), мм")
    g.add_argument("--dot-size", type=float, default=d.dot_size,
                   help="елементи, менші за це, лишати контуром — крапки (centerline), мм")
    g.add_argument("--smooth", type=float, default=d.smooth, help="згладжування сплайном, 0..1")
    g.add_argument("--step", type=float, default=d.step, help="крок дискретизації кривих, мм")
    g.add_argument("--corner-angle", type=float, default=d.corner_angle,
                   help="кути гостріші за це не згладжуються, ° (типово 50 для outline, 140 для centerline)")
    g.add_argument("--simplify", type=float, default=d.simplify, help="допуск спрощення контуру, мм")
    g.add_argument("--slant", type=float, default=d.slant, help="додатковий нахил (штучний курсив), °")
    g.add_argument("--letter-spacing", type=float, default=d.letter_spacing, help="трекінг, мм")
    g.add_argument("--word-spacing", type=float, default=d.word_spacing, help="додаткова ширина пробілу, мм")
    g.add_argument("--no-kerning", action="store_true", help="вимкнути кернінг")
    g.add_argument("--no-substitute", action="store_true", help="не замінювати типографські символи")

    g = p.add_argument_group("сторінка")
    g.add_argument("--page-width", type=float, default=d.page_w, help="ширина сторінки, мм")
    g.add_argument("--page-height", type=float, default=d.page_h, help="висота сторінки, мм")
    g.add_argument("--offset-x", type=float, default=d.offset_x, help="загальний зсув по X, мм")
    g.add_argument("--offset-y", type=float, default=d.offset_y,
                   help="загальний зсув тексту і сітки по Y, мм (мінус = вгору)")
    g.add_argument("--even-offset-y", type=float, default=d.even_offset_y,
                   help="додаток до зсуву по Y для парних аркушів, мм (мінус = вгору, 0 — як непарні)")
    g.add_argument("--date-offset-x", type=float, default=d.date_offset_x, help="додатковий зсув колонки дати, мм")
    g.add_argument("--number-offset-x", type=float, default=d.number_offset_x,
                   help="додатковий зсув колонки номерів, мм (мінус = ліворуч)")
    g.add_argument("--baseline-offset", type=float, default=d.baseline_offset, help="підстроювання базової лінії, мм")
    g.add_argument("--origin", choices=["bottom-left", "top-left"], default=d.origin)
    g.add_argument("--mirror-x", action="store_true", help="дзеркалити по X")
    g.add_argument("--scale-y", type=float, default=d.scale_y,
                   help="компенсація недоходу плотера по Y; 1 — без компенсації")

    g = p.add_argument_group("розкладка")
    g.add_argument("--header-rows", type=int, default=d.header_rows,
                   help="скільки перших рядків пропустити перед письмом (-1 — визначити автоматично)")
    g.add_argument("--start-number", type=int, default=d.start_number, help="номер першого запису")
    g.add_argument("--no-numbers", action="store_true", help="не писати номери в колонці 1")
    g.add_argument("--no-justify", action="store_true", help="не вирівнювати текст по ширині")
    g.add_argument("--no-split", action="store_true", help="не розривати запис між аркушами")
    g.add_argument("--max-rows", type=int, default=d.max_rows, help="скільки рядків використовувати на аркуші")
    g.add_argument("--max-pages", type=int, default=d.max_pages or 0,
                   help="максимум файлів .gcode; кегль підбирається автоматично (не більше 12 pt, 0 — без ліміту)")
    g.add_argument("--grid", choices=["none", "calib", "full"], default=d.grid,
                   help="none — писати на надрукований бланк; calib — тільки сітка; full — сітка + текст")
    g.add_argument("--optimize", action="store_true",
                   help="оптимізувати порядок контурів усередині рядка (типово увімкнено)")
    g.add_argument("--no-optimize", action="store_true",
                   help="не оптимізувати порядок контурів — довші холості ходи, швидша генерація")

    g = p.add_argument_group("вивід та діагностика")
    g.add_argument("--no-comments", action="store_true",
                   help="не писати заголовкові коментарі у G-code")
    g.add_argument("--single-file", action="store_true", help="усі аркуші в один файл із паузою M0")
    g.add_argument("--preview", type=Path, default=d.preview,
                   help="шлях для SVG-прев'ю (1:1); типово — поруч із G-code")
    g.add_argument("--no-preview", action="store_true", help="не створювати SVG-прев'ю")
    g.add_argument("--check-font", action="store_true", help="перевірити покриття символів і вийти")
    g.add_argument("--list-fonts", metavar="ЗАПИТ", nargs="?", const="", help="показати знайдені шрифти")
    g.add_argument("--rebuild-font-cache", action="store_true", help="перебудувати кеш шрифтів")
    g.add_argument("--dry-run", action="store_true",
                   help="не записувати G-code (прев'ю, якщо задане, усе одно створюється)")
    g.add_argument("-v", "--verbose", action="store_true")
    return p


def config_from_args(args) -> Config:
    # Прапорці-перемикачі лише зсувають типове значення з Config: --no-* здатні
    # його вимкнути, а --optimize і подібні — увімкнути. Тому правка булевого
    # поля в Config теж працює як новий дефолт.
    d = Config()
    return Config(
        feed=args.feed,
        feed_travel=args.feed_travel,
        feed_z=args.feed_z,
        z_safe=args.z_safe,
        z_draw=args.z_draw,
        font_query=args.font,
        font_index=args.font_index,
        fallback_font_query=args.fallback_font,
        font_size_pt=args.font_size,
        smooth=args.smooth,
        step=args.step,
        corner_angle=args.corner_angle,
        simplify=args.simplify,
        mode=args.mode,
        raster_ppmm=args.raster_ppmm,
        prune=args.prune,
        dot_size=args.dot_size,
        slant=args.slant,
        letter_spacing=args.letter_spacing,
        word_spacing=args.word_spacing,
        kerning=d.kerning and not args.no_kerning,
        substitute=d.substitute and not args.no_substitute,
        page_w=args.page_width,
        page_h=args.page_height,
        offset_x=args.offset_x,
        offset_y=args.offset_y,
        even_offset_y=args.even_offset_y,
        date_offset_x=args.date_offset_x,
        number_offset_x=args.number_offset_x,
        baseline_offset=args.baseline_offset,
        origin=args.origin,
        mirror_x=d.mirror_x or args.mirror_x,
        scale_y=args.scale_y,
        header_rows=None if args.header_rows is not None and args.header_rows < 0
        else args.header_rows,
        start_number=args.start_number,
        number_column=d.number_column and not args.no_numbers,
        justify=d.justify and not args.no_justify,
        split_records=d.split_records and not args.no_split,
        max_rows=args.max_rows,
        max_pages=args.max_pages if args.max_pages and args.max_pages > 0 else None,
        grid=args.grid,
        optimize=(d.optimize or args.optimize) and not args.no_optimize,
        comments=d.comments and not args.no_comments,
        single_file=d.single_file or args.single_file,
        preview=args.preview,
        write_preview=d.write_preview and not args.no_preview,
        verbose=d.verbose or args.verbose,
    )


def resolve_template(args) -> Path:
    """Бланк із -t або типовий «1 лист.docx» поруч із джерелом чи скриптом."""
    if args.template:
        if not args.template.exists():
            raise SystemExit(f"Шаблон не знайдено: {args.template}")
        return args.template

    for base in (args.source.parent, Path(__file__).resolve().parent, Path.cwd()):
        candidate = base / DEFAULT_TEMPLATE
        if candidate.exists():
            return candidate

    raise SystemExit(
        f"Не знайдено типовий бланк «{DEFAULT_TEMPLATE}».\n"
        f"Покладіть його поруч із {args.source.name} або вкажіть шлях:\n"
        "  -t /шлях/до/бланка.docx"
    )


def cmd_list_fonts(query: str, rebuild: bool) -> int:
    fonts = build_font_index(rebuild)
    target = _normalise(query) if query else ""
    rows: list[tuple[str, str, int]] = []
    for entry in fonts:
        names = entry["names"]
        if target and not any(target in _normalise(n) for n in names):
            continue
        rows.append((names[0], entry["path"], entry["index"]))
    rows.sort()
    if not rows:
        print(f"Шрифтів за запитом «{query}» не знайдено." if query else "Шрифтів не знайдено.")
        return 1
    for name, path, index in rows:
        suffix = f"  [#{index}]" if index else ""
        print(f"{name:<42} {path}{suffix}")
    print(f"\nВсього: {len(rows)}")
    return 0


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
            print(f"Попередження: запасний шрифт «{cfg.fallback_font_query}» не знайдено.", file=sys.stderr)
        else:
            engine.fallback = FontEngine(fb[0], fb[1], size_pt, cfg)
    return engine


def check_coverage(engine: FontEngine, records: Sequence[SourceRecord], cfg: Config) -> int:
    chars: set[str] = set()
    for r in records:
        for s in [r.text, r.note, r.number, *r.date_lines]:
            chars.update(s)
    chars.discard("\t")
    text_chars = {c for c in chars if not c.isspace()}

    missing = sorted(c for c in text_chars if not engine.has(c))
    # Символ може бути в cmap, але вести на порожній гліф — на папері
    # це виглядає як пропуск, тому перевіряємо ще й наявність контурів.
    empty = sorted(
        c for c in text_chars
        if engine.has(c) and not engine.glyph_contours(engine.glyph_name(c))
    )
    covered = len(text_chars) - len(missing) - len(empty)
    print(f"Шрифт: {engine.path.name}")
    print(f"Символів у документі: {len(text_chars)}   покрито: {covered}   "
          f"відсутні: {len(missing)}   порожні: {len(empty)}")

    def report(chars: list[str], label: str) -> None:
        if not chars:
            return
        print(f"\n{label}:")
        for ch in chars:
            name = unicodedata.name(ch, "?")
            repl = TYPOGRAPHIC_SUBSTITUTES.get(ch)
            hint = f"  → буде замінено на «{repl}»" if repl and cfg.substitute else ""
            print(f"  U+{ord(ch):04X}  «{ch}»  {name}{hint}")

    report(missing, "Немає у шрифті")
    report(empty, "Є в шрифті, але гліф порожній")

    if missing or empty:
        print(
            "\nПідказка: якщо бракує кирилиці, візьміть шрифт із кириличним\n"
            "покриттям або допишіть відсутні символи:  --fallback-font \"Segoe Script\""
        )
        return 1
    print("\nУсі символи документа підтримуються шрифтом.")
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    if os.name == "nt":  # консоль Windows типово не в UTF-8
        for stream in (sys.stdout, sys.stderr):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
    args = build_parser().parse_args(argv)

    if args.list_fonts is not None:
        return cmd_list_fonts(args.list_fonts, args.rebuild_font_cache)

    if args.source is None:
        build_parser().print_help()
        return 2
    if not args.source.exists():
        raise SystemExit(f"Файл не знайдено: {args.source}")

    template = resolve_template(args)

    cfg = config_from_args(args)

    records = read_source(args.source)
    if not records:
        raise SystemExit(f"У {args.source} не знайдено записів.")
    if cfg.substitute:
        for r in records:
            r.text = substitute_text(r.text)
            r.note = substitute_text(r.note)
            r.number = substitute_text(r.number)
            r.date_lines = [substitute_text(d) for d in r.date_lines]

    geo = read_template(template, cfg)
    size_pt = min(cfg.font_size_pt or geo.font_size_pt, MAX_FONT_SIZE_PT)
    cfg = replace(cfg, font_size_pt=size_pt)

    engine = resolve_font(cfg, size_pt, args.rebuild_font_cache)

    if args.check_font:
        return check_coverage(engine, records, cfg)

    if cfg.verbose:
        print(f"Джерело:  {args.source.name}  — {len(records)} записів")
        print(f"Бланк:    {template.name}  — {geo.data_rows} рядків, "
              f"висота рядка {geo.row_height:.2f} мм")
        print(f"Колонки:  " + ", ".join(f"{a:.1f}–{b:.1f}" for a, b in
                                        zip(geo.col_x, geo.col_x[1:])) + " мм")
        print(f"Шрифт:    {engine.path.name}  {size_pt:g} pt")

    engine, fit_info, wrap_cache = fit_font_to_max_pages(records, geo, engine, cfg)
    cfg = replace(cfg, font_size_pt=engine.size_pt)
    if cfg.verbose and fit_info.uncompressed_pages != fit_info.pages:
        print(
            f"Кегль:     {fit_info.requested_pt:g} → {fit_info.font_size_pt:g} pt  "
            f"({fit_info.uncompressed_pages} → {fit_info.pages} аркушів, ліміт {cfg.max_pages})"
        )
    elif cfg.verbose:
        print(f"Кегль:     {fit_info.font_size_pt:g} pt  ({fit_info.pages} аркушів)")

    pages = layout(records, geo, engine, cfg, _wrap_cache=wrap_cache)
    if not pages:
        raise SystemExit("Нічого розкладати.")

    cfgs = page_configs(len(pages), cfg)

    rendered: list[list[Polyline]] = []
    for page, page_cfg in zip(pages, cfgs):
        groups = page_line_groups(page, geo, engine, page_cfg)
        # Оптимізація тільки всередині рядка: наскрізне перевпорядкування
        # шукає найближчий контур по всьому аркушу і перо починає блукати
        # між колонками замість того, щоб дописати рядок до кінця.
        if page_cfg.optimize:
            groups = [optimize_order(g) for g in groups]
        polys = [p for g in groups for p in g]
        if page_cfg.grid in ("calib", "full"):
            grid = grid_polylines(geo, page_cfg)
            polys = grid if page_cfg.grid == "calib" else grid + polys
        rendered.append(polys)

    if engine.missing:
        chars = " ".join(f"«{c}»" for c in sorted(engine.missing))
        print(
            f"Попередження: {len(engine.missing)} символів відсутні у шрифті і пропущені: {chars}\n"
            f"Перевірте покриття:  --check-font",
            file=sys.stderr,
        )

    # --- вивід ---
    if args.output is None:
        # Типово — тека gcodeOutput поруч із текою проєкту, щоб згенеровані
        # файли не змішувалися з джерелами. Створюється за потреби.
        out = args.source.resolve().parent.parent / DEFAULT_OUTPUT_DIR
        is_file = False
        outdir, stem = out, args.source.stem
    else:
        # -o може бути файлом (.gcode/.nc/.ngc), наявною текою або префіксом імені.
        out = args.output
        is_file = out.suffix.lower() in (".gcode", ".nc", ".ngc")
        if is_file:
            outdir, stem = out.parent, out.stem
        elif out.is_dir():
            outdir, stem = out, args.source.stem
        else:
            outdir, stem = out.parent, out.name

    totals = dict(contours=0, points=0, draw_mm=0.0, travel_mm=0.0, z_mm=0.0, minutes=0.0)
    written: list[Path] = []

    if cfg.single_file or len(rendered) == 1:
        chunks: list[str] = []
        for i, polys in enumerate(rendered):
            text, stats = emit_gcode(polys, geo, cfgs[i], f"sheet {i + 1} of {len(rendered)}")
            for k in totals:
                totals[k] += stats[k]
            if i:
                chunks.append("M0 ; change sheet, then resume")
            chunks.append(text)
        target = out if is_file else outdir / f"{stem}.gcode"
        if not args.dry_run:
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text("\n".join(chunks), "utf-8")
        written.append(target)
    else:
        for i, polys in enumerate(rendered):
            text, stats = emit_gcode(polys, geo, cfgs[i], f"sheet {i + 1} of {len(rendered)}")
            for k in totals:
                totals[k] += stats[k]
            target = outdir / f"{stem}_p{i + 1:02d}.gcode"
            if not args.dry_run:
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_text(text, "utf-8")
            written.append(target)

    # Прев'ю пишеться і під --dry-run: це і є сценарій «подивитися,
    # не генеруючи G-code».
    preview_path = None
    if cfg.write_preview:
        preview_path = cfg.preview or outdir / f"{stem}_preview.svg"
        preview_path.parent.mkdir(parents=True, exist_ok=True)
        preview_path.write_text(emit_svg(rendered, geo, cfgs), "utf-8")

    # --- звіт ---
    print(f"Записів: {len(records)}   аркушів: {len(rendered)}")
    if fit_info.uncompressed_pages > fit_info.pages:
        print(
            f"Кегль підібрано під ліміт {cfg.max_pages} аркушів: "
            f"{fit_info.requested_pt:g} → {fit_info.font_size_pt:g} pt "
            f"({fit_info.uncompressed_pages} → {fit_info.pages})"
        )
    elif cfg.max_pages:
        print(f"Кегль: {fit_info.font_size_pt:g} pt  (ліміт {cfg.max_pages} аркушів)")
    if cfg.max_pages and len(rendered) > cfg.max_pages:
        print(
            f"Попередження: навіть при {fit_info.font_size_pt:g} pt вийшло "
            f"{len(rendered)} аркушів (ліміт {cfg.max_pages}).",
            file=sys.stderr,
        )
    print(f"Контурів: {totals['contours']}   точок: {totals['points']}")
    print(f"Довжина письма: {totals['draw_mm'] / 1000:.2f} м   "
          f"холостих ходів: {totals['travel_mm'] / 1000:.2f} м   "
          f"хід по Z: {totals['z_mm'] / 1000:.2f} м")
    print(f"Орієнтовний час: {totals['minutes']:.1f} хв")
    if cfg.scale_y != 1.0 and len(geo.row_tops) > 1:
        span = geo.row_tops[-1] - geo.row_tops[0]
        print(f"Калібровка по Y: ×{cfg.scale_y:.4f} — останній рядок опущено на "
              f"{span * (cfg.scale_y - 1):.2f} мм (прев'ю показує бланк без компенсації)")
    if cfg.even_offset_y and len(rendered) > 1:
        verb = "піднято" if cfg.even_offset_y < 0 else "опущено"
        print(f"Парні аркуші: {verb} на {abs(cfg.even_offset_y):.2f} мм "
              f"(offset_y {cfg.offset_y + cfg.even_offset_y:+.2f} проти {cfg.offset_y:+.2f})")
    if args.dry_run:
        print("(--dry-run: G-code не записано)")
    else:
        for w in written:
            print(f"→ {w}")
    if preview_path is not None:
        print(f"→ {preview_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
