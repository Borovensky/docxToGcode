"""Рендер рядків бланка, координати плотера, G-code і SVG.

Інтерфейс: page_line_groups, page_configs, emit_gcode, emit_svg, optimize_order.
"""

from __future__ import annotations

import unicodedata
from dataclasses import replace
from typing import Sequence

from .config import Config
from .docx import TableGeometry
from .fonts import FontEngine
from .geom import Point, Polyline, dist
from .layout import Page

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
