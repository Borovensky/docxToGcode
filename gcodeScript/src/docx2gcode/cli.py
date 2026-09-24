"""CLI і оркестрація пайплайна: аргументи → записи → розкладка → файли.

Інтерфейс пакета для людини і droplet — main(). Усі правила письма
залишаються в layout/fonts/emit; тут лише збір і звіт.
"""

from __future__ import annotations

import argparse
import os
import sys
import unicodedata
from dataclasses import replace
from pathlib import Path
from typing import Sequence

from .config import (
    DEFAULT_OUTPUT_DIR,
    DEFAULT_TEMPLATE,
    MAX_FONT_SIZE_PT,
    PROJECT_DIR,
    TYPOGRAPHIC_SUBSTITUTES,
    Config,
)
from .docx import SourceRecord, read_source, read_template
from .emit import (
    emit_gcode,
    emit_svg,
    grid_polylines,
    optimize_order,
    page_configs,
    page_line_groups,
)
from .fonts import FontEngine, _normalise, build_font_index, resolve_font
from .geom import Polyline
from .layout import fit_font_to_max_pages, layout


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

    for base in (args.source.parent, PROJECT_DIR, Path.cwd()):
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
