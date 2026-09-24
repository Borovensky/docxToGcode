"""Розкладка журналу: перенос, підпис командира, підбір кегля під ліміт аркушів.

Інтерфейс: layout(), fit_font_to_max_pages(), Page, PlacedLine.
Правила «N 231» і порожнього рядка дати сховані за wrap_record_text.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Sequence

from .config import (
    FONT_SIZE_STEP_PT,
    MAX_FONT_SIZE_PT,
    MIN_FONT_SIZE_PT,
    Config,
    snap_font_pt,
)
from .docx import SourceRecord, TableGeometry
from .fonts import FontEngine

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
        mid = snap_font_pt((lo + hi) / 2.0)
        mid = min(max(mid, lo), hi)
        if mid == lo or mid == hi:
            break
        mid_pages = _count_pages(records, geo, engine.derive(mid), cfg, wrap_cache)
        if mid_pages <= max_pages:
            best, best_pages = mid, mid_pages
            lo = mid
        else:
            hi = mid

    best = snap_font_pt(best)
    best = max(MIN_FONT_SIZE_PT, min(best, requested_pt))
    fitted = engine.derive(best)
    best_pages = _count_pages(records, geo, fitted, cfg, wrap_cache)
    while best_pages > max_pages and best > MIN_FONT_SIZE_PT + 1e-9:
        best = snap_font_pt(best - FONT_SIZE_STEP_PT)
        best = max(best, MIN_FONT_SIZE_PT)
        fitted = engine.derive(best)
        best_pages = _count_pages(records, geo, fitted, cfg, wrap_cache)
    while True:
        nxt = snap_font_pt(best + FONT_SIZE_STEP_PT)
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
