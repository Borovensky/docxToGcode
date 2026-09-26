"""Розкладка журналу: перенос, підпис командира, підбір кегля під ліміт аркушів.

Інтерфейс: layout(), fit_font_to_max_pages(), Page, PlacedLine.
Правила «N 231» і порожнього рядка дати сховані за wrap_record_text.
"""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass, field
from typing import Sequence

from .config import (
    FONT_SIZE_STEP_PT,
    MAX_FONT_SIZE_PT,
    MIN_FONT_SIZE_PT,
    PROJECT_DIR,
    Config,
    snap_font_pt,
)
from .docx import SourceRecord, TableGeometry
from .fonts import FontEngine

_LOCAL_CONFIG = PROJECT_DIR / "local_config.toml"


def _line_break_markers() -> tuple[str, ...]:
    """Позначки з local_config.toml поруч із бланком. Немає файлу — немає розривів.

    Пошкоджений файл або поле не того типу зупиняє програму.
    """
    if not _LOCAL_CONFIG.is_file():
        return ()
    if sys.version_info < (3, 11):
        raise SystemExit(f"Щоб прочитати {_LOCAL_CONFIG.name}, потрібен Python 3.11+.")
    import tomllib

    try:
        with _LOCAL_CONFIG.open("rb") as fh:
            data = tomllib.load(fh)
    except tomllib.TOMLDecodeError as exc:
        raise SystemExit(f"Не вдалося прочитати {_LOCAL_CONFIG}: {exc}") from exc
    if "markers" not in data:
        raise SystemExit(f"{_LOCAL_CONFIG}: немає поля «markers».")
    markers = data["markers"]
    if not isinstance(markers, list) or not all(isinstance(item, str) for item in markers):
        raise SystemExit(f"{_LOCAL_CONFIG}: «markers» має бути списком рядків.")
    return tuple(item for item in markers if item)


def _compile_marker_re() -> re.Pattern[str]:
    """Розрив перед дужкою, тож дужка лишається на початку нового рядка."""
    markers = _line_break_markers()
    if not markers:
        return re.compile(r"(?!)")
    body = "|".join(re.escape(item) for item in markers)
    return re.compile(rf"(?=\(\s*(?:{body})\b)", re.IGNORECASE)


_MARKER_RE = _compile_marker_re()

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


#: «N» і наступне слово — один шматок, хоч би що стояло після нього:
#: «N 231», «N БР-313/1/ТГр/362», «N Купянськ». Голе «N» не лишається
#: в кінці рядка.
_DOC_MARK = r"(?:N|№|No\.?)"
_DOC_NUM_RE = re.compile(
    rf"(?<!\w){_DOC_MARK}\s+\S+",
    re.IGNORECASE,
)
_DOC_MARK_RE = re.compile(rf"^{_DOC_MARK}$", re.IGNORECASE)


def wrap_tokens(text: str) -> list[str]:
    """Слова для переносу: «N» і слово одразу після нього — один шматок."""
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
        if _DOC_MARK_RE.fullmatch(last) and rest:
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
            if _DOC_MARK_RE.fullmatch(piece.rstrip().rsplit(" ", 1)[-1]) and rest:
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
    is_signature: bool = False

    @property
    def is_blank(self) -> bool:
        return not any(text for text, _align in self.fragments)


def is_commander_signature(text: str) -> bool:
    """Запис, у якому немає тексту крім підпису командира.

    Такому рядку не ставимо номер у крайній лівій колонці.
    """
    parts = list(iter_signature_parts(text))
    return bool(parts) and all(kind == "commander" for kind, _payload in parts)


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
            out.append(BodyLine([(title, "left")], ends_paragraph=True, is_signature=True))
            out.append(
                BodyLine([(rank, "left"), (name, "right")], ends_paragraph=True, is_signature=True)
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
    is_signature: bool = False


@dataclass
class Page:
    lines: list[PlacedLine] = field(default_factory=list)
    #: Якщо підпис не вміщався внизу аркуша, цей аркуш набрано щільніше.
    font_size_pt: float | None = None
    word_spacing: float | None = None
    #: Записи, що почалися на цьому аркуші, і скільки рядків уже зайняв хвіст попереднього.
    records: list[tuple[int, SourceRecord]] = field(default_factory=list)
    fixed_rows: int = 0
    number_at_start: int = 1


def _wrap_cache_key(rec_i: int, engine: FontEngine) -> tuple[int, float, float]:
    """Ключ без id(): CPython повторно видає адреси GC-нутих двигунів."""
    return rec_i, engine.size_pt, engine.cfg.word_spacing


def _wrapped_body(
    rec_i: int,
    record: SourceRecord,
    engine: FontEngine,
    text_width: float,
    wrap_cache: dict[tuple[int, float, float], list[BodyLine]] | None,
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


def _signature_span(body: Sequence[BodyLine], index: int) -> int:
    """Скільки рядків підпису йде підряд, починаючи з index. 0 — це не початок блоку."""
    if index >= len(body) or not body[index].is_signature:
        return 0
    if index > 0 and body[index - 1].is_signature:
        return 0
    span = 0
    while index + span < len(body) and body[index + span].is_signature:
        span += 1
    return span


def _signature_fits(body: Sequence[BodyLine], row: int, rows_per_page: int) -> bool:
    """Блок підпису, що починається на цьому аркуші, закінчується на ньому ж."""
    for i, line in enumerate(body):
        span = _signature_span(body, i)
        if not span:
            continue
        start = row + i
        if start < rows_per_page and start + span > rows_per_page:
            return False
    return True


def layout(
    records: Sequence[SourceRecord],
    geo: TableGeometry,
    engine: FontEngine,
    cfg: Config,
    *,
    _wrap_cache: dict[tuple[int, float, float], list[BodyLine]] | None = None,
) -> list[Page]:
    """Розкладає записи по рядках бланка, повертає список аркушів.

    Блок підпису командира не розривається між аркушами. Якщо він припадає
    на кінець аркуша і не вміщається цілком, цей аркуш набирається щільніше
    (менша відстань між словами, за потреби менший кегль), щоб підпис лишився.
    """
    text_left, text_right = geo.text_bounds(2)
    text_width = text_right - text_left

    pages: list[Page] = [Page()]
    row = 0
    rows_per_page = geo.data_rows
    number = cfg.start_number
    page_engine = engine
    # Записи, що почалися на поточному аркуші: їх можна перекласти щільніше.
    flex: list[tuple[int, SourceRecord]] = []
    flex_number = number
    fixed_rows = 0

    def new_page() -> None:
        nonlocal row, page_engine, fixed_rows
        pages[-1].records = list(flex)
        pages[-1].fixed_rows = fixed_rows
        pages[-1].number_at_start = flex_number
        pages.append(Page())
        row = 0
        page_engine = engine
        flex.clear()
        fixed_rows = 0

    def take_label(record: SourceRecord, num: int) -> tuple[str, int, bool]:
        explicit = record.number.strip()
        hide = is_commander_signature(record.text)
        if hide:
            return "", num, True
        if explicit:
            nxt = int(explicit) + 1 if explicit.isdigit() else num
            return explicit, nxt, False
        return str(num), num + 1, False

    def place_lines(
        rec_i: int,
        record: SourceRecord,
        eng: FontEngine,
        num: int,
        *,
        guard_signature: bool,
    ) -> tuple[int, bool]:
        """Ставить запис. Повертає (новий номер, чи підпис не вмістився)."""
        nonlocal row
        body_lines = _wrapped_body(rec_i, record, eng, text_width, _wrap_cache)
        needed = max(len(body_lines), len(record.date_lines), 1)
        label, num, hide_number = take_label(record, num)
        page = pages[-1]

        if row >= rows_per_page:
            new_page()
            page = pages[-1]
        if (
            not cfg.split_records
            and not any(line.is_signature for line in body_lines)
            and row + needed > rows_per_page
            and needed <= rows_per_page
        ):
            new_page()
            page = pages[-1]

        for i in range(needed):
            body = body_lines[i] if i < len(body_lines) else None
            is_blank = body is not None and body.is_blank
            span = _signature_span(body_lines, i) if body is not None else 0
            if (
                guard_signature
                and span
                and span <= rows_per_page
                and row < rows_per_page
                and row + span > rows_per_page
            ):
                return num, True
            if row >= rows_per_page:
                if is_blank:
                    continue
                new_page()
                page = pages[-1]
            if i == 0 and cfg.number_column and not hide_number:
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
                        PlacedLine(
                            row, 2, text, align,
                            is_last=last,
                            is_signature=body.is_signature,
                        )
                    )
            row += 1
        return num, False

    def signature_fits(eng: FontEngine) -> bool:
        cursor = fixed_rows
        for rec_i, record in flex:
            body = _wrapped_body(rec_i, record, eng, text_width, _wrap_cache)
            if not _signature_fits(body, cursor, rows_per_page):
                return False
            cursor += max(len(body), len(record.date_lines), 1)
        return True

    def tighter_engine() -> FontEngine | None:
        if signature_fits(page_engine):
            return None
        current_pt = page_engine.size_pt
        current_ws = page_engine.cfg.word_spacing

        def at(pt: float, spacing: float) -> FontEngine:
            return page_engine.derive(pt).with_word_spacing(_snap_word_spacing(spacing))

        if signature_fits(at(current_pt, 0.0)):
            lo, hi = 0.0, current_ws
            best = at(current_pt, 0.0)
            while hi - lo > WORD_SPACING_STEP_MM + 1e-9:
                mid = _snap_word_spacing((lo + hi) / 2.0)
                if mid <= lo or mid >= hi:
                    break
                if signature_fits(at(current_pt, mid)):
                    best, lo = at(current_pt, mid), mid
                else:
                    hi = mid
            return best

        tight = at(MIN_FONT_SIZE_PT, 0.0)
        if not signature_fits(tight):
            return tight
        lo, hi = MIN_FONT_SIZE_PT, current_pt
        best = tight
        while hi - lo > FONT_SIZE_STEP_PT + 1e-9:
            mid = snap_font_pt((lo + hi) / 2.0)
            if mid <= lo or mid >= hi:
                break
            trial = at(mid, 0.0)
            if signature_fits(trial):
                best, lo = trial, mid
            else:
                hi = mid
        return best

    def _layout_rows(eng: FontEngine, grouped: list[tuple[int, SourceRecord]], start: int):
        """Куди лягає кожен рядок: (номер аркуша від поточного, чи це підпис, чи це звичайний запис)."""
        cursor = start
        sheet = 0
        out: list[tuple[int, bool, bool]] = []
        for rec_i, record in grouped:
            body = _wrapped_body(rec_i, record, eng, text_width, _wrap_cache)
            needed = max(len(body), len(record.date_lines), 1)
            entry = not is_commander_signature(record.text)
            for i in range(needed):
                line = body[i] if i < len(body) else None
                blank = line is not None and line.is_blank
                if cursor >= rows_per_page:
                    if blank:
                        continue
                    sheet += 1
                    cursor = 0
                out.append((sheet, bool(line and line.is_signature), entry and not blank))
                cursor += 1
        return out

    def _fits_on_previous(eng: FontEngine, grouped, start: int) -> bool:
        laid = _layout_rows(eng, grouped, start)
        sig_sheets = [sheet for sheet, sig, _entry in laid if sig]
        return bool(sig_sheets) and all(sheet == 0 for sheet in sig_sheets)

    def _shares_sheet(eng: FontEngine, grouped, start: int) -> bool:
        laid = _layout_rows(eng, grouped, start)
        sig_sheets = {sheet for sheet, sig, _entry in laid if sig}
        entry_sheets = {sheet for sheet, _sig, entry in laid if entry}
        # Підпис лишається цілим і ділить аркуш хоча б з одним записом.
        return len(sig_sheets) == 1 and bool(sig_sheets & entry_sheets)

    def _engine_at(base: FontEngine, pt: float, spacing: float) -> FontEngine:
        return base.derive(pt).with_word_spacing(_snap_word_spacing(spacing))

    def _keep_signature_with_entry() -> None:
        """Підпис не лишається на аркуші сам.

        Спершу переносимо його на попередній аркуш, стиснувши кегль і пробіл.
        Якщо так не вмістити — розтягуємо попередній аркуш, щоб разом із підписом
        на останній аркуш переїхав хоча б один запис.
        """
        nonlocal row, number, page_engine, fixed_rows
        if len(pages) < 2 or not _page_is_signature_only(pages[-1]):
            return
        prev, last = pages[-2], pages[-1]
        grouped = list(prev.records) + list(last.records)
        if not grouped:
            return
        start = prev.fixed_rows
        base = engine
        if prev.font_size_pt is not None:
            base = engine.derive(prev.font_size_pt).with_word_spacing(
                prev.word_spacing if prev.word_spacing is not None else engine.cfg.word_spacing
            )

        chosen: FontEngine | None = None
        if _fits_on_previous(_engine_at(base, base.size_pt, 0.0), grouped, start):
            lo, hi = 0.0, base.cfg.word_spacing
            chosen = _engine_at(base, base.size_pt, 0.0)
            while hi - lo > WORD_SPACING_STEP_MM + 1e-9:
                mid = _snap_word_spacing((lo + hi) / 2.0)
                if mid <= lo or mid >= hi:
                    break
                if _fits_on_previous(_engine_at(base, base.size_pt, mid), grouped, start):
                    chosen, lo = _engine_at(base, base.size_pt, mid), mid
                else:
                    hi = mid
        elif _fits_on_previous(_engine_at(base, MIN_FONT_SIZE_PT, 0.0), grouped, start):
            lo, hi = MIN_FONT_SIZE_PT, base.size_pt
            chosen = _engine_at(base, MIN_FONT_SIZE_PT, 0.0)
            while hi - lo > FONT_SIZE_STEP_PT + 1e-9:
                mid = snap_font_pt((lo + hi) / 2.0)
                if mid <= lo or mid >= hi:
                    break
                if _fits_on_previous(_engine_at(base, mid, 0.0), grouped, start):
                    chosen, lo = _engine_at(base, mid, 0.0), mid
                else:
                    hi = mid

        pull_back = chosen is not None and _fits_on_previous(chosen, grouped, start)
        if not pull_back:
            # Найменше збільшення пробілу, потім кегля, при якому підпис
            # ділить аркуш із записом і сам не розривається.
            chosen = None
            spacing = _snap_word_spacing(base.cfg.word_spacing + WORD_SPACING_STEP_MM)
            while spacing <= WORD_SPACING_MAX_MM + 1e-9:
                trial = _engine_at(base, base.size_pt, spacing)
                if _shares_sheet(trial, grouped, start):
                    chosen = trial
                    break
                spacing = _snap_word_spacing(spacing + WORD_SPACING_STEP_MM)
            if chosen is None:
                pt = snap_font_pt(base.size_pt + FONT_SIZE_STEP_PT)
                while pt <= MAX_FONT_SIZE_PT + 1e-9:
                    trial = _engine_at(base, pt, WORD_SPACING_MAX_MM)
                    if _shares_sheet(trial, grouped, start):
                        chosen = trial
                        break
                    pt = snap_font_pt(pt + FONT_SIZE_STEP_PT)
            if chosen is None:
                return

        prev.lines = [line for line in prev.lines if line.row < start]
        prev.font_size_pt = chosen.size_pt
        prev.word_spacing = chosen.cfg.word_spacing
        pages.pop()
        row = start
        number = prev.number_at_start
        page_engine = chosen
        fixed_rows = start
        flex.clear()
        for rec_i, record in grouped:
            flex.append((rec_i, record))
            number, _overflow = place_lines(
                rec_i, record, chosen, number, guard_signature=False
            )
        for page in pages[pages.index(prev):]:
            page.font_size_pt = chosen.size_pt
            page.word_spacing = chosen.cfg.word_spacing

    rec_i = 0
    while rec_i < len(records):
        record = records[rec_i]
        if row >= rows_per_page:
            new_page()
        if not flex:
            flex_number = number
            fixed_rows = row
        flex.append((rec_i, record))

        number, overflow = place_lines(
            rec_i, record, page_engine, number, guard_signature=True
        )
        if not overflow:
            rec_i += 1
            continue

        fitted = tighter_engine()
        page = pages[-1]
        page.lines = [line for line in page.lines if line.row < fixed_rows]
        row = fixed_rows
        number = flex_number
        if fitted is not None:
            page_engine = fitted
            page.font_size_pt = fitted.size_pt
            page.word_spacing = fitted.cfg.word_spacing
        saved = list(flex)
        flex.clear()
        for fr_i, fr in saved:
            if not flex:
                flex_number = number
            flex.append((fr_i, fr))
            number, _overflow = place_lines(
                fr_i, fr, page_engine, number, guard_signature=False
            )
        rec_i = saved[-1][0] + 1

    pages[-1].records = list(flex)
    pages[-1].fixed_rows = fixed_rows
    pages[-1].number_at_start = flex_number
    _keep_signature_with_entry()
    return [p for p in pages if p.lines]


def _page_is_signature_only(page: Page) -> bool:
    """На аркуші є підпис і немає жодного звичайного запису."""
    body = [line for line in page.lines if line.column == 2 and line.text]
    if not body:
        return False
    return all(line.is_signature for line in body)


def _count_pages(
    records: Sequence[SourceRecord],
    geo: TableGeometry,
    engine: FontEngine,
    cfg: Config,
    wrap_cache: dict[tuple[int, float, float], list[BodyLine]] | None = None,
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


#: Крок і стеля додаткової ширини пробілу, коли аркушів менше за мінімум.
WORD_SPACING_STEP_MM = 0.05
WORD_SPACING_MAX_MM = 12.0


def _snap_word_spacing(spacing: float) -> float:
    step = WORD_SPACING_STEP_MM
    return round(round(spacing / step) * step, 2)


@dataclass
class FontFitInfo:
    """Що змінили, щоб увійти в діапазон аркушів."""

    requested_pt: float
    font_size_pt: float
    pages: int
    uncompressed_pages: int
    word_spacing: float = 0.0


def _ensure_min_pages(
    records: Sequence[SourceRecord],
    geo: TableGeometry,
    engine: FontEngine,
    cfg: Config,
    info: FontFitInfo,
    wrap_cache: dict[tuple[int, float, float], list[BodyLine]],
) -> tuple[FontEngine, FontFitInfo, dict[tuple[int, float, float], list[BodyLine]]]:
    """Якщо аркушів менше за мінімум — збільшує відстань між словами.

    Береться найменший крок, при якому аркушів уже не менше мінімуму
    і ще не більше максимуму. Якщо в діапазон не потрапити, лишається
    варіант, найближчий до нього, але не вищий за максимум.
    """
    min_pages = cfg.min_pages
    info.word_spacing = engine.cfg.word_spacing
    if min_pages is None or min_pages <= 0 or info.pages >= min_pages:
        return engine, info, wrap_cache

    max_pages = cfg.max_pages if cfg.max_pages and cfg.max_pages > 0 else None
    base = engine.cfg.word_spacing

    def pages_at(spacing: float) -> tuple[float, int, FontEngine]:
        spacing = _snap_word_spacing(spacing)
        fitted = engine.with_word_spacing(spacing)
        return spacing, _count_pages(records, geo, fitted, cfg, wrap_cache), fitted

    lo = base
    lo_pages = info.pages
    hi = base
    hi_pages = info.pages
    hi_engine = engine
    step = 0.2
    while hi_pages < min_pages and hi < WORD_SPACING_MAX_MM - 1e-9:
        hi, hi_pages, hi_engine = pages_at(min(WORD_SPACING_MAX_MM, hi + step))
        step = min(step * 2.0, 2.0)
        if max_pages is not None and hi_pages > max_pages and hi - lo <= WORD_SPACING_STEP_MM:
            break

    if hi_pages < min_pages:
        info.word_spacing = hi
        info.pages = hi_pages
        return hi_engine, info, wrap_cache

    best_spacing, best_pages, best_engine = hi, hi_pages, hi_engine
    while hi - lo > WORD_SPACING_STEP_MM + 1e-9:
        mid, mid_pages, mid_engine = pages_at((lo + hi) / 2.0)
        if mid <= lo or mid >= hi:
            break
        if mid_pages < min_pages:
            lo, lo_pages = mid, mid_pages
        else:
            best_spacing, best_pages, best_engine = mid, mid_pages, mid_engine
            hi = mid

    if max_pages is not None and best_pages > max_pages:
        # Найменший крок уже перескочив максимум — відступаємо, доки
        # аркушів не стане не більше ліміту.
        while best_spacing > base + 1e-9 and best_pages > max_pages:
            nxt = _snap_word_spacing(best_spacing - WORD_SPACING_STEP_MM)
            if nxt >= best_spacing:
                break
            nxt, nxt_pages, nxt_engine = pages_at(nxt)
            best_spacing, best_pages, best_engine = nxt, nxt_pages, nxt_engine
            if nxt_pages < min_pages:
                break

    info.word_spacing = best_spacing
    info.pages = best_pages
    return best_engine, info, wrap_cache


def fit_font_to_max_pages(
    records: Sequence[SourceRecord],
    geo: TableGeometry,
    engine: FontEngine,
    cfg: Config,
) -> tuple[FontEngine, FontFitInfo, dict[tuple[int, float, float], list[BodyLine]]]:
    """Найбільший кегль ≤ 12 pt, при якому аркушів не більше cfg.max_pages."""
    requested_pt = min(engine.size_pt, MAX_FONT_SIZE_PT)
    engine = engine.derive(requested_pt)
    wrap_cache: dict[tuple[int, float, float], list[BodyLine]] = {}
    start_pages = _count_pages(records, geo, engine, cfg, wrap_cache)
    max_pages = cfg.max_pages

    info = FontFitInfo(
        requested_pt=requested_pt,
        font_size_pt=requested_pt,
        pages=start_pages,
        uncompressed_pages=start_pages,
    )
    if max_pages is None or max_pages <= 0 or start_pages <= max_pages:
        return _ensure_min_pages(records, geo, engine, cfg, info, wrap_cache)

    lo = MIN_FONT_SIZE_PT
    hi = requested_pt
    lo_pages = _count_pages(records, geo, engine.derive(lo), cfg, wrap_cache)
    if lo_pages > max_pages:
        info.font_size_pt = lo
        info.pages = lo_pages
        return _ensure_min_pages(records, geo, engine.derive(lo), cfg, info, wrap_cache)

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
    return _ensure_min_pages(records, geo, fitted, cfg, info, wrap_cache)
