"""Читання заповненої таблиці і геометрії порожнього бланка.

Інтерфейс: read_source(), read_template(), SourceRecord, TableGeometry.
Word XML і twips лишаються всередині модуля.
"""

from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from .config import PT_PER_MM, TNR_FALLBACK_METRICS, W, Config
from .geom import mm_from_twips
from .fonts import find_font
from ._fonttools import TTFont

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
