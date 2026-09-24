"""Параметри перетворення і сталі проєкту.

Config — це інтерфейс усього пайплайна: CLI лише заповнює його,
а розкладка, шрифт і вивід читають ті самі поля. Тому дефолти живуть
тут один раз, а argparse бере їх з Config(), а не дублює числа.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

#: Тека пакета (src/docx2gcode) і тека проєкту (gcodeScript) — там лежать
#: бланк і TTF, які шукаємо відносно коду, а не робочої теки.
PACKAGE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = PACKAGE_DIR.parent.parent

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


def snap_font_pt(size_pt: float) -> float:
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


@dataclass
class Config:
    """Усі параметри перетворення. Значення за замовчуванням — зі специфікації плотера."""

    # --- плотер ---
    feed: float = 2000.0  # мм/хв, робоча подача (знижена через глибокий натиск)
    feed_travel: float = 3000.0  # мм/хв, холості переміщення (G0 ігнорує, але лишаємо)
    feed_z: float = 2000.0  # мм, подача по Z
    z_safe: float = 1.0  # мм, висота підйому пера
    z_draw: float = -0.4  # мм, сила натиску (заглиблення; має бути нижче паперу)

    # --- шрифт ---
    #: Назва родини, а не імʼя файлу: «Segoe_Script.ttf» лежить у теці проєкту
    #: й індексується разом із системними шрифтами (див. fontindex).
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
