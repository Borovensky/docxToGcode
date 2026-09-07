#!/bin/bash
# Запускач для drag & drop: викликається з «Створити G-code.app», коли на її
# іконку кидають .docx. Сам знаходить Python із fonttools і друкує звіт,
# який обгортка показує у вікні.
set -u

DOCX="${1:-}"
DIR="$(cd "$(dirname "$0")" && pwd)"
SCRIPT="$DIR/docx2gcode.py"
OUT="$(cd "$DIR/.." && pwd)/gcodeOutput"

if [ -z "$DOCX" ]; then
    echo "Перетягніть заповнену таблицю (.docx) на іконку програми." >&2
    exit 1
fi

case "$DOCX" in
    *.docx | *.DOCX) ;;
    *)
        echo "Потрібен файл .docx, а не «$(basename "$DOCX")»." >&2
        exit 1
        ;;
esac

if [ ! -f "$SCRIPT" ]; then
    echo "Не знайдено docx2gcode.py поруч із цим скриптом." >&2
    exit 1
fi

# Обовʼязково: для -o без розширення скрипт розрізняє теку й префікс імені за
# тим, чи тека існує. Без mkdir файли лягли б як «…/gcodeOutput_p01.gcode».
mkdir -p "$OUT"

# Перший Python, у якому вже є fonttools. .venv поруч зі скриптом — пріоритет.
PY=""
for candidate in \
    "$DIR/.venv/bin/python3" \
    /opt/homebrew/bin/python3 \
    /usr/local/bin/python3 \
    /Library/Frameworks/Python.framework/Versions/Current/bin/python3 \
    /usr/bin/python3; do
    [ -x "$candidate" ] || continue
    "$candidate" -c 'import fontTools' >/dev/null 2>&1 || continue
    PY="$candidate"
    break
done

if [ -z "$PY" ]; then
    echo "Не знайдено Python із бібліотекою fonttools." >&2
    echo "Відкрийте Термінал і виконайте:" >&2
    echo "    python3 -m pip install fonttools" >&2
    exit 1
fi

exec "$PY" "$SCRIPT" "$DOCX" -o "$OUT"
