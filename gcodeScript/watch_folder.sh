#!/bin/bash
# Викликається launchd, коли в теці проєкту щось змінилося. Знаходить нові
# .docx і генерує для них G-code, показуючи результат системним повідомленням.
# Увімкнути/вимкнути:  ./install_watcher.sh  /  ./install_watcher.sh --off
set -u

# Інакше Python створить __pycache__ у теці, за якою ми ж і стежимо,
# і watcher запуститься ще раз через власний побічний ефект.
export PYTHONDONTWRITEBYTECODE=1

DIR="$(cd "$(dirname "$0")" && pwd)"
OUT="$(cd "$DIR/.." && pwd)/gcodeOutput"
STATE="$OUT/.processed"        # поза текою, за якою стежимо
TEMPLATE="1 лист.docx"

mkdir -p "$OUT"
touch "$STATE"

notify() {
    /usr/bin/osascript -e "display notification \"${2//\"/}\" with title \"${1//\"/}\"" \
        >/dev/null 2>&1
}

# Файл міг ще копіюватися в момент спрацювання — даємо йому дописатися.
sleep 1

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
    notify "docx → G-code" "Не знайдено Python із бібліотекою fonttools."
    exit 1
fi

shopt -s nullglob
for file in "$DIR"/*.docx "$DIR"/*.DOCX; do
    name="$(basename "$file")"
    [ "$name" = "$TEMPLATE" ] && continue
    case "$name" in
        '~$'*) continue ;;  # тимчасові файли Word
    esac

    # Ключ із розміру й часу зміни: той самий файл двічі не обробляємо,
    # а відредагований — обробляємо знову.
    key="$name|$(stat -f '%m %z' "$file")"
    grep -qxF "$key" "$STATE" && continue

    if report="$("$PY" "$DIR/docx2gcode.py" "$file" -o "$OUT" 2>&1)"; then
        printf '%s\n' "$key" >>"$STATE"
        sheets="$(printf '%s\n' "$report" | grep -c '_p[0-9][0-9]\.gcode$')"
        notify "Готово: $name" "Аркушів: $sheets. Файли — у теці gcodeOutput."
    else
        notify "Не вдалося: $name" "$(printf '%s\n' "$report" | tail -n 1)"
    fi
done
