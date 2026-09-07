#!/bin/bash
# Збирає «Створити G-code.app», вшиваючи в неї шлях до цієї теки. Завдяки
# цьому програму можна тримати будь-де — теку проєкту відкривати не треба.
#
#   ./make_droplet.sh                зібрати поруч із текою проєкту (типово)
#   ./make_droplet.sh ~/Desktop      зібрати на робочому столі
#   ./make_droplet.sh .              зібрати всередині теки проєкту
#
# Після перенесення або перейменування теки проєкту скрипт треба запустити
# знову, щоб оновити вшитий шлях.
set -eu

DIR="$(cd "$(dirname "$0")" && pwd)"
NAME="Створити G-code.app"
BUNDLE_ID="com.docx2gcode.droplet"
# Типово — на рівень вище, щоб обгортки лежали поруч із текою проєкту разом
# із текою gcodeOutput, і теку не треба було відкривати.
DEST="$(cd "${1:-$DIR/..}" && pwd)"
APP="$DEST/$NAME"

tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT

# Шлях не містить «|», тому його безпечно взяти як розділювач sed.
sed "s|__PROJECT_DIR__|$DIR|" "$DIR/droplet.applescript" >"$tmp/droplet.applescript"

rm -rf "$APP"
osacompile -o "$APP" "$tmp/droplet.applescript"

# Стала назва в системі. Без неї macOS не має за що «зачепити» дозвіл на доступ
# до тек, і після кожної перезбірки дозвіл доводилося б давати заново.
/usr/libexec/PlistBuddy \
    -c "add :CFBundleIdentifier string $BUNDLE_ID" \
    "$APP/Contents/Info.plist" >/dev/null

# osacompile уже дозволяє кидати на іконку будь-який файл. Додатково оголошуємо
# саме .docx — тоді програма зʼявляється в меню «Відкрити у програмі», тобто
# файл можна обробити навіть без перетягування. Rank «Alternate», щоб не
# перебивати Word чи Pages як типову програму для .docx.
/usr/libexec/PlistBuddy \
    -c 'add :CFBundleDocumentTypes:0 dict' \
    -c 'add :CFBundleDocumentTypes:0:CFBundleTypeName string "Word document"' \
    -c 'add :CFBundleDocumentTypes:0:CFBundleTypeRole string Viewer' \
    -c 'add :CFBundleDocumentTypes:0:LSHandlerRank string Alternate' \
    -c 'add :CFBundleDocumentTypes:0:LSItemContentTypes array' \
    -c 'add :CFBundleDocumentTypes:0:LSItemContentTypes:0 string org.openxmlformats.wordprocessingml.document' \
    "$APP/Contents/Info.plist" >/dev/null

# ОБОВʼЯЗКОВО останнім кроком: osacompile підписує програму сам, а будь-яка
# правка Info.plist після того ламає підпис. Зі зламаним підписом macOS не може
# встановити особу програми й відмовляє в доступі до тек навіть без запиту —
# у вікні видно «Operation not permitted».
codesign --force --sign - "$APP" 2>/dev/null
codesign --verify "$APP" || {
    echo "Підпис не пройшов перевірку — програма не отримає доступу до файлів." >&2
    exit 1
}

# Щоб Finder і Dock побачили новий тип без перезаходу в систему.
/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister \
    -f "$APP" 2>/dev/null || true

# Windows-обгортка сама знаходить проєкт поруч із собою, тому достатньо копії.
if [ "$DEST" != "$DIR" ]; then
    cp "$DIR/Створити G-code.bat" "$DEST/Створити G-code.bat"
fi

echo "Готово: $APP"
echo "Вшитий шлях до проєкту: $DIR"
if [ "$DEST" != "$DIR" ]; then
    echo "Копія Windows-обгортки: $DEST/Створити G-code.bat"
    echo
    echo "Тепер кидайте .docx на цю іконку — файл залишиться на місці,"
    echo "а G-code зʼявиться в $(cd "$DIR/.." && pwd)/gcodeOutput"
fi
