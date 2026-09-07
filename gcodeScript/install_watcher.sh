#!/bin/bash
# Вмикає спостереження за текою проєкту: киньте .docx у цю теку — G-code
# зʼявиться сам, відкривати теку й тягнути файл на застосунок не потрібно.
#
#   ./install_watcher.sh        увімкнути
#   ./install_watcher.sh --off  вимкнути
set -eu

DIR="$(cd "$(dirname "$0")" && pwd)"
LABEL="com.docx2gcode.watch"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
DOMAIN="gui/$(id -u)"

if [ "${1:-}" = "--off" ]; then
    launchctl bootout "$DOMAIN/$LABEL" 2>/dev/null || true
    rm -f "$PLIST"
    echo "Спостереження вимкнено."
    exit 0
fi

# macOS не дає ф    новим агентам читати Documents, Desktop і Downloads, причому
# молча: агент запуститься, але отримає «Operation not permitted». Краще
# відмовитися одразу, ніж залишити конфігурацію, яка тихо не працює.
case "$DIR/" in
    "$HOME"/Documents/* | "$HOME"/Desktop/* | "$HOME"/Downloads/*)
        echo "Спостереження неможливе: тека лежить у захищеному місці." >&2
        echo >&2
        echo "  $DIR" >&2
        echo >&2
        echo "macOS блокує фоновим агентам доступ до Documents, Desktop і" >&2
        echo "Downloads. Перенесіть теку проєкту, наприклад у ~/gcodeScript," >&2
        echo "і запустіть цей скрипт знову." >&2
        echo >&2
        echo "Якщо переносити не хочеться — використайте «Створити G-code.app»:" >&2
        echo "    ./make_droplet.sh ~/Desktop" >&2
        exit 1
        ;;
esac

mkdir -p "$(dirname "$PLIST")"
cat >"$PLIST" <<PLIST_END
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>$LABEL</string>
    <key>ProgramArguments</key>
    <array>
        <string>/bin/bash</string>
        <string>$DIR/watch_folder.sh</string>
    </array>
    <key>WatchPaths</key>
    <array>
        <string>$DIR</string>
    </array>
    <key>StandardOutPath</key>
    <string>/tmp/$LABEL.log</string>
    <key>StandardErrorPath</key>
    <string>/tmp/$LABEL.log</string>
</dict>
</plist>
PLIST_END

chmod +x "$DIR/watch_folder.sh"
launchctl bootout "$DOMAIN/$LABEL" 2>/dev/null || true
launchctl bootstrap "$DOMAIN" "$PLIST"

echo "Спостереження увімкнено за текою:"
echo "  $DIR"
echo
echo "Киньте .docx у цю теку — G-code зʼявиться в ../gcodeOutput."
echo "Журнал помилок:  /tmp/$LABEL.log"
