@echo off
rem Windows-варіант drag & drop: киньте .docx на цей файл. Сам файл при цьому
rem залишається на місці — сюди він не переміщується.
rem
rem Файл працює у двох положеннях: поруч із текою gcodeScript або всередині
rem неї — потрібне сам знаходить. Якщо теку проєкту перейменували чи поклали
rem кудись інде, впишіть шлях до неї у PROJECT нижче вручну.
setlocal

set "PROJECT=%~dp0"
if not exist "%PROJECT%docx2gcode.py" set "PROJECT=%~dp0gcodeScript\"
rem Приклад для нетипового розташування:
rem set "PROJECT=C:\Users\Artem\Documents\gcodeScript\"

if "%~1"=="" (
    echo Перетягніть заповнену таблицю ^(.docx^) на цей файл.
    echo.
    pause
    exit /b 1
)

if /I not "%~x1"==".docx" (
    echo Потрібен файл .docx, а не "%~nx1".
    echo.
    pause
    exit /b 1
)

set "SCRIPT=%PROJECT%docx2gcode.py"
if not exist "%SCRIPT%" (
    echo Не знайдено docx2gcode.py за шляхом:
    echo     %SCRIPT%
    echo.
    echo Тримайте цей файл поруч із текою gcodeScript або всередині неї.
    echo Інакше впишіть шлях до теки проєкту у змінну PROJECT
    echo на початку цього файлу.
    echo.
    pause
    exit /b 1
)

rem Теку виводу створюємо заздалегідь: для -o без розширення скрипт розрізняє
rem теку й префікс імені за тим, чи тека вже існує.
set "OUT=%PROJECT%..\gcodeOutput"
if not exist "%OUT%" mkdir "%OUT%"

python "%SCRIPT%" "%~1" -o "%OUT%"
if errorlevel 1 (
    echo.
    echo Не вдалося створити G-code.
    echo Якщо бракує бібліотеки, виконайте:  python -m pip install fonttools
) else (
    echo.
    echo Готово. Файли — у теці gcodeOutput.
)

echo.
pause
