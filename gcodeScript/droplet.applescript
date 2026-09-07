-- Обгортка для запуску без термінала: .docx кидають на іконку програми.
-- Файл при цьому лишається на своєму місці — програма лише читає його.
--
-- Програму можна тримати будь-де (Dock, стіл, /Applications): шлях до теки
-- проєкту вшивається під час збірки. Перезібрати після правок:
--     ./make_droplet.sh

property projectDir : "__PROJECT_DIR__"

on run
	display dialog "Перетягніть заповнену таблицю (.docx) на іконку цієї програми." & return & return & "Файл залишиться там, де лежав, а готовий G-code зʼявиться в теці gcodeOutput." buttons {"Зрозуміло"} default button 1 with title "docx → G-code"
end run

on open droppedItems
	repeat with anItem in droppedItems
		handleFile(anItem)
	end repeat
end open

-- Тека проєкту: спершу вшитий шлях, потім — поруч із програмою (так
-- працювали старі збірки, коли .app мусила лежати всередині проєкту).
on resolveProject()
	if projectDir does not start with "__" and my fileExists(projectDir & "/docx2gcode.py") then
		return projectDir
	end if
	set appDir to do shell script "dirname " & quoted form of (POSIX path of (path to me))
	if my fileExists(appDir & "/docx2gcode.py") then return appDir
	return ""
end resolveProject

on fileExists(posixPath)
	try
		do shell script "test -f " & quoted form of posixPath
		return true
	on error
		return false
	end try
end fileExists

on handleFile(anItem)
	set docxPath to POSIX path of (anItem as alias)

	set root to my resolveProject()
	if root is "" then
		display alert "Не знайдено docx2gcode.py" message "Вшитий шлях до проєкту не працює:" & return & return & projectDir & return & return & "Якщо теку проєкту перенесли або перейменували, перезберіть програму: відкрийте теку проєкту в Терміналі й виконайте ./make_droplet.sh" as critical
		return
	end if

	try
		set report to do shell script "/bin/bash " & quoted form of (root & "/run_droplet.sh") & " " & quoted form of docxPath
	on error errorMessage
		display alert "Не вдалося створити G-code" message errorMessage as critical
		return
	end try

	display dialog report buttons {"Показати теку", "Готово"} default button 1 with title "Готово"
	if button returned of result is "Показати теку" then
		do shell script "open " & quoted form of (root & "/../gcodeOutput")
	end if
end handleFile
