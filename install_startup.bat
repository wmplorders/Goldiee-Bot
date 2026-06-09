@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut([Environment]::GetFolderPath('Startup') + '\Goldiee Purchase Bot.lnk'); $s.TargetPath='%~dp0start_bot.bat'; $s.WorkingDirectory='%~dp0'; $s.Save()"
echo Startup shortcut installed.
pause
