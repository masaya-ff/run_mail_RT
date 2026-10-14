@echo off
setlocal
chcp 65001 > nul
cd /d "%~dp0"
python "%~dp0run_mail_routine.py"
echo.
pause
endlocal
