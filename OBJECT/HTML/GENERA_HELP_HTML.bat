@echo off
rem ==========================================================================
rem  GENERA_HELP_HTML - rigenera tutte le pagine help .html dai file .md
rem  cartelle: OBJECT\HTML\[TIPO]\[NomeOggetto]\[NomeOggetto]_HELP_nn_CHIAVE.md
rem  regole del formato: OBJECT\HTML\LEGGIMI_HELP.md
rem ==========================================================================
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (py -3 "_STRUMENTI\genera_help.py") else (python "_STRUMENTI\genera_help.py")
echo.
pause
