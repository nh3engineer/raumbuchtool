@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ================================================
echo  Raumbuch-Tool Kaeltetechnik wird gestartet...
echo ================================================
echo.
where python >nul 2>nul
if errorlevel 1 (
    echo FEHLER: Python wurde nicht gefunden.
    echo Bitte Python von python.org installieren mit Haekchen "Add python.exe to PATH".
    pause
    exit /b 1
)
python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo openpyxl fehlt - wird jetzt installiert...
    python -m pip install openpyxl
)
echo.
python main.py
if errorlevel 1 (
    echo.
    echo ================================================
    echo  Das Programm wurde mit einem Fehler beendet.
    echo  Meldung siehe oben.
    echo ================================================
)
pause
