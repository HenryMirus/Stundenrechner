@echo off
echo ============================================
echo   Stundenrechner - Build-Skript
echo ============================================
echo.

echo [1/3] Installiere Abhaengigkeiten...
pip install -r requirements.txt
if errorlevel 1 (
    echo FEHLER: Abhaengigkeiten konnten nicht installiert werden.
    pause
    exit /b 1
)

echo.
echo [2/3] Installiere PyInstaller...
pip install pyinstaller
if errorlevel 1 (
    echo FEHLER: PyInstaller konnte nicht installiert werden.
    pause
    exit /b 1
)

echo.
echo [3/3] Erstelle ausfuehrbare Datei...
pyinstaller --onefile --windowed --name Stundenrechner --collect-all ttkbootstrap --collect-all msal --hidden-import=requests --hidden-import=msal --clean app.py
if errorlevel 1 (
    echo FEHLER: Build fehlgeschlagen.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Build erfolgreich!
echo.
echo   Die Datei befindet sich in: dist\Stundenrechner.exe
echo   Diese EXE koennen Sie weitergeben.
echo ============================================
pause
