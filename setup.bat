@echo off
chcp 65001 >nul
echo ============================================
echo  Nastaveni - Validator kapacit
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo CHYBA: Python neni nalezen.
    echo Stahni Python z https://www.python.org/downloads/
    echo Pri instalaci zatrhni "Add Python to PATH"!
    pause
    exit /b 1
)

echo Instaluji zavislosti...
pip install -r requirements.txt
if errorlevel 1 (
    echo CHYBA: Instalace selhala.
    pause
    exit /b 1
)

echo.
echo Instalace dokoncena!
echo Spust aplikaci prikazem:  python validator.py
echo.
pause
