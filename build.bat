@echo off
chcp 65001 >nul
echo ============================================
echo  Build - Validator kapacit (.exe)
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
pip install pyinstaller openpyxl Pillow
if errorlevel 1 (
    echo CHYBA: Instalace selhala.
    pause
    exit /b 1
)

echo.
echo Generuji PNG assety (logo)...
python graphics\create_assets.py
if errorlevel 1 (
    echo CHYBA: Generovani assestu selhalo.
    pause
    exit /b 1
)

echo.
echo Stavim .exe soubor...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name ValidatorKapacit ^
    --add-data "graphics\parama-symbol.png;graphics" ^
    --add-data "graphics\parama-icon.png;graphics" ^
    --exclude-module numpy ^
    --exclude-module lxml ^
    --exclude-module PIL.ImageQt ^
    --exclude-module PIL.ImageTk._imagingtk ^
    validator.py

if errorlevel 1 (
    echo CHYBA: Build selhal.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Hotovo!
echo  Spustitelny soubor: dist\ValidatorKapacit.exe
echo ============================================
echo.
pause
