@echo off
echo =============================================
echo  Build PV Generator — executable Windows
echo =============================================

:: Installer PyInstaller si absent
pip install pyinstaller --quiet

:: Construire l'exe (un seul fichier, sans console, avec icone)
pyinstaller ^
    --onefile ^
    --windowed ^
    --icon=assets\app_icon.ico ^
    --name="PV_Generator" ^
    --add-data="assets;assets" ^
    --add-data="filieres.json;." ^
    --add-data="modules;modules" ^
    app.py

echo.
echo === Termine ! L'executable se trouve dans le dossier dist\ ===
pause
