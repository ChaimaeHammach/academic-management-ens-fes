@echo off
echo Build Rattrapage Manager...
python -m pip install pyinstaller --quiet
python -m PyInstaller --onefile --windowed --icon=assets\app_icon.ico --name="Rattrapage_Manager" --add-data="assets;assets" app.py
echo === Termine ! Executable dans dist\ ===
pause
