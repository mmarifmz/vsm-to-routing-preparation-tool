@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if errorlevel 1 (
 echo Install Python 3.11 or 3.12 from python.org and enable PATH.
 pause
 exit /b 1
)

if not exist ".venv\Scripts\python.exe" py -3 -m venv .venv
call ".venv\Scripts\activate.bat"

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

python -m PyInstaller ^
 --noconfirm ^
 --clean ^
 --onefile ^
 --windowed ^
 --name "VSM_to_ROUTING_Preparation_Tool_v1.3.8" ^
 --icon "assets\app_icon.ico" ^
 --collect-all ttkbootstrap ^
 --hidden-import PIL._tkinter_finder ^
 --hidden-import win32com ^
 --hidden-import win32com.client ^
 main.py

if errorlevel 1 goto :error

echo Build complete:
echo %CD%\dist\VSM_to_ROUTING_Preparation_Tool_v1.3.8.exe
pause
exit /b 0

:error
echo Build failed.
pause
exit /b 1
