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
python main.py
if errorlevel 1 pause
