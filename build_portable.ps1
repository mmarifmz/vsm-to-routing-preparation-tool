$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    py -3 -m venv .venv
}

& ".venv\Scripts\python.exe" -m pip install --upgrade pip
& ".venv\Scripts\python.exe" -m pip install -r requirements.txt

& ".venv\Scripts\python.exe" -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name "VSM_to_ROUTING_Preparation_Tool_v1.3.8" `
    --icon "assets\app_icon.ico" `
    --collect-all ttkbootstrap `
    --hidden-import PIL._tkinter_finder `
    --hidden-import win32com `
    --hidden-import win32com.client `
    main.py

Write-Host "Build complete: $PSScriptRoot\dist\VSM_to_ROUTING_Preparation_Tool_v1.3.8.exe" -ForegroundColor Green
