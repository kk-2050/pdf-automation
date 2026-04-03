@echo off
setlocal
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
  call ".venv\Scripts\activate.bat"
  python highlight_gui_v2.py
  goto :eof
)

python highlight_gui_v2.py
