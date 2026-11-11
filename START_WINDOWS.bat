@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py") else (set "PY=python")
%PY% -m pip install -r requirements.txt
if errorlevel 1 (echo Could not install requirements. Install Python 3.10+ and retry.&pause&exit /b 1)
start "" http://localhost:8765
%PY% app.py
pause
