@echo off
REM Doble clic para abrir Control Logistico.
cd /d "%~dp0"
where py >nul 2>nul && (set PY=py) || (set PY=python)
%PY% -c "import openpyxl" 2>nul || %PY% -m pip install --quiet openpyxl
start "Control Logistico" /min %PY% logistica.py
