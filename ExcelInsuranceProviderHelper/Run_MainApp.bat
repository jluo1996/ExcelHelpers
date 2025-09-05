@echo off
REM Get the directory of the batch file
set SCRIPT_DIR=%~dp0

REM Get the path for python.exe
set PYTHON_EXE_DIR="%SCRIPT_DIR%Python313\python.exe"

REM Get the path for MainApp.py
set MAIN_APP_PY_DIR = "%SCRIPT_DIR%MainApp.py"

REM Run the Python script using pythonw.exe with optimization (-O)
start "" /b "%SCRIPT_DIR%Python313\pythonw.exe" -O "%SCRIPT_DIR%MainApp.py"