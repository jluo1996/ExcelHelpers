@echo off
REM Get the directory of the batch file
set SCRIPT_DIR=%~dp0

REM Get the path for python.exe
set PYTHON_EXE_DIR="%SCRIPT_DIR%Python313\python.exe"

REM Get the path for MainApp.py
set MAIN_APP_PY_DIR = "%SCRIPT_DIR%MainApp.py"

REM Run the Python script using python.exe in a subfolder (replace subfolder name if needed)
start "" /b "%SCRIPT_DIR%Python313\pythonw.exe" "%SCRIPT_DIR%MainApp.py"