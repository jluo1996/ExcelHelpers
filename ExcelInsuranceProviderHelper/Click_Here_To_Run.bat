@echo off
REM ------------------------------
REM Navigate to the app folder
REM ------------------------------
cd /d D:\SRC\ExcelHelpers\ExcelHelpers

REM ------------------------------
REM Check/install PyQt5
REM ------------------------------
python -c "import PyQt5" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo PyQt5 not found. Installing...
    pip install pyqt5
    if %ERRORLEVEL% NEQ 0 (
        echo Failed to install PyQt5! Please install it manually.
        pause
        exit /b 1
    )
)

REM ------------------------------
REM Check/install pandas
REM ------------------------------
python -c "import pandas" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo pandas not found. Installing...
    pip install pandas
    if %ERRORLEVEL% NEQ 0 (
        echo Failed to install pandas! Please install it manually.
        pause
        exit /b 1
    )
)

REM ------------------------------
REM Check/install colorama
REM ------------------------------
python -c "import colorama" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo colorama not found. Installing...
    pip install colorama
    if %ERRORLEVEL% NEQ 0 (
        echo Failed to install colorama! Please install it manually.
        pause
        exit /b 1
    )
)

REM ------------------------------
REM Start the PyQt5 app without console
REM Start with pythonw.exe to hide python console
REM ------------------------------
start "" pythonw.exe -m ExcelInsuranceProviderHelper.MainApp 