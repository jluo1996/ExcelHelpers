# Excel Insurance Provider Helper

This project provides a GUI tool to assist with processing and validating insurance provider Excel files, including ADP and other insurance data files.

## Features
- Select and validate ADP and insurance Excel files (.xlsx)
- Choose insurance provider and plan type
- Specify output folder for reports
- Generate status reports with validation and error handling


## Requirements
- Python 3.13+
- PyQt5
- pandas
- colorama
- openpyxl
- tkinter (standard with Python on Windows)

## How to Run
### Recommended: Use the batch file (Windows)
1. Double-click `Click_Here_To_Run.bat` in the `ExcelInsuranceProviderHelper` folder.
   - This script will:
     - Navigate to the correct folder
   - Check for and install required dependencies (`PyQt5`, `pandas`, `colorama`, `openpyxl`) automatically if missing
     - Launch the application using `pythonw.exe` (no console window)

### Manual method
1. Install dependencies:
   ```bash
   pip install PyQt5 pandas colorama openpyxl
   ```
2. Run the application:
   ```bash
   python MainApp.py
   ```


## File Structure
- `MainApp.py` - Main GUI application
- `InsuranceStatusHelper.py` - Core logic for insurance status
- `InsuranceStatusHelperEnum.py` - Enum definitions for plan types, providers, statuses
- `GetUniqueEmployeeForLife.py` - Legacy script for specific use case
- `InsuranceProviderExcelFileConverter/ExcelFileConverter.py` - Excel file conversion utilities
- `Click_Here_To_Run.bat` - Windows batch file for easy startup

## Usage
1. Launch the app.
2. Select the ADP and insurance Excel files.
3. Choose the insurance provider and plan type.
4. Select the output folder.
5. Click "Generate Status Report" to process and generate results.


## Notes
- Only `.xlsx` files are supported for input.
- Output folder must be a valid directory.
