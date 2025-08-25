# Excel Insurance Provider Helper

This project provides a GUI tool to assist with processing and validating insurance provider Excel files, including ADP and other insurance data files.

## Features
- Select and validate ADP and insurance Excel files (.xlsx)
- Choose insurance provider and plan type
- Specify output folder for reports
- Generate status reports with validation and error handling

## Requirements
- Python 3.9+
- PyQt5
- tkinter

## How to Run
1. Install dependencies:
   ```bash
   pip install PyQt5
   ```
2. Run the application:
   ```bash
   python MainApp.py
   ```
   Or use the provided `Run_This.bat` on Windows.

## File Structure
- `MainApp.py` - Main GUI application
- `InsuranceStatusHelper.py` - Core logic for insurance status
- `InsuranceStatusHelperEnum.py` - Enum definitions for plan types, providers, statuses
- `GetUniqueEmployeeForLife.py` - Legacy script for specific use case
- `InsuranceProviderExcelFileConverter/ExcelFileConverter.py` - Excel file conversion utilities

## Usage
1. Launch the app.
2. Select the ADP and insurance Excel files.
3. Choose the insurance provider and plan type.
4. Select the output folder.
5. Click "Generate Status Report" to process and generate results.

## Notes
- Only `.xlsx` files are supported for input.
- Output folder must be a valid directory.
