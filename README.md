# ExcelHelpers

ExcelHelpers is a collection of Python utilities for processing, analyzing, and managing insurance-related data in Excel files. It is designed to help with employee insurance status, unique employee identification, and related reporting tasks.

## Project Structure

- `ExcelInsuranceProviderHelper/`
	- `GetUniqueEmployeeForLife_legacy.py`: Legacy script for specific use case
	- `InsuranceStatusHelper.py`: Main logic for insurance status processing
	- `InsuranceStatusHelperEnum.py`: Enum definitions for insurance status
	- `logger.py`: Logging utility
	- `MainApp.py`: Main application entry point
	- `Run_MainApp.bat`: Batch file to run the main app
	- `Python313/`: Embedded Python 3.13 environment (no external Python installation required)

## Features

- Insurance status processing from Excel files
- Unique employee identification for life insurance
- Self-contained Python environment (portable)
- Built-in logging

## Getting Started

1. Go to the `ExcelInsuranceProviderHelper` directory.
2. Run `Run_MainApp.bat` (double-click or from command line) to start the main application.

## Customization

- Edit the Python scripts in `ExcelInsuranceProviderHelper` to modify or extend logic.
- Update the embedded Python environment by replacing files in `Python313`.

## License

See `LICENSE.txt` in the `Python313` directory for Python distribution license. Add project-specific licensing as needed.
