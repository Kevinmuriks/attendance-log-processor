# Attendance Log Processor

## Overview
This Python script processes an Excel-based attendance log by:
- Creating a separate sheet for each employee.
- Consolidating check-in and check-out times into a single row per day.
- Calculating overtime (hours worked beyond 8 hours per day).

## Features
- Automatically detects employee names or ID numbers.
- Removes duplicate rows by merging check-in and check-out records.
- Computes overtime for each workday.
- Outputs a structured Excel file with individual sheets per employee.

## Requirements
Ensure you have the following dependencies installed:
```sh
pip install pandas openpyxl xlsxwriter
```

## Usage
1. Place your attendance log file in the project directory.
2. Modify `input_path` in `split_attendance_sheets.py` to match your file name.
3. Run the script:
```sh
python split_attendance_sheets.py
```
4. The processed Excel file will be generated as `Processed_Attendance_Logs.xlsx`.

## Contributing
Feel free to submit issues or pull requests to improve the script.

## License
This project is open-source and available under the MIT License.

