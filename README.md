# Chrome Passwords Exporter

A lightweight Windows utility that converts a Chrome/Chromium exported `Passwords.csv` into:
- `credentials.txt` with tab-separated columns for Domain, Username, and Password
- `credentials.xlsx` with the same data in a spreadsheet

This project includes both:
- a Python GUI version: `chrome_passwords_exporter.py`
- a compiled Windows executable: `chrome_passwords_exporter.exe`

## Features

- Reads Chrome password exports in CSV format
- Automatically detects common column names like `url`, `username`, and `password`
- Normalizes domains to a clean hostname
- Exports data to:
  - a UTF-8 text file for quick use
  - an Excel workbook for easier viewing and editing
- Includes a simple GUI for selecting the input CSV and output folder

## Requirements

If you are running the Python script directly:
- Python 3.x
- `openpyxl`

Install the dependency with:

```bash
python -m pip install openpyxl