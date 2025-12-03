# Stock Spreadsheets

I use this to quickly clean up/export brokerage position files and run my own analysis on whether paid stock-picking services add value versus sticking with simple index tracking. The goal is to automate the repetitive Excel steps so the process becomes a one-click workflow.

This repo includes:
- An Excel macro workbook that processes a brokerage positions CSV  
- A Python helper script to automate file selection and macro execution  
- Sample positions files and an exclusion list  

---

## How It Works

1. Download your latest positions CSV from your brokerage.  
2. Run the Python script (manually or through a batch file).  
3. The script:
   - Finds the newest positions file in your download folder  
   - Copies/renames it to a consistent working file (e.g., `price_sheet.xlsx`)  
   - Loads Excel  
   - Executes the macro inside the `.xlsm` workbook  
4. The macro processes and summarizes the positions into your working Excel file.

This removes the manual steps of opening Excel files, navigating sheets, selecting macros, and handling the CSV each time.

---

## Python Helper Script

`api/main.py` automates the workflow using five arguments:

1. **Download folder** — directory containing brokerage CSV files  
2. **Output Excel file** — destination for the processed data  
3. **Exclusion list file** — Excel file listing symbols to exclude  
4. **Macro workbook** (`.xlsm`) — contains the Excel macro logic  
5. **Macro name** — the macro to execute (e.g., `ProcessExclusionsAndTotals`)  

The script uses Python and COM automation to load Excel, run the macro, and save the updated workbook.

---

## Example Windows Batch File

```bat
@echo off

call "C:\path\to\stock_spreadsheets\venv\Scripts\activate.bat"
python "C:\path\to\stock_spreadsheets\api\main.py" ^
  "C:\Users\YourName\Downloads" ^
  "C:\Users\YourName\Documents\price_sheet.xlsx" ^
  "C:\path\to\stock_spreadsheets\ExcludeSymbolsList.xlsx" ^
  "C:\path\to\stock_spreadsheets\StockSymbolMacro.xlsm" ^
  "ProcessExclusionsAndTotals"

pause
```
---

## Project Structure
```stock_spreadsheets/
├── api/
│   └── main.py            # Python automation script
├── ExcludeSymbolsList.xlsx
├── StockSymbolMacro.xlsm  # Excel macro for processing positions
├── Sample Positions1.csv
├── Sample Positions2.csv
└── README.md
```

---

## Requirements

- **Windows + Microsoft Excel** (required for COM automation)  
- **Python 3.x**  
- **pywin32** (used to control Excel from Python)  
- Optional: a **Python virtual environment**

Install `pywin32` with:

```bash
pip install pywin32
```

## Notes

- The Python script automates Excel; the macro workbook contains the actual processing logic.
- Sample CSVs are included for testing.
- Paths in the examples are placeholders—update them for your environment.
- The batch file is optional but convenient for repeat runs.

## License
Personal project — no license specified.
