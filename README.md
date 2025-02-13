# Stock Spreadsheet Processor VBA Macro

## 📌 Overview
This VBA macro automates the processing of stock and bond position spreadsheets downloaded from a brokerage firm.  
It allows users to **filter out specific stock symbols** (from an exclusion list) and **calculate key totals** for the remaining positions.  
Additionally, the macro applies **basic formatting** to highlight important financial changes, making it easier to track gains and losses.

## 🚀 Features
- **Filters out excluded stock/bond symbols** so you can focus on relevant holdings.
- **Sums up key financial values** (e.g., total gains/losses, current value).
- **Formats the spreadsheet** for easier readability.
- **Supports `.csv` and `.xlsx` files**.

## 📂 Installation – How to Import the Macro
To use this VBA macro in **Microsoft Excel**, follow these steps:

1. **Open Excel** and press `Alt + F11` to open the **VBA Editor**.
2. In the **VBA Project Explorer** (left panel), right-click on `"Modules"` and choose `"Import File..."`.
3. Select the file: **`Module1.bas`** and click **Open**.
4. The macro will now appear under `"Modules"` in your VBA Editor.

## 🛠️ How to Run the Macro
1. I have a batch file (windows - each line has a number below) 
2. The arguments are: folder_path (where main.py is), where to look for symbol files, output_file, exclusion_file, macro_file, macro_name
3. @echo off
4. python "C:\SomeLocation\stock_spreadsheets\api\main.py" "C:\SomeLocation\Downloads" "C:\SomeLocation\Documents\price_sheet.xlsx" "C:\SomeLocation\stock_spreadsheets\ExcludeSymbolsList.xlsx" "C:\SomeLocation\stock_spreadsheets\StockSymbolMacro.xlsm" "ProcessExclusionsAndTotals"
5. pause


## 📌 Requirements
- **Microsoft Excel (2016 or later recommended)**.
- **Macros must be enabled** (`File` → `Options` → `Trust Center` → `Enable Macros`).
- **An exclusion list must be maintained in a separate sheet** (if required for filtering).

## ❓ Troubleshooting
- **Macro doesn't run?** Ensure macros are **enabled** in Excel settings.
- **Wrong symbols filtered?** Check if the **exclusion list is correctly formatted**.
- **Numbers stored as text?** Try converting columns to **Number Format** manually.

## 📜 License
This project is open-source. Feel free to modify and use it for personal finance tracking.  

---
