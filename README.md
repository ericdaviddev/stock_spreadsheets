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
1. **Ensure your downloaded brokerage file is open in Excel**.
2. Open the **VBA Editor (`Alt + F11`)**.
3. Select the **Module1** module.
4. Run the macro by pressing `F5` or using `"Run"` from the VBA Editor toolbar.
5. The macro will:
   - Filter out excluded symbols.
   - Sum relevant columns.
   - Apply basic formatting.
6. **Check the final processed spreadsheet** for key changes.

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
