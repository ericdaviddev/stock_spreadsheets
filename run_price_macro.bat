@echo off

call "F:\backup\code\stock_spreadsheets\venv\Scripts\activate.bat"
python "F:\backup\code\stock_spreadsheets\api\main.py" "C:\Users\erics\Downloads" "C:\Users\erics\Documents\price_sheet.xlsx" "F:\backup\code\stock_spreadsheets\ExcludeSymbolsList.xlsx" "F:\backup\code\stock_spreadsheets\StockSymbolMacro.xlsm" "ProcessExclusionsAndTotals"
pause
