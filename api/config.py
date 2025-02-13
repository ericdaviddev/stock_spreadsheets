# Specify data types for certain columns
data_types = {
    "Account Number": str,
    "Account Name": str,
    "Symbol": str,
    "Description": str,
    "Quantity": float,
    "Last Price": float,
    "Last Price Change": float,
    "Current Value": float,
    "Today's Gain/Loss Dollar": float,
    "Today's Gain/Loss Percent": float,
    "Total Gain/Loss Dollar": float,
    "Total Gain/Loss Percent": float,
    "Percent Of Account": float,
    "Cost Basis Total": float,
    "Average Cost Basis": float,
    "Type": str
}

# Explicitly convert numeric and percentage columns after cleaning
numeric_columns = ["Last Price", "Last Price Change", "Current Value", "Today's Gain/Loss Dollar",
                   "Total Gain/Loss Dollar", "Cost Basis Total", "Average Cost Basis" ]  # Adjust as needed
percentage_columns = ["Today's Gain/Loss Percent", "Total Gain/Loss Percent", "Percent Of Account"]  # Add percentage columns here

startsWithColumns = ["The data and information", "Brokerage services are", "Date downloaded"]

